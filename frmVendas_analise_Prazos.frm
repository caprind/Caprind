VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmVendas_analise_prazos 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Outros - Análise crítica - Definir prazos"
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7635
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
   ForeColor       =   &H8000000D&
   Icon            =   "frmVendas_analise_Prazos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   7635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   90
      TabIndex        =   5
      Top             =   0
      Width           =   7455
      _ExtentX        =   13150
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
      ButtonCaption2  =   "Excluir"
      ButtonEnabled2  =   0   'False
      ButtonIconSize2 =   32
      ButtonToolTipText2=   "Excluir (F4)"
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
      ButtonLeft2     =   42
      ButtonTop2      =   2
      ButtonWidth2    =   39
      ButtonHeight2   =   21
      ButtonUseMaskColor2=   0   'False
      ButtonEnabled3  =   0   'False
      ButtonIconSize3 =   32
      ButtonAlignment3=   2
      ButtonType3     =   1
      ButtonStyle3    =   -1
      BeginProperty ButtonFont3 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState3    =   -1
      ButtonLeft3     =   83
      ButtonTop3      =   4
      ButtonWidth3    =   2
      ButtonHeight3   =   54
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
      ButtonLeft4     =   87
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
      ButtonLeft5     =   125
      ButtonTop5      =   2
      ButtonWidth5    =   26
      ButtonHeight5   =   21
      ButtonUseMaskColor5=   0   'False
      ButtonEnabled6  =   0   'False
      ButtonIconSize6 =   32
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
      ButtonState6    =   5
      ButtonLeft6     =   153
      ButtonTop6      =   2
      ButtonWidth6    =   24
      ButtonHeight6   =   24
      ButtonUseMaskColor6=   0   'False
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   5670
         Top             =   180
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmVendas_analise_Prazos.frx":0442
         Count           =   1
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Prazos"
      Height          =   885
      Left            =   90
      TabIndex        =   6
      Top             =   990
      Width           =   7455
      Begin MSMask.MaskEdBox txtEngenharia 
         Height          =   315
         Left            =   180
         TabIndex        =   0
         ToolTipText     =   "Prazo de engenharia."
         Top             =   435
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtProcesso 
         Height          =   315
         Left            =   1627
         TabIndex        =   1
         ToolTipText     =   "Prazo do processo."
         Top             =   435
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtPCP 
         Height          =   315
         Left            =   3074
         TabIndex        =   2
         ToolTipText     =   "Prazo do PCP."
         Top             =   435
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtQualidade 
         Height          =   315
         Left            =   4521
         TabIndex        =   3
         ToolTipText     =   "Prazo da qualidade."
         Top             =   435
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtCompras 
         Height          =   315
         Left            =   5970
         TabIndex        =   4
         ToolTipText     =   "Prazo de compras."
         Top             =   435
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Compras"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   6150
         TabIndex        =   11
         Top             =   240
         Width           =   630
      End
      Begin VB.Image imgCompras 
         Height          =   360
         Left            =   6960
         Picture         =   "frmVendas_analise_Prazos.frx":2F11
         Stretch         =   -1  'True
         ToolTipText     =   "Abrir calendário."
         Top             =   405
         Width           =   330
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Qualidade"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   4656
         TabIndex        =   10
         Top             =   240
         Width           =   720
      End
      Begin VB.Image imgQualidade 
         Height          =   360
         Left            =   5520
         Picture         =   "frmVendas_analise_Prazos.frx":3394
         Stretch         =   -1  'True
         ToolTipText     =   "Abrir calendário."
         Top             =   405
         Width           =   330
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "PCP"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   3427
         TabIndex        =   9
         Top             =   240
         Width           =   285
      End
      Begin VB.Image imgPCP 
         Height          =   360
         Left            =   4080
         Picture         =   "frmVendas_analise_Prazos.frx":3817
         Stretch         =   -1  'True
         ToolTipText     =   "Abrir calendário."
         Top             =   405
         Width           =   330
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Processo"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   1800
         TabIndex        =   8
         Top             =   240
         Width           =   645
      End
      Begin VB.Image imgProcesso 
         Height          =   360
         Left            =   2610
         Picture         =   "frmVendas_analise_Prazos.frx":3C9A
         Stretch         =   -1  'True
         ToolTipText     =   "Abrir calendário."
         Top             =   405
         Width           =   330
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Engenharia"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   270
         TabIndex        =   7
         Top             =   220
         Width           =   810
      End
      Begin VB.Image imgEngenharia 
         Height          =   360
         Left            =   1170
         Picture         =   "frmVendas_analise_Prazos.frx":411D
         Stretch         =   -1  'True
         ToolTipText     =   "Abrir calendário."
         Top             =   405
         Width           =   330
      End
   End
End
Attribute VB_Name = "frmVendas_analise_prazos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CodReferencia   As String 'OK

Private Sub ProcSalvar()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

With frmVendas_analise
    If .Txt_status = "REVISADA" Then
        USMsgBox ("Não é permitida a alteração dos prazos da análise crítica revisada."), vbExclamation, "CAPRIND v5.0"
        Exit Sub
    End If
    
    Set TBGravar = CreateObject("adodb.recordset")
    TBGravar.Open "Select Status from vendas_analise where ID = " & .txtID, Conexao, adOpenKeyset, adLockOptimistic
    If TBGravar.EOF = False Then
        If TBGravar!status = "APROVADA" And .Txt_status = "APROVADA" Then
            USMsgBox ("Não é permitido alterar prazos desta análise, pois a mesma já está aprovada."), vbExclamation, "CAPRIND v5.0"
            TBGravar.Close
            Exit Sub
        End If
        If TBGravar!status = "APROVADA" And .Txt_status <> "APROVADA" Then
            Set TBFIltro = CreateObject("adodb.recordset")
            TBFIltro.Open "Select IDAnalise from vendas_carteira where IDAnalise = " & .txtID, Conexao, adOpenKeyset, adLockOptimistic
            If TBFIltro.EOF = False Then
                USMsgBox ("Não é permitido alterar prazos desta análise, pois a mesma está vinculada a uma proposta."), vbExclamation, "CAPRIND v5.0"
                TBFIltro.Close
                TBGravar.Close
                Exit Sub
            End If
            TBFIltro.Close
        End If
    End If
    TBGravar.Close
    
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select Prazo_engenharia, Prazo_processos, Prazo_pcp, Prazo_qualidade, Prazo_compras from Vendas_analise where ID = " & .txtID & "", Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
        TBProduto!Prazo_engenharia = IIf(txtEngenharia = "__/__/____", Null, txtEngenharia)
        TBProduto!Prazo_processos = IIf(txtProcesso = "__/__/____", Null, txtProcesso)
        TBProduto!Prazo_pcp = IIf(txtPCP = "__/__/____", Null, txtPCP)
        TBProduto!Prazo_qualidade = IIf(txtQualidade = "__/__/____", Null, txtQualidade)
        TBProduto!Prazo_compras = IIf(txtCompras = "__/__/____", Null, txtCompras)
        TBProduto.Update
        
        .txtPrazo_Engenharia = IIf(IsNull(TBProduto!Prazo_engenharia), "", Format(TBProduto!Prazo_engenharia, "dd/mm/yy"))
        .txtPrazo_Processo = IIf(IsNull(TBProduto!Prazo_processos), "", Format(TBProduto!Prazo_processos, "dd/mm/yy"))
        .txtPrazo_PCP = IIf(IsNull(TBProduto!Prazo_pcp), "", Format(TBProduto!Prazo_pcp, "dd/mm/yy"))
        .txtPrazo_Qualidade = IIf(IsNull(TBProduto!Prazo_qualidade), "", Format(TBProduto!Prazo_qualidade, "dd/mm/yy"))
        .txtPrazo_Compras = IIf(IsNull(TBProduto!Prazo_compras), "", Format(TBProduto!Prazo_compras, "dd/mm/yy"))
    End If
    TBProduto.Close
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluir()
On Error GoTo tratar_erro
    
If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

With frmVendas_analise
    If .Txt_status = "REVISADA" Then
        USMsgBox ("Não é permitido excluir os prazos da análise crítica revisada."), vbExclamation, "CAPRIND v5.0"
        Exit Sub
    End If
    
    Set TBGravar = CreateObject("adodb.recordset")
    TBGravar.Open "Select Status from vendas_analise where ID = " & .txtID, Conexao, adOpenKeyset, adLockOptimistic
    If TBGravar.EOF = False Then
        If TBGravar!status = "APROVADA" And .Txt_status = "APROVADA" Then
            USMsgBox ("Não é permitido o excluir os prazos desta análise, pois a mesma já está aprovada."), vbExclamation, "CAPRIND v5.0"
            TBGravar.Close
            Exit Sub
        End If
        If TBGravar!status = "APROVADA" And .Txt_status <> "APROVADA" Then
            Set TBFIltro = CreateObject("adodb.recordset")
            TBFIltro.Open "Select IDAnalise from vendas_carteira where IDAnalise = " & .txtID, Conexao, adOpenKeyset, adLockOptimistic
            If TBFIltro.EOF = False Then
                USMsgBox ("Não é permitido o excluir os prazos desta análise, pois a mesma está vinculada a uma proposta."), vbExclamation, "CAPRIND v5.0"
                TBFIltro.Close
                TBGravar.Close
                Exit Sub
            End If
            TBFIltro.Close
        End If
    End If
    TBGravar.Close
    
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select Prazo_engenharia, Prazo_processos, Prazo_pcp, Prazo_qualidade, Prazo_compras from Vendas_analise where ID = " & .txtID & "", Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
        TBProduto!Prazo_engenharia = Null
        TBProduto!Prazo_processos = Null
        TBProduto!Prazo_pcp = Null
        TBProduto!Prazo_qualidade = Null
        TBProduto!Prazo_compras = Null
        TBProduto.Update
        txtEngenharia = "__/__/____"
        txtProcesso = "__/__/____"
        txtPCP = "__/__/____"
        txtQualidade = "__/__/____"
        txtCompras = "__/__/____"
        
        .txtPrazo_Engenharia = ""
        .txtPrazo_Processo = ""
        .txtPrazo_PCP = ""
        .txtPrazo_Qualidade = ""
        .txtPrazo_Compras = ""
    End If
    TBProduto.Close
    USMsgBox ("Cadastro excluído com sucesso."), vbInformation, "CAPRIND v5.0"
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro
  
Select Case KeyCode
    Case vbKeyF3: ProcSalvar
    Case vbKeyF4: ProcExcluir
    Case vbKeyEscape: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 7455, 6, True
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select Prazo_engenharia, Prazo_processos, Prazo_pcp, Prazo_qualidade, Prazo_compras from Vendas_analise where ID = " & frmVendas_analise.txtID & "", Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    txtEngenharia = IIf(IsNull(TBProduto!Prazo_engenharia), "__/__/____", TBProduto!Prazo_engenharia)
    txtProcesso = IIf(IsNull(TBProduto!Prazo_processos), "__/__/____", TBProduto!Prazo_processos)
    txtPCP = IIf(IsNull(TBProduto!Prazo_pcp), "__/__/____", TBProduto!Prazo_pcp)
    txtQualidade = IIf(IsNull(TBProduto!Prazo_qualidade), "__/__/____", TBProduto!Prazo_qualidade)
    txtCompras = IIf(IsNull(TBProduto!Prazo_compras), "__/__/____", TBProduto!Prazo_compras)
End If
TBProduto.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub imgCompras_Click()
On Error GoTo tratar_erro

Faturamento = False
Compras_Pedido = False
Compras_Requisicao = False
Compras_Fallow_up = False
Vendas_Carteira = False
Vendas_Proposta = False
Vendas_PI = False
Manutencao = False
Compras_Cotacao = False
Usuarios = False
Inspecao_recebimento = False
Funcionario = False
RNC = False
SolicitacaoAcao = False
Troca_Duplicata = False
Financeiro_Contas_Recebidas = False
Engenharia_Normas = False
Qualidade_PPAP_PSW = False
Qualidade_PPAP_Plano = False
Qualidade_PPAP_FMEA = False
Qualidade_sistema = False
Engenharia = False
Compras_Fornecedores = False
Vendas_Programacao = False
Vendas_Analise = True
Estoque_recebimento = False
Sit_Data = 5
FrmCalendario.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub imgEngenharia_Click()
On Error GoTo tratar_erro

Faturamento = False
Compras_Pedido = False
Compras_Requisicao = False
Compras_Fallow_up = False
Vendas_Carteira = False
Vendas_Proposta = False
Vendas_PI = False
Manutencao = False
Compras_Cotacao = False
Usuarios = False
Inspecao_recebimento = False
Funcionario = False
RNC = False
SolicitacaoAcao = False
Troca_Duplicata = False
Financeiro_Contas_Recebidas = False
Engenharia_Normas = False
Qualidade_PPAP_PSW = False
Qualidade_PPAP_Plano = False
Qualidade_PPAP_FMEA = False
Qualidade_sistema = False
Engenharia = False
Compras_Fornecedores = False
Vendas_Programacao = False
Vendas_Analise = True
Estoque_recebimento = False
Sit_Data = 1
FrmCalendario.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub imgPCP_Click()
On Error GoTo tratar_erro

Faturamento = False
Compras_Pedido = False
Compras_Requisicao = False
Compras_Fallow_up = False
Vendas_Carteira = False
Vendas_Proposta = False
Vendas_PI = False
Manutencao = False
Compras_Cotacao = False
Usuarios = False
Inspecao_recebimento = False
Funcionario = False
RNC = False
SolicitacaoAcao = False
Troca_Duplicata = False
Financeiro_Contas_Recebidas = False
Engenharia_Normas = False
Qualidade_PPAP_PSW = False
Qualidade_PPAP_Plano = False
Qualidade_PPAP_FMEA = False
Qualidade_sistema = False
Engenharia = False
Compras_Fornecedores = False
Vendas_Programacao = False
Vendas_Analise = True
Estoque_recebimento = False
Sit_Data = 3
FrmCalendario.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub imgProcesso_Click()
On Error GoTo tratar_erro

Faturamento = False
Compras_Pedido = False
Compras_Requisicao = False
Compras_Fallow_up = False
Vendas_Carteira = False
Vendas_Proposta = False
Vendas_PI = False
Manutencao = False
Compras_Cotacao = False
Usuarios = False
Inspecao_recebimento = False
Funcionario = False
RNC = False
SolicitacaoAcao = False
Troca_Duplicata = False
Financeiro_Contas_Recebidas = False
Engenharia_Normas = False
Qualidade_PPAP_PSW = False
Qualidade_PPAP_Plano = False
Qualidade_PPAP_FMEA = False
Qualidade_sistema = False
Engenharia = False
Compras_Fornecedores = False
Vendas_Programacao = False
Vendas_Analise = True
Estoque_recebimento = False
Sit_Data = 2
FrmCalendario.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub imgQualidade_Click()
On Error GoTo tratar_erro

Faturamento = False
Compras_Pedido = False
Compras_Requisicao = False
Compras_Fallow_up = False
Vendas_Carteira = False
Vendas_Proposta = False
Vendas_PI = False
Manutencao = False
Compras_Cotacao = False
Usuarios = False
Inspecao_recebimento = False
Funcionario = False
RNC = False
SolicitacaoAcao = False
Troca_Duplicata = False
Financeiro_Contas_Recebidas = False
Engenharia_Normas = False
Qualidade_PPAP_PSW = False
Qualidade_PPAP_Plano = False
Qualidade_PPAP_FMEA = False
Qualidade_sistema = False
Engenharia = False
Compras_Fornecedores = False
Vendas_Programacao = False
Vendas_Analise = True
Estoque_recebimento = False
Sit_Data = 4
FrmCalendario.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtCompras_LostFocus()
On Error GoTo tratar_erro

If txtCompras <> "__/__/____" Then
    VerifData = txtCompras
    ProcVerificaData
    If VerifData = False Then
        txtCompras.Text = "__/__/____"
        txtCompras.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtEngenharia_LostFocus()
On Error GoTo tratar_erro

If txtEngenharia.Text <> "__/__/____" Then
    VerifData = txtEngenharia.Text
    ProcVerificaData
    If VerifData = False Then
        txtEngenharia.Text = "__/__/____"
        txtEngenharia.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtPCP_LostFocus()
On Error GoTo tratar_erro

If txtPCP <> "__/__/____" Then
    VerifData = txtPCP
    ProcVerificaData
    If VerifData = False Then
        txtPCP.Text = "__/__/____"
        txtPCP.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtProcesso_LostFocus()
On Error GoTo tratar_erro

If txtProcesso <> "__/__/____" Then
    VerifData = txtProcesso
    ProcVerificaData
    If VerifData = False Then
        txtProcesso.Text = "__/__/____"
        txtProcesso.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtQualidade_LostFocus()
On Error GoTo tratar_erro

If txtQualidade <> "__/__/____" Then
    VerifData = txtQualidade
    ProcVerificaData
    If VerifData = False Then
        txtQualidade.Text = "__/__/____"
        txtQualidade.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcSalvar
    Case 2: ProcExcluir
    'Case 4: ProcAjuda
    Case 5: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
