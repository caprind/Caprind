VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmPlano_producao_localizar_OS_adicionar 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Adicionar OS no plano de produção"
   ClientHeight    =   1845
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10545
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   10545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   825
      Left            =   60
      TabIndex        =   5
      Top             =   990
      Width           =   10425
      Begin VB.TextBox Txt_OS 
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
         TabIndex        =   0
         TabStop         =   0   'False
         ToolTipText     =   "Número da OS."
         Top             =   375
         Width           =   1125
      End
      Begin VB.TextBox Txt_qtde 
         Alignment       =   1  'Right Justify
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
         Left            =   7770
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade."
         Top             =   375
         Width           =   1215
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
         Left            =   2790
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Descrição."
         Top             =   375
         Width           =   4965
      End
      Begin VB.TextBox Txt_codigo_interno 
         Alignment       =   2  'Center
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
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Código interno."
         Top             =   375
         Width           =   1455
      End
      Begin VB.TextBox Txt_qtde_adicionar 
         Alignment       =   1  'Right Justify
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
         Left            =   9000
         TabIndex        =   3
         ToolTipText     =   "Quantidade à adicionar."
         Top             =   375
         Width           =   1215
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantidade"
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
         Left            =   7950
         TabIndex        =   11
         Top             =   180
         Width           =   840
      End
      Begin VB.Label Label6 
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
         Height          =   195
         Left            =   4927
         TabIndex        =   10
         ToolTipText     =   "Quantidade à liberada."
         Top             =   180
         Width           =   690
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código interno"
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
         Left            =   1515
         TabIndex        =   9
         ToolTipText     =   "Quantidade à liberada."
         Top             =   180
         Width           =   1050
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OS"
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
         Left            =   630
         TabIndex        =   8
         Top             =   180
         Width           =   225
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qtde. adic."
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
         Left            =   9150
         TabIndex        =   6
         ToolTipText     =   "Quantidade à liberada."
         Top             =   180
         Width           =   915
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   7
      Top             =   0
      Width           =   10425
      _ExtentX        =   18389
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
      ButtonCaption1  =   "Adicionar"
      ButtonEnabled1  =   0   'False
      ButtonIconSize1 =   32
      ButtonToolTipText1=   "Adicionar (F3)"
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
      ButtonWidth1    =   52
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
      ButtonLeft2     =   56
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
      ButtonLeft3     =   60
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
      ButtonLeft4     =   98
      ButtonTop4      =   2
      ButtonWidth4    =   26
      ButtonHeight4   =   21
      ButtonUseMaskColor4=   0   'False
      ButtonEnabled5  =   0   'False
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
      ButtonLeft5     =   126
      ButtonTop5      =   2
      ButtonWidth5    =   24
      ButtonHeight5   =   24
      ButtonUseMaskColor5=   0   'False
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   2550
         Top             =   150
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmPlano_producao_localizar_OS_adicionar.frx":0000
         Count           =   1
      End
   End
End
Attribute VB_Name = "frmPlano_producao_localizar_OS_adicionar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ProcAdicionar()
On Error GoTo tratar_erro

Acao = "adicionar"
valor = Txt_qtde
Valor1 = IIf(Txt_qtde_adicionar = "", 0, Txt_qtde_adicionar)
If Valor1 <= 0 Then
    NomeCampo = "a quantidade à adicionar"
    ProcVerificaAcao
    Txt_qtde_adicionar.SetFocus
    Exit Sub
End If
If Valor1 > valor Then
    USMsgBox ("A quantidade à adicionar não pode ser maior que a quantidade da OS."), vbExclamation, "CAPRIND v5.0"
    Txt_qtde_adicionar.SetFocus
    Exit Sub
End If

Set TBOS = CreateObject("adodb.recordset")
TBOS.Open "Select * from Ordemservico where IDproducao = " & IDlista, Conexao, adOpenKeyset, adLockOptimistic
If TBOS.EOF = False Then
    If Valor1 < valor Then
        'Cria nova OS
        Set TBGravar = CreateObject("adodb.recordset")
        TBGravar.Open "Select * from Ordemservico", Conexao, adOpenKeyset, adLockOptimistic
        TBGravar.AddNew
        TBGravar!Pronto = "NÃO"
        TBGravar!status = Aguardando
        TBGravar!Fase = TBOS!Fase
        TBGravar!Rev_Fase = TBOS!Rev_Fase
        TBGravar!Grupo_op = TBOS!Grupo_op
        TBGravar!IDFase = TBOS!IDFase
        TBGravar!maquina = TBOS!maquina
        TBGravar!quantidade = Valor1
        
        TBOS!quantidade = TBOS!quantidade - Valor1
        
        Set TBFases = CreateObject("adodb.recordset")
        TBFases.Open "Select TESegundos, TPSegundos from Fases where IDFase = " & TBOS!IDFase, Conexao, adOpenKeyset, adLockOptimistic
        If TBFases.EOF = False Then
            DecimoSegundos = (IIf(IsNull(TBFases!TESegundos), 0, TBFases!TESegundos) * TBGravar!quantidade) + IIf(IsNull(TBFases!TPSegundos), 0, TBFases!TPSegundos)
            TBGravar!TTLPREVS = DecimoSegundos 'Tempo total do lote previsto em segundos
            TBGravar!TempoTotalLote = FormataTempo(DecimoSegundos) 'Tempo total do lote previsto
            
            DecimoSegundos = (IIf(IsNull(TBFases!TESegundos), 0, TBFases!TESegundos) * TBOS!quantidade) + IIf(IsNull(TBFases!TPSegundos), 0, TBFases!TPSegundos)
            TBOS!TTLPREVS = DecimoSegundos 'Tempo total do lote previsto em segundos
            TBOS!TempoTotalLote = FormataTempo(DecimoSegundos) 'Tempo total do lote previsto
        End If
        TBFases.Close
        
        TBGravar!custos = TBOS!custos
        TBGravar!Valor_hs_prep = TBOS!Valor_hs_prep
        TBGravar!Valor_hs_exec = TBOS!Valor_hs_exec
        TBGravar!IDPROCESSO = TBOS!IDPROCESSO
        TBGravar!Ordem = TBOS!Ordem
        TBGravar!pecahora = TBOS!pecahora
        TBGravar!Pcshora = TBOS!Pcshora
        TBGravar!pc_te = TBOS!pc_te
        TBGravar!Preparacao = TBOS!Preparacao
        TBGravar!Execucao = TBOS!Execucao
        TBGravar!TempoPreparacao = TBOS!TempoPreparacao
        TBGravar!TempoExecucao = TBOS!TempoExecucao
        TBGravar!descfase = TBOS!descfase
        TBGravar!TESegundos = TBOS!TESegundos
        TBGravar!OSControlada = TBOS!OSControlada
        TBGravar!Processo_controlado = TBOS!Processo_controlado
        TBGravar!IDPlano = TBOS!IDPlano
        TBGravar!PrazoFinal = TBOS!PrazoFinal
        TBGravar!prazofinalinicio = TBOS!prazofinalinicio
        TBGravar.Update
        TBOS.Update
        Conexao.Execute "Update Ordemservico Set ID_apontamento = " & frmPlano_producao.Txt_ID & " where IDproducao = " & TBGravar!IDProducao
        ID_documento = TBGravar!IDProducao
        Documento1 = "OS: " & TBGravar!IDProducao
        TBGravar.Close
        TBOS.Close
    Else
        Conexao.Execute "Update Ordemservico Set ID_apontamento = " & frmPlano_producao.Txt_ID & " where IDproducao = " & IDlista
        ID_documento = IDlista
        Documento1 = "OS: " & IDlista
    End If
    USMsgBox ("OS adicionada com sucesso."), vbInformation, "CAPRIND v5.0"
    '==================================
    Modulo = "PCP/Plano da produção"
    Evento = "Nova OS"
    Documento = "Nº plano: " & frmPlano_producao.Txt_numero_plano
    ProcGravaEvento
    '==================================
End If
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

Permitido2 = False
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro
    
Select Case KeyCode
    Case vbKeyF3: ProcAdicionar
    'Case vbKeyF1: ProcAjuda
    Case vbKeyEsc: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 10425, 5, True
Set TBOS = CreateObject("adodb.recordset")
TBOS.Open "Select P.Desenho, P.Produto, OS.Quantidade from Ordemservico OS INNER JOIN Producao P ON P.Ordem = OS.Ordem where OS.IDproducao = " & IDlista, Conexao, adOpenKeyset, adLockOptimistic
If TBOS.EOF = False Then
    Txt_OS = IDlista
    Txt_codigo_interno = IIf(IsNull(TBOS!Desenho), "", TBOS!Desenho)
    Txt_descricao = IIf(IsNull(TBOS!Produto), "", TBOS!Produto)
    Txt_qtde = IIf(IsNull(TBOS!quantidade), 0, TBOS!quantidade)
    Txt_qtde_adicionar = IIf(IsNull(TBOS!quantidade), 0, TBOS!quantidade)
End If
TBOS.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_qtde_adicionar_Change()
On Error GoTo tratar_erro

If Txt_qtde_adicionar <> "" Then
    VerifNumero = Txt_qtde_adicionar
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_qtde_adicionar = ""
        Txt_qtde_adicionar.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_qtde_adicionar_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus Txt_qtde_adicionar

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcAdicionar
    'Case 3: ProcAjuda
    Case 4: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
