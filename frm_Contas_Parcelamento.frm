VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frm_Contas_parcelamento 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Contas a pagar - Copiar"
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3315
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
   Icon            =   "frm_Contas_Parcelamento.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   3315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   2460
      Top             =   180
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frm_Contas_Parcelamento.frx":1042
      Count           =   1
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1665
      Left            =   55
      TabIndex        =   4
      Top             =   990
      Width           =   3210
      Begin VB.TextBox txtvalor 
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
         Height          =   315
         Left            =   1605
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Valor das parcelas."
         Top             =   870
         Width           =   1395
      End
      Begin VB.TextBox txtprazo 
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
         Left            =   1605
         TabIndex        =   1
         ToolTipText     =   "Número de parcelas."
         Top             =   525
         Width           =   1395
      End
      Begin VB.CheckBox chk_fixar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fixar dia de pagamento"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   630
         TabIndex        =   3
         Top             =   1320
         Width           =   2385
      End
      Begin MSComCtl2.DTPicker txtinicio 
         Height          =   315
         Left            =   1605
         TabIndex        =   0
         ToolTipText     =   "Data de início do pagamento."
         Top             =   180
         Width           =   1395
         _ExtentX        =   2461
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
         Format          =   133496833
         CurrentDate     =   39057
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         X1              =   30
         X2              =   3180
         Y1              =   1260
         Y2              =   1260
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Início pagamento :"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   195
         MousePointer    =   1  'Arrow
         TabIndex        =   7
         Top             =   180
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Nº parcelas :*"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   510
         MousePointer    =   1  'Arrow
         TabIndex        =   6
         Top             =   525
         Width           =   1020
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Valor :"
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
         Left            =   1005
         MousePointer    =   1  'Arrow
         TabIndex        =   5
         Top             =   870
         Width           =   525
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   55
      TabIndex        =   8
      Top             =   0
      Width           =   3210
      _ExtentX        =   5662
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
      ButtonCaption1  =   "Copiar"
      ButtonEnabled1  =   0   'False
      ButtonIconSize1 =   32
      ButtonToolTipText1=   "Copiar (F3)"
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
      ButtonWidth1    =   44
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
      ButtonLeft2     =   48
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
      ButtonLeft3     =   52
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
      ButtonLeft4     =   95
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
      ButtonLeft5     =   127
      ButtonTop5      =   2
      ButtonWidth5    =   24
      ButtonHeight5   =   24
      ButtonUseMaskColor5=   0   'False
   End
End
Attribute VB_Name = "frm_Contas_parcelamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro
  
Select Case KeyCode
    Case vbKeyF3: ProcGerarPagtos
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

ProcCarregaToolBar1 Me, 3210, 5, True

With frmContas_Pagar
    txtinicio.Value = .txtDtpagto
    txtValor.Text = Format(.txtValorTotal, "###,##0.00")
End With
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcGerarPagtos()
On Error GoTo tratar_erro

valor = IIf(txtprazo = "", 0, txtprazo)
If valor <= 0 Then
    USMsgBox ("Informe o prazo/parcela do pagamento antes de copiar."), vbExclamation, "CAPRIND v5.0"
    txtprazo.SetFocus
    Exit Sub
End If
PAGTO = (txtinicio.Value)
With frmContas_Pagar.txtparcela
    .PromptInclude = False
    Controle = Left(.Text, 3)
    .PromptInclude = True
End With
'Insere Contas programadas
With frmContas_Pagar.lst_contas
    For InitFor = 1 To valor
        If chk_fixar.Value = 1 Then
            DT = PAGTO
            DiaX = Day(DT)
            MesX = Month(DT)
            AnoX = Year(DT)
        Else
            DT = PAGTO
            DiaX = Day(DT)
            MesX = Month(DT)
            AnoX = Year(DT)
            If Weekday(DT) = vbSunday Then
                Dataini = DT
                Dataini = Dataini + 1
                DT = Dataini
            End If
            If Weekday(DT) = vbSaturday Then
                Dataini = DT
                Dataini = Dataini + 2
                DT = Dataini
            End If
        End If
        With frmContas_Pagar
            Set TBContas = CreateObject("adodb.recordset")
            TBContas.Open "Select * from tbl_contaspagar where IDintconta = " & .txtidintconta, Conexao, adOpenKeyset, adLockOptimistic
            If TBContas.EOF = False Then
                Set TBProduto = CreateObject("adodb.recordset")
                TBProduto.Open "Select * from tbl_contaspagar", Conexao, adOpenKeyset, adLockOptimistic
                TBProduto.AddNew
                TBProduto!Antecipacao = TBContas!Antecipacao
                If TBContas!Antecipacao = True Then TBProduto!Saldo_antecipacao = txtValor
                TBProduto!Devolucao = TBContas!Devolucao
                TBProduto!Data_transacao = Date
                TBProduto!Class_conta = TBContas!Class_conta
                TBProduto!txt_ndocumento = TBContas!txt_ndocumento
                TBProduto!Txt_pedido = TBContas!Txt_pedido
                TBProduto!dt_Pagamento = DT
                TBProduto!Dt_emissao = TBContas!Dt_emissao
                TBProduto!int_codforn = TBContas!int_codforn
                TBProduto!dbl_valorpagto = txtValor
                
                Controle = Controle + 1
                Par1 = Controle
                .txtparcela.PromptInclude = False
                Par2 = Right(.txtparcela.Text, 3)
                .txtparcela.PromptInclude = True
                If Len(Par1) = 1 Then
                    Par1 = "00" & Par1
                ElseIf Len(Par1) = 2 Then
                    Par1 = "0" & Par1
                End If
                If Len(Par2) = 1 Then
                    Par2 = "00" & Par2
                ElseIf Len(Par2) = 2 Then
                    Par2 = "0" & Par2
                End If
                If Par1 <= Par2 Then TBProduto!txt_Parcela = Par1 & "/" & Par2 Else TBProduto!txt_Parcela = Par2 & "/" & Par2
                            
                TBProduto!Banco = TBContas!Banco
                TBProduto!FormaBaixa = TBContas!FormaBaixa
                TBProduto!txt_observacoes = TBContas!txt_observacoes
                TBProduto!Txt_fornecedor = TBContas!Txt_fornecedor
                TBProduto!Logsit = "N"
                TBProduto!status = "TÍTULO EM ABERTO"
                TBProduto!Responsavel = pubUsuario
                TBProduto!ID_empresa = TBContas!ID_empresa
                TBProduto!Bloqueado = False
                TBProduto!Tipo = TBContas!Tipo
                TBProduto.Update
            End If
            TBContas.Close
            
            'Fluxo de Caixa
            Set TBFluxo = CreateObject("adodb.recordset")
            TBFluxo.Open "Select * from tbl_Fluxo_de_caixa where IDFluxo = " & IIf(IsNull(TBProduto!IDFluxo), 0, TBProduto!IDFluxo), Conexao, adOpenKeyset, adLockOptimistic
            If TBFluxo.EOF = True Then TBFluxo.AddNew
            TBFluxo!IDintconta = TBProduto!IDintconta
            TBFluxo!Operacao = "À Debitar"
            TBFluxo!Data = DT
            TBFluxo!valor = txtValor
            TBFluxo!Descricao = .txtFornec
            TBFluxo!status = "N"
            TBFluxo!int_NotaFiscal = .txtNDocumento
            TBFluxo!Bloqueado = False
            TBFluxo!ID_empresa = .Cmb_empresa.ItemData(.Cmb_empresa.ListIndex)
            TBFluxo!Instituicao = TBProduto!Banco
            TBFluxo.Update
            Conexao.Execute "UPDATE tbl_contaspagar set IDFluxo = " & TBFluxo!IDFluxo & " where IdIntConta = " & TBProduto!IDintconta
            TBFluxo.Close
                        
            'Conta contábil
            If .txtidintconta <> "" Then
                Set TBCiclo = CreateObject("adodb.recordset")
                TBCiclo.Open "select * from familia_financeiro where IDconta = " & .txtidintconta & " and TipoConta = 'P' and Deposito_transf = 'False'", Conexao, adOpenKeyset, adLockOptimistic
                Do While TBCiclo.EOF = False
                    Set TBFamilia = CreateObject("adodb.recordset")
                    TBFamilia.Open "Select * from familia_financeiro", Conexao, adOpenKeyset, adLockOptimistic
                    TBFamilia.AddNew
                    TBFamilia!ID_PC = TBCiclo!ID_PC
                    TBFamilia!IDConta = TBProduto!IDintconta
                    TBFamilia!TipoConta = TBCiclo!TipoConta
                    TBFamilia!valor = TBCiclo!valor
                    TBFamilia.Update
                    TBFamilia.Close
                    TBCiclo.MoveNext
                Loop
                
                'Centro de custo
                Set TBCiclo = CreateObject("adodb.recordset")
                TBCiclo.Open "Select * from CC_realizado where ID_financeiro = " & .txtidintconta & " Order by ID", Conexao, adOpenKeyset, adLockOptimistic
                Do While TBCiclo.EOF = False
                    Set TBFamilia = CreateObject("adodb.recordset")
                    TBFamilia.Open "Select * from CC_realizado", Conexao, adOpenKeyset, adLockOptimistic
                    TBFamilia.AddNew
                    ProcEnviaDadosCCRealizado TBCiclo!ID_CC
                    TBFamilia.Update
                    TBFamilia.Close
                    
                    'Grava movimentação no centro consolidado
                    Set TBAfericao = CreateObject("adodb.recordset")
                    TBAfericao.Open "Select * from Usuarios_Setor_Consolidacao where ID_CC_consolidado = " & TBCiclo!ID_CC, Conexao, adOpenKeyset, adLockOptimistic
                    If TBAfericao.EOF = False Then
                        Do While TBAfericao.EOF = False
                            Set TBFamilia = CreateObject("adodb.recordset")
                            TBFamilia.Open "Select * from CC_realizado", Conexao, adOpenKeyset, adLockOptimistic
                            TBFamilia.AddNew
                            ProcEnviaDadosCCRealizado TBAfericao!ID_CC
                            TBFamilia.Update
                            TBFamilia.Close
                            
                            Set TBFI = CreateObject("adodb.recordset")
                            TBFI.Open "Select * from Usuarios_Setor_Consolidacao where ID_CC_consolidado = " & TBAfericao!ID_CC, Conexao, adOpenKeyset, adLockOptimistic
                            If TBFI.EOF = False Then
                                Do While TBFI.EOF = False
                                    Set TBFamilia = CreateObject("adodb.recordset")
                                    TBFamilia.Open "Select * from CC_realizado", Conexao, adOpenKeyset, adLockOptimistic
                                    TBFamilia.AddNew
                                    ProcEnviaDadosCCRealizado TBFI!ID_CC
                                    TBFamilia.Update
                                    TBFamilia.Close
                                    TBFI.MoveNext
                                Loop
                            End If
                            TBFI.Close
                            
                            TBAfericao.MoveNext
                        Loop
                    End If
                    TBAfericao.Close
                    
                    TBCiclo.MoveNext
                Loop
                TBCiclo.Close
            End If
                        
            '==================================
            Modulo = "Financeiro/Contas a pagar"
            Evento = "Novo"
            ID_documento = TBProduto!IDintconta
            Documento = "Documento: " & .txtNDocumento
            Documento1 = ""
            ProcGravaEvento
            '==================================

            TBProduto.Close
        End With
        MesX = MesX + 1
        If MesX > 12 Then
            AnoX = AnoX + 1
            MesX = 1
        End If
        If DiaX = 29 And MesX = 2 Then DiaX = 28
        If DiaX = 30 And MesX = 2 Then DiaX = 28
        If DiaX = 31 And MesX = 2 Then DiaX = 28
        If DiaX = 31 And MesX = 4 Then DiaX = 30
        If DiaX = 31 And MesX = 6 Then DiaX = 30
        If DiaX = 31 And MesX = 9 Then DiaX = 30
        If DiaX = 31 And MesX = 11 Then DiaX = 30
        PAGTO = Format(DiaX, "00") & "/" & Format(MesX, "00") & "/" & Format(AnoX, "0000")
    Next InitFor
End With
USMsgBox ("Conta copiada com sucesso."), vbInformation, "CAPRIND v5.0"
With frmContas_Pagar
    TextoFiltro = ""
    TextoFiltroRel = ""
    If .txtNPedido <> "" Then
        TextoFiltro = " and CP.txt_pedido = '" & .txtNPedido.Text & "'"
        TextoFiltroRel = " and {tbl_ContasPagar.txt_pedido} = '" & .txtNPedido.Text & "'"
    End If
    If .Chk_antecipacao.Value = 1 Then
        TextoFiltro = TextoFiltro & " and CP.Antecipacao = 'True'"
        TextoFiltroRel = TextoFiltroRel & " and {tbl_ContasPagar.Antecipacao} = True"
    End If
    If .Chk_devolucao.Value = 1 Then
        TextoFiltro = TextoFiltro & " and CP.Devolucao = 'True'"
        TextoFiltroRel = TextoFiltroRel & " and {tbl_ContasPagar.Devolucao} = True"
    End If
    .ProcConstruirFiltroPadrao "CP.int_codforn = " & .txtIDFornec & " and CP.txt_NDocumento = '" & .txtNDocumento & "'" & TextoFiltro, "{tbl_ContasPagar.int_codforn} = " & .txtIDFornec & " and {tbl_ContasPagar.txt_NDocumento} = '" & .txtNDocumento & "' " & TextoFiltroRel, False, False
    .lst_contas.ListItems.Clear
    .ProcCarregaLista (1)
    .Novo_Pagar = False
End With
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviaDadosCCRealizado(ID_CC As Long)
On Error GoTo tratar_erro

TBFamilia!ID_financeiro = TBProduto!IDintconta
TBFamilia!Data = DT
TBFamilia!Responsavel = pubUsuario
TBFamilia!ID_empresa = TBCiclo!ID_empresa
TBFamilia!Operacao = "Débito"
TBFamilia!ID_CC = ID_CC
TBFamilia!valor = TBCiclo!valor
TBFamilia!Percentual = TBCiclo!Percentual
TBFamilia!Bloqueado = False

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

Private Sub txtinicio_Change()
On Error GoTo tratar_erro

If Left(txtinicio.Value, 2) = "31" Or Left(txtinicio.Value, 5) = "30/01" Or Left(txtinicio.Value, 5) = "31/01" Then
    chk_fixar.Enabled = False
Else
    chk_fixar.Enabled = True
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtprazo_Change()
On Error GoTo tratar_erro

If txtprazo.Text <> "" Then
    VerifNumero = txtprazo.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtprazo.Text = ""
        txtprazo.SetFocus
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
    Case 1: ProcGerarPagtos
    'Case 3: ProcAjuda
    Case 4: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
