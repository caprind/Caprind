VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Begin VB.Form Gerar_receb 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Contas a receber - Parcelar recebimentos"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5250
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Gerar_receb.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   5250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   2835
      Left            =   55
      TabIndex        =   6
      Top             =   990
      Width           =   5160
      Begin VB.TextBox txtDTEmissao 
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
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         ToolTipText     =   "Data de emissão."
         Top             =   390
         Width           =   1095
      End
      Begin VB.CheckBox chk_fixar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fixar dia de recebimento"
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
         Left            =   2895
         TabIndex        =   4
         Top             =   810
         Width           =   2085
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1290
         TabIndex        =   1
         ToolTipText     =   "Número de parcelas."
         Top             =   390
         Width           =   975
      End
      Begin VB.TextBox txtdescricao 
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
         Height          =   1665
         Left            =   180
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         ToolTipText     =   "Obeservações."
         Top             =   1020
         Width           =   4770
      End
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3690
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Valor das parcelas."
         Top             =   390
         Width           =   1275
      End
      Begin MSComCtl2.DTPicker txtinicio 
         Height          =   315
         Left            =   2280
         TabIndex        =   2
         ToolTipText     =   "Início do recebimento."
         Top             =   390
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
         Format          =   183304195
         CurrentDate     =   39057
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Observações"
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
         Left            =   270
         TabIndex        =   11
         Top             =   810
         Width           =   945
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Início recto."
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
         Left            =   2550
         TabIndex        =   10
         Top             =   180
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Nº parcelas*"
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
         Left            =   1320
         TabIndex        =   9
         Top             =   180
         Width           =   915
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Valor"
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
         Left            =   4110
         TabIndex        =   8
         Top             =   180
         Width           =   435
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Dt. emissão"
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
         Left            =   307
         TabIndex        =   7
         Top             =   180
         Width           =   840
      End
   End
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   3330
      Top             =   150
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "Gerar_receb.frx":1042
      Count           =   1
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   55
      TabIndex        =   12
      Top             =   0
      Width           =   5160
      _ExtentX        =   9102
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
      ButtonCaption1  =   "Parcelar"
      ButtonEnabled1  =   0   'False
      ButtonIconSize1 =   32
      ButtonToolTipText1=   "Parcelar (F3)"
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
      ButtonWidth1    =   55
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
      ButtonLeft2     =   59
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
      ButtonLeft3     =   63
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
      ButtonLeft4     =   106
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
      ButtonLeft5     =   138
      ButtonTop5      =   2
      ButtonWidth5    =   24
      ButtonHeight5   =   24
      ButtonUseMaskColor5=   0   'False
   End
End
Attribute VB_Name = "Gerar_receb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro
  
Select Case KeyCode
    Case vbKeyF3: ProcGerarRectos
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

ProcCarregaToolBar1 Me, 5160, 5, True

With frmContas_Receber
    txtDTEmissao = .mskEmissao.Value
    txtinicio.Value = .mskVencimento.Value
    txtValor.Text = Format(.txtValor, "###,##0.00")
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

Private Sub ProcGerarRectos()
On Error GoTo tratar_erro

valor = IIf(txtprazo = "", 0, txtprazo)
If valor <= 0 Then
    USMsgBox ("Informe o número de parcelas antes de parcelar."), vbExclamation, "CAPRIND v5.0"
    txtprazo.SetFocus
    Exit Sub
End If

Controle = 0
PAGTO = (txtinicio.Value)
'Insere Contas programadas
With frmContas_Receber.Lista
    For InitFor = 1 To txtprazo
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
        With frmContas_Receber
            Set TBContas = CreateObject("adodb.recordset")
            TBContas.Open "Select * from tbl_contas_receber where IDintconta = " & .txtidintconta, Conexao, adOpenKeyset, adLockOptimistic
            If TBContas.EOF = False Then
                Set TBProduto = CreateObject("adodb.recordset")
                TBProduto.Open "Select * from tbl_contas_receber", Conexao, adOpenKeyset, adLockOptimistic
                TBProduto.AddNew
                TBProduto!Antecipacao = TBContas!Antecipacao
                If TBContas!Antecipacao = True Then TBProduto!Saldo_antecipacao = txtValor
                TBProduto!Devolucao = TBContas!Devolucao
                TBProduto!Data_transacao = TBContas!Data_transacao
                TBProduto!Tipo_doc = TBContas!Tipo_doc
                If Permitido = True Then
                    TBProduto!txt_ndocumento = TBContas!txt_ndocumento
                    TBProduto!NFiscal = TBContas!NFiscal
                End If
                TBProduto!Proposta = TBContas!Proposta
                TBProduto!Vencimento = DT
                TBProduto!emissao = TBContas!emissao
                TBProduto!IDCliente = TBContas!IDCliente
                TBProduto!Nome_Razao = TBContas!Nome_Razao
                TBProduto!Cidade = TBContas!Cidade
                TBProduto!Estado = TBContas!Estado
                TBProduto!Observacoes = TBContas!Observacoes
                TBProduto!Banco = TBContas!Banco
                TBProduto!FormaBaixa = TBContas!FormaBaixa
                TBProduto!valor = txtValor
                TBProduto!ValorExtenso = FunValorExtenso(txtValor)
                TBProduto!ID_empresa = TBContas!ID_empresa
                TBProduto!Bloqueado = False
                TBProduto!Parcial = False
                TBProduto!IDtrocatitulo = 0
                
                Controle = Controle + 1
                Par1 = Controle
                Par2 = txtprazo
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
                TBProduto!Parcela = Par1 & "/" & Par2
                
                TBProduto!Logsit = "N"
                TBProduto!status = "TÍTULO EM ABERTO"
                TBProduto!Responsavel = pubUsuario
                TBProduto!Tipo = TBContas!Tipo
                TBProduto.Update
            End If
            TBContas.Close
        
            'Fluxo de Caixa
            Set TBFluxo = CreateObject("adodb.recordset")
            TBFluxo.Open "Select * from tbl_Fluxo_de_caixa where IDFluxo = " & IIf(IsNull(TBProduto!IDFluxo), 0, TBProduto!IDFluxo), Conexao, adOpenKeyset, adLockOptimistic
            If TBFluxo.EOF = True Then TBFluxo.AddNew
            TBFluxo!IDintconta = TBProduto!IDintconta
            TBFluxo!Operacao = "À Creditar"
            TBFluxo!Data = DT
            TBFluxo!valor = txtValor
            TBFluxo!Descricao = .txtNome_Razao
            TBFluxo!status = "N"
            TBFluxo!int_NotaFiscal = IIf(.txtNFiscal = "", Null, .txtNFiscal)
            TBFluxo!Documento = .txtDocumento
            TBFluxo!Bloqueado = False
            TBFluxo!ID_empresa = .Cmb_empresa.ItemData(.Cmb_empresa.ListIndex)
            TBFluxo!Instituicao = TBProduto!Banco
            TBFluxo.Update
            Conexao.Execute "Update tbl_contas_receber set IDFLUXO = " & TBFluxo!IDFluxo & " where IDIntconta = " & TBProduto!IDintconta
            TBFluxo.Close
            
            'Conta contábil
            If .txtidintconta <> "" Then
                Set TBCiclo = CreateObject("adodb.recordset")
                TBCiclo.Open "select * from familia_financeiro where IDconta = " & .txtidintconta & " and TipoConta = 'R' and Deposito_transf = 'False'", Conexao, adOpenKeyset, adLockOptimistic
                Do While TBCiclo.EOF = False
                    Set TBFamilia = CreateObject("adodb.recordset")
                    TBFamilia.Open "Select * from familia_financeiro", Conexao, adOpenKeyset, adLockOptimistic
                    TBFamilia.AddNew
                    TBFamilia!ID_PC = TBCiclo!ID_PC
                    TBFamilia!IDConta = TBProduto!IDintconta
                    TBFamilia!TipoConta = TBCiclo!TipoConta
                    TBFamilia!valor = TBCiclo!valor / txtprazo
                    TBFamilia.Update
                    TBFamilia.Close
                    TBCiclo.MoveNext
                Loop
                TBCiclo.Close
            End If
            
            '==================================
            Modulo = "Financeiro/Contas a receber"
            Evento = "Novo"
            ID_documento = TBProduto!IDintconta
            Documento = "Documento: " & .txtDocumento
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
USMsgBox ("Conta parcelada com sucesso."), vbInformation, "CAPRIND v5.0"
With frmContas_Receber
    If USMsgBox("Deseja excluir a conta utilizada para o parcelamento?", vbYesNo, "CAPRIND v5.0") = vbYes Then
        Conexao.Execute "DELETE from FC from tbl_Fluxo_de_caixa FC INNER JOIN tbl_contas_receber CR ON FC.IDFluxo = CR.IDFluxo Where CR.idintconta = " & .txtidintconta
        Conexao.Execute "DELETE FROM Familia_financeiro WHERE IDconta = " & .txtidintconta & " and TipoConta = 'R' and Deposito_transf = 'False'"
        Conexao.Execute "DELETE FROM tbl_contas_receber WHERE idintconta = " & .txtidintconta
        '==================================
        Modulo = "Financeiro/Contas a receber"
        Evento = "Excluir"
        ID_documento = .txtidintconta
        Documento = "Documento: " & .txtNFiscal
        Documento1 = ""
        ProcGravaEvento
        '==================================
    End If
    
    TextoFiltro = ""
    TextoFiltroRel = ""
    If .txtProposta <> "" Then
        TextoFiltro = "and CR.Proposta = '" & .txtProposta.Text & "'"
        TextoFiltroRel = "and {tbl_Contas_receber.Proposta} = '" & .txtProposta.Text & "'"
    End If
    If .Chk_antecipacao.Value = 1 Then
        TextoFiltro = TextoFiltro & " and CR.Antecipacao = 'True'"
        TextoFiltroRel = TextoFiltroRel & " and {tbl_Contas_receber.Antecipacao} = True"
    End If
    If .Chk_devolucao.Value = 1 Then
        TextoFiltro = TextoFiltro & " and CR.Devolucao = 'True'"
        TextoFiltroRel = TextoFiltroRel & " and {tbl_Contas_receber.Devolucao} = True"
    End If
    .ProcConstruirFiltroPadrao "CR.idcliente = " & .txtIDcliente & " and CR.Nfiscal = '" & .txtNFiscal.Text & "' " & TextoFiltro, "{tbl_Contas_receber.idcliente} = " & .txtIDcliente & " and {tbl_Contas_receber.Nfiscal} = '" & .txtNFiscal.Text & "' " & TextoFiltroRel, False, False
    .ProcCarregaLista (1)
    .Novo_Receber = False
End With
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

If txtprazo.Text = "" Or txtprazo.Text = "0" Then
    txtprazo = ""
    Exit Sub
End If
VerifNumero = txtprazo.Text
ProcVerificaNumero
If VerifNumero = False Then
    txtprazo.Text = ""
    txtprazo.SetFocus
    Exit Sub
End If
ValorTotal = frmContas_Receber.txtValor.Text / txtprazo.Text
txtValor.Text = Format(ValorTotal, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcGerarRectos
    'Case 3: ProcAjuda
    Case 4: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
