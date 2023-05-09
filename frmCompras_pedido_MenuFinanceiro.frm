VERSION 5.00
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Begin VB.Form frmCompras_pedido_MenuFinanceiro 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Compras - Pedido de compra"
   ClientHeight    =   1140
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3645
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1140
   ScaleWidth      =   3645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
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
      ForeColor       =   &H00000000&
      Height          =   1095
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   3525
      Begin DrawSuite2022.USButton cmdFinanceiro 
         Height          =   360
         Left            =   180
         TabIndex        =   1
         Top             =   180
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   635
         Caption         =   "Enviar para o financeiro"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   8421504
         BorderColorDisabled=   0
         BorderColorDown =   15048022
         BorderColorOver =   15381630
         GradientColor2  =   14737632
         GradientColor3  =   12632256
         GradientColor4  =   12632256
         PicSizeH        =   48
         PicSizeW        =   48
         Theme           =   1
      End
      Begin DrawSuite2022.USButton cmdFinanceiro_porcentagem 
         Height          =   360
         Left            =   180
         TabIndex        =   2
         Top             =   630
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   635
         Caption         =   "Enviar para o financeiro em porcentagem"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   8421504
         BorderColorDisabled=   0
         BorderColorDown =   15048022
         BorderColorOver =   15381630
         GradientColor2  =   14737632
         GradientColor3  =   12632256
         GradientColor4  =   12632256
         PicSizeH        =   48
         PicSizeW        =   48
         Theme           =   1
      End
   End
End
Attribute VB_Name = "frmCompras_pedido_MenuFinanceiro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdFinanceiro_Click()
On Error GoTo tratar_erro

If FunVerifDados = False Then Exit Sub
With frmCompras_Pedido
    If USMsgBox("Deseja realmente gerar contas a pagar do pedido de compra n° " & .txtPedido & "?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
    
    QtdeSaida = 0
    Contador = 1
    Contador2 = 0
    Controle = 1
    nPagto = 0
    Valor_Duplicatas = 0
    If Not IsNumeric(Left(.cmbpagamento, 1)) Then
        USMsgBox ("O campo condições de pagamento tem que ser em dias, favor alterar."), vbExclamation, "CAPRIND v5.0"
        Unload Me
        Exit Sub
    End If
    
    QtdeSaida = Len(.cmbpagamento)
            
    TextoCond = ""
    Do While Contador <= QtdeSaida
        If Mid(.cmbpagamento, Contador, 1) = "/" Or Mid(.cmbpagamento, Contador, 1) = "," Or IsNumeric(Mid(.cmbpagamento, Contador, 1)) = True Then
            If TextoCond = "" Then TextoCond = Mid(.cmbpagamento, Contador, 1) Else TextoCond = TextoCond & Mid(.cmbpagamento, Contador, 1)
        End If
        Contador = Contador + 1
    Loop
    
    If .cmbMoeda <> "" And .cmbMoeda <> "REAL" Then
Mensagem1:
        reposicao = InputBox("Favor informar o valor do " & .cmbMoeda & ".")
        If reposicao = "" Then Exit Sub
        If IsNumeric(reposicao) = False Then
            USMsgBox ("Só é permitido número neste campo."), vbExclamation, "CAPRIND v5.0"
            GoTo Mensagem1
        End If
        Qtd = reposicao
        If Qtd <= 0 Then
            USMsgBox ("So é permitido número maior que 0."), vbExclamation, "CAPRIND v5.0"
            GoTo Mensagem1
        End If
    Else
        Qtd = 1
    End If
    
    'Verifica qtde. de parcelas
    Contador = 1
    QtdeSaida = Len(TextoCond)
    Do While Contador <= QtdeSaida
       Do While Mid(TextoCond, Contador, 1) <> "/" And Contador <= QtdeSaida
            Contador2 = Contador2 + 1
            Contador = Contador + 1
        Loop
        nPagto = nPagto + 1
        Contador = Contador + 1
    Loop
        
    TotalProduto = .txtTotalPedido
    mxValorPag = Format(TotalProduto / nPagto, "###,##0.00")
    
    Contador = 1
    Contador3 = 1
    Dataini = Date
    Controle = 0
    Do While Contador <= QtdeSaida
        
        Contador2 = 0
        Do While Mid(.cmbpagamento, Contador, 1) <> "/" And Contador <= QtdeSaida
            Contador2 = Contador2 + 1
            Contador = Contador + 1
        Loop
        
        mxCondpag = ReturnNumbersOnly(Mid(.cmbpagamento, Contador3, Contador2))
        Contador3 = Contador3 + Contador2 + 1
        
        Controle = Controle + 1
        Par1 = Controle
        Par2 = nPagto
        If Len(Par1) = 1 Then
            Par1 = "00" & Par1
        ElseIf Len(Par1) = 1 Then
                Par1 = "0" & Par1
        End If
        If Len(Par2) = 1 Then
            Par2 = "00" & Par2
        ElseIf Len(Par2) = 1 Then
            Par2 = "0" & Par2
        End If
        
        Set TBReceber = CreateObject("adodb.recordset")
        TBReceber.Open "Select * from tbl_ContasPagar where txt_pedido = '" & .txtPedido & "' and txt_parcela = '" & Par1 & "/" & Par2 & "' order by dt_pagamento", Conexao, adOpenKeyset, adLockOptimistic
        If TBReceber.EOF = True Then
            TBReceber.AddNew
            TBReceber!Data_transacao = Date
            TBReceber!Parcial = False
            TBReceber!impresso = False
            TBReceber!Bloqueado = False
            TBReceber!Logsit = "N"
            TBReceber!Despesas_NF = False
            TBReceber!Antecipacao = False
            TBReceber!Devolucao = False
            TBReceber!status = "TÍTULO EM ABERTO"
            TBReceber!Responsavel = pubUsuario
            TBReceber!ID_nota = 0
            TBReceber!txt_ndocumento = ""
        End If
        TBReceber!dt_Pagamento = Dataini + mxCondpag
        TBReceber!Txt_pedido = .txtPedido.Text
        TBReceber!Dt_emissao = Date
        TBReceber!int_codforn = .txtIDfornecedor.Text
        TBReceber!txt_condpag = ""
        TBReceber!dbl_valorpagto = mxValorPag * Qtd
        TBReceber!txt_Parcela = Par1 & "/" & Par2
        TBReceber!Txt_fornecedor = .txtFornecedor.Text
        TBReceber!ID_empresa = .Cmb_empresa.ItemData(.Cmb_empresa.ListIndex)
        TBReceber!Tipo = "FO"
        If .chkObs_Financeiro.Value = 1 Then TBReceber!txt_observacoes = .txtObservacoes Else TBReceber!txt_observacoes = Null
        TBReceber.Update
        
        'Fluxo de Caixa
        Set TBFluxo = CreateObject("adodb.recordset")
        TBFluxo.Open "Select * from tbl_Fluxo_de_caixa where IDFluxo = " & IIf(IsNull(TBReceber!IDFluxo), 0, TBReceber!IDFluxo), Conexao, adOpenKeyset, adLockOptimistic
        If TBFluxo.EOF = True Then TBFluxo.AddNew
        TBFluxo!Operacao = "À Debitar"
        TBFluxo!Data = TBReceber!dt_Pagamento
        TBFluxo!valor = TBReceber!dbl_valorpagto
        TBFluxo!Descricao = TBReceber!Txt_fornecedor
        TBFluxo!status = "N"
        TBFluxo!Documento = TBReceber!Txt_pedido
        TBFluxo!Bloqueado = False
        TBFluxo!ID_empresa = .Cmb_empresa.ItemData(.Cmb_empresa.ListIndex)
        TBFluxo!IDintconta = TBReceber!IDintconta
        
        TBFluxo.Update
        Conexao.Execute "Update tbl_ContasPagar Set IDFluxo = " & TBFluxo!IDFluxo & " where IdIntConta = " & TBFluxo!IDintconta
        TBFluxo.Close
        
        ProcCriaFamiliaFinanceiro .txtTotalPedido, .txtIDPedido
        TBReceber.Close
        Contador = Contador + 1
    Loop
        
    USMsgBox ("Nova(s) conta(s) enviada(s) para o financeiro com sucesso."), vbInformation, "CAPRIND v5.0"
    '==================================
    Modulo = "Compras/Pedido"
    Evento = "Enviar p/ financeiro"
    ID_documento = .txtIDPedido
    Documento = "Nº pedido: " & .txtPedido.Text
    Documento1 = ""
    ProcGravaEvento
    '==================================
End With
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdFinanceiro_porcentagem_Click()
On Error GoTo tratar_erro

If FunVerifDados = False Then Exit Sub
With frmCompras_Pedido
    If .cmbMoeda <> "" And .cmbMoeda <> "REAL" Then
Mensagem1:
        reposicao = InputBox("Favor informar o valor do " & .cmbMoeda & ".")
        If reposicao = "" Then Exit Sub
        If IsNumeric(reposicao) = True Then
            Qtd = reposicao
        Else
            USMsgBox ("Só é permitido número neste campo."), vbExclamation, "CAPRIND v5.0"
            GoTo Mensagem1
        End If
        If Qtd = "0" Then Qtd = 1
    Else
        Qtd = 1
    End If
    
    Contador = 0
    ValoresParcelas = 0
    Set TBContas = CreateObject("adodb.recordset")
    TBContas.Open "Select Sum(dbl_valorpagto) as ValoresParcelas from tbl_contaspagar where txt_pedido = '" & .txtPedido & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBContas.EOF = False Then
        ValoresParcelas = Format(IIf(IsNull(TBContas!ValoresParcelas), 0, TBContas!ValoresParcelas), "###,##0.00")
    End If
    TBContas.Close
    
    With frmCompras_Pedido_Duplicata_Porcento
        .txtValorDuplicata = Format((frmCompras_Pedido.txtTotalPedido * Qtd) - ValoresParcelas, "###,##0.00")
        .txtsaldodin = Format(.txtValorDuplicata, "###,##0.00")
        .Show 1
    End With
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Function FunVerifDados() As Boolean
On Error GoTo tratar_erro

FunVerifDados = True
Acao = "enviar para o financeiro"
With frmCompras_Pedido
    If .cmbpagamento = "" Then
        NomeCampo = "as condições de pagamento"
        ProcVerificaAcao
        FunVerifDados = False
        Exit Function
    End If
    valor = 0
    Set TBContas = CreateObject("adodb.recordset")
    TBContas.Open "Select Sum(ValorPago) as Valor from tbl_ContasPagar where txt_pedido = '" & .txtPedido & "' and Status <> 'TÍTULO EM ABERTO'", Conexao, adOpenKeyset, adLockOptimistic
    If TBContas.EOF = False Then
        valor = IIf(IsNull(TBContas!valor), 0, TBContas!valor)
        Valor1 = .txtTotalPedido
        If valor >= Valor1 Then
            USMsgBox ("Não é permitido enviar para o financeiro, pois existem duplicatas que já foram alteradas no financeiro."), vbExclamation, "CAPRIND v5.0"
            TBContas.Close
            FunVerifDados = False
        End If
    End If
    TBContas.Close
End With

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyEscape: Unload Me
End Select
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCriaFamiliaFinanceiro(ValorTotal As Double, IDpedido As Long)
On Error GoTo tratar_erro

Conexao.Execute "DELETE from familia_financeiro where IDConta = " & TBReceber!IDintconta & " and TipoConta = 'P' and Deposito_transf = 'False'"
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select * from Compras_pedido_lista where IDpedido = " & IDpedido & " and Remessa = 'False'", Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    Do While TBProduto.EOF = False
        Set TBFamilia = CreateObject("adodb.recordset")
        TBFamilia.Open "select * from projproduto where desenho = '" & TBProduto!Desenho & "' and ID_PC is not null and ID_PC <> 0", Conexao, adOpenKeyset, adLockOptimistic
        If TBFamilia.EOF = False Then
            If TBFamilia!ID_PC <> "" Then
                'Verifica a porcentagem representada pelo valor da duplicata
                Valor1 = Format((TBReceber!dbl_valorpagto * 100) / ValorTotal, "###,##0.0000000000")
                
                Set TBCiclo = CreateObject("adodb.recordset")
                TBCiclo.Open "select * from familia_financeiro where ID_PC = " & TBFamilia!ID_PC & " and idconta = " & TBReceber!IDintconta & " and TipoConta = 'P'", Conexao, adOpenKeyset, adLockOptimistic
                If TBCiclo.EOF = True Then TBCiclo.AddNew
                TBCiclo!ID_PC = TBFamilia!ID_PC
                TBCiclo!IDConta = TBReceber!IDintconta
                TBCiclo!IDnota = 0
                
                Qtde = TBProduto!preco_total
                valor = Format((Qtde * Valor1) / 100, "###,##0.00")
                valor = IIf(IsNull(TBCiclo!valor), 0, TBCiclo!valor) + valor
                
                TBCiclo!valor = Format(valor, "###,##0.00")
                TBCiclo!TipoConta = "P"
                TBCiclo.Update
                TBCiclo.Close
            End If
        End If
        TBFamilia.Close
        TBProduto.MoveNext
    Loop
End If
TBProduto.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
