VERSION 5.00
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Begin VB.Form frmVendas_propostaII_MenuFinanceiro 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Vendas | Proposta comercial"
   ClientHeight    =   5400
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   4665
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   4665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   7
      Top             =   4995
      Width           =   4665
      _ExtentX        =   8229
      _ExtentY        =   714
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   435
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   4665
      _ExtentX        =   8229
      _ExtentY        =   767
      DibPicture      =   "frmVendas_propostaII_MenuFinanceiro.frx":0000
      CaptionDelimiter=   "|"
      CaptionOnCenter =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Icon            =   "frmVendas_propostaII_MenuFinanceiro.frx":7180
   End
   Begin VB.Frame Frame1 
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
      Height          =   495
      Left            =   540
      TabIndex        =   5
      Top             =   720
      Width           =   3525
      Begin VB.CheckBox chk_Servicos 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Serviços"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1230
         TabIndex        =   1
         Top             =   180
         Value           =   1  'Checked
         Width           =   945
      End
      Begin VB.CheckBox chk_Produtos 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Produtos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   150
         TabIndex        =   0
         Top             =   180
         Value           =   1  'Checked
         Width           =   945
      End
   End
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
      Height          =   2745
      Left            =   540
      TabIndex        =   4
      Top             =   1170
      Width           =   3525
      Begin DrawSuite2022.USButton cmdFinanceiro 
         Height          =   1080
         Left            =   180
         TabIndex        =   2
         Top             =   240
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   1905
         DibPicture      =   "frmVendas_propostaII_MenuFinanceiro.frx":749A
         Caption         =   "Gerar duplicatas"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   4960354
         BorderColorDisabled=   13160660
         BorderColorDown =   4210752
         BorderColorOver =   49152
         GradientColor1  =   4960354
         GradientColor2  =   4960354
         GradientColor3  =   4960354
         GradientColor4  =   4960354
         GradientColorDisabled1=   14215660
         GradientColorDisabled2=   14215660
         GradientColorDisabled3=   14215660
         GradientColorDisabled4=   14215660
         GradientColorOver1=   49152
         GradientColorOver2=   49152
         GradientColorOver3=   49152
         GradientColorOver4=   49152
         GradientColorDown1=   32768
         GradientColorDown2=   32768
         GradientColorDown3=   32768
         GradientColorDown4=   32768
         PicAlign        =   7
         PicSize         =   4
         PicSizeH        =   48
         PicSizeW        =   48
         ShowFocusRect   =   0   'False
         Theme           =   3
      End
      Begin DrawSuite2022.USButton cmdFinanceiro_porcentagem 
         Height          =   1020
         Left            =   180
         TabIndex        =   3
         Top             =   1530
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   1799
         DibPicture      =   "frmVendas_propostaII_MenuFinanceiro.frx":115BD
         Caption         =   "Gerar duplicatas (em porcentagem%)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   5263559
         BorderColorDisabled=   13160660
         BorderColorDown =   4013465
         BorderColorOver =   4408288
         GradientColor1  =   5263559
         GradientColor2  =   5263559
         GradientColor3  =   5263559
         GradientColor4  =   5263559
         GradientColorDisabled1=   13160660
         GradientColorDisabled2=   13160660
         GradientColorDisabled3=   13160660
         GradientColorDisabled4=   13160660
         GradientColorOver1=   4408288
         GradientColorOver2=   4408288
         GradientColorOver3=   4408288
         GradientColorOver4=   4408288
         GradientColorDown1=   4013465
         GradientColorDown2=   4013465
         GradientColorDown3=   4013465
         GradientColorDown4=   4013465
         PicAlign        =   7
         PicSize         =   4
         PicSizeH        =   48
         PicSizeW        =   48
         Theme           =   4
      End
   End
   Begin DrawSuite2022.USOptionButton opthoje 
      Height          =   255
      Left            =   660
      TabIndex        =   8
      Top             =   4110
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   450
      Caption         =   "Calcular vencimento(s) a partir da data de hoje"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   128
      ShowFocusRect   =   0   'False
   End
   Begin DrawSuite2022.USOptionButton optdataPI 
      Height          =   255
      Left            =   630
      TabIndex        =   9
      Top             =   4410
      Width           =   3345
      _ExtentX        =   5900
      _ExtentY        =   450
      Caption         =   "Calcular vencimento(s) a partir da data da venda"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   8388608
      ShowFocusRect   =   0   'False
      Value           =   -1  'True
   End
End
Attribute VB_Name = "frmVendas_propostaII_MenuFinanceiro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdFinanceiro_Click()
On Error GoTo tratar_erro

ProcPI_Financeiro

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdFinanceiro_porcentagem_Click()
On Error GoTo tratar_erro

ProcPI_Financeiro_porc

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

If KeyCode = vbKeyEscape Then Unload Me
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcVerifDescontos()
On Error GoTo tratar_erro

TotalDesconto = 0
TotalDescontoServico = 0
'Verifica desconto produtos
If Vendas_PI = True Then IDlista = frmVendas_PI.txtId Else IDlista = frmVendas_proposta.txtId
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select Sum(ROUND(preco_unitario * Quantidade, 2)) as TotalProduto, Sum(ROUND(Preco_lote, 2)) as Valor from vendas_carteira where cotacao = " & IDlista & " and Tipo = 'P' and Retorno = 'False'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    TotalProduto = IIf(IsNull(TBAbrir!TotalProduto), 0, TBAbrir!TotalProduto)
    valor = IIf(IsNull(TBAbrir!valor), 0, TBAbrir!valor)
    If valor > TotalProduto Then TotalProduto = valor
    TotalDesconto = Format(TotalProduto - valor, "###,##0.00")
End If
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select Sum(ROUND(preco_unitario * Quantidade, 2)) as TotalServicos, Sum(ROUND(Preco_lote, 2)) as Valor from vendas_carteira where cotacao = " & IDlista & " and Tipo = 'S'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    TotalServicos = IIf(IsNull(TBAbrir!TotalServicos), 0, TBAbrir!TotalServicos)
    valor = IIf(IsNull(TBAbrir!valor), 0, TBAbrir!valor)
    If valor > TotalServicos Then TotalServicos = valor
    TotalDescontoServico = Format(TotalServicos - valor, "###,##0.00")
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcVerificaImposto()
On Error GoTo tratar_erro

DestacaImpostos = "NÃO"
With IIf(Vendas_PI = True, frmVendas_PI, frmVendas_proposta)
    IDlista = .txtId
    ID_empresa = .Cmb_empresa.ItemData(.Cmb_empresa.ListIndex)
End With
Set TBFI = CreateObject("adodb.recordset")
TBFI.Open "Select CFOP.* from vendas_carteira VC INNER JOIN tbl_NaturezaOperacao CFOP ON CFOP.IDCountCfop = VC.ID_CFOP where VC.Cotacao = " & IDlista & " and CFOP.Retem = 'True'", Conexao, adOpenKeyset, adLockOptimistic
If TBFI.EOF = False Then
    DestacaImpostos = "SIM"
End If

Set TBFI = CreateObject("adodb.recordset")
TBFI.Open "select * from Empresa where Codigo = " & ID_empresa, Conexao, adOpenKeyset, adLockOptimistic
    If TBFI.EOF = False Then
        Regime = 0
        If TBFI!Simples = True Then Regime = 1
        If TBFI!Presumido = True Then Regime = 2
        If TBFI!Real = True Then Regime = 3
        If TBFI!Simples1 = True Then Regime = 4
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from Impostos where Regime = " & Regime, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            If valor > TBAbrir!Acima And TBAbrir!Acima > 0 And DestacaImpostos = "SIM" Then
                Set TBCotacao = CreateObject("adodb.recordset")
                TBCotacao.Open "Select * FROM vendas_proposta where Cotacao = " & IDlista, Conexao, adOpenKeyset, adLockOptimistic
                If TBCotacao.EOF = False Then
                    ValorTotal = Format(valor - (IIf(IsNull(TBCotacao!Total_PIS_serv) = False, TBCotacao!Total_PIS_serv, 0) + IIf(IsNull(TBCotacao!Total_Cofins_serv) = False, TBCotacao!Total_Cofins_serv, 0) + IIf(IsNull(TBCotacao!Total_CSLL_serv) = False, TBCotacao!Total_CSLL_serv, 0) + IIf(IsNull(TBCotacao!Total_INSS_serv) = False, TBCotacao!Total_INSS_serv, 0) + IIf(IsNull(TBCotacao!Total_IRRF_serv) = False, TBCotacao!Total_IRRF_serv, 0)), "###,##0.00")
                End If
                TBCotacao.Close
            ElseIf valor >= 667 And valor <= 5000 And DestacaImpostos = "SIM" Then
                    Set TBCotacao = CreateObject("adodb.recordset")
                    TBCotacao.Open "Select * FROM vendas_proposta where Cotacao = " & IDlista, Conexao, adOpenKeyset, adLockOptimistic
                    If TBCotacao.EOF = False Then
                        ValorTotal = valor - IIf(IsNull(TBCotacao!Total_IRRF_serv), 0, TBCotacao!Total_IRRF_serv)
                    End If
                    TBCotacao.Close
                Else
                    ValorTotal = valor
            End If
        End If
        TBAbrir.Close
End If
TBFI.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCriaFamiliaFinanceiro(ValorTotal As Double, IDpedido As Long)
On Error GoTo tratar_erro

Conexao.Execute "DELETE from familia_financeiro where IDConta = " & TBReceber!IDintconta & " and TipoConta = 'R' and Deposito_transf = 'False'"
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "select * from Vendas_carteira where cotacao = " & IDpedido & " and Retorno = 'False'", Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    Do While TBProduto.EOF = False
        Set TBFamilia = CreateObject("adodb.recordset")
        TBFamilia.Open "select * from projproduto where desenho = '" & TBProduto!Desenho & "' and ID_PC1 is not null and ID_PC1 <> 0", Conexao, adOpenKeyset, adLockOptimistic
        If TBFamilia.EOF = False Then
            If TBFamilia!ID_PC1 <> "" Then
                'Verifica a porcentagem representada pelo valor da duplicata
                Valor1 = Format((TBReceber!valor * 100) / ValorTotal, "###,##0.0000000000")
                
                Set TBCiclo = CreateObject("adodb.recordset")
                TBCiclo.Open "select * from familia_financeiro where ID_PC = " & TBFamilia!ID_PC1 & " and idconta = " & TBReceber!IDintconta & " and TipoConta = 'R'", Conexao, adOpenKeyset, adLockOptimistic
                If TBCiclo.EOF = True Then TBCiclo.AddNew
                TBCiclo!ID_PC = TBFamilia!ID_PC1
                TBCiclo!IDConta = TBReceber!IDintconta
                TBCiclo!IDnota = 0
                
                Qtde = TBProduto!preco_lote + TBProduto!dbl_valoripi
                valor = Format((Qtde * Valor1) / 100, "###,##0.00")
                valor = IIf(IsNull(TBCiclo!valor), 0, TBCiclo!valor) + valor
                
                TBCiclo!valor = Format(valor, "###,##0.00")
                TBCiclo!TipoConta = "R"
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

Sub ProcPI_Financeiro()
On Error GoTo tratar_erro

With IIf(Vendas_PI = True, frmVendas_PI, frmVendas_proposta)
    If USMsgBox("Deseja realmente gerar contas a receber " & IIf(Vendas_PI = True, "do pedido interno", "da proposta comercial") & " n° " & .txtCotacao & "?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
    If FunVerifDados = False Then Exit Sub
    QtdeSaida = 0
    Contador = 1
    Contador2 = 0
    Controle = 1
    nPagto = 0
    Valor_Duplicatas = 0
    If Not IsNumeric(Left(.txtCondicoes, 1)) Then
        USMsgBox ("O campo condições de pagamento tem que ser em dias, favor alterar."), vbExclamation, "CAPRIND v5.0"
        Unload Me
        Exit Sub
    End If
    
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
    
    Set TBCotacao = CreateObject("adodb.recordset")
    TBCotacao.Open "Select Total_retencao_PIS, Total_retencao_Cofins FROM vendas_proposta where Cotacao = " & .txtId.Text, Conexao, adOpenKeyset, adLockOptimistic
    If TBCotacao.EOF = False Then
        Valor_Retencao_PIS = IIf(IsNull(TBCotacao!Total_retencao_PIS), 0, TBCotacao!Total_retencao_PIS)
        Valor_Retencao_Cofins = IIf(IsNull(TBCotacao!Total_retencao_Cofins), 0, TBCotacao!Total_retencao_Cofins)
    End If
    TBCotacao.Close
        
    QtdeSaida = Len(.txtCondicoes)
            
    TextoCond = ""
    Do While Contador <= QtdeSaida
        If Mid(.txtCondicoes, Contador, 1) = "/" Or IsNumeric(Mid(.txtCondicoes, Contador, 1)) = True Then
            If TextoCond = "" Then TextoCond = Mid(.txtCondicoes, Contador, 1) Else TextoCond = TextoCond & Mid(.txtCondicoes, Contador, 1)
        End If
        Contador = Contador + 1
    Loop
    
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
    
    ProcVerifDescontos
    ProcCalculaValores
                                    
    mxValorPag = Format(ValorTotal / nPagto, "###,##0.00")
    
    Contador = 1
    Contador3 = 1
    
    If optdataPI.Value = True Then
        Dataini = IIf(Vendas_PI = True, .txtDatavendas_PI, .txtDatavendas)
    Else
        Dataini = Date
    End If
    

    Controle = 0
    Do While Contador <= QtdeSaida
        
        Contador2 = 0
        Do While Mid(.txtCondicoes, Contador, 1) <> "/" And Contador <= QtdeSaida
            Contador2 = Contador2 + 1
            Contador = Contador + 1
        Loop
        
        mxCondpag = ReturnNumbersOnly(Mid(.txtCondicoes, Contador3, Contador2))
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
        TBReceber.Open "Select * from tbl_contas_receber where proposta = '" & .txtCotacao & "' and parcela = '" & Par1 & "/" & Par2 & "' order by vencimento", Conexao, adOpenKeyset, adLockOptimistic
        If TBReceber.EOF = True Then
            TBReceber.AddNew
            TBReceber!Data_transacao = Date
            TBReceber!Parcial = False
            TBReceber!titulodesc = False
            TBReceber!Bloqueado = False
            TBReceber!Logsit = "N"
            TBReceber!IDtrocatitulo = 0
            TBReceber!Antecipacao = False
            TBReceber!Devolucao = False
            TBReceber!status = "TÍTULO EM ABERTO"
            TBReceber!Responsavel = pubUsuario
            TBReceber!NFiscal = "000000000"
            TBReceber!ID_nota = 0
        End If
        TBReceber!Proposta = .txtCotacao
        TBReceber!emissao = Date
        TBReceber!Nome_Razao = .txtCliente
        TBReceber!IDCliente = .txtIDcliente
        TBReceber!Cidade = .txtCidade
        TBReceber!Estado = .txtuf
        TBReceber!ID_empresa = .Cmb_empresa.ItemData(.Cmb_empresa.ListIndex)
        TBReceber!Tipo_doc = "REQ"
        TBReceber!txt_ndocumento = IIf(frmVendas_PI.txtreferente.Text <> "", frmVendas_PI.txtreferente, "")
        If Controle = nPagto Then
            TBReceber!valor = Format(ValorTotal - Valor_Duplicatas, "###,##0.00")
        Else
            TBReceber!valor = mxValorPag
        End If
        TBReceber!ValorExtenso = FunValorExtenso(TBReceber!valor)
        Valor_Duplicatas = Valor_Duplicatas + mxValorPag
'================================================================
' Vencimento das duplicatas
'================================================================
        TBReceber!Vencimento = Dataini + mxCondpag
'================================================================
        TBReceber!Parcela = Par1 & "/" & Par2
        TBReceber!Responsavel = pubUsuario
        TBReceber!Tipo = "CL"
        TBReceber.Update
        intidconta = TBReceber!IDintconta
        
        'Fluxo de Caixa
        Set TBFluxo = CreateObject("adodb.recordset")
        TBFluxo.Open "Select * from tbl_Fluxo_de_caixa where IDFluxo = " & IIf(IsNull(TBReceber!IDFluxo), 0, TBReceber!IDFluxo), Conexao, adOpenKeyset, adLockOptimistic
        If TBFluxo.EOF = True Then TBFluxo.AddNew
        TBFluxo!IDintconta = TBReceber!IDintconta
        TBFluxo!Operacao = "À Creditar"
        TBFluxo!Data = TBReceber!Vencimento
        TBFluxo!valor = TBReceber!valor
        TBFluxo!Descricao = TBReceber!Nome_Razao
        TBFluxo!status = "N"
        TBFluxo!int_NotaFiscal = 0
        TBFluxo!Bloqueado = False
        TBFluxo!ID_empresa = .Cmb_empresa.ItemData(.Cmb_empresa.ListIndex)
        TBFluxo.Update
        Conexao.Execute "UPDATE tbl_contas_receber set IDFluxo = " & TBFluxo!IDFluxo & " where IdIntConta = " & TBReceber!IDintconta
        TBFluxo.Close
        
        ProcCriaFamiliaFinanceiro ValorTotal, .txtId
        TBReceber.Close
        Contador = Contador + 1
    Loop
        
    USMsgBox ("Nova(s) conta(s) enviada(s) para o financeiro com sucesso."), vbInformation, "CAPRIND v5.0"
    '==================================
    If Vendas_PI = True Then Modulo = "Vendas/Pedido interno" Else Modulo = "Vendas/Proposta comercial"
    Evento = "Enviar p/ financeiro"
    ID_documento = .txtId
    Documento = "Nº pedido: " & .txtCotacao & " - Rev.: " & .txtrevisao
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

Sub ProcPI_Financeiro_porc()
On Error GoTo tratar_erro

If FunVerifDados = False Then Exit Sub
With IIf(Vendas_PI = True, frmVendas_PI, frmVendas_proposta)
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
    
    Set TBCotacao = CreateObject("adodb.recordset")
    TBCotacao.Open "Select * FROM vendas_proposta where Cotacao = " & .txtId.Text, Conexao, adOpenKeyset, adLockOptimistic
    If TBCotacao.EOF = False Then
        Valor_Retencao_PIS = IIf(IsNull(TBCotacao!Total_retencao_PIS), 0, TBCotacao!Total_retencao_PIS)
        Valor_Retencao_Cofins = IIf(IsNull(TBCotacao!Total_retencao_Cofins), 0, TBCotacao!Total_retencao_Cofins)
    End If
    TBCotacao.Close
    
    Contador = 0
    ValoresParcelas = 0
    Set TBContas = CreateObject("adodb.recordset")
    TBContas.Open "Select Sum(Valor) as ValoresParcelas from tbl_contas_receber where proposta = '" & .txtCotacao & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBContas.EOF = False Then
        ValoresParcelas = IIf(IsNull(TBContas!ValoresParcelas), 0, TBContas!ValoresParcelas)
    End If
    TBContas.Close
    
    ProcVerifDescontos
    ProcCalculaValores
    frmVendas_PI_Duplicata_Porcento.Show 1
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
With IIf(Vendas_PI = True, frmVendas_PI, frmVendas_proposta)
    If .txtStatus <> "VENDIDA" And .txtStatus <> "VENDIDA PARCIAL" Then
        USMsgBox ("Só é permitido enviar para o financeiro " & IIf(Vendas_PI = True, "pedido interno", "proposta comercial") & " com o status vendida ou vendida parcial."), vbExclamation, "CAPRIND v5.0"
        FunVerifDados = False
        Exit Function
    End If
    If chk_Produtos.Value = 0 And chk_Servicos.Value = 0 Then
        NomeCampo = "a opção de produto ou serviço"
        ProcVerificaAcao
        FunVerifDados = False
        Exit Function
    End If
    If .txttotalproposta = "" Or .txttotalproposta = "0,00" Then
        NomeCampo = "o valor total"
        ProcVerificaAcao
        FunVerifDados = False
        Exit Function
    End If
    valor = 0
    Set TBContas = CreateObject("adodb.recordset")
    TBContas.Open "Select Sum(valortitulorecebido) as Valor from tbl_contas_receber where proposta = '" & .txtCotacao & "' and Status <> 'TÍTULO EM ABERTO'", Conexao, adOpenKeyset, adLockOptimistic
    If TBContas.EOF = False Then
        valor = IIf(IsNull(TBContas!valor), 0, TBContas!valor)
        Valor1 = .txttotalproposta
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

Private Sub ProcCalculaValores()
On Error GoTo tratar_erro

With IIf(Vendas_PI = True, frmVendas_PI, frmVendas_proposta)
    TotalProduto = IIf(.txt_vlrtotalprod = "", 0, .txt_vlrtotalprod) - TotalDesconto
    TotalIPI = IIf(.txt_TotalIPI = "", 0, .txt_TotalIPI)
    ICMSCST = IIf(.txt_ICMSs = "", 0, .txt_ICMSs)
    TotalServicos = IIf(.txttotalservicos = "", 0, .txttotalservicos) - TotalDescontoServico
    
    If USMsgBox("Deseja utilizar o valor com imposto?", vbYesNo, "CAPRIND v5.0") = vbYes Then
        If chk_Produtos.Value = 1 And chk_Servicos.Value = 1 Then
            valor = Format((TotalProduto + TotalIPI + ICMSCST - Valor_Retencao_PIS - Valor_Retencao_Cofins) + TotalServicos, "###,##0.00")
        ElseIf chk_Produtos.Value = 1 And chk_Servicos.Value = 0 Then
                valor = Format((TotalProduto + TotalIPI + ICMSCST - Valor_Retencao_PIS - Valor_Retencao_Cofins), "###,##0.00")
            ElseIf chk_Produtos.Value = 0 And chk_Servicos.Value = 1 Then
                valor = Format(TotalServicos, "###,##0.00")
        End If
        ValorTotal = valor * Qtd
        If TotalServicos <> 0 And chk_Servicos.Value = 1 Then
            valor = TotalServicos
            .ProcVerificaEmpresaCliente
            If Permitido = True Then ProcVerificaImposto
            If chk_Produtos.Value = 1 Then
                ValorTotal = (ValorTotal + (TotalProduto + TotalIPI + ICMSCST - Valor_Retencao_PIS - Valor_Retencao_Cofins)) * Qtd
            Else
                ValorTotal = ValorTotal * Qtd
            End If
        End If
    Else
        If chk_Produtos.Value = 1 And chk_Servicos.Value = 1 Then
            valor = Format((TotalProduto - Valor_Retencao_PIS - Valor_Retencao_Cofins) + TotalServicos, "###,##0.00")
        ElseIf chk_Produtos.Value = 1 And chk_Servicos.Value = 0 Then
                valor = Format((TotalProduto - Valor_Retencao_PIS - Valor_Retencao_Cofins), "###,##0.00")
            ElseIf chk_Produtos.Value = 0 And chk_Servicos.Value = 1 Then
                valor = Format(TotalServicos, "###,##0.00")
        End If
    
        valor = Format((TotalProduto - Valor_Retencao_PIS - Valor_Retencao_Cofins) + TotalServicos, "###,##0.00")
        ValorTotal = valor * Qtd
        If TotalServicos <> 0 And chk_Servicos.Value = 1 Then
            valor = TotalServicos
            .ProcVerificaEmpresaCliente
            If Permitido = True Then ProcVerificaImposto
            If chk_Produtos.Value = 1 Then
                ValorTotal = (ValorTotal + (TotalProduto - Valor_Retencao_PIS - Valor_Retencao_Cofins)) * Qtd
            Else
                ValorTotal = ValorTotal * Qtd
            End If
        End If
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

If Vendas_Proposta = True Then Caption = "Vendas - Proposta comercial" Else Caption = "Vendas - Pedido interno"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

