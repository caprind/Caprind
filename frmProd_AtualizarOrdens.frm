VERSION 5.00
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Begin VB.Form frmProd_AtualizarOrdens 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "PCP | Atualizar dados da ordem"
   ClientHeight    =   4140
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   4215
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
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
   ScaleHeight     =   4140
   ScaleWidth      =   4215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   8
      Top             =   3735
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   714
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   405
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   714
      DibPicture      =   "frmProd_AtualizarOrdens.frx":0000
      CaptionDelimiter=   "|"
      CaptionOnCenter =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Icon            =   "frmProd_AtualizarOrdens.frx":62E4
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   2565
      Left            =   360
      TabIndex        =   5
      Top             =   660
      Width           =   3405
      Begin DrawSuite2022.USButton cmdProcessos 
         Height          =   390
         Left            =   330
         TabIndex        =   0
         Top             =   270
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Criar processo do produto"
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
         ShowFocusRect   =   0   'False
         Theme           =   1
      End
      Begin DrawSuite2022.USButton cmdMaterial 
         Height          =   390
         Left            =   330
         TabIndex        =   2
         Top             =   1110
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Valor de material"
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
         ShowFocusRect   =   0   'False
         Theme           =   1
      End
      Begin DrawSuite2022.USButton cmdEstoque 
         Height          =   390
         Left            =   330
         TabIndex        =   4
         Top             =   1950
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Valor unitário no estoque"
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
         ShowFocusRect   =   0   'False
         Theme           =   1
      End
      Begin DrawSuite2022.USButton cmdTerceiros 
         Height          =   390
         Left            =   330
         TabIndex        =   3
         Top             =   1530
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Valor de terceiros"
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
         ShowFocusRect   =   0   'False
         Theme           =   1
      End
      Begin DrawSuite2022.USButton Cmd_MO 
         Height          =   390
         Left            =   330
         TabIndex        =   1
         Top             =   690
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Valor de mão de obra"
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
         ShowFocusRect   =   0   'False
         Theme           =   1
      End
   End
   Begin DrawSuite2022.USProgressBar PBlista 
      Height          =   225
      Left            =   0
      TabIndex        =   6
      Top             =   3510
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   397
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
Attribute VB_Name = "frmProd_AtualizarOrdens"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ProcSair()
On Error GoTo tratar_erro

Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_MO_Click()
On Error GoTo tratar_erro

If FunVerifValidacaoRegistro("atualizar", frmprod.txtDtValidacao_custo, "resultado da ordem", "valor de mão de obra", True) = False Then Exit Sub
Unload Me
ProcMO

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdEstoque_Click()
On Error GoTo tratar_erro

If frmprod.txtDtValidacao_custo = "" Then
    USMsgBox ("Não é permitido atualizar o valor unitário no estoque, pois o resultado da ordem ainda não foi validado."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
ProcEstoque
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdMaterial_Click()
On Error GoTo tratar_erro

If FunVerifValidacaoRegistro("atualizar", frmprod.txtDtValidacao_custo, "resultado da ordem", "valor de material", True) = False Then Exit Sub
ProcMaterial
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdProcessos_Click()
On Error GoTo tratar_erro

ProcProcesso
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdTerceiros_Click()
On Error GoTo tratar_erro

If FunVerifValidacaoRegistro("atualizar", frmprod.txtDtValidacao_custo, "resultado da ordem", "valor de terceiros", True) = False Then Exit Sub
ProcTerceiros
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyEscape: ProcSair
End Select
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcProcesso()
On Error GoTo tratar_erro

With frmprod
    If .txtdesenho = "" Then Exit Sub
    If USMsgBox("Deseja realmente criar um processo deste produto?", vbYesNo, "CAPRIND v5.0") = vbYes Then
        Set TBItem = CreateObject("adodb.recordset")
        TBItem.Open "Select * from projproduto where desenho = '" & .txtdesenho & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBItem.EOF = False Then
            Set TBProcessos = CreateObject("adodb.recordset")
            TBProcessos.Open "Select * from Processos where codproduto = " & TBItem!Codproduto & " and tipo <> 'C'", Conexao, adOpenKeyset, adLockOptimistic
            If TBProcessos.EOF = True Then
                TBItem!PCusto = .Txt_custo_total_peca
                TBItem.Update
                
                TBProcessos.AddNew
                TBProcessos!RespValidacao = pubUsuario
                TBProcessos!DtValidacao = Date
                TBProcessos!cronometrado = "SIM"
                TBProcessos!Custo = .Txt_custo_total_peca
                TBProcessos!TTotal = .txtpecautil
                TBProcessos!Codproduto = TBItem!Codproduto
                TBProcessos!Revisao = 0
                TBProcessos!Contador = TBItem!RevDesenho
                If .Opt_componente.Value = 1 Then
                    TBProcessos!Tipo = "F"
                ElseIf .Opt_subconjunto.Value = 1 Then
                        TBProcessos!Tipo = "M"
                    Else
                        TBProcessos!Tipo = "E"
                End If
                TBProcessos.Update
                IDPROCESSO = TBProcessos!IDPROCESSO
                TBProcessos.Close
                
                'Atualizar ordem com processo
                Conexao.Execute "Update producao Set IDPROCESSO = " & IDPROCESSO & " where desenho = '" & .txtdesenho.Text & "'"
                
                'Atualizar fases
                Set TBOrdem = CreateObject("adodb.recordset")
                TBOrdem.Open "Select * from ordemservico where Ordem = " & .txtof.Text & " order by idproducao", Conexao, adOpenKeyset, adLockOptimistic
                If TBOrdem.EOF = False Then
                    Do While TBOrdem.EOF = False
                        Set TBFases = CreateObject("adodb.recordset")
                        TBFases.Open "Select * from fases where idfase = " & IIf(IsNull(TBOrdem!IDFase), 0, TBOrdem!IDFase), Conexao, adOpenKeyset, adLockOptimistic
                        If TBFases.EOF = True Then
                            TBFases.AddNew
                            TBFases!versao = "A"
                        End If
                        TBFases!IDPROCESSO = IDPROCESSO
                        TBFases!Fase = TBOrdem!Fase
                        TBFases!maquina = TBOrdem!maquina
                        TBFases!Descricao = TBOrdem!descfase
                        If IsNull(TBOrdem!TPUTIL) = False And TBOrdem!TPUTIL <> "" Then TBFases!Preparacao = TBOrdem!TPUTIL Else TBFases!Preparacao = "00:00:00"
                        If IsNull(TBOrdem!TEUTIL) = False And TBOrdem!TEUTIL <> "" Then TBFases!Execucao = Left(TBOrdem!TEUTIL, 8) Else TBFases!Execucao = "00:00:00"
                        If IsNull(TBOrdem!TPUTIL) = False And TBOrdem!TPUTIL <> "" Then TBFases!TempoPreparacao = TBOrdem!TPUTIL Else TBFases!TempoPreparacao = "00:00:00"
                        If IsNull(TBOrdem!TEUTIL) = False And TBOrdem!TEUTIL <> "" Then TBFases!TempoExecucao = TBOrdem!TEUTIL Else TBFases!TempoExecucao = "00:00:00"
                        If TBOrdem!custos = True Then
                            TBFases!Custo = TBOrdem!CRPECA
                        Else
                            If TBOrdem!Totalprod <> 0 Then TBFases!Custo = TBOrdem!CTServico / TBOrdem!Totalprod Else TBFases!Custo = TBOrdem!CTServico
                        End If
                        TBFases!cronometrado = True
                        TBOrdem!IDFase = TBFases!IDFase
                        TBOrdem.Update
                        TBFases.Update
                        TBOrdem.MoveNext
                    Loop
                End If
                TBOrdem.Close
                USMsgBox ("Novo processo cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
                '==================================
                Modulo = "PCP/Gerenciamento de ordem"
                Evento = "Novo processo"
                ID_documento = IDPROCESSO
                Documento = "Cód. interno: " & frmprod.txtdesenho
                Documento1 = ""
                ProcGravaEvento
                '==================================
            Else
                USMsgBox ("Já existe processo cadastrado para este produto."), vbExclamation, "CAPRIND v5.0"
            End If
        Else
            USMsgBox ("Produto não cadastrado, favor cadastrar."), vbExclamation, "CAPRIND v5.0"
            Exit Sub
        End If
        TBItem.Close
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEstoque()
On Error GoTo tratar_erro

With frmprod
    If USMsgBox("Deseja realmente atualizar o valor do estoque desta ordem?", vbYesNo, "CAPRIND v5.0") = vbYes Then
        .ProcAtualizarVlrUnitOrdemEst "Ordem = " & OF, PBLista
        USMsgBox ("Valor unitário atualizado no estoque com sucesso."), vbInformation, "CAPRIND v5.0"
        '==================================
        Modulo = "PCP/Gerenciamento de ordem"
        Evento = "Atualizar valor unitário no estoque"
        ID_documento = .txtof
        Documento = "Ordem: " & .txtof.Text & " - Cód. interno: " & .txtdesenho
        Documento1 = ""
        ProcGravaEvento
        '==================================
        .ProcAbrirRe
    End If
End With
        
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcMO()
On Error GoTo tratar_erro

With frmprod
    If USMsgBox("Deseja realmente atualizar o valor de mão de obra desta ordem?", vbYesNo, "CAPRIND v5.0") = vbYes Then
        .SSTab1.Tab = 4
        
        Set TBAcessos = CreateObject("adodb.recordset")
        TBAcessos.Open "Select PF.IDFase from Producao P INNER JOIN " & NomeTabelaAp & " PF ON PF.Ordem = P.Ordem where P.Ordem = " & OF & " Group by PF.IDFase", Conexao, adOpenKeyset, adLockReadOnly
        If TBAcessos.EOF = False Then
            Do While TBAcessos.EOF = False
                Set TBFI = CreateObject("adodb.recordset")
                TBFI.Open "Select Ordem from OrdemServico where Idproducao = " & TBAcessos!IDFase, Conexao, adOpenKeyset, adLockOptimistic
                If TBFI.EOF = False Then
                    .txtof = TBFI!Ordem
                    .txtAPOF = TBFI!Ordem
                    
                    .ProcCarregaAPOS
                    
                    .cmbAPOS = TBAcessos!IDFase
                    
                    .ProcCorrigeProdFasesTotalizacao
                    .ProcAtualizaPrepExecUtil
                    .ProcGravarStatusOSOF
                    .ProcGravaValoresOS
                    .ProcAtualizaOFAtualizacao
                End If
                TBAcessos.MoveNext
            Loop
        End If
        TBAcessos.Close
        USMsgBox ("Valor de mão de obra atualizada com sucesso."), vbInformation, "CAPRIND v5.0"
        '==================================
        Modulo = "PCP/Gerenciamento de ordem"
        Evento = "Atualizar valor de material utilizado na ordem"
        ID_documento = .txtof
        Documento = "Ordem: " & .txtof.Text & " - Cód. interno: " & .txtdesenho
        Documento1 = ""
        ProcGravaEvento
        '==================================
        .SSTab1.Tab = 5
        .ProcAbrirRe
    End If
End With
        
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcMaterial()
On Error GoTo tratar_erro

If USMsgBox("Deseja realmente atualizar o valor do material desta ordem?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    Set TBproducao = CreateObject("adodb.recordset")
    TBproducao.Open "Select * from producao where Ordem = " & OF, Conexao, adOpenKeyset, adLockOptimistic
    If TBproducao.EOF = False Then
        TBproducao!CTMaterial = 0
        TBproducao.Update
        
        valor = 0
        Valor1 = 0
        Set TBFI = CreateObject("adodb.recordset")
        
        StrSql = "Select PM.Codigo, sum(EM.VlrUnit*EM.saida) as Valor, PM.Valor_saida_estoque from (producaomaterial PM INNER JOIN Estoque_movimentacao EM ON EM.Desenho = PM.Codigo) INNER JOIN Estoque_Controle EC ON EC.IDestoque = EM.IDestoque where PM.Ordem = " & TBproducao!Ordem & " and EM.Documento = '" & TBproducao!Ordem & "' and EC.Consignacao = 'False' and (EM.Operacao = 'SAIDA_ORDEM' or EM.Operacao = 'SAIDA_ORDEM_PARCIAL') group by PM.codigo, PM.Valor_saida_estoque"
        'Debug.print StrSql
        
        TBFI.Open StrSql, Conexao, adOpenKeyset, adLockReadOnly
        
        If TBFI.EOF = False Then
            PBLista.Min = 0
            PBLista.Max = TBFI.RecordCount
            PBLista.Value = 1
            Contador = 0
            Do While TBFI.EOF = False
                valor = IIf(IsNull(TBFI!valor), 0, TBFI!valor)
                NovoValor = Replace(valor, ",", ".")
                Conexao.Execute "UPDATE producaomaterial Set Valor_saida_estoque = " & NovoValor & " where Codigo = '" & TBFI!CODIGO & "' and Ordem = " & TBproducao!Ordem
                
                Valor1 = Valor1 + valor
                TBFI.MoveNext
                Contador = Contador + 1
                PBLista.Value = Contador
            Loop
        End If
        TBFI.Close
        If Valor1 > 0 Then
            TBproducao!CTMaterial = Format(Valor1, "###,##0.00")
            TBproducao.Update
        Else
            Set TBAliquota = CreateObject("adodb.recordset")
            TBAliquota.Open "Select * from Empresa where Codigo = " & TBproducao!ID_empresa, Conexao, adOpenKeyset, adLockOptimistic
            If TBAliquota.EOF = False Then
                If TBAliquota!Real = True Then
                    TextoFiltro = "Sum(NFP.dbl_ValorTotal - (CST_ICMS.Valor_ICMS + CST_ICMS.Valor_ICMS_ST + CST_ICMS.Valor_ICMS_SN) - (NFP.Total_PIS_prod + NFP.Total_Cofins_prod))"
                Else
                    TextoFiltro = "Sum(NFP.dbl_ValorTotal - (CST_ICMS.Valor_ICMS + CST_ICMS.Valor_ICMS_ST + CST_ICMS.Valor_ICMS_SN))"
                End If
                Set TBCompras_Lista = CreateObject("adodb.recordset")
                StrSql = "Select ROUND(" & TextoFiltro & ", 2) AS Valortotal from ((tbl_Detalhes_Nota NFP INNER JOIN Compras_pedido_lista PP ON NFP.Int_codigo = PP.IDlista and NFP.int_Cod_Produto = PP.Desenho) INNER JOIN tbl_Detalhes_Nota_CST_ICMS CST_ICMS ON CST_ICMS.ID_item = NFP.Int_codigo) INNER JOIN Compras_pedido P ON P.IDpedido = PP.IDpedido where NFP.Ordem = '" & TBproducao!Ordem & "' and PP.remessa = 'False' and (PP.OS is null or PP.OS = 0)"
                'Debug.print StrSql
                
                
                TBCompras_Lista.Open "Select ROUND(" & TextoFiltro & ", 2) AS Valortotal from ((tbl_Detalhes_Nota NFP INNER JOIN Compras_pedido_lista PP ON NFP.Int_codigo = PP.IDlista and NFP.int_Cod_Produto = PP.Desenho) INNER JOIN tbl_Detalhes_Nota_CST_ICMS CST_ICMS ON CST_ICMS.ID_item = NFP.Int_codigo) INNER JOIN Compras_pedido P ON P.IDpedido = PP.IDpedido where NFP.Ordem = '" & TBproducao!Ordem & "' and PP.remessa = 'False' and (PP.OS is null or PP.OS = 0)", Conexao, adOpenKeyset, adLockOptimistic
                If TBCompras_Lista.EOF = False Then
                    TBproducao!CTMaterial = IIf(IsNull(TBCompras_Lista!ValorTotal), 0, TBCompras_Lista!ValorTotal)
                End If
                TBCompras_Lista.Close
                TBproducao.Update
            End If
            TBAliquota.Close
        End If
    End If
    TBproducao.Close
    USMsgBox ("Valor de material atualizado com sucesso."), vbInformation, "CAPRIND v5.0"
    '==================================
    Modulo = "PCP/Gerenciamento de ordem"
    Evento = "Atualizar valor de material utilizado na ordem"
    ID_documento = frmprod.txtof
    Documento = "Ordem: " & frmprod.txtof.Text & " - Cód. interno: " & frmprod.txtdesenho
    Documento1 = ""
    ProcGravaEvento
    '==================================
    frmprod.ProcAbrirRe
End If
        
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcTerceiros()
On Error GoTo tratar_erro

If USMsgBox("Deseja realmente atualizar o valor de terceiros desta ordem?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    valor = 0
    Set TBproducao = CreateObject("adodb.recordset")
    TBproducao.Open "Select * from producao where Ordem = " & OF, Conexao, adOpenKeyset, adLockOptimistic
    If TBproducao.EOF = False Then
        'Custo terceiros utilizado pela OS da ordem
        Set TBOrdem = CreateObject("adodb.recordset")
        TBOrdem.Open "Select * from ordemservico where Ordem = " & OF & " and Custos = 'False'", Conexao, adOpenKeyset, adLockOptimistic
        If TBOrdem.EOF = False Then
            PBLista.Min = 0
            PBLista.Max = TBOrdem.RecordCount
            PBLista.Value = 1
            Contador = 0
            Do While TBOrdem.EOF = False
                Valor1 = 0
                Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open "Select Sum(ROUND(NFP.dbl_ValorUnitario * NFPP.Quantidade, 2)) as Valor1 from (Compras_pedido_lista CPL INNER JOIN tbl_Detalhes_Nota_pedidos NFPP ON NFPP.ID_carteira = CPL.IDlista and NFPP.Codinterno = CPL.Desenho) INNER JOIN tbl_Detalhes_Nota NFP ON NFP.Int_codigo = NFPP.ID_prod_NF where CPL.Ordem = " & TBOrdem!Ordem & " and CPL.OS  = " & TBOrdem!IDProducao & " and CPL.Remessa = 'False' and (CPL.Status_Item = 'RECEBIDO' or CPL.Status_Item = 'PARCIAL')", Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    Valor1 = IIf(IsNull(TBAbrir!Valor1), 0, TBAbrir!Valor1)
                End If
                TBOrdem!CTServico = Format(Valor1, "###,##0.00")
                TBOrdem.Update
                
                valor = valor + Valor1
                
                TBOrdem.MoveNext
                Contador = Contador + 1
                PBLista.Value = Contador
            Loop
        End If
        TBOrdem.Close
    
        'Custo terceiros utilizado pela ordem
        TBproducao!CTServico = Format(valor, "###,##0.00")
        TBproducao.Update
    End If
    TBproducao.Close
    USMsgBox ("Valor de terceiros atualizado com sucesso."), vbInformation, "CAPRIND v5.0"
    '==================================
    Modulo = "PCP/Gerenciamento de ordem"
    Evento = "Atualizar valor de terceiros utilizado na ordem"
    ID_documento = frmprod.txtof
    Documento = "Ordem: " & frmprod.txtof.Text & " - Cód. interno: " & frmprod.txtdesenho
    Documento1 = ""
    ProcGravaEvento
    '==================================
    frmprod.ProcAbrirRe
End If
        
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
