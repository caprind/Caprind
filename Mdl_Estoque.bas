Attribute VB_Name = "Mdl_Estoque"
Public Estoque As Boolean
Public ID_empresa As Integer
Public IDnota As Long
Public IDTipo As Integer

Public strUN As String
Public strGrupo As String
Public strFramilia As String

Public Movimentou As Boolean
Public Consignacao As Boolean
Public StatusRE As String

'=====================================
' Variaveis baixa requisicao de materiais do estoque
'===================================================
Public QTUnitario As Double 'Quantidade a baixar por item
Public QTBaixar As Double 'Quantidade a baixar requisitado
Public QTBaixado As Double 'Quantidade já baixado da requisicao
Public QTSaldo As Double 'Quantidade de saldo a baixar da requisicao
Public QTEmpenhado As Double 'Quantidade empenhado na requisicao
Public QTLOTE As Double 'Quantidade a produzir da ordem
Public QTEstoque As Double 'Quantidade em estoque
Public QTSaida As Double 'Quantidade baixada antes da baixa
Public QTEntrada As Double
Public QTBaixadoItemNota As Double 'Total baixado do item da nota


Public TBRE As ADODB.Recordset
Public TBEntrada As ADODB.Recordset
Public TBMovimentacao As ADODB.Recordset

Public TBListaNFe           As ADODB.Recordset
Public TBListaPI            As ADODB.Recordset

Public Function FunEstoque_SaldoItem(IDempresa As Integer, Desenho As String, Data As Date) As Double
On Error GoTo tratar_erro
    
'=================================================================================
'Dados de estoque do item
'=================================================================================
    Set TBFI = CreateObject("adodb.recordset")
    StrSql = "Select Sum(Entrada)-Sum(Saida) as Saldo from estoque_movimentacao Where Desenho = '" & Desenho & "' and ID_Empresa = '" & IDempresa & "'  and Data <= '" & Data & "' group By Desenho"
    'Debug.print StrSql
    
    TBFI.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
    
    If TBFI.EOF = False Then
        FunSaldoItem = IIf(IsNull(TBFI!Saldo), 0, TBFI!Saldo)
    End If


Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function


Public Sub procAtualizaProdutosWEB()
On Error GoTo tratar_erro

FunAbreBDSite
If ConexaoMySql.State = 1 Then

    
Set TBAbrir = CreateObject("adodb.recordset")
    
StrSql = "Select pp.Desenho,PP.Unidade,PP.Descricao,PP.Prevenda, isnull(QEPP.Estoque_disponivel,0) as Estoque_Disponivel, isnull(QEPP.Qtde_empenhada,0) as Vendido, isnull(Sum(QEPP.Estoque_Disponivel-QEPP.Qtde_empenhada),0) As Saldo from projproduto PP Left join Qtde_estoque_produto QEPP on QEPP.Desenho = pp.Desenho Where PP.Vendas = '1' group by pp.Desenho,PP.Unidade,PP.Descricao,PP.Prevenda,QEPP.Estoque_disponivel,QEPP.Qtde_empenhada"
    TBAbrir.Open StrSql & " order by Descricao", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Do While TBAbrir.EOF = False
            Set TBMySQL = New ADODB.Recordset
            '=================================================================
            ' Salvar produtos na nuvem
            '=================================================================
            StrSql = "Select * From Vendas_Produtos where Descricao = '" & TBAbrir!Descricao & "' and CNPJ_Empresa = '" & CNPJ_Empresa & "'"
            'Debug.print StrSql
             TBMySQL.Open StrSql, ConexaoMySql, adOpenKeyset, adLockOptimistic, adCmdText
              If TBMySQL.EOF = True Then
                    TBMySQL.AddNew
              End If
                    TBMySQL.Fields!CODIGO = TBAbrir!Desenho
                    TBMySQL.Fields!Descricao = TBAbrir!Descricao
                    TBMySQL.Fields!Unidade = TBAbrir!Unidade
                    TBMySQL.Fields!vlr_unit = TBAbrir!PRevenda
                    TBMySQL.Fields!Estoque = TBAbrir!Estoque_Disponivel
                    TBMySQL.Fields!Vendido = TBAbrir!Vendido
                    TBMySQL.Fields!Estoque = TBAbrir!Saldo
                    TBMySQL.Fields!CNPJ_Empresa = CNPJ_Empresa
                    
                    TBMySQL.Update
        TBAbrir.MoveNext
        Loop
        End If
        TBAbrir.Close
        'USMsgBox ("Atualização efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
End If


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Function ProcBuscaValorEstoque(IDEstoque As Integer)
On Error GoTo tratar_erro


Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Public Function ApagarMovimentacaoNFe()
On Error GoTo tratar_erro

Set TBAbrir_NFe = CreateObject("adodb.recordset")

StrSql = "Select EM.IDEstoque,EM.ID_Carteira, SUM(EM.saida) as saida , EP.Qtde_empenhada, EP.Qtde_saida from Estoque_movimentacao EM INNER JOIN Estoque_Controle_Empenho_Vendas EP ON EM.IdEstoque = EP.ID_estoque where EM.ID_Nota = '" & ID_nota & "' GROUP BY EM.IDEstoque, EM.ID_carteira, EP.Qtde_empenhada, EP.Qtde_saida"
TBAbrir_NFe.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir_NFe.EOF = False Then

Do While TBAbrir_NFe.EOF = False
Conexao.Execute "UPDATE Estoque_Controle_Empenho_Vendas SET Qtde_saida = " & Replace(TBAbrir_NFe!Qtde_saida, ",", ".") - Replace(TBAbrir_NFe!Saida, ",", ".") & " where ID_estoque = '" & TBAbrir_NFe!IDEstoque & "'"
Conexao.Execute "UPDATE vendas_carteira SET Qtdeexpedida = " & Replace(TBAbrir_NFe!Qtde_saida, ",", ".") - Replace(TBAbrir_NFe!Saida, ",", ".") & " where Codigo = '" & TBAbrir_NFe!ID_carteira & "'"
TBAbrir_NFe.MoveNext
Loop


End If
TBAbrir_NFe.Close

Conexao.Execute "Delete from Estoque_Controle where ID_Nota = '" & ID_nota & "'"
Conexao.Execute "Delete from Estoque_movimentacao where ID_Nota = '" & ID_nota & "'"
Conexao.Execute "Delete from Faturamento_ImportacaoXML where ID_Nota = '" & ID_nota & "'"

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Public Function EntrarEstoqueNF()
On Error GoTo tratar_erro

'===================================================================
' Abre a nota fiscal para entrada no estoque
'===================================================================
Set TBAbrir_NFe = CreateObject("adodb.recordset")
TBAbrir_NFe.Open "Select * from tbl_Dados_Nota_Fiscal where ID = '" & ID_nota & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir_NFe.EOF = False Then

'===================================================================
' Abre a lista da nota fiscal
'===================================================================
Set TBListaNFe = CreateObject("adodb.recordset")
TBListaNFe.Open "Select * from tbl_detalhes_nota where ID_nota = '" & ID_nota & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBListaNFe.EOF = False Then
Do While TBListaNFe.EOF = False
    Devolucao = False
    Remessa = False
'===================================================================
' Verifica se CFOP é de devolução ou remessa para industrialização
'===================================================================
Set TBCFOP = CreateObject("adodb.recordset")
TBCFOP.Open "Select Devolucao, remessa from tbl_NaturezaOperacao where IDCountCfop = '" & TBListaNFe!ID_CFOP & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBCFOP.EOF = False Then
        Devolucao = TBCFOP!Devolucao
        Remessa = TBCFOP!Remessa
    End If
TBCFOP.Close
If Devolucao = True Or Remessa = True Then
'===================================================================
' Busca os dados do produto se controla estoque para cadastro da RE
'===================================================================
status = IIf(Devolucao = True, "ENTRADA_NOTA_FISCAL_DEVOLUÇÃO", "ENTRADA_NOTA_REM_INDUST")

Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select * from projproduto where desenho = '" & TBListaNFe!int_Cod_Produto & "' and Estoque = 'TRUE' ", Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then

'===================================================================
' Cria a RE
'===================================================================
Set TBRE = CreateObject("adodb.recordset")
TBRE.Open "Select * from Estoque_controle", Conexao, adOpenKeyset, adLockOptimistic
TBRE.AddNew
TBRE!ID_nota = ID_nota
TBRE!Data = Date
TBRE!LOTE = TBListaNFe!int_NotaFiscal
TBRE!Desenho = TBListaNFe!int_Cod_Produto
TBRE!Descricao = TBListaNFe!Txt_descricao
TBRE!estoque_venda = TBListaNFe!int_Qtd
TBRE!estoque_real = TBListaNFe!int_Qtd
TBRE!Estoque_minimo = TBListaNFe!int_Qtd
TBRE!Un = TBListaNFe!txt_Unid
TBRE!Responsavel = pubUsuario
TBRE!Fornecedor = TBAbrir_NFe!txt_Razao_Nome
TBRE!Certificado = ""
TBRE!Classe = TBProduto!Classe
TBRE!descricaotecnica = TBProduto!descricaotecnica
TBRE!peso_unit = TBProduto!PBruto
TBRE!local_armaz = LocalArmazenamento
TBRE!Consignacao = Consignacao
TBRE!status = StatusRE
TBRE!Qtde = TBListaNFe!int_Qtd
TBRE!NF = TBListaNFe!int_NotaFiscal
TBRE!ID_Cliente = TBAbrir_NFe!Id_Int_Cliente
TBRE!Cliente = TBAbrir_NFe!txt_Razao_Nome
TBRE!Ref = ""
TBRE!Corrida = ""
TBRE!emissaonf = TBAbrir_NFe!dt_DataEmissao
TBRE!valor_unitario = TBListaNFe!dbl_ValorUnitario
TBRE!Valor_total = TBListaNFe!dbl_ValorTotal
TBRE!qtde_fisica = TBListaNFe!int_Qtd
TBRE!ID_empresa = ID_empresa
TBRE!Bloqueado = False
TBRE!resp_Status = pubUsuario
TBRE.Update

'===============================================
'Criar entrada na movimentação de Estoque
'===============================================

Set TBMovimentacao = CreateObject("adodb.recordset")
TBMovimentacao.Open "Select * from Estoque_movimentacao", Conexao, adOpenKeyset, adLockOptimistic
TBMovimentacao.AddNew
TBMovimentacao!IDEstoque = TBRE!IDEstoque
TBMovimentacao!ID_nota = ID_nota
TBMovimentacao!Operacao = status
TBMovimentacao!Desenho = TBListaNFe!int_Cod_Produto
TBMovimentacao!Descricao = TBListaNFe!Txt_descricao
TBMovimentacao!Data = Date
TBMovimentacao!Entrada = TBListaNFe!int_Qtd
TBMovimentacao!Responsavel = pubUsuario
TBMovimentacao!Cliente = TBAbrir_NFe!txt_Razao_Nome
TBMovimentacao!Documento = TBListaNFe!int_NotaFiscal
TBMovimentacao!estoque_venda = TBListaNFe!int_Qtd
TBMovimentacao!LOTE = TBListaNFe!int_NotaFiscal
TBMovimentacao!DtEmissao = TBAbrir_NFe!dt_DataEmissao
TBMovimentacao!Ordem = 0
TBMovimentacao!OE = TBListaNFe!int_NotaFiscal
TBMovimentacao!VlrUnit = TBListaNFe!dbl_ValorUnitario
TBMovimentacao!vlrTotal = TBListaNFe!dbl_ValorTotal
TBMovimentacao!IDEstoque_recebimento = 0
TBMovimentacao!idlista_recebimento = 0
TBMovimentacao!Familia = TBProduto!Classe
TBMovimentacao!Destino = "Interno"
TBMovimentacao!IDpedido = 0
TBMovimentacao!Pedidocompra = ""
TBMovimentacao!Terceiros = False
TBMovimentacao!Id_cfi = 0
TBMovimentacao!Obs = ""
TBMovimentacao!ID_prod_NF = TBListaNFe!Int_codigo
TBMovimentacao!ID_inventario = 0
TBMovimentacao!ID_prod_RM = 0
TBMovimentacao!ID_Tipo = TBProduto!ID_Tipo
TBMovimentacao.Update
TBMovimentacao.Close

TBRE.Close
'End If


TBProduto.Close
End If
End If
TBListaNFe.MoveNext
Loop
End If
TBListaNFe.Close
End If
TBAbrir_NFe.Close


Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Public Function FunRecebePedidoCompra()
On Error GoTo tratar_erro

Set TBAbrir_Pedido = CreateObject("adodb.recordset")
TBAbrir_Pedido.Open "Select * from Compras_pedido where IDPedido = '" & IDpedido & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir_Pedido.EOF = False Then

                
Set TBListaPedido = CreateObject("adodb.recordset")
TBListaPedido.Open "Select * from Compras_pedido_lista where IDlista = '" & IDlista & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBListaPedido.EOF = False Then
'Busca os dados do produto para cadastro da RE
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select PP.*, PF.Grupo from projproduto PP inner join Projfamilia PF on PP.Classe = PF.Familia where PP.desenho = '" & TBListaPedido!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then

'Cria a RE
Set TBRE = CreateObject("adodb.recordset")
TBRE.Open "Select * from Estoque_controle", Conexao, adOpenKeyset, adLockOptimistic
TBRE.AddNew
TBRE!Data = Date
TBRE!LOTE = TBAbrir_Pedido!Pedido
TBRE!Desenho = TBListaPedido!Desenho
TBRE!Descricao = TBListaPedido!Descricao
TBRE!estoque_venda = TBListaPedido!Quant_Comp
TBRE!estoque_real = TBListaPedido!Quant_Comp
TBRE!Estoque_minimo = TBListaPedido!Quant_Comp
TBRE!Un = TBListaPedido!Un
TBRE!Responsavel = pubUsuario
TBRE!Fornecedor = TBAbrir_Pedido!Fornecedor
TBRE!Certificado = ""
TBRE!Classe = TBProduto!Classe
TBRE!descricaotecnica = TBProduto!descricaotecnica
TBRE!peso_unit = TBProduto!PBruto

If frmEstoque_Recebimento.Chk_LA.Value = 0 Then
    TBRE!local_armaz = "ESTOQUE PADRÃO"
Else
    TBRE!local_armaz = frmEstoque_Recebimento.cmbLocal_armaz.Text
End If

TBRE!status = "ENTRADA_NOTA_FISCAL"
TBRE!Qtde = TBListaPedido!Quant_Comp
TBRE!NF = frmEstoque_Recebimento.txtnotafiscal.Text
TBRE!ID_Cliente = TBAbrir_Pedido!IDFornecedor
TBRE!Cliente = TBAbrir_Pedido!Fornecedor
TBRE!Ref = ""
TBRE!Corrida = ""
TBRE!emissaonf = frmEstoque_Recebimento.txtDataemissao.Text
TBRE!Consignacao = False
TBRE!valor_unitario = TBListaPedido!preco_unitario
TBRE!Valor_total = TBListaPedido!preco_total
TBRE!qtde_fisica = TBListaPedido!Quant_Comp
TBRE!ID_empresa = TBAbrir_Pedido!ID_empresa
TBRE!Bloqueado = False
TBRE!resp_Status = pubUsuario
TBRE.Update

'===========================================================
'Grava movimentação na tabela estoque_controle_recebimento
'===========================================================
Set TBEstoque = CreateObject("adodb.recordset")
TBEstoque.Open "Select * from estoque_controle_recebimento", Conexao, adOpenKeyset, adLockOptimistic
TBEstoque.AddNew

If frmEstoque_Recebimento.Chk_Dt_rcbto.Value = 0 Then
TBEstoque!Data_recebimento = Date
Else
TBEstoque!Data_recebimento = frmEstoque_Recebimento.Txt_data_recebimento
End If

TBEstoque!IDpedido = TBListaPedido!IDpedido
TBEstoque!IDlista = TBListaPedido!IDlista
TBEstoque!Desenho = TBListaPedido!Desenho
TBEstoque!Certificado = ""
TBEstoque!Corrida = ""

If frmEstoque_Recebimento.Chk_LA.Value = 0 Then
    TBEstoque!local_armaz = "ESTOQUE PADRÃO"
Else
    TBEstoque!local_armaz = frmEstoque_Recebimento.cmbLocal_armaz.Text
End If

TBEstoque!Nota_fiscal = frmEstoque_Recebimento.txtnotafiscal.Text
TBEstoque!Serie = frmEstoque_Recebimento.txtSerie
TBEstoque!Data_emissao = frmEstoque_Recebimento.txtDataemissao
TBEstoque!Responsavel = pubUsuario
TBEstoque!Recebido = Format(TBListaPedido!Quant_Comp, "###.##0.000")
TBEstoque!Recebido_PC = Format(TBListaPedido!Quant_Comp, "###.##0.000")
TBEstoque!Parcial = False
TBEstoque!Programacao = False
TBEstoque!ID_empresa = TBAbrir_Pedido!ID_empresa
TBEstoque!Obs = ""
TBEstoque.Update
IDEstoque_recebimento = TBEstoque!ID
TBEstoque.Close

'frmEstoque_Recebimento.ProcAtualizaQtdeRecebEmp Txt_ID_pedido, txtidlista
'===============================================
'Criar entrada na movimentação de Estoque
'===============================================

Set TBMovimentacao = CreateObject("adodb.recordset")
TBMovimentacao.Open "Select * from Estoque_movimentacao", Conexao, adOpenKeyset, adLockOptimistic
TBMovimentacao.AddNew
TBMovimentacao!IDEstoque = TBRE!IDEstoque
TBMovimentacao!Operacao = "ENTRADA_NOTA_FISCAL"
TBMovimentacao!Desenho = TBListaPedido!Desenho
TBMovimentacao!Descricao = TBListaPedido!Descricao

Desenho = TBListaPedido!Desenho

If frmEstoque_Recebimento.Chk_Dt_rcbto.Value = 0 Then
TBMovimentacao!Data = Date
Else
TBMovimentacao!Data = frmEstoque_Recebimento.Txt_data_recebimento
End If

TBMovimentacao!Entrada = TBListaPedido!Quant_Comp
TBMovimentacao!Responsavel = pubUsuario
TBMovimentacao!Cliente = TBAbrir_Pedido!Fornecedor
TBMovimentacao!Documento = frmEstoque_Recebimento.txtnotafiscal.Text
TBMovimentacao!estoque_venda = TBListaPedido!Quant_Comp
TBMovimentacao!LOTE = TBAbrir_Pedido!Pedido
TBMovimentacao!DtEmissao = TBAbrir_Pedido!Data
TBMovimentacao!Ordem = 0
TBMovimentacao!OE = frmEstoque_Recebimento.txtnotafiscal.Text
TBMovimentacao!VlrUnit = TBListaPedido!preco_unitario
TBMovimentacao!vlrTotal = TBListaPedido!preco_total
TBMovimentacao!IDEstoque_recebimento = 0
TBMovimentacao!idlista_recebimento = 0
TBMovimentacao!Familia = TBProduto!Classe
TBMovimentacao!Destino = "Interno"
TBMovimentacao!IDpedido = IDpedido
TBMovimentacao!Pedidocompra = TBAbrir_Pedido!Pedido
TBMovimentacao!Terceiros = False
TBMovimentacao!Id_cfi = 0
TBMovimentacao!Obs = ""
TBMovimentacao!ID_prod_NF = 0
TBMovimentacao!ID_inventario = 0
TBMovimentacao!ID_prod_RM = 0
TBMovimentacao!ID_Tipo = TBProduto!ID_Tipo

If frmEstoque_Recebimento.Chk_LA.Value = 0 Then
    TBMovimentacao!local_armaz = "ESTOQUE PADRÃO"
Else
    TBMovimentacao!local_armaz = frmEstoque_Recebimento.cmbLocal_armaz.Text
End If

TBMovimentacao!Grupo = TBProduto!Classe
TBMovimentacao!ID_empresa = TBAbrir_Pedido!ID_empresa
TBMovimentacao!IDEstoque_recebimento = IDEstoque_recebimento
TBMovimentacao.Update

TBMovimentacao.Close
TBRE.Close

End If
TBListaPedido!Status_Item = "RECEBIDO"
TBListaPedido.Update
TBListaPedido.Close
End If
TBAbrir_Pedido.Close
End If
TBProduto.Close

'==================================================================
'Gravar evento realizado pelo usuario
'==================================================================
Modulo = "Estoque/Recebimento/Pedido de compra"
Evento = "Receber"
ID_documento = 0 ' frmEstoque_Recebimento.TXTIDLista
Documento = "Cód. interno: " & Desenho & " - Nº lote: " & frmEstoque_Recebimento.txtProg_pedido & " - Nº corrida: " & IIf(frmEstoque_Recebimento.txtcorrida = "", 0, frmEstoque_Recebimento.txtcorrida) & " - Nº certificado: " & IIf(frmEstoque_Recebimento.txtCertificado = "", 0, frmEstoque_Recebimento.txtCertificado) & " - Local armaz.: " & frmEstoque_Recebimento.cmbLocal_armaz
Documento1 = "Operação: ENTRADA_NOTA_FISCAL - Documento: " & frmEstoque_Recebimento.txtnotafiscal.Text
ProcGravaEvento
'==================================================================
Desenho = ""

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Public Function CriarRE()
On Error GoTo tratar_erro

Set TBAbrir_NFe = CreateObject("adodb.recordset")
TBAbrir_NFe.Open "Select * from tbl_Dados_Nota_Fiscal where ID = '" & ID_nota & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir_NFe.EOF = False Then

Set TBListaNFe = CreateObject("adodb.recordset")
TBListaNFe.Open "Select * from tbl_detalhes_nota where ID_nota = '" & ID_nota & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBListaNFe.EOF = False Then
Do While TBListaNFe.EOF = False
'Busca os dados do produto para cadastro da RE
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select * from projproduto where desenho = '" & TBListaNFe!int_Cod_Produto & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then

'Cria a RE
Set TBRE = CreateObject("adodb.recordset")
TBRE.Open "Select * from Estoque_controle", Conexao, adOpenKeyset, adLockOptimistic
TBRE.AddNew
TBRE!LOTE = TBListaNFe!int_NotaFiscal
TBRE!Desenho = TBListaNFe!int_Cod_Produto
TBRE!Descricao = TBListaNFe!Txt_descricao
TBRE!estoque_venda = TBListaNFe!int_Qtd
TBRE!estoque_real = TBListaNFe!int_Qtd
TBRE!Estoque_minimo = TBListaNFe!int_Qtd
TBRE!Un = TBListaNFe!txt_Unid
TBRE!Responsavel = pubUsuario
TBRE!Fornecedor = TBAbrir_NFe!txt_Razao_Nome
TBRE!Certificado = ""
TBRE!Classe = TBProduto!Classe
TBRE!descricaotecnica = TBProduto!descricaotecnica
TBRE!peso_unit = TBProduto!PBruto
TBRE!local_armaz = "ESTOQUE PADRÃO"
TBRE!status = "ENTRADA_NOTA_FISCAL"
TBRE!Qtde = TBListaNFe!int_Qtd
TBRE!NF = TBListaNFe!int_NotaFiscal
TBRE!ID_Cliente = TBAbrir_NFe!Id_Int_Cliente
TBRE!Cliente = TBAbrir_NFe!txt_Razao_Nome
TBRE!Ref = ""
TBRE!Corrida = ""
TBRE!emissaonf = TBAbrir_NFe!dt_DataEmissao
TBRE!Consignacao = False
TBRE!valor_unitario = TBListaNFe!dbl_ValorUnitario
TBRE!Valor_total = TBListaNFe!dbl_ValorTotal
TBRE!qtde_fisica = TBListaNFe!int_Qtd
TBRE!ID_empresa = ID_empresa
TBRE!Bloqueado = False
TBRE!resp_Status = pubUsuario
TBRE.Update

'===============================================
'Criar entrada na movimentação de Estoque
'===============================================

Set TBMovimentacao = CreateObject("adodb.recordset")
TBMovimentacao.Open "Select * from Estoque_movimentacao", Conexao, adOpenKeyset, adLockOptimistic
TBMovimentacao.AddNew
TBMovimentacao!IDEstoque = TBRE!IDEstoque
TBMovimentacao!Operacao = "ENTRADA_NOTA_FISCAL"
TBMovimentacao!Desenho = TBListaNFe!int_Cod_Produto
TBMovimentacao!Descricao = TBListaNFe!Txt_descricao
TBMovimentacao!Data = Date
TBMovimentacao!Entrada = TBListaNFe!int_Qtd
TBMovimentacao!Responsavel = pubUsuario
TBMovimentacao!Cliente = TBAbrir_NFe!txt_Razao_Nome
TBMovimentacao!Documento = TBListaNFe!int_NotaFiscal
TBMovimentacao!estoque_venda = TBListaNFe!int_Qtd
TBMovimentacao!LOTE = TBListaNFe!int_NotaFiscal
TBMovimentacao!DtEmissao = TBAbrir_NFe!dt_DataEmissao
TBMovimentacao!Ordem = 0
TBMovimentacao!OE = TBListaNFe!int_NotaFiscal
TBMovimentacao!VlrUnit = TBListaNFe!dbl_ValorUnitario
TBMovimentacao!vlrTotal = TBListaNFe!dbl_ValorTotal
TBMovimentacao!IDEstoque_recebimento = 0
TBMovimentacao!idlista_recebimento = 0
TBMovimentacao!Familia = TBProduto!Classe
TBMovimentacao!Destino = "Interno"
TBMovimentacao!IDpedido = 0
TBMovimentacao!Pedidocompra = ""
TBMovimentacao!Terceiros = False
TBMovimentacao!Id_cfi = 0
TBMovimentacao!Obs = ""
TBMovimentacao!ID_prod_NF = TBListaNFe!Int_codigo
TBMovimentacao!ID_inventario = 0
TBMovimentacao!ID_prod_RM = 0
TBMovimentacao!ID_Tipo = TBProduto!ID_Tipo
TBMovimentacao.Update
End If
TBMovimentacao.Close
TBRE.Close
TBProduto.Close
TBListaNFe.MoveNext
Loop
End If
TBListaNFe.Close
End If
TBAbrir_NFe.Close

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Public Function CriarMovEntrada()
On Error GoTo tratar_erro


Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Public Function BaixarEstoqueNF()
On Error GoTo tratar_erro

'Verifica se movimenta estoque
Set TBListaNFe = CreateObject("adodb.recordset")
StrSql = "Select TBL.Int_codigo, TBL.int_Cod_Produto, TBL.int_Qtd, TBL.txt_Unid, TBL.Unidade_com, TBL.qtde_estoque, TBL.ID_nota, TBL.N_Referencia,TBL.dbl_ValorUnitario, TBL.Int_NotaFiscal, P.ID_PC, P.Codproduto from (tbl_detalhes_nota TBL INNER JOIN ProjProduto P ON TBL.Codproduto = P.Codproduto) INNER JOIN tbl_NaturezaOperacao CFOP ON CFOP.IDCountCfop = TBL.ID_CFOP where TBL.ID_Nota = " & ID_nota & " and P.Estoque = 'True' and (CFOP.Retorno IS NULL or CFOP.Retorno = 'False')"
'Debug.print StrSql
'=================================================================
' Busca o item na lista da nota com CFOP que não seja retorno
'=================================================================
TBListaNFe.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
 Do While TBListaNFe.EOF = False
 QTBaixar = TBListaNFe!int_Qtd
  Set TBEstoque = CreateObject("adodb.recordset")
  
  StrSql = "SELECT * from Estoque_Controle_Saldo_RE WHERE Codigo = '" & TBListaNFe!int_Cod_Produto & "' AND Saldo > 0 AND local_armaz <> 'DEGUSTAÇÃO' "
   'Debug.print StrSql
   TBEstoque.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
    If TBEstoque.EOF = False Then
    VerificaSaldoBaixarItem (TBListaNFe!Int_codigo)
    
     If QTBaixadoItemNota < QTBaixar Then
        SaidaEstoqueNFItem
        Sair = True
     Else
        Sair = False
     End If
    End If
  TBEstoque.Close
 TBListaNFe.MoveNext
 Loop
TBListaNFe.Close

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Public Function SaidaEstoqueNFItem()
On Error GoTo tratar_erro

CampoFiltro = "Saida"
QTBaixar = TBListaNFe!int_Qtd
QTBaixado = 0

Do While TBEstoque.EOF = False And QTBaixar > 0
'Debug.print TBEstoque!IDEstoque
QTSaldo = IIf(TBEstoque!Saldo <> "", TBEstoque!Saldo, o)

Set TBListaPI = CreateObject("adodb.recordset")
TBListaPI.Open "Select * from tbl_Detalhes_Nota_pedidos where ID_prod_NF = " & TBListaNFe!Int_codigo, Conexao, adOpenKeyset, adLockOptimistic
QtdeSaida = QTBaixar
'=========================================================================
'Se existe vinculo da nota com pedido interno, atualiza saida no pedido
'=========================================================================
If TBListaPI.EOF = False Then
'==================================================================
'Verifica se tem empenho
'==================================================================
Set TBEmpenho = CreateObject("adodb.recordset")
StrSql = "Select ECEV.*, EC.Lote from Estoque_Controle_Empenho_Vendas ECEV Inner Join Estoque_Controle EC on ECEV.ID_estoque = EC.IdEstoque where ID_Carteira = " & TBListaPI!ID_carteira & " and QTde_saida = '0'"
'Debug.print StrSql

TBEmpenho.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
If TBEmpenho.EOF = False Then
REempenho = TBEmpenho!ID_estoque
Loteempenho = TBEmpenho!LOTE
QTEmpenho = TBEmpenho!Qtde_empenhada
Else
REempenho = ""
Loteempenho = ""
QTEmpenho = ""
End If
ID_carteira = TBListaPI!ID_carteira
End If

'============================================================================
' Verifica se usa código interno ou codigo do cliente na NFe
'============================================================================
    If FunVerifCodRefCliDANFE(IDempresa) = True Then
        Conexao.Execute "UPDATE Estoque_controle Set REF = '" & TBListaNFe!N_referencia & "' where IDestoque = " & TBEstoque!IDEstoque
    End If
 '===========================================================================
 'Acerta saldo e valores (Valor total) na RE
 '===========================================================================
  Set TBProduto = CreateObject("adodb.recordset")
  TBProduto.Open "select PP.*, PF.Grupo from projproduto PP inner join Projfamilia PF on PP.Classe = PF.Familia where Desenho = '" & TBEstoque!CODIGO & "'", Conexao, adOpenKeyset, adLockOptimistic
   If TBProduto.EOF = False Then
   SaldoREantigo = TBEstoque!Saldo
   If TBEstoque!Saldo > QTBaixar Then
   'Debug.print TBEstoque!IDEstoque
   SaldoRE = TBEstoque!Saldo - QTBaixar
   Else
   SaldoRE = 0
   End If
   
   IDTipo = IIf(IsNull(TBProduto!ID_Tipo), 0, TBProduto!ID_Tipo)
   strUN = TBProduto!Unidade
   strGrupo = TBProduto!Grupo
   strFamilia = TBProduto!Classe
   
   SaldoRE = Replace(SaldoRE, ",", ".")
   NovoValor1 = Replace(QuantEmpenhoPC, ",", ".")
   
   Conexao.Execute "UPDATE Estoque_controle Set Estoque_real = " & SaldoRE & ", Estoque_real_PC = " & NovoValor1 & ", Estoque_venda = " & SaldoRE & ", peso_unit = '" & IIf(IsNull(TBProduto!peso_metro), "", TBProduto!peso_metro) & "', Pedido = '" & IIf(IsNull(TBProduto!Un_Kg), "", TBProduto!Un_Kg) & "' where IDestoque = " & TBEstoque!IDEstoque
   Conexao.Execute "UPDATE Estoque_controle Set Valor_Total = ROUND(valor_unitario * Estoque_real, 2) where IDestoque = " & TBEstoque!IDEstoque
   End If
  TBProduto.Close
  
 '===========================================================================
 'Acerta saldo do empenho
 '===========================================================================
 IDEmp = 0
 If ID_carteira <> "" And Loteempenho <> "" Then
 Set TBAbrir = CreateObject("adodb.recordset")
 TBAbrir.Open "Select * from Estoque_Controle_Empenho_Vendas where id_Carteira = " & ID_carteira & " and ID_estoque = " & REempenho, Conexao, adOpenKeyset, adLockOptimistic
  If TBAbrir.EOF = False Then
  IDEmp = TBAbrir!ID
    If TBAbrir!Qtde_empenhada <= QTBaixar Then
     TBAbrir!Qtde_saida = TBAbrir!Qtde_empenhada
     QTEmpenhado = TBAbrir!Qtde_empenhada
    Else
     TBAbrir!Qtde_saida = TBAbrir!Qtde_saida + QTBaixar
    End If
    TBAbrir.Update
   End If
 TBAbrir.Close
 End If
 
 
 '===========================================================================
 'Cria a movimentação de saida do estoque
 '===========================================================================
    Set TBMovimentacao = CreateObject("adodb.recordset")
    TBMovimentacao.Open "Select * from Estoque_movimentacao", Conexao, adOpenKeyset, adLockOptimistic
    TBMovimentacao.AddNew
    TBMovimentacao!ID_nota = ID_nota
    TBMovimentacao!Destino = "Interno"
    TBMovimentacao!Terceiros = False
    TBMovimentacao!Documento = TBListaNFe!int_NotaFiscal
    TBMovimentacao!LOTE = IIf(Loteempenho = "", TBEstoque!LOTE, Loteempenho)
    TBMovimentacao!Desenho = TBEstoque!CODIGO
    TBMovimentacao!Data = frmFaturamento_Prod_Serv.txt_DtEmissao.Text
    TBMovimentacao!Descricao = IIf(IsNull(TBEstoque!Descricao), "", TBEstoque!Descricao)
    TBMovimentacao!Familia = IIf(IsNull(TBEstoque!Classe), "", TBEstoque!Classe)
    TBMovimentacao!Requisitante = pubUsuario
    TBMovimentacao!Responsavel = pubUsuario
    TBMovimentacao!IDEstoque = IIf(REempenho = "", TBEstoque!IDEstoque, REempenho)
    TBMovimentacao!OE = TBListaNFe!int_NotaFiscal
    
'====================================================================
' Grava dados  na movimentação
'====================================================================
   TBMovimentacao!ID_Tipo = IDTipo
   TBMovimentacao!Unidade = strUN
   TBMovimentacao!Grupo = strGrupo
   TBMovimentacao!Familia = strFamilia
   TBMovimentacao!ID_carteira = IIf(ID_carteira <> "", ID_carteira, 0)
'====================================================================
    
    TBMovimentacao!ID_prod_NF = TBListaNFe!Int_codigo
    TBMovimentacao!Operacao = "SAIDA_NOTA"
    
   If REempenho <> "" Then
    TBMovimentacao!Saida = QTEmpenhado
    TBMovimentacao!estoque_venda = TBMovimentacao!Saida
    'TBProduto!VlrTotal = Format(TBProduto!Saida * TBListaNFe!dbl_ValorUnitario, "###,##0.000000")
    TBMovimentacao!vlrTotal = Format(TBMovimentacao!Saida * TBEstoque!valor_unitario, "###,##0.000000")
   Else
       If TBEstoque!Saldo >= QTBaixar Then
        TBMovimentacao!Saida = QTBaixar
        TBMovimentacao!estoque_venda = TBMovimentacao!Saida
        'TBProduto!VlrTotal = Format(TBProduto!Saida * TBListaNFe!dbl_ValorUnitario, "###,##0.000000")
        TBMovimentacao!vlrTotal = Format(TBMovimentacao!Saida * TBEstoque!valor_unitario, "###,##0.000000")
       Else
        TBMovimentacao!Saida = SaldoREantigo
        TBMovimentacao!estoque_venda = SaldoREantigo
        'TBProduto!VlrTotal = Format(SaldoREantigo * TBListaNFe!dbl_ValorUnitario, "###,##0.000000")
        TBMovimentacao!vlrTotal = Format(SaldoREantigo * TBEstoque!valor_unitario, "###,##0.000000")
       End If
   End If
   QTBaixar = QTBaixar - TBMovimentacao!Saida
   QTBaixado = QTBaixado + TBMovimentacao!Saida
   
    'Atualiza valor do material no estoque
    TBMovimentacao!VlrUnit = Format(TBEstoque!valor_unitario, "###,##0.000000") 'Format(TBListaNFe!dbl_ValorUnitario, "###,##0.000000")
'===================================================
TBMovimentacao.Update
'==================================
Modulo = "Faturamento/Nota fiscal/Própria"
Evento = "Retirar"
ID_documento = TBListaNFe!Int_codigo
Documento = "Cód. interno: " & TBEstoque!CODIGO & " - RE: " & TBEstoque!IDEstoque
Documento1 = ""
ProcGravaEvento
'==================================
Permitido1 = True

'Centro de custo
ProcCriaCreditoCCProdutoItemSelecionada TBListaNFe!Codproduto, Date, IDempresa, TBMovimentacao!IDoperacao, TBMovimentacao!vlrTotal

TBMovimentacao.Close
Proximo:
'End If
TBEstoque.MoveNext
Loop
'=================================================================================================================
' Acerta Saldo expedido na carteira de vendas
'=================================================================================================================
 Conexao.Execute "UPDATE vendas_carteira SET Qtdeexpedida = " & Replace(QTBaixado, ",", ".") & " where Codigo = '" & ID_carteira & "'"
'=================================================================================================================

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Public Function VerificaSaldoBaixarItem(IDitem As Long) As Double
On Error GoTo tratar_erro

Set TBAfericao = CreateObject("adodb.recordset")
StrSql = "Select ISNULL(sum(Saida), 0) as Saldo from Estoque_movimentacao where ID_prod_NF =" & IDitem & ""
'Debug.print StrSql

TBAfericao.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic

If TBAfericao.EOF = False Then
    QTBaixadoItemNota = IIf(TBAfericao!Saldo <> "", TBAfericao!Saldo, 0)
Else
    QTBaixadoItemNota = 0
End If
TBAfericao.Close

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Public Sub ProcVeficaSaldoEstoqueListaNF(IDnota As Long)
On Error GoTo tratar_erro

Set TBLISTA = CreateObject("adodb.recordset")
StrSql = "Select TDN.ID_Nota , TDN.int_Cod_Produto,TDN.int_Qtd as Saida,EC.saldo As Saldo from tbl_Detalhes_Nota TDN inner join Estoque_Controle_Saldo_Item EC on EC.Desenho = TDN.int_Cod_Produto Where TDN.ID_Nota = '" & IDnota & "' and TDN.remessa = 'False' and TDN.retorno = 'False'"

'Debug.print StrSql

TBLISTA.Open StrSql, Conexao, adOpenKeyset, adLockReadOnly

If TBLISTA.EOF = False Then


Do While TBLISTA.EOF = False

If TBLISTA!Saldo < TBLISTA!Saida Then
    USMsgBox "ATENÇÃO" & vbCrLf & vbCrLf & " ITEM CÓDIGO: " & TBLISTA!int_Cod_Produto & vbCrLf & vbCrLf & "SALDO INSUFICIENTE NO ESTOQUE!" & vbCrLf & vbCrLf & "Solicitado : " & Format(TBLISTA!Saida, "###,##0.00") & vbCrLf & " Saldo estoque : " & Format(TBLISTA!Saldo, "###,##0.00") & vbCrLf & vbCrLf & "Por favor revisar a quantidade ou excluir o item da lista.", vbCritical, "CAPRIND v5.0"
    Validar = False
    TBLISTA.Close
Exit Sub
End If

TBLISTA.MoveNext
Loop

End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Function VerificaFaturarSemSaldo(IDempresa As Integer) As Boolean
On Error GoTo tratar_erro


Set TBAcessos = CreateObject("adodb.recordset")
TBAcessos.Open "Select SemEstoque from empresa where codigo = '" & IDempresa & "'", Conexao, adOpenKeyset, adLockReadOnly
If TBAcessos.EOF = False Then
VerificaFaturarSemSaldo = IIf(IsNull(TBAcessos!SemEstoque), 0, TBAcessos!SemEstoque)
End If
TBAcessos.Close

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Public Function VerificaSaldoItem(CodigoItem As String) As Double
On Error GoTo tratar_erro

Set TBAfericao = CreateObject("adodb.recordset")
StrSql = "Select Sum(Entrada) - Sum(Saida) as Saldo from Estoque_movimentacao where Desenho ='" & CodigoItem & "'"
'Debug.print StrSql

TBAfericao.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic

If TBAfericao.EOF = False Then
Saldo_Atual = IIf(TBAfericao!Saldo <> "", TBAfericao!Saldo, 0)
Else
Saldo_Atual = 0
End If
TBAfericao.Close

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

