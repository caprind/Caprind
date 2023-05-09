Attribute VB_Name = "Mdl_caprind"
'=====================================================
Public Const SWP_FRAMECHANGED = &H20
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_SHOWME = SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOSIZE
Public Const HWND_NOTOPMOST = -2
Declare Function SetWindowLong Lib "user32" _
Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, _
ByVal dwNewLong As Long) As Long
'=============================================================
' Desabilitar "x" do MDI - 2021
'=============================================================
Public Declare Function DeleteMenu Lib "user32" _
(ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long

Public Declare Function GetSystemMenu Lib "user32" _
(ByVal hWnd As Long, ByVal bRevert As Long) As Long

Public Const MF_BYPOSITION = &H400&

Declare Function SetWindowPos Lib "user32" _
(ByVal hWnd As Long, ByVal hWndInsertAfter _
As Long, ByVal x As Long, ByVal Y As Long, _
ByVal cx As Long, ByVal cy As Long, _
ByVal wFlags As Long) As Long

'==============================================================
' Variaveis de emiss„o e retorno de NFe
'==============================================================
Public UF_Empresa As String
Public UF_Destinatario As String
Public NFCe As Boolean ' Informa se trata-se de uma NFCe ou n„o
Public RetornoXML As String
Public resposta As String
Public statusEnvio As String
Public statusConsulta As String
Public statusDownload As String
Public motivo As String
Public xMotivo As String
Public erros As String
Public nsNRec As String
Public chNFe As String
Public cStat As String
Public nProt As String
'Public JSON As String
Public url As String
'Public XML As String
Public pdf As String
Public chNFeCanc As String
Public respostaDownload As String
'==============================================================
Public ID_Tipo As Long
Public Familia As String
Public LocalArmazenamento As String
Public codAutomatico As Boolean
'==============================================================
Public Tipo_Produto As Boolean
Public TipoServico As Boolean
'==============================================================
Public Individual As Boolean
Public OrdemRastreavel As Boolean
Public OSRastreavel As Boolean
Public NumeroSerie As String
Public TotalAnexado As Integer
'==============================================================
Public LiberarData As Boolean
'==============================================================
Public TemaCaprind As String
Public TemaINI As String
'==============================================================
Public CNPJ_Empresa As String
'==============================================================
Public TotalICMS As Double
Public TotalIPI As Double
'==============================================================
Public NProposta As String
Public VendedorExterno As String
Public DataVenda As String
Public condicoes As String
Public Observacoes As String
Public CNPJCliente As String
'=================================================
' Variaveis da importaÁ„o dados cliente Sintegra
'=================================================
Public NomeRazao As String
Public NomeContato As String
Public UF As String
Public Bairro As String
Public Endereco As String
Public Numero As String
Public CEP As String
Public Cidade As String
Public NomeFantasia As String
Public RegimeTributario As String
Public RG_IE As String
Public CPF_CNPJ As String
Public Categoria As String
Public SintTelefone As String
Public SintEmail As String
Public IDPedidoWEB As Integer
Public TipoEmpresa As String

'===========================================
' Variavel de recebimento de nota
'===========================================
Public NF_Recebimento As Boolean
'===========================================
' Variaveis da ordem de faturamento
'===========================================
Public Faturamento_PI As Boolean
Public Faturamento_Produtos As Boolean
Public Faturamento_Comercial As Boolean
Public StrSQL_OF As String
Public StrSQLTotais As String
Public PaginaAtual As Integer

Public ID_CFOP As Long

Public Agrupar_Produtos As Boolean
Public FormulaRelOF As String
'===========================================
'
'===========================================
Public LiberarAlteracao As Boolean
Public AlterarProduto As Boolean
Public Tipo_Nota As String

Public UseNomeFantasia As Boolean

Public vICMSCST As String

Public Transporte1 As Boolean
Public Transporte2 As Boolean

'====================================================
Public StatusItem As String
Public StatusItemRel As String
'====================================================
' Variaveis da empresa
'====================================================
Public IDempresa As Integer
Public TabelaSN As Integer
Public RegimeEmpresa As Integer
Public AliquotaSN As Double
'====================================================

Public FormulaRelatorio As String
Public FormulaRelatorioCampo As String

Public VarST As Boolean


Public ComPedido As Boolean
Public Liberado As Boolean
'=======================================================
' Variavel de controle de Estoque
'=======================================================
Public RE As Long

Public TotalEntrada As Double
Public TotalSaida As Double
Public TotalEmpenho As Double
Public TotalSaldo As Double

Public EstoqueSaida As Double
Public EstoqueEntrada As Double
Public EstoqueEmpenho As Double
Public EstoqueSaldo As Double
'=======================================================
' Variavel de controle de ediÁ„o
'=======================================================
Public continuar As Boolean
'=======================================================
' Variavel de controle da nota fiscal
'=======================================================
Public NotaFiscalPronta As Boolean
Public DescricaoProduto As String

'=====================================
'Variaveis de frete
'=====================================
Public FRETE_ICMS As Boolean
Public CIF As Boolean
Public FOB As Boolean
'====================================
' VARIAVEIS BOLETO BANC¡RIO         =
'====================================
Public CobreBemX As CobreBemX.ContaCorrente 'OK
Public CobreBemX1 As New ContaCorrente


Public Titulosselecionados  As Integer

Public nfDocumento          As String
Public ArquivoLicensa       As String
Public Email_Contato        As String
Public EmailCopia           As String
Public Tipo_Documento       As String
Public Diretorio            As String 'OK
Public Arquivo              As String 'OK
Public Layout               As String 'OK
Public Agencia              As String 'OK
Public ContaCorrente        As String 'OK
Public NomeAgencia          As String 'OK
Public OutrosDadosConfiguracao1 As String 'OK
Public OutrosDadosConfiguracao2 As String 'OK
Public Instrucoes           As String 'OK
Public Remessa              As Boolean 'OK
Public Enviar_Email         As Boolean 'OK
Public Seq                  As Long 'OK
Public Especie              As String 'OK
'====================================
' Editado em 22/05/2019             =
'====================================
' VARIAVEIS EMPRESA - 2019          =
'====================================
Public NomeEmpresa          As String
Public IDEmp                As Integer
Public TPCertificado        As String
Public SerialCertificado    As String
Public DiretorioEnvio       As String
Public DiretorioXMLDanfe    As String
Public DiretorioDanfe       As String
Public DiretorioXML         As String
Public DiretorioRetorno     As String
Public NomeArquivo          As String

Public ID_nota              As Long
Public ID_produto_nota As Long

Public CnpjNF               As String
Public CodUF                As String
Public ClienteVendedor      As Boolean
Public SemEstoque           As Boolean

Public Validar              As Boolean
'====================================
' VARIAVEIS ENVIO DAFE - 2019       =
'====================================
Public EmailEnvioNFe        As String
Public EmailUsuario         As String
Public EmailCliente         As String
Public EmailFornecedor      As String
Public EmailTransportadora  As String
'====================================
' VARIAVEIS BOLETO - 2019           =
'====================================
Public AgenciaBol           As String
Public ContaCorrenteBol     As String
Public Codigocedente        As String
Public SiglaCedente         As String
'====================================
' VARIAVEIS CHAVE DE ACESSO - 2019  =
'====================================
Public chChave              As String
Public chCodUF              As String
Public chDTEmissao          As String
Public chCNPJ               As String
Public chModelo             As String
Public chSerie              As String
Public chNNfe               As String
Public chFormaEmissao       As String
Public chCodNumerico        As String
Public chdVer               As String
'====================================
' Encerrar ordem 2019               =
'====================================
Public DataOrdem            As String
Public NOrdem               As Long
Public OpOrdem              As Boolean
'====================================
' Variavel de serie da Nota fiscal
'====================================
Public NF_Serie             As Integer
'====================================
'   VARIAVEIS DE NFe TEXTO 2019     =
'====================================
Public Texto_Envio          As String 'Texto contendo xml da nota fiscal para envio para o Sefaz
Public Texto_Cancelamento   As String 'Texto para cancelamento da nota fiscal
Public ResultadoNFe         As String 'Texto contendo resultado consulta NfeAPI
Public tpAmb                As String 'Texto do tipo do ambiente na SEFAZ
Public vFimCST              As String 'Texto para CST variavel
Public TextoRetorno         As String 'Texto contendo retorno montado
Public VarObjetonome        As String
Public nsUF                 As String
Public idDest               As String
Public indFinal             As String
Public indIEDest            As String
Public vICMSUFDest          As Double
Public ttvICMSUFDest        As Double
'====================================
Public Certificado          As String 'Numero do certificado - OK
Public Corrida              As String 'Numero da corrida - OK
Public Cliente              As String 'Nome do cliente - OK
Public DesenhoProduto       As String 'CÛdigo interno na pesquisa da estrutura do produto - OK
Public TipoProcesso         As String 'Informa o tipo do processo - OK
Public Letra                As String 'OK
Public Familiatext          As String 'OK
Public Transportadora       As String 'OK
Public Acao                 As String 'Controle de mensagens - OK
Public Tipo                 As String 'OK
Public TipoNF               As String 'OK
Public fotopadrao           As String 'OK
Public fotos                As String 'OK
Public Salvarrel            As String 'OK
Public Codproduto           As String 'OK
Public Instalado            As String 'OK
Public Ordem                As String 'OK
Public Pesquisa             As String 'OK
Public caminho              As String 'OK
Public maquina              As String 'OK
Public material             As String 'OK
Public Diasemana            As String 'OK
Public Formulario           As String 'OK
Public Formulario_nota      As String 'OK
Public Formulario_necessidade  As String 'OK
Public Formulario_produto      As String 'OK
Public Formulario_familia      As String 'OK

Public SQL                  As String 'OK
Public StrSql               As String 'OK
Public StrSqlLocProdPadrao  As String 'OK
Public StrSqlLocCliPadrao   As String 'OK
Public StrSqlLocFornPadrao  As String 'OK

Public Desenho              As String 'OK
Public FormatoData          As String 'OK
Public FormatoHora          As String 'OK
Public Liberacao            As String 'OK

Public pubUsuario           As String 'OK
Public pubIDUsuario         As String 'OK
Public pubNome              As String 'OK
Public pubSetor             As String 'OK
Public pubSenha             As String 'OK
Public pubEmail             As String 'OK
Public ValorFatura          As Double
Public condicao_duplicata   As String 'OK
Public DT                   As String 'OK
Public TemICMS              As String 'OK
Public TemIPI               As String 'OK
Public SomarIPI             As String 'OK
Public DestacaImpostos      As String 'OK
Public DiaX                 As String 'OK
Public MesX                 As String 'OK
Public MesX1                As String 'OK
Public AnoX                 As String 'OK
Public Produto              As String 'OK
Public Filter               As String 'OK
Public mxCondpag            As String 'OK
Public Mensagem             As String 'OK
Public CFOP_vendas          As String 'OK
Public FamiliaAntiga        As String 'OK
Public pc                   As String 'OK
Public Tipo_endereco        As String 'OK
'Public Endereco             As String 'OK
'Public Bairro               As String 'OK
'Public Cidade               As String 'OK
'Public UF                   As String 'OK
Public Par1                 As String 'OK
Public Par2                 As String 'OK
Public Revisao              As String 'OK
Public Ordenar              As String 'OK - Variavel utilizada para ordenar coluna
Public NomeTabelaAp         As String 'OK
Public NomeTabelaApTotalizacao     As String 'OK
Public Nome_anexo           As String 'OK
Public Moeda                As String 'OK
Public NumeroInteiro()      As String 'OK - Variavel para arredondamento
Public MsgTexto             As String 'OK
Public Modulo_caprind       As String 'OK

'C·lculo de tempos e custos do processo
Public TempoPreparacao      As String 'OK
Public TempoExecucao        As String 'OK

'Criar local do bd e do relatÛrio
Public NomeServidor         As String 'OK
Public Nome_banco           As String 'OK
Public Localrel             As String 'OK
Public Usuario_banco        As String 'OK
Public Senha_banco          As String 'OK

Public Var                  As String 'OK
Public VarE                 As String 'OK
Public VarR                 As String 'OK
Public VarU                 As String 'OK
Public VarS                 As String 'OK
Public VarLAC               As String 'OK
Public VarLNC               As String 'OK
Public VarLAG               As String 'OK
Public VarLNG               As String 'OK

Public NomeServidor1        As String 'OK
Public Nome_banco1          As String 'OK
Public Localrel1            As String 'OK
Public Usuario_banco1       As String 'OK
Public Senha_banco1         As String 'OK

Public Var1                 As String 'OK
Public VarE1                As String 'OK
Public VarR1                As String 'OK
Public VarU1                As String 'OK
Public VarS1                As String 'OK
Public VarLAC1              As String 'OK
Public VarLNC1              As String 'OK
Public VarLAG1              As String 'OK
Public VarLNG1              As String 'OK

Public NomeServidor2        As String 'OK
Public Nome_banco2          As String 'OK
Public Localrel2            As String 'OK
Public Usuario_banco2       As String 'OK
Public Senha_banco2         As String 'OK

Public Var2                 As String 'OK
Public VarE2                As String 'OK
Public VarR2                As String 'OK
Public VarU2                As String 'OK
Public VarS2                As String 'OK
Public VarLAC2              As String 'OK
Public VarLNC2              As String 'OK
Public VarLAG2              As String 'OK
Public VarLNG2              As String 'OK

Public NomeServidor3        As String 'OK
Public Nome_banco3          As String 'OK
Public Localrel3            As String 'OK
Public Usuario_banco3       As String 'OK
Public Senha_banco3         As String 'OK

Public Var3                 As String 'OK
Public VarE3                As String 'OK
Public VarR3                As String 'OK
Public VarU3                As String 'OK
Public VarS3                As String 'OK
Public VarLAC3              As String 'OK
Public VarLNC3              As String 'OK
Public VarLAG3              As String 'OK
Public VarLNG3              As String 'OK

'Criar n˙meros de solicitaÁ„o de compras e cotaÁ„o
Public a                    As String 'OK
Public A2                   As String 'OK

'Gravacao de evento
Public Modulo               As String 'OK
Public Evento               As String 'OK
Public Documento            As String 'OK
Public Documento1           As String 'OK
Public ID_documento         As String 'OK
Public Simbolos             As String 'OK
Public pubRegistrado        As String 'OK
Public pubLicenca           As String 'OK

'Filtros carteira de produÁ„o e engenharia
Public FiltroMRP As String 'OK
Public FiltroMRPRel As String 'OK
Public TemOrdem As String 'OK
Public TemOrdemRel As String 'OK
Public Expedido As String 'OK
Public AtrasadoFiltro As String 'OK
Public AtrasadoFiltroRel As String 'OK
Public StatusFiltro As String 'OK
Public StatusFiltroRel As String 'OK
Public SerieNF As String 'OK

Public InicioCST               As String 'OK
Public FimCST                  As String 'OK

'===========================
'=   VARIAVEIS DE N⁄MERO   =
'===========================
Public Saldo            As Long 'Saldo estoque consignacao
Public TipoItem         As Integer 'Tipo do produto Ex. Comercial, Industrial - OK
Public SubTipoItem      As Integer 'Sub Tipo do produto Ex: Montar, Fabricar, Expedir - OK
Public Aplic            As Integer 'OK
Public NumeroErro       As Integer 'OK
Public diaprep          As Integer 'OK
Public diaexec          As Integer 'OK
Public situacao         As Integer 'OK
Public Sit_REG          As Integer 'OK
Public Sit_Data         As Integer 'OK
Public Init             As Integer 'OK
Public Sit_Nota         As Integer 'OK
Public nPagto           As Integer 'OK
Public Controle         As Integer 'OK
Public CompLetra        As Integer 'OK
Public w                As Integer 'OK
Public EventFlag        As Integer 'Para sinalizar qual evento ocorreu - OK
Public IntCounter       As Integer 'Contador para a barra de progresso - OK
Public Id_Item          As Long 'OK
Public Status_nota      As Integer 'OK
Public Regime           As Integer 'OK
Public Qtlicencas_caprind As Integer 'OK
Public Qtlicencas_gerprod As Integer 'OK

'VerificaÁ„o de recebimento
Public Embalagem        As Integer 'OK
Public Laudos           As Integer 'OK
Public Visual           As Integer 'OK
Public Dimensoes        As Integer 'OK
Public Outros           As Integer 'OK

'Criar n˙meros de solicitaÁ„o de compras e cotaÁ„o
Public Cont             As Integer 'OK
Public Cont2            As Integer 'OK

Public IDConta          As Integer  'Nova criada em 22-01-2008 por francisco
Public IDTransf         As Integer  'Nova Criada em 22-01-2008 por francisco

Public CountReg             As Integer
Public CountDias            As Integer

Public Comissao             As Double

Public quantidade           As Double 'OK
Public Desejada             As Double 'OK
Public Encontrada           As Double 'OK
Public TolSup               As Double 'OK
Public TolInf               As Double 'OK
Public Desconto             As Double 'OK
Public TTE                  As Double 'OK
Public QuantEmpenho         As Double 'OK
Public QuantEmpenhoPC       As Double 'OK
Public NEstoqueMinimo       As Double 'Necessidade pare estoque minimo - OK
Public EstoqueMinimo        As Double 'OK
Public NPrevista            As Double 'Necessidade de compra prevista - OK
Public NReal                As Double 'Necessidade de compra Real - OK
Public qtdeliberar          As Double 'OK
Public qtdeliberada         As Double 'OK
Public qtdeliberadaPC       As Double 'OK
Public quantnovo            As Double 'OK
Public quantestoque         As Double 'OK
Public QuantComprado        As Double 'OK
Public QuantComprado1       As Double 'OK
Public QuantSolicitado      As Double 'OK
Public QuantsolicitadoN1    As Double 'OK
Public QuantsolicitadoN2    As Double 'OK
Public QuantsolicitadoN3    As Double 'OK
Public QuantsolicitadoN4    As Double 'OK
Public QuantsolicitadoN5    As Double 'OK
Public QuantsolicitadoN6    As Double 'OK
Public QuantsolicitadoN7    As Double 'OK
Public QuantsolicitadoN8    As Double 'OK
Public QuantsolicitadoN9    As Double 'OK
Public QuantsolicitadoN10   As Double 'OK
Public QuantsolicitadoN11   As Double 'OK
Public QuantsolicitadoN12   As Double 'OK
Public QuantsolicitadoN13   As Double 'OK
Public QuantsolicitadoN14   As Double 'OK
Public Estoquereal          As Double 'OK
Public Estoquevenda         As Double 'OK
Public IntIPI               As Double 'OK
Public IntICMS              As Double 'OK
Public PesoMetro            As Double 'OK
Public Pesoestoque          As Double 'OK
Public Comprimento          As Double 'OK
Public Pesolote             As Double 'OK
Public Maxima               As Double 'OK
Public Minima               As Double 'OK
Public Media                As Double 'OK
Public Totalduplicata       As Double 'OK
Public Valorparcela         As Double 'OK
Public ValorICMS            As Double 'OK
Public ValorIPI             As Double 'OK
Public SaldoDuplicata       As Double 'OK
Public ValorPorc            As Double 'OK
Public SaldoPorc            As Double 'OK
Public TotalGeral           As Double 'OK
Public Reparar              As Double 'OK
Public Substituir           As Double 'OK
Public TotalProduto         As Double 'OK
Public TotalServicos        As Double 'OK
Public TotalDesconto        As Double 'OK
Public TotalDescontoServico As Double 'OK
Public SubTotal             As Double 'OK
Public SubTotalServicos     As Double 'OK
Public Qtd                  As Double 'OK
Public mxValorPag           As Double 'OK
Public ValoresParcelas      As Double 'OK
Public Pendente             As Double 'OK
Public TotalReceber         As Double 'OK
Public TotalPagar           As Double 'OK
Public ValorConta           As Double 'OK
Public Qtde                 As Double 'OK
Public EstoqueAtual         As Double 'OK
Public EstoqueAtualPC       As Double 'OK
Public QtdeSaida            As Double 'OK
Public Entrada              As Double 'OK
Public QtdeSaidaPC          As Double 'OK
Public VltUnit              As Double 'OK
Public qt                   As Double 'OK
Public VlttTotal            As Double 'OK
Public TotalISS             As Double 'OK
Public ValorParcial         As Double 'OK
Public DecimoSegundos       As Double 'OK
Public VlrSubTotal          As Double 'OK
Public vlrTotal             As Double 'OK
Public VlrTotalServ         As Double 'OK
Public VlrTotalRetorno      As Double 'OK
Public VlrTotalRetornoNF    As Double 'OK
Public ValorNC              As Double 'OK
Public BASECALCULO          As Double 'OK
Public ValorMoeda           As Double 'OK
Public ValorHora            As Double 'OK
Public ValorHoraPrep        As Double 'OK

Public NomeView             As String

Public ValorTotalPagar  As Double   'Nova criada em 22-01-2008 por francisco
Public ValorPagar       As Double   'Nova criada em 22-01-2008 por francisco
Public ValorTotalPago   As Double   'Nova criada em 22-01-2008 por francisco
Public ValorPago        As Double   'Nova criada em 22-01-2008 por francisco
Public ValorPagoParcial As Double   'Nova criada em 22-01-2008 por francisco
Public ValorPendente    As Double   'Nova criada em 22-01-2008 por francisco
Public Valor1           As Double   'Nova criada em 22-01-2008 por francisco
Public Valor2           As Double   'Nova criada em 22-01-2008 por francisco
Public Valor3           As Double   'Nova criada em 22-01-2008 por francisco
Public SaqueValorTotal  As Double   'Nova criada em 22-01-2008 por francisco
Public SaqueUtilizado   As Double   'Nova criada em 22-01-2008 por francisco
Public SaqueSaldo       As Double   'Nova criada em 22-01-2008 por francisco

'C·lculo de tempos e custos do processo
Public Porcentagem      As Double 'OK
Public Vlrmateriallu    As Double 'OK
Public Vlrmaodeobralu   As Double 'OK
Public Vlrterceiroslu   As Double 'OK
Public Vlrtotallu       As Double 'OK
Public Vlrcomvend       As Double 'OK
Public Vlrtotalcom      As Double 'OK
Public VlrPIS           As Double 'OK
Public VlrCofins        As Double 'OK
Public VlrCSLL          As Double 'OK
Public VlrIRPJ          As Double 'OK
Public CMO              As Double 'OK
Public CMP              As Double 'OK
Public CPREP            As Double 'OK
Public CTERC            As Double 'OK
Public CTT              As Double 'OK
Public Total            As Double 'OK
Public Total1           As Double 'OK
Public TotalProcCusto   As Double 'OK
Public TotalPreparacao  As Double 'OK
Public CustoMaterial    As Double 'OK
Public ICMS             As Double 'OK
Public ValorTotal       As Double 'OK
Public ICMSOUTROS       As Double 'Calculo de imposto - OK
Public vlrTotalProd     As Double 'OK
Public Precofinal       As Double 'OK
Public Precounitario    As Double 'OK
Public Totalliberado    As Double 'OK
Public TotalProposta    As Double 'OK
Public BC               As Double 'Variavel da base de calculo do ICMS - OK
Public BCST             As Double 'Variavel da base de calculo do ICMS ST - OK
Public PV               As Double 'Variavel do preÁo de venda com icms incluso R$ - OK
Public PV1              As Double 'Variavel do preÁo de venda com icms incluso R$ - OK
Public SumTotProdutos   As Double 'OK
Public SumTotNota       As Double 'OK
Public SumIPI           As Double 'OK
Public SumICMS          As Double 'OK
Public VlrIPI           As Double 'OK
Public TotContas        As Double 'OK
Public valor            As Double 'OK
Public Valores          As Double 'OK
Public VP               As Double 'OK
Public VD               As Double 'OK
Public TotalCredito     As Double 'OK
Public TotalDebito      As Double 'OK
Public TotalCreditar    As Double 'OK
Public TotalDebitar     As Double 'OK
Public Valor_IPI        As Double 'OK
Public Valor_total      As Double 'OK
Public Saldo_Anterior   As Double 'OK
Public Saldo_Atual      As Double 'OK
Public Valor_Produto    As Double 'OK
Public VlISS            As Double 'OK
Public Credito_ICMS     As Double 'OK
Public ICMS_suframa     As Double 'OK
Public VlrICMS_suframa  As Double 'OK
Public IPI              As Double 'OK
Public TTICMS           As Double 'OK
Public Total_ICMS       As Double 'OK
Public Calcula_ICMS     As Double 'OK
Public CTMedioEst       As Double 'OK

'Vari·veis para c·lculo da CST
Public BCICMSCST       As Double 'OK
Public ICMSCST         As Double 'OK
Public TotalBCICMSCST  As Double 'OK
Public TotalICMSCST    As Double 'OK

'Calculo de reduÁ„o de base de calculo icms
Public CT               As Double 'Variavel da carga tributaria % - OK
Public CTDE             As Double 'Variavel da carga tributaria dentro do estado % - OK
Public CTSS             As Double 'Variavel da carga tributaria no sul sudeste % - OK
Public CTNN             As Double 'Variavel da carga tributaria no norte nordeste % - OK
Public CTCO             As Double 'Variavel da carga tributaria no centro oeste % - OK
Public CTEX             As Double 'Variavel da carga tributaria no centro oeste % - OK

'C·lculo de ICMS diferido
Public DIF               As Double 'OK
Public DIFDE             As Double 'OK
Public DIFSS             As Double 'OK
Public DIFNN             As Double 'OK
Public DIFCO             As Double 'OK
Public DIFEX             As Double 'OK

'C·lculo do FCP (fundo de combate ‡ pobreza)
Public FCP               As Double 'OK

'Calculo de reduÁ„o de PIS/Cofins
Public Valor_Retencao_PIS    As Double 'OK
Public Valor_Retencao_Cofins As Double 'OK

Public Valor_Retencao_Servico As Double 'OK

'Calculo de imposto
Public VLFRETE              As Double 'OK
Public VLSEGURO             As Double 'OK
Public VLOUTROS             As Double 'OK
Public VLICMSOUTROS         As Double 'OK
Public PIS_Serv             As Double 'OK
Public Valor_PIS_Serv       As Double 'OK
Public Cofins_Serv          As Double 'OK
Public Valor_Cofins_Serv    As Double 'OK
Public CSLL_Serv            As Double 'OK
Public Valor_CSLL_Serv      As Double 'OK
Public ISS_Serv             As Double 'OK
Public Valor_ISS_Serv       As Double 'OK
Public INSS_Serv            As Double 'OK
Public Valor_INSS_Serv      As Double 'OK
Public IRPJ_Serv            As Double 'OK
Public Valor_IRPJ_Serv      As Double 'OK
Public IRRF_Serv            As Double 'OK
Public Valor_IRRF_Serv      As Double 'OK
Public PIS_Prod             As Double 'OK
Public Valor_PIS_Prod       As Double 'OK
Public Cofins_Prod          As Double 'OK
Public Valor_Cofins_Prod    As Double 'OK
Public CSLL_Prod            As Double 'OK
Public Valor_CSLL_Prod      As Double 'OK
Public IRPJ_Prod            As Double 'OK
Public Valor_IRPJ_Prod      As Double 'OK
Public CPP_Prod             As Double 'OK
Public Valor_CPP_Prod       As Double 'OK
Public CPP_Serv             As Double 'OK
Public Valor_CPP_Serv       As Double 'OK

Public DAS                  As Double 'OK
Public Valor_DAS            As Double 'OK
Public Valor_DAS_Serv       As Double 'OK
Public ICMS_SN              As Double 'OK
Public IPI_SN               As Double 'OK
Public Valor_ICMS_SN        As Double 'OK

Public CRLOTE               As Double 'OK
Public CRPECA               As Double 'OK
Public Qtd_Prog             As Double 'OK

'Carga de maquina
Public TotalSegDisp     As Double 'OK

Public TotalUtilPeriodo As Double 'OK
Public TotalRestPeriodo As Double 'OK
Public TotalRestMaquina As Double 'OK

Public IDlogon              As Long 'OK
Public NumeroME         As Long 'OK
Public IDlista          As Long 'OK
Public IDpedido         As Long 'OK
Public IDPlano          As Long 'OK
Public ContOE           As Long 'contador de ordem de expediÁ„o - OK
Public ContOF           As Long 'contador de ordem de fabricaÁ„o - OK
Public ContOM           As Long 'contador de ordem de montagem - OK
Public ContOrdem        As Long 'contador de todos os tipos de ordem - OK
Public IDFase           As Long 'OK
Public IDPROCESSO       As Long 'OK
Public IDMaquina        As Long 'OK
Public IDUsuario        As Long 'OK
Public ID_Familia       As Long 'OK
Public IDAntigo         As Long 'OK
Public OF               As Long 'OK
Public Quant            As Double 'OK
Public IDCliente        As Long 'OK
Public Contador As Long, Contador1 As Long, Contador2 As Long, Contador3 As Long, ContadorReg As Long
Public NotaFiscal       As Long 'OK
Public i                As Long 'OK
Public produzidas       As Double 'OK
Public OS               As Long 'OK
Public Posicao          As Long 'OK
Public InitFor          As Long 'OK
Public InitFor1         As Long 'OK

'===========================
'=   VARIAVEIS DE TREEVIEW     =
'===========================
Public Pagar As TreeNode, Produtos As TreeNode, NivelP1 As TreeNode, NivelP2 As TreeNode, NivelP3 As TreeNode, NivelP4 As TreeNode, NivelP5 As TreeNode, NivelP6 As TreeNode, NivelP7 As TreeNode, NivelP8 As TreeNode, NivelP9 As TreeNode, NivelP10 As TreeNode, NivelP11 As TreeNode, NivelP12 As TreeNode, NivelP13 As TreeNode, NivelP14 As TreeNode, NivelP15 As TreeNode
Public Receber As TreeNode, NivelR1 As TreeNode, NivelR2 As TreeNode, NivelR3 As TreeNode, NivelR4 As TreeNode, NivelR5 As TreeNode, NivelR6 As TreeNode, NivelR7 As TreeNode, NivelR8 As TreeNode

'===========================
'=   VARIAVEIS DE DATA     =
'===========================
Public Tempoprocesso    As Date 'OK
Public Dataini          As Date 'OK
Public DataFim          As Date 'OK
Public Inicio           As Date 'OK
Public intervalo        As Date 'OK
Public TempoTotalDias   As Date 'OK
Public Inicio_prazo     As Date 'OK
Public Fim_prazo        As Date 'OK

'C·lculo de tempos e custos do processo
Public DataResultado    As Date 'OK
Public HoraResultado    As Date 'OK
Public Preparacao       As Date 'OK
Public Execucao         As Date 'OK

Public TRLOTE      As Double 'OK

'Carga de maquina
Public TotalDisponivel  As Date 'OK
Public TotalDispPeriodo As Date 'OK

'===========================
'=   VARIAVEIS DE DECIS√O  =
'===========================
Public EmailEnviado     As Boolean  'Nova criada para envio de email DanfeXML
Public PagarParcial     As Boolean  'Nova criada em 22-01-2008 por francisco
Public PagarIntegrar    As Boolean  'Nova criada em 22-01-2008 por francisco
Public PagarVarias      As Boolean  'Nova criada em 22-01-2008 por francisco
Public PagarUnica       As Boolean  'Nova criada em 22-01-2008 por francisco
Public Semsolicitacao                   As Boolean 'Criar n˙meros de solicitaÁ„o de compras e cotaÁ„o - OK
Public CarregaListaForm                 As Boolean 'Carrega lista nos formularios - OK
Public CodManual                        As Boolean 'Cadastro de produto com codigo manual - OK
Public Programacao                      As Boolean 'ProgramaÁ„o de compra - OK
Public Permitido                        As Boolean 'Vari·vel de verificaÁ„o - OK
Public Permitido1                       As Boolean 'Vari·vel de verificaÁ„o - OK
Public Permitido2                       As Boolean 'Vari·vel de verificaÁ„o - OK
Public ParaProcesso                     As Boolean 'OK
Public Urgencia                         As Boolean 'OK
Public Acessos                          As Boolean 'OK
Public NotaProposta                     As Boolean 'OK
Public Comercial                        As Boolean 'OK
Public Industrial                       As Boolean 'OK
Public PropostaVendida                  As Boolean 'OK
Public VE                               As Boolean 'OK
Public VI                               As Boolean 'OK
Public Abrir                            As Boolean 'OK
Public Retirar                          As Boolean 'OK
Public Modificado                       As Boolean 'OK
Public Reducao                          As Boolean 'Variavel de controle se tem reduÁ„o na base de calculo - OK
Public Retencao_PIS_Cofins              As Boolean 'Variavel de controle se tem retenÁ„o de PIS/Cofins - OK
Public Novo                             As Boolean 'OK
Public Sair                             As Boolean 'OK
Public Parcial                          As Boolean 'OK
Public Recebido                         As Boolean 'OK
Public Aberto                           As Boolean 'OK
Public Previsao                         As Boolean 'OK
Public Servicos                         As Boolean 'OK
Public Prod                             As Boolean 'OK
Public Encontrou                        As Boolean 'OK
Public Logoff                           As Boolean 'OK
Public Avancar                          As Boolean 'OK
Public Tipo_Processo                    As Boolean 'OK
Public CompactarRepararBanco            As Boolean 'OK
Public Soma_retorno                     As Boolean 'OK
Public Soma_retorno_total_NF            As Boolean 'OK
Public Item                             As Boolean 'OK
Public Valido                           As Boolean 'OK
Public N_Valido                         As Boolean 'OK
Public Atualizacao                      As Boolean 'OK Usar somente para o bot„o de atualizar
Public InfLicenca                       As Boolean 'OK
Public TemPIS                           As Boolean 'OK
Public TemCOFINS                        As Boolean 'OK
Public Desconto_Suframa                 As Boolean 'OK
Public Suframa                          As Boolean 'OK
Public MPA                              As Boolean 'OK
Public TemReducaoBC                     As Boolean 'OK
Public Antecipacao                      As Boolean 'OK
Public Devolucao                        As Boolean 'OK
Public TemInternet                      As Boolean 'OK
Public ErroDriverMYSQL                  As Boolean 'OK
Public OutraMoeda                       As Boolean 'OK
Public SomarIPIST                       As Boolean 'OK
Public Clientes_Grupos                  As Boolean 'OK
Public VerifDadosPadraoFamilia          As Boolean 'OK

'Gravar CST NFe
Public CST_ICMS                As Boolean 'OK
Public CST_IPI                 As Boolean 'OK
Public CST_PIS                 As Boolean 'OK
Public CST_Cofins              As Boolean 'OK

'UtilizaÁ„o do mesmo form em mas de um mÛdulo
Public Analise_critica                      As Boolean 'OK
Public Atualizacao_GNFe                     As Boolean 'OK
Public Atualizacao_GMRE                     As Boolean 'OK
Public Atualizacao_versao                   As Boolean 'OK
Public Atualizacao_TeamViewer               As Boolean 'OK
Public Atualizacao_TeamViewerQS             As Boolean 'OK
Public CadMaquinas                          As Boolean 'OK
Public Clientes                             As Boolean 'OK
Public CC_RM                                As Boolean 'OK
Public ClassFiscal                          As Boolean 'OK
Public Compras_Familia                      As Boolean 'OK
Public Compras                              As Boolean 'Produtos
Public Compras_Fallow_up                    As Boolean 'OK
Public Compras_Requisicao                   As Boolean 'OK
Public Compras_Programacao                  As Boolean 'OK
Public Compras_Cotacao                      As Boolean 'OK
Public Compras_Pedido                       As Boolean 'OK
Public Compras_Produtos                     As Boolean 'OK
Public Compras_Fornecedores                 As Boolean 'OK
Public Compras_Necessidade                  As Boolean 'OK
Public Compras_Relatorio_IndiceAtraso       As Boolean 'OK
Public Custos_justificativa                 As Boolean 'OK
Public Engenharia_Localfornecedor           As Boolean 'OK
Public Engenharia_Localfornecedor1          As Boolean 'OK
Public Engenharia_Localcliente              As Boolean 'OK
Public Engenharia_Localcliente1             As Boolean 'OK
Public Engenharia                           As Boolean 'OK
Public Engenharia_Familia                   As Boolean 'OK
Public Engenharia_Conjuntos                 As Boolean 'OK
Public Engenharia_Normas                    As Boolean 'OK
Public Engenharia_Produtos                  As Boolean 'OK
Public Estoque_recebimento                  As Boolean 'OK
Public Estoque_entrada                      As Boolean 'OK
Public Estoque_Requisicao                   As Boolean 'OK
Public Estoque_Consignacao                  As Boolean 'OK
Public Estoque_Inventario                   As Boolean 'OK
Public Estoque_Local_Armazenamento          As Boolean 'OK
Public Faturamento                          As Boolean 'OK
Public Faturamento_NF_Saida                 As Boolean 'OK
'==========================================================
Public Faturamento_NF_Terceiro              As Boolean
Public Faturamento_NF_Propria               As Boolean
'==========================================================
Public Faturamento_ListaProdudos            As Boolean 'OK
Public Financeiro_Contas_Pagar              As Boolean 'OK
Public Financeiro_Forma_Pgto_Pagar          As Boolean 'OK
Public Financeiro_Forma_Pgto_Receber        As Boolean 'OK
Public Financeiro_Contas_Pagas              As Boolean 'OK
Public Financeiro_Contas_Receber            As Boolean 'OK
Public Financeiro_Contas_Recebidas          As Boolean 'OK
Public Fiscal_NaturezaOperacao              As Boolean 'OK
Public Funcionario                          As Boolean 'OK
Public Imprimir                             As Boolean 'OK
Public Imprimir1                            As Boolean 'OK
Public Inspecao_recebimento                 As Boolean 'OK
Public Inspecaorecebimento_AnexarPlano      As Boolean 'OK
Public Financeiro_Instituicao               As Boolean 'OK
Public Liquido                              As Boolean 'OK
Public Manutencao                           As Boolean 'OK
Public Manutencao_Relatorio_Historico       As Boolean 'OK
Public Minuta                               As Boolean 'OK
Public Outros_solicitacaoPCP                As Boolean 'OK
Public OpcoesGerais                         As Boolean 'OK
Public PCP_Ordem                            As Boolean 'OK
Public PCP_AlterarRM                        As Boolean 'OK
Public PCP_Imprimir                         As Boolean 'OK
Public PCP_relatorios_indice_atraso         As Boolean 'OK
Public PCP_Necessidade                      As Boolean 'OK
Public Plano_contas_produtos                As Boolean 'OK
Public Plano_contas_familias                As Boolean 'OK
Public Plano_centro_de_custo                As Boolean 'OK
Public Plano_instituicao                    As Boolean 'OK
Public Plano_opcoesgerais                   As Boolean 'OK
Public Plano_Faturamento                    As Boolean 'OK
Public Plano_PCP                            As Boolean 'OK
Public PI_Produtos                          As Boolean 'OK
Public PI_Servicos                          As Boolean 'OK
Public Proposta_Servicos                    As Boolean 'OK
Public Processos_instrucoes                 As Boolean 'OK
Public Qualidade_Almox                      As Boolean 'OK
Public Qualidade_Familia                    As Boolean 'OK
Public Qualidade_PPAP_Plano                 As Boolean 'OK
Public Qualidade_PPAP_PSW                   As Boolean 'OK
Public Qualidade_PPAP_FMEA                  As Boolean 'OK
Public Qualidade_sistema                    As Boolean 'OK
Public Qualidade_NC                         As Boolean 'OK
Public Qualidade_Plano                      As Boolean 'OK
Public RH_Funcionarios                      As Boolean 'OK
Public RNC_Inspecao_Recebimento             As Boolean 'OK
Public RNC_Controle_Medicao                 As Boolean 'OK
Public RNC_Nao_Conformidade                 As Boolean 'OK
Public RNC_Solicitacao_Desvio               As Boolean 'OK
Public RNC                                  As Boolean 'OK
Public SolicitacaoAcao                      As Boolean 'OK
Public Substituicao                         As Boolean
Public Telemarketing                        As Boolean 'OK
Public Troca_Duplicata                      As Boolean 'OK
Public Usuarios                             As Boolean 'OK
Public Ultrasom                             As Boolean 'OK
Public Vendas_Familia                       As Boolean 'OK
Public Vendas                               As Boolean 'OK
Public Vendas_Produtos                      As Boolean 'Produtos - OK
Public Vendas_Programacao                   As Boolean 'OK
Public Vendas_Analise                       As Boolean 'OK
Public Vendas_Carteira                      As Boolean 'OK
Public Vendas_Proposta                      As Boolean 'OK
Public Vendas_PI                            As Boolean 'OK
Public Vendas_Vendedores                    As Boolean 'OK
Public Vendas_AtualizaÁ„o_Valores           As Boolean 'OK
Public Vendas_Relatorio_Historico           As Boolean 'OK
Public Vendas_Relatorio_IndiceAtraso        As Boolean 'OK
Public Vendas_Relatorio_Comissao            As Boolean 'OK
Public Downloads_NF                         As Boolean 'OK
Public Chat                                 As Boolean 'OK
Public Video_ajuda                          As Boolean 'OK
Public Familia_NCM                          As Boolean 'OK

'===========================
'=   VARIAVEIS INDEFINIDA  =
'===========================
Public LocalAntigoCaprind   As Variant 'OK
Public LocalNovoCaprind     As Variant 'OK
Public LocalAntigoGerprod   As Variant 'OK
Public LocalNovoGerprod     As Variant 'OK
Public LocalAntigoCaprind1  As Variant 'OK
Public LocalNovoCaprind1    As Variant 'OK
Public LocalAntigoGerprod1  As Variant 'OK
Public LocalNovoGerprod1    As Variant 'OK
Public LocalAntigoCaprind2  As Variant 'OK
Public LocalNovoCaprind2    As Variant 'OK
Public LocalAntigoGerprod2  As Variant 'OK
Public LocalNovoGerprod2    As Variant 'OK
Public LocalAntigoCaprind3  As Variant 'OK
Public LocalNovoCaprind3    As Variant 'OK
Public LocalAntigoGerprod3  As Variant 'OK
Public LocalNovoGerprod3    As Variant 'OK

Public Incluir              As Variant 'OK
Public Excluir              As Variant 'OK
Public Alterar              As Variant 'OK
Public TotalUtilizado       As Variant 'OK
Public Dia                  As Variant 'OK
Public TotalSegUtil         As Variant 'OK
Public Cod_produto          As Variant 'OK
Public Peca                 As Variant 'OK
Public NF                   As Variant 'Numero da nota fiscal de material consignado - OK
Public Formula              As Variant 'OK
Public x                    As Variant 'Criar n˙meros de solicitaÁ„o de compras e cotaÁ„o - OK
Public VerifNumero          As Variant 'Verifica campo n˙mero - OK
Public VerifData            As Variant 'Verifica campo data - OK
Public NomeCampo            As Variant 'Controle de mensagens - OK
Public Hora                 As Variant 'OK
Public PAGTO                As Variant 'OK
Public Periodo              As Variant 'OK
Public vlrICMS(0, 6)        As Variant 'OK
Public vRegiao(0, 1)        As Variant 'OK
Public NCliente             As Variant 'OK
Public Parcela              As Variant 'OK
Public TotalOS              As Variant 'OK
Public TotalOrdem           As Variant 'OK
Public mxPagto(1, 16)       As Variant 'OK
Public Data_Prog            As Variant 'OK
Public Prog                 As Variant 'OK
Public Usuario              As Variant 'OK
Public Salvar               As Variant 'OK
Public PrecoHora            As Variant 'OK
Public intIndex             As Variant 'Variavel para o index do NÛ atual - OK
Public TempoTotalProcesso   As Variant 'OK
Public CustoProcesso        As Variant 'OK
Public TotalFaseSeg         As Variant 'OK
Public CustoFase            As Variant 'OK
Public CustohoraSeg         As Variant 'OK
Public CustoTotalPrep       As Variant 'OK
Public TotalFase            As Variant 'OK
Public TotalProcesso        As Variant 'OK
Public TotalProcessoSeg     As Variant 'OK
Public TotalFaseCusto       As Variant 'OK
Public TotalProcessoCusto   As Variant 'OK
Public TotalProcessoHora    As Variant 'OK
Public TotalProcessoMin     As Variant 'OK
Public Disponivel           As Variant 'OK
Public Revisar              As Variant 'OK
Public Proposta             As Variant 'OK

'Apos atualizar lista, manter o ˙ltimo registro selecionado
Public CodigoLista          As Variant 'OK
Public CodigoLista1         As Variant 'OK
Public CodigoLista2         As Variant 'OK
Public CodigoLista3         As Variant 'OK
Public CodigoLista4         As Variant 'OK
Public CodigoLista5         As Variant 'OK
Public CodigoLista6         As Variant 'OK
Public CodigoLista7         As Variant 'OK
Public CodigoLista8         As Variant 'OK
Public CodigoLista9         As Variant 'OK
Public CodigoLista10         As Variant 'OK
Public CodigoLista11         As Variant 'OK

'========================================
'=   VARIAVEIS ABERTURA BANCO DE DADOS  =
'========================================
Public TBnatOp                 As ADODB.Recordset
Public BD                    As Database
Public TBAcessos             As ADODB.Recordset
Public TBNivel1              As ADODB.Recordset
Public TBNivel2              As ADODB.Recordset
Public TBNivel3              As ADODB.Recordset
Public TBNivel4              As ADODB.Recordset
Public TBNivel5              As ADODB.Recordset
Public TBNivel6              As ADODB.Recordset
Public TBNivel7              As ADODB.Recordset
Public TBNivel8              As ADODB.Recordset
Public TBNivel9              As ADODB.Recordset
Public TBNivel10             As ADODB.Recordset
Public TBNivel11             As ADODB.Recordset
Public TBNivel12             As ADODB.Recordset
Public TBNivel13             As ADODB.Recordset
Public TBNivel14             As ADODB.Recordset
Public TBNivel15             As ADODB.Recordset
Public TBSolicitacao         As ADODB.Recordset
Public TBTotaisnota          As ADODB.Recordset
Public TBCorretiva           As ADODB.Recordset
Public TBControleNF          As ADODB.Recordset
Public TBPedido              As ADODB.Recordset
Public TBFIltro              As ADODB.Recordset
Public TBFluxo               As ADODB.Recordset
Public TBOrdem               As ADODB.Recordset
Public TBCompras_Pedido      As ADODB.Recordset
Public TBCompras_Lista       As ADODB.Recordset
Public TBCompras             As ADODB.Recordset
Public TBReceber             As ADODB.Recordset
Public TBRecebidos           As ADODB.Recordset
Public TBTransporte          As ADODB.Recordset
Public TBGravar              As ADODB.Recordset
Public TBExecucao            As ADODB.Recordset
Public TBLISTA               As ADODB.Recordset
Public TBAbrir               As ADODB.Recordset
Public TBSaldo               As ADODB.Recordset
Public TBVendas              As ADODB.Recordset
Public TBCotacao             As ADODB.Recordset
Public TBProposta            As ADODB.Recordset
Public TBCarteira            As ADODB.Recordset
Public TBPI_Lista_produto    As ADODB.Recordset
Public TBLogon               As ADODB.Recordset
Public TBOSC                 As ADODB.Recordset
Public TBAliquota            As ADODB.Recordset
Public TBEstoque             As ADODB.Recordset
Public TBFornecedor          As ADODB.Recordset
Public TBMaterial            As ADODB.Recordset
Public TBTipo                As ADODB.Recordset
Public TBMateriaprima        As ADODB.Recordset
Public TBAfericao            As ADODB.Recordset
Public TBProduto             As ADODB.Recordset
Public TBItem                As ADODB.Recordset
Public TBComponente          As ADODB.Recordset
Public TBInstrumentos        As ADODB.Recordset
Public TBplano               As ADODB.Recordset
Public TBplanolaudo          As ADODB.Recordset
Public TBplanomedicao        As ADODB.Recordset
Public TBCFOP                As ADODB.Recordset
Public TBClientes            As ADODB.Recordset
Public TBProcessos           As ADODB.Recordset
Public TBUsuarios            As ADODB.Recordset
Public TBFases               As ADODB.Recordset
Public TBFerramentas         As ADODB.Recordset
Public TBHistProc            As ADODB.Recordset
Public TBProgramas           As ADODB.Recordset
Public TBCiclo               As ADODB.Recordset
Public TBMaquinas            As ADODB.Recordset
Public TBproducao            As ADODB.Recordset
Public TBCodigoDesc          As ADODB.Recordset
Public TBProdutividade       As ADODB.Recordset
Public TBProducaoFases       As ADODB.Recordset
Public TBFI                  As ADODB.Recordset
Public TBFamilia             As ADODB.Recordset
Public TBContas              As ADODB.Recordset
Public TBEventos             As ADODB.Recordset
Public TBCST                 As ADODB.Recordset
Public TBMascara             As ADODB.Recordset
Public TBTempo               As ADODB.Recordset
Public TBGravar_NFe          As ADODB.Recordset
Public TBGravar_NFe1         As ADODB.Recordset
Public TBGravar_NFe_Status   As ADODB.Recordset
Public TBAbrir_NFe           As ADODB.Recordset
Public TBOS                  As ADODB.Recordset
Public TBCarregarCombo       As ADODB.Recordset
Public TBOrdemServico        As ADODB.Recordset
Public TBCQ                  As ADODB.Recordset
Public TBFerramenta          As ADODB.Recordset
Public TBPrograma            As ADODB.Recordset
Public TBMySQL               As ADODB.Recordset
Public TBSubreport           As ADODB.Recordset

Public TBLocalizar_produto_padrao As ADODB.Recordset
Public TBLocalizar_produto_padrao1 As ADODB.Recordset
Public TBLocalizar_produto_padrao2 As ADODB.Recordset
Public TBLocalizar_cliente_padrao As ADODB.Recordset
Public TBLocalizar_fornecedor_padrao As ADODB.Recordset

'Conex„o Link
Public Conexao              As ADODB.Connection
Public Conexao_NFe          As ADODB.Connection
Public ConexaoMySql         As ADODB.Connection
Public Conexao_Configuracao As ADODB.Connection

'=============
'=   OUTROS  =
'=============

Public Lista_Cor As ListItem 'Vari·vel para alterar cor da linha da lista

Public NewNode  As Node ' para um novo NÛ - OK
Public MNode    As Node 'OK
Public licensa  As LicensaTFB 'OK
Public Serial   As serialTFB 'OK

Type LicensaTFB
    Senha As String * 10
    Contador As Long
    Numero As Integer
    licensa As String * 20
    Serie As Integer
    Data1 As Date
End Type

Type serialTFB
    Contador As Long
    Serial As String * 10
    Texto As String * 200
    Data1 As Date
End Type

'RelatÛrios crystal 11
Public crAPP As New CRAXDDRT.Application
Public Report As CRAXDDRT.Report
Public crxExport As CRAXDDRT.ExportOptions
Public CPProperty As CRAXDDRT.ConnectionProperty
Public DBTable As CRAXDDRT.DatabaseTable
Public SubReport As CRAXDDRT.SubreportObject

Public NomeRel As String
Public PermitidoRel As Boolean
Public LocalRelPersonalizado As String 'OK

'Ordenar dados do relatÛrio via cÛdigo
Public TabelaRel As Long
Public CampoRel As Long
Public OrdenarRel As Long
Public TabelaRel1 As Long
Public CampoRel1 As Long
Public OrdenarRel1 As Long
Public TabelaRel2 As Long
Public CampoRel2 As Long
Public OrdenarRel2 As Long

'Atualiza subreport
Public NomeSubReport As String
Public NomeSubReport1 As String
Public NomeSubReport2 As String
Public NomeSubReport3 As String
Public NomeSubReport4 As String
Public NomeSubReport5 As String
Public NomeSubReport6 As String
Public NomeSubReport7 As String
Public NomeSubReport8 As String
Public NomeSubReport9 As String

'Localizar pasta
Public Const BIF_RETURNONLYFSDIRS = 1 'OK
Public Const BIF_DONTGOBELOWDOMAIN = 2 'OK
Public Const MAX_PATH = 260 'OK

Public Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long 'OK
Public Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long 'OK
Public Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long 'OK

Public Type BrowseInfo 'OK
    hwndOwner As Long 'OK
    pIDLRoot As Long 'OK
    pszDisplayName As Long 'OK
    lpszTitle As Long 'OK
    ulFlags As Long 'OK
    lpfnCallback As Long 'OK
    lParam As Long 'OK
    iImage As Long 'OK
End Type

Public lpIDList As Long 'OK
Public sBuffer As String 'OK
Public szTitle As String 'OK
Public tBrowseInfo As BrowseInfo 'OK

'Criando backup do banco de dados
Public Declare Function SHFileOperation Lib _
"shell32.dll" Alias "SHFileOperationA" _
(lpFileOp As Any) As Long

Public Declare Sub SHFreeNameMappings Lib _
"shell32.dll" (ByVal hNameMappings As Long)

Public Type SHFILEOPSTRUCT
    hWnd As Long
    wFunc As FO_Functions
    pFrom As String
    pTo As String
    fFlags As FOF_Flags
    fAnyOperationsAborted As Long
    hNameMappings As Long
    lpszProgressTitle As String 'only used if FOF_SIMPLEPROGRESS
End Type

Public Enum FO_Functions
    FO_MOVE = &H1
    FO_COPY = &H2
    FO_DELETE = &H3
    FO_RENAME = &H4
End Enum

Public Enum FOF_Flags
    FOF_MULTIDESTFILES = &H1
    FOF_CONFIRMMOUSE = &H2
    FOF_SILENT = &H4
    FOF_RENAMEONCOLLISION = &H8
    FOF_NOCONFIRMATION = &H10
    FOF_WANTMAPPINGHANDLE = &H20
    FOF_ALLOWUNDO = &H40
    FOF_FILESONLY = &H80
    FOF_SIMPLEPROGRESS = &H100
    FOF_NOCONFIRMMKDIR = &H200
    FOF_NOERRORUI = &H400
    FOF_NOCOPYSECURITYATTRIBS = &H800
    FOF_NORECURSION = &H1000
    FOF_NO_CONNECTED_ELEMENTS = &H2000
    FOF_WANTNUKEWARNING = &H4000
End Enum

Public Type SHNAMEMAPPING
    pszOldPath As String
    pszNewPath As String
    cchOldPath As Long
    cchNewPath As Long
End Type

'Enviar e-mail abrindo o programa do outlook
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const SW_SHOWNORMAL = 1

'Enviar e-mail
Public Email_email As String, Nome_email As String, Usuario_email As String, Senha_email As String, Servidor_SMTP As String
Public Porta_email As Integer

'K-MAIL
Public chartab(17) As String
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
'Open filename
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData1 As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Const OFN_ALLOWMULTISELECT = 512
Const OFN_EXPLORER = 524288
Const OFN_FILEMUSTEXIST = 4096
Const OFN_HIDEREADONLY = 4

'//Charset
Public Enum Charsets
   DEFAULT_ISO_8859_1 = 0
   EUROPE_WINDOWS_1252 = 1
   CYRILLIC_KOI8_R = 2
   ARRABIC_WINDOWS_1256 = 3
   BALTIC_WINDOWS_1257 = 4
   GREEK_WINDOWS_1253 = 5
   HEBREW_WINDOWS_1255 = 6
   UNICODE_UTF7 = 7
   UNICODE_UTF8 = 8
   CYRILLIC_KOI8_U = 9
   CYRILLIC_CP1251 = 10
   CYRILLIC_ISO_8859_5 = 11
   BALTIC_ISO_8859_4 = 12
   CHINESE_TRADITIONAL_BIG5 = 13
   CHINESE_SIMPLIFIED_GB2312 = 14
   CHINESE_SIMPLIFIED_HZ = 15
   THAI_874 = 16
   JAPANESE_ISO_2022 = 17
End Enum

'Ordenar lista
Public OrdAsc As Boolean 'serve para poder classificar como ascendente/descendente
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

'Criar aquivos (.txt), gerenciar pastas e arquivos
Public GerArqPastas As New FileSystemObject 'OK
Public ArqTXT As TextStream 'OK

'ResoluÁ„o da tela/Monitor
Public xTwips%, yTwips%, xPixels#, YPixels# 'OK

'Atualizando o sistema
Public LocalAntigoSincCaprind As Variant
Public LocalNovoSincCaprind As Variant
Public LocalAntigoSincGerprod As Variant
Public LocalNovoSincGerprod As Variant
Public Fso, f, fG, Fsu, FU, FUG
Public Caprind As String
Public Gerprod As String
Public Atualizando As Boolean

'Acesso a internet
'Public ie As Object 'Chat online OK
Public IE As InternetExplorer

Private Declare Function InternetGetConnectedState Lib "wininet" (ByRef dwFlags As Long, ByVal dwReserved As Long) As Long
Private Const CONNECT_LAN As Long = &H2
Private Const CONNECT_MODEM As Long = &H1
Private Const CONNECT_PROXY As Long = &H4
Private Const CONNECT_RAS As Long = &H10
Private Const CONNECT_OFFLINE As Long = &H20
Private Const CONNECT_CONFIGURED As Long = &H40

'Baixar arquivo
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Dim lngRetVal As Long

'Verifica se o sistema esta aberto
'Option Explicit
Private Const TH32CS_SNAPPROCESS As Long = 2
Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * MAX_PATH
End Type
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal lFlags As Long, _
                                                                  ByVal lProcessID As Long) As Long
Private Declare Function Process32First Lib "kernel32" (ByVal hSnapShot As Long, _
                                                        typProcess As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapShot As Long, _
                                                       typProcess As PROCESSENTRY32) As Long
Private Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)

'Encerra o processo
'DeclaraÁ„o de APIs
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long

'DeclaraÁ„o de Vari·veis
Public ProcessId As Long
Public ProcessHandle As Long
Public ThreadId As Long

'DeclaraÁ„o de Constantes
Public Const PROCESS_ALL_ACCESS = &H1F0FFF

'Abrir arquivos
Const SW_HIDE = 0
Const SW_NORMAL = 1
Const SW_SHOWMINIMIZED = 2
Const SW_SHOWMAXIMIZED = 3
Const SW_MAXIMIZE = 3
Const SW_SHOWNOACTIVATE = 4
Const SW_SHOW = 5
Const SW_MINIMIZE = 6
Const SW_SHOWMINNOACTIVE = 7
Const SW_SHOWNA = 8
Const SW_RESTORE = 9
Const SW_SHOWDEFAULT = 10
Const SW_MAX = 10



'Carrega Inst‚ncia do SQL automaticamente
Const SEPARATOR        As String = ""
Public vSplit          As Variant
Public sSrv            As String
Public sDb             As String
Public sUser           As String
Public sPass           As String
Public vSrv            As Variant
Public vDb             As Variant
Public sText           As String
Public m_bEnumSrv      As Boolean
Public m_bOk           As Boolean

'Carregar fuso hor·rio
Private Type SYSTEMTIME
   wYear         As Integer
   wMonth        As Integer
   wDayOfWeek    As Integer
   wDay          As Integer
   wHour         As Integer
   wMinute       As Integer
   wSecond       As Integer
   wMilliseconds As Integer
End Type

Private Type TIME_ZONE_INFORMATION
   Bias As Long
   StandardName(63) As Byte  'unicode (0-based)
   StandardDate As SYSTEMTIME
   StandardBias As Long
   DaylightName(63) As Byte  'unicode (0-based)
   DaylightDate As SYSTEMTIME
   DaylightBias As Long
End Type

Private Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long

'Enviar e receber arquivos do servidor online
Public Declare Function FtpGetFile Lib "Wininet.dll" Alias "FtpGetFileA" (ByVal hFtp As Long, ByVal lpszRemoteFile As String, ByVal lpszNewFile As String, ByVal fFailIfExists As Boolean, ByVal dwFlagsAndAttributes As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Public Declare Function FtpPutFile Lib "wininet" Alias "FtpPutFileA" (ByVal hFtp As Long, ByVal lpszLocalFile As String, ByVal lpszNewRemoteFile As String, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Public Declare Function InternetOpen Lib "Wininet.dll" Alias "InternetOpenA" (ByVal lpszAgent As String, ByVal dwAccessType As Long, ByVal lpszProxyName As String, ByVal lpszProxyBypass As String, ByVal dwFlags As Long) As Long
Public Declare Function InternetConnect Lib "Wininet.dll" Alias "InternetConnectA" (ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Long, ByVal sUsername As String, ByVal sPassword As String, ByVal lService As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Public hopen As Long
Public hconnect As Long
Public ftp As Long

'Redimensionar a coluna do ListView
Private Const LVM_FIRST = &H1000
Private Const LVM_SETCOLUMNWIDTH = (LVM_FIRST + 30)
Private Const LVSCW_AUTOSIZE = -1
Private Const LVSCW_AUTOSIZE_USEHEADER = -2

'Verifica hora do servidor
Private Declare Function NetRemoteTOD Lib "NETAPI32.DLL" (ByVal server As String, buffer As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function NetApiBufferFree Lib "NETAPI32.DLL" (buffer As Any) As Long

Private Type TIME_OF_DAY
  t_elapsedt As Long
  t_msecs As Long
  t_hours As Long
  t_mins As Long
  t_secs As Long
  t_hunds As Long
  t_timezone As Long
  t_tinterval As Long
  t_day As Long
  t_month As Long
  t_year As Long
  t_weekday As Long
End Type

'Estrutura
Private Type NodeData
    Level As Integer
    Text As String
End Type
Public arrNodes(666) As NodeData

Public ArrayQtdeDescNC() As String 'OK 'NC por descriÁ„o no apontado
'=====================================================================
'Componente tecnospeed
'=====================================================================
'Public spdNFSe As NFSex.spdNFSeX
'Public spdProxyNFSe As NFSex.spdProxyNFSeX
'Public spdNFe As spdNFeX
'Public spdNFeDataSet As spdNFeDataSetX
'Public spdNFSeConverter As NFSeConverterX.spdNFSeConverterX


Public Sub ProcGerarPDF2()
'On Error GoTo tratar_erro

'Dim crxApp As New CRAXDRT.Application
'Dim Report As CRAXDRT.Report
'Dim crxExport As CRAXDRT.ExportOptions
'
'Set Report = crxApp.OpenReport(Diretorio & "\399Hsbce.rpt", 1)
'
'Set crxExport = Report.ExportOptions
'
'crxExport.FormatType = crEFTPortableDocFormat
'
'crxExport.DestinationType = crEDTDiskFile
'
'Report.SQLQueryString = "{Boleto_Temp.codigo_usuario} = " & Val(lp_cod_usuario) & " and {Boleto_Temp.Emitir} = true "
'
'crxExport.DiskFileName = Diretorio & "\contador.pdf"
'
'Report.Export False
'Set crxExport = Nothing
'Set Report = Nothing
'
'Exit Sub
'tratar_erro:
'    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
'    Exit Sub
End Sub

Public Sub ListaCertificados()
On Error GoTo tratar_erro

Dim i As Long        ',: Cardinal;
Dim oNode As IXMLDOMNode
Dim SetT As New settings, Certs As ICertificates, StoreSrc As New Store
    'Dim StoreDst As New Store
Dim Cert As Certificate        ': OleVariant;
    'Dim oRps As IXMLDOMNodeList, oLote As IXMLDOMNodeList, oSigs As IXMLDOMNodeList
Dim s1 As String, s2 As String

    'Sett = CoSettings.Create
    'On Error GoTo ListaCert_Error

    SetT.EnablePromptForCertificateUI = True
    'StoreSrc = CoStore.Create
    Call StoreSrc.Open(CAPICOM_CURRENT_USER_STORE, "My", CAPICOM_STORE_OPEN_EXISTING_ONLY)
    'StoreDst = CoStore.Create
    'Call StoreDst.Open(CAPICOM_CURRENT_USER_STORE, "TMP", CAPICOM_STORE_OPEN_MAXIMUM_ALLOWED)
    Set Certs = StoreSrc.Certificates

    '//Remove certificados sem a private key.
    If Certs.Count > 0 Then
             Set Certs = Certs.Find(CAPICOM_CERTIFICATE_FIND_EXTENDED_PROPERTY, CAPICOM_PROPID_KEY_PROV_INFO)
    End If
    '//Somente certificados com data v·lida.
    If Certs.Count > 0 Then
            Set Certs = Certs.Find(CAPICOM_CERTIFICATE_FIND_TIME_VALID)
    End If
    'usMsgbox Certs.Item(1).SubjectName

    'Certs.Select
    'lbCerti = "Nenhum certificado selecionado"
    For i = Certs.Count To 1 Step -1
        USMsgBox Certs.Item(i).SubjectName, vbInformation, "CAPRIND v5.0"
        frmFaturamento_Prod_Serv_NFe_NS.txtSerialCertificado = Certs.Item(i).SerialNumber
'        If mCertificadoSel = Certs.Item(i).SerialNumber Then
'            lbCerti = Certs.Item(i).SubjectName
'        End If
    Next


    If Certs.Count = 0 Then
        lbCerti.Caption = "Sem certificados"
    Else
        Set Cert = Certs.Item(1)
    End If

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcGravarNumeroBoleto(IDConta As Long, IDnota As Long)
On Error GoTo tratar_erro

If Chk_novo.Value = 1 Or Chk_atualizar.Value = 1 Then
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from tbl_Detalhes_Recebimento_Nboletos where IDContaReceber = " & IDConta & " and Nosso_numero = '" & Txt_nosso_numero & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = True Then TBAbrir.AddNew
    TBAbrir!Data = Date
    TBAbrir!Responsavel = pubUsuario
    TBAbrir!IdContaReceber = IDConta
    TBAbrir!Nosso_Numero = Txt_nosso_numero
    TBAbrir!ID_nota = IDnota
    TBAbrir.Update
    TBAbrir.Close
End If

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcCarregaComboEmpresaBoleto()
On Error GoTo tratar_erro

With frm_Instituicoes.Cmb_empresa
    .Clear
    Set TBCarregarCombo = CreateObject("adodb.recordset")
    TBCarregarCombo.Open "Select * from Empresa order by Razao", Conexao, adOpenKeyset, adLockOptimistic
    If TBCarregarCombo.EOF = False Then

        Do While TBCarregarCombo.EOF = False
            If IsNull(TBCarregarCombo!Razao) = False And TBCarregarCombo!Razao <> "" Then
                .AddItem TBCarregarCombo!Razao
                .ItemData(.NewIndex) = TBCarregarCombo!CODIGO
            End If
            TBCarregarCombo.MoveNext
        Loop
        TBCarregarCombo.MoveFirst
    End If
End With
If TBCarregarCombo.RecordCount = 1 Then frm_Instituicoes.Cmb_empresa.Text = TBCarregarCombo!Razao
TBCarregarCombo.Close

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcCarregaComboEmpresaConciliacao()
On Error GoTo tratar_erro

With Frm_InstituicoesConciliacao.cmbempresa
    .Clear
    Set TBCarregarCombo = CreateObject("adodb.recordset")
    TBCarregarCombo.Open "Select * from Empresa order by Razao", Conexao, adOpenKeyset, adLockOptimistic
    If TBCarregarCombo.EOF = False Then
        
        Do While TBCarregarCombo.EOF = False
            If IsNull(TBCarregarCombo!Razao) = False And TBCarregarCombo!Razao <> "" Then
                .AddItem TBCarregarCombo!Razao
                .ItemData(.NewIndex) = TBCarregarCombo!CODIGO
            End If
            TBCarregarCombo.MoveNext
        Loop
        TBCarregarCombo.MoveFirst
    End If
End With
If TBCarregarCombo.RecordCount = 1 Then
    Frm_InstituicoesConciliacao.cmbempresa.Text = TBCarregarCombo!Razao
    Frm_InstituicoesConciliacao.txtCodcedente = TBCarregarCombo!CODIGO
End If

TBCarregarCombo.Close

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcCarregaComboCliente()
On Error GoTo tratar_erro
Dim NomeRazao As String

StrSql = "select Nome_Razao , FormaBaixa from tbl_contas_receber where LogSit = 'N' and formabaixa ='BOLETO' group by Nome_Razao, FormaBaixa ORDER BY Nome_Razao"

With frm_Instituicoes
    .cmbcliente.Clear
    .cmbcliente.AddItem ""
    Set TBCarregarCombo = CreateObject("adodb.recordset")
    TBCarregarCombo.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
    If TBCarregarCombo.EOF = False Then
        
        Do While TBCarregarCombo.EOF = False
            If NomeRazao <> TBCarregarCombo!Nome_Razao Then
                .cmbcliente.AddItem TBCarregarCombo!Nome_Razao
            End If
            TBCarregarCombo.MoveNext
        Loop
        TBCarregarCombo.MoveFirst
    End If
    TBCarregarCombo.Close

End With

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Function AcertaHora(Datahora As String)
On Error GoTo tratar_erro
    
    vetor = Split(Datahora, ",")
    AcertaHora = vetor(0)
                
Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Public Sub ProcMontaStatusBar()
On Error GoTo tratar_erro

ProcVerifQtdeLicencasModulo
Familiatext = "LicenÁas liberadas: " & Qtlicencas_caprind & " Caprind - " & Qtlicencas_gerprod & " Gerprod - MÛdulo: " & Modulo_caprind & ")"

TemInternet = DS.IsInternetOnline

With frmMDI.StatusBar1
    .Clear
    .AddPanel "Usu·rio : " & pubUsuario & " ", [On Center], True, False, , , , , vbWhite, frmMDI.TreeView1.Font, , 300
    .AddPanel " Empresa : " & Nome_banco & " ", [On Center], True, False, , , , , vbWhite, frmMDI.TreeView1.Font, , 500
    .AddPanel " " & Familiatext & " ", [On Center], True, False, , , , , vbWhite, frmMDI.TreeView1.Font, , 1500
    .AddPanel " Internet: " & IIf(TemInternet = True, "On Line", "Off line") & " ", [On Center], True, False, , , , , IIf(TemInternet = False, vbRed, vbWhite), frmMDI.TreeView1.Font, , 500
End With

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Public Sub ProcCarregaComboFornecedor()
On Error GoTo tratar_erro
Dim NomeRazao As String

StrSql = "SELECT tbl_Detalhes_Recebimento.IDContaReceber, tbl_contas_receber.Nome_Razao " _
& "FROM tbl_contas_receber INNER JOIN" _
& " tbl_Detalhes_Recebimento ON tbl_contas_receber.IDIntconta = tbl_Detalhes_Recebimento.IDContaReceber" _
& " ORDER BY tbl_contas_receber.Nome_Razao"

With frm_Instituicoes.cmbcliente
    .Clear
    .AddItem ""
    Set TBCarregarCombo = CreateObject("adodb.recordset")
    TBCarregarCombo.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
    If TBCarregarCombo.EOF = False Then
        
        Do While TBCarregarCombo.EOF = False
            If IsNull(TBCarregarCombo!Nome_Razao) = False And TBCarregarCombo!Nome_Razao <> "" Then
            If NomeRazao <> TBCarregarCombo!Nome_Razao Then
                .AddItem TBCarregarCombo!Nome_Razao
                .ItemData(.NewIndex) = TBCarregarCombo!IdContaReceber
                NomeRazao = TBCarregarCombo!Nome_Razao
            End If
            End If
            TBCarregarCombo.MoveNext
        Loop
        TBCarregarCombo.MoveFirst
    End If
    TBCarregarCombo.Close
End With

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcCarregacomboCarteira()
On Error GoTo tratar_erro

With frm_Instituicoes.cmbCarteira
    .Clear
        
Select Case frm_Instituicoes.txtNBanco
    Case "341": 'Ita˙
           .AddItem "109 - Direta EletrÙnica Sem Emiss„o - Simples":
           .AddItem "112 - Escritual EletrÙnica - simples / contratual":
           .AddItem "175 - Sem Registro Sem Emiss„o":
           .Text = "109 - Direta EletrÙnica Sem Emiss„o - Simples":
    Case "001": 'Banco do brasil
            .AddItem "11 - Simples - Com Registro":
            .AddItem "11 - Vinculada - Com Registro":
            .AddItem "17 - Direta Especial - Com Registro":
            .AddItem "17Simples - Direta Especial Simples - Com Registro":
            .AddItem "17-7 - Direta Especial - Com Registro ConvÍnio 7 dÌgitos":
            .AddItem "18 - Simples - Sem Registro":
            .AddItem "18-7 - Simples - Sem Registro - ConvÍnio 7 dÌgitos":
            .Text = "11 - Vinculada - Com Registro":
    Case "033": 'Santander
            .AddItem "COB - CobranÁa Simples":
            .AddItem "COBR - CobranÁa Simples - R·pida Com Registro":
            .AddItem "COBR-Nova - CobranÁa Simples - R·pida Com Registro"
            .AddItem "CSR - CobranÁa Simples Sem Registro":
            .AddItem "ECR - CobranÁa Simples Com Registro":
            .Text = "COBR - CobranÁa Simples - R·pida Com Registro":
    Case "104": 'Caixa
            .AddItem "CR - CobranÁa R·pida":
            .AddItem "SR - CobranÁa Sem Registro":
            .AddItem "SIG14 - SIG Com Registro - Emiss„o pelo Cedente":
            .Text = "SIG14 - SIG Com Registro - Emiss„o pelo Cedente":
    Case "237": 'Bradesco
            .AddItem "09 - Com Registro":
            .Text = "09 - Com Registro"
    
    
    Case "356": 'ABN e Real
            .AddItem "20 - CobranÁa Simples":
            .Text = "20 - CobranÁa Simples":
    Case "399": 'HSBC
            .AddItem "CNR - Sem Registro":
            .Text = "CNR - Sem Registro":
    Case "409": 'Unibanco
            .AddItem "Especial":
            .Text = "Especial":
        End Select
End With

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcCarregaInstrucoesBoleto()
On Error GoTo tratar_erro

With frm_Instituicoes
            .Txtpercentual_juros = ""
            .Txtpercentual_desconto = ""
            .Txtpercentual_multa = ""
            .Txtdias_protesto = ""
            .txtInstrucoes = ""
            .txtAssunto = ""

Set TBBoleto = CreateObject("adodb.recordset")
    TBBoleto.Open "Select * from tbl_Instituicoes_Instrucoes_Boleto where ID_Instituicao = " & .txtId & "", Conexao, adOpenKeyset, adLockOptimistic
        If TBBoleto.EOF = False Then
            .Txtpercentual_juros = TBBoleto!Juros
            .Txtpercentual_desconto = TBBoleto!Desconto
            .Txtpercentual_multa = TBBoleto!Multa
            .Txtdias_protesto = TBBoleto!Dias_Protesto
            .txtInstrucoes = TBBoleto!Instrucoes_protesto
            .txtAssunto = TBBoleto!AssuntoEmail
    TBBoleto.Close
        End If
End With
  
Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub

End Sub

Public Sub ProcCarregaInstituicaoBoleto()
On Error GoTo tratar_erro

Set TBFIltro = CreateObject("adodb.recordset")
TBFIltro.Open "Select * from Tbl_Instituicoes where txt_descricao = '" & frm_Instituicoes.txtdescricao.Text & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBFIltro.EOF = False Then
With frm_Instituicoes
    .txtConta = IIf(IsNull(TBFIltro!txt_Conta) = False, TBFIltro!txt_Conta, "")
    .txtId = TBFIltro!ID
    If DS.FileOrDirExists(Localrel & "\Imagens\Bancos\" & TBFIltro!int_NBanco & ".jpg") = True Then
    .Logo_Banco.Picture = LoadPicture(Localrel & "\Imagens\Bancos\" & TBFIltro!int_NBanco & ".jpg")
    .Logo_Banco.Visible = True
    Else
    .Logo_Banco.Visible = False
    End If
    .txtAgencia = IIf(IsNull(TBFIltro!txt_Agencia) = False, TBFIltro!txt_Agencia, "")
    .txtNBanco = IIf(IsNull(TBFIltro!int_NBanco) = False, TBFIltro!int_NBanco, "")
    .Txt_codigo_cedente1 = IIf(IsNull(TBFIltro!Codigo_cedente_registrado) = False, TBFIltro!Codigo_cedente_registrado, "")
    .txtcarteiraconf = ArquivoLicensa
    .Txt_nome_agencia = IIf(IsNull(TBFIltro!Nome_agencia) = False, TBFIltro!Nome_agencia, "")
    If TBFIltro!DiretorioRemessa = "" Or IsNull(TBFIltro!DiretorioRemessa) Then
    .txtlocal = Localrel & "\Boletos\Arquivos remessa\" & frm_Instituicoes.txtdescricao.Text
    Else
    .txtlocal = TBFIltro!DiretorioRemessa
    End If
    If TBFIltro!DiretorioBoleto = "" Or IsNull(TBFIltro!DiretorioBoleto) Then
    .txtLocalBoleto = Localrel & "\Boletos\Arquivos pdf\" & frm_Instituicoes.txtdescricao.Text
    Else
    .txtLocalBoleto = TBFIltro!DiretorioBoleto
    End If
End With
End If


TBFIltro.Close

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcCarregaComboBancoConciliacao()
On Error GoTo tratar_erro

With Frm_InstituicoesConciliacao.cmbBanco
    .Clear
    Set TBCarregarCombo = CreateObject("adodb.recordset")
    TBCarregarCombo.Open "Select * from tbl_Instituicoes order by txt_Descricao", Conexao, adOpenKeyset, adLockOptimistic
    If TBCarregarCombo.EOF = False Then
        
        Do While TBCarregarCombo.EOF = False
            If IsNull(TBCarregarCombo!Txt_descricao) = False And TBCarregarCombo!Txt_descricao <> "" Then
                .AddItem TBCarregarCombo!Txt_descricao
                .ItemData(.NewIndex) = TBCarregarCombo!ID
            End If
            TBCarregarCombo.MoveNext
        Loop
        TBCarregarCombo.MoveFirst
        Frm_InstituicoesConciliacao.cmbBanco.Text = TBCarregarCombo!Txt_descricao
    End If
End With
TBCarregarCombo.Close

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcBuscaClienteSintegra(CNPJCliente As String)
On Error GoTo tratar_erro
'======================================================
' Buscar dados do cnpj no sintegra
'======================================================

Dim resposta As String
Dim obj As MSXML2.ServerXMLHTTP50
Set obj = New MSXML2.ServerXMLHTTP50
Dim Plugin As String

Plugin = "ST"


CnpjDados = ReturnNumbersOnly(CNPJCliente)

1:
obj.Open "GET", "https://www.sintegraws.com.br/api/v1/execute-api.php?token=1F718E4E-3222-42F1-95D6-995FC9E69C9C&cnpj=" & CnpjDados & "&plugin=" & Plugin & ""

conteudo = CnpjDados
obj.send conteudo
resposta = obj.responseText
'Debug.print resposta

If LerDadosJSON(resposta, "status", "", "") = "OK" And LerDadosJSON(resposta, "code", "", "") = "0" Then

If Plugin = "ST" Then
    NomeRazao = UCase(LerDadosJSON(resposta, "nome_empresarial", "", ""))
    NomeFantasia = UCase(LerDadosJSON(resposta, "nome_fantasia", "", ""))
    RegimeTributario = IIf(LerDadosJSON(resposta, "regime_tributacao", "", "") = "Normal - regime periÛdico de apuraÁ„o", "Lucro presumido", "Simples Nacional")
    RG_IE = Trim(LerDadosJSON(resposta, "inscricao_estadual", "", ""))
Else
    NomeRazao = UCase(LerDadosJSON(resposta, "nome", "", ""))
    NomeFantasia = UCase(LerDadosJSON(resposta, "fantasia", "", ""))
End If
    UF = UCase(LerDadosJSON(resposta, "uf", "", ""))
    Bairro = UCase(LerDadosJSON(resposta, "bairro", "", ""))
    Endereco = UCase(LerDadosJSON(resposta, "logradouro", "", ""))
    Numero = LerDadosJSON(resposta, "numero", "", "")
    CEP = LerDadosJSON(resposta, "cep", "", "")
    Cidade = UCase(LerDadosJSON(resposta, "municipio", "", ""))

'    'Debug.print resposta
    CPF_CNPJ = CNPJCliente
    Categoria = "A"
Else
    Plugin = "RF"
    GoTo 1
    'USMsgBox LerDadosJSON(resposta, "message", "", ""), vbInformation, "CAPRIND v5.0"
End If

Exit Sub
tratar_erro:
    MousePointer = 0
    If Err.Number = 91 Then
        USMsgBox ("N„o foi possÌvel carregar todos os dados referentes a este CEP."), vbInformation, "CAPRIND v5.0"
        Exit Sub
    End If
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcCriarNumeroProposta()
On Error GoTo tratar_erro

Set TBAbrir = CreateObject("adodb.recordset")
StrSql = "Select Ncotacao from vendas_proposta where Year(Data) = '" & Year(Date) & "' order by Ordenarproposta desc"
'Debug.print StrSql

    TBAbrir.Open "Select Ncotacao from vendas_proposta where Year(Data) = '" & Year(Date) & "' order by Ordenarproposta desc", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
    TBAbrir.MoveFirst
        Cotacao = Left(TBAbrir!Ncotacao, Len(TBAbrir!Ncotacao) - 3) + 1
    Else
        Cotacao = 1
    End If
    Ano = Right(Year(Date), 2)
    Select Case Len(Cotacao)
        Case 1: NumeroCotacao = "000" & Cotacao & "/" & Ano
        Case 2: NumeroCotacao = "00" & Cotacao & "/" & Ano
        Case 3: NumeroCotacao = "0" & Cotacao & "/" & Ano
        Case 4: NumeroCotacao = Cotacao & "/" & Ano
        Case 5: NumeroCotacao = Cotacao & "/" & Ano
    End Select

NProposta = NumeroCotacao

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcCarregaComboBancoBoleto()
On Error GoTo tratar_erro
'
'With frm_Instituicoes.txtDescricao
'    .Clear
'    Set TBCarregarCombo = CreateObject("adodb.recordset")
'    TBCarregarCombo.Open "Select * from tbl_Instituicoes order by txt_Descricao", Conexao, adOpenKeyset, adLockOptimistic
'    If TBCarregarCombo.EOF = False Then
'
'        Do While TBCarregarCombo.EOF = False
'            If IsNull(TBCarregarCombo!Txt_descricao) = False And TBCarregarCombo!Txt_descricao <> "" Then
'                .AddItem TBCarregarCombo!Txt_descricao
'                .ItemData(.NewIndex) = TBCarregarCombo!ID
'            End If
'            TBCarregarCombo.MoveNext
'        Loop
'        TBCarregarCombo.MoveFirst
'        frm_Instituicoes.cmbBanco.Text = TBCarregarCombo!Txt_descricao
'        ProcCarregacomboCarteira
'    End If
'End With
'TBCarregarCombo.Close
'
Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaListaDuplicatas()
On Error GoTo tratar_erro
Dim SQLBusca As String

Init = 0
Sit_REG = 0
valor = 0
If StrSql = "" Then Exit Sub

frm_Instituicoes.lst_Duplicata.ListItems.Clear

Set TBLISTA = CreateObject("adodb.recordset")
'Debug.print StrSql

TBLISTA.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    'TBLISTA.MoveLast
    Contador = 0
    'TBLISTA.MoveFirst
    Do While TBLISTA.EOF = False
        With frm_Instituicoes.lst_Duplicata.ListItems
            .Add , , TBLISTA!IdContaReceber
            .Item(.Count).SubItems(1) = TBLISTA!int_NotaFiscal
            .Item(.Count).SubItems(2) = TBLISTA!Nome_Razao
            .Item(.Count).SubItems(3) = Format(TBLISTA!dt_Vencimento, "dd/mm/yyyy")
            .Item(.Count).SubItems(4) = TBLISTA!txt_Parcela
            .Item(.Count).SubItems(5) = Format(TBLISTA!dbl_Valor, "###,##0.00")
            .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA!Nosso_Numero), "", TBLISTA!Nosso_Numero)
            
            If IsNull(TBLISTA!Seq_remessa) = False And TBLISTA!Seq_remessa <> "" And TBLISTA!txt_Portador_Banco <> "" Then
                ProcPeganumeroremessa
                .Item(.Count).SubItems(7) = Arquivo
            End If
            
            If TBLISTA!IdContaReceber = "" Then
            .Item(.Count).SubItems(8) = "N„o"
            Else
            .Item(.Count).SubItems(8) = "Sim"
            End If
            
            If TBLISTA!Enviado = True Then
            .Item(.Count).SubItems(9) = "Sim"
            Else
            .Item(.Count).SubItems(9) = "N„o"
            End If
            
            
            Init = Init + 1
            .Item(Init).Checked = False
        End With
                
        TBLISTA.MoveNext
        Contador = Contador + 1
    Loop

End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Function ValidaEmail(strEmail As String) As Boolean
    Dim nCharacter As Integer
    Dim Count As Integer
    Dim sLetra As String
    'Verifica se o e-mail tem no MÕNIMO 5
    'caracteres (a@b.c)
    If Len(strEmail) < 5 Then
    'O e-mail È inv·lido, pois tem menos
    'de 5 caracteres
        ValidaEmail = False
        MsgBox "O e-mail digitado tem menos de 5 caracteres", vbInformation, TITULO_AVISO
        Exit Function
    'Verificar a existencia de arrobas
    ElseIf InStr(strEmail, "@") = Len(strEmail) Then
    'O e-mail È inv·lido, pois termina
    'com uma @
        ValidaEmail = False
        MsgBox "O e-mail termina com uma arroba", vbInformation, TITULO_AVISO
        Exit Function
    ElseIf Mid$(strEmail, (InStr(strEmail, "@")) + 1, 1) = "." Then
    'Valida se h· ponto apÛs o arroba
        ValidaEmail = False
        MsgBox "ApÛs a arroba, contÈm um ponto inv·lido", vbInformation, TITULO_AVISO
        Exit Function
    ElseIf InStr(Mid$(strEmail, (InStr(strEmail, "@") + 1), (Len(strEmail) - (InStr(strEmail, "@") + 1))), "@") > 0 Then
 'Valida se h· mais de um arroba
        ValidaEmail = False
        MsgBox "E-mail contÈm mais de um arroba", vbInformation, TITULO_AVISO
        Exit Function
    End If
    nCharacter = 0
    Count = 0
    'Verificar a existencia de pontos (.) no e-mail
    For nCharacter = 1 To Len(strEmail)
        If Mid(strEmail, nCharacter, 1) = "." Then
    'OPA!!! Achou um ponto!!!
    'Soma 1 ao contador
            Count = Count + 1
        End If
    Next
    'Verifica o n˙mero de pontos.
    'TEM que ter PELO MENOS UM ponto.
    If Count < 1 Then
    'O e-mail È inv·lido, pois n„o tem pontos.
        ValidaEmail = False
        MsgBox "O e-mail È inv·lido, pois n„o contÈm (.) pontos.", vbInformation, TITULO_AVISO
        Exit Function
    Else
    'O e-mail tem pelo menos 1 ponto.
    'Verificar a posiÁ„o do ponto:
        If InStr(strEmail, ".") = 1 Then
        'O e-mail È inv·lido, pois comeÁa
        'com um ponto
            ValidaEmail = False
            MsgBox "O e-mail foi iniciado com um ponto (.)", vbInformation, TITULO_AVISO
            Exit Function
        ElseIf InStr(strEmail, ".") = Len(strEmail) Then
        'O e-mail È inv·lido, pois termina
        'com um ponto.
            ValidaEmail = False
            MsgBox "O e-mail termina com um ponto (.)", vbInformation, TITULO_AVISO
            Exit Function
        ElseIf InStr(InStr(strEmail, "@"), strEmail, ".") = 0 Then
        'O e-mail È inv·lido, pois termina
        'com um ponto.
            ValidaEmail = False
            MsgBox "O e-mail n„o tem nenhum ponto (.) apÛs " & _
            "a arroba.", vbInformation, TITULO_AVISO
            Exit Function
        End If
    End If
    nCharacter = 0
    Count = 0
    'Verifica se o e-mail n„o tem pontos
    'consecutivos (..) apÛs a arroba http://www.babooforum.com.br/idealbb/images/smilies/22.gif'>!!! Exit Function
    For nCharacter = 1 To Len(strEmail)
        sLetra = Mid$(strEmail, nCharacter, 1)
        If Not (LCase(sLetra) Like "[a-z]" Or sLetra = _
        "@" Or sLetra = "." Or sLetra = "-" Or _
        sLetra = "_" Or IsNumeric(sLetra)) Then
    'O e-mail È inv·lido, pois tem
    'caracteres inv·lidos
            ValidaEmail = False
            MsgBox "Foi digitado um caracter inv·lido no e-mail", vbInformation, TITULO_AVISO
            Exit Function
        End If
    Next
    nCharacter = 0
    'Bem, se a verificaÁ„o chegou atÈ aqui
    'È porque o e-mail È v·lido, ent„o...
    ValidaEmail = True
End Function

Public Sub ProcPeganumeroremessa()
On Error GoTo tratar_erro

Arquivo = ""

'Verifica o ˙ltimo sequencial no banco para gerar o prÛximo
If IsNull(TBLISTA!data_envio) = True Then Exit Sub

Dia = Day(TBLISTA!data_envio)
If Len(Dia) = 1 Then Dia = "0" & Dia
Mes = Month(TBLISTA!data_envio)
If Len(Mes) = 1 Then Mes = "0" & Mes
Ano = Year(TBLISTA!data_envio)

    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select Seq_remessa from tbl_Detalhes_Recebimento where IDContaReceber = '" & TBLISTA!IdContaReceber & "' order by Seq_remessa desc", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        If IsNull(TBAbrir!Seq_remessa) = False And TBAbrir!Seq_remessa <> "" Then Seq = TBAbrir!Seq_remessa Else Seq = 1
    End If
    TBAbrir.Close
    
If frm_Instituicoes.txtNBanco.Text = "341" Then 'Itau then
    'O nome do arquivo remessa do Ita˙ sÛ aceita no m·ximo 8 caracteres
    'seqremessa = Seq
    If Seq < 10 Then SeqRemessa = "0" & Seq & ".txt" Else SeqRemessa = Seq & ".txt"
    SeqRemessaTexto = Left(SeqRemessa, Len(SeqRemessa) - 4)
    Select Case Len(SeqRemessaTexto)
        Case 1: RemessaTexto = "0" & Right(SeqRemessaTexto, 1)
        Case 2: RemessaTexto = SeqRemessaTexto
        Case Is >= 3: RemessaTexto = Right(SeqRemessaTexto, 2)
    End Select
    Arquivo = Dia & Mes & Right(Ano, 2) & RemessaTexto & ".txt"
   ' Layout = "CNAB400"
    'CobreBemX1.ArquivoRemessa.Sequencia = Left(SeqRemessa, Len(SeqRemessa) - 4)
End If

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaVariaveisPrincipais()
On Error GoTo tratar_erro

StrSql = ""
SQL = ""
NomeRel = ""
TabelaRel = 0
TabelaRel1 = 0
TabelaRel2 = 0
CampoRel = 0
CampoRel1 = 0
CampoRel = 0
CarregaListaForm = False
Sair = False
CodigoLista = 0
CodigoLista1 = 0
CodigoLista2 = 0
CodigoLista3 = 0
CodigoLista4 = 0
CodigoLista5 = 0
CodigoLista6 = 0
ID_documento = 0
Documento1 = ""
StrSqlLocProdPadrao = ""
StrSqlLocCliPadrao = ""
StrSqlLocFornPadrao = ""

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcVerificaAcao()
On Error GoTo tratar_erro

USMsgBox ("Informe " & NomeCampo & " antes de " & Acao & "."), vbExclamation, "CAPRIND v5.0"

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcVerificaSalvar()
On Error GoTo tratar_erro

USMsgBox ("Clique no bot„o <Novo> e preencha todos os campos antes de salvar."), vbExclamation, "CAPRIND v5.0"

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcVerificaNumero()
On Error GoTo tratar_erro

If (IsNumeric(VerifNumero)) = False Or InStr(VerifNumero, "-") <> 0 And InStr(VerifNumero, "-") <> 1 Then
'    usMsgbox ("SÛ È permitido n˙mero neste campo."), vbExclamation, "CAPRIND v5.0"
    VerifNumero = False
Else
    VerifNumero = True
End If

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritica
End Sub

Sub ProcVerificaData()
On Error GoTo tratar_erro

If IsDate(VerifData) = False Then
    USMsgBox ("A data foi digitada incorretamente."), vbExclamation, "CAPRIND v5.0"
    VerifData = False
Else
    VerifData = True
End If

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Function FunVerificaDataFinal(Dataini As Date, DataFim As Date) As Boolean
On Error GoTo tratar_erro

FunVerificaDataFinal = True
If DataFim < Dataini Then
    USMsgBox ("A data final n„o pode ser menor que a data inicial."), vbExclamation, "CAPRIND v5.0"
    FunVerificaDataFinal = False
End If

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Sub ProcImprimirRel(FormulaRel As String, FormulaRelSubReport As String)
On Error GoTo tratar_erro

ProcVerifRelPersonalizado

If PermitidoRel = False Then LocalrelNovo = Localrel Else LocalrelNovo = LocalRelPersonalizado
'Debug.print NomeRel
Set Report = crAPP.OpenReport(LocalrelNovo & "\" & NomeRel)
'Login SQL
Contador = Report.Database.Tables.Count

Do While Contador > 0
    Set DBTable = Report.Database.Tables(Contador)
    ProcLogonBDSQL
    Contador = Contador - 1
Loop

ProcVerifSubReport FormulaRelSubReport
'Debug.print FormulaRel

frmimprimir.CrystalActiveXReportViewer1.ReportSource = Report
Report.FormulaSyntax = crCrystalSyntaxFormula
'Debug.print FormulaRel
Report.RecordSelectionFormula = FormulaRel
frmimprimir.CrystalActiveXReportViewer1.ViewReport
frmimprimir.CrystalActiveXReportViewer1.DisplayGroupTree = False
frmimprimir.WindowState = vbMaximized
frmimprimir.Show 1
2:
    Set Report = Nothing
    Set crAPP = Nothing

Exit Sub
tratar_erro:
    If Err.Number = "-2147206461" Then
        USMsgBox ("N„o foi encontrado o relatÛrio " & NomeRel & " na pasta " & LocalrelNovo), vbExclamation, "CAPRIND v5.0"
        GoTo 2
    End If
    If Err.Number = "-2147483638" Then
        USMsgBox ("N„o foi possÌvel visualizar o relatÛrio, favor reiniciar o sistema."), vbExclamation, "CAPRIND v5.0"
        GoTo 2
    End If
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcVerifRelPersonalizado()
On Error GoTo tratar_erro

PermitidoRel = False
Set TBMaterial = CreateObject("adodb.recordset")
TBMaterial.Open "Select * from Qualidade_revisao_relatorios where Nome_relatorio = '" & NomeRel & "' and Personalizado = 'True' order by Revisao desc, ID desc", Conexao, adOpenKeyset, adLockOptimistic
If TBMaterial.EOF = False Then
    LocalRelPersonalizado = Localrel & "\Personalizados"
    PermitidoRel = True
End If
TBMaterial.Close

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcVerifSubReport(FormulaRelSubReport As String)
On Error GoTo tratar_erro

Contador2 = 0
NomeSubReport = ""
NomeSubReport1 = ""
NomeSubReport2 = ""
NomeSubReport3 = ""
NomeSubReport4 = ""
NomeSubReport5 = ""
NomeSubReport6 = ""
NomeSubReport7 = ""
NomeSubReport8 = ""
NomeSubReport9 = ""
NomeSubReport10 = ""

Set TBSubreport = CreateObject("adodb.recordset")
TBSubreport.Open "Select * from Qualidade_revisao_relatorios_subreports where Nome_relatorio = '" & NomeRel & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBSubreport.EOF = False Then
    Do While TBSubreport.EOF = False
        Select Case Contador2
            Case 0: NomeSubReport = TBSubreport!SubReport
            Case 1: NomeSubReport1 = TBSubreport!SubReport
            Case 2: NomeSubReport2 = TBSubreport!SubReport
            Case 3: NomeSubReport3 = TBSubreport!SubReport
            Case 4: NomeSubReport4 = TBSubreport!SubReport
            Case 5: NomeSubReport5 = TBSubreport!SubReport
            Case 6: NomeSubReport6 = TBSubreport!SubReport
            Case 7: NomeSubReport7 = TBSubreport!SubReport
            Case 8: NomeSubReport8 = TBSubreport!SubReport
            Case 9: NomeSubReport9 = TBSubreport!SubReport
            Case 10: NomeSubReport10 = TBSubreport!SubReport
        End Select
        Contador2 = Contador2 + 1
        TBSubreport.MoveNext
    Loop
End If
TBSubreport.Close
If Contador2 > 0 Then ProcSubReport FormulaRelSubReport

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcExcluirSaida_NFe(ID_nota As Long)
On Error GoTo tratar_erro

'Conta gerada pela nota
Conexao.Execute "DELETE from CC from CC_realizado CC INNER JOIN tbl_contas_receber CR ON CR.IDIntconta = CC.ID_financeiro Where CR.ID_Nota = " & ID_nota & " and CC.Operacao = 'CrÈdito'"
Conexao.Execute "DELETE from FF from Familia_financeiro FF INNER JOIN tbl_contas_receber CR ON CR.IDIntconta = FF.IDconta Where CR.ID_Nota = " & ID_nota & " and FF.Tipoconta = 'R' and (CR.Proposta IS NULL or CR.Proposta = N'')"
Conexao.Execute "DELETE from FC from tbl_Fluxo_de_caixa FC INNER JOIN tbl_contas_receber CR ON CR.IDFluxo = FC.IDFluxo Where CR.ID_Nota = " & ID_nota & " and (CR.Proposta IS NULL or CR.Proposta = N'')"
Conexao.Execute "DELETE from tbl_contas_receber where ID_Nota = " & ID_nota & " and (Proposta IS NULL or Proposta = N'')"

'Conta gerada pelo pedido
Conexao.Execute "Update FC set FC.int_NotaFiscal = NULL from tbl_Fluxo_de_caixa FC INNER JOIN tbl_contas_receber CR ON CR.IDFluxo = FC.IDFluxo Where CR.ID_Nota = " & ID_nota & " and CR.Proposta IS NOT NULL"
Conexao.Execute "Update tbl_contas_receber Set ID_nota = NULL, NFiscal = NULL where ID_Nota = " & ID_nota & " and Proposta IS NOT NULL"

'Conta gerada pela nota
Conexao.Execute "DELETE from FF from Familia_financeiro FF INNER JOIN tbl_ContasPagar CP ON CP.IDIntconta = FF.IDconta Where CP.ID_Nota = " & ID_nota & " and FF.Tipoconta = 'P' and (CP.txt_pedido IS NULL or CP.txt_pedido = N'')"
Conexao.Execute "DELETE from FC from tbl_Fluxo_de_caixa FC INNER JOIN tbl_ContasPagar CP ON CP.IDFluxo = FC.IDFluxo Where CP.ID_Nota = " & ID_nota & " and (CP.txt_pedido IS NULL or CP.txt_pedido = N'')"
Conexao.Execute "DELETE from tbl_ContasPagar where ID_Nota = " & ID_nota & "  and (txt_pedido IS NULL or txt_pedido = N'')"
'Conta gerada pelo pedido
Conexao.Execute "Update FC set FC.int_NotaFiscal = NULL from tbl_Fluxo_de_caixa FC INNER JOIN tbl_ContasPagar CP ON CP.IDFluxo = FC.IDFluxo Where CP.ID_Nota = " & ID_nota & " and CP.txt_pedido IS NOT NULL"
Conexao.Execute "Update tbl_ContasPagar Set ID_nota = NULL, txt_ndocumento = NULL where ID_Nota = " & ID_nota & " and txt_pedido IS NOT NULL"
Conexao.Execute "Update CC set CC.ID_Financeiro = 0 from CC_realizado CC INNER JOIN tbl_Detalhes_Recebimento TBL on CC.ID_duplicata = TBL.ID where TBL.ID_nota = " & ID_nota
         
Set TBNivel10 = CreateObject("adodb.recordset")
TBNivel10.Open "Select TBL.Int_codigo, TBL.int_Cod_Produto, TBL.int_Qtd, TBL.ID_nota, TBL.N_Referencia, TBL.Int_NotaFiscal, P.ID_PC, P.Codproduto from (tbl_detalhes_nota TBL INNER JOIN ProjProduto P ON TBL.Codproduto = P.Codproduto) INNER JOIN tbl_NaturezaOperacao CFOP ON CFOP.IDCountCfop = TBL.ID_CFOP where TBL.ID_Nota = " & ID_nota & " and P.Estoque = 'True' and (CFOP.Remessa IS NULL or CFOP.Remessa = 'False')", Conexao, adOpenKeyset, adLockOptimistic
Do While TBNivel10.EOF = False
    quantidade = 0
    Qtde = 0
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from estoque_movimentacao where ID_Prod_NF = " & TBNivel10!Int_codigo, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Do While TBAbrir.EOF = False
            valor = IIf(IsNull(TBAbrir!VlrUnit), 0, TBAbrir!VlrUnit)
            'Atualiza qtde. expedida
            Qtd = IIf(IsNull(TBAbrir!Saida), 0, TBAbrir!Saida)
            
            Set TBGravar = CreateObject("adodb.recordset")
            TBGravar.Open "Select PP.Qtdeexpedida, PP.Dataexpedicao, NFPP.* from vendas_carteira PP INNER JOIN tbl_Detalhes_Nota_pedidos NFPP ON NFPP.ID_carteira = PP.Codigo and NFPP.Codinterno = PP.Desenho where NFPP.ID_prod_NF = " & TBAbrir!ID_prod_NF & " order by PP.PrazoFinal", Conexao, adOpenKeyset, adLockOptimistic
            If TBGravar.EOF = False Then
                Do While TBGravar.EOF = False
                    If Qtd >= TBGravar!qtdeexpedida Then qt = TBGravar!qtdeexpedida Else qt = Qtd
                    TBGravar!qtdeexpedida = TBGravar!qtdeexpedida - qt
                    Qtd = Qtd - qt
                    
                    Set TBFI = CreateObject("adodb.recordset")
                    TBFI.Open "Select Data from Estoque_movimentacao where Idoperacao <> " & TBAbrir!IDoperacao & " and ID_prod_NF = " & IIf(IsNull(TBAbrir!ID_prod_NF), 0, TBAbrir!ID_prod_NF) & " and (Operacao = 'SAIDA_NOTA' or Operacao = 'SAIDA_NOTA_PARCIAL') order by Data desc", Conexao, adOpenKeyset, adLockOptimistic
                    If TBFI.EOF = False Then
                        TBGravar!dataexpedicao = TBFI!Data
                    Else
                        TBGravar!dataexpedicao = Null
                    End If
                    TBFI.Close
                    TBGravar.Update
                    
                    'Desvincula pedido da ordem para estoque
                    If IsNumeric(TBAbrir!LOTE) = True Then
                        Set TBAliquota = CreateObject("adodb.recordset")
                        TBAliquota.Open "Select * from Producao_pedidos where IDcarteira = " & IIf(IsNull(TBGravar!ID_carteira), 0, TBGravar!ID_carteira) & " and Ordem = " & TBAbrir!LOTE & " and Expedicao = 'True'", Conexao, adOpenKeyset, adLockOptimistic
                        If TBAliquota.EOF = False Then
                            TBAliquota!Qtde_empenho = TBAliquota!Qtde_empenho - qt
                            TBAliquota!Qtde_entrada = TBAliquota!Qtde_empenho
                            TBAliquota.Update
                            
                            If TBAliquota!Qtde_empenho <= 0 Then Conexao.Execute "DELETE from Producao_pedidos where IDcarteira = " & IIf(IsNull(TBGravar!ID_carteira), 0, TBGravar!ID_carteira) & " and Ordem = " & TBAbrir!LOTE & " and Expedicao = 'True'"
                        End If
                        TBAliquota.Close
                    End If
                    
                    Do While qt > 0
                        'Atualiza qtde. de saÌda no empenho
                        Set TBAliquota = CreateObject("adodb.recordset")
                        TBAliquota.Open "Select EE.Qtde_saida from Estoque_Controle_Empenho_Vendas EE INNER JOIN tbl_Detalhes_Nota_pedidos NFPP ON EE.ID_carteira = NFPP.ID_carteira where NFPP.ID_prod_NF = " & TBAbrir!ID_prod_NF & " and EE.ID_estoque = " & TBAbrir!IDEstoque & " and EE.Qtde_saida > 0", Conexao, adOpenKeyset, adLockOptimistic
                        If TBAliquota.EOF = False Then
                            If TBAliquota!Qtde_saida >= qt Then
                                TBAliquota!Qtde_saida = TBAliquota!Qtde_saida - qt
                                qt = 0
                            Else
                                qt = qt - TBAliquota!Qtde_saida
                                TBAliquota!Qtde_saida = 0
                            End If
                            TBAliquota.Update
                        Else
                            GoTo Prosseguir
                        End If
                        TBAliquota.Close
                    Loop
Prosseguir:
                    If Qtd <= 0 Then GoTo Prosseguir1
                    TBGravar.MoveNext
                Loop
            End If
    
Prosseguir1:
            IDEstoque = TBAbrir!IDEstoque
            quantidade = IIf(IsNull(TBAbrir!Saida), 0, TBAbrir!Saida)
            TBAbrir.Delete
            
            'Centro de custo
            Conexao.Execute "DELETE from CC_realizado where ID_estoque = " & IDEstoque
                    
            'Corrige estoque real
            Set TBEstoque = CreateObject("adodb.recordset")
            TBEstoque.Open "Select * from estoque_controle where idestoque = " & IDEstoque, Conexao, adOpenKeyset, adLockOptimistic
            If TBEstoque.EOF = False Then
                '==================================
                Modulo = "Faturamento/Nota fiscal/PrÛpria"
                Evento = "Excluir movimentaÁ„o"
                ID_documento = IDEstoque
                Documento = "CÛd. interno: " & TBEstoque!Desenho & " - N∫ lote: " & TBEstoque!LOTE & " - N∫ corrida: " & TBEstoque!Corrida & " - N∫ certificado: " & TBEstoque!Certificado & " - Local armaz.: " & TBEstoque!local_armaz
                Documento1 = ""
                ProcGravaEvento
                '==================================
                Qtd = IIf(IsNull(TBEstoque!estoque_real), 0, TBEstoque!estoque_real)
                Qtde = Format(Qtd + quantidade, "###,##0.0000")
                TBEstoque!estoque_real = Format(Qtde, "###,##0.0000")
                TBEstoque!Qtde = Qtde
                TBEstoque!estoque_real_PC = 0
                                                
                'Atualiza valor do material no estoque
                TBEstoque!Valor_total = Format(IIf(IsNull(TBEstoque!valor_unitario), 0, TBEstoque!valor_unitario) * Qtde, "###,##0.00")
                        
                TBEstoque.Update
                Set TBMaterial = CreateObject("adodb.recordset")
                TBMaterial.Open "Select * from Estoque_movimentacao where IDEstoque = " & TBEstoque!IDEstoque, Conexao, adOpenKeyset, adLockOptimistic
                If TBMaterial.EOF = True Then TBEstoque.Delete
                TBMaterial.Close
            End If
            TBEstoque.Close

            TBAbrir.MoveNext
        Loop
    End If
    TBAbrir.Close
    TBNivel10.MoveNext
Loop
TBNivel10.Close


Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub ProcSubReport(FormulaRelSubReport As String)
On Error GoTo tratar_erro

Do While Contador2 > 0
    Select Case Contador2
        Case "10": SubReportRel = NomeSubReport9
        Case "9": SubReportRel = NomeSubReport8
        Case "8": SubReportRel = NomeSubReport7
        Case "7": SubReportRel = NomeSubReport6
        Case "6": SubReportRel = NomeSubReport5
        Case "5": SubReportRel = NomeSubReport4
        Case "4": SubReportRel = NomeSubReport3
        Case "3": SubReportRel = NomeSubReport2
        Case "2": SubReportRel = NomeSubReport1
        Case "1": SubReportRel = NomeSubReport
    End Select
    
    Contador = Report.OpenSubreport(SubReportRel).Database.Tables.Count
    Do While Contador > 0
        Set DBTable = Report.OpenSubreport(SubReportRel).Database.Tables(Contador)
        ProcLogonBDSQL
        Contador = Contador - 1
        
        'Coloca a formula no subreport
        If FormulaRelSubReport <> "" And SubReportRel <> "RevisaoRelatorio.rpt" And SubReportRel <> "Responsavel_relatorio" Then
            Report.OpenSubreport(SubReportRel).FormulaSyntax = crCrystalSyntaxFormula
            Report.OpenSubreport(SubReportRel).RecordSelectionFormula = FormulaRelSubReport
        End If
        
        'Coloca a formula no subreport de responsavel
        If SubReportRel = "Responsavel_relatorio" Then
            'verifica se esta marcado para imprimir o relatorio com o responsavel
            Set TBMaterial = CreateObject("adodb.recordset")
            TBMaterial.Open "Select Responsavel_rel from Qualidade_revisao_relatorios where Nome_relatorio = '" & NomeRel & "' and Responsavel_rel = 1", Conexao, adOpenKeyset, adLockOptimistic
            If TBMaterial.EOF = False Then
                Report.OpenSubreport(SubReportRel).FormulaSyntax = crCrystalSyntaxFormula
                Report.OpenSubreport(SubReportRel).RecordSelectionFormula = "{Usuarios.IDusuario} = " & pubIDUsuario
            End If
            TBMaterial.Close
        End If
    Loop
1:
    Contador2 = Contador2 - 1
Loop

Exit Sub
tratar_erro:
    If Err.Number = "-2147190528" Then GoTo 1
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLogonBDSQL()
On Error GoTo tratar_erro

Set CPProperty = DBTable.ConnectionProperties("Data Source")
CPProperty.Value = NomeServidor
Set CPProperty = DBTable.ConnectionProperties("User ID")
CPProperty.Value = IIf(Usuario_banco = "", "Procam", Usuario_banco)
Set CPProperty = DBTable.ConnectionProperties("Password")
CPProperty.Value = IIf(Senha_banco = "", "PRO0902loc$?", Senha_banco)
Set CPProperty = DBTable.ConnectionProperties("Initial Catalog")
CPProperty.Value = Nome_banco
DBTable.Location = "authors2"

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcImprimirRelOrdenado(FormulaRel As String, FormulaRelSubReport As String)
On Error GoTo tratar_erro

ProcVerifRelPersonalizado

If PermitidoRel = False Then LocalrelNovo = Localrel Else LocalrelNovo = LocalRelPersonalizado
Set Report = crAPP.OpenReport(LocalrelNovo & "\" & NomeRel)
'Login SQL
Contador = Report.Database.Tables.Count
Do While Contador > 0
    Set DBTable = Report.Database.Tables(Contador)
    ProcLogonBDSQL
    Contador = Contador - 1
Loop
ProcVerifSubReport FormulaRelSubReport

frmimprimir.CrystalActiveXReportViewer1.ReportSource = Report
Report.FormulaSyntax = crCrystalSyntaxFormula
Report.RecordSelectionFormula = FormulaRel

If TabelaRel <> 0 Then
    Set CRXDatabaseField = Report.Database.Tables.Item(TabelaRel).Fields.Item(CampoRel)
    Report.RecordSortFields.Add CRXDatabaseField, OrdenarRel
End If
If TabelaRel1 <> 0 Then
    Set CRXDatabaseField = Report.Database.Tables.Item(TabelaRel1).Fields.Item(CampoRel1)
    Report.RecordSortFields.Add CRXDatabaseField, OrdenarRel1
End If
If TabelaRel2 <> 0 Then
    Set CRXDatabaseField = Report.Database.Tables.Item(TabelaRel2).Fields.Item(CampoRel2)
    Report.RecordSortFields.Add CRXDatabaseField, OrdenarRel2
End If

frmimprimir.CrystalActiveXReportViewer1.ViewReport
frmimprimir.Show 1
2:
    Set Report = Nothing
    Set crAPP = Nothing

Exit Sub
tratar_erro:
    If Err.Number = "-2147206461" Then
        USMsgBox ("N„o foi encontrado o relatÛrio " & NomeRel & " na pasta " & LocalrelNovo), vbExclamation, "CAPRIND v5.0"
        GoTo 2
    End If
    If Err.Number = "-2147483638" Then
        USMsgBox ("N„o foi possÌvel visualizar o relatÛrio, favor reiniciar o sistema."), vbExclamation, "CAPRIND v5.0"
        GoTo 2
    End If
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcImprimirRelGrafico(FormulaRel As String, FormulaRelSubReport As String)
On Error GoTo tratar_erro

ProcVerifRelPersonalizado

If PermitidoRel = False Then LocalrelNovo = Localrel Else LocalrelNovo = LocalRelPersonalizado
Set Report = crAPP.OpenReport(LocalrelNovo & "\" & NomeRel)
'Login SQL
Contador = Report.Database.Tables.Count
Do While Contador > 0
    Set DBTable = Report.Database.Tables(Contador)
    ProcLogonBDSQL
    Contador = Contador - 1
Loop
ProcVerifSubReport FormulaRelSubReport

frmimprimir.CrystalActiveXReportViewer1.ReportSource = Report
Report.FormulaSyntax = crCrystalSyntaxFormula
Report.RecordSelectionFormula = FormulaRel
frmimprimir.CrystalActiveXReportViewer1.ViewReport
frmimprimir.Show 1
2:
    Set Report = Nothing
    Set crAPP = Nothing

Exit Sub
tratar_erro:
    If Err.Number = "-2147206461" Then
        USMsgBox ("N„o foi encontrado o relatÛrio " & NomeRel & " na pasta " & LocalrelNovo), vbExclamation, "CAPRIND v5.0"
        GoTo 2
    End If
    If Err.Number = "-2147483638" Then
        USMsgBox ("N„o foi possÌvel visualizar o relatÛrio, favor reiniciar o sistema."), vbExclamation, "CAPRIND v5.0"
        GoTo 2
    End If
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcImprimirDireto(FormulaRel As String, FormulaRelSubReport As String)
On Error GoTo tratar_erro

ProcVerifRelPersonalizado
            
If PermitidoRel = False Then LocalrelNovo = Localrel Else LocalrelNovo = LocalRelPersonalizado
Set Report = crAPP.OpenReport(LocalrelNovo & "\" & NomeRel, crptToPrinter)
'Login SQL
Contador = Report.Database.Tables.Count
Do While Contador > 0
    Set DBTable = Report.Database.Tables(Contador)
    ProcLogonBDSQL
    Contador = Contador - 1
Loop
ProcVerifSubReport FormulaRelSubReport

Report.FormulaSyntax = crCrystalSyntaxFormula 'Configura a sintaxe da formula
Report.RecordSelectionFormula = FormulaRel 'Formula de seleÁ„o do relatÛrio
Report.PrintOut False 'Configura a seleÁ„o de impressora com false, enviando para impressora padr„o
2:
    Set Report = Nothing
    Set crAPP = Nothing

Exit Sub
tratar_erro:
    If Err.Number = "-2147206461" Then
        USMsgBox ("N„o foi encontrado o relatÛrio " & NomeRel & " na pasta " & LocalrelNovo), vbExclamation, "CAPRIND v5.0"
        GoTo 2
    End If
    If Err.Number = "-2147483638" Then
        USMsgBox ("N„o foi possÌvel visualizar o relatÛrio, favor reiniciar o sistema."), vbExclamation, "CAPRIND v5.0"
        GoTo 2
    End If
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcGerarPDF(caminho As String, FormulaRel As String, FormulaRelSubReport As String)
On Error GoTo tratar_erro

ProcVerifRelPersonalizado

If PermitidoRel = False Then LocalrelNovo = Localrel Else LocalrelNovo = LocalRelPersonalizado
Set Report = crAPP.OpenReport(LocalrelNovo & "\" & NomeRel, 1)
'Login SQL
Contador = Report.Database.Tables.Count
Do While Contador > 0
    Set DBTable = Report.Database.Tables(Contador)
    ProcLogonBDSQL
    Contador = Contador - 1
Loop
ProcVerifSubReport FormulaRelSubReport

Report.FormulaSyntax = crCrystalSyntaxFormula
Report.RecordSelectionFormula = FormulaRel
Report.DiscardSavedData

Set crxExport = Report.ExportOptions
crxExport.DiskFileName = caminho
crxExport.DestinationType = crEDTDiskFile
crxExport.PDFExportAllPages = True
crxExport.FormatType = crEFTPortableDocFormat
Report.Export False
2:
    Set Report = Nothing
    Set crAPP = Nothing

Exit Sub
tratar_erro:
    If Err.Number = "-2147206461" Then
        USMsgBox ("N„o foi encontrado o relatÛrio " & NomeRel & " na pasta " & LocalrelNovo), vbExclamation, "CAPRIND v5.0"
        GoTo 2
    End If
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCalculaBC(ID_empresa As Integer, CFOP As String, VlrMPA As Double, PV As Double, VlrIPI As Double, SomarIPI As String, SomarIPIST As Boolean, TemReducaoBC As Boolean, NaoArredondar As Boolean, CSTICMS As String, Aplicacao As String, IDforn As Long, NomeForn As String)
On Error GoTo tratar_erro

If Reducao = True And TemReducaoBC = True And (Right(CSTICMS, 2) = "20" Or Right(CSTICMS, 2) = "70" Or Right(CSTICMS, 3) = "900") Then

    'Calcula base de calculo do icms
    If SomarIPI = "SIM" Then PV = PV + VlrIPI
    If IntICMS > 0 Or Suframa = True And Desconto_Suframa = True Then
        If NaoArredondar = True Then ' n„o arredondar valor
            BC = (PV * CT) / 100
            BC = PV - BC
            If BC <> Int(BC) Then
                NumeroInteiro() = Split(BC, ",")
                BC = NumeroInteiro(0) & "," & Left(NumeroInteiro(1), 2)
            End If
            BC = Format(BC, "###,##0.00")
        Else ' arredondar valor
            BC = Format((PV * CT) / 100, "###,##0.00")
            BC = PV - BC
        End If
    End If
        
    If SomarIPIST = True Then
        BCST = Format(((PV + VlrIPI) * CT) / 100, "###,##0.00")
        BCST = (PV + VlrIPI) - BCST
    Else
        BCST = Format((PV * CT) / 100, "###,##0.00")
        BCST = PV - BCST
    End If
Else
    If IntICMS > 0 Or Suframa = True And Desconto_Suframa = True Then
        If SomarIPI = "SIM" Then BC = PV + VlrIPI Else BC = PV
    End If
    If SomarIPIST = True Then BCST = PV + VlrIPI Else BCST = PV
    
    'Verifica se È industrializaÁ„o e calcula a base de acordo com a aliquota de ICMS da empresa
    If MPA = True And VlrMPA <> 0 Then
        Set TBAfericao = CreateObject("adodb.recordset")
        If Aplicacao = "P" Then
            TBAfericao.Open "Select ICMS_ind from Impostos where ID_empresa = " & ID_empresa & " and ICMS_ind is not null and ICMS_ind <> N''", Conexao, adOpenKeyset, adLockOptimistic
        Else
            TBAfericao.Open "Select ICMS_ind from Compras_fornecedores where IDcliente = " & IDforn & " and Nome_Razao = '" & NomeForn & "' and ICMS_ind is not null and ICMS_ind <> N''", Conexao, adOpenKeyset, adLockOptimistic
        End If
        If TBAfericao.EOF = False Then
            'Calcula ICMS sem arredondar ou arredondando
            If NaoArredondar = True Then
                BC = (PV * TBAfericao!ICMS_ind) / 100
                If BC <> Int(BC) Then
                    NumeroInteiro() = Split(BC, ",")
                    BC = Format(NumeroInteiro(0) & "," & Left(NumeroInteiro(1), 2), "###,##0.00")
                End If
            Else
                BC = Format((PV * TBAfericao!ICMS_ind) / 100, "###,##0.00")
            End If
            BCST = BC
        End If
        TBAfericao.Close
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Function Funconsistir_CgcCpf(Vl_CgcCpf As String)
On Error GoTo tratar_erro

'Esta Rotina Devolver· True se o Cgc/Cpf Informado For valido ou False Se o Cgc/Cpf N„o For Correto
Dim VA_CgcCpf           As String
Dim VA_Digito           As String
Static Numero(15)       As Integer
Dim VA_Resto            As Integer
Dim VA_Resultado        As Integer
Dim VA_SomaDigito10     As Integer
Dim VA_resto1           As Integer

Funconsistir_CgcCpf = False
VA_CgcCpf = Format(ReturnNumbersOnly(Vl_CgcCpf), "@@@@@@@@@@@@@@")
VA_Digito = Mid(VA_CgcCpf, 13, 2)
Numero(1) = Val(Mid(VA_CgcCpf, 1, 1))
Numero(2) = Val(Mid(VA_CgcCpf, 2, 1))
Numero(3) = Val(Mid(VA_CgcCpf, 3, 1))
Numero(4) = Val(Mid(VA_CgcCpf, 4, 1))
Numero(5) = Val(Mid(VA_CgcCpf, 5, 1))
Numero(6) = Val(Mid(VA_CgcCpf, 6, 1))
Numero(7) = Val(Mid(VA_CgcCpf, 7, 1))
Numero(8) = Val(Mid(VA_CgcCpf, 8, 1))
Numero(9) = Val(Mid(VA_CgcCpf, 9, 1))
Numero(10) = Val(Mid(VA_CgcCpf, 10, 1))
Numero(11) = Val(Mid(VA_CgcCpf, 11, 1))
Numero(12) = Val(Mid(VA_CgcCpf, 12, 1))
Numero(13) = Val(Mid(VA_CgcCpf, 13, 1))
Numero(14) = Val(Mid(VA_CgcCpf, 14, 1))

If Len(Trim(VA_CgcCpf)) > 11 Then  ' Cgc
    Formato = Format(numeros, "@@.@@@.@@@/@@@@-@@")
    VA_Resultado = (Numero(1) * 5) + (Numero(2) * 4) _
            + (Numero(3) * 3) + (Numero(4) * 2) _
            + (Numero(5) * 9) + (Numero(6) * 8) + _
            (Numero(7) * 7) + (Numero(8) * 6) + _
            (Numero(9) * 5) + (Numero(10) * 4) + _
            (Numero(11) * 3) + (Numero(12) * 2)
    ' Atribui para resto o resto da divis„o
    ' de VA_resultado dividido por 11
    VA_Resto = VA_Resultado Mod 11
    If VA_Resto < 2 Then
        VA_resto1 = 0
    Else
        VA_resto1 = 11 - VA_Resto
    End If
    If VA_resto1 <> Numero(13) Then
        Exit Function
    End If
    VA_Resultado = (Numero(1) * 6) + _
            (Numero(2) * 5) + (Numero(3) * 4) + _
            (Numero(4) * 3) + (Numero(5) * 2) + _
            (Numero(6) * 9) + (Numero(7) * 8) + _
            (Numero(8) * 7) + (Numero(9) * 6) + _
            (Numero(10) * 5) + (Numero(11) * 4) + _
            (Numero(12) * 3) + (Numero(13) * 2)
    ' Atribui para resto o resto da divis„o
    ' de VA_resultado dividido por 11
    VA_Resto = VA_Resultado Mod 11
    If VA_Resto < 2 Then
        VA_resto1 = 0
    Else
        VA_resto1 = 11 - VA_Resto
    End If
    If VA_resto1 <> Numero(14) Then
        Exit Function
    End If
Else  ' Cpf
    Formato = Format(numeros, "@@@.@@@.@@@ - @@")
    VA_Resultado = (Numero(4) * 1) + (Numero(5) * 2) _
            + (Numero(6) * 3) + (Numero(7) * 4) _
            + (Numero(8) * 5) + (Numero(9) * 6) _
            + (Numero(10) * 7) + (Numero(11) * 8) + (Numero(12) * 9)
    VA_Resto = VA_Resultado Mod 11

    If VA_Resto > 9 Then
        VA_resto1 = VA_Resto - 10
    Else
        VA_resto1 = VA_Resto
    End If
    'frmCritica_CPCCGC.LblC1.Caption = VA_resto1
    If VA_resto1 <> Numero(13) Then
        GoTo Sairr
    End If
    VA_Resultado = (Numero(5) * 1) _
            + (Numero(6) * 2) + (Numero(7) * 3) _
            + (Numero(8) * 4) + (Numero(9) * 5) + _
            (Numero(10) * 6) + (Numero(11) * 7) + _
            (Numero(12) * 8) + (VA_resto1 * 9)
    VA_Resto = VA_Resultado Mod 11
    If VA_Resto > 9 Then
        VA_resto1 = VA_Resto - 10
    Else
        VA_resto1 = VA_Resto
    End If
    'frmCritica_CPCCGC.LblC2.Caption = VA_resto1
    If VA_resto1 <> Numero(14) Then
        Exit Function
    End If
End If

Funconsistir_CgcCpf = True
Exit Function
Sairr:
VA_Resultado = (Numero(5) * 1) _
        + (Numero(6) * 2) + (Numero(7) * 3) _
        + (Numero(8) * 4) + (Numero(9) * 5) + _
        (Numero(10) * 6) + (Numero(11) * 7) + _
        (Numero(12) * 8) + (VA_resto1 * 9)
VA_Resto = VA_Resultado Mod 11
If VA_Resto > 9 Then
    VA_resto1 = VA_Resto - 10
Else
    VA_resto1 = VA_Resto
End If
If VA_resto1 <> Numero(14) Then
    Exit Function
End If

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Sub ProcGravaEvento()
On Error GoTo tratar_erro

ProcINSERTINTO "Mascara", "Modulo, Usuario, Operacao, Data, Hora, Documento, Documento1, ID_documento", "'" & Modulo & "', '" & pubUsuario & "', '" & Evento & "', '" & Date & "', '" & Time & "', '" & Documento & "', '" & Documento1 & "' , " & ID_documento & ""

Prosseguir:
    Set TBLogon = CreateObject("adodb.recordset")
    TBLogon.Open "Select Hora_ultimo_evento from Logon where Usuario = '" & pubUsuario & "' and Data = '" & Date & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBLogon.EOF = False Then
        TBLogon!Hora_ultimo_evento = Time
        TBLogon.Update
    End If
    TBLogon.Close

Exit Sub
tratar_erro:
    If Err.Number = "13" Or Err.Number = "3022" Then GoTo Prosseguir
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Function FunLogonIn(VerifConexao As Boolean)
On Error GoTo tratar_erro

Set TBLogon = CreateObject("adodb.recordset")
If VerifConexao = True And pubUsuario <> "PROCAM" Then
    TBLogon.Open "Select * from Logon where usuario = '" & pubUsuario & "' and Tipo = 'C'", Conexao, adOpenKeyset, adLockOptimistic
    If TBLogon.EOF = False Then
        If USMsgBox("O usu·rio " & pubUsuario & " j· est· conectado, deseja desconectar a conex„o antiga antes de iniciar uma nova conex„o?", vbYesNo, "CAPRIND v5.0") = vbYes Then
            Conexao.Execute "DELETE from Logon where IDlogon = " & TBLogon!IDlogon
            
            'Conta 11 segundos para desconectar a outra conex„o
            Dataini = Format(Now, "hh:mm:ss")
            Dataini = Dataini + "00:00:11"
            Do While Format(Now, "hh:mm:ss") < Dataini
            
            Loop
        End If
        Set TBLogon = CreateObject("adodb.recordset")
        TBLogon.Open "Select * from Logon", Conexao, adOpenKeyset, adLockOptimistic
        TBLogon.AddNew
    Else
        TBLogon.AddNew
    End If
Else
    TBLogon.Open "Select * from Logon", Conexao, adOpenKeyset, adLockOptimistic
    TBLogon.AddNew
End If
TBLogon!Usuario = pubUsuario
TBLogon!Data = Date
TBLogon!Hora = Time
TBLogon!Hora_ultimo_evento = Time
TBLogon!Tipo = "C"
TBLogon.Update
IDlogon = TBLogon!IDlogon
TBLogon.Close

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Sub ProcLogonOut()
On Error GoTo tratar_erro

With frmMDI
    .Timer_logon.Enabled = False
    .Timer_logoff_automatico.Enabled = False
End With
InfLicenca = False
ProcLogonOut1 IDlogon, pubUsuario, "C"
                
Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLogonOut1(IDlogon As Long, Usuario As String, TextoFiltro As String)
On Error GoTo tratar_erro

Conexao.Execute "DELETE from Logon where IDLogon = " & IDlogon
Conexao.Execute "DELETE from Producao_Relatorios where Responsavel = '" & Usuario & "'"
Conexao.Execute "DELETE from Producao_Relatorios_Total where Responsavel = '" & Usuario & "'"
Conexao.Execute "DELETE from Estoque_relatorios where Responsavel = '" & Usuario & "'"
Conexao.Execute "DELETE from Troca_titulo_relatorio where Responsavel = '" & Usuario & "'"
Conexao.Execute "DELETE from Plano_de_contas_totalizacao where Responsavel = '" & Usuario & "'"
Conexao.Execute "DELETE from Compras_Recebimento_Relatorios where responsavel = '" & Usuario & "'"
Conexao.Execute "DELETE from Etiqueta where responsavel = '" & Usuario & "'"

'Efetua logoff do usu·rio no site
If TemInternet = True And ErroDriverMYSQL = False Then
    Set TBAfericao = CreateObject("adodb.recordset")
    TBAfericao.Open "Select * from Empresa", Conexao, adOpenKeyset, adLockOptimistic
    If TBAfericao.EOF = False Then
        FunAbreBDSite
        If ConexaoMySql.State = 1 Then ConexaoMySql.Execute "Update usuarios Set Logado = 'N√O' where CNPJ = '" & TBAfericao!CNPJ & "' and Usuario = '" & Usuario & "'"
    End If
    TBAfericao.Close
End If
                
Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLogonOutSemUtilizacao()
On Error GoTo tratar_erro

Set TBLogon = CreateObject("adodb.recordset")
TBLogon.Open "Select Usuario from Logon where Data < '" & Date & "' and Tipo = 'C'", Conexao, adOpenKeyset, adLockOptimistic
If TBLogon.EOF = False Then
    Do While TBLogon.EOF = False
        Conexao.Execute "DELETE from Logon WHERE usuario = '" & TBLogon!Usuario & "' and Tipo = 'C'"
        Conexao.Execute "DELETE from Producao_Relatorios where Responsavel = '" & TBLogon!Usuario & "'"
        Conexao.Execute "DELETE from Producao_Relatorios_Total where Responsavel = '" & TBLogon!Usuario & "'"
        Conexao.Execute "DELETE from Estoque_relatorios where Responsavel = '" & TBLogon!Usuario & "'"
        Conexao.Execute "DELETE from Troca_titulo_relatorio where Responsavel = '" & TBLogon!Usuario & "'"
        Conexao.Execute "DELETE from Plano_de_contas_totalizacao where Responsavel = '" & TBLogon!Usuario & "'"
        Conexao.Execute "DELETE from Compras_Recebimento_Relatorios where responsavel = '" & TBLogon!Usuario & "'"
        
        'Efetua logoff do usu·rio no site
        If TemInternet = True And ErroDriverMYSQL = False Then
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select IDLogon from Logon where usuario = '" & TBLogon!Usuario & "' and Data = '" & Date & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = True Then
                Set TBFIltro = CreateObject("adodb.recordset")
                TBFIltro.Open "Select * from Empresa", Conexao, adOpenKeyset, adLockOptimistic
                If TBFIltro.EOF = False Then
                    FunAbreBDSite
                    If ConexaoMySql.State = 1 Then ConexaoMySql.Execute "Update usuarios Set Logado = 'N√O', Logado_Gerprod = 'N√O' where CNPJ = '" & TBFIltro!CNPJ & "' and Usuario = '" & TBLogon!Usuario & "'"
                End If
                TBFIltro.Close
            End If
            TBAbrir.Close
        End If
        TBLogon.MoveNext
    Loop
End If
TBLogon.Close
                
Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Function RetornarNumeros(x As String) As String
On Error GoTo tratar_erro
Dim Temp As String
Dim j As Integer

Temp = ""
For j = 1 To Len(x)
    If Mid(x, j, 1) = "0" Or _
        Mid(x, j, 1) = "1" Or _
        Mid(x, j, 1) = "2" Or _
        Mid(x, j, 1) = "3" Or _
        Mid(x, j, 1) = "4" Or _
        Mid(x, j, 1) = "5" Or _
        Mid(x, j, 1) = "6" Or _
        Mid(x, j, 1) = "7" Or _
        Mid(x, j, 1) = "8" Or _
        Mid(x, j, 1) = "9" Then
        Temp = Temp + Mid(x, j, 1)
    End If
Next
RetornarNumeros = Temp

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Sub PulaCampo(KA As Integer)
On Error GoTo tratar_erro

If KA = 13 Then SendKeys "{TAB}"

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Function FunAbreBD() As Boolean
On Error GoTo tratar_erro

If NomeServidor = "" Or Nome_banco = "" Then Exit Function

Abrir = True
FunAbreBD = True

NomeCampo = "Caprind"
Set Conexao = New ADODB.Connection
With Conexao
    .Provider = "SQLOLEDB"
    .Properties("Data Source").Value = NomeServidor
    .Properties("Initial catalog").Value = Nome_banco
    .Properties("User ID").Value = IIf(Usuario_banco = "", "Procam", Usuario_banco)
    .Properties("Password").Value = IIf(Senha_banco = "", "PRO0902loc$?", Senha_banco)
    .Properties("Persist Security Info") = "False"
    .Open
End With

Exit Function
tratar_erro:
    If Err.Number = "-2147467259" Then
        Abrir = False
        FunAbreBD = False
        Exit Function
    End If
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Public Function FunAbreBDSite() As Boolean
On Error GoTo tratar_erro

'MYSQL Site
Permitido = True
Atualizando = False
ErroDriverMYSQL = False
Contador = Contador + 1
Set ConexaoMySql = New ADODB.Connection
With ConexaoMySql
    'Porta 3306
    .ConnectionTimeout = 120
    .CommandTimeout = 400
    .CursorLocation = adUseClient

Conectar:
    If Permitido = False Then
        .Open "DRIVER={MySQL ODBC 3.51 Driver};" & "user=caprind11" & ";password=cap0902loc" & ";database=caprind11" & ";server=mysql02.caprind1.hospedagemdesites.ws" & ";option=20499"
    Else
        .Open "DRIVER={MySQL ODBC 5.1 Driver};" & "user=caprind11" & ";password=cap0902loc" & ";database=caprind11" & ";server=mysql02.caprind1.hospedagemdesites.ws" & ";option=20499"
    End If
End With

Exit Function
tratar_erro:
    If Err.Number = "-2147467259" Then
        
        If FunVerifFormAberto(frmabertura) = True And Contador = 1 Then
            caminho = "C:\Program Files (x86)\MySQL\Connector ODBC 5.1"
            Set GerArqPastas = CreateObject("Scripting.FileSystemObject")
            If GerArqPastas.FolderExists(caminho) = False Then
                caminho = "C:\Program Files\MySQL\Connector ODBC 5.1"
                Set GerArqPastas = CreateObject("Scripting.FileSystemObject")
                If GerArqPastas.FolderExists(caminho) = False Then
                    Call USMsgBox("… obrigatÛrio atualizar o driver MySQL antes de logar." & vbCrLf & "IMPORTANTE: clique no bot„o Next> da instalaÁ„o do driver para prosseguir." & vbCrLf & "O sistema ser· encerrado apÛs a atualizaÁ„o.", vbInformation, "CAPRIND v5.0", "Driver MySQL desatualizado")
                    Atualizando = True
                    With frmabertura
                        .kftp.DisableRESTCommand
                        If FunConectaKFTP(.kftp, "public_html/phocadownload/userupload/Atualizacao", True) = False Then
                            Permitido = False
                            Atualizando = False
                            GoTo Conectar
                        End If
                        If FunDownloadKFTP(.kftp, "mysql-connector-odbc-5.1.13-win32.msi", App.Path & "\mysql-connector-odbc-5.1.13-win32.msi") = False Then
                            Permitido = False
                            Atualizando = False
                            GoTo Conectar
                        End If
                    End With
                    ProcAbrirArquivo (App.Path & "\mysql-connector-odbc-5.1.13-win32.msi")
                    End
                End If
            End If
        End If
        ErroDriverMYSQL = True
        Exit Function
    End If
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Public Sub ProcSalvarPedidoWEB(IDpedido)
On Error GoTo tratar_erro

FunAbreBDSite

'=================================================================
' Localiza dados do pedido de compras localmente
'=================================================================
Set TBCompras_Pedido = CreateObject("adodb.recordset")
StrSql = "Select CP.* , CC.condicoes from Compras_pedido CP Inner Join Compras_comercial CC ON CC.IdPedido = CP.IDPedido where CP.IDPedido = " & IDpedido & ""
'Debug.print StrSql

TBCompras_Pedido.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
If TBCompras_Pedido.EOF = False Then
    If ConexaoMySql.State = 1 Then
        Set TBMySQL = New ADODB.Recordset
'=================================================================
' Salvar pedido na WEB
'=================================================================
        StrSql = "Select * From Compras_Pedido where Pedido = '" & TBCompras_Pedido!Pedido & "' and CNPJ_Empresa = '" & CNPJ_Empresa & "'"
        'Debug.print StrSql
        
       TBMySQL.Open StrSql, ConexaoMySql, adOpenKeyset, adLockOptimistic, adCmdText
            If TBMySQL.EOF = True Then
                TBMySQL.AddNew
                TBMySQL.Fields!Pedido = TBCompras_Pedido!Pedido
                TBMySQL.Fields!Fornecedor = TBCompras_Pedido!Fornecedor
                TBMySQL.Fields!Status_pedido = TBCompras_Pedido!Status_pedido
                TBMySQL.Fields!Data = TBCompras_Pedido!Data
                TBMySQL.Fields!Responsavel = TBCompras_Pedido!Responsavel
                TBMySQL.Fields!dbl_valor_total = TBCompras_Pedido!dbl_valor_total
                TBMySQL.Fields!ID_empresa = TBCompras_Pedido!ID_empresa
                TBMySQL.Fields!CNPJ_Empresa = CNPJ_Empresa
                TBMySQL.Update
                IDpedido = TBMySQL.Fields!IDpedido

        
'========================================================
'Salvar dados comerciais na WEB
'========================================================
        Set TBMySQL = New ADODB.Recordset
        StrSql = "Select * From Compras_comercial where IDPedido = '" & IDpedido & "'"
        'Debug.print StrSql
        
        TBMySQL.Open StrSql, ConexaoMySql, adOpenKeyset, adLockOptimistic, adCmdText
            If TBMySQL.EOF = True Then
                TBMySQL.AddNew

                TBMySQL.Fields!IDpedido = IDpedido
                TBMySQL.Fields!condicoes = TBCompras_Pedido!condicoes
                TBMySQL.Update
                'IDpedido = .Fields!IDpedido
            End If


 
'=========================================================
' Salvar itens do pedido na Web
'=========================================================
Set TBCompras_Pedido_Lista = CreateObject("adodb.recordset")
TBCompras_Pedido_Lista.Open "Select * from Compras_pedido_lista where IDPedido = " & TBCompras_Pedido!IDpedido & "", Conexao, adOpenKeyset, adLockOptimistic
   If TBCompras_Pedido.EOF = False Then
   Do While TBCompras_Pedido_Lista.EOF = False
       If ConexaoMySql.State = 1 Then
           Set TBMySQL = New ADODB.Recordset
           StrSql = "Select * From Compras_pedido_lista where IDPedido = '" & IDpedido & "' and Desenho = '" & TBCompras_Pedido_Lista!Desenho & "'"
           'Debug.print StrSql
           
           TBMySQL.Open StrSql, ConexaoMySql, adOpenKeyset, adLockOptimistic, adCmdText
   
               If TBMySQL.EOF = True Then
                   TBMySQL.AddNew
                   TBMySQL.Fields!IDpedido = IDpedido
                   TBMySQL.Fields!Desenho = TBCompras_Pedido_Lista!Desenho
                   TBMySQL.Fields!Descricao = TBCompras_Pedido_Lista!Descricao
                   TBMySQL.Fields!Quant_Comp = TBCompras_Pedido_Lista!Quant_Comp
                   TBMySQL.Fields!preco_unitario = TBCompras_Pedido_Lista!preco_unitario
                   TBMySQL.Fields!preco_total = TBCompras_Pedido_Lista!preco_total
                   TBMySQL.Fields!Status_Item = TBCompras_Pedido_Lista!Status_Item
                   TBMySQL.Update
               End If
       End If
   TBCompras_Pedido_Lista.MoveNext
   Loop
   TBCompras_Pedido_Lista.Close
   End If
End If
End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Public Sub ProCarregaTema()
On Error GoTo tratar_erro

'With frmMDI.SkinFramework1
'
'.LoadSkin TemaCaprind
'
'End With

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcBuscarPedidoWEB(IDpedido)
On Error GoTo tratar_erro

FunAbreBDSite

'=================================================================
' Localiza dados do pedido de compras
'=================================================================
    If ConexaoMySql.State = 1 Then
        Set TBMySQL = New ADODB.Recordset
        StrSql = "Select * From Compras_Pedido where Pedido = '" & TBCompras_Pedido!Pedido & "' and CNPJ_Empresa = '" & CNPJ_Empresa & "'"
        'Debug.print StrSql
        
        TBMySQL.Open StrSql, ConexaoMySql, adOpenKeyset, adLockOptimistic, adCmdText
        With TBMySQL
            If .EOF = False Then
             Status_pedido = .Fields!Status_pedido
             Data_aprovado = .Fields!Data_aprovado
             Resp_aprovado = .Fields!Resp_aprovado
             If Status_pedido = "APROVADO" And TBCompras_Pedido!Status_pedido = "AGUARDANDO APROVA«√O" Then
               StrSql = "Update Compras_pedido set Status_pedido = '" & Status_pedido & "' , Resp_aprovado = '" & Resp_aprovado & "' ,  Data_aprovado = '" & Format(Data_aprovado, "dd/mm/yy") & "' where IDPedido = '" & TBCompras_Pedido!IDpedido & "'"
               'Debug.print StrSql
               
               Conexao.Execute StrSql
  
               frmCompras_Pedido.txtStatus = IIf(Status_pedido <> "", Status_pedido, "")
               frmCompras_Pedido.txtData_aprovacao = IIf(Data_aprovado <> "", Format(Data_aprovado, "dd-mm-yyyy"), "")
               frmCompras_Pedido.txtResponsavel_aprovacao = IIf(Resp_aprovado <> "", Resp_aprovado, "")
               frmCompras_Pedido.listapedido.SelectedItem.ListSubItems.Item(5).Text = Status_pedido
               frmCompras_Pedido.listapedido.SelectedItem.ListSubItems.Item(7).Text = IIf(Status_pedido = "APROVADO", "Sim", "N„o")
             End If
            End If
        End With
    End If

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Function FunAbreBDWEB() As Boolean
On Error GoTo tratar_erro

'MYSQL Site
Permitido = True
Atualizando = False
ErroDriverMYSQL = False
Contador = Contador + 1
Set Conexao = New ADODB.Connection
With Conexao
    'Porta 3306
    .ConnectionTimeout = 60
    .CommandTimeout = 400
    .CursorLocation = adUseClient

Conectar:
    If Permitido = False Then
        .Open "DRIVER={MySQL ODBC 3.51 Driver};" & "user=caprind_ple" & ";password=C@p0902loc" & ";database=caprind_ple" & ";server=caprind_ple.mysql.dbaas.com.br" & ";option=20499"
    Else
        .Open "DRIVER={MySQL ODBC 5.1 Driver};" & "user=caprind_ple" & ";password=C@p0902loc" & ";database=caprind_ple" & ";server=caprind_ple.mysql.dbaas.com.br" & ";option=20499"
      '  .Open "DRIVER={MySQL ODBC 5.1 Driver};" & "user=root" & ";password=" & ";database=caprind_ple" & ";server=127.0.0.1" & ";option=20499"
        FunAbreBDWEB = True
    End If
End With

Exit Function
tratar_erro:
    If Err.Number = "-2147467259" Then
        If FunVerifFormAberto(frmabertura) = True And Contador = 1 Then
            caminho = "C:\Program Files (x86)\MySQL\Connector ODBC 5.1"
            Set GerArqPastas = CreateObject("Scripting.FileSystemObject")
            If GerArqPastas.FolderExists(caminho) = False Then
                caminho = "C:\Program Files\MySQL\Connector ODBC 5.1"
                Set GerArqPastas = CreateObject("Scripting.FileSystemObject")
                If GerArqPastas.FolderExists(caminho) = False Then
                    Call USMsgBox("… obrigatÛrio atualizar o driver MySQL antes de logar." & vbCrLf & "IMPORTANTE: clique no bot„o Next> da instalaÁ„o do driver para prosseguir." & vbCrLf & "O sistema ser· encerrado apÛs a atualizaÁ„o.", vbInformation, "CAPRIND v5.0", "Driver MySQL desatualizado")
                    Atualizando = True
                    With frmabertura
                        .kftp.DisableRESTCommand
                        If FunConectaKFTP(.kftp, "public_html/phocadownload/userupload/Atualizacao", True) = False Then
                            Permitido = False
                            Atualizando = False
                            GoTo Conectar
                        End If
                        If FunDownloadKFTP(.kftp, "mysql-connector-odbc-5.1.13-win32.msi", App.Path & "\mysql-connector-odbc-5.1.13-win32.msi") = False Then
                            Permitido = False
                            Atualizando = False
                            GoTo Conectar
                        End If
                    End With
                    ProcAbrirArquivo (App.Path & "\mysql-connector-odbc-5.1.13-win32.msi")
                    End
                End If
            End If
        End If
        ErroDriverMYSQL = True
        Exit Function
    End If
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function


Function FunFechaBD()
On Error GoTo tratar_erro

Conexao.Close
'Conexao_NFe.Close
'BD.Close

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Public Function FunFechaBDSite()
On Error GoTo tratar_erro

If ConexaoMySql.State = 1 Then ConexaoMySql.Close

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Function FunGotFocus(Objeto)
On Error GoTo tratar_erro

Objeto.SelStart = 0
Objeto.SelLength = Len(Objeto.Text)

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Public Function RemoverCaracter(valor As String) As String
On Error GoTo tratar_erro
Dim Remover As String, i As Byte, Temp As String
Contador = 0

Remover = "?*∞‡…∫z%[^!#$%&''()*+./:;<=>?@\^_`|{}~ÄÅÇÉÖÜáàâäãåé''""o--òôöõúûü°¢£§•¶ß®©"
Remover = Remover & "™ÆØ∞±≤≥¥µ∑∏π∫ºΩæø¿¡¬√ƒ≈∆«»… ÀÃÕŒœ–—“”‘’÷◊ÿŸ⁄€‹›ﬁﬂ‡·‚„‰ÂÊÁËÈÍÎÏÌÓÔÒÚÛÙıˆ˜¯˘˙˚¸˝˛ˇ%"
Temp = valor

Do While Contador < Len(valor)
Contador = Contador + 1
    Temp = Replace(Temp, Mid(Remover, Contador, 1), "")
Loop
RemoverCaracter = Temp

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Public Function RemoverCaracterELetras(valor As String) As String
On Error GoTo tratar_erro
Dim Remover As String, i As Byte, Temp As String
Contador = 0

Remover = "?*∞‡…∫z%[^!#$%&''()*+.:;<=>?@\^_`|{}~ÄÅÇÉÖÜáàâäãåé''""o--òôöõúûü°¢£§•¶ß®©"
Remover = Remover & "™ÆØ∞±≤≥¥µ∑∏π∫ºΩæø¿¡¬√ƒ≈∆«»… ÀÃÕŒœ–—“”‘’÷◊ÿŸ⁄€‹›ﬁﬂ‡·‚„‰ÂÊÁËÈÍÎÏÌÓÔÒÚÛÙıˆ˜¯˘˙˚¸˝˛ˇ%"
Remover = Remover & "ABCDEFGHIJKLMNOPWRSTUVXYWZabcdefghijklmnopqrstuvxywz"
Temp = valor
'Debug.print Remover

Do While Contador < Len(valor)
Contador = Contador + 1
    Temp = Replace(Temp, Mid(Remover, Contador, 1), "")
Loop
RemoverCaracterELetras = Temp

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Public Sub ProcCriarPastasDFE()
On Error GoTo tratar_erro


Diretorio = Left(Localrel, Len(Localrel) - 10)

   If DS.FileOrDirExists(Diretorio & "Nota fiscal\") = False Then
    MkDir Diretorio & "Nota fiscal\"
   End If
   
   If DS.FileOrDirExists(Diretorio & "Nota fiscal\Envio\") = False Then
    MkDir Diretorio & "Nota fiscal\Envio\"
   End If
   
   If DS.FileOrDirExists(Diretorio & "Nota fiscal\Retorno\") = False Then
    MkDir Diretorio & "Nota fiscal\Retorno\"
   End If
   
   If DS.FileOrDirExists(Diretorio & "Nota fiscal\Danfe-XML\") = False Then
    MkDir Diretorio & "Nota fiscal\Danfe-XML\"
   End If

Diretorio = ""

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub procVerificacodigoUF(UF As String)
On Error GoTo tratar_erro

Select Case UF
   Case "MA": CodUF = "21" '21 Maranh„o MA
   Case "Pi": CodUF = "22" '22 PiauÌ Pi
   Case "CE": CodUF = "23" '23 Cear· CE
   Case "RN": CodUF = "24" '24 Rio Grande do Norte RN
   Case "PB": CodUF = "25" '25 ParaÌba PB
   Case "PE": CodUF = "26" '26 Pernambuco PE
   Case "AL": CodUF = "27" '27 Alagoas AL
   Case "SE": CodUF = "28" '28 Sergipe SE
   Case "BA": CodUF = "29" '29 Bahia BA
   Case "MG": CodUF = "31" '31 Minas Gerais MG
   Case "ES": CodUF = "32" '32 EspÌrito Santo ES
   Case "RJ": CodUF = "33" '33 Rio de Janeiro RJ
   Case "SP": CodUF = "35" '35 S„o Paulo SP
   Case "PR": CodUF = "41" '41 Paran· PR
   Case "SC": CodUF = "42" '42 Santa Catarina SC
   Case "RS": CodUF = "43" '43 Rio Grande do Sul RS
   Case "MS": CodUF = "50" '50 Mato Grosso do Sul MS
   Case "MT": CodUF = "51" '51 Mato Grosso MT
   Case "GO": CodUF = "52" '52 Goi·s GO
   Case "DF": CodUF = "53" '53 Distrito Federal DF
End Select

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub Main()
On Error GoTo tratar_erro

'Verifica drawSuite 2022
'DownloadDrawSuite2022
'================================================================================


'Verifica local de execuÁ„o do exe
caminho = App.Path
Contador = Len(caminho)
Debug.Print Right(caminho, 27)
If Right(caminho, 27) <> "Projeto Caprind\Projeto VB6" And Right(caminho, 15) <> "Projeto Caprind" And Mid(caminho, 4, 21) <> "Arquivos de Programas" And Mid(caminho, 4, 21) <> "Arquivos de programas" And Mid(caminho, 4, 27) <> "Arquivos de Programas (x86)" And Mid(caminho, 4, 27) <> "Arquivos de programas (x86)" And Mid(caminho, 4, 13) <> "Program Files" And Mid(caminho, 4, 13) <> "Program files" And Mid(caminho, 4, 19) <> "Program Files (x86)" And Mid(caminho, 4, 19) <> "Program files (x86)" Then
    USMsgBox ("N„o È permitido abrir o Caprind deste caminho " & caminho & "."), vbCritical, "CAPRIND v5.0"
    End
End If

'Verifica resoluÁ„o do computador
xTwips = Screen.TwipsPerPixelX
yTwips = Screen.TwipsPerPixelY
xPixels = Screen.Width / xTwips
YPixels = Screen.Height / yTwips
If xPixels < 1024 And YPixels < 768 Then
    USMsgBox ("O Caprind requer resoluÁ„o mÌnima de 1024x768 pÌxels."), vbCritical, "CAPRIND v5.0"
    End
End If

ProcCarregaBancoDados
FormatoData = GetSetting("Procam", "CaprindSQL", "FormatoData", "dd/mm/yyyy")
FormatoHora = GetSetting("Procam", "CaprindSQL", "FormatoHora", "hh:mm:ss")
Simbolos = "ÿ±ºΩæ≤≥™∞°¢£§•¶ß®©≠ÆØ¥µ∑∏π∫øÊ˜¯◊"
pubLicenca = "M·ximo 10 liberaÁıes"
pubRegistrado = "DemonstraÁ„o"
InfLicenca = False
If Salvarrel = "" Then Salvarrel = False
'Se o banco de dados for localizado.
If Salvarrel = False Then
    With frmabertura
        .Timer1.Enabled = True
        .Timer2.Enabled = True
        .Show
    End With
End If

VerifDadosPadraoFamilia = True

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Function FunVerifAtualizacaoObrigatoria(FlexCell As Boolean, ControleLayout As Boolean) As Boolean
On Error GoTo tratar_erro

FunVerifAtualizacaoObrigatoria = False
If FlexCell = True Then
    If DS.IsWow64ProcessEx = True Then
        caminho = "C:\Windows\SysWOW64"
'        FamiliaAntiga = "ATCaprind_x64_v4.9.163.zip"
    Else
        caminho = "C:\Windows\System32"
'        FamiliaAntiga = "ATCaprind_x86_v4.9.163.zip"
    End If
    If DS.FileOrDirExists(caminho & "\FlexCell.ocx") = False Then
'        FunVerifAtualizacaoObrigatoria = True
'        Call usMsgbox("… necess·rio instalar um novo controle antes de abrir este mÛdulo." & vbCrLf & "IMPORTANTE: clique no bot„o Next> da instalaÁ„o do controle para prosseguir.", vbInformation, "CAPRIND v5.0", "InstalaÁ„o de novo controle")
'        Atualizacao_GNFe = False
'        Atualizacao_GMRE = False
'        Atualizacao_versao = True
'        Frm_atualizacao_sistema.Show 1
    End If
End If


Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Function

Sub ProcCarregaBancoDados()
On Error GoTo tratar_erro

NomeServidor = GetSetting("Procam", "CaprindSQL", "NomeServidor")
Localrel = GetSetting("Procam", "CaprindSQL", "Localrel")
Nome_banco = GetSetting("Procam", "CaprindSQL", "Nome_banco")
Usuario_banco = GetSetting("Procam", "CaprindSQL", "Usuario_banco")
Senha_banco = GetSetting("Procam", "CaprindSQL", "Senha_banco")
LocalAntigoCaprind = GetSetting("Procam", "CaprindSQL", "LocalAntigoCaprind")
LocalNovoCaprind = GetSetting("Procam", "CaprindSQL", "LocalNovoCaprind")
LocalAntigoGerprod = GetSetting("Procam", "CaprindSQL", "LocalAntigoGerprod")
LocalNovoGerprod = GetSetting("Procam", "CaprindSQL", "LocalNovoGerprod")

NomeServidor1 = GetSetting("Procam", "CaprindSQL", "NomeServidor1")
Localrel1 = GetSetting("Procam", "CaprindSQL", "Localrel1")
Nome_banco1 = GetSetting("Procam", "CaprindSQL", "Nome_banco1")
Usuario_banco1 = GetSetting("Procam", "CaprindSQL", "Usuario_banco1")
Senha_banco1 = GetSetting("Procam", "CaprindSQL", "Senha_banco1")
LocalAntigoCaprind1 = GetSetting("Procam", "CaprindSQL", "LocalAntigoCaprind1")
LocalNovoCaprind1 = GetSetting("Procam", "CaprindSQL", "LocalNovoCaprind1")
LocalAntigoGerprod1 = GetSetting("Procam", "CaprindSQL", "LocalAntigoGerprod1")
LocalNovoGerprod1 = GetSetting("Procam", "CaprindSQL", "LocalNovoGerprod1")

NomeServidor2 = GetSetting("Procam", "CaprindSQL", "NomeServidor2")
Localrel2 = GetSetting("Procam", "CaprindSQL", "Localrel2")
Nome_banco2 = GetSetting("Procam", "CaprindSQL", "Nome_banco2")
Usuario_banco2 = GetSetting("Procam", "CaprindSQL", "Usuario_banco2")
Senha_banco2 = GetSetting("Procam", "CaprindSQL", "Senha_banco2")
LocalAntigoCaprind2 = GetSetting("Procam", "CaprindSQL", "LocalAntigoCaprind2")
LocalNovoCaprind2 = GetSetting("Procam", "CaprindSQL", "LocalNovoCaprind2")
LocalAntigoGerprod2 = GetSetting("Procam", "CaprindSQL", "LocalAntigoGerprod2")
LocalNovoGerprod2 = GetSetting("Procam", "CaprindSQL", "LocalNovoGerprod2")

NomeServidor3 = GetSetting("Procam", "CaprindSQL", "NomeServidor3")
Localrel3 = GetSetting("Procam", "CaprindSQL", "Localrel3")
Nome_banco3 = GetSetting("Procam", "CaprindSQL", "Nome_banco3")
Usuario_banco3 = GetSetting("Procam", "CaprindSQL", "Usuario_banco3")
Senha_banco3 = GetSetting("Procam", "CaprindSQL", "Senha_banco3")
LocalAntigoCaprind3 = GetSetting("Procam", "CaprindSQL", "LocalAntigoCaprind3")
LocalNovoCaprind3 = GetSetting("Procam", "CaprindSQL", "LocalNovoCaprind3")
LocalAntigoGerprod3 = GetSetting("Procam", "CaprindSQL", "LocalAntigoGerprod3")
LocalNovoGerprod3 = GetSetting("Procam", "CaprindSQL", "LocalNovoGerprod3")

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Function FunTemAcesso(Formulario As String) As Boolean
On Error GoTo tratar_erro

If Left(Formulario, 32) = "ConfiguraÁ„o do sistema/Usu·rios" And (pubUsuario = "Admin" Or pubUsuario = "Cliente") Then
    Acessos = True
Else
    Set TBAcessos = CreateObject("adodb.recordset")
    TBAcessos.Open "Select * FROM Acessos WHERE IDUsuario = " & pubIDUsuario & " AND Acesso = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBAcessos.EOF = True Then
        Acessos = False
    Else
        Acessos = True
    End If
    TBAcessos.Close
End If
FunTemAcesso = Acessos

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Sub ProcLiberaAcessos(MostrarMsg As Boolean)
On Error GoTo tratar_erro

If Left(Formulario, 32) <> "ConfiguraÁ„o do sistema/Usu·rios" Then
    Set TBAcessos = CreateObject("adodb.recordset")
    TBAcessos.Open "Select Codigo from Empresa", Conexao, adOpenKeyset, adLockOptimistic
    If TBAcessos.EOF = True Then
        If MostrarMsg = True Then USMsgBox ("… necess·rio cadastrar a empresa antes de acessar este mÛdulo."), vbExclamation, "CAPRIND v5.0"
        Acessos = False
        Exit Sub
    End If
    TBAcessos.Close
End If
If Formulario = "Vendas/Proposta comercial" Or Formulario = "Vendas/Pedido interno" Or Left(Formulario, 11) = "Faturamento" Or Formulario = "Estoque/Ordem de faturamento" Or Formulario = "Estoque/Nota fiscal" Then
    Set TBAcessos = CreateObject("adodb.recordset")
    TBAcessos.Open "Select CODIGO, Simples, presumido, Real, simples1 from Empresa", Conexao, adOpenKeyset, adLockOptimistic
    If TBAcessos.EOF = False Then
        Do While TBAcessos.EOF = False
            'If TBAcessos!Simples = False And TBAcessos!Presumido = False And TBAcessos!Real = False And TBAcessos!Simples1 = False Then
            If TBAcessos!Simples = False And TBAcessos!Presumido = False And TBAcessos!Real = False And TBAcessos!Simples1 = False Then
                If MostrarMsg = True Then USMsgBox ("… necess·rio cadastrar o regime tribut·rio da empresa antes de acessar este mÛdulo."), vbExclamation, "CAPRIND v5.0"
                Acessos = False
                Exit Sub
            End If
            
            'Verifica se foi cadastrado a tabela do simples nacional
            If TBAcessos!Simples = True Then
                Set TBTempo = CreateObject("adodb.recordset")
                TBTempo.Open "Select ID from Impostos_TabelaDAS where ID_empresa = " & TBAcessos!CODIGO, Conexao, adOpenKeyset, adLockOptimistic
                If TBTempo.EOF = True Then
                    If MostrarMsg = True Then USMsgBox ("… necess·rio cadastrar a tabela do simples nacional antes de acessar este mÛdulo."), vbExclamation, "CAPRIND v5.0"
                    Acessos = False
                    Exit Sub
                End If
                TBTempo.Close
            End If
            TBAcessos.MoveNext
        Loop
    End If
    TBAcessos.Close
End If

If FunTemAcesso(Formulario) = False Then
    If MostrarMsg = True Then USMsgBox ("Usu·rio " & pubUsuario & " n„o tem acesso ao mÛdulo " & Formulario & " ou a empresa ainda n„o foi cadastrada."), vbExclamation, "CAPRIND v5.0"
End If

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub Direitos()
On Error GoTo tratar_erro

If pubIDUsuario <> "" Then
    Set TBAcessos = CreateObject("adodb.recordset")
    TBAcessos.Open "Select * FROM Acessos WHERE IDUsuario = " & pubIDUsuario & "AND Acesso = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBAcessos.EOF = False Then
        Incluir = TBAcessos!Incluir
        Alterar = TBAcessos!Alterar
        Excluir = TBAcessos!Excluir
    End If
    TBAcessos.Close
End If

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub Jump(KeyAscii As Integer)
On Error GoTo tratar_erro

If KeyAscii = 13 Then SendKeys "{TAB}"

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcValorImposto(Documento As String, ID_CF As Long, IDClienteImposto As Long, ClienteImposto As String, UF As String, ID_empresa As Integer, Terceiros As Boolean, ID_CFOP As Long, Regime As Integer)
On Error GoTo tratar_erro

If Documento = "" Then Exit Sub

'Verifica se È nota prÛpria ou de terceiros para verif. o regime tribut·rio
Set TBCFOP = CreateObject("adodb.recordset")
If Terceiros = False Then
ProcVerificaRegime
If RegimeEmpresa = 1 Then Simples = True
Else
    Set TBCFOP = CreateObject("adodb.recordset")
    TBCFOP.Open "Select * from clientes where IDCliente = " & IDClienteImposto & " and NomeRazao = '" & ClienteImposto & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBCFOP.EOF = True Then
        Set TBCFOP = CreateObject("adodb.recordset")
        TBCFOP.Open "Select * from compras_fornecedores where IDCliente = " & IDClienteImposto & " and Nome_Razao = '" & ClienteImposto & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBCFOP.EOF = False Then
            If TBCFOP!Simples = True Then Simples = True
        End If
    End If
End If

If Simples = True Then
    Set TBTempo = CreateObject("adodb.recordset")
    TBTempo.Open "Select * from tbl_NaturezaOperacao where IDCountCfop = " & ID_CFOP & " and (Devolucao = 'true' or Left(id_CFOP, 1) = '3')", Conexao, adOpenKeyset, adLockOptimistic
    If TBTempo.EOF = False Then
        ProcBuscaTributos (ID_CF)
        ProcVerificaRegiao UF, IDClienteImposto, ClienteImposto
    Else
        vlrICMS(0, 0) = 0
        vlrICMS(0, 1) = 0
        vlrICMS(0, 2) = 0
        vlrICMS(0, 3) = 0
        vlrICMS(0, 4) = 0
        vlrICMS(0, 5) = 0
        vlrICMS(0, 6) = 0
        CTDE = 0
        CTSS = 0
        CTNN = 0
        CTCO = 0
        CTEX = 0
        DIFDE = 0
        DIFSS = 0
        DIFNN = 0
        DIFCO = 0
        DIFEX = 0
        FCP = 0
        Reducao = False
        Retencao_PIS_Cofins = False
        
        vRegiao(0, 0) = 0
        vRegiao(0, 1) = 0
        CT = 0
    End If
Else
    ProcBuscaTributos (ID_CF)
    ProcVerificaRegiao UF, IDClienteImposto, ClienteImposto
End If
IntIPI = vRegiao(0, 0)
IntICMS = vRegiao(0, 1)
      
Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcBuscaTributos(ID_CF As Long)
On Error GoTo tratar_erro

vlrICMS(0, 0) = 0
vlrICMS(0, 1) = 0
vlrICMS(0, 2) = 0
vlrICMS(0, 3) = 0
vlrICMS(0, 4) = 0
vlrICMS(0, 5) = 0
vlrICMS(0, 6) = 0
CTDE = 0
CTSS = 0
CTNN = 0
CTCO = 0
CTEX = 0
DIFDE = 0
DIFSS = 0
DIFNN = 0
DIFCO = 0
DIFEX = 0
Reducao = False
Retencao_PIS_Cofins = False
Set TBCorretiva = CreateObject("adodb.recordset")
TBCorretiva.Open "Select * from tbl_ClassificacaoFiscal where Idclass = " & ID_CF, Conexao, adOpenKeyset, adLockOptimistic
If TBCorretiva.EOF = False Then
    vlrICMS(0, 0) = IIf(IsNull(TBCorretiva!IDIntClasse), 0, TBCorretiva!IDIntClasse)
    vlrICMS(0, 1) = IIf(IsNull(TBCorretiva!dbl_IPI), 0, TBCorretiva!dbl_IPI)
    vlrICMS(0, 2) = IIf(IsNull(TBCorretiva!dbl_ICMS_de), 0, TBCorretiva!dbl_ICMS_de)
    vlrICMS(0, 3) = IIf(IsNull(TBCorretiva!dbl_ICMS_ss), 0, TBCorretiva!dbl_ICMS_ss)
    vlrICMS(0, 4) = IIf(IsNull(TBCorretiva!dbl_ICMS_nn), 0, TBCorretiva!dbl_ICMS_nn)
    vlrICMS(0, 5) = IIf(IsNull(TBCorretiva!dbl_ICMS_co), 0, TBCorretiva!dbl_ICMS_co)
    vlrICMS(0, 6) = IIf(IsNull(TBCorretiva!dbl_ICMS_ex), 0, TBCorretiva!dbl_ICMS_ex)
    CTDE = IIf(IsNull(TBCorretiva!CTDE), 0, TBCorretiva!CTDE)
    CTSS = IIf(IsNull(TBCorretiva!CTSS), 0, TBCorretiva!CTSS)
    CTNN = IIf(IsNull(TBCorretiva!CTNN), 0, TBCorretiva!CTNN)
    CTCO = IIf(IsNull(TBCorretiva!CTCO), 0, TBCorretiva!CTCO)
    CTEX = IIf(IsNull(TBCorretiva!CTEX), 0, TBCorretiva!CTEX)
    DIFDE = IIf(IsNull(TBCorretiva!DIFDE), 0, TBCorretiva!DIFDE)
    DIFSS = IIf(IsNull(TBCorretiva!DIFSS), 0, TBCorretiva!DIFSS)
    DIFNN = IIf(IsNull(TBCorretiva!DIFNN), 0, TBCorretiva!DIFNN)
    DIFCO = IIf(IsNull(TBCorretiva!DIFCO), 0, TBCorretiva!DIFCO)
    DIFEX = IIf(IsNull(TBCorretiva!DIFEX), 0, TBCorretiva!DIFEX)
    If TBCorretiva!basereduzida = True Then Reducao = True Else Reducao = False
    If TBCorretiva!Retem_PIS_Cofins = True Then Retencao_PIS_Cofins = True Else Retencao_PIS_Cofins = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcControleImposto(IDCFOP As Long, IDCliente As Long)
On Error GoTo tratar_erro

TemICMS = "N√O"
TemIPI = "N√O"
TemPIS = False
TemCOFINS = False
TemReducaoBC = False
SomarIPI = "N√O"
SomarIPIST = False
Soma_retorno_total_NF = False
Desconto_Suframa = False
Suframa = False
MPA = False
DestacaImpostos = "N√O"

Set TBCFOP = CreateObject("adodb.recordset")
TBCFOP.Open "Select * FROM tbl_NaturezaOperacao WHERE IDCountCfop = " & IDCFOP, Conexao, adOpenKeyset, adLockOptimistic
If TBCFOP.EOF = False Then
    TemICMS = IIf(IsNull(TBCFOP!Txt_ICMS), True, TBCFOP!Txt_ICMS)
    TemIPI = IIf(IsNull(TBCFOP!txt_IPI), True, TBCFOP!txt_IPI)
    TemPIS = IIf(IsNull(TBCFOP!TemPIS), True, TBCFOP!TemPIS)
    TemCOFINS = IIf(IsNull(TBCFOP!TemCOFINS), True, TBCFOP!TemCOFINS)
    TemReducaoBC = IIf(TBCFOP!TemReducaoBC = True, True, False)
    SomarIPI = IIf(IsNull(TBCFOP!txt_Somar), True, TBCFOP!txt_Somar)
    SomarIPIST = IIf(TBCFOP!Somar_IPI_BC_ICMSST = True, True, False)
    Soma_retorno_total_NF = IIf(TBCFOP!Soma_retorno_totalnf = True, True, False)
    Desconto_Suframa = IIf(TBCFOP!Suframa = True, True, False)
    MPA = IIf(IsNull(TBCFOP!MPA), True, TBCFOP!MPA)
    MPA = IIf(MPA = True, True, False)
    DestacaImpostos = IIf(TBCFOP!Retem = True, "SIM", "N√O")
    
    Suframa = False
    Set TBCFOP = CreateObject("adodb.recordset")
    TBCFOP.Open "Select IDCliente from Clientes where IDCliente = " & IDCliente & " and chkSuframa = 'True'", Conexao, adOpenKeyset, adLockOptimistic
    If TBCFOP.EOF = False Then
        Suframa = True
    End If
End If
TBCFOP.Close
    
Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcVerificaRegiao(UF As String, IDClienteRegiao As Long, ClienteRegiao As String)
On Error GoTo tratar_erro

vRegiao(0, 0) = 0
vRegiao(0, 1) = 0
CT = 0
DIF = 0
FCP = 0

vRegiao(0, 0) = vlrICMS(0, 1) 'IPI
Set TBCorretiva = CreateObject("adodb.recordset")
TBCorretiva.Open "Select * from Clientes where IDCliente = " & IDClienteRegiao & " and NomeRazao = '" & ClienteRegiao & "' and idTipoEmpresa = 0", Conexao, adOpenKeyset, adLockOptimistic
If TBCorretiva.EOF = False Then
    vRegiao(0, 1) = vlrICMS(0, 6)
    CT = CTEX
    DIF = DIFEX
Else
    Set TBCorretiva = CreateObject("adodb.recordset")
    TBCorretiva.Open "Select * from Compras_fornecedores where IDCliente = " & IDClienteRegiao & " and Nome_Razao = '" & ClienteRegiao & "' and idTipoEmpresa = 0", Conexao, adOpenKeyset, adLockOptimistic
    If TBCorretiva.EOF = False Then
        vRegiao(0, 1) = vlrICMS(0, 6)
        CT = CTEX
        DIF = DIFEX
    Else
        Set TBCorretiva = CreateObject("adodb.recordset")
        TBCorretiva.Open "Select regiao, FCP from regioes where uf = '" & UF & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBCorretiva.EOF = False Then
            Select Case TBCorretiva!regiao
                Case "DE":
                    vRegiao(0, 1) = vlrICMS(0, 2)
                    CT = CTDE
                    DIF = DIFDE
                Case "SS":
                    vRegiao(0, 1) = vlrICMS(0, 3)
                    CT = CTSS
                    DIF = DIFSS
                Case "NN":
                    vRegiao(0, 1) = vlrICMS(0, 4)
                    CT = CTNN
                    DIF = DIFNN
                Case "CO":
                    vRegiao(0, 1) = vlrICMS(0, 5)
                    CT = CTCO
                    DIF = DIFCO
            End Select
            FCP = IIf(IsNull(TBCorretiva!FCP), 0, TBCorretiva!FCP)
        End If
    End If
End If
TBCorretiva.Close

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Function FunVerificaICMS(vNatOp As String) As String
On Error GoTo tratar_erro

Set TBnatOp = New ADODB.Recordset
StrSql = "Select * FROM tbl_NaturezaOperacao WHERE id_CFOP =" & vNatOp & ";"
TBnatOp.CursorLocation = adUseClient
TBnatOp.Open StrSql, Conexao, adOpenDynamic, adLockOptimistic
If TBnatOp.RecordCount = 0 Then
    USMsgBox "Natureza n„o cadastrada", vbCritical, "CAPRIND v5.0"
    Exit Function
End If
FunVerificaICMS = TBnatOp("txt_ICMS")
TBnatOp.Close
Set TBnatOp = Nothing

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Public Function FunVerificaIPI(vNatOp As String) As String
On Error GoTo tratar_erro

Set TBnatOp = New ADODB.Recordset
StrSql = "Select * FROM tbl_NaturezaOperacao WHERE id_CFOP =" & vNatOp & ";"
TBnatOp.CursorLocation = adUseClient
TBnatOp.Open StrSql, Conexao, adOpenDynamic, adLockOptimistic
If TBnatOp.RecordCount = 0 Then
    USMsgBox "Natureza n„o cadastrada", vbCritical, "CAPRIND v5.0"
    Exit Function
End If
FunVerificaIPI = TBnatOp("txt_IPI")
TBnatOp.Close
Set TBnatOp = Nothing

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Public Function FunValorExtenso(nvalor)
On Error GoTo tratar_erro

'Vari·veis
Dim nContador, nTamanho As Integer
Dim cValor, cParte, cFinal As String
ReDim aGrupo(4), aTexto(4) As String

'Valida Argumento
If IsNull(nvalor) Or nvalor <= 0 Or nvalor > 9999999.99 Then Exit Function

'Matrizes de FunValorExtensos (Parciais)
ReDim aUnid(19) As String
aUnid(1) = "um ": aUnid(2) = "dois ": aUnid(3) = "tres "
aUnid(4) = "quatro ": aUnid(5) = "cinco ": aUnid(6) = "seis "
aUnid(7) = "sete ": aUnid(8) = "oito ": aUnid(9) = "nove "
aUnid(10) = "dez ": aUnid(11) = "onze ": aUnid(12) = "doze "
aUnid(13) = "treze ": aUnid(14) = "quatorze ": aUnid(15) = "quinze "
aUnid(16) = "dezesseis ": aUnid(17) = "dezessete ": aUnid(18) = "dezoito "
aUnid(19) = "dezenove "

ReDim aDezena(9) As String
aDezena(1) = "dez ": aDezena(2) = "vinte ": aDezena(3) = "trinta "
aDezena(4) = "quarenta ": aDezena(5) = "cinquenta "
aDezena(6) = "sessenta ": aDezena(7) = "setenta ": aDezena(8) = "oitenta "
aDezena(9) = "noventa "

ReDim aCentena(9) As String
aCentena(1) = "cento ": aCentena(2) = "duzentos "
aCentena(3) = "trezentos ": aCentena(4) = "quatrocentos "
aCentena(5) = "quinhentos ": aCentena(6) = "seiscentos "
aCentena(7) = "setecentos ": aCentena(8) = "oitocentos "
aCentena(9) = "novecentos "

'Separa valor em grupos
cValor = Format$(nvalor, "0000000000.00")
aGrupo(1) = Mid$(cValor, 2, 3)
aGrupo(2) = Mid$(cValor, 5, 3)
aGrupo(3) = Mid$(cValor, 8, 3)
aGrupo(4) = "0" + Mid$(cValor, 12, 2)

'Calcula cada grupo
For nContador = 1 To 4
  cParte = aGrupo(nContador)
  nTamanho = Switch(Val(cParte) < 10, 1, Val(cParte) < 100, 2, Val(cParte) < 1000, 3)
  If nTamanho = 3 Then
    If Right$(cParte, 2) <> "00" Then
      aTexto(nContador) = aTexto(nContador) + aCentena(Left(cParte, 1)) + "e "
      nTamanho = 2
    Else
      aTexto(nContador) = aTexto(nContador) + IIf(Left$(cParte, 1) = "1", "cem ", aCentena(Left(cParte, 1)))
    End If
  End If
  If nTamanho = 2 Then
    If Val(Right(cParte, 2)) < 20 Then
      aTexto(nContador) = aTexto(nContador) + aUnid(Right(cParte, 2))
    Else
      aTexto(nContador) = aTexto(nContador) + aDezena(Mid(cParte, 2, 1))
      If Right$(cParte, 1) <> "0" Then
        aTexto(nContador) = aTexto(nContador) + "e "
        nTamanho = 1
      End If
    End If
  End If
  If nTamanho = 1 Then
    aTexto(nContador) = aTexto(nContador) + aUnid(Right(cParte, 1))
  End If
Next

'Final
If Val(aGrupo(1) + aGrupo(2) + aGrupo(3)) = 0 And Val(aGrupo(4)) <> 0 Then
  cFinal = aTexto(4) + IIf(Val(aGrupo(4)) = 1, "centavo", "centavos")
Else
  cFinal = ""
  cFinal = cFinal + IIf(Val(aGrupo(1)) <> 0, aTexto(1) + IIf(Val(aGrupo(1)) > 1, "milhıes ", "milh„o "), "")
  If Val(aGrupo(2) + aGrupo(3)) = 0 Then
    cFinal = cFinal + "de "
  Else
    cFinal = cFinal + IIf(Val(aGrupo(2)) <> 0, aTexto(2) + "mil ", "")
  End If
  cFinal = cFinal + aTexto(3) + IIf(Val(aGrupo(1) + aGrupo(2) + aGrupo(3)) = 1, "real ", "reais ")
  cFinal = cFinal + IIf(Val(aGrupo(4)) <> 0, "E " + aTexto(4) + IIf(Val(aGrupo(4)) = 1, "centavo", "centavos"), "")
End If
FunValorExtenso = UCase$(cFinal)

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Public Function FunCalculaSegPC(TExec As String, PcHora As Double)
On Error GoTo tratar_erro

If IsDate(TExec) = True Then
    Dataini = TExec
    ElapsedTime (Dataini)
Else
    ProcFormataHora (TExec)
End If
Valor1 = s
Valor2 = PcHora

If Valor1 And Valor2 <> 0 Then
'If Valor2 >= 1 Then
FunCalculaSegPC = Valor1 / Valor2
'Else
'FunCalculaSegPC = Valor1 * Valor2

'End If

Else
FunCalculaSegPC = 0
End If


Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Public Function FunTiraAcentosTexto(Texto As String)
On Error GoTo tratar_erro
Dim Texto_limpo As String 'OK
Dim Codigo_Tabela_Asc As Integer 'OK

i = Len(Texto)
For Posicao = 1 To Len(Texto)
    Codigo_Tabela_Asc = Asc(Mid(Texto, Posicao, 1))
    Select Case Codigo_Tabela_Asc
        Case 10:
            If Posicao = 1 Then
                Codigo_Tabela_Asc = Asc("-")
            ElseIf Posicao = i Then
                    Codigo_Tabela_Asc = Asc(".")
                Else
                    Codigo_Tabela_Asc = Asc(" ")
            End If
        Case 13:
            If Posicao = 1 Then
                Codigo_Tabela_Asc = Asc("-")
            ElseIf Posicao = i Then
                    Codigo_Tabela_Asc = Asc(".")
                Else
                    Codigo_Tabela_Asc = Asc("|")
            End If
        Case 39: Codigo_Tabela_Asc = Asc(" ")
        Case 192 To 197: Codigo_Tabela_Asc = Asc("A")
        Case 224 To 229: Codigo_Tabela_Asc = Asc("a")
        Case 200 To 203: Codigo_Tabela_Asc = Asc("E")
        Case 232 To 235: Codigo_Tabela_Asc = Asc("e")
        Case 204 To 207: Codigo_Tabela_Asc = Asc("I")
        Case 236 To 239: Codigo_Tabela_Asc = Asc("i")
        Case 186: Codigo_Tabela_Asc = Asc(".")
        Case 199: Codigo_Tabela_Asc = Asc("C")
        Case 231: Codigo_Tabela_Asc = Asc("c")
        Case 210 To 214: Codigo_Tabela_Asc = Asc("O")
        Case 242 To 246: Codigo_Tabela_Asc = Asc("o")
        Case 154: Codigo_Tabela_Asc = Asc("U")
        Case 217 To 220: Codigo_Tabela_Asc = Asc("U")
        Case 249 To 252: Codigo_Tabela_Asc = Asc("u")
    End Select
    Texto_limpo = Texto_limpo & Chr(Codigo_Tabela_Asc)
Next
Texto_limpo = Replace(Texto_limpo, "`", "")

FunTiraAcentosTexto = Texto_limpo
'Debug.print Texto_limpo

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Public Function FunVerificaCodUF(Cidade As String, UF As String)
On Error GoTo tratar_erro

Set TBAliquota = CreateObject("adodb.recordset")
TBAliquota.Open "Select * from CEP where Municipio = '" & Cidade & "' and Sigla_UF = '" & UF & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBAliquota.EOF = False Then
     FunVerificaCodUF = TBAliquota!Codigo_UF
End If
TBAliquota.Close
    
Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Public Function FunVerificaCodMunicipio(Cidade As String, UF As String)
On Error GoTo tratar_erro

Set TBAliquota = CreateObject("adodb.recordset")
'Debug.print Cidade
Cidade2 = Replace(Cidade, "'", "o ")

'Debug.print Cidade2

TBAliquota.Open "Select * from CEP where Municipio = '" & Cidade2 & "' and Sigla_UF = '" & UF & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBAliquota.EOF = False Then
     FunVerificaCodMunicipio = TBAliquota!Codigo_municipio
End If
TBAliquota.Close
    
Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Public Function FunVerificaCodMunicipioDIPAM(Cidade As String, UF As String)
On Error GoTo tratar_erro

Set TBAliquota = CreateObject("adodb.recordset")
TBAliquota.Open "Select * from CEP where Municipio = '" & Cidade & "' and Sigla_UF = '" & UF & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBAliquota.EOF = False Then
     FunVerificaCodMunicipioDIPAM = IIf(IsNull(TBCFOP!Codigo_municipio_DIPAM), "", TBCFOP!Codigo_municipio_DIPAM)
End If
TBAliquota.Close
    
Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Public Function FunTamanhoTextoZeroEsq(Texto As Variant, Tamanho As Integer) As String
On Error GoTo tratar_erro
Dim QuantZeroEsq As Double 'OK

Texto1 = ""
If Len(Texto) < Tamanho Then
QuantZeroEsq = Tamanho - Len(Texto)
Else
QuantZeroEsq = Len(Texto) - Tamanho
End If

If QuantZeroEsq > 0 Then
    Do While QuantZeroEsq > 0
        If Texto1 = "" Then Texto1 = "0" Else Texto1 = Texto1 & "0"
        QuantZeroEsq = QuantZeroEsq - 1
    Loop
    FunTamanhoTextoZeroEsq = Texto1 & Texto
Else
    FunTamanhoTextoZeroEsq = Texto
End If

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Public Function FunTamanhoTextoZeroDir(Texto As Variant, Tamanho As Integer) As String
On Error GoTo tratar_erro
Dim QuantZeroDir As Double 'OK

Texto1 = ""
QuantZeroDir = Tamanho - Len(Texto)
If QuantZeroDir > 0 Then
    Do While QuantZeroDir > 0
        If Texto1 = "" Then Texto1 = "0" Else Texto1 = Texto1 & "0"
        QuantZeroDir = QuantZeroDir - 1
    Loop
    FunTamanhoTextoZeroDir = Texto & Texto1
Else
    FunTamanhoTextoZeroDir = Texto
End If

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Function

Public Function FunTamanhoTextoVazioDir(Texto As Variant, Tamanho As Integer) As String
On Error GoTo tratar_erro
Dim QuantVazioDir As Double 'OK

Texto1 = ""
QuantVazioDir = Tamanho - Len(Texto)
If QuantVazioDir > 0 Then
    Do While QuantVazioDir > 0
        If Texto1 = "" Then Texto1 = " " Else Texto1 = Texto1 & " "
        QuantVazioDir = QuantVazioDir - 1
    Loop
    FunTamanhoTextoVazioDir = Texto & Texto1
Else
    FunTamanhoTextoVazioDir = Texto
End If

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Sub ProcOrdenaListView(ByVal lvw As MSComctlLib.ListView, ByVal Coluna_Cabecalho As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If Coluna_Cabecalho.Tag = "N" Then
    ProcSortListView lvw, Coluna_Cabecalho.index, "ldtNumber", OrdAsc
ElseIf Coluna_Cabecalho.Tag = "T" Then
    ProcSortListView lvw, Coluna_Cabecalho.index, "ldtString", OrdAsc
ElseIf Coluna_Cabecalho.Tag = "D" Then
        ProcSortListView lvw, Coluna_Cabecalho.index, "ldtDateTime", OrdAsc
End If
OrdAsc = Not OrdAsc

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcSortListView(ListView As ListView, ByVal index As Integer, ByVal DataType As String, ByVal Ascending As Boolean)
On Error GoTo tratar_erro
Dim i As Integer
Dim L As Long
Dim strFormat As String
Dim lngCursor As Long
Dim blnRestoreFromTag As Boolean
Dim dte As Date

Permitido = False

lngCursor = ListView.MousePointer
ListView.MousePointer = vbHourglass
LockWindowUpdate ListView.hWnd

Select Case DataType
    Case "ldtString": blnRestoreFromTag = False
    Case "ldtNumber":
        strFormat = String$(20, "0") & "." & String$(10, "0")
        With ListView.ListItems
            If (index = 1) Then
                For L = 1 To .Count
                    With .Item(L)
                        If IsNumeric(.Text) Or Right(.Text, 1) = "%" Then
                            If Right(.Text, 1) = "%" Then
                                valor = Len(.Text) - 1
                                Familiatext = Mid(.Text, 1, valor)
                                Permitido = True
                            Else
                                Familiatext = .Text
                            End If
                            
                            .Tag = Familiatext & Chr$(0) & .Tag
                            If CDbl(Familiatext) >= 0 Then .Text = Format(CDbl(Familiatext), strFormat) Else .Text = "&" & Format(0 - CDbl(Familiatext), strFormat)
                        Else
                            .Tag = .Text & Chr$(0) & .Tag
                            .Text = ""
                        End If
                    End With
                Next L
            Else
                For L = 1 To .Count
                    With .Item(L).ListSubItems(index - 1)
                        If IsNumeric(.Text) Or Right(.Text, 1) = "%" Then
                            If Right(.Text, 1) = "%" Then
                                valor = Len(.Text) - 1
                                Familiatext = Mid(.Text, 1, valor)
                                Permitido = True
                            Else
                                Familiatext = .Text
                            End If
                            
                            .Tag = Familiatext & Chr$(0) & .Tag
                            If CDbl(Familiatext) >= 0 Then .Text = Format(CDbl(Familiatext), strFormat) Else .Text = "&" & Format(0 - CDbl(Familiatext), strFormat)
                        Else
                            .Tag = .Text & Chr$(0) & .Tag
                            .Text = ""
                        End If
                    End With
                Next L
            End If
        End With
        blnRestoreFromTag = True
    Case "ldtDateTime":
        strFormat = "YYYYMMDDHhNnSs"
        With ListView.ListItems
            If (index = 1) Then
                For L = 1 To .Count
                    With .Item(L)
                        If .Text <> "" Then
                            .Tag = .Text & Chr$(0) & .Tag
                            dte = (.Text)
                            .Text = Format$(dte, strFormat)
                        End If
                    End With
                Next L
            Else
                For L = 1 To .Count
                    With .Item(L).ListSubItems(index - 1)
                        If .Text <> "" Then
                            .Tag = .Text & Chr$(0) & .Tag
                            dte = (.Text)
                            .Text = Format$(dte, strFormat)
                        End If
                    End With
                Next L
            End If
        End With
        blnRestoreFromTag = True
End Select
  
ListView.SortOrder = IIf(Ascending, lvwAscending, lvwDescending)
ListView.SortKey = index - 1
ListView.Sorted = True
    
If blnRestoreFromTag Then
    With ListView.ListItems
        If (index = 1) Then
            For L = 1 To .Count
                With .Item(L)
                    If .Tag <> "" Then
                        i = InStr(.Tag, Chr$(0))
                        .Text = Left$(.Tag, i - 1)
                        .Tag = Mid$(.Tag, i + 1)
                    End If
                End With
            Next L
        Else
            For L = 1 To .Count
                With .Item(L).ListSubItems(index - 1)
                    If .Tag <> "" Then
                        i = InStr(.Tag, Chr$(0))
                        .Text = Left$(.Tag, i - 1)
                        .Tag = Mid$(.Tag, i + 1)
                    End If
                    If Permitido = True Then .Text = .Text & "%"
                End With
            Next L
        End If
    End With
End If
    
LockWindowUpdate 0&
ListView.MousePointer = lngCursor
ListView.Sorted = False

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Function FunCalculaValorUnitOrdem(OF As Long, QuantsolicitadoN1 As Double, qt As Double, Qtde As Double, CTLote As Double, CTPecaReal As Double, CTTerceiros As Double, CTMaterial As Double, CTOutras As Double, Consignada As Boolean)
On Error GoTo tratar_erro
Dim TBMaterialVlrUnitOrdem As ADODB.Recordset
Dim TBEstoqueVlrUnitOrdem As ADODB.Recordset
Dim TBproducaoVlrUnitOrdem As ADODB.Recordset
Dim TBAbrirVlrUnitOrdem As ADODB.Recordset
Dim Permitido_Calula_Vlr_Unit_Ordem As Boolean

'Valor NC
Valor_Cofins_Prod = 0
Valor1 = 0 'ServiÁo
Valor2 = 0 'Material
Valor3 = 0 'M„o de obra
ValorConta = 0 'Outras
Valor_CSLL_Serv = 0
Valor_INSS_Serv = 0
Valor_IPI = 0
Permitido_Calula_Vlr_Unit_Ordem = False

'Custo de material
If Consignada = False Then
    Set TBMaterialVlrUnitOrdem = CreateObject("adodb.recordset")
    TBMaterialVlrUnitOrdem.Open "Select Valor_saida_estoque, Saida from Producaomaterial where Ordem = " & OF & " order by Codigo", Conexao, adOpenKeyset, adLockOptimistic
    If TBMaterialVlrUnitOrdem.EOF = False Then
        Do While TBMaterialVlrUnitOrdem.EOF = False
            
            'Verifica valor total do material
            Valor_CSLL_Prod = IIf(IsNull(TBMaterialVlrUnitOrdem!Valor_saida_estoque), 0, TBMaterialVlrUnitOrdem!Valor_saida_estoque)
            If TBMaterialVlrUnitOrdem!Saida <> "N√O" Then
                Set TBproducaoVlrUnitOrdem = CreateObject("adodb.recordset")
                TBproducaoVlrUnitOrdem.Open "Select Totalprod from ordemservico where Ordem = " & OF & " ORDER BY fase, retrabalho, IDproducao", Conexao, adOpenKeyset, adLockOptimistic
                If TBproducaoVlrUnitOrdem.EOF = False Then
                    Qtd_Prog = IIf(IsNull(TBproducaoVlrUnitOrdem!Totalprod), 0, TBproducaoVlrUnitOrdem!Totalprod) 'Qtde. produzida
                    If Qtd_Prog <> 0 Then
                        Valor_CSLL_Serv = Format(Valor_CSLL_Serv + (Valor_CSLL_Prod / Qtd_Prog), "###,##0.0000000000")
                    ElseIf Qtde <> 0 Then
                            Valor_CSLL_Serv = Format(Valor_CSLL_Serv + (Valor_CSLL_Prod / Qtde), "###,##0.0000000000")
                    End If
                End If
                TBproducaoVlrUnitOrdem.Close
            End If
            TBMaterialVlrUnitOrdem.MoveNext
        Loop
    End If
    TBMaterialVlrUnitOrdem.Close
End If

'Verifica qtde NC da ordem
QuantComprado = 0
Set TBAbrirVlrUnitOrdem = CreateObject("adodb.recordset")
TBAbrirVlrUnitOrdem.Open "Select Sum(TTNC) as QtdeNC from CQ_NC_FABRICA where Ordem = " & OF & " and PARECERCQ = 'Rejeitar'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrirVlrUnitOrdem.EOF = False Then
    QuantComprado = IIf(IsNull(TBAbrirVlrUnitOrdem!qtdeNC), 0, TBAbrirVlrUnitOrdem!qtdeNC)
End If

'Verifica ˙ltima OS com NC
Set TBAbrirVlrUnitOrdem = CreateObject("adodb.recordset")
TBAbrirVlrUnitOrdem.Open "Select OS.Fase FROM ordemservico OS INNER JOIN CQ_NC_FABRICA CQNC ON OS.Idproducao = CQNC.OS where OS.Ordem = " & OF & " and CQNC.PARECERCQ = 'Rejeitar' order by OS.fase, OS.retrabalho, OS.IDproducao", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrirVlrUnitOrdem.EOF = False Then
    TBAbrirVlrUnitOrdem.MoveLast
    OS = TBAbrirVlrUnitOrdem!Fase
End If
Set TBproducaoVlrUnitOrdem = CreateObject("adodb.recordset")
TBproducaoVlrUnitOrdem.Open "Select * from ordemservico where Ordem = " & OF & " and Fase <= " & OS & " ORDER BY fase, retrabalho, IDproducao", Conexao, adOpenKeyset, adLockOptimistic
If TBproducaoVlrUnitOrdem.EOF = False Then
    Do While TBproducaoVlrUnitOrdem.EOF = False
        'Soma valor unit·rio do SERVI«O na OS
        If IsNull(TBproducaoVlrUnitOrdem!Totalprod) = False And TBproducaoVlrUnitOrdem!Totalprod <> "" And TBproducaoVlrUnitOrdem!Totalprod <> "0" Then Valor_IPI = Format(Valor_IPI + (TBproducaoVlrUnitOrdem!CTServico / TBproducaoVlrUnitOrdem!Totalprod), "###,##0.0000000000")
        
        'Soma valor unit·rio da M√O DE OBRA na OS
        Valor_Cofins_Prod = Format(Valor_Cofins_Prod + TBproducaoVlrUnitOrdem!CRPECA, "###,##0.0000000000")
        'Verifica qtde. regufada na OS
        Qtd_Prog = 0
        Set TBAbrirVlrUnitOrdem = CreateObject("adodb.recordset")
        TBAbrirVlrUnitOrdem.Open "Select Sum(TTNC) as QtdeNC from CQ_NC_FABRICA where OS = " & TBproducaoVlrUnitOrdem!IDProducao & " and PARECERCQ = 'Rejeitar'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrirVlrUnitOrdem.EOF = False Then
            Permitido_Calula_Vlr_Unit_Ordem = True
            Qtd_Prog = IIf(IsNull(TBAbrirVlrUnitOrdem!qtdeNC), 0, TBAbrirVlrUnitOrdem!qtdeNC)
            Valor1 = Format(Valor1 + (Valor_IPI * Qtd_Prog), "###,##0.00") 'Valor total unit·rio serviÁo x qtde. refugada da OS
            Valor3 = Format(Valor3 + (Valor_Cofins_Prod * Qtd_Prog), "###,##0.00") 'Valor total unit·rio m„o de obra x qtde. refugada da OS
        End If
        TBAbrirVlrUnitOrdem.Close
        TBproducaoVlrUnitOrdem.MoveNext
    Loop
End If
If Permitido_Calula_Vlr_Unit_Ordem = True Then
    'Valor do material por peÁa x qtde. refugada
    If QuantsolicitadoN1 <> 0 Then Valor2 = Format(Valor_CSLL_Serv * QuantComprado, "###,##0.00")
                       'SE  +   MT   +   MO
    ValorNC = Format(Valor1 + Valor2 + Valor3, "###,##0.00")
Else
    ValorNC = 0
End If

'Custo de MO do lote
Valor_Cofins_Prod = CTLote

'Custo de terceiros por peÁa
Valor2 = 0
Set TBproducaoVlrUnitOrdem = CreateObject("adodb.recordset")
TBproducaoVlrUnitOrdem.Open "Select Totalprod, CTServico from ordemservico where Ordem = " & OF & " and Custos = 'False' ORDER BY fase, retrabalho, IDproducao", Conexao, adOpenKeyset, adLockOptimistic
If TBproducaoVlrUnitOrdem.EOF = False Then
    Do While TBproducaoVlrUnitOrdem.EOF = False
        If TBproducaoVlrUnitOrdem!Totalprod <> 0 Then
            Valor2 = Valor2 + (IIf(IsNull(TBproducaoVlrUnitOrdem!CTServico), 0, TBproducaoVlrUnitOrdem!CTServico) / IIf(IsNull(TBproducaoVlrUnitOrdem!Totalprod), 0, TBproducaoVlrUnitOrdem!Totalprod))
        ElseIf Qtde <> 0 Then
                Valor2 = Valor2 + (IIf(IsNull(TBproducaoVlrUnitOrdem!CTServico), 0, TBproducaoVlrUnitOrdem!CTServico) / Qtde)
        End If
        TBproducaoVlrUnitOrdem.MoveNext
    Loop
End If
TBproducaoVlrUnitOrdem.Close

'Valor por peÁa
'Custo total
Valor_Cofins_Prod = Format(CTTerceiros + CTMaterial + CTLote + CTOutras, "###,##0.00")
If ValorNC <> 0 And QuantsolicitadoN1 = Qtde Then
                                            'Custo total  - Valor NC / QTDE. OK                                                               Custo total  - Valor NC
    If qt <> 0 Then FunCalculaValorUnitOrdem = Format((Valor_Cofins_Prod - ValorNC) / qt, "###,##0.0000000000") Else FunCalculaValorUnitOrdem = Format(Valor_Cofins_Prod - ValorNC, "###,##0.0000000000")
Else
    If Qtde > 0 Then CTOutras = CTOutras / Qtde Else CTOutras = 0
                                    'SE    +        MT       +  MO        + OU
    FunCalculaValorUnitOrdem = Format((Valor2 + Valor_CSLL_Serv + CTPecaReal + CTOutras), "###,##0.0000000000")
End If

'Variaveis para o relatÛrio de custo resumido
Valor3 = CTPecaReal
Valor_Produto = Valor_CSLL_Serv
ValorPagar = Valor2
Valor_DAS = CTOutras

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

                              'Tipo do cÛd.         Filtro cÛd. aut.
Function FunCriaNovoProdServ(CodManual As Boolean, TextoFiltroAut As String, Codinterno As String, CodRef As String, Revisao As String, Descricao As String, DescricaoCom As String, Familia As String, PConsumo As Double, PRevenda As Double, PCusto As Double, Unidade As String, UnidadeCom As String, ID_CF As Long, Compras As Boolean, Vendas As Boolean, Producao As Boolean, Qualidade As Boolean, SubTipoItem As Integer, Tipo As String, Obs As String, Comprimento As Double, Largura As Double, Espessura As Double, Dureza As String, IDCliFornCodRef As Long, CliFornCodRef As String, TipoCliFornCodRef As String)
On Error GoTo tratar_erro

Permitido2 = False
'Verifica se no cadastro da empresa esta marcado para gerar codigo sequencial para produto final
If SubTipoItem = 1 Then
    Set TBAliquota = CreateObject("adodb.recordset")
    TBAliquota.Open "Select Codigo_sequencial from Empresa where Codigo_sequencial = 'True'", Conexao, adOpenKeyset, adLockReadOnly
    If TBAliquota.EOF = False Then Permitido2 = True
    TBAliquota.Close
End If

If CodManual = False Then
    Set TBFamilia = CreateObject("adodb.recordset")
    TBFamilia.Open "Select * from projfamilia where familia = '" & Familia & "'", Conexao, adOpenKeyset, adLockReadOnly
    If TBFamilia.EOF = False Then Letra = TBFamilia!Letra
    TBFamilia.Close
    
    CompLetra = Len(Letra)
    valor = 6 + CompLetra
    Set TBComponente = CreateObject("adodb.recordset")
    If Permitido2 = False Then
        If SubTipoItem <> 3 Then
        StrSql = "Select * from projproduto where classe = '" & Familia & "' and Right(Desenho, " & CompLetra & ") = '" & Letra & "' and Len(Desenho) = " & valor & " and " & TextoFiltroAut & " order by codproduto desc"
        Else
        StrSql = "Select * from projproduto where classe = '" & Familia & "' and left(Desenho, " & CompLetra & ") = '" & Letra & "' and Len(Desenho) = " & valor & " and " & TextoFiltroAut & " order by codproduto desc"
        End If
    'Debug.print StrSql
    
        TBComponente.Open StrSql, Conexao, adOpenKeyset, adLockReadOnly
    Else
        'O Codigo sequencial n„o verifica a familia para gerar o codigo interno
        TBComponente.Open "Select * from projproduto where codmanual = 'False' and subtipoitem = 1 order by codproduto desc", Conexao, adOpenKeyset, adLockReadOnly
    End If
    
    If TBComponente.EOF = False Then
        If SubTipoItem <> 3 Then
                If Permitido2 = False Then
                    Numero = Left(TBComponente!Desenho, Len(TBComponente!Desenho) - (CompLetra + 1))
                Else
                    Numero = Left(TBComponente!Desenho, 5)
                End If
                    Numero = Numero + 1
                Select Case Len(Numero)
                    Case 5: Desenho = Numero & "-" & Letra
                    Case 4: Desenho = "0" & Numero & "-" & Letra
                    Case 3: Desenho = "00" & Numero & "-" & Letra
                    Case 2: Desenho = "000" & Numero & "-" & Letra
                    Case 1: Desenho = "0000" & Numero & "-" & Letra
                End Select
        Else
                If Permitido2 = False Then
                    Numero = Right(TBComponente!Desenho, Len(TBComponente!Desenho) - (CompLetra + 1))
                Else
                    Numero = Right(TBComponente!Desenho, 5)
                 End If
                    Numero = Numero + 1
                Select Case Len(Numero)
                    Case 5: Desenho = Letra & "-" & Numero
                    Case 4: Desenho = Letra & "-" & "0" & Numero
                    Case 3: Desenho = Letra & "-" & "00" & Numero
                    Case 2: Desenho = Letra & "-" & "000" & Numero
                    Case 1: Desenho = Letra & "-" & "0000" & Numero  '"0000" & Numero & "-" & Letra
                End Select
        End If
        
    Else
            If SubTipoItem <> 3 Then
                Desenho = "00001" & "-" & Letra
            Else
                 Desenho = Letra & "-" & "00001"
            End If
    
    End If
    TBComponente.Close
    
VerifCodigo:
    Set TBFIltro = CreateObject("adodb.recordset")
    TBFIltro.Open "Select * from projproduto where desenho = '" & Desenho & "'", Conexao, adOpenKeyset, adLockReadOnly
    If TBFIltro.EOF = False Then
        Numero = Left(Desenho, Len(Desenho) - (CompLetra + 1))
        Numero = Numero + 1
        Select Case Len(Numero)
            Case 5: Desenho = Numero & "-" & Letra
            Case 4: Desenho = "0" & Numero & "-" & Letra
            Case 3: Desenho = "00" & Numero & "-" & Letra
            Case 2: Desenho = "000" & Numero & "-" & Letra
            Case 1: Desenho = "0000" & Numero & "-" & Letra
        End Select
        GoTo VerifCodigo
    End If
    
    Codinterno = Desenho
Else
    Desenho = Codinterno
End If
Set TBComponente = CreateObject("adodb.recordset")
TBComponente.Open "Select * from projproduto", Conexao, adOpenKeyset, adLockOptimistic
TBComponente.AddNew
TBComponente!Data = Date
TBComponente!DtValidacao = Date
TBComponente!Responsavel = pubUsuario
TBComponente!RespValidacao = pubUsuario
TBComponente!Desenho = Codinterno
TBComponente!RevDesenho = Revisao
TBComponente!Descricao = Descricao
TBComponente!descricaotecnica = DescricaoCom
TBComponente!Classe = Familia
TBComponente!Unidade = Unidade
TBComponente!Unidade_com = UnidadeCom
TBComponente!ID_CF = IIf(ID_CF = 0, Null, ID_CF)
TBComponente!PConsumo = PConsumo
TBComponente!PRevenda = PRevenda
TBComponente!PCusto = PCusto
TBComponente!Compras = Compras
TBComponente!Vendas = Vendas
TBComponente!Producao = Producao
TBComponente!Qualidade = Qualidade
TBComponente!SubTipoItem = SubTipoItem
TBComponente!Leadtime = 0
TBComponente!CodManual = CodManual
TBComponente!Tipo = Tipo
If Vendas = False Then
    If Tipo = "P" Then TBComponente!Estoque = True Else TBComponente!Estoque = False
Else
    TBComponente!Estoque = True
End If
TBComponente!Processo = False
TBComponente!Observacoes = Obs
TBComponente!Espessura = IIf(Espessura = 0, Null, Espessura)
TBComponente!Largura = IIf(Largura = 0, Null, Largura)
TBComponente!Comprimento = IIf(Comprimento = 0, Null, Comprimento)
TBComponente!Dureza = Dureza
TBComponente!peso_metro = 0
TBComponente!Un_Kg = "N/a"
Set TBFamilia = CreateObject("adodb.recordset")
TBFamilia.Open "select ID_PC, ID_PC1, ID_CC from projfamilia where Familia = '" & Familia & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBFamilia.EOF = False Then
    If IsNull(TBFamilia!ID_PC) = False And TBFamilia!ID_PC <> "" Then TBComponente!ID_PC = TBFamilia!ID_PC
    If IsNull(TBFamilia!ID_PC1) = False And TBFamilia!ID_PC1 <> "" Then TBComponente!ID_PC1 = TBFamilia!ID_PC1
    If IsNull(TBFamilia!ID_CC) = False And TBFamilia!ID_CC <> "" Then TBComponente!ID_CC = TBFamilia!ID_CC
End If
TBFamilia.Close
TBComponente.Update
Codproduto = TBComponente!Codproduto

If CodRef <> "" Then
    Set TBItem = CreateObject("adodb.recordset")
    TBItem.Open "Select * from item_aplicacoes where n_referencia = '" & CodRef & "'", Conexao, adOpenKeyset, adLockReadOnly
    If TBItem.EOF = False Then
        Set TBProduto = CreateObject("adodb.recordset")
        TBProduto.Open "Select * from projproduto where codproduto = " & TBItem!Codproduto, Conexao, adOpenKeyset, adLockReadOnly
        If TBProduto.EOF = False Then
            If TBProduto!Desenho <> Codinterno Then
                If USMsgBox("Este cÛdigo de referÍncia est· sendo utilizado em outro(s) produto(s), deseja excluir para ser salvo no produto " & Codinterno & "?", vbYesNo) = vbYes Then
                    If USMsgBox("Deseja realmente excluir o cÛdigo de referÍncia " & CodRef & " no(s) outro(s) produto(s)?", vbYesNo) = vbYes Then
                        Conexao.Execute "DELETE from item_aplicacoes where  n_referencia = '" & CodRef & "'"
                    End If
                End If
            End If
        End If
        TBProduto.Close
    End If
    TBItem.Close
    
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select * from item_aplicacoes where codproduto = " & TBComponente!Codproduto, Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = True Then TBProduto.AddNew
    TBProduto!N_referencia = CodRef
    TBProduto!Codproduto = TBComponente!Codproduto
    TBProduto!Descricao = Descricao
    TBProduto!Aplicacao = CliFornCodRef
    TBProduto!ID_cliente_forn = IDCliFornCodRef
    TBProduto!Tipo = TipoCliFornCodRef
    TBProduto.Update
    TBProduto.Close
End If
FunCriaNovoProdServ = Desenho
TBComponente.Close

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Function

Sub ProcCarregaComboEmpresa(Combo As ComboBox, CarregarTodas As Boolean)
On Error GoTo tratar_erro

With Combo
    .Clear
    Set TBCarregarCombo = CreateObject("adodb.recordset")
    TBCarregarCombo.Open "Select Empresa, CODIGO, NF_Serie from Empresa order by codigo, Empresa", Conexao, adOpenKeyset, adLockReadOnly
    If TBCarregarCombo.EOF = False Then
        If CarregarTodas = True Then .AddItem "Todas"
        Do While TBCarregarCombo.EOF = False
            If IsNull(TBCarregarCombo!Empresa) = False And TBCarregarCombo!Empresa <> "" Then
                .AddItem TBCarregarCombo!Empresa
                .ItemData(.NewIndex) = TBCarregarCombo!CODIGO
            End If
            TBCarregarCombo.MoveNext
        Loop
        TBCarregarCombo.MoveFirst
         NF_Serie = TBCarregarCombo!NF_Serie
        .Text = TBCarregarCombo!Empresa
   End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Sub ProcPuxaDadosComboEmpresa(Combo As ComboBox, ID_empresa As Integer)
On Error GoTo tratar_erro

Set TBCarregarCombo = CreateObject("adodb.recordset")
TBCarregarCombo.Open "Select * from Empresa where Codigo = " & ID_empresa, Conexao, adOpenKeyset, adLockOptimistic
If TBCarregarCombo.EOF = False Then
    Combo = TBCarregarCombo!Empresa
End If

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaComboFamilia(Combo As ComboBox, TextoFiltro As String, CarregarVazio As Boolean)
On Error GoTo tratar_erro

With Combo
    .Clear
    Set TBCarregarCombo = CreateObject("adodb.recordset")
    TBCarregarCombo.Open "Select CODIGO, Familia from Projfamilia where " & TextoFiltro & " group by CODIGO, Familia order by Familia", Conexao, adOpenKeyset, adLockOptimistic
    If TBCarregarCombo.EOF = False Then
        If CarregarVazio = True Then .AddItem ""
        Do While TBCarregarCombo.EOF = False
            If IsNull(TBCarregarCombo!Familia) = False And TBCarregarCombo!Familia <> "" Then
                .AddItem TBCarregarCombo!Familia
                .ItemData(.NewIndex) = TBCarregarCombo!CODIGO
            End If
            TBCarregarCombo.MoveNext
        Loop
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaComboGrupoFamilia(Combo As ComboBox, TextoFiltro As String, CarregarVazio As Boolean)
On Error GoTo tratar_erro

With Combo
    .Clear
    Set TBCarregarCombo = CreateObject("adodb.recordset")
    TBCarregarCombo.Open "Select Grupo from Projfamilia where " & TextoFiltro & " group by Grupo", Conexao, adOpenKeyset, adLockOptimistic
    If TBCarregarCombo.EOF = False Then
        If CarregarVazio = True Then .AddItem ""
        Do While TBCarregarCombo.EOF = False
            If IsNull(TBCarregarCombo!Grupo) = False And TBCarregarCombo!Grupo <> "" Then .AddItem TBCarregarCombo!Grupo
            TBCarregarCombo.MoveNext
        Loop
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaComboOperacao(Combo As ComboBox, TextoFiltro As String, CarregarVazio As Boolean)
On Error GoTo tratar_erro

With Combo
    .Clear
    Set TBCarregarCombo = CreateObject("adodb.recordset")
    TBCarregarCombo.Open "Select Status from Estoque_controle  group by Status Order By Status", Conexao, adOpenKeyset, adLockOptimistic
    If TBCarregarCombo.EOF = False Then
        If CarregarVazio = True Then .AddItem ""
        Do While TBCarregarCombo.EOF = False
            If IsNull(TBCarregarCombo!status) = False And TBCarregarCombo!status <> "" Then .AddItem TBCarregarCombo!status
            TBCarregarCombo.MoveNext
        Loop
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Sub ProcCarregaComboUnidade(Combo As ComboBox, BloquearUnConversao As Boolean)
On Error GoTo tratar_erro

With Combo
    .Clear
    If BloquearUnConversao = True Then TextoFiltro = " and Unidade <> 'KG' and Unidade <> 'MT' and Unidade <> 'MM'" Else TextoFiltro = ""
    Set TBCarregarCombo = CreateObject("adodb.recordset")
    TBCarregarCombo.Open "Select Unidade, Codigo from Unidade_Medida where Unidade IS NOT NULL " & TextoFiltro & " group by Unidade, Codigo", Conexao, adOpenKeyset, adLockOptimistic
    If TBCarregarCombo.EOF = False Then
        Do While TBCarregarCombo.EOF = False
            .AddItem TBCarregarCombo!Unidade
            .ItemData(.NewIndex) = TBCarregarCombo!CODIGO
            TBCarregarCombo.MoveNext
        Loop
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaComboUsuario(Combo As ComboBox, TextoFiltro As String, CarregarVazio As Boolean)
On Error GoTo tratar_erro

With Combo
    .Clear
    Set TBCarregarCombo = CreateObject("adodb.recordset")
    TBCarregarCombo.Open "Select Usuario from Usuarios U LEFT JOIN Acessos A ON A.IDUsuario = U.IDUsuario where " & TextoFiltro & " group by U.usuario", Conexao, adOpenKeyset, adLockOptimistic
    If TBCarregarCombo.EOF = False Then
        If CarregarVazio = True Then .AddItem ""
        
        Do While TBCarregarCombo.EOF = False
            .AddItem TBCarregarCombo!Usuario
            TBCarregarCombo.MoveNext
        Loop
    End If
    TBCarregarCombo.Close
End With

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaComboTranspMarcaEspecie(Combo As ComboBox, Tipo As String)
On Error GoTo tratar_erro

With Combo
    .Clear
    Set TBCarregarCombo = CreateObject("adodb.recordset")
    TBCarregarCombo.Open "Select Descricao from Embalagem_Marca_Especie where Tipo = '" & Tipo & "' group by Descricao", Conexao, adOpenKeyset, adLockOptimistic
    If TBCarregarCombo.EOF = False Then
        .AddItem ""
        Do While TBCarregarCombo.EOF = False
            .AddItem TBCarregarCombo!Descricao
            '.ItemData(.NewIndex) = TBCarregarCombo!ID
            TBCarregarCombo.MoveNext
        Loop
    End If
    TBCarregarCombo.Close
End With

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaComboPostoTrab(Combo As ComboBox, TextoFiltro As String, CarregarVazio As Boolean, CarregarDescricao As Boolean)
On Error GoTo tratar_erro

With Combo
    .Clear
    Set TBCarregarCombo = CreateObject("adodb.recordset")
    TBCarregarCombo.Open "Select IDMaquina, Maquina, Descricao from CadMaquinas where " & TextoFiltro & " Group by IDMaquina, Maquina, Descricao", Conexao, adOpenKeyset, adLockOptimistic
    If TBCarregarCombo.EOF = False Then
        If CarregarVazio = True Then .AddItem ""
        Do While TBCarregarCombo.EOF = False
            If CarregarDescricao = True Then .AddItem TBCarregarCombo!maquina & " - " & TBCarregarCombo!Descricao Else .AddItem TBCarregarCombo!maquina
            .ItemData(.NewIndex) = TBCarregarCombo!IDMaquina
            TBCarregarCombo.MoveNext
        Loop
    End If
    TBCarregarCombo.Close
End With

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaComboCodigoTrab(Combo As ComboBox, TextoFiltro As String, CarregarVazio As Boolean)
On Error GoTo tratar_erro

With Combo
    .Clear
    Set TBCarregarCombo = CreateObject("adodb.recordset")
    TBCarregarCombo.Open "Select * from CodigoDesc where " & TextoFiltro & " order by Descricao", Conexao, adOpenKeyset, adLockOptimistic
    If TBCarregarCombo.EOF = False Then
        If CarregarVazio = True Then .AddItem ""
        
        Do While TBCarregarCombo.EOF = False
            .AddItem TBCarregarCombo!Descricao
            .ItemData(.NewIndex) = TBCarregarCombo!CODIGO
            TBCarregarCombo.MoveNext
        Loop
    End If
    TBCarregarCombo.Close
End With

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaComboFuncionario(Combo As ComboBox, TextoFiltro As String, CarregarVazio As Boolean)
On Error GoTo tratar_erro

With Combo
    .Clear
    Set TBCarregarCombo = CreateObject("adodb.recordset")
    TBCarregarCombo.Open "Select * from Funcionarios where " & TextoFiltro & " order by Nome", Conexao, adOpenKeyset, adLockOptimistic
    If TBCarregarCombo.EOF = False Then
        If CarregarVazio = True Then .AddItem ""
        
        Do While TBCarregarCombo.EOF = False
            .AddItem TBCarregarCombo!Nome
            .ItemData(.NewIndex) = TBCarregarCombo!ID
            TBCarregarCombo.MoveNext
        Loop
    End If
    TBCarregarCombo.Close
End With

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaComboSetor(Combo As ComboBox, TextoFiltro As String, Codinterno As String, VerifCCProd As Boolean, CarregarVazio As Boolean, VerifRespCC As Boolean, Responsavel_CC As String, CarregarCodigo As Boolean, VerificaPostoTrab As Boolean)
On Error GoTo tratar_erro

With Combo
    .Clear
    Set TBCarregarCombo = CreateObject("adodb.recordset")
    If VerifRespCC = True Then
        TBCarregarCombo.Open "Select US.ID, US.Codigo, US.Setor from Usuarios_Setor_Responsavel USR INNER JOIN Usuarios_setor US ON USR.ID_CC = US.ID where USR.Responsavel_CC = '" & Responsavel_CC & "' and " & TextoFiltro & " group by US.ID, US.Codigo, US.Setor order by US.Codigo", Conexao, adOpenKeyset, adLockReadOnly
    ElseIf VerificaPostoTrab = True Then
            TBCarregarCombo.Open "Select US.ID, US.Codigo, US.Setor from Cadmaquinas CM INNER JOIN Usuarios_setor US ON CM.Setor = US.Setor where " & TextoFiltro & " group by US.ID, US.Codigo, US.Setor order by US.Codigo", Conexao, adOpenKeyset, adLockReadOnly
        Else
            TBCarregarCombo.Open "Select * from Usuarios_Setor US where " & TextoFiltro & " order by Codigo", Conexao, adOpenKeyset, adLockOptimistic
    End If
    If TBCarregarCombo.EOF = False Then
        If CarregarVazio = True Then .AddItem ""
        Do While TBCarregarCombo.EOF = False
            Permitido = True
            If VerifCCProd = True Then
                Set TBCQ = CreateObject("adodb.recordset")
                TBCQ.Open "Select * from projproduto where Desenho = '" & Codinterno & "' and ID_CC = " & TBCarregarCombo!ID, Conexao, adOpenKeyset, adLockOptimistic
                If TBCQ.EOF = False Then Permitido = False
                TBCQ.Close
            End If
            
            If Permitido = True Then
                If IsNull(TBCarregarCombo!CODIGO) = False And TBCarregarCombo!CODIGO <> "" Then
                    If CarregarCodigo = True Then .AddItem TBCarregarCombo!CODIGO & " - " & TBCarregarCombo!Setor Else .AddItem TBCarregarCombo!Setor
                Else
                    .AddItem TBCarregarCombo!Setor
                End If
                .ItemData(.NewIndex) = TBCarregarCombo!ID
            End If
            TBCarregarCombo.MoveNext
        Loop
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaComboSetorPT(Combo As ComboBox, CarregarVazio As Boolean)
On Error GoTo tratar_erro

With Combo
    .Clear
    Set TBCarregarCombo = CreateObject("adodb.recordset")
    TBCarregarCombo.Open "Select Setor from CadMaquinas where Setor IS NOT NULL and Bloqueado = 'False' group by Setor", Conexao, adOpenKeyset, adLockReadOnly
    If TBCarregarCombo.EOF = False Then
        If CarregarVazio = True Then .AddItem ""
        Do While TBCarregarCombo.EOF = False
            .AddItem TBCarregarCombo!Setor
            TBCarregarCombo.MoveNext
        Loop
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaComboGrupoPT(Combo As ComboBox, CarregarVazio As Boolean)
On Error GoTo tratar_erro

With Combo
    .Clear
    Set TBCarregarCombo = CreateObject("adodb.recordset")
    TBCarregarCombo.Open "Select Grupo from CadMaquinas where Grupo IS NOT NULL and Bloqueado = 'False' group by Grupo", Conexao, adOpenKeyset, adLockReadOnly
    If TBCarregarCombo.EOF = False Then
        If CarregarVazio = True Then .AddItem ""
        Do While TBCarregarCombo.EOF = False
            .AddItem TBCarregarCombo!Grupo
            TBCarregarCombo.MoveNext
        Loop
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaComboAno(Combo As ComboBox, Ano_inicial As Integer, Ano_final As Integer)
On Error GoTo tratar_erro

With Combo
    .Clear
    Do While Ano_inicial <> (Year(Date) + Ano_final)
        .AddItem Ano_inicial
        Ano_inicial = Ano_inicial + 1
    Loop
    .Text = Year(Date)
End With

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaComboCodRef(Combo As ComboBox, TextoFiltro As String, IDCliForn As Long, Tipo As String, FiltrarCliForn As Boolean, CarregarCampo As Boolean)
On Error GoTo tratar_erro

With Combo
    .Clear
    Set TBCarregarCombo = CreateObject("adodb.recordset")
    If FiltrarCliForn = True Then
        TBCarregarCombo.Open "Select IA.N_Referencia from item_aplicacoes IA INNER JOIN projproduto P ON IA.codproduto = P.codproduto where " & TextoFiltro & " and IA.ID_cliente_forn = " & IDCliForn & " and IA.Tipo = '" & Tipo & "' and IA.N_Referencia is not null group by IA.n_referencia", Conexao, adOpenKeyset, adLockOptimistic
        If TBCarregarCombo.EOF = True Then
            Set TBCarregarCombo = CreateObject("adodb.recordset")
            TBCarregarCombo.Open "Select IA.N_Referencia from item_aplicacoes IA INNER JOIN projproduto P ON IA.codproduto = P.codproduto where " & TextoFiltro & " and (IA.ID_cliente_forn = 0 or IA.ID_cliente_forn IS NULL) and IA.N_Referencia is not null group by IA.n_referencia", Conexao, adOpenKeyset, adLockOptimistic
        End If
    Else
        TBCarregarCombo.Open "Select IA.N_Referencia from item_aplicacoes IA INNER JOIN projproduto P ON IA.codproduto = P.codproduto where " & TextoFiltro & " and IA.N_Referencia is not null group by IA.n_referencia", Conexao, adOpenKeyset, adLockOptimistic
    End If
    If TBCarregarCombo.EOF = False Then
        .AddItem ""
        Do While TBCarregarCombo.EOF = False
            If TBCarregarCombo!N_referencia <> "" Then .AddItem TBCarregarCombo!N_referencia
            TBCarregarCombo.MoveNext
        Loop
        If CarregarCampo = True Then
            TBCarregarCombo.MoveFirst
            .Text = TBCarregarCombo!N_referencia
        End If
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaComboPartNumberFab(Combo As ComboBox, TextoFiltro As String)
On Error GoTo tratar_erro

With Combo
    .Clear
    Set TBCarregarCombo = CreateObject("adodb.recordset")
    StrSql = "Select PF.* from Projproduto_fabricante PF INNER JOIN Projproduto P ON P.Codproduto = PF.Codproduto where " & TextoFiltro & " order by PF.Part_number"
    'Debug.print StrSql
    
    TBCarregarCombo.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
    If TBCarregarCombo.EOF = False Then
        Do While TBCarregarCombo.EOF = False
            If TBCarregarCombo!Part_number <> "" Then
                .AddItem TBCarregarCombo!Part_number
                .ItemData(.NewIndex) = TBCarregarCombo!ID
                
                If TBCarregarCombo.RecordCount = 1 Then .Text = TBCarregarCombo!Part_number
            End If
            TBCarregarCombo.MoveNext
        Loop
    End If
    TBCarregarCombo.Close
End With

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Sub ProcCarregaComboCF(Combo As ComboBox, TextoFiltro As String, CarregarVazio As Boolean)
On Error GoTo tratar_erro

With Combo
    .Clear
    Set TBCarregarCombo = CreateObject("adodb.recordset")
    TBCarregarCombo.Open "Select txt_ref from tbl_classificacaofiscal where txt_ref is not null group by txt_ref", Conexao, adOpenKeyset, adLockOptimistic
    If TBCarregarCombo.EOF = False Then
        If CarregarVazio = True Then .AddItem ""
        
        Do While TBCarregarCombo.EOF = False
            If TBCarregarCombo!txt_ref <> "" Then .AddItem TBCarregarCombo!txt_ref
            TBCarregarCombo.MoveNext
        Loop
        TBCarregarCombo.MoveFirst
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaComboLA(Combo As ComboBox, CarregarVazio As Boolean, CarregarBloq As Boolean)
On Error GoTo tratar_erro

With Combo
    TextoFiltro = ""
    If CarregarBloq = False Then TextoFiltro = " and DtValidacao IS NOT NULL and Status = 'Liberado'"
    
    .Clear
    Set TBCarregarCombo = CreateObject("adodb.recordset")
    TBCarregarCombo.Open "Select * from Estoque_Localarmazenamento_criar where Descricao is not null " & TextoFiltro & " order by Descricao", Conexao, adOpenKeyset, adLockOptimistic
    If TBCarregarCombo.EOF = False Then
        If CarregarVazio = True Then .AddItem ""
        .AddItem "ESTOQUE PADR√O"
        Do While TBCarregarCombo.EOF = False
            .AddItem TBCarregarCombo!Descricao
            .ItemData(.NewIndex) = TBCarregarCombo!ID
            TBCarregarCombo.MoveNext
        Loop
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaComboBancoFinanceiro(Combo As ComboBox, TextoFiltro As String, CarregarVazio As Boolean)
On Error GoTo tratar_erro
StrSql = "Select * from tbl_Instituicoes where " & TextoFiltro & " order by txt_Descricao"
'Debug.print StrSql

With Combo
    .Clear
    Set TBCarregarCombo = CreateObject("adodb.recordset")
    TBCarregarCombo.Open "Select * from tbl_Instituicoes where " & TextoFiltro & " order by txt_Descricao", Conexao, adOpenKeyset, adLockOptimistic
    If TBCarregarCombo.EOF = False Then
        If CarregarVazio = True Then .AddItem ""
        Do While TBCarregarCombo.EOF = False
            If IsNull(TBCarregarCombo!Txt_descricao) = False And TBCarregarCombo!Txt_descricao <> "" Then
                .AddItem TBCarregarCombo!Txt_descricao
                .ItemData(.NewIndex) = TBCarregarCombo!ID
            End If
            TBCarregarCombo.MoveNext
        Loop
        TBCarregarCombo.MoveFirst
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaComboFormaPgtoRcbto(Combo As ComboBox, TextoFiltro As String)
On Error GoTo tratar_erro

With Combo
    .Clear
    Set TBCarregarCombo = CreateObject("adodb.recordset")
    TBCarregarCombo.Open "Select * from tbl_FormaPagto where " & TextoFiltro & " order by Descricao", Conexao, adOpenKeyset, adLockOptimistic
    If TBCarregarCombo.EOF = False Then
        Do While TBCarregarCombo.EOF = False
            If IsNull(TBCarregarCombo!Descricao) = False And TBCarregarCombo!Descricao <> "" Then
                .AddItem TBCarregarCombo!Descricao
                '.ItemData(.NewIndex) = TBCarregarCombo!IdForma
            End If
            TBCarregarCombo.MoveNext
        Loop
        TBCarregarCombo.MoveFirst
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaComboTipoDocto(Combo As ComboBox, TextoFiltro As String)
On Error GoTo tratar_erro

With Combo
    .Clear
    Set TBCarregarCombo = CreateObject("adodb.recordset")
    TBCarregarCombo.Open "Select * from tbl_ContasPagar_Tipo_Docto where " & TextoFiltro & " order by Descricao", Conexao, adOpenKeyset, adLockOptimistic
    If TBCarregarCombo.EOF = False Then
        .AddItem ""
        Do While TBCarregarCombo.EOF = False
            If IsNull(TBCarregarCombo!Descricao) = False And TBCarregarCombo!Descricao <> "" Then
                .AddItem TBCarregarCombo!Descricao
                '.ItemData(.NewIndex) = TBCarregarCombo!IdForma
            End If
            TBCarregarCombo.MoveNext
        Loop
        TBCarregarCombo.MoveFirst
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaComboUF(Combo As ComboBox, TextoFiltro As String, Origem As String)
On Error GoTo tratar_erro

With Combo
    .Clear
    .AddItem ""
    If Origem = "Estrangeiro" Then
        .AddItem "EX"
        .Text = "EX"
    Else
        Set TBCarregarCombo = CreateObject("adodb.recordset")
        TBCarregarCombo.Open "Select * from Regioes where " & TextoFiltro & " order by UF", Conexao, adOpenKeyset, adLockOptimistic
        If TBCarregarCombo.EOF = False Then
            If Origem = "" Then .AddItem "EX"
            
            Do While TBCarregarCombo.EOF = False
                If IsNull(TBCarregarCombo!UF) = False And TBCarregarCombo!UF <> "" Then
                    .AddItem TBCarregarCombo!UF
                    '.ItemData(.NewIndex) = TBCarregarCombo!IdForma
                End If
                TBCarregarCombo.MoveNext
            Loop
            TBCarregarCombo.MoveFirst
        End If
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaComboCidade(Combo As ComboBox, TextoFiltro As String, CarregarVazio As Boolean)
On Error GoTo tratar_erro

With Combo
    .Clear
    Set TBCarregarCombo = CreateObject("adodb.recordset")
    TBCarregarCombo.Open "Select * from CEP where " & TextoFiltro & " order by Municipio", Conexao, adOpenKeyset, adLockOptimistic
    If TBCarregarCombo.EOF = False Then
        If CarregarVazio = True Then .AddItem ""
        
        Do While TBCarregarCombo.EOF = False
            .AddItem UCase(TBCarregarCombo!municipio)
            TBCarregarCombo.MoveNext
        Loop
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaComboGrupoCliente(Combo As ComboBox, CarregarVazio As Boolean)
On Error GoTo tratar_erro

With Combo
    .Clear
    Set TBCarregarCombo = CreateObject("adodb.recordset")
    TBCarregarCombo.Open "Select CG.ID, CG.Texto from Clientes_grupos CG INNER JOIN Clientes C ON C.IDGrupo = CG.ID group by CG.ID, CG.Texto order by CG.Texto", Conexao, adOpenKeyset, adLockOptimistic
    If TBCarregarCombo.EOF = False Then
        If CarregarVazio = True Then .AddItem ""
        
        Do While TBCarregarCombo.EOF = False
            .AddItem TBCarregarCombo!Texto
            .ItemData(.NewIndex) = TBCarregarCombo!ID
            TBCarregarCombo.MoveNext
        Loop
    End If
    TBCarregarCombo.Close
End With

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaComboCliForn(Combo As ComboBox, Cliente As Boolean)
On Error GoTo tratar_erro

With Combo
    .Clear
    If Cliente = True Then
        NomeCampo = "NomeRazao"
        NomeTabela = "Clientes"
    Else
        NomeCampo = "Nome_Razao"
        NomeTabela = "Compras_fornecedores"
    End If
    Set TBCarregarCombo = CreateObject("adodb.recordset")
    TBCarregarCombo.Open "Select IDCliente, " & NomeCampo & " as NomeRazao from " & NomeTabela & " where " & NomeCampo & " IS NOT NULL order by " & NomeCampo & "", Conexao, adOpenKeyset, adLockOptimistic
    If TBCarregarCombo.EOF = False Then
        Do While TBCarregarCombo.EOF = False
            .AddItem Trim(TBCarregarCombo!NomeRazao)
            .ItemData(.NewIndex) = TBCarregarCombo!IDCliente
            TBCarregarCombo.MoveNext
        Loop
    End If
    TBCarregarCombo.Close
End With

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaComboFontes(Combo As ComboBox)
On Error GoTo tratar_erro

With Combo
    .Clear
    For i = 0 To Screen.FontCount - 1
        .AddItem Screen.Fonts(i)
    Next i
    .Text = "Arial"
End With

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaComboTamanhoFonte(Combo As ComboBox, De As Double, Ate As Double)
On Error GoTo tratar_erro

valor = "0,25"
With Combo
    .Clear
    For Qtde = De To Ate Step valor
         .AddItem Qtde
    Next Qtde
    .ListIndex = 0
End With

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaComboModulos(Combo As ComboBox, CarregarVazio As Boolean, Tipo As String)
On Error GoTo tratar_erro

With Combo
    .Clear
    If CarregarVazio = True Then .AddItem ""
    
    If Tipo = "InstalaÁ„o" Or Tipo = "ConfiguraÁ„o" Then
        .AddItem "Caprind"
        .AddItem "Gerprod"
        .AddItem "GNFE"
        .AddItem "Caprind e GNFe"
        .AddItem "SQL Server"
    ElseIf Tipo = "ReindexaÁ„o" Then
            .AddItem "Banco de dados"
        Else
            If Tipo <> "D˙vida" And Tipo <> "" Then .AddItem "Novo mÛdulo"
                
            If Tipo <> "" Then
                .AddItem "Todos os mÛdulos necess·rios"
                .AddItem "Gerprod"
                .AddItem "GNFE"
                .AddItem "Aviso di·rio"
                .AddItem "Menu"
            End If
                    
            .AddItem "ConfiguraÁ„o do sistema/OpÁıes gerais"
                .AddItem "ConfiguraÁ„o do sistema/OpÁıes gerais/ConfiguraÁ„o do sistema"
                .AddItem "ConfiguraÁ„o do sistema/OpÁıes gerais/Cadastro de empresa"
                .AddItem "ConfiguraÁ„o do sistema/OpÁıes gerais/Cadastro de moedas"
                .AddItem "ConfiguraÁ„o do sistema/OpÁıes gerais/Cadastro de unidades"
                .AddItem "ConfiguraÁ„o do sistema/OpÁıes gerais/Cadastro de condiÁ„o de pagamento/recebimento"
                .AddItem "ConfiguraÁ„o do sistema/OpÁıes gerais/Cadastro de feriados"
                
            .AddItem "ConfiguraÁ„o do sistema/Usu·rios"
            .AddItem "ConfiguraÁ„o do sistema/Usu·rios/Eventos realizados"
            .AddItem "ConfiguraÁ„o do sistema/Usu·rios/Conectados"
            .AddItem "ConfiguraÁ„o do sistema/Criar backup/ConfiguraÁıes"
            .AddItem "ConfiguraÁ„o do sistema/Criar backup/Apontamentos"
            .AddItem "ConfiguraÁ„o do sistema/Criar backup/Eventos"
            .AddItem "ConfiguraÁ„o do sistema/Reindexar BD/Caprind e Gerprod"
            .AddItem "ConfiguraÁ„o do sistema/Reindexar BD/GNFe"
                
            .AddItem "RH/Funcion·rios"
            .AddItem "RH/RelatÛrios/DesoneraÁ„o da folha de pagamento"
        
            .AddItem "Compras/FamÌlias"
            .AddItem "Compras/Produtos e serviÁos"
                .AddItem "Compras/Produtos e serviÁos/Valores e descontos"
                .AddItem "Compras/Produtos e serviÁos/Clientes e fornecedores"
                .AddItem "Compras/Produtos e serviÁos/Validar estrutura"
                .AddItem "Compras/Produtos e serviÁos/Validar plano de inspeÁ„o"

            .AddItem "Compras/Fornecedores"
            .AddItem "Compras/ProgramaÁ„o"
            .AddItem "Compras/CotaÁ„o"
                .AddItem "Compras/CotaÁ„o/Liberar cotaÁ„o"
            .AddItem "Compras/Pedido"
            .AddItem "Compras/Pedido/Aprovar"
            .AddItem "Outros/Follow up de compras"
            .AddItem "Compras/Necessidade"
            .AddItem "Compras/N„o conformidade"
            .AddItem "Compras/AtualizaÁ„o de valores"
            .AddItem "Compras/RelatÛrios/HistÛrico"
            .AddItem "Compras/RelatÛrios/Õndice de atraso"
            .AddItem "Compras/AutorizaÁ„o de centro de custo sem previs„o"

            .AddItem "Vendas/FamÌlias"
            .AddItem "Vendas/Produtos e serviÁos"
                .AddItem "Vendas/Produtos e serviÁos/Valores e descontos"
                    .AddItem "Vendas/Produtos e serviÁos/Valores e descontos/Visualizar valor de custo"
                .AddItem "Vendas/Produtos e serviÁos/Clientes e fornecedores"
                .AddItem "Vendas/Produtos e serviÁos/Validar estrutura"
                .AddItem "Vendas/Produtos e serviÁos/Validar plano de inspeÁ„o"
                
            .AddItem "Vendas/Clientes"
            .AddItem "Vendas/Vendedores"
            .AddItem "Vendas/Telemarketing"
            .AddItem "Vendas/SimulaÁ„o"
            .AddItem "Vendas/Empenho"
            .AddItem "Vendas/ProgramaÁ„o"
            .AddItem "Vendas/Proposta comercial"
            .AddItem "Vendas/Pedido interno"
            .AddItem "Vendas/Follow up"
            .AddItem "Vendas/SituaÁ„o da produÁ„o"
            .AddItem "Vendas/PÛs-vendas/AssistÍncia tÈcnica"
            .AddItem "Vendas/Controle de vendas"
            .AddItem "Vendas/InformaÁıes faturamento"
            .AddItem "Vendas/AtualizaÁ„o de valores"
                .AddItem "Vendas/RelatÛrios/HistÛrico"
                .AddItem "Vendas/RelatÛrios/Õndice de atraso"
                .AddItem "Vendas/RelatÛrios/Comiss„o"
        
            .AddItem "Financeiro/Plano de contas"
            .AddItem "Financeiro/InstituiÁıes"
            .AddItem "Financeiro/Contas a pagar"
                .AddItem "Financeiro/Contas a pagar/Visualizar contas dos funcion·rios"
            .AddItem "Financeiro/Contas pagas"
                .AddItem "Financeiro/Contas pagas/Visualizar contas dos funcion·rios"
            .AddItem "Financeiro/Contas a receber"
            .AddItem "Financeiro/Contas recebidas"
            .AddItem "Financeiro/Desconto de duplicata"
            .AddItem "Financeiro/Fluxo de caixa"
            .AddItem "Financeiro/RelatÛrios/HistÛrico"
            .AddItem "Financeiro/RelatÛrios/Raz„o"
            .AddItem "Financeiro/AutorizaÁ„o de centro de custo sem previs„o"
        
            .AddItem "Faturamento/Fiscal/ClassificaÁ„o fiscal"
            .AddItem "Faturamento/Fiscal/Natureza de operaÁ„o"
            .AddItem "Faturamento/Nota fiscal/Terceiros"
            .AddItem "Faturamento/Nota fiscal/PrÛpria"
                .AddItem "Faturamento/Nota fiscal/Cancelar nota"
                .AddItem "Faturamento/Nota fiscal/Excluir duplicatas"
            .AddItem "Faturamento/Nota fiscal/SPED"
            .AddItem "Faturamento/Nota fiscal/Exportar"
            .AddItem "Faturamento/Carta de correÁ„o"
            .AddItem "Faturamento/Minuta de despacho"
            .AddItem "Faturamento/RelatÛrios/HistÛrico"
            .AddItem "Faturamento/RelatÛrios/Relacionamento de notas fiscais"
            .AddItem "Faturamento/RelatÛrios/Impostos"
            .AddItem "Faturamento/RelatÛrios/Doze ˙ltimos meses"
            .AddItem "Faturamento/AutorizaÁ„o de centro de custo sem previs„o"
            
            .AddItem "Custos/Centro de custo"
                .AddItem "Custos/Centro de custo/Visualizar todos"
                .AddItem "Custos/Centro de custo/Visualizar lanÁamentos realizados"
            .AddItem "Custos/RelatÛrios/Detalhado"
            .AddItem "Custos/RelatÛrios/Resumido"
            .AddItem "Custos/RelatÛrios/Previsto x Realizado"
        
            .AddItem "Engenharia/FamÌlias"
            .AddItem "Engenharia/Produtos e serviÁos"
                .AddItem "Engenharia/Produtos e serviÁos/Validar estrutura"
                .AddItem "Engenharia/Produtos e serviÁos/Validar plano de inspeÁ„o"
            .AddItem "Engenharia/Conjuntos"
            .AddItem "Engenharia/Estrutura/Resumida"
            .AddItem "Engenharia/Estrutura/Detalhada"
                .AddItem "Engenharia/Estrutura/Visualizar valor de custo"
            .AddItem "Engenharia/Controle de projetos"
            .AddItem "Engenharia/Processos"
            .AddItem "Engenharia/Processos/HistÛrico"
            .AddItem "Engenharia/Normas"
        
            .AddItem "PCP/Postos de trabalho"
            .AddItem "PCP/CÛdigos de trabalho"
            .AddItem "PCP/Carga de posto de trabalho"
            .AddItem "PCP/Gerenciamento de ordem"
                .AddItem "PCP/Gerenciamento de ordem/Validar resultados"
            .AddItem "PCP/Monitor de trabalho"
            .AddItem "PCP/Programas CNC"
            .AddItem "PCP/SituaÁ„o da produÁ„o"
            .AddItem "PCP/Necessidade"
            .AddItem "PCP/N„o conformidade"
            .AddItem "PCP/ProgramaÁ„o da produÁ„o"
            .AddItem "PCP/Plano da produÁ„o"
            .AddItem "PCP/RequisiÁ„o da ordem"
                .AddItem "PCP/RelatÛrios/Produtividade"
                .AddItem "PCP/RelatÛrios/N„o conformidade"
                .AddItem "PCP/RelatÛrios/Monitor de eventos"
                .AddItem "PCP/RelatÛrios/Õndice de atraso"
                .AddItem "PCP/RelatÛrios/Resultados da ordem"
            
            .AddItem "Qualidade/FamÌlias"
            .AddItem "Qualidade/Instrumentos"
            .AddItem "Qualidade/Almoxarifado"
            .AddItem "Qualidade/Plano de inspeÁ„o"
            .AddItem "Qualidade/Controle de mediÁ„o"
            .AddItem "Qualidade/InspeÁ„o de recebimento"
            .AddItem "Qualidade/Ensaios/Ultra-som"
            .AddItem "Qualidade/Ensaios/LÌquido penetrante"
            .AddItem "Qualidade/Ensaios/Controle de certificados"
            .AddItem "Qualidade/Controle de certificados"
            .AddItem "Qualidade/N„o conformidade"
                .AddItem "Qualidade/N„o conformidade/DescriÁ„o da n„o conformidade"
                .AddItem "Qualidade/N„o conformidade"
            
            .AddItem "Qualidade/SolicitaÁ„o de aÁ„o"
            .AddItem "Qualidade/SolicitaÁ„o de desvio"
            .AddItem "Qualidade/RNC"
            .AddItem "Qualidade/RelatÛrios/N„o conformidade"
            .AddItem "Qualidade/RelatÛrios/HistÛrico"
            .AddItem "Qualidade/PPAP"
                .AddItem "Qualidade/PPAP/PSW"
                .AddItem "Qualidade/PPAP/FMEA"
                .AddItem "Qualidade/PPAP/Plano de controle"
            .AddItem "Qualidade/HistÛrico de revis„o dos relatÛrios"
                    
            .AddItem "Estoque/Almoxarifado"
            .AddItem "Estoque/Local de armazenamento"
            .AddItem "Estoque/RequisiÁ„o de materiais"
            .AddItem "Estoque/Recebimento/Pedido de compra"
            .AddItem "Estoque/Recebimento/ConsignaÁ„o"
            .AddItem "Estoque/Recebimento/Importar nota de terceiros"
            
            .AddItem "Estoque/Invent·rio"
            .AddItem "Estoque/MovimentaÁ„o"
                .AddItem "Estoque/MovimentaÁ„o/Entrada"
                .AddItem "Estoque/MovimentaÁ„o/Retirada"
            .AddItem "Estoque/Necessidade"
            .AddItem "Estoque/Ordem de faturamento"
            .AddItem "Estoque/Nota fiscal"
            .AddItem "Estoque/AutorizaÁ„o de centro de custo sem previs„o"
        
            .AddItem "ManutenÁ„o/Equipamentos"
                .AddItem "ManutenÁ„o/Equipamentos/Aprovar manutenÁ„o"
            .AddItem "ManutenÁ„o/AssistÍncia tÈcnica"
            .AddItem "ManutenÁ„o/RelatÛrios/HistÛrico"
            
            .AddItem "Suporte/Chamado"
            .AddItem "Suporte/Chat (online)"
            .AddItem "Suporte/SolicitaÁ„o de atendimento"
            .AddItem "Suporte/AtualizaÁ„o/Caprind e Gerprod"
            .AddItem "Suporte/AtualizaÁ„o/GNFe"
            .AddItem "Suporte/AtualizaÁ„o/GMRE (relatÛrios)"
            
            .AddItem "Outros/An·lise crÌtica"
                .AddItem "Outros/An·lise crÌtica/Engenharia"
                .AddItem "Outros/An·lise crÌtica/Processos"
                .AddItem "Outros/An·lise crÌtica/Pcp"
                .AddItem "Outros/An·lise crÌtica/Qualidade"
                .AddItem "Outros/An·lise crÌtica/Compras"
                .AddItem "Outros/An·lise crÌtica/Vendas"
                .AddItem "Outros/An·lise crÌtica/Documentos"
            .AddItem "Outros/SolicitaÁ„o"
                .AddItem "Outros/SolicitaÁ„o/Autorizar solicitaÁ„o"
            .AddItem "Outros/SolicitaÁ„o de produÁ„o"
                .AddItem "Outros/SolicitaÁ„o de produÁ„o/Autorizar solicitaÁ„o"
            .AddItem "Outros/ValidaÁ„o de procedimentos"
            .AddItem "Qualidade/Controle de documentos e dados"
            
            .AddItem "Outros/Downloads/Nota fiscal"
            .AddItem "Outros/Downloads/Boleto"
            
            .AddItem "Avisos di·rio/SolicitaÁ„o"
            .AddItem "Avisos di·rio/Contas a pagar"
            .AddItem "Avisos di·rio/Contas a receber"
            .AddItem "Avisos di·rio/ManutenÁ„o"
            .AddItem "Avisos di·rio/RequisiÁ„o de materiais"
            .AddItem "Avisos di·rio/Compras/Necessidade"
            .AddItem "Avisos di·rio/PCP/Necessidade"
            .AddItem "Avisos di·rio/Estoque/Necessidade"
            .AddItem "Avisos di·rio/Faturamento/Carteira de faturamento"
            .AddItem "Avisos di·rio/PCP/OSs em atraso"
            .AddItem "Avisos di·rio/Custos/Centro de custo"
            .AddItem "Avisos di·rio/An·lise crÌtica/Engenharia"
            .AddItem "Avisos di·rio/An·lise crÌtica/Processos"
            .AddItem "Avisos di·rio/An·lise crÌtica/Pcp"
            .AddItem "Avisos di·rio/An·lise crÌtica/Qualidade"
            .AddItem "Avisos di·rio/An·lise crÌtica/Compras"
            .AddItem "Avisos di·rio/Terceiros"
            .AddItem "Avisos di·rio/Compras/Pedidos em atraso"
            .AddItem "Avisos di·rio/Qualidade/CalibraÁ„o a vencer"
            .AddItem "Avisos di·rio/Qualidade/N„o conformidades"
            .AddItem "Avisos di·rio/Estoque/Produtos · vencer"
            .AddItem "Avisos di·rio/Processos/Sugestıes"
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaComboPais(Combo As ComboBox)
On Error GoTo tratar_erro

With Combo
    .Clear
    Set TBCarregarCombo = CreateObject("adodb.recordset")
    TBCarregarCombo.Open "Select * from Codigos_pais order by Pais", Conexao, adOpenKeyset, adLockOptimistic
    If TBCarregarCombo.EOF = False Then
        Do While TBCarregarCombo.EOF = False
            .AddItem Trim(TBCarregarCombo!Pais)
            .ItemData(.NewIndex) = TBCarregarCombo!CODIGO
            TBCarregarCombo.MoveNext
        Loop
    End If
    TBCarregarCombo.Close
End With

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaComboVersao(Combo As ComboBox, CarregarVazio As Boolean, ValidacaoObr As Boolean, Estrutura As Boolean, Processo As Boolean, Cod_interno As String)
On Error GoTo tratar_erro

With Combo
    .Clear
    If ValidacaoObr = True Then
    
        If Estrutura = True Then
            INNERJOINTEXTO = "PCDV.Versao from Projproduto P INNER JOIN Projconjunto_desc_versao PCDV ON PCDV.Codproduto = P.Codproduto where P.Desenho = '" & Cod_interno & "' and PCDV.DtValidacao IS NOT NULL group by PCDV.Versao"
        Else
            INNERJOINTEXTO = "F.Versao from (Projproduto P INNER JOIN Processos PR ON PR.Codproduto = P.Codproduto) INNER JOIN Fases F ON F.IDprocesso = PR.IDprocesso where P.Desenho = '" & Cod_interno & "' and PR.DtValidacao IS NOT NULL group by F.Versao"
        End If
       
        Set TBCarregarCombo = CreateObject("adodb.recordset")
        TBCarregarCombo.Open "Select " & INNERJOINTEXTO, Conexao, adOpenKeyset, adLockOptimistic
        If TBCarregarCombo.EOF = False Then
            If CarregarVazio = True Then .AddItem ""
            Do While TBCarregarCombo.EOF = False
                .AddItem TBCarregarCombo!versao
                TBCarregarCombo.MoveNext
            Loop
        End If
        TBCarregarCombo.Close
    Else
        .AddItem "A"
        .AddItem "B"
        .AddItem "C"
        .AddItem "D"
        .AddItem "E"
        .AddItem "F"
        .AddItem "G"
        .AddItem "H"
        .AddItem "I"
        .AddItem "J"
        .AddItem "K"
        .AddItem "L"
        .AddItem "M"
        .AddItem "N"
        .AddItem "O"
        .AddItem "P"
        .AddItem "Q"
        .AddItem "R"
        .AddItem "S"
        .AddItem "T"
        .AddItem "U"
        .AddItem "V"
        .AddItem "W"
        .AddItem "X"
        .AddItem "Y"
        .AddItem "Z"
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Function FunVerificaCidade(Cidade As String, UF As String) As Boolean
On Error GoTo tratar_erro

FunVerificaCidade = True
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from CEP where Municipio = '" & Cidade & "' and Sigla_UF = '" & UF & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = True Then
    USMsgBox ("N„o foi encontrado a cidade " & Cidade & " no estado de " & UF & ", favor revisar."), vbExclamation, "CAPRIND v5.0"
    FunVerificaCidade = False
End If
TBAbrir.Close

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Sub ProcVerificaRegistroUtilizado(NomeTabela As String, FiltroPadrao As String, Modulo As String)
On Error GoTo tratar_erro

Permitido = True
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from " & NomeTabela & " where " & FiltroPadrao, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    USMsgBox (Mensagem & " " & Modulo & "."), vbExclamation, "CAPRIND v5.0"
    Permitido = False
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcVerificaRegistroUtilizadoSemMsg(NomeTabela As String, FiltroPadrao As String)
On Error GoTo tratar_erro

Permitido = True
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from " & NomeTabela & " where " & FiltroPadrao, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then Permitido = False
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Function FunVerificaRegistroValidado(NomeTabela As String, FiltroPadrao As String, NomeRegistroPrincipal As String, NomeRegistro As String, NomeOperacao As String, MascPrincipal As Boolean, VerifValid As Boolean) As Boolean
On Error GoTo tratar_erro

'NOVO
FunVerificaRegistroValidado = True
If Formulario = "Estoque/Ordem de faturamento" Then
    TextoFiltro = "DtValidacaoOF"
ElseIf Formulario = "Vendas/Pedido interno" Then
        TextoFiltro = "DtValidacaoPI"
    Else
        TextoFiltro = "DtValidacao"
End If
If VerifValid = True Then
    TextoFiltro = TextoFiltro & " IS NOT NULL"
    MsgTexto = " j·"
Else
    TextoFiltro = TextoFiltro & " IS NULL"
    MsgTexto = " n„o"
End If

Set TBTempo = CreateObject("adodb.recordset")

StrSql = "Select * from " & NomeTabela & " where " & FiltroPadrao & " and " & TextoFiltro
'Debug.print StrSql

TBTempo.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
If TBTempo.EOF = False Then
    If MascPrincipal = True Then TextoMasc = "o" Else TextoMasc = "a"
    If NomeRegistro <> "" Then InicioMsg = NomeOperacao & " " & NomeRegistro Else InicioMsg = NomeOperacao
    USMsgBox ("N„o È permitido " & InicioMsg & ", pois " & TextoMasc & " " & NomeRegistroPrincipal & MsgTexto & " foi validad" & TextoMasc & "."), vbExclamation, "CAPRIND v5.0"
    FunVerificaRegistroValidado = False
End If
TBTempo.Close

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Function FunVerificaRegistroValidadoSemMsg(NomeTabela As String, FiltroPadrao As String, VerifValid As Boolean) As Boolean
On Error GoTo tratar_erro

'NOVO
FunVerificaRegistroValidadoSemMsg = True
If Formulario = "Estoque/Ordem de faturamento" Then
    TextoFiltro = "DtValidacaoOF"
ElseIf Formulario = "Vendas/Pedido interno" Then
        TextoFiltro = "DtValidacaoPI"
    Else
        TextoFiltro = "DtValidacao"
End If
If VerifValid = True Then TextoFiltro = TextoFiltro & " IS NOT NULL" Else TextoFiltro = TextoFiltro & " IS NULL"
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from " & NomeTabela & " where " & FiltroPadrao & " and " & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then FunVerificaRegistroValidadoSemMsg = False
TBAbrir.Close

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Sub ProcCarregaToolBar1(Formulario As Form, Comprimento As Integer, QtdeBotao As Integer, Visivel As Boolean)
On Error GoTo tratar_erro

With Formulario.USToolBar1
    .Theme = 1
    .DrawButtonsEx Formulario.USImageList1
    .Height = 975
    .Width = Comprimento
    Contador = QtdeBotao
    Do While QtdeBotao > 0
        .ButtonForeColor(QtdeBotao) = &H0&
        .ButtonFont(QtdeBotao).Bold = True
        QtdeBotao = QtdeBotao - 1
    Loop
    .Refresh
    If Visivel = True Then .Visible = True Else .Visible = False
End With

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaToolBar2(Formulario As Form, Comprimento As Integer, QtdeBotao As Integer, Visivel As Boolean)
On Error GoTo tratar_erro

With Formulario.USToolBar2
    .Theme = 1
    .DrawButtonsEx Formulario.USImageList2
    .Height = 975
    .Width = Comprimento
    Contador = QtdeBotao
    Do While QtdeBotao > 0
        .ButtonForeColor(QtdeBotao) = &H0&
        .ButtonFont(QtdeBotao).Bold = True
        QtdeBotao = QtdeBotao - 1
    Loop
    .Refresh
    If Visivel = True Then .Visible = True Else .Visible = False
End With

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaToolBar3(Formulario As Form, Comprimento As Integer, QtdeBotao As Integer, Visivel As Boolean)
On Error GoTo tratar_erro

With Formulario.USToolBar3
    .Theme = 1
    .DrawButtonsEx Formulario.USImageList3
    .Height = 975
    .Width = Comprimento
    Contador = QtdeBotao
    Do While QtdeBotao > 0
        .ButtonForeColor(QtdeBotao) = &H0&
        .ButtonFont(QtdeBotao).Bold = True
        QtdeBotao = QtdeBotao - 1
    Loop
    .Refresh
    If Visivel = True Then .Visible = True Else .Visible = False
End With

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaToolBar4(Formulario As Form, Comprimento As Integer, QtdeBotao As Integer, Visivel As Boolean)
On Error GoTo tratar_erro

With Formulario.USToolBar4
    .Theme = 1
    .DrawButtonsEx Formulario.USImageList4
    .Height = 975
    .Width = Comprimento
    Contador = QtdeBotao
    Do While QtdeBotao > 0
        .ButtonForeColor(QtdeBotao) = &H0&
        .ButtonFont(QtdeBotao).Bold = True
        QtdeBotao = QtdeBotao - 1
    Loop
    .Refresh
    If Visivel = True Then .Visible = True Else .Visible = False
End With

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaToolBar5(Formulario As Form, Comprimento As Integer, QtdeBotao As Integer, Visivel As Boolean)
On Error GoTo tratar_erro

With Formulario.USToolBar5
    .Theme = 1
    .DrawButtonsEx Formulario.USImageList5
    .Height = 975
    .Width = Comprimento
    Contador = QtdeBotao
    Do While QtdeBotao > 0
        .ButtonForeColor(QtdeBotao) = &H0&
        .ButtonFont(QtdeBotao).Bold = True
        QtdeBotao = QtdeBotao - 1
    Loop
    .Refresh
    If Visivel = True Then .Visible = True Else .Visible = False
End With

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaToolBar6(Formulario As Form, Comprimento As Integer, QtdeBotao As Integer, Visivel As Boolean)
On Error GoTo tratar_erro

With Formulario.USToolBar6
    .Theme = 1
    .DrawButtonsEx Formulario.USImageList6
    .Height = 975
    .Width = Comprimento
    Contador = QtdeBotao
    Do While QtdeBotao > 0
        .ButtonForeColor(QtdeBotao) = &H0&
        .ButtonFont(QtdeBotao).Bold = True
        QtdeBotao = QtdeBotao - 1
    Loop
    .Refresh
    If Visivel = True Then .Visible = True Else .Visible = False
End With

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaToolBar7(Formulario As Form, Comprimento As Integer, QtdeBotao As Integer, Visivel As Boolean)
On Error GoTo tratar_erro

With Formulario.USToolBar7
    .Theme = 1
    .DrawButtonsEx Formulario.USImageList7
    .Height = 975
    .Width = Comprimento
    Contador = QtdeBotao
    Do While QtdeBotao > 0
        .ButtonForeColor(QtdeBotao) = &H0&
        .ButtonFont(QtdeBotao).Bold = True
        QtdeBotao = QtdeBotao - 1
    Loop
    .Refresh
    If Visivel = True Then .Visible = True Else .Visible = False
End With

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaToolBar8(Formulario As Form, Comprimento As Integer, QtdeBotao As Integer, Visivel As Boolean)
On Error GoTo tratar_erro

With Formulario.USToolBar8
    .Theme = 1
    .DrawButtonsEx Formulario.USImageList8
    .Height = 975
    .Width = Comprimento
    Contador = QtdeBotao
    Do While QtdeBotao > 0
        .ButtonForeColor(QtdeBotao) = &H0&
        .ButtonFont(QtdeBotao).Bold = True
        QtdeBotao = QtdeBotao - 1
    Loop
    .Refresh
    If Visivel = True Then .Visible = True Else .Visible = False
End With

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaToolBar9(Formulario As Form, Comprimento As Integer, QtdeBotao As Integer, Visivel As Boolean)
On Error GoTo tratar_erro

With Formulario.USToolBar9
    .Theme = 1
    .DrawButtonsEx Formulario.USImageList9
    .Height = 975
    .Width = Comprimento
    Contador = QtdeBotao
    Do While QtdeBotao > 0
        .ButtonForeColor(QtdeBotao) = &H0&
        .ButtonFont(QtdeBotao).Bold = True
        
        QtdeBotao = QtdeBotao - 1
    Loop
    .Refresh
    If Visivel = True Then .Visible = True Else .Visible = False
End With

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaToolBar10(Formulario As Form, Comprimento As Integer, QtdeBotao As Integer, Visivel As Boolean)
On Error GoTo tratar_erro

With Formulario.USToolBar10
    .Theme = 1
    .DrawButtonsEx Formulario.USImageList10
    .Height = 975
    .Width = Comprimento
    Contador = QtdeBotao
    Do While QtdeBotao > 0
        .ButtonForeColor(QtdeBotao) = &H0&
        .ButtonFont(QtdeBotao).Bold = True
        QtdeBotao = QtdeBotao - 1
    Loop
    .Refresh
    If Visivel = True Then .Visible = True Else .Visible = False
End With

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaToolBar11(Formulario As Form, Comprimento As Integer, QtdeBotao As Integer, Visivel As Boolean)
On Error GoTo tratar_erro

With Formulario.USToolBar11
    .Theme = 1
    .DrawButtonsEx Formulario.USImageList11
    .Height = 975
    .Width = Comprimento
    Contador = QtdeBotao
    Do While QtdeBotao > 0
        .ButtonForeColor(QtdeBotao) = &H0&
        .ButtonFont(QtdeBotao).Bold = True
        QtdeBotao = QtdeBotao - 1
    Loop
    .Refresh
    If Visivel = True Then .Visible = True Else .Visible = False
End With

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

'############################################################################################
'VERIFICA NIVEL DO PLANO DE CONTAS

Sub ProcNivelPC7()
On Error GoTo tratar_erro

Nivel7A = Mid(TBNivel8!CODIGO, 1, 19)
Set TBNivel7 = CreateObject("adodb.recordset")
TBNivel7.Open "Select * from tbl_familia where Left(Codigo,19) = '" & Nivel7A & "' and Nivel = 7 order by Codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBNivel7.EOF = False Then
    Do While TBNivel7.EOF = False
        Set TBGravar = CreateObject("adodb.recordset")
        TBGravar.Open "Select * from Plano_de_contas_totalizacao where Codigo = '" & TBNivel7!CODIGO & "' and Nivel = " & TBNivel7!Nivel & " and Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
        If Formulario = "Financeiro/RelatÛrios/HistÛrico" Then ProcEnviaDadosPCRelFinanc TBNivel7!int_codfamilia, TBNivel7!CODIGO, TBNivel7!Txt_descricao, TBNivel7!Nivel Else ProcEnviaDadosPC TBNivel7!int_codfamilia, TBNivel7!CODIGO, TBNivel7!Txt_descricao, TBNivel7!Nivel
        
        ProcNivelPC6
        
        TBNivel7.MoveNext
    Loop
End If
TBNivel7.Close

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcNivelPC6()
On Error GoTo tratar_erro

Nivel6A = Mid(TBNivel7!CODIGO, 1, 16)
Set TBNivel6 = CreateObject("adodb.recordset")
TBNivel6.Open "Select * from tbl_familia where Left(Codigo,16) = '" & Nivel6A & "' and Nivel = 6 order by Codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBNivel6.EOF = False Then
    Do While TBNivel6.EOF = False
        Set TBGravar = CreateObject("adodb.recordset")
        TBGravar.Open "Select * from Plano_de_contas_totalizacao where Codigo = '" & TBNivel6!CODIGO & "' and Nivel = " & TBNivel6!Nivel & " and Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
        If Formulario = "Financeiro/RelatÛrios/HistÛrico" Then ProcEnviaDadosPCRelFinanc TBNivel6!int_codfamilia, TBNivel6!CODIGO, TBNivel6!Txt_descricao, TBNivel6!Nivel Else ProcEnviaDadosPC TBNivel6!int_codfamilia, TBNivel6!CODIGO, TBNivel6!Txt_descricao, TBNivel6!Nivel
        
        ProcNivelPC5
        
        TBNivel6.MoveNext
    Loop
End If
TBNivel6.Close

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcNivelPC5()
On Error GoTo tratar_erro

Nivel5A = Mid(TBNivel6!CODIGO, 1, 13)
Set TBNivel5 = CreateObject("adodb.recordset")
TBNivel5.Open "Select * from tbl_familia where Left(Codigo,13) = '" & Nivel5A & "' and Nivel = 5 order by Codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBNivel5.EOF = False Then
    Do While TBNivel5.EOF = False
        Set TBGravar = CreateObject("adodb.recordset")
        TBGravar.Open "Select * from Plano_de_contas_totalizacao where Codigo = '" & TBNivel5!CODIGO & "' and Nivel = " & TBNivel5!Nivel & " and Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
        If Formulario = "Financeiro/RelatÛrios/HistÛrico" Then ProcEnviaDadosPCRelFinanc TBNivel5!int_codfamilia, TBNivel5!CODIGO, TBNivel5!Txt_descricao, TBNivel5!Nivel Else ProcEnviaDadosPC TBNivel5!int_codfamilia, TBNivel5!CODIGO, TBNivel5!Txt_descricao, TBNivel5!Nivel
        
        ProcNivelPC4
        
        TBNivel5.MoveNext
    Loop
End If
TBNivel5.Close

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcNivelPC4()
On Error GoTo tratar_erro

Nivel4A = Mid(TBNivel5!CODIGO, 1, 10)
Set TBNivel4 = CreateObject("adodb.recordset")
TBNivel4.Open "Select * from tbl_familia where Left(Codigo,10) = '" & Nivel4A & "' and Nivel = 4 order by Codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBNivel4.EOF = False Then
    Do While TBNivel4.EOF = False
        Set TBGravar = CreateObject("adodb.recordset")
        TBGravar.Open "Select * from Plano_de_contas_totalizacao where Codigo = '" & TBNivel4!CODIGO & "' and Nivel = " & TBNivel4!Nivel & " and Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
        If Formulario = "Financeiro/RelatÛrios/HistÛrico" Then ProcEnviaDadosPCRelFinanc TBNivel4!int_codfamilia, TBNivel4!CODIGO, TBNivel4!Txt_descricao, TBNivel4!Nivel Else ProcEnviaDadosPC TBNivel4!int_codfamilia, TBNivel4!CODIGO, TBNivel4!Txt_descricao, TBNivel4!Nivel
        
        ProcNivelPC3
        
        TBNivel4.MoveNext
    Loop
End If
TBNivel4.Close

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcNivelPC3()
On Error GoTo tratar_erro

Nivel3A = Mid(TBNivel4!CODIGO, 1, 7)
Set TBNivel3 = CreateObject("adodb.recordset")
TBNivel3.Open "Select * from tbl_familia where Left(Codigo,7) = '" & Nivel3A & "' and Nivel = 3 order by Codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBNivel3.EOF = False Then
    Do While TBNivel3.EOF = False
        Set TBGravar = CreateObject("adodb.recordset")
        TBGravar.Open "Select * from Plano_de_contas_totalizacao where Codigo = '" & TBNivel3!CODIGO & "' and Nivel = " & TBNivel3!Nivel & " and Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
        If Formulario = "Financeiro/RelatÛrios/HistÛrico" Then ProcEnviaDadosPCRelFinanc TBNivel3!int_codfamilia, TBNivel3!CODIGO, TBNivel3!Txt_descricao, TBNivel3!Nivel Else ProcEnviaDadosPC TBNivel3!int_codfamilia, TBNivel3!CODIGO, TBNivel3!Txt_descricao, TBNivel3!Nivel
        
        ProcNivelPC2
        
        TBNivel3.MoveNext
    Loop
End If
TBNivel3.Close

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcNivelPC2()
On Error GoTo tratar_erro

Nivel2A = Mid(TBNivel3!CODIGO, 1, 4)
Set TBNivel2 = CreateObject("adodb.recordset")
TBNivel2.Open "Select * from tbl_familia where Left(Codigo,4) = '" & Nivel2A & "' and Nivel = 2 order by Codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBNivel2.EOF = False Then
    Do While TBNivel2.EOF = False
        Set TBGravar = CreateObject("adodb.recordset")
        TBGravar.Open "Select * from Plano_de_contas_totalizacao where Codigo = '" & TBNivel2!CODIGO & "' and Nivel = " & TBNivel2!Nivel & " and Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
        If Formulario = "Financeiro/RelatÛrios/HistÛrico" Then ProcEnviaDadosPCRelFinanc TBNivel2!int_codfamilia, TBNivel2!CODIGO, TBNivel2!Txt_descricao, TBNivel2!Nivel Else ProcEnviaDadosPC TBNivel2!int_codfamilia, TBNivel2!CODIGO, TBNivel2!Txt_descricao, TBNivel2!Nivel
        
        ProcNivelPC1
        
        TBNivel2.MoveNext
    Loop
End If
TBNivel2.Close

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcNivelPC1()
On Error GoTo tratar_erro

Nivel1A = Mid(TBNivel2!CODIGO, 1, 1)
Set TBNivel1 = CreateObject("adodb.recordset")
TBNivel1.Open "Select * from tbl_familia where Left(Codigo,1) = '" & Nivel1A & "' and Nivel = 1 order by Codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBNivel1.EOF = False Then
    Do While TBNivel1.EOF = False
        Set TBGravar = CreateObject("adodb.recordset")
        TBGravar.Open "Select * from Plano_de_contas_totalizacao where Codigo = '" & TBNivel1!CODIGO & "' and Nivel = " & TBNivel1!Nivel & " and Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
        If Formulario = "Financeiro/RelatÛrios/HistÛrico" Then ProcEnviaDadosPCRelFinanc TBNivel1!int_codfamilia, TBNivel1!CODIGO, TBNivel1!Txt_descricao, TBNivel1!Nivel Else ProcEnviaDadosPC TBNivel1!int_codfamilia, TBNivel1!CODIGO, TBNivel1!Txt_descricao, TBNivel1!Nivel
        TBNivel1.MoveNext
    Loop
End If
TBNivel1.Close

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcEnviaDadosPC(IDPC As Long, CODIGO As String, Descricao As String, Nivel As Integer)
On Error GoTo tratar_erro

If TBGravar.EOF = True Then
    TBGravar.AddNew
    TBGravar!valor = Valor1
Else
    TBGravar!valor = TBGravar!valor + Valor1
End If
TBGravar!Modulo = Formulario
TBGravar!Responsavel = pubUsuario
TBGravar!ID_PC = IDPC
TBGravar!CODIGO = CODIGO
TBGravar!Descricao = Descricao
TBGravar!Nivel = Nivel
TBGravar.Update
TBGravar.Close

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcEnviaDadosPCRelFinanc(IDPC As Long, CODIGO As String, Descricao As String, Nivel As Integer)
On Error GoTo tratar_erro

TBGravar.AddNew
TBGravar!valor = IIf(IsNull(TBFamilia!valor), 0, TBFamilia!valor)
TBGravar!Modulo = Formulario
TBGravar!Responsavel = pubUsuario
TBGravar!ID_PC = IDPC

Select Case Nivel
    Case 1: Codigo1 = Left(CODIGO, 1)
    Case 2: Codigo1 = Left(CODIGO, 4)
    Case 3: Codigo1 = Left(CODIGO, 7)
    Case 4: Codigo1 = Left(CODIGO, 10)
    Case 5: Codigo1 = Left(CODIGO, 13)
    Case 6: Codigo1 = Left(CODIGO, 16)
    Case 7: Codigo1 = Left(CODIGO, 19)
    Case 8: Codigo1 = Left(CODIGO, 22)
End Select

TBGravar!CODIGO = Codigo1
TBGravar!Descricao = Descricao
TBGravar!Nivel = Nivel

With frmFinanceiro_Relatorios
    If .Cmb_tipo = "A receber" Or .Cmb_tipo = "Recebidas" Or .Cmb_tipo = "A receber e recebidas" Or .Cmb_tipo = "A receber e a pagar" Or .Cmb_tipo = "Recebidas e pagas" Then
        TBGravar!emissao = TBFamilia!emissao
        TBGravar!Mes_Emissao = Month(TBFamilia!emissao)
        TBGravar!Ano_Emissao = Year(TBFamilia!emissao)
        TBGravar!Vencimento = TBFamilia!Vencimento
        TBGravar!Mes_venc = Month(TBFamilia!Vencimento)
        TBGravar!Ano_Venc = Year(TBFamilia!Vencimento)
        If IsNull(TBFamilia!Data_pagamento) = False And TBFamilia!Data_pagamento <> "" Then
            TBGravar!Pagamento_receb = TBFamilia!Data_pagamento
            TBGravar!Mes_pgto_receb = Month(TBFamilia!Data_pagamento)
            TBGravar!Ano_pgto_receb = Year(TBFamilia!Data_pagamento)
        End If
        TBGravar!Valor_pgto_receb = IIf(IsNull(TBFamilia!valor), 0, TBFamilia!valor)
    Else
        TBGravar!emissao = TBFamilia!Dt_emissao
        TBGravar!Mes_Emissao = Month(TBFamilia!Dt_emissao)
        TBGravar!Ano_Emissao = Year(TBFamilia!Dt_emissao)
        TBGravar!Vencimento = TBFamilia!dt_Pagamento
        TBGravar!Mes_venc = Month(TBFamilia!dt_Pagamento)
        TBGravar!Ano_Venc = Year(TBFamilia!dt_Pagamento)
        If IsNull(TBFamilia!DataBaixa) = False And TBFamilia!DataBaixa <> "" Then
            TBGravar!Pagamento_receb = TBFamilia!DataBaixa
            TBGravar!Mes_pgto_receb = Month(TBFamilia!DataBaixa)
            TBGravar!Ano_pgto_receb = Year(TBFamilia!DataBaixa)
        End If
        TBGravar!Valor_pgto_receb = IIf(IsNull(TBFamilia!valor), 0, TBFamilia!valor)
    End If
    If .Cmb_tipo = "A receber e a pagar" Or .Cmb_tipo = "Recebidas e pagas" Then TBGravar!Tipo = TBFamilia!TipoConta
End With

TBGravar!Logsit = TBFamilia!Logsit
TBGravar.Update
TBGravar.Close

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Function FunVerifNumPagina(Variavel As String)
On Error GoTo tratar_erro

QtdeLetras = Len(Variavel)
Texto = ""
Pagina = ""
Numero = 8
Do While QtdeLetras > 8
    Texto = Left(Variavel, (Numero + 1))
    Texto = Right(Texto, Len(Texto) - Numero)
    If Texto = " " Then Exit Function
    If FunVerifNumPagina = "" Then FunVerifNumPagina = Texto Else FunVerifNumPagina = FunVerifNumPagina & Texto
    Numero = Numero + 1
    QtdeLetras = QtdeLetras - 1
Loop

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Function FunLiberaCamposEstrutura() As Boolean
On Error GoTo tratar_erro

FunLiberaCamposEstrutura = False
Set TBFIltro = CreateObject("adodb.recordset")
TBFIltro.Open "Select * from Empresa where Liberar_campos_estrutura = 'True'", Conexao, adOpenKeyset, adLockOptimistic
If TBFIltro.EOF = False Then
    FunLiberaCamposEstrutura = True
End If
TBFIltro.Close

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Function FunVerificaMes(Mes As String)
On Error GoTo tratar_erro

Select Case Mes
    Case "Jan": FunVerificaMes = 1
    Case "Fev": FunVerificaMes = 2
    Case "Mar": FunVerificaMes = 3
    Case "Abr": FunVerificaMes = 4
    Case "Maio": FunVerificaMes = 5
    Case "Jun": FunVerificaMes = 6
    Case "Jul": FunVerificaMes = 7
    Case "Ago": FunVerificaMes = 8
    Case "Set": FunVerificaMes = 9
    Case "Out": FunVerificaMes = 10
    Case "Nov": FunVerificaMes = 11
    Case "Dez": FunVerificaMes = 12
End Select

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Function FunVerificaNumeroMes(Mes As Integer)
On Error GoTo tratar_erro

Select Case Mes
    Case 1: FunVerificaNumeroMes = "Jan"
    Case 2: FunVerificaNumeroMes = "Fev"
    Case 3: FunVerificaNumeroMes = "Mar"
    Case 4: FunVerificaNumeroMes = "Abr"
    Case 5: FunVerificaNumeroMes = "Maio"
    Case 6: FunVerificaNumeroMes = "Jun"
    Case 7: FunVerificaNumeroMes = "Jul"
    Case 8: FunVerificaNumeroMes = "Ago"
    Case 9: FunVerificaNumeroMes = "Set"
    Case 10: FunVerificaNumeroMes = "Out"
    Case 11: FunVerificaNumeroMes = "Nov"
    Case 12: FunVerificaNumeroMes = "Dez"
End Select

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Sub ProcAtualizaSaldoSaque(ID_saque As Long)
On Error GoTo tratar_erro

Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from tbl_instituicoes_transf where ID_transf = " & ID_saque, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = False Then
    Valor_total = 0
    Set TBMaquinas = CreateObject("adodb.recordset")
    TBMaquinas.Open "Select Sum(Valor_utilizado) as Valor_Total from tbl_ContasPagar_Saque where IDSaque = " & TBGravar!id_transf, Conexao, adOpenKeyset, adLockOptimistic
    If TBMaquinas.EOF = False Then
        Valor_total = IIf(IsNull(TBMaquinas!Valor_total), 0, TBMaquinas!Valor_total)
    End If
    TBMaquinas.Close
    
    TBGravar!Saldo = TBGravar!valor_transf - Valor_total
    TBGravar.Update
End If
TBGravar.Close

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcNivel2EstruturaAcima(NomeFormulario As Form, versao As String, TodaEstrutura As Boolean)
On Error GoTo tratar_erro

If versao = "" Then TextoFiltro = "" Else TextoFiltro = "and Versao_desenho = '" & versao & "'"
Set TBNivel2 = CreateObject("adodb.recordset")
TBNivel2.Open "Select codproduto, Versao from projconjunto where desenho = '" & DesenhoProduto & "' " & TextoFiltro & " order by Posicao, Codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBNivel2.EOF = False Then
    Do While TBNivel2.EOF = False
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select Desenho from projproduto where codproduto = " & TBNivel2!Codproduto, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            DesenhoProduto = TBAbrir!Desenho
            Letra = TBNivel2!versao
            If TodaEstrutura = True Then ProcNivel3EstruturaAcima NomeFormulario, TBNivel2!versao
        End If
        If Familiatext = "" Then
            Familiatext = "P.Desenho = '" & DesenhoProduto & "' and PC.Versao = '" & Letra & "'"
        Else
            Familiatext = Familiatext & " or P.Desenho = '" & DesenhoProduto & "' and PC.Versao = '" & Letra & "'"
            FamiliaAntiga = "("
            Pesquisa = ")"
        End If
        TBNivel2.MoveNext
    Loop
End If

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcNivel3EstruturaAcima(NomeFormulario As Form, versao As String)
On Error GoTo tratar_erro

If versao = "" Then TextoFiltro = "" Else TextoFiltro = "and Versao_desenho = '" & versao & "'"
Set TBNivel3 = CreateObject("adodb.recordset")
TBNivel3.Open "Select codproduto, Versao from projconjunto where desenho = '" & DesenhoProduto & "' " & TextoFiltro & " order by Posicao, Codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBNivel3.EOF = False Then
    Do While TBNivel3.EOF = False
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select Desenho from projproduto where codproduto = " & TBNivel3!Codproduto, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            DesenhoProduto = TBAbrir!Desenho
            ProcNivel4EstruturaAcima NomeFormulario, TBNivel3!versao
        End If
        TBNivel3.MoveNext
    Loop
End If

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcNivel4EstruturaAcima(NomeFormulario As Form, versao As String)
On Error GoTo tratar_erro

If versao = "" Then TextoFiltro = "" Else TextoFiltro = "and Versao_desenho = '" & versao & "'"
Set TBNivel4 = CreateObject("adodb.recordset")
TBNivel4.Open "Select codproduto, versao from projconjunto where desenho = '" & DesenhoProduto & "' " & TextoFiltro & " order by Posicao, Codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBNivel4.EOF = False Then
    Do While TBNivel4.EOF = False
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select Desenho from projproduto where codproduto = " & TBNivel4!Codproduto, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            DesenhoProduto = TBAbrir!Desenho
            ProcNivel5EstruturaAcima NomeFormulario, TBNivel4!versao
        End If
        TBNivel4.MoveNext
    Loop
End If

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcNivel5EstruturaAcima(NomeFormulario As Form, versao As String)
On Error GoTo tratar_erro

If versao = "" Then TextoFiltro = "" Else TextoFiltro = "and Versao_desenho = '" & versao & "'"
Set TBNivel5 = CreateObject("adodb.recordset")
TBNivel5.Open "Select codproduto, versao from projconjunto where desenho = '" & DesenhoProduto & "' " & TextoFiltro & " order by Posicao, Codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBNivel5.EOF = False Then
    Do While TBNivel5.EOF = False
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select Desenho from projproduto where codproduto = " & TBNivel5!Codproduto, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            DesenhoProduto = TBAbrir!Desenho
            ProcNivel6EstruturaAcima NomeFormulario, TBNivel5!versao
        End If
        TBNivel5.MoveNext
    Loop
End If

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcNivel6EstruturaAcima(NomeFormulario As Form, versao As String)
On Error GoTo tratar_erro

If versao = "" Then TextoFiltro = "" Else TextoFiltro = "and Versao_desenho = '" & versao & "'"
Set TBNivel6 = CreateObject("adodb.recordset")
TBNivel6.Open "Select codproduto, versao from projconjunto where desenho = '" & DesenhoProduto & "' " & TextoFiltro & " order by Posicao, Codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBNivel6.EOF = False Then
    Do While TBNivel6.EOF = False
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select Desenho from projproduto where codproduto = " & TBNivel6!Codproduto, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            DesenhoProduto = TBAbrir!Desenho
            ProcNivel7EstruturaAcima NomeFormulario, TBNivel6!versao
        End If
        TBNivel6.MoveNext
    Loop
End If

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcNivel7EstruturaAcima(NomeFormulario As Form, versao As String)
On Error GoTo tratar_erro

If versao = "" Then TextoFiltro = "" Else TextoFiltro = "and Versao_desenho = '" & versao & "'"
Set TBNivel7 = CreateObject("adodb.recordset")
TBNivel7.Open "Select codproduto, versao from projconjunto where desenho = '" & DesenhoProduto & "' " & TextoFiltro & " order by Posicao, Codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBNivel7.EOF = False Then
    Do While TBNivel7.EOF = False
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select Desenho from projproduto where codproduto = " & TBNivel7!Codproduto, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            DesenhoProduto = TBAbrir!Desenho
            ProcNivel8EstruturaAcima NomeFormulario, TBNivel7!versao
        End If
        TBNivel7.MoveNext
    Loop
End If

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcNivel8EstruturaAcima(NomeFormulario As Form, versao As String)
On Error GoTo tratar_erro

If versao = "" Then TextoFiltro = "" Else TextoFiltro = "and Versao_desenho = '" & versao & "'"
Set TBNivel8 = CreateObject("adodb.recordset")
TBNivel8.Open "Select codproduto, versao from projconjunto where desenho = '" & DesenhoProduto & "' " & TextoFiltro & " order by Posicao, Codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBNivel8.EOF = False Then
    Do While TBNivel8.EOF = False
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select Desenho from projproduto where codproduto = " & TBNivel8!Codproduto, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            DesenhoProduto = TBAbrir!Desenho
            ProcNivel9EstruturaAcima NomeFormulario, TBNivel8!versao
        End If
        TBNivel8.MoveNext
    Loop
End If

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcNivel9EstruturaAcima(NomeFormulario As Form, versao As String)
On Error GoTo tratar_erro

If versao = "" Then TextoFiltro = "" Else TextoFiltro = "and Versao_desenho = '" & versao & "'"
Set TBNivel9 = CreateObject("adodb.recordset")
TBNivel9.Open "Select codproduto, versao from projconjunto where desenho = '" & DesenhoProduto & "' " & TextoFiltro & " order by Posicao, Codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBNivel9.EOF = False Then
    Do While TBNivel9.EOF = False
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select Desenho from projproduto where codproduto = " & TBNivel9!Codproduto, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            DesenhoProduto = TBAbrir!Desenho
            ProcNivel10EstruturaAcima NomeFormulario, TBNivel9!versao
        End If
        TBNivel9.MoveNext
    Loop
End If

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcNivel10EstruturaAcima(NomeFormulario As Form, versao As String)
On Error GoTo tratar_erro

If versao = "" Then TextoFiltro = "" Else TextoFiltro = "and Versao_desenho = '" & versao & "'"
Set TBNivel10 = CreateObject("adodb.recordset")
TBNivel10.Open "Select codproduto, versao from projconjunto where desenho = '" & DesenhoProduto & "' " & TextoFiltro & " order by Posicao, Codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBNivel10.EOF = False Then
    Do While TBNivel10.EOF = False
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select Desenho from projproduto where codproduto = " & TBNivel10!Codproduto, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            DesenhoProduto = TBAbrir!Desenho
            ProcNivel11EstruturaAcima NomeFormulario, TBNivel10!versao
        End If
        TBNivel10.MoveNext
    Loop
End If

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcNivel11EstruturaAcima(NomeFormulario As Form, versao As String)
On Error GoTo tratar_erro

If versao = "" Then TextoFiltro = "" Else TextoFiltro = "and Versao_desenho = '" & versao & "'"
Set TBNivel11 = CreateObject("adodb.recordset")
TBNivel11.Open "Select codproduto, versao from projconjunto where desenho = '" & DesenhoProduto & "' " & TextoFiltro & " order by Posicao, Codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBNivel11.EOF = False Then
    Do While TBNivel11.EOF = False
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select Desenho from projproduto where codproduto = " & TBNivel11!Codproduto, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            DesenhoProduto = TBAbrir!Desenho
            ProcNivel12EstruturaAcima NomeFormulario, TBNivel11!versao
        End If
        TBNivel11.MoveNext
    Loop
End If

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcNivel12EstruturaAcima(NomeFormulario As Form, versao As String)
On Error GoTo tratar_erro

If versao = "" Then TextoFiltro = "" Else TextoFiltro = "and Versao_desenho = '" & versao & "'"
Set TBNivel12 = CreateObject("adodb.recordset")
TBNivel12.Open "Select codproduto, versao from projconjunto where desenho = '" & DesenhoProduto & "' " & TextoFiltro & " order by Posicao, Codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBNivel12.EOF = False Then
    Do While TBNivel12.EOF = False
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select Desenho from projproduto where codproduto = " & TBNivel12!Codproduto, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            DesenhoProduto = TBAbrir!Desenho
            ProcNivel13EstruturaAcima NomeFormulario, TBNivel12!versao
        End If
        TBNivel12.MoveNext
    Loop
End If

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcNivel13EstruturaAcima(NomeFormulario As Form, versao As String)
On Error GoTo tratar_erro

If versao = "" Then TextoFiltro = "" Else TextoFiltro = "and Versao_desenho = '" & versao & "'"
Set TBNivel13 = CreateObject("adodb.recordset")
TBNivel13.Open "Select codproduto, versao from projconjunto where desenho = '" & DesenhoProduto & "' " & TextoFiltro & " order by Posicao, Codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBNivel13.EOF = False Then
    Do While TBNivel13.EOF = False
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select Desenho from projproduto where codproduto = " & TBNivel13!Codproduto, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            DesenhoProduto = TBAbrir!Desenho
            ProcNivel14EstruturaAcima NomeFormulario, TBNivel13!versao
        End If
        TBNivel13.MoveNext
    Loop
End If

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcNivel14EstruturaAcima(NomeFormulario As Form, versao As String)
On Error GoTo tratar_erro

If versao = "" Then TextoFiltro = "" Else TextoFiltro = "and Versao_desenho = '" & versao & "'"
Set TBNivel14 = CreateObject("adodb.recordset")
TBNivel14.Open "Select codproduto, versao from projconjunto where desenho = '" & DesenhoProduto & "' " & TextoFiltro & " order by Posicao, Codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBNivel14.EOF = False Then
    Do While TBNivel14.EOF = False
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select Desenho from projproduto where codproduto = " & TBNivel14!Codproduto, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            DesenhoProduto = TBAbrir!Desenho
            ProcNivel15EstruturaAcima NomeFormulario, TBNivel14!versao
        End If
        TBNivel14.MoveNext
    Loop
End If

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcNivel15EstruturaAcima(NomeFormulario As Form, versao As String)
On Error GoTo tratar_erro

If versao = "" Then TextoFiltro = "" Else TextoFiltro = "and Versao_desenho = '" & versao & "'"
Set TBNivel15 = CreateObject("adodb.recordset")
TBNivel15.Open "Select codproduto, versao from projconjunto where desenho = '" & DesenhoProduto & "' " & TextoFiltro & " order by Posicao, Codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBNivel15.EOF = False Then
    Do While TBNivel15.EOF = False
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select Desenho from projproduto where codproduto = " & TBNivel15!Codproduto, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            DesenhoProduto = TBAbrir!Desenho
        End If
        TBNivel15.MoveNext
    Loop
End If

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Function FunCarregaValorEstrutura(IDestrutura As Long, Codinterno As String, MostrarValores As Boolean, Simulacao_vendas As Boolean, Quant As Double, SomarValorTotal As Boolean) As String
On Error GoTo tratar_erro

Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select Responsavel, Modulo, Ordem, Qtdeprev, QtdeOK, QtdeNC, Qtdetotalprod, Eficiencia, Terceiros from Producao_Relatorios where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
TBGravar.AddNew
TBGravar!Responsavel = pubUsuario
If Simulacao_vendas = True Then TBGravar!Modulo = "Vendas/SimulaÁ„o" Else TBGravar!Modulo = "Engenharia/Estrutura"
TBGravar!Ordem = IDestrutura

If Simulacao_vendas = True Then
    QuantComprado = FunVerificaQtdeEstoque(Codinterno, 0, "") 'Verifica quantidade no estoque com material consignado
    'MsgTexto = "   .   Qtde. estoque: " & Format(QuantComprado, "###,##0.0000")
    TBGravar!qtdeOK = Quant
    TBGravar!qtdeNC = QuantComprado
    TBGravar!Qtdetotalprod = IIf(QuantComprado - Quant < 0, (QuantComprado - Quant) * -1, 0)
End If

If MostrarValores = True Then
    'Verifica custo do estoque
    Call FunVerificaQtdeEstoque(Codinterno, 0, "and Consignacao = 'False'") 'Verifica valor no estoque sem material consignado
    If CTMedioEst <> 0 Then
        valor = Format(CTMedioEst * Quant, "###,##0.00000000") 'Medio custo
        If Simulacao_vendas = True Then
            Set TBComponente = CreateObject("adodb.recordset")
            TBComponente.Open "Select ROUND(Max(Valor_total / estoque_real), 8) as MAX, ROUND(Min(Valor_total / estoque_real), 8) as Min from Qtde_estoque_produto where desenho = '" & Codinterno & "' and Consignacao = 'False' and Valor_total > 0", Conexao, adOpenKeyset, adLockOptimistic
            If TBComponente.EOF = False Then
                Valor_Cofins_Prod = Format(IIf(IsNull(TBComponente!Max), 0, TBComponente!Max) * Quant, "###,##0.00000000") 'Max custo
                Valor_Cofins_Serv = Format(IIf(IsNull(TBComponente!Min), 0, TBComponente!Min) * Quant, "###,##0.00000000") 'Min custo
            End If
            TBComponente.Close
        End If
    Else
        'Verifica custo de compras
        valor = FunVerificaVlrUltCompra(Codinterno) 'Verifica valor unit·rio da ˙ltima compra
        If valor <> 0 Then
            valor = Format(valor * Quant, "###,##0.00000000")
            If Simulacao_vendas = True Then
                Set TBComponente = CreateObject("adodb.recordset")
                TBComponente.Open "Select ROUND(MAX(CPL.preco_unitario * ISNULL(CC.Valor_moeda, 1)), 10) AS max, ROUND(MIN(CPL.preco_unitario * ISNULL(CC.Valor_moeda, 1)), 10) AS Min from dbo.Compras_pedido_lista AS CPL LEFT OUTER JOIN dbo.Compras_comercial AS CC ON CPL.IDPedido = CC.IdPedido where CPL.Desenho = '" & Codinterno & "' and CPL.IDpedido <> 0 AND CPL.preco_unitario > 0", Conexao, adOpenKeyset, adLockOptimistic
                If TBComponente.EOF = False Then
                    Valor_Cofins_Prod = Format(IIf(IsNull(TBComponente!Max), 0, TBComponente!Max) * Quant, "###,##0.00000000") 'Max custo
                    Valor_Cofins_Serv = Format(IIf(IsNull(TBComponente!Min), 0, TBComponente!Min) * Quant, "###,##0.00000000") 'Min custo
                End If
                TBComponente.Close
            End If
        Else
            'Verifica custo do item
            valor = Format(IIf(IsNull(TBAbrir!PCusto), 0, TBAbrir!PCusto) * Quant, "###,##0.00000000")
            If Simulacao_vendas = True Then
                Valor_Cofins_Prod = valor 'Max custo
                Valor_Cofins_Serv = valor 'Min custo
            End If
        End If
    End If
    If SomarValorTotal = True Then ValorPago = ValorPago + valor 'Valor total custo
    TBGravar!QtdePrev = Format(valor, "###,##0.00000000")
    If Simulacao_vendas = True Then
        If SomarValorTotal = True Then
            ValorTotalPagar = ValorTotalPagar + Valor_Cofins_Prod 'Valor total Max custo
            ValorTotalPago = ValorTotalPago + Valor_Cofins_Serv 'Valor total Min custo
        End If
        TBGravar!Eficiencia = Format(Valor_Cofins_Prod, "###,##0.00000000")
        TBGravar!Terceiros = Format(Valor_Cofins_Serv, "###,##0.00000000")
    End If
Else
    'Mensagem = ""
    valor = 0
    Valor_Cofins_Prod = 0
    Valor_Cofins_Serv = 0
    TBGravar!QtdePrev = 0
    TBGravar!Eficiencia = 0
    TBGravar!Terceiros = 0
End If
TBGravar.Update
TBGravar.Close

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Public Function FunVerifAcessoContasFunc(Modulo As String) As Boolean
On Error GoTo tratar_erro

Set TBAcessos = CreateObject("adodb.recordset")
TBAcessos.Open "Select * FROM Acessos where IDUsuario = " & pubIDUsuario & " AND Acesso = '" & Modulo & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBAcessos.EOF = True Then
    FunVerifAcessoContasFunc = False
Else
    FunVerifAcessoContasFunc = True
End If
TBAcessos.Close

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Sub ProcAtualizaValorProdServ(Valor_custo As Boolean, PCusto As Double, Valor_consumo As Boolean, PConsumo As Double, PRevenda As Double, Codinterno As String)
On Error GoTo tratar_erro

Set TBItem = CreateObject("adodb.recordset")
TBItem.Open "Select * from projproduto where Desenho = '" & Codinterno & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBItem.EOF = False Then
    If Valor_custo = True Then
        If IsNull(TBItem!importacao) = True Or TBItem!importacao = "" Or TBItem!importacao = False Then TBItem!PCusto = PCusto
    ElseIf Valor_consumo = True Then
            If IsNull(TBItem!exportacao) = True Or TBItem!exportacao = "" Or TBItem!exportacao = False Then TBItem!PConsumo = PConsumo
        Else
            If IsNull(TBItem!exportacao) = True Or TBItem!exportacao = "" Or TBItem!exportacao = False Then TBItem!PRevenda = PRevenda
    End If
    TBItem.Update
End If
TBItem.Close

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcSubstituicaoTributaria(UFcliente As String, CST As String, ID_CF As Long, IDClienteImp As Long, ClienteImp As String, VlrUnit As Double, QtdeProd As Double, VlrBC As Double, VlrBCST As Double, VlrFrete As Double, VlrSeguro As Double, VlrAcessorias As Double, NF As Boolean, consumidorFinal As Boolean, IDempresaST As Integer)
On Error GoTo tratar_erro
Dim ICMSSTSimples As Boolean
Dim naoContribuinteICMS As Boolean
Dim UFempresa As String

QtdeSaida = 0
BCICMSCST = 0
TTICMS = 0
ICMSCST = 0
If UFcliente <> "" And CST <> "" And ID_CF <> "0" Then
    Set TBMaquinas = CreateObject("adodb.recordset")
    TBMaquinas.Open "Select * from regioes where uf = '" & UFcliente & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBMaquinas.EOF = False Then
        Set TBFI = CreateObject("adodb.recordset")
        TBFI.Open "Select Margem, Aliquota, Aliquota_Aplicavel from CST where Id_UF = " & TBMaquinas!ID & " and CST = '" & CST & "' and ID_CF = " & ID_CF, Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False Then
                If Right(CST, 2) = "10" Or Right(CST, 2) = "20" Or Right(CST, 2) = "51" Or Right(CST, 2) = "60" Or Right(CST, 2) = "70" Or Right(CST, 3) = "201" Or Right(CST, 3) = "202" Or Right(CST, 3) = "203" Or Right(CST, 3) = "500" Or Right(CST, 3) = "900" Then
                    ProcBuscaTributos (ID_CF)
                    ProcVerificaRegiao UFcliente, IDClienteImp, ClienteImp
                    
                    ICMSSTSimples = False
                    naoContribuinteICMS = False
                    If UFcliente = "MT" Then
                        Set TBCorretiva = CreateObject("adodb.recordset")
                        TBCorretiva.Open "Select SimplesICMSST, Nao_contribuinte_ICMS from Clientes where IDCliente = " & IDClienteImp & " and NomeRazao = '" & ClienteImp & "'", Conexao, adOpenKeyset, adLockReadOnly
                        If TBCorretiva.EOF = False Then
                            If TBCorretiva!SimplesICMSST = True Then ICMSSTSimples = True
                            If TBCorretiva!Nao_contribuinte_ICMS = True Then naoContribuinteICMS = True
                        Else
                            Set TBCorretiva = CreateObject("adodb.recordset")
                            TBCorretiva.Open "Select ICMSSTSimples, contribuinteICMS from Compras_fornecedores where IDCliente = " & IDClienteImp & " and Nome_Razao = '" & ClienteImp & "'", Conexao, adOpenKeyset, adLockReadOnly
                            If TBCorretiva.EOF = False Then
                                If TBCorretiva!SimplesICMSST = True Then ICMSSTSimples = True
                                If TBCorretiva!Nao_contribuinte_ICMS = True Then naoContribuinteICMS = True
                            End If
                        End If
                        TBCorretiva.Close
                    End If
                    
                    Set TBCorretiva = CreateObject("adodb.recordset")
                    TBCorretiva.Open "Select UF from Empresa where codigo = " & IDempresaST, Conexao, adOpenKeyset, adLockReadOnly
                        If TBCorretiva.EOF = False Then UFempresa = IIf(IsNull(TBCorretiva!UF), "", TBCorretiva!UF)
                        TBCorretiva.Close
                        
                            If ICMSSTSimples = True Then 'entra quando for do estado de MT e cliente marcado como ST simplificado
                                'A (Valor total da operaÁ„o)
                                'VlrBCST = Vlr. total produto - Vlr. total desc. + Vlr. IPI (se o usu·rio configurar) + Vlr. frete
                                'Vlr BCST + Vlr. seguro + Vlr. acessorias
                                ValorConta = VlrBCST + VlrSeguro + VlrAcessorias
            
                                'B (AlÌquota interno do ICMS normal)
                                ICMS = vRegiao(0, 1)
            
                                'C (Valor de ICMS normal)
                                TTICMS = Format((Format((VlrUnit * QtdeProd), "0.00") * ICMS) / 100, "0.00")
            
                                'D (AlÌquota do ICMS ST)
                                ICMSST = IIf(IsNull(TBFI!Aliquota), 0, TBFI!Aliquota)
            
                                'E (Margem do ICMS ST)
                                QtdeSaida = 16
            
                                'F (Valor do ICMS ST)
                                ICMSCST = Format((ValorConta * QtdeSaida) / 100, "0.00")
            
                                'G (Base do ICMS ST)
                                BCICMSCST = Format(((TTICMS + ICMSCST) / ICMSST) * 100, "0.00")
 '==============================================================================================
 ' CST 0202 - Se cliente for contribuinte icms e consumidor final e for de fora do estado e usar CST 202
 '==============================================================================================
                            ElseIf Right(CST, 3) = "202" And naoContribuinteICMS = False And consumidorFinal = False And UFempresa <> UFcliente Then 'Entra quando a CST È 202 + consumidor final + contribuinte de ICMS + estados diferentes
                                'Valor total da operaÁ„o
                                'VlrBCST = Vlr. total produto - Vlr. total desc. + Vlr. IPI (se o usu·rio configurar) + Vlr. frete
                                'Vlr BCST + Vlr. seguro + Vlr. acessorias
                                ValorConta = VlrBCST + VlrSeguro + VlrAcessorias
                                
                                'Aliquota destino
                                ICMS = vRegiao(0, 1) / 100 'Aliquota interestadual

                                'Aliquota interna
                                TTICMS = IIf(IsNull(TBFI!Aliquota), 0, TBFI!Aliquota)
                                ICMS2 = TTICMS / 100
                                
                                

                                'Calculo para base de calculo do Difal
                                QtdeSaida = Format(ValorConta - (ValorConta * ICMS), "0.00")
                                BCICMSCST = Format(QtdeSaida / (1 - ICMS2), "0.00")
                                QtdeSaida = 0
                                
                                'Valor ICMS Difal
                                ICMSCST = Format((BCICMSCST * ICMS2) - (ValorConta * ICMS), "0.00")
                            Else
                                'A (Valor total da operaÁ„o)
                                'VlrBCST = Vlr. total produto - Vlr. total desc. + Vlr. IPI (se o usu·rio configurar) + Vlr. frete
                                'Vlr BCST + Vlr. seguro + Vlr. acessorias
                                ValorConta = VlrBCST + VlrSeguro + VlrAcessorias
                                
                                'B (Margem do ICMS ST)
                                QtdeSaida = IIf(IsNull(TBFI!Margem), 0, TBFI!Margem)
                                
                                'C (Base do ICMS ST)
                                BCICMSCST = Format(((ValorConta * QtdeSaida) / 100) + ValorConta, "0.0000")
                                'Debug.print BCICMSCST
                                
                                '===========================================
                                ' ReduÁ„o base de calculo ICMS ST
                                '===========================================
                                If Right(CST, 2) = "70" Or Right(CST, 3) = "201" Or Right(CST, 3) = "202" Or Right(CST, 3) = "203" Then
                                    If TBFI!Aliquota_aplicavel <> 0 And IsNull(TBFI!Aliquota_aplicavel) = False Then
                                        
                                        valor = 100 - ((TBFI!Aliquota_aplicavel * 100) / TBFI!Aliquota)
                                        valor = Format(valor, "###,##0.00")
        
                                        BCICMSCST = BCICMSCST - ((BCICMSCST * valor) / 100)
                                        'S6 -(S6 * R6)
                                    'Debug.print BCICMSCST
                                    End If
                                End If

                                '============================================
                                'D (Aliquota interna do ICMS)
                                TTICMS = IIf(IsNull(TBFI!Aliquota), 0, TBFI!Aliquota)
                                
                                'E (Valor de ICMS normal
                                ValorConta = IIf(VlrBC = 0, Format((VlrUnit * QtdeProd) + VlrSeguro + VlrAcessorias + VlrFrete, "0.0000"), VlrBC)
                                ICMS1 = Format((ValorConta * vRegiao(0, 1)) / 100, "0.00")
                                                
                                ICMS2 = Format((BCICMSCST * TTICMS) / 100, "0.00")
                                ICMSCST = Format(ICMS2 - ICMS1, "0.00")
                            End If
                    End If
            End If
        TBFI.Close
    End If
    TBMaquinas.Close
End If

If NF = True Then
    NovoValor = Replace(QtdeSaida, ",", ".")
    NovoValor1 = Replace(BCICMSCST, ",", ".")
    NovoValor2 = Replace(TTICMS, ",", ".")
    NovoValor3 = Replace(ICMSCST, ",", ".")
    Conexao.Execute "Update tbl_Detalhes_Nota_CST_ICMS Set Percentual_margem_ICMS_ST = " & NovoValor & ", Valor_BC_ST = " & NovoValor1 & ", Aliquota_imposto_ST = " & NovoValor2 & ", Valor_ICMS_ST = " & NovoValor3 & " where ID_item = " & TBProduto!Int_codigo
End If

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Public Function FunInternetConectada(Optional ByRef ConnType As String) As Boolean
On Error GoTo tratar_erro
Dim dwFlags As Long
Dim WebTest As Boolean

ConnType = ""
WebTest = InternetGetConnectedState(dwFlags, 0&)
Select Case WebTest
    Case dwFlags And CONNECT_LAN: ConnType = "LAN"
    Case dwFlags And CONNECT_MODEM: ConnType = "Modem"
    Case dwFlags And CONNECT_PROXY: ConnType = "Proxy"
    Case dwFlags And CONNECT_OFFLINE: ConnType = "Offline"
    Case dwFlags And CONNECT_CONFIGURED: ConnType = "Configurada"
    Case dwFlags And CONNECT_RAS: ConnType = "Remota"
End Select
FunInternetConectada = WebTest

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Sub ProcEnviarEmailAutomatico(MAPIS As MAPISession, MAPIM As MAPIMessages, Endereco As String, Assunto As String, Mensagem As String, CaminhoAnexo As String, Nome_anexo As String)
On Error GoTo tratar_erro
Const SESSION_SIGNON = 1
Const MESSAGE_COMPOSE = 6
Const ATTACHTYPE_DATA = 0
Const RECIPTYPE_TO = 1
Const RECIPTYPE_CC = 2
Const MESSAGE_RESOLVENAME = 13
Const MESSAGE_SEND = 3
Const SESSION_SIGNOFF = 2

Inicio:
    MAPIS.Action = SESSION_SIGNON 'Abre up a MAPI session
    With MAPIM
        .SessionID = MAPIS.SessionID
        .Action = MESSAGE_COMPOSE 'Cria uma nova mensagem
        .MsgSubject = Assunto
        .MsgNoteText = Mensagem
        
        If CaminhoAnexo <> "" Then
            .AttachmentPosition = 0 'Anexa no final da mensagem
            .AttachmentType = ATTACHTYPE_DATA 'Define tipo do anexo
            .AttachmentName = Nome_anexo 'Nome do arquivo
            .AttachmentPathName = CaminhoAnexo 'Caminho e nome do arquivo a anexar
        End If
        
        'Verifica quantidade de e-mails
        Texto = ""
        qt = 0
        Numero1 = Len(Endereco)
        Contador = 0
        Endereco1 = ""
        Do While Numero1 <> 0
            If Texto = ";" Then GoTo Pula
Proximo:
            Texto = Left(Endereco, (qt + 1))
            Texto = Right(Texto, Len(Texto) - qt)
            
            If Texto <> ";" And Texto <> " " Then
                If Endereco1 = "" Then Endereco1 = Texto Else Endereco1 = Endereco1 & Texto
            End If
            
            qt = qt + 1
            Numero1 = Numero1 - 1
        Loop
        
Pula:
        'Destinat·rio
        .RecipIndex = Contador 'ID
        .RecipType = RECIPTYPE_TO
        .RecipDisplayName = Endereco1 'E-mail destinat·rio
        
        If Numero1 <> 0 Then
            Endereco1 = ""
            Contador = Contador + 1
            GoTo Proximo
        End If
        
        '.RecipIndex = 1 'ID
        '.RecipType = RECIPTYPE_TO
        '.RecipDisplayName = Endereco 'E-mail destinat·rio
        
        '.Action = MESSAGE_RESOLVENAME 'Verifica se o e-mail È v·lido
        .Action = MESSAGE_SEND 'Envia a mensagem
    End With
1:
    MAPIS.Action = SESSION_SIGNOFF 'Fecha up a MAPI session

Exit Sub
tratar_erro:
    Permitido = False
    If Err.Number = 32050 Then
        MAPIS.Action = SESSION_SIGNOFF 'Fecha up a MAPI session
        GoTo Inicio
    End If
    If Err.Number = 32014 Then
        USMsgBox ("O e-mail informado n„o È v·lido, favor alterar."), vbExclamation, "CAPRIND v5.0"
        GoTo 1
    End If
        
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Function FunAbrirVideoWeb(EnderecoWeb As String)
On Error GoTo tratar_erro

'Chat = False
'Video_ajuda = True
'With Frm_web
'    .Show
'    .Web.Navigate EnderecoWeb
'End With

Set IE = New InternetExplorer
IE.Navigate EnderecoWeb
IE.Visible = True

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Public Function FunBaixarArquivoNET(url As String, LocalFileName As String)
On Error GoTo tratar_erro

lngRetVal = URLDownloadToFile(0, url, LocalFileName, 0, 0)
If lngRetVal = 0 Then Permitido1 = True

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Public Function FunCalculaAmostragem(Combo As ComboBox, valor As Double)
On Error GoTo tratar_erro

FunCalculaAmostragem = ""
    
If valor <= 0 Then
    FunCalculaAmostragem = "0,0000"
    Exit Function
End If

Select Case Combo
    Case "0": FunCalculaAmostragem = Format(valor, "###,##0.0000")
    Case "S1":
        If valor > 1 And valor < 9 Then FunCalculaAmostragem = "2,000"
        If valor > 8 And valor < 16 Then FunCalculaAmostragem = "2,000"
        If valor > 15 And valor < 26 Then FunCalculaAmostragem = "2,000"
        If valor > 25 And valor < 51 Then FunCalculaAmostragem = "2,000"
        If valor > 50 And valor < 91 Then FunCalculaAmostragem = "3,000"
        If valor > 90 And valor < 151 Then FunCalculaAmostragem = "3,000"
        If valor > 150 And valor < 281 Then FunCalculaAmostragem = "3,000"
        If valor > 280 And valor < 501 Then FunCalculaAmostragem = "3,000"
        If valor > 500 And valor < 1201 Then FunCalculaAmostragem = "5,000"
        If valor > 1200 And valor < 3201 Then FunCalculaAmostragem = "5,000"
        If valor > 3200 And valor < 10001 Then FunCalculaAmostragem = "5,000"
        If valor > 10000 And valor < 35001 Then FunCalculaAmostragem = "5,000"
        If valor > 35000 And valor < 150001 Then FunCalculaAmostragem = "8,000"
        If valor > 150000 And valor < 500001 Then FunCalculaAmostragem = "8,000"
        If valor > 500000 Then FunCalculaAmostragem = "8,000"
    Case "S2":
        If valor > 1 And valor < 9 Then FunCalculaAmostragem = "2,000"
        If valor > 8 And valor < 16 Then FunCalculaAmostragem = "2,000"
        If valor > 15 And valor < 26 Then FunCalculaAmostragem = "2,000"
        If valor > 25 And valor < 51 Then FunCalculaAmostragem = "3,000"
        If valor > 50 And valor < 91 Then FunCalculaAmostragem = "3,000"
        If valor > 90 And valor < 151 Then FunCalculaAmostragem = "3,000"
        If valor > 150 And valor < 281 Then FunCalculaAmostragem = "5,000"
        If valor > 280 And valor < 501 Then FunCalculaAmostragem = "5,000"
        If valor > 500 And valor < 1201 Then FunCalculaAmostragem = "5,000"
        If valor > 1200 And valor < 3201 Then FunCalculaAmostragem = "8,000"
        If valor > 3200 And valor < 10001 Then FunCalculaAmostragem = "8,000"
        If valor > 10000 And valor < 35001 Then FunCalculaAmostragem = "8,000"
        If valor > 35000 And valor < 150001 Then FunCalculaAmostragem = "13,000"
        If valor > 150000 And valor < 500001 Then FunCalculaAmostragem = "13,000"
        If valor > 500000 Then FunCalculaAmostragem = "13,000"
    Case "S3":
        If valor > 1 And valor < 9 Then FunCalculaAmostragem = "2,000"
        If valor > 8 And valor < 16 Then FunCalculaAmostragem = "2,000"
        If valor > 15 And valor < 26 Then FunCalculaAmostragem = "3,000"
        If valor > 25 And valor < 51 Then FunCalculaAmostragem = "3,000"
        If valor > 50 And valor < 91 Then FunCalculaAmostragem = "5,000"
        If valor > 90 And valor < 151 Then FunCalculaAmostragem = "5,000"
        If valor > 150 And valor < 281 Then FunCalculaAmostragem = "8,000"
        If valor > 280 And valor < 501 Then FunCalculaAmostragem = "8,000"
        If valor > 500 And valor < 1201 Then FunCalculaAmostragem = "13,000"
        If valor > 1200 And valor < 3201 Then FunCalculaAmostragem = "13,000"
        If valor > 3200 And valor < 10001 Then FunCalculaAmostragem = "13,000"
        If valor > 10000 And valor < 35001 Then FunCalculaAmostragem = "20,000"
        If valor > 35000 And valor < 150001 Then FunCalculaAmostragem = "32,000"
        If valor > 150000 And valor < 500001 Then FunCalculaAmostragem = "32,000"
        If valor > 500000 Then FunCalculaAmostragem = "50,000"
    Case "S4":
        If valor > 1 And valor < 9 Then FunCalculaAmostragem = "2,000"
        If valor > 8 And valor < 16 Then FunCalculaAmostragem = "2,000"
        If valor > 15 And valor < 26 Then FunCalculaAmostragem = "3,000"
        If valor > 25 And valor < 51 Then FunCalculaAmostragem = "5,000"
        If valor > 50 And valor < 91 Then FunCalculaAmostragem = "5,000"
        If valor > 90 And valor < 151 Then FunCalculaAmostragem = "8,000"
        If valor > 150 And valor < 281 Then FunCalculaAmostragem = "13,000"
        If valor > 280 And valor < 501 Then FunCalculaAmostragem = "13,000"
        If valor > 500 And valor < 1201 Then FunCalculaAmostragem = "20,000"
        If valor > 1200 And valor < 3201 Then FunCalculaAmostragem = "32,000"
        If valor > 3200 And valor < 10001 Then FunCalculaAmostragem = "32,000"
        If valor > 10000 And valor < 35001 Then FunCalculaAmostragem = "50,000"
        If valor > 35000 And valor < 150001 Then FunCalculaAmostragem = "80,000"
        If valor > 150000 And valor < 500001 Then FunCalculaAmostragem = "80,000"
        If valor > 500000 Then FunCalculaAmostragem = "125,000"
    Case "I":
        If valor > 1 And valor < 9 Then FunCalculaAmostragem = "2,000"
        If valor > 8 And valor < 16 Then FunCalculaAmostragem = "2,000"
        If valor > 15 And valor < 26 Then FunCalculaAmostragem = "3,000"
        If valor > 25 And valor < 51 Then FunCalculaAmostragem = "5,000"
        If valor > 50 And valor < 91 Then FunCalculaAmostragem = "5,000"
        If valor > 90 And valor < 151 Then FunCalculaAmostragem = "8,000"
        If valor > 150 And valor < 281 Then FunCalculaAmostragem = "13,000"
        If valor > 280 And valor < 501 Then FunCalculaAmostragem = "20,000"
        If valor > 500 And valor < 1201 Then FunCalculaAmostragem = "32,000"
        If valor > 1200 And valor < 3201 Then FunCalculaAmostragem = "50,000"
        If valor > 3200 And valor < 10001 Then FunCalculaAmostragem = "80,000"
        If valor > 10000 And valor < 35001 Then FunCalculaAmostragem = "125,000"
        If valor > 35000 And valor < 150001 Then FunCalculaAmostragem = "200,000"
        If valor > 150000 And valor < 500001 Then FunCalculaAmostragem = "315,000"
        If valor > 500000 Then FunCalculaAmostragem = "500,000"
    Case "II":
        If valor > 1 And valor < 9 Then FunCalculaAmostragem = "2,000"
        If valor > 8 And valor < 16 Then FunCalculaAmostragem = "3,000"
        If valor > 15 And valor < 26 Then FunCalculaAmostragem = "5,000"
        If valor > 25 And valor < 51 Then FunCalculaAmostragem = "8,000"
        If valor > 50 And valor < 91 Then FunCalculaAmostragem = "13,000"
        If valor > 90 And valor < 151 Then FunCalculaAmostragem = "20,000"
        If valor > 150 And valor < 281 Then FunCalculaAmostragem = "32,000"
        If valor > 280 And valor < 501 Then FunCalculaAmostragem = "50,000"
        If valor > 500 And valor < 1201 Then FunCalculaAmostragem = "80,000"
        If valor > 1200 And valor < 3201 Then FunCalculaAmostragem = "125,000"
        If valor > 3200 And valor < 10001 Then FunCalculaAmostragem = "200,000"
        If valor > 10000 And valor < 35001 Then FunCalculaAmostragem = "315,000"
        If valor > 35000 And valor < 150001 Then FunCalculaAmostragem = "500,000"
        If valor > 150000 And valor < 500001 Then FunCalculaAmostragem = "800,000"
        If valor > 500000 Then FunCalculaAmostragem = "1250,000"
    Case "III":
        If valor > 1 And valor < 9 Then FunCalculaAmostragem = "3,000"
        If valor > 8 And valor < 16 Then FunCalculaAmostragem = "5,000"
        If valor > 15 And valor < 26 Then FunCalculaAmostragem = "8,000"
        If valor > 25 And valor < 51 Then FunCalculaAmostragem = "13,000"
        If valor > 50 And valor < 91 Then FunCalculaAmostragem = "20,000"
        If valor > 90 And valor < 151 Then FunCalculaAmostragem = "32,000"
        If valor > 150 And valor < 281 Then FunCalculaAmostragem = "50,000"
        If valor > 280 And valor < 501 Then FunCalculaAmostragem = "80,000"
        If valor > 500 And valor < 1201 Then FunCalculaAmostragem = "125,000"
        If valor > 1200 And valor < 3201 Then FunCalculaAmostragem = "200,000"
        If valor > 3200 And valor < 10001 Then FunCalculaAmostragem = "315,000"
        If valor > 10000 And valor < 35001 Then FunCalculaAmostragem = "500,000"
        If valor > 35000 And valor < 150001 Then FunCalculaAmostragem = "800,000"
        If valor > 150000 And valor < 500001 Then FunCalculaAmostragem = "1250,000"
        If valor > 500000 Then
            ValorTotal = 500000
            quantestoque = 1250
            quantidade = valor * quantestoque
            quantidade = quantidade / ValorTotal
            FunCalculaAmostragem = Format(quantidade, "###,##0.0000")
        End If
End Select

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Public Function FunBissexto(Ano As Integer) As Boolean
On Error GoTo tratar_erro

FunBissexto = False
If Ano Mod 4 = 0 Then
   If Ano Mod 100 = 0 Then
      If intAno Mod 400 = 0 Then FunBissexto = True
   Else
        FunBissexto = True
   End If
End If

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Public Function FunPrimeiraLetraMaiuscula(Texto As String) As String
On Error GoTo tratar_erro

FunPrimeiraLetraMaiuscula = Strings.StrConv(Texto, VbStrConv.vbProperCase)

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Public Function FunVerificaQtdeEstoque(Codinterno As String, ID_empresa As Integer, Pesquisa As String) As Double
On Error GoTo tratar_erro

FunVerificaQtdeEstoque = 0
QuantEmpenho = 0
quantestoque = 0
qt = 0
QuantEmpenhoPC = 0
quantnovo = 0
Valor_total = 0
CTMedioEst = 0
Set TBComponente = CreateObject("adodb.recordset")

If ID_empresa <> 0 Then TextoFiltro = "desenho = '" & Codinterno & "' and ID_empresa = " & ID_empresa Else TextoFiltro = "desenho = '" & Codinterno & "'"
TBComponente.Open "Select Sum(Estoque_disponivel) as Qtde, Sum(Qtde_empenhada) as QuantEmpenho, Sum(Estoque_disponivel) as quantestoque, Sum(estoque_real_PC) as Qt, Sum(Qtde_empenhada_PC) as QuantEmpenhoPC, Sum(Estoque_disponivel_PC) as quantnovo, Sum(Valor_total) as Valor_Total from Qtde_estoque_produto where  " & TextoFiltro & " " & Pesquisa, Conexao, adOpenKeyset, adLockOptimistic
If TBComponente.EOF = False Then
    FunVerificaQtdeEstoque = IIf(IsNull(TBComponente!Qtde), 0, TBComponente!Qtde)
    QuantEmpenho = IIf(IsNull(TBComponente!QuantEmpenho), 0, TBComponente!QuantEmpenho)
    quantestoque = IIf(IsNull(TBComponente!quantestoque), 0, TBComponente!quantestoque)
    qt = IIf(IsNull(TBComponente!qt), 0, TBComponente!qt)
    QuantEmpenhoPC = IIf(IsNull(TBComponente!QuantEmpenhoPC), 0, TBComponente!QuantEmpenhoPC)
    quantnovo = IIf(IsNull(TBComponente!quantnovo), 0, TBComponente!quantnovo)
    Valor_total = IIf(IsNull(TBComponente!Valor_total), 0, TBComponente!Valor_total)
    
    If IIf(IsNull(TBComponente!Qtde), 0, TBComponente!Qtde) > 0 Then CTMedioEst = Format(Valor_total / TBComponente!Qtde, "###,##0.00000000") Else CTMedioEst = 0
End If
TBComponente.Close

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Public Function FunVerificaNecessidadeVenda(Codinterno As String, ID_empresa As Integer) As Double
On Error GoTo tratar_erro

FunVerificaNecessidadeVenda = 0
Set TBComponente = CreateObject("adodb.recordset")
TBComponente.Open "Select * from Necessidade_vendas_produto where desenho = '" & Codinterno & "' and ID_empresa = " & ID_empresa, Conexao, adOpenKeyset, adLockOptimistic
If TBComponente.EOF = False Then
    FunVerificaNecessidadeVenda = IIf(IsNull(TBComponente!Necessidade), 0, TBComponente!Necessidade)
End If
TBComponente.Close

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Public Function FunVerificaQtdeEmpenhoEstVenda(Codinterno As String, ID_empresa As Integer) As Double
On Error GoTo tratar_erro

FunVerificaQtdeEmpenhoEstVenda = 0
Set TBComponente = CreateObject("adodb.recordset")
TBComponente.Open "Select Sum(Qtde_requisitar) as Qtde_requisitar from Qtde_empenhada_produto_venda where desenho = '" & Codinterno & "' and ID_empresa = " & ID_empresa, Conexao, adOpenKeyset, adLockOptimistic
If TBComponente.EOF = False Then
    FunVerificaQtdeEmpenhoEstVenda = IIf(IsNull(TBComponente!Qtde_requisitar), 0, TBComponente!Qtde_requisitar)
End If
TBComponente.Close

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Public Function FunVerificaQtdeEmpenhoEst(Codinterno As String, ID_empresa As Integer) As Double
On Error GoTo tratar_erro

FunVerificaQtdeEmpenhoEst = 0
Set TBComponente = CreateObject("adodb.recordset")
TBComponente.Open "Select Sum(Qtde_empenhar) as Qtde_empenhar from Qtde_empenhada_produto where Codigo = '" & Codinterno & "' and ID_empresa = " & ID_empresa, Conexao, adOpenKeyset, adLockOptimistic
If TBComponente.EOF = False Then
    FunVerificaQtdeEmpenhoEst = IIf(IsNull(TBComponente!Qtde_empenhar), 0, TBComponente!Qtde_empenhar)
End If
TBComponente.Close

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Public Function FunVerificaQtdeProduzindo(Codinterno As String, ID_empresa As Integer) As Double
On Error GoTo tratar_erro

FunVerificaQtdeProduzindo = 0
Set TBComponente = CreateObject("adodb.recordset")
TBComponente.Open "Select Sum(Qtde_produzindo) as Qtde_produzindo from Qtde_produzindo_produto where desenho = '" & Codinterno & "' and ID_empresa = " & ID_empresa, Conexao, adOpenKeyset, adLockOptimistic
If TBComponente.EOF = False Then
    FunVerificaQtdeProduzindo = IIf(IsNull(TBComponente!Qtde_produzindo), 0, TBComponente!Qtde_produzindo)
End If
TBComponente.Close

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Public Function FunVerificaQtdeEmpenhoProduzindo(Codinterno As String, ID_empresa As Integer) As Double
On Error GoTo tratar_erro

FunVerificaQtdeEmpenhoProduzindo = 0
Set TBComponente = CreateObject("adodb.recordset")
TBComponente.Open "Select Sum(Qtde_requisitar) as Qtde_requisitar from Qtde_empenhada_produto_produzindo where desenho = '" & Codinterno & "' and ID_empresa = " & ID_empresa, Conexao, adOpenKeyset, adLockOptimistic
If TBComponente.EOF = False Then
    FunVerificaQtdeEmpenhoProduzindo = IIf(IsNull(TBComponente!Qtde_requisitar), 0, TBComponente!Qtde_requisitar)
End If
TBComponente.Close

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Public Function FunVerificaQtdeEmpenhoREOrdem(TextoFiltro As String, VerifQtdePC As Boolean) As Double
On Error GoTo tratar_erro

FunVerificaQtdeEmpenhoREOrdem = 0
If VerifQtdePC = True Then CamposFiltro = "Qtde_empenhar_PC" Else CamposFiltro = "Qtde_empenhar"
Set TBComponente = CreateObject("adodb.recordset")
TBComponente.Open "Select Sum(" & CamposFiltro & ") as Qtde_requisitar from Qtde_empenhada_produto_detalhado where " & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
If TBComponente.EOF = False Then
    FunVerificaQtdeEmpenhoREOrdem = IIf(IsNull(TBComponente!Qtde_requisitar), 0, TBComponente!Qtde_requisitar)
End If
TBComponente.Close

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Public Function FunVerificaQtdeEmpenhoREPI(TextoFiltro As String) As Double
On Error GoTo tratar_erro

FunVerificaQtdeEmpenhoREPI = 0
Set TBComponente = CreateObject("adodb.recordset")
TBComponente.Open "Select Sum(Qtde_requisitar) as Qtde_requisitar from Qtde_empenhada_produto_venda_detalhado where " & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
If TBComponente.EOF = False Then
    FunVerificaQtdeEmpenhoREPI = IIf(IsNull(TBComponente!Qtde_requisitar), 0, TBComponente!Qtde_requisitar)
End If
TBComponente.Close

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Public Function FunVerificaTabelaConversaoUnidade(Unidade_de As String, Unidade_para As String) As Double
On Error GoTo tratar_erro

FunVerificaTabelaConversaoUnidade = 1
If Unidade_de <> Unidade_para Then
    Set TBCorretiva = CreateObject("adodb.recordset")
    TBCorretiva.Open "Select * from Tabela_conversao_unidade where Unidade_de = '" & Unidade_de & "' and Unidade_para = '" & Unidade_para & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBCorretiva.EOF = False Then
        FunVerificaTabelaConversaoUnidade = IIf(IsNull(TBCorretiva!Qtde_para), 1, TBCorretiva!Qtde_para)
    End If
    TBCorretiva.Close
End If

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Public Function FunVerifValidadeCertForn(ID_empresa As Integer, Data_emissao As Date, MostrarMsg As Boolean) As Boolean
On Error GoTo tratar_erro

FunVerifValidadeCertForn = True
If IsNull(TBFornecedor!Data_venc) = True Or TBFornecedor!Data_venc = "" Or IsNull(TBFornecedor!Fornecedor) = True Or TBFornecedor!Fornecedor = "" Then
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select Codigo from Empresa where Codigo = " & ID_empresa & " and Bloquear_fornecedores = 'True'", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        If MostrarMsg = True Then USMsgBox ("N„o È permitido utlilizar este fornecedor, pois o mesmo n„o est· homologado."), vbExclamation, "CAPRIND v5.0"
        FunVerifValidadeCertForn = False
    End If
    TBAbrir.Close
Else
    Dataini = Data_emissao
    If TBFornecedor!Data_venc < Dataini Then
        If MostrarMsg = True Then
            If TBFornecedor!Fornecedor = "A" Then
                NomeCampo = "a avaliaÁ„o vencida"
            ElseIf TBFornecedor!Fornecedor = "A" Then
                NomeCampo = "o certificado vencido"
            Else
                NomeCampo = "a fonte aprovada pelo cliente vencida"
            End If
            USMsgBox ("N„o È permitido utlilizar este fornecedor, pois o mesmo est· com " & NomeCampo & "."), vbExclamation, "CAPRIND v5.0"
        End If
        FunVerifValidadeCertForn = False
    End If
End If

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Public Function FunVerifRegimeTribCliForn(IDempresa As Integer, Fornecedor As Boolean, MostrarMsg As Boolean) As Boolean
On Error GoTo tratar_erro

FunVerifRegimeTribCliForn = True
Permitido = True
If Fornecedor = True Then
    NomeCampo = "fornecedor"
    If TBFornecedor!Simples = False And TBFornecedor!Presumido = False And TBFornecedor!Real = False Then Permitido = False
Else
    NomeCampo = "cliente"
    If TBClientes!Simples = False And TBClientes!Presumido = False And TBClientes!Real = False Then Permitido = False
End If
If Permitido = False Then
    Set TBFI = CreateObject("adodb.recordset")
    TBFI.Open "Select * from Empresa where Codigo = " & IDempresa & " and Bloquear_cli_forn_regime = 'True'", Conexao, adOpenKeyset, adLockOptimistic
    If TBFI.EOF = False Then
        If MostrarMsg = True Then USMsgBox ("N„o È permitido utlilizar este " & NomeCampo & ", pois n„o foi informado o regime tribut·rio em seu cadastro."), vbExclamation, "CAPRIND v5.0"
        FunVerifRegimeTribCliForn = False
        Exit Function
    Else
'        If USMsgBox("N„o foi informado o regime tribut·rio do " & NomeCampo & ", deseja prosseguir assim mesmo?", vbyesno, "CAPRIND v5.0") = vbNo Then
'            FunVerifRegimeTribCliForn = False
'            Exit Function
'        End If
    End If
    TBFI.Close
End If

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Sub ProcGravarDataFiltroRel(DataInicio As Date, DataFinal As Date, MostrarDataRel As Boolean, ID_empresa As Integer, Texto As String)
On Error GoTo tratar_erro

ProcExcluirDadosProducaoRelatoriosTotal

Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from Producao_Relatorios_Total where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then TBGravar.AddNew
TBGravar!Data_inicial = DataInicio
TBGravar!Data_final = DataFinal
TBGravar!Responsavel = pubUsuario
TBGravar!Modulo = Formulario
If MostrarDataRel = True Then TBGravar!Valor1 = 1 Else TBGravar!Valor1 = 0
TBGravar!QtdePrevista = ID_empresa
TBGravar!Texto = IIf(Texto = "", Null, Texto)
TBGravar.Update
TBGravar.Close

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Function FunAtualizaStatusPropPI(Cotacao As Long) As Boolean
On Error GoTo tratar_erro

FunAtualizaStatusPropPI = False

StatusTexto1 = ""
'VERIFICA STATUS = ABERTA EM ANALISE
Set TBAliquota = CreateObject("adodb.recordset")
TBAliquota.Open "Select Codigo from vendas_carteira where cotacao = " & Cotacao & " and liberacao <> 'ABERTA EM ANALISE'", Conexao, adOpenKeyset, adLockOptimistic
If TBAliquota.EOF = True Then
    StatusTexto = "ABERTA EM ANALISE"
    StatusTexto1 = ", Tipo = 'PR'"
    FunAtualizaStatusPropPI = True
Else
    'VERIFICA STATUS = CANCELADO
    Set TBAliquota = CreateObject("adodb.recordset")
    TBAliquota.Open "Select Codigo from vendas_carteira where cotacao = " & Cotacao & " and liberacao <> 'CANCELADO'", Conexao, adOpenKeyset, adLockOptimistic
    If TBAliquota.EOF = True Then
        StatusTexto = "CANCELADA"
    Else
        'VERIFICA STATUS = FATURADA
        Set TBAliquota = CreateObject("adodb.recordset")
        TBAliquota.Open "Select Codigo from vendas_carteira where cotacao = " & Cotacao & " and liberacao <> 'FATURADO' and liberacao <> 'CANCELADO'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAliquota.EOF = True Then
            StatusTexto = "FATURADA"
        Else
            'VERIFICA STATUS = FATURADA PARCIAL
            Set TBAliquota = CreateObject("adodb.recordset")
            TBAliquota.Open "Select Codigo from vendas_carteira where cotacao = " & Cotacao & " and (liberacao = 'FATURADO' or liberacao = 'FATURADO PARCIAL')", Conexao, adOpenKeyset, adLockOptimistic
            If TBAliquota.EOF = False Then
                Set TBProduto = CreateObject("adodb.recordset")
                TBProduto.Open "Select * from vendas_carteira where cotacao = " & Cotacao, Conexao, adOpenKeyset, adLockOptimistic
                If TBAliquota.RecordCount <= TBProduto.RecordCount Then StatusTexto = "FATURADA PARCIAL"
                GoTo Prosseguir
            End If
            'VERIFICA STATUS = VENDIDA or VENDIDA PARCIAL
            Set TBAliquota = CreateObject("adodb.recordset")
            TBAliquota.Open "Select Codigo from vendas_carteira where cotacao = " & Cotacao & " and (liberacao = 'VENDIDA' or liberacao = 'FATURAR' or liberacao = 'FATURAR PARCIAL')", Conexao, adOpenKeyset, adLockOptimistic
            If TBAliquota.EOF = False Then
                Set TBProduto = CreateObject("adodb.recordset")
                TBProduto.Open "Select * from vendas_carteira where cotacao = " & Cotacao, Conexao, adOpenKeyset, adLockOptimistic
                If TBAliquota.RecordCount < TBProduto.RecordCount Then
                    StatusTexto = "VENDIDA PARCIAL"
                Else
                    StatusTexto = "VENDIDA"
                End If
            End If
        End If
    End If
End If
Prosseguir:
    TBAliquota.Close
    Conexao.Execute "UPDATE vendas_proposta Set Status = '" & StatusTexto & "' " & StatusTexto1 & " where cotacao = " & Cotacao

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Function FunAtualizaStatusPC(IDpedido As Long) As String
On Error GoTo tratar_erro

Set TBPedido = CreateObject("adodb.recordset")
TBPedido.Open "Select CP.Status_pedido, CPL.IDlista from compras_pedido CP LEFT JOIN compras_pedido_lista CPL ON CP.Idpedido = CPL.Idpedido where CP.IDpedido = " & IDpedido, Conexao, adOpenKeyset, adLockOptimistic
If TBPedido.EOF = False Then
    If IsNull(TBPedido!IDlista) = False Then
        Set TBCompras = CreateObject("adodb.recordset")
        TBCompras.Open "Select IDlista from compras_pedido_lista where idpedido = " & IDpedido & " and Status_Item <> 'CANCELADO'", Conexao, adOpenKeyset, adLockOptimistic
        If TBCompras.EOF = True Then
            TBPedido!Status_pedido = "CANCELADO"
        Else
            Set TBCompras = CreateObject("adodb.recordset")
            TBCompras.Open "Select IDlista from compras_pedido_lista where idpedido = " & IDpedido & " and Status_Item <> 'N√O APROVADO' and status_item <> 'CANCELADO'", Conexao, adOpenKeyset, adLockOptimistic
            If TBCompras.EOF = True Then
                TBPedido!Status_pedido = "N√O APROVADO"
            Else
                Set TBCompras = CreateObject("adodb.recordset")
                TBCompras.Open "Select IDlista from compras_pedido_lista where idpedido = " & IDpedido & " and status_item <> 'APROVADO' and status_item <> 'CANCELADO'", Conexao, adOpenKeyset, adLockOptimistic
                If TBCompras.EOF = True Then
                    TBPedido!Status_pedido = "APROVADO"
                Else
                    Set TBCompras = CreateObject("adodb.recordset")
                    TBCompras.Open "Select IDlista from compras_pedido_lista where idpedido = " & IDpedido & " and status_item <> 'RECEBIDO' and status_item <> 'CANCELADO'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBCompras.EOF = True Then
                        TBPedido!Status_pedido = "ENCERRADO"
                    Else
                        Set TBCompras = CreateObject("adodb.recordset")
                        TBCompras.Open "Select IDlista from compras_pedido_lista where idpedido = " & IDpedido & " and Status_Item <> 'AGUARDANDO APROVA«√O' and Status_Item <> 'CANCELADO'", Conexao, adOpenKeyset, adLockOptimistic
                        If TBCompras.EOF = True Then
                            TBPedido!Status_pedido = "AGUARDANDO APROVA«√O"
                        Else
                            Set TBCompras = CreateObject("adodb.recordset")
                            TBCompras.Open "Select IDlista from compras_pedido_lista where idpedido = " & IDpedido & " and Status_Item <> 'N_RECEBIDO' and Status_Item <> 'CANCELADO'", Conexao, adOpenKeyset, adLockOptimistic
                            If TBCompras.EOF = True Then
                                TBPedido!Status_pedido = "ABERTO"
                            Else
                                TBPedido!Status_pedido = "PARCIAL"
                            End If
                        End If
                    End If
                End If
            End If
        End If
        TBCompras.Close
    Else
        TBPedido!Status_pedido = "AGUARDANDO APROVA«√O"
    End If
    TBPedido.Update
    FunAtualizaStatusPC = TBPedido!Status_pedido
End If
TBPedido.Close

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Function FunVerificaManutencaoAtiva() As Boolean
On Error GoTo tratar_erro

FunVerificaManutencaoAtiva = False
Set TBFIltro = CreateObject("adodb.recordset")
TBFIltro.Open "Select * from Empresa", Conexao, adOpenKeyset, adLockOptimistic
If TBFIltro.EOF = False Then
    Do While TBFIltro.EOF = False
        FunAbreBDSite
        If ConexaoMySql.State = 1 Then
            Set TBMySQL = New ADODB.Recordset
            TBMySQL.Open "Select ID From Clientes Where CNPJ = '" & TBFIltro!CNPJ & "' and Manutencao_ativa = 'SIM'", ConexaoMySql, adOpenKeyset, adLockOptimistic, adCmdText
            If TBMySQL.EOF = False Then
                TBMySQL.Close
                FunVerificaManutencaoAtiva = True
                Exit Function
            End If
            TBMySQL.Close
        End If
        TBFIltro.MoveNext
    Loop
End If
TBFIltro.Close

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Function FunVerificaVersaoAtualizacaoCaprind() As Boolean
On Error GoTo tratar_erro

FunVerificaVersaoAtualizacaoCaprind = True
FunAbreBDSite
If ConexaoMySql.State = 1 Then
    Set TBMySQL = New ADODB.Recordset
    TBMySQL.Open "Select * From Atualizacao_liberada", ConexaoMySql, adOpenKeyset, adLockOptimistic, adCmdText
    If TBMySQL.EOF = False Then
        Set TBFIltro = CreateObject("adodb.recordset")
        TBFIltro.Open "Select * from Versao", Conexao, adOpenKeyset, adLockOptimistic
        If TBFIltro.EOF = False Then
            VersaoAtual = ReturnNumbersOnly(IIf(TBFIltro!versao = "", 0, TBFIltro!versao))
            VersaoNova = ReturnNumbersOnly(IIf(TBMySQL!versao = "", 0, TBMySQL!versao))
            If VersaoAtual = VersaoNova Then
                USMsgBox ("O sistema j· est· atualizado."), vbExclamation, "CAPRIND v5.0"
                TBMySQL.Close
                TBFIltro.Close
                FunVerificaVersaoAtualizacaoCaprind = False
                Exit Function
            End If
        End If
        TBFIltro.Close
    End If
    TBMySQL.Close
End If

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Sub ProcPassaDadosContaCorrenteParaCobreBemX(Carteira As String, Carteira1 As String, Codigocedente As String, ID_empresa As Integer, EmitirBoleto As Boolean, Assunto_email As String)
On Error GoTo tratar_erro

Permitido1 = True
OutrosDadosConfiguracao1 = ""
OutrosDadosConfiguracao2 = ""
SeqRemessa = 0

'Verifica data de emiss„o do boleto
Dataini = Date
Set TBFI = CreateObject("adodb.recordset")
If Financeiro_Contas_Receber = False Then
    TBFI.Open "Select * from tbl_Detalhes_Recebimento where ID = " & IDlista, Conexao, adOpenKeyset, adLockOptimistic
    If TBFI.EOF = False Then
        If IsNull(TBFI!Data_emissao) = False And TBFI!Data_emissao <> "" Then Dataini = TBFI!Data_emissao
    End If
Else
    TBFI.Open "Select * from tbl_Detalhes_Recebimento where IDContaReceber = " & frmContas_Receber.txtidintconta, Conexao, adOpenKeyset, adLockOptimistic
    If TBFI.EOF = False Then
        If IsNull(TBFI!Data_emissao) = False And TBFI!Data_emissao <> "" Then Dataini = TBFI!Data_emissao
    End If
End If
TBFI.Close
Dia = Format(Dataini, "dd")
Mes = Format(Dataini, "mm")
Ano = Format(Dataini, "yyyy")

If Financeiro_Contas_Receber = True Then
    With frmContas_Receber
        Set TBContas = CreateObject("adodb.recordset")
        TBContas.Open "Select * from tbl_contas_receber where IDIntconta = " & IIf(.txtidintconta = "", 0, .txtidintconta), Conexao, adOpenKeyset, adLockOptimistic
        If TBContas.EOF = False Then
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select * from tbl_Instituicoes where txt_Descricao = '" & TBContas!Banco & "' and ID_empresa = " & ID_empresa, Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                Familiatext = TBAbrir!int_NBanco
                Agencia = TBAbrir!txt_Agencia
                ContaCorrente = TBAbrir!txt_Conta
                NomeAgencia = IIf(IsNull(TBAbrir!Nome_agencia), "", TBAbrir!Nome_agencia)
                If retorno = True Then Codigocedente = TBAbrir!Codigo_cedente_registrado
            End If
            TBAbrir.Close
        End If
        TBContas.Close
    End With
Else
    With frmFaturamento_Prod_Serv
        Set TBFI = CreateObject("adodb.recordset")
        TBFI.Open "Select tbl_Detalhes_Recebimento.*, tbl_Dados_Nota_Fiscal.ID_empresa from tbl_Detalhes_Recebimento INNER JOIN tbl_Dados_Nota_Fiscal on tbl_Detalhes_Recebimento.ID_nota = tbl_Dados_Nota_Fiscal.ID where tbl_Detalhes_Recebimento.Id = " & IDlista, Conexao, adOpenKeyset, adLockOptimistic
        If TBFI.EOF = False Then
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select * from tbl_Instituicoes where txt_Descricao = '" & TBFI!txt_Portador_Banco & "' and ID_empresa = " & ID_empresa, Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                Familiatext = TBAbrir!int_NBanco
                Agencia = TBAbrir!txt_Agencia
                ContaCorrente = TBAbrir!txt_Conta
                NomeAgencia = IIf(IsNull(TBAbrir!Nome_agencia), "", TBAbrir!Nome_agencia)
            End If
            TBAbrir.Close
        End If
        TBFI.Close
    End With
End If

'Verifica o ˙ltimo sequencial no banco para gerar o prÛximo
If EmitirBoleto = True Then
    If Financeiro_Contas_Receber = False Then TextoFiltro = frmFaturamento_Prod_Serv.cbo_PortBanco Else TextoFiltro = frmContas_Receber.cmbBanco
    'If Familiatext = "341" Then TextoFiltro1 = " and Data_emissao = '" & Date & "'" Else TextoFiltro1 = ""
    Seq = 1
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select Seq_remessa from tbl_Detalhes_Recebimento where txt_Portador_Banco = '" & TextoFiltro & "' order by Seq_remessa desc", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        If IsNull(TBAbrir!Seq_remessa) = False And TBAbrir!Seq_remessa <> "" Then Seq = TBAbrir!Seq_remessa + 1
    End If
    TBAbrir.Close
Else
    If Financeiro_Contas_Receber = False Then TextoFiltro = "ID = " & IDlista Else TextoFiltro = "IDContaReceber = " & frmContas_Receber.txtidintconta
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select Seq_remessa from tbl_Detalhes_Recebimento where " & TextoFiltro & " and Seq_remessa IS NOT NULL", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Seq = TBAbrir!Seq_remessa
    End If
    TBAbrir.Close
End If
If Remessa = True Then
    If Financeiro_Contas_Receber = False Then TextoFiltro = "Id = " & IDlista Else TextoFiltro = "IDContaReceber = " & frmContas_Receber.txtidintconta
    Conexao.Execute "UPDATE tbl_Detalhes_Recebimento Set Seq_remessa = '" & Seq & "' where " & TextoFiltro
End If
If Seq < 10 Then SeqRemessa = "0" & Seq & ".txt" Else SeqRemessa = Seq & ".txt"

'InÌcio dos par‚metros obrigatÛrios da ContaCorrente corrente
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Empresa where codigo = " & ID_empresa, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    Select Case Familiatext
        Case "001": 'Banco do brasil
            Select Case Carteira
                Case "11 - Simples - Com Registro":
                    ArquivoLicenca = TBAbrir!Registro_boleto & "-001-11.conf"
                    OutrosDadosConfiguracao1 = Carteira1
                Case "17 - Direta Especial - Com Registro":
                    ArquivoLicenca = TBAbrir!Registro_boleto & "-001-17.conf"
                    OutrosDadosConfiguracao1 = Carteira1
                Case "17Simples - Direta Especial Simples - Com Registro":
                    ArquivoLicenca = TBAbrir!Registro_boleto & "-001-17SIMPLES.conf"
                    OutrosDadosConfiguracao1 = Carteira1
                Case "17-7 - Direta Especial - Com Registro ConvÍnio 7 dÌgitos":
                    ArquivoLicenca = TBAbrir!Registro_boleto & "-001-17-7.conf"
                    OutrosDadosConfiguracao1 = Carteira1
                    OutrosDadosConfiguracao2 = "0000000000"
                Case "18 - Simples - Sem Registro":
                    ArquivoLicenca = TBAbrir!Registro_boleto & "-001-18.conf"
                    OutrosDadosConfiguracao1 = Carteira1
                Case "18-7 - Simples - Sem Registro - ConvÍnio 7 dÌgitos":
                    ArquivoLicenca = TBAbrir!Registro_boleto & "-001-18-7.conf"
                    OutrosDadosConfiguracao1 = Carteira1
            End Select
            Select Case Len(Agencia)
                Case 1: AgenciaBol = "0000-" & Agencia
                Case 2: AgenciaBol = "000" & Left(Agencia, 1) & "-" & Right(Agencia, 1)
                Case 3: AgenciaBol = "00" & Left(Agencia, 2) & "-" & Right(Agencia, 1)
                Case 4: AgenciaBol = "0" & Left(Agencia, 3) & "-" & Right(Agencia, 1)
                Case Is >= 5: AgenciaBol = Left(Agencia, 4) & "-" & Right(Agencia, 1)
            End Select
            Select Case Len(ContaCorrente)
                Case 1: ContaCorrenteBol = "00000000-" & ContaCorrente
                Case 2: ContaCorrenteBol = "0000000" & Left(ContaCorrente, 1) & "-" & Right(ContaCorrente, 1)
                Case 3: ContaCorrenteBol = "000000" & Left(ContaCorrente, 2) & "-" & Right(ContaCorrente, 1)
                Case 4: ContaCorrenteBol = "00000" & Left(ContaCorrente, 3) & "-" & Right(ContaCorrente, 1)
                Case 5: ContaCorrenteBol = "0000" & Left(ContaCorrente, 4) & "-" & Right(ContaCorrente, 1)
                Case 6: ContaCorrenteBol = "000" & Left(ContaCorrente, 5) & "-" & Right(ContaCorrente, 1)
                Case 7: ContaCorrenteBol = "00" & Left(ContaCorrente, 6) & "-" & Right(ContaCorrente, 1)
                Case 8: ContaCorrenteBol = "0" & Left(ContaCorrente, 7) & "-" & Right(ContaCorrente, 1)
                Case Is >= 9: ContaCorrenteBol = Left(ContaCorrente, 8) & "-" & Right(ContaCorrente, 1)
            End Select
            If Carteira = "17-7 - Direta Especial - Com Registro ConvÍnio 7 dÌgitos" Or Carteira = "18-7 - Simples - Sem Registro - ConvÍnio 7 dÌgitos" Then
                Codigocedente = FunTamanhoTextoZeroEsq(Left(Codigocedente, 7), 7)
            Else
                Codigocedente = FunTamanhoTextoZeroEsq(Left(Codigocedente, 6), 6)
            End If
            
            Diretorio = Localrel & "\Boletos\Arquivos remessa\Banco do brasil"
            Arquivo = "CBR" & Dia & Mes & "." & SeqRemessa
            Layout = "FEBRABAN240"
        Case "033": 'Santander
            If Carteira = "CSR - CobranÁa Simples Sem Registro" Or Carteira = "ECR - CobranÁa Simples Com Registro" Or Carteira = "COBR-Nova - CobranÁa Simples - R·pida Com Registro" Then
                Select Case Carteira
                    Case "CSR - CobranÁa Simples Sem Registro": ArquivoLicenca = TBAbrir!Registro_boleto & "-033-CSR.conf"
                    Case "ECR - CobranÁa Simples Com Registro":
                        ArquivoLicenca = TBAbrir!Registro_boleto & "-033-ECR.conf"
                        'OutrosDadosConfiguracao1 = Left(Agencia, 4) & FunTamanhoTextoZeroEsq(Codigocedente, 7) & Right(ContaCorrente, 9) forma antiga com 11 digitos no nosso numero
                        OutrosDadosConfiguracao1 = Left(Agencia, 5) & FunTamanhoTextoZeroEsq(Codigocedente, 7) & Left(ContaCorrente, 9)
                    Case "COBR-Nova - CobranÁa Simples - R·pida Com Registro":
                        ArquivoLicenca = TBAbrir!Registro_boleto & "-033-COBR-NOVA.conf"
                        OutrosDadosConfiguracao1 = Left(Agencia, 5) & FunTamanhoTextoZeroEsq(Codigocedente, 7) & Left(ContaCorrente, 9)
                End Select
                Select Case Len(Agencia)
                    Case 1: AgenciaBol = "0000-" & Agencia
                    Case 2: AgenciaBol = "000" & Left(Agencia, 1) & "-" & Right(Agencia, 1)
                    Case 3: AgenciaBol = "00" & Left(Agencia, 2) & "-" & Right(Agencia, 1)
                    Case 4: AgenciaBol = "0" & Left(Agencia, 3) & "-" & Right(Agencia, 1)
                    Case Is >= 5: AgenciaBol = Left(Agencia, 4) & "-" & Right(Agencia, 1)
                End Select
                Select Case Len(ContaCorrente)
                    Case 1: ContaCorrenteBol = "000000000-" & ContaCorrente
                    Case 2: ContaCorrenteBol = "00000000" & Left(ContaCorrente, 1) & "-" & Right(ContaCorrente, 1)
                    Case 3: ContaCorrenteBol = "0000000" & Left(ContaCorrente, 2) & "-" & Right(ContaCorrente, 1)
                    Case 4: ContaCorrenteBol = "000000" & Left(ContaCorrente, 3) & "-" & Right(ContaCorrente, 1)
                    Case 5: ContaCorrenteBol = "00000" & Left(ContaCorrente, 4) & "-" & Right(ContaCorrente, 1)
                    Case 6: ContaCorrenteBol = "0000" & Left(ContaCorrente, 5) & "-" & Right(ContaCorrente, 1)
                    Case 7: ContaCorrenteBol = "000" & Left(ContaCorrente, 6) & "-" & Right(ContaCorrente, 1)
                    Case 8: ContaCorrenteBol = "00" & Left(ContaCorrente, 7) & "-" & Right(ContaCorrente, 1)
                    Case 9: ContaCorrenteBol = "0" & Left(ContaCorrente, 8) & "-" & Right(ContaCorrente, 1)
                    Case Is >= 10: ContaCorrenteBol = Left(ContaCorrente, 9) & "-" & Right(ContaCorrente, 1)
                End Select
                Select Case Len(Codigocedente)
                    Case 1: Codigocedente = "000000-" & Codigocedente
                    Case 2: Codigocedente = "00000" & Left(Codigocedente, 1) & "-" & Right(Codigocedente, 1)
                    Case 3: Codigocedente = "0000" & Left(Codigocedente, 2) & "-" & Right(Codigocedente, 1)
                    Case 4: Codigocedente = "000" & Left(Codigocedente, 3) & "-" & Right(Codigocedente, 1)
                    Case 5: Codigocedente = "00" & Left(Codigocedente, 4) & "-" & Right(Codigocedente, 1)
                    Case 6: Codigocedente = "0" & Left(Codigocedente, 5) & "-" & Right(Codigocedente, 1)
                    Case Is >= 7: Codigocedente = Left(Codigocedente, 6) & "-" & Right(Codigocedente, 1)
                End Select
            Else
                Select Case Carteira
                    Case "COB - CobranÁa Simples": ArquivoLicenca = TBAbrir!Registro_boleto & "-033-COB.conf"
                    Case "COBR - CobranÁa Simples - R·pida Com Registro": ArquivoLicenca = TBAbrir!Registro_boleto & "-033-COBR.conf"
                End Select
                AgenciaBol = Mid(Agencia, 2, 3)
                ContaCorrente = Codigocedente
                Select Case Len(ContaCorrente)
                    Case 1: ContaCorrenteBol = "00" & " " & "00000" & " " & ContaCorrente
                    Case 2: ContaCorrenteBol = "00" & " " & "0000" & Left(ContaCorrente, 1) & " " & Mid(ContaCorrente, 2, 1)
                    Case 3: ContaCorrenteBol = "00" & " " & "000" & Left(ContaCorrente, 2) & " " & Mid(ContaCorrente, 3, 1)
                    Case 4: ContaCorrenteBol = "00" & " " & "00" & Left(ContaCorrente, 3) & " " & Mid(ContaCorrente, 4, 1)
                    Case 5: ContaCorrenteBol = "00" & " " & "0" & Left(ContaCorrente, 4) & " " & Mid(ContaCorrente, 5, 1)
                    Case 6: ContaCorrenteBol = "00" & " " & Left(ContaCorrente, 5) & " " & Mid(ContaCorrente, 6, 1)
                    Case 7: ContaCorrenteBol = "0" & Left(ContaCorrente, 1) & " " & Mid(ContaCorrente, 2, 5) & " " & Mid(ContaCorrente, 7, 1)
                    Case Is >= 8: ContaCorrenteBol = Left(ContaCorrente, 2) & " " & Mid(ContaCorrente, 3, 5) & " " & Mid(ContaCorrente, 8, 1)
                End Select
                Codigocedente = FunTamanhoTextoVazioDir(Left(NomeAgencia, 20), 20)
            End If
            Diretorio = Localrel & "\Boletos\Arquivos remessa\Santander"
            Arquivo = "DB" & Dia & Mes & Right(Ano, 2) & "." & SeqRemessa
            Layout = "CNAB400"
        Case "104": 'Caixa
            Arquivo = "CB" & Dia & Mes & "." & SeqRemessa
            If Carteira = "SIG14 - SIG Com Registro - Emiss„o pelo Cedente" Or Carteira = "SR5 - SINCO - Sem Registro" Then
                If Carteira = "SIG14 - SIG Com Registro - Emiss„o pelo Cedente" Then
                    ArquivoLicenca = TBAbrir!Registro_boleto & "-104-SIG14.conf"
                    Layout = "SIGCB240"
                    Arquivo = Replace(Arquivo, ".txt", ".rem")
                Else
                    ArquivoLicenca = TBAbrir!Registro_boleto & "-104-SR5.conf"
                    Layout = "FEBRABAN240"
                    OutrosDadosConfiguracao1 = "S"
                    OutrosDadosConfiguracao2 = "S"
                End If
                AgenciaBol = FunTamanhoTextoZeroEsq(Left(Agencia, 4), 4)
                ContaCorrenteBol = ""
                Codigocedente = FunTamanhoTextoZeroEsq(Left(ReturnNumbersOnly(Codigocedente), 6), 6)
            Else
                Select Case Carteira
                    Case "CR - CobranÁa R·pida": ArquivoLicenca = TBAbrir!Registro_boleto & "-104-CR.conf"
                    Case "CS - CobranÁa Simples": ArquivoLicenca = TBAbrir!Registro_boleto & "-104-CS.conf"
                    Case "SR - CobranÁa Sem Registro": ArquivoLicenca = TBAbrir!Registro_boleto & "-104-SR.conf"
                End Select
                AgenciaBol = ""
                ContaCorrenteBol = ""
                Codigocedente = ReturnNumbersOnly(Codigocedente)
                Codigocedente = Left(Codigocedente, 4) & "." & Mid(Codigocedente, 5, 3) & "." & Mid(Codigocedente, 8, 8) & "-" & Right(Codigocedente, 1)
                Layout = "CNAB400"
            End If
            Diretorio = Localrel & "\Boletos\Arquivos remessa\Caixa"
            
        Case "237": 'Bradesco
            Select Case Carteira
                Case "06 - Sem Registro": ArquivoLicenca = TBAbrir!Registro_boleto & "-237-06.conf"
                Case "09 - Com Registro": ArquivoLicenca = TBAbrir!Registro_boleto & "-237-09.conf"
            End Select
            Select Case Len(Agencia)
                Case 1: AgenciaBol = "0000-" & Agencia
                Case 2: AgenciaBol = "000" & Left(Agencia, 1) & "-" & Right(Agencia, 1)
                Case 3: AgenciaBol = "00" & Left(Agencia, 2) & "-" & Right(Agencia, 1)
                Case 4: AgenciaBol = "0" & Left(Agencia, 3) & "-" & Right(Agencia, 1)
                Case Is >= 5: AgenciaBol = Left(Agencia, 4) & "-" & Right(Agencia, 1)
            End Select
            Select Case Len(ContaCorrente)
                Case 1: ContaCorrenteBol = "0000000-" & ContaCorrente
                Case 2: ContaCorrenteBol = "000000" & Left(ContaCorrente, 1) & "-" & Right(ContaCorrente, 1)
                Case 3: ContaCorrenteBol = "00000" & Left(ContaCorrente, 2) & "-" & Right(ContaCorrente, 1)
                Case 4: ContaCorrenteBol = "0000" & Left(ContaCorrente, 3) & "-" & Right(ContaCorrente, 1)
                Case 5: ContaCorrenteBol = "000" & Left(ContaCorrente, 4) & "-" & Right(ContaCorrente, 1)
                Case 6: ContaCorrenteBol = "00" & Left(ContaCorrente, 5) & "-" & Right(ContaCorrente, 1)
                Case 7: ContaCorrenteBol = "0" & Left(ContaCorrente, 6) & "-" & Right(ContaCorrente, 1)
                Case Is >= 8: ContaCorrenteBol = Left(ContaCorrente, 7) & "-" & Right(ContaCorrente, 1)
            End Select
            Codigocedente = FunTamanhoTextoZeroEsq(Left(Codigocedente, 15), 15)
            Diretorio = Localrel & "\Boletos\Arquivos remessa\Bradesco"
            Arquivo = "CB" & Dia & Mes & "." & SeqRemessa
            Layout = "CNAB400"
        Case "341": 'Ita˙
            Select Case Carteira
                Case "109 - Direta EletrÙnica Sem Emiss„o - Simples": ArquivoLicenca = TBAbrir!Registro_boleto & "-341-109.conf"
                Case "112 - Escritual EletrÙnica - simples / contratual": ArquivoLicenca = TBAbrir!Registro_boleto & "-341-112.conf"
                Case "175 - Sem Registro Sem Emiss„o": ArquivoLicenca = TBAbrir!Registro_boleto & "-341-175.conf"
            End Select
            AgenciaBol = FunTamanhoTextoZeroEsq(Left(Agencia, 4), 4)
            Select Case Len(ContaCorrente)
                Case 1: ContaCorrenteBol = "00000-" & ContaCorrente
                Case 2: ContaCorrenteBol = "0000" & Left(ContaCorrente, 1) & "-" & Right(ContaCorrente, 1)
                Case 3: ContaCorrenteBol = "000" & Left(ContaCorrente, 2) & "-" & Right(ContaCorrente, 1)
                Case 4: ContaCorrenteBol = "00" & Left(ContaCorrente, 3) & "-" & Right(ContaCorrente, 1)
                Case 5: ContaCorrenteBol = "0" & Left(ContaCorrente, 4) & "-" & Right(ContaCorrente, 1)
                Case Is >= 6: ContaCorrenteBol = Left(ContaCorrente, 5) & "-" & Right(ContaCorrente, 1)
            End Select
            Codigocedente = ContaCorrente
            Diretorio = Localrel & "\Boletos\Arquivos remessa\Ita˙"
            
            'Debug.print Diretorio
            'O nome do arquivo remessa do Ita˙ sÛ aceita no m·ximo 8 caracteres
            SeqRemessaTexto = Left(SeqRemessa, Len(SeqRemessa) - 4)
            Select Case Len(SeqRemessaTexto)
                Case 1: RemessaTexto = "0" & Right(SeqRemessaTexto, 1)
                Case 2: RemessaTexto = SeqRemessaTexto
                Case Is >= 3: RemessaTexto = Right(SeqRemessaTexto, 2)
            End Select
            Arquivo = Dia & Mes & Right(Ano, 2) & RemessaTexto & ".txt"
            Layout = "CNAB400"
        Case "356": 'ABN e Real
            Select Case Carteira
                Case "20 - CobranÁa Simples": ArquivoLicenca = TBAbrir!Registro_boleto & "-356-20.conf"
            End Select
            AgenciaBol = FunTamanhoTextoZeroEsq(Left(Agencia, 4), 4)
            Select Case Len(ContaCorrente)
                Case 1: ContaCorrenteBol = "000000-" & ContaCorrente
                Case 2: ContaCorrenteBol = "00000" & Left(ContaCorrente, 1) & "-" & Right(ContaCorrente, 1)
                Case 3: ContaCorrenteBol = "0000" & Left(ContaCorrente, 2) & "-" & Right(ContaCorrente, 1)
                Case 4: ContaCorrenteBol = "000" & Left(ContaCorrente, 3) & "-" & Right(ContaCorrente, 1)
                Case 5: ContaCorrenteBol = "00" & Left(ContaCorrente, 4) & "-" & Right(ContaCorrente, 1)
                Case 6: ContaCorrenteBol = "0" & Left(ContaCorrente, 5) & "-" & Right(ContaCorrente, 1)
                Case Is > 7: ContaCorrenteBol = Left(ContaCorrente, 6) & "-" & Right(ContaCorrente, 1)
            End Select
            Codigocedente = FunTamanhoTextoZeroEsq(Left(Codigocedente, 9), 9)
            Diretorio = Localrel & "\Boletos\Arquivos remessa\ABN e Real"
            Arquivo = "CB" & Dia & Mes & "." & SeqRemessa
            Layout = "CNAB400"
        Case "399": 'HSBC
            Select Case Carteira
                Case "CNR - Sem Registro": ArquivoLicenca = TBAbrir!Registro_boleto & "-399-CNR.conf"
            End Select
            Codigocedente = FunTamanhoTextoZeroEsq(Left(Codigocedente, 7), 7)
            Diretorio = Localrel & "\Boletos\Arquivos remessa\HSBC"
            Arquivo = "D" & Dia & Mes & Ano & "." & SeqRemessa
            Layout = "CNAB400"
        Case "409": 'Unibanco
            Select Case Carteira
                Case "Especial": ArquivoLicenca = TBAbrir!Registro_boleto & "-409-ESPECIAL.conf"
            End Select
            AgenciaBol = FunTamanhoTextoZeroEsq(Left(Agencia, 4), 4)
            Select Case Len(ContaCorrente)
                Case 1: ContaCorrenteBol = "000" & "." & "000" & "-" & ContaCorrente
                Case 2: ContaCorrenteBol = "000" & "." & "00" & Left(ContaCorrente, 1) & "-" & Mid(ContaCorrente, 2, 1)
                Case 3: ContaCorrenteBol = "000" & "." & "0" & Left(ContaCorrente, 2) & "-" & Mid(ContaCorrente, 3, 1)
                Case 4: ContaCorrenteBol = "000" & "." & Left(ContaCorrente, 3) & "-" & Mid(ContaCorrente, 4, 1)
                Case 5: ContaCorrenteBol = "00" & Left(ContaCorrente, 1) & "." & Mid(ContaCorrente, 2, 3) & "-" & Mid(ContaCorrente, 5, 1)
                Case 6: ContaCorrenteBol = "0" & Left(ContaCorrente, 2) & "." & Mid(ContaCorrente, 3, 3) & "-" & Mid(ContaCorrente, 6, 1)
                Case Is >= 7: ContaCorrenteBol = Left(ContaCorrente, 3) & "." & Mid(ContaCorrente, 4, 3) & "-" & Mid(ContaCorrente, 7, 1)
            End Select
            Codigocedente = ContaCorrente
            Diretorio = Localrel & "\Boletos\Arquivos remessa\Unibanco"
            Arquivo = "CBR" & Dia & Mes & "." & SeqRemessa
            Layout = "CNAB240"
    End Select
End If

Set GerArqPastas = CreateObject("Scripting.FileSystemObject")
If GerArqPastas.FileExists(Localrel & "\Boletos\Carteiras\" & ArquivoLicenca) = False Then
    If Remessa = False Then TextoMsg = "emitir o boleto" Else TextoMsg = "gerar o arquivo remessa"
    USMsgBox ("N„o ser· possÌvel " & TextoMsg & ", pois n„o foi encontrado o arquivo licenÁa " & ArquivoLicenca & " na pasta " & Localrel & "\Boletos\Carteiras."), vbExclamation, "CAPRIND v5.0"
    Permitido1 = False
    Exit Sub
End If
    
CobreBemX1.ArquivoLicenca = Localrel & "\Boletos\Carteiras\" & ArquivoLicenca
CobreBemX1.CodigoAgencia = AgenciaBol
CobreBemX1.NumeroContaCorrente = ContaCorrenteBol
CobreBemX1.Codigocedente = Codigocedente
CobreBemX1.OutroDadoConfiguracao1 = OutrosDadosConfiguracao1
CobreBemX1.OutroDadoConfiguracao2 = OutrosDadosConfiguracao2
If Remessa = True Then CobreBemX1.ArquivoRemessa.Sequencia = Left(SeqRemessa, Len(SeqRemessa) - 4)

'CobreBemX1.InicioNossoNumero =
'CobreBemX1.FimNossoNumero =
'CobreBemX1.ProximoNossoNumero =

If Enviar_Email = True Then
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select EE.*, E.Empresa from Empresa E INNER JOIN Empresa_email EE ON EE.ID_empresa = E.Codigo where EE.ID_empresa = " & ID_empresa & " and EE.Aplicacao = 'F'", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        'InÌcio da configuraÁ„o dos dados do Cedente para envio de boletos por email
        CobreBemX1.PadroesBoleto.PadroesBoletoEmail.SMTP.Servidor = TBAbrir!Servidor_SMTP ' Trocar para apontar para o seu servidor SMTP
        CobreBemX1.PadroesBoleto.PadroesBoletoEmail.SMTP.Porta = TBAbrir!Porta
        CobreBemX1.PadroesBoleto.PadroesBoletoEmail.SMTP.Usuario = TBAbrir!Usuario 'utilizar esta propriedade para acesso a servidores SMTP seguros
        CobreBemX1.PadroesBoleto.PadroesBoletoEmail.SMTP.Senha = TBAbrir!Senha 'utilizar esta propriedade para acesso a servidores SMTP seguros
        CobreBemX1.PadroesBoleto.PadroesBoletoEmail.URLImagensCodigoBarras = "http://www.bptob.com/imagenscbe/"
        CobreBemX1.PadroesBoleto.PadroesBoletoEmail.URLLogotipo = "http://www.thisf.com.br/banners/BannerCBE.gif"
        CobreBemX1.PadroesBoleto.PadroesBoletoEmail.PadroesEmail.Assunto = Assunto_email
        CobreBemX1.PadroesBoleto.PadroesBoletoEmail.PadroesEmail.EmailFrom.Endereco = TBAbrir!Email
        CobreBemX1.PadroesBoleto.PadroesBoletoEmail.PadroesEmail.EmailFrom.Nome = TBAbrir!Nome
        'CobreBemX1.PadroesBoleto.PadroesBoletoEmail.PadroesEmail.FormaEnvio = feeSMTPBoletoHTML
        'CobreBemX1.PadroesBoleto.PadroesBoletoEmail.PadroesEmail.Mensagem = "Mensagem quando for utilizado HTLM ANEXADO"
        CobreBemX1.PadroesBoleto.PadroesBoletoEmail.PadroesEmail.FormaEnvio = feeSMTPMensagemBoletoHTMLAnexo 'HTML ANEXADO
        'CobreBemX1.PadroesBoleto.PadroesBoletoEmail.PadroesEmail.FormaEnvio = feeSMTPMensagemBoletoPDFAnexo 'PDF
        If Left(TBAbrir!Empresa, 7) = "CAPRIND" Then
            CobreBemX1.PadroesBoleto.PadroesBoletoEmail.PadroesEmail.CopiaReply = True
            CobreBemX1.PadroesBoleto.PadroesBoletoEmail.PadroesEmail.EmailReply.Endereco = "caprind@caprind.com.br"
            CobreBemX1.PadroesBoleto.PadroesBoletoEmail.PadroesEmail.EmailReply.Nome = "CAPRIND SISTEMAS"
        End If
    End If
End If

'Logotipo do cedente na parte superior do boleto
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select Logotipo from Empresa where Codigo = " & ID_empresa & " and Logotipo <> 'Null'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    If TBAbrir!Logotipo <> "" Then CobreBemX1.PadroesBoleto.PadroesBoletoImpresso.ArquivoLogotipo = TBAbrir!Logotipo
End If
TBAbrir.Close
       
CobreBemX1.PadroesBoleto.PadroesBoletoImpresso.CaminhoImagensCodigoBarras = Localrel & "\Imagens\Bancos\"

'Utilize o par‚metro abaixo para efetuar ajustes na impress„o do boleto subindo ou descendo o mesmo na folha de papel
'Os valores devem ser informados em milÌmetros e quanto maior o valor mais para baixo ser· iniciado o boleto
'Se este par‚metro n„o for passado ser· assumido o valor 15 que È o indicado para a maioria das impressoras Jato de Tinta }
CobreBemX1.PadroesBoleto.PadroesBoletoImpresso.MargemSuperior = 3

'A prÛxima linha È utilizada para exibir um texto do lado direito do logotipo nos boletos impressos ou enviados por email
'CobreBemX1.PadroesBoleto.IdentificacaoCedente =

Exit Sub
tratar_erro:
    If Err.Number = "91" Then
        USMsgBox ("N„o foi encontrado o caminho " & TBAbrir!Logotipo & "."), vbExclamation, "CAPRIND v5.0"
        Permitido1 = False
        Exit Sub
    End If
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Function FunVerificaUsuariosConectados(Usuario As String) As Boolean
On Error GoTo tratar_erro

FunVerificaUsuariosConectados = True
Set TBFIltro = CreateObject("adodb.recordset")
TBFIltro.Open "Select * from Logon where Usuario <> '" & Usuario & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBFIltro.EOF = False Then
    FunVerificaUsuariosConectados = False
End If
TBFIltro.Close

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Sub ProcReindexarBDCaprindeGerprod()
On Error GoTo tratar_erro

With frmMenucaprind_menulateral
    Conexao.Execute "Truncate Table Logon"
    Conexao.Execute "DBCC CheckIdent('Logon',Reseed,1)"
    Conexao.Execute "Truncate Table Producao_Relatorios"
    Conexao.Execute "DBCC CheckIdent('Producao_Relatorios',Reseed,1)"
    Conexao.Execute "Truncate Table Producao_Relatorios_Total"
    Conexao.Execute "DBCC CheckIdent('Producao_Relatorios_Total',Reseed,1)"
    Conexao.Execute "Truncate Table Estoque_relatorios"
    Conexao.Execute "DBCC CheckIdent('Estoque_relatorios',Reseed,1)"
    Conexao.Execute "Truncate Table Troca_titulo_relatorio"
    Conexao.Execute "DBCC CheckIdent('Troca_titulo_relatorio',Reseed,1)"
    Conexao.Execute "Truncate Table Plano_de_contas_totalizacao"
    Conexao.Execute "DBCC CheckIdent('Plano_de_contas_totalizacao',Reseed,1)"
    Conexao.Execute "Truncate Table Compras_Recebimento_Relatorios"
    Conexao.Execute "DBCC CheckIdent('Compras_Recebimento_Relatorios',Reseed,1)"
    Conexao.Execute "ReindexarBD"
    USMsgBox ("ReindexaÁ„o do BD efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
End With

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcReindexarBDGNFe()
On Error GoTo tratar_erro

With frmMenucaprind_menulateral
    'Conexao_NFe.Execute "ReindexarBD"
    USMsgBox ("ReindexaÁ„o do BD efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
End With

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Function FunFormataNumeroArqRemessa(DataEmissao As Date, ID_instituicao As Long, Seq_remessa As Long) As String
On Error GoTo tratar_erro
Dim seql As String

Dia = Format(DataEmissao, "dd")
Mes = Format(DataEmissao, "mm")
Ano = Format(DataEmissao, "yyyy")
If Seq_remessa < 10 Then Seq1 = "0" & Seq_remessa Else Seq1 = Seq_remessa

Set TBFI = CreateObject("adodb.recordset")
TBFI.Open "Select * from tbl_Instituicoes where ID = " & ID_instituicao, Conexao, adOpenKeyset, adLockOptimistic
If TBFI.EOF = False Then
    If IsNull(TBFI!int_NBanco) = False And TBFI!int_NBanco <> "" Then
        Select Case TBFI!int_NBanco
            Case "001": FunFormataNumeroArqRemessa = "CBR" & Dia & Mes  'Banco do brasil
            Case "033": FunFormataNumeroArqRemessa = "DB" & Dia & Mes & Right(Ano, 2) 'Santander
            Case "104": FunFormataNumeroArqRemessa = "CB" & Dia & Mes 'Caixa
            Case "237": FunFormataNumeroArqRemessa = "CB" & Dia & Mes 'Bradesco
            Case "341": FunFormataNumeroArqRemessa = DS.FormatWithZeros(Seq1, 10 - Len(Seq1)) 'Dia & Mes & Right(Ano, 2) 'Ita˙
            Case "356": FunFormataNumeroArqRemessa = "CB" & Dia & Mes & "." 'ABN e Real
            Case "399": FunFormataNumeroArqRemessa = "D" & Dia & Mes & Ano & "." 'HSBC
            Case "409": FunFormataNumeroArqRemessa = "CBR" & Dia & Mes & "." 'Unibanco
        End Select
    End If
    If TBFI!int_NBanco <> "341" Then
    FunFormataNumeroArqRemessa = FunFormataNumeroArqRemessa & "." & Seq1
    End If
End If
TBFI.Close

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Sub ProcCopiarPlanoInspecao(Codinterno As String, Descricao As String, IDFase As Long, Fase As String, Grupo_op As String)
On Error GoTo tratar_erro

'Grava dados principais
Set TBplano = CreateObject("adodb.recordset")
TBplano.Open "select * from Plano", Conexao, adOpenKeyset, adLockOptimistic
TBplano.AddNew
TBplano!Rev = 0
TBplano!Desenho = Codinterno
If IDFase <> 0 Then
    TBplano!IDFase = IDFase
    Conexao.Execute "UPDATE Fases Set Plano_inspecao = 'True' where IDFase = " & IDFase
Else
    Conexao.Execute "UPDATE projproduto Set Plano_inspecao = 'True' where Desenho = '" & Codinterno & "'"
End If
If Fase <> "" Then TBplano!Fase = Fase
If Grupo_op <> "" Then TBplano!Grupo_op = Grupo_op
TBplano!Descricao = Descricao
TBplano!Inspetor = pubUsuario
TBplano!Data = Date
TBplano!Nivel = TBAbrir!Nivel
TBplano.Update

'Grava medidas
Do While TBAbrir.EOF = False
    Set TBGravar = CreateObject("adodb.recordset")
    TBGravar.Open "Select * from Planodimensao", Conexao, adOpenKeyset, adLockOptimistic
    TBGravar.AddNew
    TBGravar!IDPlano = TBplano!IDPlano
    TBGravar!Tipo = TBAbrir!Tipo
    TBGravar!dimdesejada = TBAbrir!dimdesejada
    TBGravar!TolSup = TBAbrir!TolSup
    TBGravar!TolInf = TBAbrir!TolInf
    TBGravar!Dim_superior = TBAbrir!Dim_superior
    TBGravar!Dim_inferior = TBAbrir!Dim_inferior
    TBGravar!Instrumento = TBAbrir!Instrumento
    TBGravar!DescInst = TBAbrir!DescInst
    TBGravar!Freq = TBAbrir!Freq
    TBGravar!Cartacontrole = TBAbrir!Cartacontrole
    TBGravar!indice = TBAbrir!indice
    TBGravar.Update
    'Grava famÌlias dos instrumentos
    Set TBFIltro = CreateObject("adodb.recordset")
    TBFIltro.Open "Select * from Planodimensao_instrumentos where ID_dimensao = " & TBAbrir!idDimensao, Conexao, adOpenKeyset, adLockOptimistic
    If TBFIltro.EOF = False Then
        Do While TBFIltro.EOF = False
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select * from Planodimensao_instrumentos", Conexao, adOpenKeyset, adLockOptimistic
            TBFI.AddNew
            TBFI!id_dimensao = TBGravar!idDimensao
            TBFI!Familia = TBFIltro!Familia
            TBFI.Update
            TBFIltro.MoveNext
        Loop
    End If
    TBFIltro.Close
    TBAbrir.MoveNext
Loop
TBGravar.Close

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Function FunVerifCalcIPISDesc(ID_empresa As Integer) As Boolean
On Error GoTo tratar_erro

FunVerifCalcIPISDesc = False
Set TBAliquota = CreateObject("adodb.recordset")
TBAliquota.Open "Select Codigo from Empresa where Codigo = " & ID_empresa & " and Calcular_IPI_sem_desc = 'True'", Conexao, adOpenKeyset, adLockOptimistic
If TBAliquota.EOF = False Then
    FunVerifCalcIPISDesc = True
End If
TBAliquota.Close

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Sub ProcVerifQtdeFaturadaProdServ(ID_carteira As String, Codinterno As String, VerificaRE As Boolean)
On Error GoTo tratar_erro

Set TBVendas = CreateObject("adodb.recordset")
TBVendas.Open "Select * from vendas_carteira where Codigo = " & ID_carteira, Conexao, adOpenKeyset, adLockOptimistic
If TBVendas.EOF = False Then
    TextoNF = ""
    TextoUNCom = " - " & TBVendas!Unidade_com
    Set TBItem = CreateObject("adodb.recordset")
    TBItem.Open "Select NF.*, NFPP.Quantidade, NFP.Unidade_com from (tbl_dados_nota_fiscal NF INNER JOIN tbl_Detalhes_Nota_pedidos NFPP ON NFPP.ID_nota = NF.ID) INNER JOIN tbl_Detalhes_Nota NFP ON NFP.Int_codigo = NFPP.ID_prod_NF where NF.int_status = 1 and NF.Pedido_interno = 'True' and NFPP.ID_carteira = " & ID_carteira & " and NFPP.Codinterno = '" & Codinterno & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBItem.EOF = False Then
        Do While TBItem.EOF = False
            Select Case TBItem!TipoNF
                Case "M1": Tipo = "Prod."
                Case "SA": Tipo = "Serv."
                Case "M1SA": Tipo = "Prod./Ser."
            End Select
            If TextoNF = "" Then
                TextoNF = "NF: " & TBItem!int_NotaFiscal & " - " & Tipo & " - Emiss„o: " & Format(TBItem!dt_DataEmissao, "dd/mm/yy") & " - Qtde.: " & Format(TBItem!quantidade, "###,##0.0000") & TextoUNCom
            Else
                TextoNF = TextoNF & vbCrLf & "NF: " & TBItem!int_NotaFiscal & " - " & Tipo & " - Emiss„o: " & Format(TBItem!dt_DataEmissao, "dd/mm/yy") & " - Qtde.: " & Format(TBItem!quantidade, "###,##0.0000") & TextoUNCom
            End If
            TBItem.MoveNext
        Loop
    End If
    
    Permitido = True
    TextoRE = ""
    If VerificaRE = True Then
        Set TBItem = CreateObject("adodb.recordset")
        TBItem.Open "Select ID_estoque, Qtde_empenhada - Qtde_saida as qtdeliberada from Estoque_Controle_Empenho_Vendas where ID_carteira = " & ID_carteira & " and Qtde_empenhada - Qtde_saida > 0", Conexao, adOpenKeyset, adLockOptimistic
        If TBItem.EOF = False Then
            Do While TBItem.EOF = False
                If TextoRE = "" Then
                    TextoRE = "RE: " & TBItem!ID_estoque & " - Qtde.: " & Format(TBItem!qtdeliberada, "###,##0.0000")
                Else
                    TextoRE = TextoRE & " | RE: " & TBItem!ID_estoque & " - Qtde.: " & Format(TBItem!qtdeliberada, "###,##0.0000")
                End If
                TBItem.MoveNext
            Loop
            Permitido = False
            USMsgBox ("CÛd. interno: " & TBVendas!Desenho & " - Rev.: " & TBVendas!Rev_codinterno & " " & vbCrLf & "Status: " & TBVendas!Liberacao & " " & vbCrLf & "Qtde. liberada p/ faturar: " & Format(TBVendas!qtdeliberada, "###,##0.0000") & " " & vbCrLf & "Qtde. faturada: " & Format(TBVendas!QtdeFaturada, "###,##0.0000") & " " & vbCrLf & " " & TextoNF & " " & vbCrLf & " IdentificaÁ„o do estoque: " & vbCrLf & " " & TextoRE), vbInformation, "CAPRIND v5.0"
        End If
    End If
    TBItem.Close
    
    If Permitido = True Then USMsgBox ("CÛd. interno: " & TBVendas!Desenho & " - Rev.: " & TBVendas!Rev_codinterno & " " & vbCrLf & "Status: " & TBVendas!Liberacao & " " & vbCrLf & "Qtde. vendida: " & Format(TBVendas!quantidade, "###,##0.0000") & TextoUNCom & vbCrLf & "Qtde. liberada p/ faturar: " & Format(TBVendas!qtdeliberada, "###,##0.0000") & TextoUNCom & vbCrLf & "Qtde. faturada: " & Format(TBVendas!QtdeFaturada, "###,##0.0000") & TextoUNCom & vbCrLf & "Qtde. a faturar: " & Format(TBVendas!qtdeliberada - TBVendas!QtdeFaturada, "###,##0.0000") & TextoUNCom & vbCrLf & "Saldo pedido: " & Format(TBVendas!quantidade - TBVendas!QtdeFaturada, "###,##0.0000") & TextoUNCom & vbCrLf & TextoNF), vbInformation, "CAPRIND v5.0"
End If
TBVendas.Close

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Function FunVerifNFProdServSemCad(ID_empresa As Integer) As Boolean
On Error GoTo tratar_erro

FunVerifNFProdServSemCad = False
Set TBAliquota = CreateObject("adodb.recordset")
TBAliquota.Open "Select * from Empresa where Codigo = " & ID_empresa & " and Bloquear_NF_prod_serv_sem_cadastro = 'True'", Conexao, adOpenKeyset, adLockOptimistic
If TBAliquota.EOF = False Then
    FunVerifNFProdServSemCad = True
End If
TBAliquota.Close

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Sub ProcConfVariaveisLocCliente(Clientes1 As Boolean, Compras_Cotacao1 As Boolean, Compras_Fornecedores1 As Boolean, Compras_Pedido1 As Boolean, Compras_Programacao1 As Boolean, Engenharia_Localfornecedor2 As Boolean, Engenharia_Localfornecedor3 As Boolean, Engenharia_Localcliente2 As Boolean, Engenharia_Localcliente3 As Boolean, Engenharia_Normas1 As Boolean, Estoque_Consignacao1 As Boolean, Faturamento1 As Boolean, Financeiro_Contas_Pagar1 As Boolean, Financeiro_Contas_Pagas1 As Boolean, Financeiro_Contas_Receber1 As Boolean, Financeiro_Contas_Recebidas1 As Boolean, Qualidade_PPAP_PSW1 As Boolean, Qualidade_PPAP_FMEA1 As Boolean, RH_Funcionarios1 As Boolean, OpcoesGerais1 As Boolean, PCP_Ordem1 As Boolean, RNC1 As Boolean, Telemarketing1 As Boolean, Vendas_Analise1 As Boolean, Vendas_PI1 As Boolean, Vendas_Proposta1 As Boolean, Vendas_Programacao1 As Boolean, Vendas_Vendedores1 As Boolean, Fiscal_NaturezaOperacao1 As Boolean, Estoque_Inventario1 As Boolean)
On Error GoTo tratar_erro

Clientes = Clientes1
Compras_Cotacao = Compras_Cotacao1
Compras_Fornecedores = Compras_Fornecedores1
Compras_Pedido = Compras_Pedido1
Compras_Programacao = Compras_Programacao1
Engenharia_Localfornecedor = Engenharia_Localfornecedor2
Engenharia_Localfornecedor1 = Engenharia_Localfornecedor3
Engenharia_Localcliente = Engenharia_Localcliente2
Engenharia_Localcliente1 = Engenharia_Localcliente3
Engenharia_Normas = Engenharia_Normas1
Estoque_Consignacao = Estoque_Consignacao1
Estoque_Inventario = Estoque_Inventario1
Faturamento = Faturamento1
Financeiro_Contas_Pagar = Financeiro_Contas_Pagar1
Financeiro_Contas_Pagas = Financeiro_Contas_Pagas1
Financeiro_Contas_Receber = Financeiro_Contas_Receber1
Financeiro_Contas_Recebidas = Financeiro_Contas_Recebidas1
Qualidade_PPAP_PSW = Qualidade_PPAP_PSW1
Qualidade_PPAP_FMEA = Qualidade_PPAP_FMEA1
RH_Funcionarios = RH_Funcionarios1
OpcoesGerais = OpcoesGerais1
PCP_Ordem = PCP_Ordem1
RNC = RNC1
Telemarketing = Telemarketing1
Vendas_Analise = Vendas_Analise1
Vendas_PI = Vendas_PI1
Vendas_Proposta = Vendas_Proposta1
Vendas_Programacao = Vendas_Programacao1
Vendas_Vendedores = Vendas_Vendedores1
Fiscal_NaturezaOperacao = Fiscal_NaturezaOperacao1

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcConfVariaveisLocForn(Clientes1 As Boolean, Compras_Cotacao1 As Boolean, Compras_Fornecedores1 As Boolean, Compras_Pedido1 As Boolean, Compras_Programacao1 As Boolean, Engenharia_Localfornecedor2 As Boolean, Engenharia_Localfornecedor3 As Boolean, Estoque_Consignacao1 As Boolean, Faturamento1 As Boolean, Financeiro_Contas_Pagar1 As Boolean, Financeiro_Contas_Pagas1 As Boolean, Financeiro_Contas_Receber1 As Boolean, Financeiro_Contas_Recebidas1 As Boolean, OpcoesGerais1 As Boolean, Qualidade_PPAP_FMEA1 As Boolean, RH_Funcionarios1 As Boolean, RNC1 As Boolean, Vendas_Proposta1 As Boolean, Vendas_PI1 As Boolean, Estoque_Inventario1 As Boolean)
On Error GoTo tratar_erro

Clientes = Clientes1
Compras_Cotacao = Compras_Cotacao1
Compras_Fornecedores = Compras_Fornecedores1
Compras_Pedido = Compras_Pedido1
Compras_Programacao = Compras_Programacao1
Engenharia_Localfornecedor = Engenharia_Localfornecedor2
Engenharia_Localfornecedor1 = Engenharia_Localfornecedor3
Estoque_Consignacao = Estoque_Consignacao1
Estoque_Inventario = Estoque_Inventario1
Faturamento = Faturamento1
Financeiro_Contas_Pagar = Financeiro_Contas_Pagar1
Financeiro_Contas_Pagas = Financeiro_Contas_Pagas1
Financeiro_Contas_Receber = Financeiro_Contas_Receber1
Financeiro_Contas_Recebidas = Financeiro_Contas_Recebidas1
OpcoesGerais = OpcoesGerais1
Qualidade_PPAP_FMEA = Qualidade_PPAP_FMEA1
RH_Funcionarios = RH_Funcionarios1
RNC = RNC1
Vendas_Proposta = Vendas_Proposta1
Vendas_PI = Vendas_PI1

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcExcluirDadosProducaoRelatorios()
On Error GoTo tratar_erro

Conexao.Execute "DELETE PD from Producao_Relatorios P INNER JOIN Producao_Relatorios_Detalhes PD ON PD.IDProd_Rel = P.ID where P.Modulo = '" & Formulario & "' and P.Responsavel = '" & pubUsuario & "'"
Conexao.Execute "DELETE from Producao_Relatorios where Modulo = '" & Formulario & "' and Responsavel = '" & pubUsuario & "'"

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcExcluirDadosProducaoRelatoriosTotal()
On Error GoTo tratar_erro

Conexao.Execute "DELETE from Producao_Relatorios_Total where Modulo = '" & Formulario & "' and Responsavel = '" & pubUsuario & "'"

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcExcluirRelacionamentoNF(ID_nota As Long)
On Error GoTo tratar_erro
Dim id_produto_selecionada As Long
Dim id_produto_relacionada As Long
Dim TBRelacionamento     As ADODB.Recordset
Dim TBRelacionamento1     As ADODB.Recordset

'Excluir relacionamento de nf e corrige saldo
Set TBRelacionamento = CreateObject("adodb.recordset")
TBRelacionamento.Open "Select ID_nota, ID_nota_relacionada, Id_produto, ID_produto_relacionada, Qtde from Faturamento_Relacionamento where Id_nota = " & ID_nota & " or ID_nota_relacionada = " & ID_nota, Conexao, adOpenKeyset, adLockOptimistic
If TBRelacionamento.EOF = False Then
    Do While TBRelacionamento.EOF = False
    
        If TBRelacionamento!ID_nota = ID_nota Then
            id_produto_selecionada = TBRelacionamento!ID_Produto
            id_produto_relacionada = TBRelacionamento!id_produto_relacionada
            procExcluirDevolucaoNF ID_nota, TBRelacionamento!ID_nota_relacionada
        Else
            id_produto_selecionada = TBRelacionamento!id_produto_relacionada
            id_produto_relacionada = TBRelacionamento!ID_Produto
            procExcluirDevolucaoNF ID_nota, TBRelacionamento!ID_nota
        End If

        'Atualiza o saldo no produto da NF de entrada
        Conexao.Execute "UPDATE tbl_Detalhes_Nota SET Saldo = saldo + " & TBRelacionamento!Qtde & " WHERE Int_codigo = " & id_produto_relacionada
        
        'Exclui o complemento da descriÁ„o e atualiza o saldo no produto da NF de saÌda
        Conexao.Execute "UPDATE tbl_Detalhes_Nota SET Complemento_descricao = Null, Saldo = saldo + " & TBRelacionamento!Qtde & " WHERE Int_codigo = " & id_produto_selecionada
        
        TBRelacionamento.MoveNext
    Loop
End If
TBRelacionamento.Close
Conexao.Execute "DELETE from Faturamento_Relacionamento where Id_nota = " & ID_nota & " or ID_nota_relacionada = " & ID_nota

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Sub ProcAtualizaStatusNFe()
On Error GoTo tratar_erro

Data = Date - 1
Set TBComponente = CreateObject("adodb.recordset")
TBComponente.Open "Select NF.ID, NF.int_NotaFiscal, NF.ID_empresa, NF.Serie, NF.Int_status, NFE.Status, NFE.Chave_acesso from (tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_Dados_Nota_Fiscal_NFe NFe on NF.ID = NFe.ID_nota) INNER JOIN Empresa E ON E.Codigo = NF.ID_empresa where E.GNFe = 'True' and NF.dt_DataEmissao >= '" & Format(Data, "Short Date") & "' and NF.Aplicacao = 'P' and int_NotaFiscal IS NOT NULL and NF.TipoNF = 'M1' order by NF.int_NotaFiscal, NF.Serie", Conexao, adOpenKeyset, adLockOptimistic
If TBComponente.EOF = False Then
    Do While TBComponente.EOF = False
        OF = TBComponente!int_NotaFiscal
        Set TBGravar_NFe_Status = CreateObject("adodb.recordset")
        TBGravar_NFe_Status.Open "Select * from NFE012 where CbdEmpCodigo = " & TBComponente!ID_empresa & " and CbdNtfNumero = " & OF & " and CbdNtfSerie = '" & TBComponente!Serie & "' and CbdSituacao <> 0 and CbdAcao <> 'V' order by CbdNtfNumero, CbdNtfSerie", Conexao_NFe, adOpenKeyset, adLockOptimistic
        If TBGravar_NFe_Status.EOF = False Then
            If IsNull(TBGravar_NFe_Status!CbdNFEChaAcesso) = False And TBGravar_NFe_Status!CbdNFEChaAcesso <> "" Then EnviadoTexto = "Imprimir = 'True'" Else EnviadoTexto = "Imprimir = 'False'"
            
            If TBGravar_NFe_Status!CbdStsRetCodigo = 101 Then ProcExcluirRelacionamentoNF TBComponente!ID
            
            If IsNull(TBGravar_NFe_Status!CbdStsRetCodigo) = False And TBGravar_NFe_Status!CbdStsRetCodigo <> "" And (TBGravar_NFe_Status!CbdProcStatus = "P" Or TBComponente!Int_status = 2 And TBGravar_NFe_Status!CbdProcStatus = "N") Then
                TBComponente!status = TBGravar_NFe_Status!CbdStsRetCodigo
            Else
                If TBComponente!Int_status = 2 Then TBComponente!status = -1
            End If
            If IsNull(TBGravar_NFe_Status!CbdSituacao) = False And TBGravar_NFe_Status!CbdSituacao <> 0 Then
                Set TBAliquota = CreateObject("adodb.recordset")
                TBAliquota.Open "Select * from Empresa where codigo = " & TBComponente!ID_empresa, Conexao, adOpenKeyset, adLockOptimistic
                If TBAliquota.EOF = False Then
                    caminho = IIf(IsNull(TBAliquota!Caminho_Nfe), "", TBAliquota!Caminho_Nfe)
                End If
                TBAliquota.Close
                
                OF = TBComponente!int_NotaFiscal
                status = TBGravar_NFe_Status!cbdAcao
                Contador2 = 2
                Do While Contador2 > 0
                    Set GerArqPastas = CreateObject("Scripting.FileSystemObject")
                    If GerArqPastas.FileExists(caminho & "\Empresa " & TBComponente!ID_empresa & " - Serie " & TBComponente!Serie & " - Nota " & OF & " - Status " & status & ".bat") = True Then Kill (caminho & "\Empresa " & TBComponente!ID_empresa & " - Serie " & TBComponente!Serie & " - Nota " & OF & " - Status " & status & ".bat")
                    If status = "C" Then status = "E" Else status = "C"
                    Contador2 = Contador2 - 1
                Loop
            End If
            TBComponente!Chave_acesso = IIf(IsNull(TBGravar_NFe_Status!CbdNFEChaAcesso), "", TBGravar_NFe_Status!CbdNFEChaAcesso)
        Else
            EnviadoTexto = "Imprimir = 'False'"
            TBComponente!Chave_acesso = ""
        End If
        TBComponente.Update
        TBGravar_NFe_Status.Close
        
        Conexao.Execute "Update tbl_Dados_Nota_Fiscal Set " & EnviadoTexto & " where ID = " & TBComponente!ID
        TBComponente.MoveNext
    Loop
End If
TBComponente.Close

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcAtualizaStatusCCe()
On Error GoTo tratar_erro

Data = Date - 1
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select CC.*, NF.int_NotaFiscal, NF.Serie, NF.txt_Razao_Nome from (NF_Carta_Correcao CC INNER JOIN tbl_Dados_Nota_Fiscal NF ON CC.ID_nota = NF.ID) INNER JOIN Empresa E ON E.Codigo = NF.ID_empresa where E.GNFe = 'True' and CC.Data_emissao >= '" & Format(Data, "Short Date") & "' order by CC.ID", Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = False Then
    Do While TBGravar.EOF = False
        OF = TBGravar!int_NotaFiscal
        
        'Verifica cartas de correÁ„o geradas para essa nota
        Set TBAliquota = CreateObject("adodb.recordset")
        TBAliquota.Open "Select * from NF_Carta_Correcao where ID_nota = " & TBGravar!ID_nota & " and ID < " & TBGravar!ID, Conexao, adOpenKeyset, adLockOptimistic
        If TBAliquota.EOF = False Then
            Contador2 = TBAliquota.RecordCount + 1
        Else
            Contador2 = 1
        End If
        TBAliquota.Close
        
        Set TBGravar_NFe = CreateObject("adodb.recordset")
        TBGravar_NFe.Open "Select * from NFE012EVENTOS where CbdNtfNumero = " & OF & " and CbdNtfSerie = '" & TBGravar!Serie & "' and CbdAcao = 'V' and CbdEvenSeq = " & Contador2 & " and CbdEveStsRetCod is not null order by CbdNtfNumero, CbdNtfSerie", Conexao_NFe, adOpenKeyset, adLockOptimistic
        If TBGravar_NFe.EOF = False Then
            TBGravar!CODIGO = TBGravar_NFe!CbdEveStsRetCod
            TBGravar!status = TBGravar_NFe!CbdEveStsRetNome
            TBGravar!Chave_acesso = IIf(IsNull(TBGravar_NFe!CbdEveId), "", TBGravar_NFe!CbdEveId)
            TBGravar.Update
            
            Set TBMaquinas = CreateObject("adodb.recordset")
            TBMaquinas.Open "Select * from Empresa where Codigo = " & TBGravar!ID_empresa & " and GNFe = 'True'", Conexao, adOpenKeyset, adLockOptimistic
            If TBMaquinas.EOF = False Then
                caminho = TBMaquinas!Caminho_Nfe & "\Empresa " & TBGravar!ID_empresa & " - Serie " & TBGravar!Serie & " - Nota " & OF & " - Status CCE.bat"
                Set GerArqPastas = CreateObject("Scripting.FileSystemObject")
                If GerArqPastas.FileExists(caminho) = True Then Kill caminho
            End If
            TBMaquinas.Close
        End If
        TBGravar_NFe.Close
        TBGravar.MoveNext
    Loop
End If
TBGravar.Close

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcEmpenhaProdutoeAtualQtdeEntEmpOrdem(LOTE As Long, Codinterno As String, Qtde_entrada As Double, ID_estoque As Long)
On Error GoTo tratar_erro

'Verifica se a ordem tem pedido vinculado, empenha o produto e atualiza a quantidade de entrada no empenho da ordem
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select VC.Qtde_produzir, VC.qtdeexpedida, VC.CODIGO, PP.Qtde_empenho, PP.Qtde_entrada from (producao_pedidos PP INNER JOIN vendas_carteira VC ON PP.IDcarteira = VC.Codigo) INNER JOIN Producao P ON P.Ordem = PP.Ordem where P.Ordem = " & LOTE & " and P.Desenho = '" & Codinterno & "' and VC.Desenho = '" & Codinterno & "' and PP.Expedicao = 'False' and VC.Cotacao <> 0 order by VC.Prazofinal", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    Do While TBAbrir.EOF = False
        Qtd = IIf(IsNull(TBAbrir!Qtde_produzir), 0, TBAbrir!Qtde_produzir) - IIf(IsNull(TBAbrir!qtdeexpedida), 0, TBAbrir!qtdeexpedida)
        
        'Verifica quantidade j· empenhada
        Set TBFI = CreateObject("adodb.recordset")
        TBFI.Open "Select Sum(Qtde_empenhada) as qtde, Sum(Qtde_saida) as Saida from Estoque_Controle_Empenho_Vendas where ID_carteira = " & TBAbrir!CODIGO, Conexao, adOpenKeyset, adLockOptimistic
        If TBFI.EOF = False Then
            Qtde = IIf(IsNull(TBFI!Qtde), 0, TBFI!Qtde) - IIf(IsNull(TBFI!Saida), 0, TBFI!Saida)
        End If
        TBFI.Close
        
        Dimensoes = Qtd - Qtde
        If Dimensoes > 0 Then
            Set TBComponente = CreateObject("adodb.recordset")
            TBComponente.Open "Select * from Estoque_Controle_Empenho_Vendas", Conexao, adOpenKeyset, adLockOptimistic
            TBComponente.AddNew
            TBComponente!Data = Date
            TBComponente!Responsavel = pubUsuario
            TBComponente!ID_estoque = ID_estoque
            TBComponente!ID_carteira = TBAbrir!CODIGO
            If Qtde_entrada >= Dimensoes Then TBComponente!Qtde_empenhada = Dimensoes Else TBComponente!Qtde_empenhada = Qtde_entrada
            TBComponente.Update
            TBComponente.Close
            Qtde_entrada = Qtde_entrada - Dimensoes
        End If
        If Qtde_entrada <= 0 Then GoTo Prosseguir
        TBAbrir.MoveNext
    Loop
End If
Prosseguir:

ProcAtualizaQtdeEntEmpProd LOTE, Codinterno

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcAtualizaQtdeEntEmpProd(LOTE As Long, Codinterno As String)
On Error GoTo tratar_erro

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select PP.* from producao_pedidos PP INNER JOIN Producao P ON P.Ordem = PP.Ordem where P.Ordem = " & LOTE & " and P.Desenho = '" & Codinterno & "' and PP.Expedicao = 'False' order by PP.IDcarteira", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    
    Set TBTempo = CreateObject("adodb.recordset")
    TBTempo.Open "Select Sum(Entrada) as Qtde_entrada from Estoque_movimentacao where Lote = '" & LOTE & "' and Desenho = '" & Codinterno & "' and (Operacao = 'ENTRADA_ORDEM' or Operacao = 'ENTRADA_ORDEM_PARCIAL')", Conexao, adOpenKeyset, adLockOptimistic
    If TBTempo.EOF = False Then
        Qtde_entrada = IIf(IsNull(TBTempo!Qtde_entrada), 0, TBTempo!Qtde_entrada)
    End If
    
    Do While TBAbrir.EOF = False
        If Qtde_entrada > 0 Then
            If Qtde_entrada > TBAbrir!Qtde_empenho Then
                TBAbrir!Qtde_entrada = TBAbrir!Qtde_empenho
                Qtde_entrada = Qtde_entrada - TBAbrir!Qtde_empenho
            Else
                TBAbrir!Qtde_entrada = Qtde_entrada
                Qtde_entrada = 0
            End If
        Else
            TBAbrir!Qtde_entrada = 0
        End If
        TBAbrir.Update
        TBAbrir.MoveNext
    Loop
End If

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaGridSitProd(Grid As MSFlexGrid, TextoFiltro As String, PBrogresso As USProgressBar, Quebra As Boolean, FiltrarDataPor As String)
On Error GoTo tratar_erro
Dim Coluna          As Integer '0K
Dim Linha           As Integer '0K
Dim Largura_Coluna  As Single '0K
Dim Largura_Campo   As Single '0K
'Debug.print TextoFiltro
With Grid
    Posicao = 0
    Contador3 = 0
    .rows = 0
    .Cols = 0
    .Refresh
    Linha = 0
    OF = 0
    OS = 0
    maquina = ""
    
    Set TBLISTA = CreateObject("adodb.recordset")
    TBLISTA.Open TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
    If TBLISTA.EOF = False Then
        Contador3 = TBLISTA.RecordCount

        Contador = 1
        .rows = TBLISTA.RecordCount
        .Cols = 1
        
        PBrogresso.Min = 0
        PBrogresso.Max = TBLISTA.RecordCount
        PBrogresso.Value = 1
        Contador2 = 0
        Do While TBLISTA.EOF = False
            If OF <> TBLISTA!Ordem Then
                .Row = Linha
                .Col = Coluna
                
                If FrmSituacao_pedido_producao.MSFlexGrid2.Visible = True Then
                    .ColWidth(Coluna) = 6500
                    .TextMatrix(Linha, Coluna) = "OP: " & TBLISTA!Ordem & " . PROD: " & TBLISTA!Desenho & " - " & TBLISTA!Produto & ""
                
                Else
                    
                    .ColWidth(Coluna) = 3500
                    Select Case TBLISTA!Tipo
                        Case "E": Tipo = "PRO"
                        Case "M": Tipo = "SUB"
                        Case "F": Tipo = "COM"
                        Case "S": Tipo = "SER"
                    End Select
                    
                    If TBLISTA!Consignacao = True Then
                    .TextMatrix(Linha, Coluna) = "OP: " & TBLISTA!Ordem & " - " & Tipo & " - Cons."
                    ElseIf TBLISTA!reposicao = True Then
                        .TextMatrix(Linha, Coluna) = "OP: " & TBLISTA!Ordem & " - " & Tipo & " - Rep."
                    Else
                        .TextMatrix(Linha, Coluna) = "OP: " & TBLISTA!Ordem & " - " & Tipo
                    End If
                    
                
            End If
                
                
                
                '.CellFontBold = True
                    
                   
                Set TBAbrir = CreateObject("adodb.recordset")
                If FiltrarDataPor = "Apontamento" And FrmSituacao_pedido_producao.chkPeriodo.Value = 1 Then
                    TBAbrir.Open "Select OSMU.Maquina, OS.Status, OS.IDproducao, OS.fase from (Ordemservico_maq_utilizadas OSMU INNER JOIN Producaofases PF ON OSMU.OS = PF.OS) INNER JOIN Ordemservico OS ON OS.Idproducao = OSMU.OS where OSMU.Ordem = " & TBLISTA!Ordem & " and OSMU.Maquina = '" & TBLISTA!maquina & "' and PF.Data Between '" & FrmSituacao_pedido_producao.txtinicio.Value & "' And '" & FrmSituacao_pedido_producao.txtFinal.Value & "' order by OS.Fase, OSMU.ID", Conexao, adOpenKeyset, adLockOptimistic
                Else
                    TBAbrir.Open "Select * from Ordemservico OS where Ordem = " & TBLISTA!Ordem & " " & FamiliaAntiga & " order by Fase, IDproducao", Conexao, adOpenKeyset, adLockOptimistic
                End If
                 
                If TBAbrir.EOF = False Then
                    Contador = Contador + TBAbrir.RecordCount
                    
                    If Quebra = False Then
                        If Contador > .Cols Then .Cols = Contador
                    Else
                        qt = Contador
                        Do While qt > 0
                            If Contador > .Cols Then
                                If Contador > 8 Then
                                    .Cols = 8
                                    If qt > 8 Then
                                        Contador3 = Contador3 + 1
                                        .rows = Contador3
                                    End If
                                    qt = qt - 7
                                Else
                                    .Cols = Contador
                                    qt = 0
                                End If
                            Else
                                qt = 0
                            End If
                        Loop
                    End If
                    
                    'Exibe o valor de cada linha
                    Do While TBAbrir.EOF = False
                        If OS <> TBAbrir!IDProducao Or OS = TBAbrir!IDProducao And maquina <> TBAbrir!maquina Then
                            Set TBFI = CreateObject("adodb.recordset")
                            If FiltrarDataPor = "Apontamento" And FrmSituacao_pedido_producao.chkPeriodo.Value = 1 Then
                                TBFI.Open "Select OSMU.Maquina, OS.Status, OS.IDproducao from (Ordemservico_maq_utilizadas OSMU INNER JOIN Ordemservico OS ON OSMU.OS = OS.IDproducao) INNER JOIN Producaofases PF ON OSMU.OS = PF.OS where OSMU.OS = " & TBAbrir!IDProducao & " and OSMU.Maquina = '" & TBLISTA!maquina & "' and PF.Data Between '" & FrmSituacao_pedido_producao.txtinicio.Value & "' And '" & FrmSituacao_pedido_producao.txtFinal.Value & "' Group by OS.IDproducao, OSMU.Maquina, OS.Status", Conexao, adOpenKeyset, adLockOptimistic
                            Else
                                TBFI.Open "Select OSMU.Maquina, OS.Status, OS.IDproducao from Ordemservico_maq_utilizadas OSMU INNER JOIN Ordemservico OS ON OSMU.OS = OS.IDproducao where OSMU.OS = " & TBAbrir!IDProducao & " and OSMU.Maquina = '" & TBLISTA!maquina & "' order by OSMU.ID", Conexao, adOpenKeyset, adLockOptimistic
                            End If
                            If TBFI.EOF = False Then
                                
                                Contador = Contador + TBFI.RecordCount
                                If Quebra = False Then
                                    If Contador > .Cols Then .Cols = Contador
                                Else
                                    If Contador > .Cols Then
                                        If Contador > 8 Then .Cols = 8 Else .Cols = Contador
                                    End If
                                End If
                                
                                Do While TBFI.EOF = False
                                    Coluna = Coluna + 1
                                    If Coluna > 7 And Quebra = True Then
                                        Linha = Linha + 1
                                        Coluna = 1
                                    End If
                                    .TextMatrix(Linha, Coluna) = "OS: " & TBFI!IDProducao & " - " & TBFI!maquina
                                    .Row = Linha
                                    .Col = Coluna
                                    .ColWidth(Coluna) = 2850
            
                                    Select Case TBFI!status
                                        Case "Aguardando": .CellBackColor = vbRed 'A produzir
                                        Case "Preparando": .CellBackColor = vbYellow 'Preparando
                                        Case "Produzindo": .CellBackColor = vbYellow 'Produzindo
                                        Case "ConcluÌda": .CellBackColor = vbGreen 'ConcluÌda
                                    End Select
                                    TBFI.MoveNext
                                Loop
                            Else
                                Coluna = Coluna + 1
                                If Coluna > 7 And Quebra = True Then
                                    Linha = Linha + 1
                                    Coluna = 1
                                End If
                                .TextMatrix(Linha, Coluna) = "OS: " & TBAbrir!IDProducao & " - " & TBAbrir!maquina
                                .Row = Linha
                                .Col = Coluna
                                .ColWidth(Coluna) = 2850
                                
                                Select Case TBAbrir!status
                                    Case "Aguardando": .CellBackColor = vbRed 'A produzir
                                    Case "Preparando": .CellBackColor = vbYellow 'Preparando
                                    Case "Produzindo": .CellBackColor = vbYellow 'Produzindo
                                    Case "ConcluÌda": .CellBackColor = vbGreen 'ConcluÌda
                                End Select
                            End If
                        End If
                        
                        OS = TBAbrir!IDProducao
                        maquina = TBAbrir!maquina
                        
                        TBAbrir.MoveNext
                    Loop
                End If
                TBAbrir.Close
                Linha = Linha + 1
                Posicao = Posicao + 1
                Coluna = 0
            End If
            Contador = 1
            Contador2 = Contador2 + 1
            PBrogresso.Value = Contador2
            OF = TBLISTA!Ordem
            TBLISTA.MoveNext
        Loop
    End If
    TBLISTA.Close
End With

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaGridSitProdOLD(Grid As MSFlexGrid, TextoFiltro As String, PBrogresso As USProgressBar, Quebra As Boolean, FiltrarDataPor As String)
On Error GoTo tratar_erro
Dim Coluna          As Integer '0K
Dim Linha           As Integer '0K
Dim Largura_Coluna  As Single '0K
Dim Largura_Campo   As Single '0K
'Debug.print TextoFiltro
With Grid
    Posicao = 0
    Contador3 = 0
    .rows = 0
    .Cols = 0
    .Refresh
    Linha = 0
    OF = 0
    OS = 0
    maquina = ""
    
    Set TBLISTA = CreateObject("adodb.recordset")
    TBLISTA.Open TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
    If TBLISTA.EOF = False Then
        Contador3 = TBLISTA.RecordCount

        Contador = 1
        .rows = TBLISTA.RecordCount
        .Cols = 1
        
        PBrogresso.Min = 0
        PBrogresso.Max = TBLISTA.RecordCount
        PBrogresso.Value = 1
        Contador2 = 0
        Do While TBLISTA.EOF = False
            If OF <> TBLISTA!Ordem Then
                .Row = Linha
                .Col = Coluna
                
                Select Case TBLISTA!Tipo
                    Case "E": Tipo = "PRO"
                    Case "M": Tipo = "SUB"
                    Case "F": Tipo = "COM"
                    Case "S": Tipo = "SER"
                End Select
                
                .ColWidth(Coluna) = 2100
                If TBLISTA!Consignacao = True Then
                    .TextMatrix(Linha, Coluna) = "Ordem: " & TBLISTA!Ordem & " - " & Tipo & " - Cons."
                ElseIf TBLISTA!reposicao = True Then
                        .TextMatrix(Linha, Coluna) = "Ordem: " & TBLISTA!Ordem & " - " & Tipo & " - Rep."
                    Else
                        .TextMatrix(Linha, Coluna) = "Ordem: " & TBLISTA!Ordem & " - " & Tipo
                End If
                .CellFontBold = True
                    
                   
                Set TBAbrir = CreateObject("adodb.recordset")
                If FiltrarDataPor = "Apontamento" And FrmSituacao_pedido_producao.chkPeriodo.Value = 1 Then
                    TBAbrir.Open "Select OSMU.Maquina, OS.Status, OS.IDproducao, OS.fase from (Ordemservico_maq_utilizadas OSMU INNER JOIN Producaofases PF ON OSMU.OS = PF.OS) INNER JOIN Ordemservico OS ON OS.Idproducao = OSMU.OS where OSMU.Ordem = " & TBLISTA!Ordem & " and OSMU.Maquina = '" & TBLISTA!maquina & "' and PF.Data Between '" & FrmSituacao_pedido_producao.txtinicio.Value & "' And '" & FrmSituacao_pedido_producao.txtFinal.Value & "' order by OS.Fase, OSMU.ID", Conexao, adOpenKeyset, adLockOptimistic
                Else
                    TBAbrir.Open "Select * from Ordemservico OS where Ordem = " & TBLISTA!Ordem & " " & FamiliaAntiga & " order by Fase, IDproducao", Conexao, adOpenKeyset, adLockOptimistic
                End If
                 
                If TBAbrir.EOF = False Then
                    Contador = Contador + TBAbrir.RecordCount
                    
                    If Quebra = False Then
                        If Contador > .Cols Then .Cols = Contador
                    Else
                        qt = Contador
                        Do While qt > 0
                            If Contador > .Cols Then
                                If Contador > 8 Then
                                    .Cols = 8
                                    If qt > 8 Then
                                        Contador3 = Contador3 + 1
                                        .rows = Contador3
                                    End If
                                    qt = qt - 7
                                Else
                                    .Cols = Contador
                                    qt = 0
                                End If
                            Else
                                qt = 0
                            End If
                        Loop
                    End If
                    
                    'Exibe o valor de cada linha
                    Do While TBAbrir.EOF = False
                        If OS <> TBAbrir!IDProducao Or OS = TBAbrir!IDProducao And maquina <> TBAbrir!maquina Then
                            Set TBFI = CreateObject("adodb.recordset")
                            If FiltrarDataPor = "Apontamento" And FrmSituacao_pedido_producao.chkPeriodo.Value = 1 Then
                                TBFI.Open "Select OSMU.Maquina, OS.Status, OS.IDproducao from (Ordemservico_maq_utilizadas OSMU INNER JOIN Ordemservico OS ON OSMU.OS = OS.IDproducao) INNER JOIN Producaofases PF ON OSMU.OS = PF.OS where OSMU.OS = " & TBAbrir!IDProducao & " and OSMU.Maquina = '" & TBLISTA!maquina & "' and PF.Data Between '" & FrmSituacao_pedido_producao.txtinicio.Value & "' And '" & FrmSituacao_pedido_producao.txtFinal.Value & "' Group by OS.IDproducao, OSMU.Maquina, OS.Status", Conexao, adOpenKeyset, adLockOptimistic
                            Else
                                TBFI.Open "Select OSMU.Maquina, OS.Status, OS.IDproducao from Ordemservico_maq_utilizadas OSMU INNER JOIN Ordemservico OS ON OSMU.OS = OS.IDproducao where OSMU.OS = " & TBAbrir!IDProducao & " and OSMU.Maquina = '" & TBLISTA!maquina & "' order by OSMU.ID", Conexao, adOpenKeyset, adLockOptimistic
                            End If
                            If TBFI.EOF = False Then
                                
                                Contador = Contador + TBFI.RecordCount
                                If Quebra = False Then
                                    If Contador > .Cols Then .Cols = Contador
                                Else
                                    If Contador > .Cols Then
                                        If Contador > 8 Then .Cols = 8 Else .Cols = Contador
                                    End If
                                End If
                                
                                Do While TBFI.EOF = False
                                    Coluna = Coluna + 1
                                    If Coluna > 7 And Quebra = True Then
                                        Linha = Linha + 1
                                        Coluna = 1
                                    End If
                                    .TextMatrix(Linha, Coluna) = "OS: " & TBFI!IDProducao & " - " & TBFI!maquina
                                    .Row = Linha
                                    .Col = Coluna
                                    .ColWidth(Coluna) = 1850
                                    Select Case TBFI!status
                                        Case "Aguardando": .CellBackColor = vbRed 'A produzir
                                        Case "Preparando": .CellBackColor = vbYellow 'Preparando
                                        Case "Produzindo": .CellBackColor = vbYellow 'Produzindo
                                        Case "ConcluÌda": .CellBackColor = vbGreen 'ConcluÌda
                                    End Select
                                    TBFI.MoveNext
                                Loop
                            Else
                                Coluna = Coluna + 1
                                If Coluna > 7 And Quebra = True Then
                                    Linha = Linha + 1
                                    Coluna = 1
                                End If
                                .TextMatrix(Linha, Coluna) = "OS: " & TBAbrir!IDProducao & " - " & TBAbrir!maquina
                                .Row = Linha
                                .Col = Coluna
                                .ColWidth(Coluna) = 1850
                                
                                Select Case TBAbrir!status
                                    Case "Aguardando": .CellBackColor = vbRed 'A produzir
                                    Case "Preparando": .CellBackColor = vbYellow 'Preparando
                                    Case "Produzindo": .CellBackColor = vbYellow 'Produzindo
                                    Case "ConcluÌda": .CellBackColor = vbGreen 'ConcluÌda
                                End Select
                            End If
                        End If
                        
                        OS = TBAbrir!IDProducao
                        maquina = TBAbrir!maquina
                        
                        TBAbrir.MoveNext
                    Loop
                End If
                TBAbrir.Close
                Linha = Linha + 1
                Posicao = Posicao + 1
                Coluna = 0
            End If
            Contador = 1
            Contador2 = Contador2 + 1
            PBrogresso.Value = Contador2
            OF = TBLISTA!Ordem
            TBLISTA.MoveNext
        Loop
    End If
    TBLISTA.Close
End With

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcVerifCFOPST(ID_CF As Long, UF As String)
On Error GoTo tratar_erro

Valido = False
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select R.*, CST.CST from Regioes R INNER JOIN CST ON CST.ID_UF = R.ID where R.UF = '" & UF & "' and CST.ID_CF = " & ID_CF, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    'Verifica se tem a CFOP cadastrada
    If TBAbrir!regiao = "DE" Then TextoFiltro = "5.401" Else TextoFiltro = "6.401"
    Set TBCFOP = CreateObject("adodb.recordset")
    TBCFOP.Open "Select * from tbl_NaturezaOperacao where id_CFOP = '" & TextoFiltro & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBCFOP.EOF = False Then
        Valido = True
        IDAntigo = TBCFOP!IDCountCfop
        FamiliaAntiga = TBCFOP!ID_CFOP
        Familiatext = TBCFOP!Txt_descricao
        If IsNull(TBAbrir!CST) = False And TBAbrir!CST <> "" Then Letra = TBAbrir!CST
    End If
    TBCFOP.Close
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcBuscarClienteNS(CnpjEmpresa As String)
On Error GoTo tratar_erro
Dim resposta As String
Dim p As Object

If UF <> "" And CNPJCliente <> "" Then
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from empresa where codigo = '" & IDempresa & "'", Conexao, adOpenKeyset, adLockReadOnly
If TBAbrir.EOF = False Then
CNPJ_Empresa = ReturnNumbersOnly(TBAbrir!CNPJ)
End If
TBAbrir.Close

If CNPJ_Empresa = "34270461000104" Then
CNPJ_Empresa = ReturnNumbersOnly("07.758.985/0001-68")
End If

If CNPJCliente <> "" Then
  resposta = consultarCadastroContribuinte(CNPJ_Empresa, UF, ReturnNumbersOnly(CNPJCliente), "CNPJ")
Else
  resposta = consultarCadastroContribuinte(CNPJ_Empresa, UF, ReturnNumbersOnly(cpfCliente), "CPF")
End If

'Debug.print resposta
status = LerDadosJSON(resposta, "status", "", "")
   If status = "200" Then
      Set p = JSON.parse(resposta)
      JSON.parse (resposta)
      
      If p.Item("retConsCad").Item("infCons").Item("xMotivo") = "CNPJ da consulta nao cadastrado como contribuinte na UF. CNPJ: 16740838000151" Then
      USMsgBox "N„o ser· possivel buscar o cadastro desse cnpj pois n„o È contribuinte do ICMS"
      Exit Sub
      End If
      
      If p.Item("retConsCad").Item("infCons").Item("xMotivo") = "RejeiÁ„o: CNPJ da consulta n„o cadastrado como contribuinte na UF" Then
      ProcBuscaClienteSintegra (ReturnNumbersOnly(CNPJCliente))
      Exit Sub
      End If
      
      
      If p.Item("retConsCad").Item("infCons").Item("xMotivo") = "Rejeicao: Sigla da UF da consulta difere da UF do Web Service" Then
      ProcBuscaClienteSintegra (ReturnNumbersOnly(CNPJCliente))
      Exit Sub
      End If
      
      If p.Item("retConsCad").Item("infCons").Item("xMotivo") = "Rejeicao: CPF da consulta nao cadastrado como contribuinte na UF" Then
      ProcBuscaClienteSintegra (ReturnNumbersOnly(CNPJCliente))
      Exit Sub
      End If
      
      If p.Item("retConsCad").Item("infCons").Item("xMotivo") = "RejeiÁ„o: UF n„o fornece consulta por CPF" Then
      ProcBuscaClienteSintegra (ReturnNumbersOnly(CNPJCliente))
      Exit Sub
      End If
      
      If p.Item("retConsCad").Item("infCons").Item("xMotivo") = "Rejeicao: Solicitante nao habilitado para emissao da NF-e" Then
      ProcBuscaClienteSintegra (ReturnNumbersOnly(CNPJCliente))
      Exit Sub
      End If


      UF = Trim(p.Item("retConsCad").Item("infCons").Item("UF"))
      
     ' If TipoEmpresa = "1" Then
      RG_IE = Trim(p.Item("retConsCad").Item("infCons").Item("infCad").Item(1).Item("IE"))
    '  End If
      
      NomeRazao = Trim(p.Item("retConsCad").Item("infCons").Item("infCad").Item(1).Item("xNome"))
      NomeFantasia = Trim(p.Item("retConsCad").Item("infCons").Item("infCad").Item(1).Item("xNome"))
      Endereco = Trim(p.Item("retConsCad").Item("infCons").Item("infCad").Item(1).Item("ender").Item("xLgr"))
      Numero = Trim(p.Item("retConsCad").Item("infCons").Item("infCad").Item(1).Item("ender").Item("nro"))
      Bairro = Trim(p.Item("retConsCad").Item("infCons").Item("infCad").Item(1).Item("ender").Item("xBairro"))
      If Trim(p.Item("retConsCad").Item("infCons").Item("infCad").Item(1).Item("ender").Item("xMun")) <> "" Then
      Cidade = Trim(p.Item("retConsCad").Item("infCons").Item("infCad").Item(1).Item("ender").Item("xMun"))
      End If
      CEP = Trim(p.Item("retConsCad").Item("infCons").Item("infCad").Item(1).Item("ender").Item("CEP"))
      RegimeTributario = IIf(p.Item("retConsCad").Item("infCons").Item("infCad").Item(1).Item("xRegApur") = "NORMAL - REGIME PERI”DICO DE APURA«√O", "Lucro presumido", "Simples Nacional")
      Categoria = "A"
      
   Else
      USMsgBox resposta, vbCritical, "CAPRIND v5.0"
   End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcEmpenharProdEstoque(ID_empresa As Integer, ID_carteira As Long, Codinterno As String, VerifEmpEmpresa As Boolean, MRP As Boolean, QtdeVendidaProd As Double)
On Error GoTo tratar_erro

Permitido2 = True
If VerifEmpEmpresa = True Then
    Set TBAliquota = CreateObject("adodb.recordset")
    TBAliquota.Open "Select Codigo from Empresa where Codigo = " & ID_empresa & " and Ativar_empenho_autom = 'True'", Conexao, adOpenKeyset, adLockOptimistic
    If TBAliquota.EOF = True Then
    Permitido2 = False
    End If
    
    TBAliquota.Close
End If
If MRP = True Then
    TextoFiltro = ""
    TextoFiltro1 = ""
Else
    TextoFiltro = " and ECEV.ID_carteira <> " & ID_carteira
    TextoFiltro1 = " and ID_carteira <> " & ID_carteira
End If

If Permitido2 = True Then
    qtdeliberar = 0
    Set TBEstoque = CreateObject("adodb.recordset")
    TBEstoque.Open "Select Sum(ECEV.Qtde_empenhada - ECEV.Qtde_saida) as qtdeliberar from Estoque_Controle_Empenho_Vendas ECEV INNER JOIN Estoque_Controle EC ON EC.IDestoque = ECEV.ID_estoque where EC.ID_empresa = " & ID_empresa & " and EC.Desenho = '" & Codinterno & "'" & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
    If TBEstoque.EOF = False Then
        qtdeliberar = IIf(IsNull(TBEstoque!qtdeliberar), 0, TBEstoque!qtdeliberar)
    End If
    qtdeliberada = 0
    Set TBEstoque = CreateObject("adodb.recordset")
    TBEstoque.Open "Select Qtde_empenhar from Qtde_empenhada_produto where ID_empresa = " & ID_empresa & " and Codigo = '" & Codinterno & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBEstoque.EOF = False Then
        qtdeliberada = IIf(IsNull(TBEstoque!Qtde_empenhar), 0, TBEstoque!Qtde_empenhar)
    End If
    TBEstoque.Close
        
    QTLOTE = FunVerificaQtdeEstoque(Codinterno, ID_empresa, "") - (qtdeliberar + qtdeliberada)
    If QTLOTE > 0 Then
        'Empenha produto em estoque para esta venda
        Set TBEstoque = CreateObject("adodb.recordset")
        TBEstoque.Open "Select * from Estoque_produtos where ID_empresa = " & ID_empresa & " and Desenho = '" & Codinterno & "' and Estoque_real > 0 and Liberado = 'SIM' order by Data", Conexao, adOpenKeyset, adLockOptimistic
        If TBEstoque.EOF = False Then
            Do While TBEstoque.EOF = False
                
                'Verifica qtde. empenhada do lote
                Set TBFI = CreateObject("adodb.recordset")
                TBFI.Open "Select Sum(Qtde_empenhada - Qtde_saida) as Qtd from Estoque_Controle_Empenho_Vendas where ID_estoque = " & TBEstoque!IDEstoque & TextoFiltro1, Conexao, adOpenKeyset, adLockOptimistic
                If TBFI.EOF = False Then
                    Qtd = IIf(IsNull(TBFI!Qtd), 0, TBFI!Qtd)
                End If
                Set TBFI = CreateObject("adodb.recordset")
                TBFI.Open "Select Sum(Quantidade - Qtde_saida) as qt from Producao_NF_Consignada where IDestoque = " & TBEstoque!IDEstoque, Conexao, adOpenKeyset, adLockOptimistic
                If TBFI.EOF = False Then
                    qt = IIf(IsNull(TBFI!qt), 0, TBFI!qt)
                End If
                TBFI.Close
                Qtd = Qtd + qt
                
                If (TBEstoque!estoque_real - Qtd) > 0 Then
                    Set TBFI = CreateObject("adodb.recordset")
                    TBFI.Open "Select * from Estoque_Controle_Empenho_Vendas where ID_estoque = " & TBEstoque!IDEstoque & " and ID_carteira = " & ID_carteira, Conexao, adOpenKeyset, adLockOptimistic
                    If TBFI.EOF = True Then TBFI.AddNew
                    TBFI!Data = Date
                    TBFI!Responsavel = pubUsuario
                    TBFI!ID_carteira = ID_carteira
                    TBFI!ID_estoque = TBEstoque!IDEstoque
                    
                    If (TBEstoque!estoque_real - Qtd) >= QtdeVendidaProd Then
                        TBFI!Qtde_empenhada = QtdeVendidaProd
                        QtdeVendidaProd = 0
                    Else
                        TBFI!Qtde_empenhada = TBEstoque!estoque_real - Qtd
                        QtdeVendidaProd = QtdeVendidaProd - TBFI!Qtde_empenhada
                    End If
                    QuantSolicitado = QuantSolicitado - TBFI!Qtde_empenhada
                    
                    TBFI.Update
                    TBFI.Close
                    If QtdeVendidaProd <= 0 Then Exit Sub
                End If
                TBEstoque.MoveNext
            Loop
        End If
        TBEstoque.Close
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcEmpenharProdProduzindo(ID_empresa As Integer, ID_carteira As Long, Codinterno As String, PrazoFinal As Date, VerifEmpEmpresa As Boolean)
On Error GoTo tratar_erro

Permitido2 = True
If VerifEmpEmpresa = True Then
    Set TBAliquota = CreateObject("adodb.recordset")
    TBAliquota.Open "Select Codigo from Empresa where Codigo = " & ID_empresa & " and Ativar_empenho_autom_prod = 'True'", Conexao, adOpenKeyset, adLockOptimistic
    If TBAliquota.EOF = True Then Permitido2 = False
    TBAliquota.Close
End If

If Permitido2 = True Then
    qtdeliberar = 0
    Set TBEstoque = CreateObject("adodb.recordset")
    TBEstoque.Open "Select Qtde_requisitar from Qtde_total_empenhada_produto_produzindo where ID_empresa = " & ID_empresa & " and Desenho = '" & Codinterno & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBEstoque.EOF = False Then
        qtdeliberar = IIf(IsNull(TBEstoque!Qtde_requisitar), 0, TBEstoque!Qtde_requisitar)
    End If
        
    QTLOTE = FunVerificaQtdeProduzindo(Codinterno, ID_empresa) - qtdeliberar
    If QTLOTE > 0 Then
        'Empenha produto em produÁ„o para esta venda
        Set TBEstoque = CreateObject("adodb.recordset")
        TBEstoque.Open "Select PPO.Ordem, PPO.Qtde_produzida - ISNULL(EEPPL.Qtde_entrada, 0) as qtdeliberada from (Qtde_produzida_produto_ordem PPO LEFT JOIN Qtde_entrada_estoque_produto_produzido_lote EEPPL ON EEPPL.Lote = PPO.Ordem) INNER JOIN Producao P ON P.Ordem = PPO.Ordem where PPO.ID_empresa = " & ID_empresa & " and PPO.Desenho = '" & Codinterno & "' and P.PrazoEntrega <= '" & PrazoFinal & "' and PPO.Qtde_produzida - ISNULL(EEPPL.Qtde_entrada, 0) > 0 and P.DtValidacao_custo IS NULL order by PPO.Ordem", Conexao, adOpenKeyset, adLockOptimistic
        If TBEstoque.EOF = False Then
            Do While TBEstoque.EOF = False
                
                'Verifica qtde. empenhada da ordem
                Qtd = 0
                Set TBFI = CreateObject("adodb.recordset")
                TBFI.Open "Select Sum(Qtde_empenho - Qtde_entrada) as Qtd from Producao_pedidos where Ordem = " & TBEstoque!Ordem, Conexao, adOpenKeyset, adLockOptimistic
                If TBFI.EOF = False Then
                    Qtd = IIf(IsNull(TBFI!Qtd), 0, TBFI!Qtd)
                End If
                                
                If (TBEstoque!qtdeliberada - Qtd) > 0 Then
                    Set TBGravar = CreateObject("adodb.recordset")
                    TBGravar.Open "Select * from Producao_pedidos", Conexao, adOpenKeyset, adLockOptimistic
                    TBGravar.AddNew
                    TBGravar!Data = Date
                    TBGravar!Responsavel = pubUsuario
                    TBGravar!IDcarteira = ID_carteira
                    TBGravar!Ordem = TBEstoque!Ordem
                    If (TBEstoque!qtdeliberada - Qtd) >= QuantSolicitado Then
                        TBGravar!Qtde_empenho = QuantSolicitado
                        QuantSolicitado = 0
                    Else
                        TBGravar!Qtde_empenho = TBEstoque!qtdeliberada - Qtd
                        QuantSolicitado = QuantSolicitado - TBGravar!Qtde_empenho
                    End If
                    TBGravar.Update
                    TBGravar.Close
                    If QuantSolicitado <= 0 Then Exit Sub
                End If
                TBEstoque.MoveNext
            Loop
        End If
        TBEstoque.Close
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcExcluirEmpenhos(ID_empresa As Integer, ID_carteira As Long, VerifEmpEmpresa As Boolean)
On Error GoTo tratar_erro

Permitido1 = True
If VerifEmpEmpresa = True Then
    Set TBAliquota = CreateObject("adodb.recordset")
    TBAliquota.Open "Select * from Empresa where Codigo = " & ID_empresa & " and Ativar_empenho_autom = 'True'", Conexao, adOpenKeyset, adLockOptimistic
    If TBAliquota.EOF = True Then Permitido1 = False
    TBAliquota.Close
End If

If Permitido1 = True Then
    Conexao.Execute "DELETE from Estoque_Controle_Empenho_Vendas where ID_carteira = " & ID_carteira
    Conexao.Execute "DELETE from Producao_pedidos where IDcarteira = " & ID_carteira
End If

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcValidarRegistros(Lista As ListView, FormularioValid As String)
On Error GoTo tratar_erro

If FormularioValid = "Outros/SolicitaÁ„o/Autorizar solicitaÁ„o" Or FormularioValid = "Compras/Pedido/Aprovar" Then
If frmCompras_Pedido.txtStatus = "RECEBIDO" Or frmCompras_Pedido.txtStatus = "RECEBIDO PARCIAL" Then
    USMsgBox "N„o È permitido mudar a autorizaÁ„o do pedido com status " & frmCompras_Pedido.txtStatus, vbInformation, "CAPRIND v5.0"
    Exit Sub
End If

TextoMsg = "aprovar/cancelar a aprovaÁ„o"
Else
TextoMsg = "validar/cancelar a validaÁ„o"
End If


If FormularioValid = "Qualidade/Plano de inspeÁ„o" Then
    If frmPlanoinspecao_validacao.Opt_produto = True Then
        If frmPlanoinspecao.txtdesenho = "" Then
            USMsgBox ("Informe o plano antes de " & TextoMsg & "."), vbExclamation, "CAPRIND v5.0"
            Exit Sub
        Else
            GoTo Validacao
        End If
    End If
End If

Permitido = False
With Lista
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If USMsgBox("Deseja realmente " & TextoMsg & " do(s) registro(s) selecionado(s)?", vbYesNo, "CAPRIND v5.0") = vbNo Then
                Exit Sub
            Else
                Permitido = True
                GoTo Validacao
            End If
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) registro(s) antes de " & TextoMsg & "."), vbExclamation, "CAPRIND v5.0"
Else
Validacao:
    Formulario = FormularioValid
    frmValidar.Show 1
End If

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcVerifLiberacaoTab(SSTab As SSTab, NTab As Integer, Formulario As String)
On Error GoTo tratar_erro

With SSTab
    Set TBAcessos = CreateObject("adodb.recordset")
    TBAcessos.Open "Select * from Acessos where IDUsuario = " & pubIDUsuario & " and Acesso = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBAcessos.EOF = True Then Acessos = False Else Acessos = True
    TBAcessos.Close
    If Acessos = False Then
        .TabVisible(NTab) = False
        Contador = Contador - 1
        .TabsPerRow = Contador
    Else
        .TabVisible(NTab) = True
        .TabsPerRow = Contador
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Function FunVerifValidacaoRegistro(Acao As String, DtValidacao As TextBox, NomeRegistroPrincipal As String, NomeRegistro As String, MascPrincipal As Boolean) As Boolean
On Error GoTo tratar_erro

'ANTIGO
FunVerifValidacaoRegistro = True
If DtValidacao <> "" Then
    If MascPrincipal = True Then
        USMsgBox ("N„o È permitido " & Acao & " " & NomeRegistro & ", pois o " & NomeRegistroPrincipal & " j· foi validado."), vbExclamation, "CAPRIND v5.0"
        FunVerifValidacaoRegistro = False
    Else
        USMsgBox ("N„o È permitido " & Acao & " " & NomeRegistro & ", pois a " & NomeRegistroPrincipal & " j· foi validada."), vbExclamation, "CAPRIND v5.0"
        FunVerifValidacaoRegistro = False
    End If
End If

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Sub ProcCopiarEstrutura(CodprodutoAnt As Long, CodprodutoNovo As Long, VersaoAnt As String, VersaoNova As String)
On Error GoTo tratar_erro

If VersaoAnt <> "" Then TextoFiltro = " and Versao = '" & VersaoAnt & "'" Else TextoFiltro = ""
Set TBNivel2 = CreateObject("adodb.recordset")
TBNivel2.Open "Select * from projconjunto where codproduto = " & CodprodutoAnt & TextoFiltro & " order by Codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBNivel2.EOF = False Then
    Do While TBNivel2.EOF = False
        VersaoAnt = TBNivel2!versao
        VersaoNova = IIf(VersaoNova = "", TBNivel2!versao, VersaoNova)
        Set TBGravar = CreateObject("adodb.recordset")
        TBGravar.Open "Select * from projconjunto where codProduto = " & CodprodutoNovo & " and Versao = '" & VersaoNova & "' and Desenho = '" & TBNivel2!Desenho & "' and Versao_desenho = '" & TBNivel2!Versao_desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBGravar.EOF = True Then TBGravar.AddNew
        TBGravar!Codproduto = CodprodutoNovo
        TBGravar!Posicao = TBNivel2!Posicao
        TBGravar!Desenho = TBNivel2!Desenho
        TBGravar!Versao_desenho = TBNivel2!Versao_desenho
        TBGravar!Descricao = TBNivel2!Descricao
        TBGravar!PesoMetro = TBNivel2!PesoMetro
        TBGravar!PesoTotal = TBNivel2!PesoTotal
        TBGravar!Percentual_perda = TBNivel2!Percentual_perda
        TBGravar!ID_partnumber_fabricante = TBNivel2!ID_partnumber_fabricante
        TBGravar!Obs = TBNivel2!Obs
        TBGravar!quantidade = TBNivel2!quantidade
        TBGravar!Peso = TBNivel2!Peso
        TBGravar!Unidade = TBNivel2!Unidade
        TBGravar!Dimensoes = TBNivel2!Dimensoes
        TBGravar!valor = TBNivel2!valor
        TBGravar!ValorTotal = TBNivel2!ValorTotal
        TBGravar!Un_Kg = TBNivel2!Un_Kg
        TBGravar!versao = VersaoNova
        TBGravar!Tipo = TBNivel2!Tipo
        TBGravar.Update
        ProcCopiarDescVersaoEstrutura CodprodutoAnt, VersaoAnt, CodprodutoNovo, VersaoNova
        TBGravar.Close
        
        ProcCopiarEstruturaN3 TBNivel2!Desenho, TBNivel2!Versao_desenho, VersaoNova
        TBNivel2.MoveNext
    Loop
End If
TBNivel2.Close

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCopiarEstruturaN3(Desenho As String, VersaoAnt As String, VersaoNova As String)
On Error GoTo tratar_erro

Set TBNivel3 = CreateObject("adodb.recordset")
TBNivel3.Open "Select PC.* from projconjunto PC INNER JOIN Projproduto P ON P.Codproduto = PC.Codproduto where P.Desenho = '" & Desenho & "' and PC.Versao = '" & VersaoAnt & "' order by PC.Codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBNivel3.EOF = False Then
    Do While TBNivel3.EOF = False
        If VersaoNova <> "" Then TextoFiltro = " and Versao = '" & VersaoNova & "'" Else TextoFiltro = ""
        Set TBGravar = CreateObject("adodb.recordset")
        TBGravar.Open "Select * from projconjunto where codProduto = " & TBNivel3!Codproduto & TextoFiltro & " and Desenho = '" & TBNivel3!Desenho & "' and Versao_desenho = '" & TBNivel3!Versao_desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBGravar.EOF = True Then TBGravar.AddNew
        TBGravar!Codproduto = TBNivel3!Codproduto
        TBGravar!Posicao = TBNivel3!Posicao
        TBGravar!Desenho = TBNivel3!Desenho
        TBGravar!Versao_desenho = TBNivel3!Versao_desenho
        TBGravar!Descricao = TBNivel3!Descricao
        TBGravar!PesoMetro = TBNivel3!PesoMetro
        TBGravar!PesoTotal = TBNivel3!PesoTotal
        TBGravar!Percentual_perda = TBNivel3!Percentual_perda
        TBGravar!ID_partnumber_fabricante = TBNivel3!ID_partnumber_fabricante
        TBGravar!Obs = TBNivel3!Obs
        TBGravar!quantidade = TBNivel3!quantidade
        TBGravar!Peso = TBNivel3!Peso
        TBGravar!Unidade = TBNivel3!Unidade
        TBGravar!Dimensoes = TBNivel3!Dimensoes
        TBGravar!valor = TBNivel3!valor
        TBGravar!ValorTotal = TBNivel3!ValorTotal
        TBGravar!Un_Kg = TBNivel3!Un_Kg
        TBGravar!Tipo = TBNivel3!Tipo

        If VersaoNova = "" Then TBGravar!versao = TBNivel3!versao Else TBGravar!versao = VersaoNova
        TBGravar.Update
        If VersaoNova <> "" Then ProcCopiarDescVersaoEstrutura TBNivel3!Codproduto, VersaoAnt, TBNivel3!Codproduto, VersaoNova
        TBGravar.Close
                
        ProcCopiarEstruturaN4 TBNivel3!Desenho, TBNivel3!Versao_desenho, VersaoNova
        TBNivel3.MoveNext
    Loop
End If
TBNivel3.Close

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCopiarEstruturaN4(Desenho As String, VersaoAnt As String, VersaoNova As String)
On Error GoTo tratar_erro

Set TBNivel4 = CreateObject("adodb.recordset")
TBNivel4.Open "Select PC.* from projconjunto PC INNER JOIN Projproduto P ON P.Codproduto = PC.Codproduto where P.Desenho = '" & Desenho & "' and PC.Versao = '" & VersaoAnt & "' order by PC.Codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBNivel4.EOF = False Then
    Do While TBNivel4.EOF = False
        If VersaoNova <> "" Then TextoFiltro = " and Versao = '" & VersaoNova & "'" Else TextoFiltro = ""
        Set TBGravar = CreateObject("adodb.recordset")
        TBGravar.Open "Select * from projconjunto where codProduto = " & TBNivel4!Codproduto & TextoFiltro & " and Desenho = '" & TBNivel4!Desenho & "' and Versao_desenho = '" & TBNivel4!Versao_desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBGravar.EOF = True Then TBGravar.AddNew
        TBGravar!Codproduto = TBNivel4!Codproduto
        TBGravar!Posicao = TBNivel4!Posicao
        TBGravar!Desenho = TBNivel4!Desenho
        TBGravar!Versao_desenho = TBNivel4!Versao_desenho
        TBGravar!Descricao = TBNivel4!Descricao
        TBGravar!PesoMetro = TBNivel4!PesoMetro
        TBGravar!PesoTotal = TBNivel4!PesoTotal
        TBGravar!Percentual_perda = TBNivel4!Percentual_perda
        TBGravar!ID_partnumber_fabricante = TBNivel4!ID_partnumber_fabricante
        TBGravar!Obs = TBNivel4!Obs
        TBGravar!quantidade = TBNivel4!quantidade
        TBGravar!Peso = TBNivel4!Peso
        TBGravar!Unidade = TBNivel4!Unidade
        TBGravar!Dimensoes = TBNivel4!Dimensoes
        TBGravar!valor = TBNivel4!valor
        TBGravar!ValorTotal = TBNivel4!ValorTotal
        TBGravar!Un_Kg = TBNivel4!Un_Kg
        TBGravar!Tipo = TBNivel4!Tipo
        
        If VersaoNova = "" Then TBGravar!versao = TBNivel4!versao Else TBGravar!versao = VersaoNova
        TBGravar.Update
        If VersaoNova <> "" Then ProcCopiarDescVersaoEstrutura TBNivel4!Codproduto, VersaoAnt, TBNivel4!Codproduto, VersaoNova
        TBGravar.Close
        
        ProcCopiarEstruturaN5 TBNivel4!Desenho, TBNivel4!Versao_desenho, VersaoNova
        TBNivel4.MoveNext
    Loop
End If
TBNivel4.Close

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCopiarEstruturaN5(Desenho As String, VersaoAnt As String, VersaoNova As String)
On Error GoTo tratar_erro

Set TBNivel5 = CreateObject("adodb.recordset")
TBNivel5.Open "Select PC.* from projconjunto PC INNER JOIN Projproduto P ON P.Codproduto = PC.Codproduto where P.Desenho = '" & Desenho & "' and PC.Versao = '" & VersaoAnt & "' order by PC.Codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBNivel5.EOF = False Then
    Do While TBNivel5.EOF = False
        If VersaoNova <> "" Then TextoFiltro = " and Versao = '" & VersaoNova & "'" Else TextoFiltro = ""
        Set TBGravar = CreateObject("adodb.recordset")
        TBGravar.Open "Select * from projconjunto where codProduto = " & TBNivel5!Codproduto & TextoFiltro & " and Desenho = '" & TBNivel5!Desenho & "' and Versao_desenho = '" & TBNivel5!Versao_desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBGravar.EOF = True Then TBGravar.AddNew
        TBGravar!Codproduto = TBNivel5!Codproduto
        TBGravar!Posicao = TBNivel5!Posicao
        TBGravar!Desenho = TBNivel5!Desenho
        TBGravar!Versao_desenho = TBNivel5!Versao_desenho
        TBGravar!Descricao = TBNivel5!Descricao
        TBGravar!PesoMetro = TBNivel5!PesoMetro
        TBGravar!PesoTotal = TBNivel5!PesoTotal
        TBGravar!Percentual_perda = TBNivel5!Percentual_perda
        TBGravar!ID_partnumber_fabricante = TBNivel5!ID_partnumber_fabricante
        TBGravar!Obs = TBNivel5!Obs
        TBGravar!quantidade = TBNivel5!quantidade
        TBGravar!Peso = TBNivel5!Peso
        TBGravar!Unidade = TBNivel5!Unidade
        TBGravar!Dimensoes = TBNivel5!Dimensoes
        TBGravar!valor = TBNivel5!valor
        TBGravar!ValorTotal = TBNivel5!ValorTotal
        TBGravar!Un_Kg = TBNivel5!Un_Kg
        If VersaoNova = "" Then TBGravar!versao = TBNivel5!versao Else TBGravar!versao = VersaoNova
        TBGravar.Update
        If VersaoNova <> "" Then ProcCopiarDescVersaoEstrutura TBNivel5!Codproduto, VersaoAnt, TBNivel5!Codproduto, VersaoNova
        TBGravar.Close
        
        ProcCopiarEstruturaN6 TBNivel5!Desenho, TBNivel5!Versao_desenho, VersaoNova
        TBNivel5.MoveNext
    Loop
End If
TBNivel5.Close

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCopiarEstruturaN6(Desenho As String, VersaoAnt As String, VersaoNova As String)
On Error GoTo tratar_erro

Set TBNivel6 = CreateObject("adodb.recordset")
TBNivel6.Open "Select PC.* from projconjunto PC INNER JOIN Projproduto P ON P.Codproduto = PC.Codproduto where P.Desenho = '" & Desenho & "' and PC.Versao = '" & VersaoAnt & "' order by PC.Codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBNivel6.EOF = False Then
    Do While TBNivel6.EOF = False
        If VersaoNova <> "" Then TextoFiltro = " and Versao = '" & VersaoNova & "'" Else TextoFiltro = ""
        Set TBGravar = CreateObject("adodb.recordset")
        TBGravar.Open "Select * from projconjunto where codProduto = " & TBNivel6!Codproduto & TextoFiltro & " and Desenho = '" & TBNivel6!Desenho & "' and Versao_desenho = '" & TBNivel6!Versao_desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBGravar.EOF = True Then TBGravar.AddNew
        TBGravar!Codproduto = TBNivel6!Codproduto
        TBGravar!Posicao = TBNivel6!Posicao
        TBGravar!Desenho = TBNivel6!Desenho
        TBGravar!Versao_desenho = TBNivel6!Versao_desenho
        TBGravar!Descricao = TBNivel6!Descricao
        TBGravar!PesoMetro = TBNivel6!PesoMetro
        TBGravar!PesoTotal = TBNivel6!PesoTotal
        TBGravar!Percentual_perda = TBNivel6!Percentual_perda
        TBGravar!ID_partnumber_fabricante = TBNivel6!ID_partnumber_fabricante
        TBGravar!Obs = TBNivel6!Obs
        TBGravar!quantidade = TBNivel6!quantidade
        TBGravar!Peso = TBNivel6!Peso
        TBGravar!Unidade = TBNivel6!Unidade
        TBGravar!Dimensoes = TBNivel6!Dimensoes
        TBGravar!valor = TBNivel6!valor
        TBGravar!ValorTotal = TBNivel6!ValorTotal
        TBGravar!Un_Kg = TBNivel6!Un_Kg
        If VersaoNova = "" Then TBGravar!versao = TBNivel6!versao Else TBGravar!versao = VersaoNova
        TBGravar.Update
        If VersaoNova <> "" Then ProcCopiarDescVersaoEstrutura TBNivel6!Codproduto, VersaoAnt, TBNivel6!Codproduto, VersaoNova
        TBGravar.Close
        
        ProcCopiarEstruturaN7 TBNivel6!Desenho, TBNivel6!Versao_desenho, VersaoNova
        TBNivel6.MoveNext
    Loop
End If
TBNivel6.Close

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCopiarEstruturaN7(Desenho As String, VersaoAnt As String, VersaoNova As String)
On Error GoTo tratar_erro

Set TBNivel7 = CreateObject("adodb.recordset")
TBNivel7.Open "Select PC.* from projconjunto PC INNER JOIN Projproduto P ON P.Codproduto = PC.Codproduto where P.Desenho = '" & Desenho & "' and PC.Versao = '" & VersaoAnt & "' order by PC.Codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBNivel7.EOF = False Then
    Do While TBNivel7.EOF = False
        If VersaoNova <> "" Then TextoFiltro = " and Versao = '" & VersaoNova & "'" Else TextoFiltro = ""
        Set TBGravar = CreateObject("adodb.recordset")
        TBGravar.Open "Select * from projconjunto where codProduto = " & TBNivel7!Codproduto & TextoFiltro & " and Desenho = '" & TBNivel7!Desenho & "' and Versao_desenho = '" & TBNivel7!Versao_desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBGravar.EOF = True Then TBGravar.AddNew
        TBGravar!Codproduto = TBNivel7!Codproduto
        TBGravar!Posicao = TBNivel7!Posicao
        TBGravar!Desenho = TBNivel7!Desenho
        TBGravar!Versao_desenho = TBNivel7!Versao_desenho
        TBGravar!Descricao = TBNivel7!Descricao
        TBGravar!PesoMetro = TBNivel7!PesoMetro
        TBGravar!PesoTotal = TBNivel7!PesoTotal
        TBGravar!Percentual_perda = TBNivel7!Percentual_perda
        TBGravar!ID_partnumber_fabricante = TBNivel7!ID_partnumber_fabricante
        TBGravar!Obs = TBNivel7!Obs
        TBGravar!quantidade = TBNivel7!quantidade
        TBGravar!Peso = TBNivel7!Peso
        TBGravar!Unidade = TBNivel7!Unidade
        TBGravar!Dimensoes = TBNivel7!Dimensoes
        TBGravar!valor = TBNivel7!valor
        TBGravar!ValorTotal = TBNivel7!ValorTotal
        TBGravar!Un_Kg = TBNivel7!Un_Kg
        If VersaoNova = "" Then TBGravar!versao = TBNivel7!versao Else TBGravar!versao = VersaoNova
        TBGravar.Update
        If VersaoNova <> "" Then ProcCopiarDescVersaoEstrutura TBNivel7!Codproduto, VersaoAnt, TBNivel7!Codproduto, VersaoNova
        TBGravar.Close
        
        ProcCopiarEstruturaN8 TBNivel7!Desenho, TBNivel7!Versao_desenho, VersaoNova
        TBNivel7.MoveNext
    Loop
End If
TBNivel7.Close

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCopiarEstruturaN8(Desenho As String, VersaoAnt As String, VersaoNova As String)
On Error GoTo tratar_erro

Set TBNivel8 = CreateObject("adodb.recordset")
TBNivel8.Open "Select PC.* from projconjunto PC INNER JOIN Projproduto P ON P.Codproduto = PC.Codproduto where P.Desenho = '" & Desenho & "' and PC.Versao = '" & VersaoAnt & "' order by PC.Codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBNivel8.EOF = False Then
    Do While TBNivel8.EOF = False
        If VersaoNova <> "" Then TextoFiltro = " and Versao = '" & VersaoNova & "'" Else TextoFiltro = ""
        Set TBGravar = CreateObject("adodb.recordset")
        TBGravar.Open "Select * from projconjunto where codProduto = " & TBNivel8!Codproduto & TextoFiltro & " and Desenho = '" & TBNivel8!Desenho & "' and Versao_desenho = '" & TBNivel8!Versao_desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBGravar.EOF = True Then TBGravar.AddNew
        TBGravar!Codproduto = TBNivel8!Codproduto
        TBGravar!Posicao = TBNivel8!Posicao
        TBGravar!Desenho = TBNivel8!Desenho
        TBGravar!Versao_desenho = TBNivel8!Versao_desenho
        TBGravar!Descricao = TBNivel8!Descricao
        TBGravar!PesoMetro = TBNivel8!PesoMetro
        TBGravar!PesoTotal = TBNivel8!PesoTotal
        TBGravar!Percentual_perda = TBNivel8!Percentual_perda
        TBGravar!ID_partnumber_fabricante = TBNivel8!ID_partnumber_fabricante
        TBGravar!Obs = TBNivel8!Obs
        TBGravar!quantidade = TBNivel8!quantidade
        TBGravar!Peso = TBNivel8!Peso
        TBGravar!Unidade = TBNivel8!Unidade
        TBGravar!Dimensoes = TBNivel8!Dimensoes
        TBGravar!valor = TBNivel8!valor
        TBGravar!ValorTotal = TBNivel8!ValorTotal
        TBGravar!Un_Kg = TBNivel8!Un_Kg
        If VersaoNova = "" Then TBGravar!versao = TBNivel8!versao Else TBGravar!versao = VersaoNova
        TBGravar.Update
        If VersaoNova <> "" Then ProcCopiarDescVersaoEstrutura TBNivel8!Codproduto, VersaoAnt, TBNivel8!Codproduto, VersaoNova
        TBGravar.Close
        
        ProcCopiarEstruturaN9 TBNivel8!Desenho, TBNivel8!Versao_desenho, VersaoNova
        TBNivel8.MoveNext
    Loop
End If
TBNivel8.Close

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCopiarEstruturaN9(Desenho As String, VersaoAnt As String, VersaoNova As String)
On Error GoTo tratar_erro

Set TBNivel9 = CreateObject("adodb.recordset")
TBNivel9.Open "Select PC.* from projconjunto PC INNER JOIN Projproduto P ON P.Codproduto = PC.Codproduto where P.Desenho = '" & Desenho & "' and PC.Versao = '" & VersaoAnt & "' order by PC.Codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBNivel9.EOF = False Then
    Do While TBNivel9.EOF = False
        If VersaoNova <> "" Then TextoFiltro = " and Versao = '" & VersaoNova & "'" Else TextoFiltro = ""
        Set TBGravar = CreateObject("adodb.recordset")
        TBGravar.Open "Select * from projconjunto where codProduto = " & TBNivel9!Codproduto & TextoFiltro & " and Desenho = '" & TBNivel9!Desenho & "' and Versao_desenho = '" & TBNivel9!Versao_desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBGravar.EOF = True Then TBGravar.AddNew
        TBGravar!Codproduto = TBNivel9!Codproduto
        TBGravar!Posicao = TBNivel9!Posicao
        TBGravar!Desenho = TBNivel9!Desenho
        TBGravar!Versao_desenho = TBNivel9!Versao_desenho
        TBGravar!Descricao = TBNivel9!Descricao
        TBGravar!PesoMetro = TBNivel9!PesoMetro
        TBGravar!PesoTotal = TBNivel9!PesoTotal
        TBGravar!Percentual_perda = TBNivel9!Percentual_perda
        TBGravar!ID_partnumber_fabricante = TBNivel9!ID_partnumber_fabricante
        TBGravar!Obs = TBNivel9!Obs
        TBGravar!quantidade = TBNivel9!quantidade
        TBGravar!Peso = TBNivel9!Peso
        TBGravar!Unidade = TBNivel9!Unidade
        TBGravar!Dimensoes = TBNivel9!Dimensoes
        TBGravar!valor = TBNivel9!valor
        TBGravar!ValorTotal = TBNivel9!ValorTotal
        TBGravar!Un_Kg = TBNivel9!Un_Kg
        If VersaoNova = "" Then TBGravar!versao = TBNivel9!versao Else TBGravar!versao = VersaoNova
        TBGravar.Update
        If VersaoNova <> "" Then ProcCopiarDescVersaoEstrutura TBNivel9!Codproduto, VersaoAnt, TBNivel9!Codproduto, VersaoNova
        TBGravar.Close
        
        ProcCopiarEstruturaN10 TBNivel9!Desenho, TBNivel9!Versao_desenho, VersaoNova
        TBNivel9.MoveNext
    Loop
End If
TBNivel9.Close

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCopiarEstruturaN10(Desenho As String, VersaoAnt As String, VersaoNova As String)
On Error GoTo tratar_erro

Set TBNivel10 = CreateObject("adodb.recordset")
TBNivel10.Open "Select PC.* from projconjunto PC INNER JOIN Projproduto P ON P.Codproduto = PC.Codproduto where P.Desenho = '" & Desenho & "' and PC.Versao = '" & VersaoAnt & "' order by PC.Codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBNivel10.EOF = False Then
    Do While TBNivel10.EOF = False
        If VersaoNova <> "" Then TextoFiltro = " and Versao = '" & VersaoNova & "'" Else TextoFiltro = ""
        Set TBGravar = CreateObject("adodb.recordset")
        TBGravar.Open "Select * from projconjunto where codProduto = " & TBNivel10!Codproduto & TextoFiltro & " and Desenho = '" & TBNivel10!Desenho & "' and Versao_desenho = '" & TBNivel10!Versao_desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBGravar.EOF = True Then TBGravar.AddNew
        TBGravar!Codproduto = TBNivel10!Codproduto
        TBGravar!Posicao = TBNivel10!Posicao
        TBGravar!Desenho = TBNivel10!Desenho
        TBGravar!Versao_desenho = TBNivel10!Versao_desenho
        TBGravar!Descricao = TBNivel10!Descricao
        TBGravar!PesoMetro = TBNivel10!PesoMetro
        TBGravar!PesoTotal = TBNivel10!PesoTotal
        TBGravar!Percentual_perda = TBNivel10!Percentual_perda
        TBGravar!ID_partnumber_fabricante = TBNivel10!ID_partnumber_fabricante
        TBGravar!Obs = TBNivel10!Obs
        TBGravar!quantidade = TBNivel10!quantidade
        TBGravar!Peso = TBNivel10!Peso
        TBGravar!Unidade = TBNivel10!Unidade
        TBGravar!Dimensoes = TBNivel10!Dimensoes
        TBGravar!valor = TBNivel10!valor
        TBGravar!ValorTotal = TBNivel10!ValorTotal
        TBGravar!Un_Kg = TBNivel10!Un_Kg
        If VersaoNova = "" Then TBGravar!versao = TBNivel10!versao Else TBGravar!versao = VersaoNova
        TBGravar.Update
        If VersaoNova <> "" Then ProcCopiarDescVersaoEstrutura TBNivel10!Codproduto, VersaoAnt, TBNivel10!Codproduto, VersaoNova
        TBGravar.Close
        
        ProcCopiarEstruturaN11 TBNivel10!Desenho, TBNivel10!Versao_desenho, VersaoNova
        TBNivel10.MoveNext
    Loop
End If
TBNivel10.Close

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCopiarEstruturaN11(Desenho As String, VersaoAnt As String, VersaoNova As String)
On Error GoTo tratar_erro

Set TBNivel11 = CreateObject("adodb.recordset")
TBNivel11.Open "Select PC.* from projconjunto PC INNER JOIN Projproduto P ON P.Codproduto = PC.Codproduto where P.Desenho = '" & Desenho & "' and PC.Versao = '" & VersaoAnt & "' order by PC.Codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBNivel11.EOF = False Then
    Do While TBNivel11.EOF = False
        If VersaoNova <> "" Then TextoFiltro = " and Versao = '" & VersaoNova & "'" Else TextoFiltro = ""
        Set TBGravar = CreateObject("adodb.recordset")
        TBGravar.Open "Select * from projconjunto where codProduto = " & TBNivel11!Codproduto & TextoFiltro & " and Desenho = '" & TBNivel11!Desenho & "' and Versao_desenho = '" & TBNivel11!Versao_desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBGravar.EOF = True Then TBGravar.AddNew
        TBGravar!Codproduto = TBNivel11!Codproduto
        TBGravar!Posicao = TBNivel11!Posicao
        TBGravar!Desenho = TBNivel11!Desenho
        TBGravar!Versao_desenho = TBNivel11!Versao_desenho
        TBGravar!Descricao = TBNivel11!Descricao
        TBGravar!PesoMetro = TBNivel11!PesoMetro
        TBGravar!PesoTotal = TBNivel11!PesoTotal
        TBGravar!Percentual_perda = TBNivel11!Percentual_perda
        TBGravar!ID_partnumber_fabricante = TBNivel11!ID_partnumber_fabricante
        TBGravar!Obs = TBNivel11!Obs
        TBGravar!quantidade = TBNivel11!quantidade
        TBGravar!Peso = TBNivel11!Peso
        TBGravar!Unidade = TBNivel11!Unidade
        TBGravar!Dimensoes = TBNivel11!Dimensoes
        TBGravar!valor = TBNivel11!valor
        TBGravar!ValorTotal = TBNivel11!ValorTotal
        TBGravar!Un_Kg = TBNivel11!Un_Kg
        If VersaoNova = "" Then TBGravar!versao = TBNivel11!versao Else TBGravar!versao = VersaoNova
        TBGravar.Update
        If VersaoNova <> "" Then ProcCopiarDescVersaoEstrutura TBNivel1!Codproduto, VersaoAnt, TBNivel1!Codproduto, VersaoNova
        TBGravar.Close
        
        ProcCopiarEstruturaN12 TBNivel11!Desenho, TBNivel11!Versao_desenho, VersaoNova
        TBNivel11.MoveNext
    Loop
End If
TBNivel11.Close

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCopiarEstruturaN12(Desenho As String, VersaoAnt As String, VersaoNova As String)
On Error GoTo tratar_erro

Set TBNivel12 = CreateObject("adodb.recordset")
TBNivel12.Open "Select PC.* from projconjunto PC INNER JOIN Projproduto P ON P.Codproduto = PC.Codproduto where P.Desenho = '" & Desenho & "' and PC.Versao = '" & VersaoAnt & "' order by PC.Codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBNivel12.EOF = False Then
    Do While TBNivel12.EOF = False
        If VersaoNova <> "" Then TextoFiltro = " and Versao = '" & VersaoNova & "'" Else TextoFiltro = ""
        Set TBGravar = CreateObject("adodb.recordset")
        TBGravar.Open "Select * from projconjunto where codProduto = " & TBNivel12!Codproduto & TextoFiltro & " and Desenho = '" & TBNivel12!Desenho & "' and Versao_desenho = '" & TBNivel12!Versao_desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBGravar.EOF = True Then TBGravar.AddNew
        TBGravar!Codproduto = TBNivel12!Codproduto
        TBGravar!Posicao = TBNivel12!Posicao
        TBGravar!Desenho = TBNivel12!Desenho
        TBGravar!Versao_desenho = TBNivel12!Versao_desenho
        TBGravar!Descricao = TBNivel12!Descricao
        TBGravar!PesoMetro = TBNivel12!PesoMetro
        TBGravar!PesoTotal = TBNivel12!PesoTotal
        TBGravar!Percentual_perda = TBNivel12!Percentual_perda
        TBGravar!ID_partnumber_fabricante = TBNivel12!ID_partnumber_fabricante
        TBGravar!Obs = TBNivel12!Obs
        TBGravar!quantidade = TBNivel12!quantidade
        TBGravar!Peso = TBNivel12!Peso
        TBGravar!Unidade = TBNivel12!Unidade
        TBGravar!Dimensoes = TBNivel12!Dimensoes
        TBGravar!valor = TBNivel12!valor
        TBGravar!ValorTotal = TBNivel12!ValorTotal
        TBGravar!Un_Kg = TBNivel12!Un_Kg
        If VersaoNova = "" Then TBGravar!versao = TBNivel12!versao Else TBGravar!versao = VersaoNova
        TBGravar.Update
        If VersaoNova <> "" Then ProcCopiarDescVersaoEstrutura TBNivel12!Codproduto, VersaoAnt, TBNivel12!Codproduto, VersaoNova
        TBGravar.Close
        
        ProcCopiarEstruturaN13 TBNivel12!Desenho, TBNivel12!Versao_desenho, VersaoNova
        TBNivel12.MoveNext
    Loop
End If
TBNivel12.Close

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCopiarEstruturaN13(Desenho As String, VersaoAnt As String, VersaoNova As String)
On Error GoTo tratar_erro

Set TBNivel13 = CreateObject("adodb.recordset")
TBNivel13.Open "Select PC.* from projconjunto PC INNER JOIN Projproduto P ON P.Codproduto = PC.Codproduto where P.Desenho = '" & Desenho & "' and PC.Versao = '" & VersaoAnt & "' order by PC.Codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBNivel13.EOF = False Then
    Do While TBNivel13.EOF = False
        If VersaoNova <> "" Then TextoFiltro = " and Versao = '" & VersaoNova & "'" Else TextoFiltro = ""
        Set TBGravar = CreateObject("adodb.recordset")
        TBGravar.Open "Select * from projconjunto where codProduto = " & TBNivel13!Codproduto & TextoFiltro & " and Desenho = '" & TBNivel13!Desenho & "' and Versao_desenho = '" & TBNivel13!Versao_desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBGravar.EOF = True Then TBGravar.AddNew
        TBGravar!Codproduto = TBNivel13!Codproduto
        TBGravar!Posicao = TBNivel13!Posicao
        TBGravar!Desenho = TBNivel13!Desenho
        TBGravar!Versao_desenho = TBNivel13!Versao_desenho
        TBGravar!Descricao = TBNivel13!Descricao
        TBGravar!PesoMetro = TBNivel13!PesoMetro
        TBGravar!PesoTotal = TBNivel13!PesoTotal
        TBGravar!Percentual_perda = TBNivel13!Percentual_perda
        TBGravar!ID_partnumber_fabricante = TBNivel13!ID_partnumber_fabricante
        TBGravar!Obs = TBNivel13!Obs
        TBGravar!quantidade = TBNivel13!quantidade
        TBGravar!Peso = TBNivel13!Peso
        TBGravar!Unidade = TBNivel13!Unidade
        TBGravar!Dimensoes = TBNivel13!Dimensoes
        TBGravar!valor = TBNivel13!valor
        TBGravar!ValorTotal = TBNivel13!ValorTotal
        TBGravar!Un_Kg = TBNivel13!Un_Kg
        If VersaoNova = "" Then TBGravar!versao = TBNivel13!versao Else TBGravar!versao = VersaoNova
        TBGravar.Update
        If VersaoNova <> "" Then ProcCopiarDescVersaoEstrutura TBNivel13!Codproduto, VersaoAnt, TBNivel13!Codproduto, VersaoNova
        TBGravar.Close
        
        ProcCopiarEstruturaN14 TBNivel13!Desenho, TBNivel13!Versao_desenho, VersaoNova
        TBNivel13.MoveNext
    Loop
End If
TBNivel13.Close

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCopiarEstruturaN14(Desenho As String, VersaoAnt As String, VersaoNova As String)
On Error GoTo tratar_erro

Set TBNivel14 = CreateObject("adodb.recordset")
TBNivel14.Open "Select PC.* from projconjunto PC INNER JOIN Projproduto P ON P.Codproduto = PC.Codproduto where P.Desenho = '" & Desenho & "' and PC.Versao = '" & VersaoAnt & "' order by PC.Codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBNivel14.EOF = False Then
    Do While TBNivel14.EOF = False
        If VersaoNova <> "" Then TextoFiltro = " and Versao = '" & VersaoNova & "'" Else TextoFiltro = ""
        Set TBGravar = CreateObject("adodb.recordset")
        TBGravar.Open "Select * from projconjunto where codProduto = " & TBNivel14!Codproduto & TextoFiltro & " and Desenho = '" & TBNivel14!Desenho & "' and Versao_desenho = '" & TBNivel14!Versao_desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBGravar.EOF = True Then TBGravar.AddNew
        TBGravar!Codproduto = TBNivel14!Codproduto
        TBGravar!Posicao = TBNivel14!Posicao
        TBGravar!Desenho = TBNivel14!Desenho
        TBGravar!Versao_desenho = TBNivel14!Versao_desenho
        TBGravar!Descricao = TBNivel14!Descricao
        TBGravar!PesoMetro = TBNivel14!PesoMetro
        TBGravar!PesoTotal = TBNivel14!PesoTotal
        TBGravar!Percentual_perda = TBNivel14!Percentual_perda
        TBGravar!ID_partnumber_fabricante = TBNivel11!ID_partnumber_fabricante
        TBGravar!Obs = TBNivel11!Obs
        TBGravar!quantidade = TBNivel14!quantidade
        TBGravar!Peso = TBNivel14!Peso
        TBGravar!Unidade = TBNivel14!Unidade
        TBGravar!Dimensoes = TBNivel14!Dimensoes
        TBGravar!valor = TBNivel14!valor
        TBGravar!ValorTotal = TBNivel14!ValorTotal
        TBGravar!Un_Kg = TBNivel14!Un_Kg
        If VersaoNova = "" Then TBGravar!versao = TBNivel14!versao Else TBGravar!versao = VersaoNova
        TBGravar.Update
        If VersaoNova <> "" Then ProcCopiarDescVersaoEstrutura TBNivel14!Codproduto, VersaoAnt, TBNivel14!Codproduto, VersaoNova
        TBGravar.Close
        
        ProcCopiarEstruturaN15 TBNivel14!Desenho, TBNivel14!Versao_desenho, VersaoNova
        TBNivel14.MoveNext
    Loop
End If
TBNivel14.Close

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCopiarEstruturaN15(Desenho As String, VersaoAnt As String, VersaoNova As String)
On Error GoTo tratar_erro

Set TBNivel15 = CreateObject("adodb.recordset")
TBNivel15.Open "Select PC.* from projconjunto PC INNER JOIN Projproduto P ON P.Codproduto = PC.Codproduto where P.Desenho = '" & Desenho & "' and PC.Versao = '" & VersaoAnt & "' order by PC.Codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBNivel15.EOF = False Then
    Do While TBNivel15.EOF = False
        If VersaoNova <> "" Then TextoFiltro = " and Versao = '" & VersaoNova & "'" Else TextoFiltro = ""
        Set TBGravar = CreateObject("adodb.recordset")
        TBGravar.Open "Select * from projconjunto where codProduto = " & TBNivel15!Codproduto & TextoFiltro & " and Desenho = '" & TBNivel15!Desenho & "' and Versao_desenho = '" & TBNivel15!Versao_desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBGravar.EOF = True Then TBGravar.AddNew
        TBGravar!Codproduto = TBNivel15!Codproduto
        TBGravar!Posicao = TBNivel15!Posicao
        TBGravar!Desenho = TBNivel15!Desenho
        TBGravar!Versao_desenho = TBNivel15!Versao_desenho
        TBGravar!Descricao = TBNivel15!Descricao
        TBGravar!PesoMetro = TBNivel15!PesoMetro
        TBGravar!PesoTotal = TBNivel15!PesoTotal
        TBGravar!Percentual_perda = TBNivel15!Percentual_perda
        TBGravar!ID_partnumber_fabricante = TBNivel15!ID_partnumber_fabricante
        TBGravar!Obs = TBNivel15!Obs
        TBGravar!quantidade = TBNivel15!quantidade
        TBGravar!Peso = TBNivel15!Peso
        TBGravar!Unidade = TBNivel15!Unidade
        TBGravar!Dimensoes = TBNivel15!Dimensoes
        TBGravar!valor = TBNivel15!valor
        TBGravar!ValorTotal = TBNivel15!ValorTotal
        TBGravar!Un_Kg = TBNivel15!Un_Kg
        If VersaoNova = "" Then TBGravar!versao = TBNivel15!versao Else TBGravar!versao = VersaoNova
        TBGravar.Update
        If VersaoNova <> "" Then ProcCopiarDescVersaoEstrutura TBNivel15!Codproduto, VersaoAnt, TBNivel15!Codproduto, VersaoNova
        TBGravar.Close
        TBNivel15.MoveNext
    Loop
End If
TBNivel15.Close

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Function FunVerifExisteNumNF(TipoNF As String, ID_empresa As Integer, NumNF As String, SerieNF As String, Modelo As String) As String
On Error GoTo tratar_erro

Set TBTempo = CreateObject("adodb.recordset")
TBTempo.Open "Select CAST(int_NotaFiscal AS int) AS NF FROM tbl_Dados_Nota_Fiscal where tipoNF = '" & TipoNF & "' and Aplicacao = 'P' and modelo = '" & Modelo & "' and ID_empresa = " & ID_empresa & " and int_NotaFiscal = '" & NumNF & "' and Serie = '" & SerieNF & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBTempo.EOF = False Then
    QuantsolicitadoN1 = TBTempo!NF + 1
    FamiliaAntiga = QuantsolicitadoN1
    Familiatext = FunTamanhoTextoZeroEsq(FamiliaAntiga, 9)
    
    Familiatext = FunVerifExisteNumNF(TipoNF, ID_empresa, Familiatext, SerieNF, Modelo)
End If

If Familiatext <> "" Then
FunVerifExisteNumNF = Familiatext
Else
Exit Function
End If

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Function FunVerificaCFOPUF(IDCFOP As Long, UF As String) As Boolean
On Error GoTo tratar_erro

FunVerificaCFOPUF = True
Permitido = False
Set TBFI = CreateObject("adodb.recordset")
TBFI.Open "Select * from tbl_NaturezaOperacao where IDCountCfop = " & IDCFOP, Conexao, adOpenKeyset, adLockOptimistic
If TBFI.EOF = False Then
    If TBFI!De = True Then
        FiltroRegiao = "Regiao = 'DE'"
        Mensagem = "fora do estado"
    ElseIf TBFI!FE = True Then
            FiltroRegiao = "Regiao <> 'DE'"
            Mensagem = "dentro do estado"
        Else
            Permitido = True
    End If
End If

If Permitido = False Then
    Set TBFI = CreateObject("adodb.recordset")
    TBFI.Open "Select * from regioes where " & FiltroRegiao & " and UF = '" & UF & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBFI.EOF = True Then
        USMsgBox ("SÛ È permitido utilizar CFOP de " & Mensagem & "."), vbExclamation, "CAPRIND v5.0"
        TBFI.Close
        FunVerificaCFOPUF = False
        Exit Function
    End If
    TBFI.Close
End If

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Sub ProcVerificaContasSelRel(Lista As ListView, Permitido As Boolean)
On Error GoTo tratar_erro

Familiatext = ""
If Permitido = True Then
    With Lista
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                If Formulario = "Financeiro/Contas a pagar" Or Formulario = "Financeiro/Contas pagas" Then NomeTabelaTexto = "tbl_ContasPagar" Else NomeTabelaTexto = "tbl_contas_receber"
                If Familiatext = "" Then Familiatext = "{" & NomeTabelaTexto & ".IdIntConta} = " & .ListItems(InitFor) Else Familiatext = Familiatext & " or {" & NomeTabelaTexto & ".IdIntConta} = " & .ListItems(InitFor)
            End If
        Next InitFor
    End With
End If

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Function FunVerificaIDProcesso(ByVal appName As String) As Boolean
On Error GoTo tratar_erro
Dim Process As PROCESSENTRY32
Dim hSnapShot As Long
Dim r As Long
    
appName = LCase$(appName)
hSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0&)
If hSnapShot <> -1 Then
    Process.dwSize = Len(Process)
    r = Process32First(hSnapShot, Process)
    Do While r
        If LCase$(Left$(Process.szExeFile, InStr(1, Process.szExeFile, vbNullChar) - 1)) = appName Then
            ProcessId = Process.th32ProcessID
            FunVerificaIDProcesso = True
            r = False
        End If
        r = Process32Next(hSnapShot, Process)
    Loop
    CloseHandle hSnapShot
End If

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Public Sub ProcEncerrarProcesso()
On Error GoTo tratar_erro

ThreadId = GetWindowThreadProcessId(hWnd, ProcessId)
ProcessHandle = OpenProcess(PROCESS_ALL_ACCESS, False, ProcessId)
TerminateProcess ProcessHandle, 0
CloseHandle ProcessHandle

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Function FunVerifFormAberto(frm As Form) As Boolean
On Error GoTo tratar_erro
Dim frmFind As Form

FunVerifFormAberto = False
For Each frmFind In Forms
    If frmFind Is frm Then
        FunVerifFormAberto = True
        Exit For
    End If
Next

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Function FunVerifStatusNFe(ID_nota As Long) As String
On Error GoTo tratar_erro

FunVerifStatusNFe = ""
Set TBTipo = CreateObject("adodb.recordset")
TBTipo.Open "Select NF.TipoNF, NFE.Status from tbl_dados_nota_fiscal_NFe NFE INNER JOIN tbl_dados_nota_fiscal NF ON NF.ID = NFE.ID_nota where NFE.ID_nota = " & ID_nota, Conexao, adOpenKeyset, adLockReadOnly
If TBTipo.EOF = False Then
    If TBTipo!TipoNF <> "SA" Then
        Select Case TBTipo!status
            Case "1": FunVerifStatusNFe = "Enviado para SEFAZ"
            Case "2": FunVerifStatusNFe = "Nota fiscal rejeitada"
            Case "100": FunVerifStatusNFe = "Autorizado o uso da NF-e"
            Case "100": FunVerifStatusNFe = "Autorizado o uso da NF-e"
            Case "200": FunVerifStatusNFe = "Autorizado o uso da NF-e"
            Case "101": FunVerifStatusNFe = "Cancelamento de NF-e homologado"
            Case "102": FunVerifStatusNFe = "InutilizaÁ„o de n˙mero homologado"
            Case "103": FunVerifStatusNFe = "Lote recebido com sucesso"
            Case "104": FunVerifStatusNFe = "Lote processado"
            Case "105": FunVerifStatusNFe = "Lote em processamento"
            Case "106": FunVerifStatusNFe = "Lote n„o localizado"
            Case "107": FunVerifStatusNFe = "ServiÁo em OperaÁ„o"
            Case "108": FunVerifStatusNFe = "ServiÁo Paralisado Momentaneamente (curto prazo)"
            Case "109": FunVerifStatusNFe = "ServiÁo Paralisado sem Previs„o"
            Case "110": FunVerifStatusNFe = "uso Denegado"
            Case "111": FunVerifStatusNFe = "Consulta cadastro com uma ocorrÍncia"
            Case "112": FunVerifStatusNFe = "Consulta cadastro com mais de uma ocorrÍncia"
            Case "201": FunVerifStatusNFe = "RejeiÁ„o: O numero m·ximo de numeraÁ„o de NF-e a inutilizar ultrapassou o limite"
            Case "202": FunVerifStatusNFe = "RejeiÁ„o: Falha no reconhecimento da autoria ou integridade do arquivo digital"
            Case "203": FunVerifStatusNFe = "RejeiÁ„o: Emissor n„o habilitado para emiss„o da NF-e"
            Case "204": FunVerifStatusNFe = "RejeiÁ„o: Duplicidade de NF-e"
            Case "205": FunVerifStatusNFe = "RejeiÁ„o: NF-e est· denegada na base de dados da SEFAZ"
            Case "206": FunVerifStatusNFe = "RejeiÁ„o: NF-e j· est· inutilizada na Base de dados da SEFAZ"
            Case "207": FunVerifStatusNFe = "RejeiÁ„o: CNPJ do emitente inv·lido"
            Case "208": FunVerifStatusNFe = "RejeiÁ„o: CNPJ do destinat·rio inv·lido"
            Case "209": FunVerifStatusNFe = "RejeiÁ„o: IE do emitente inv·lida"
            Case "210": FunVerifStatusNFe = "RejeiÁ„o: IE do destinat·rio inv·lida"
            Case "211": FunVerifStatusNFe = "RejeiÁ„o: IE do substituto inv·lida"
            Case "212": FunVerifStatusNFe = "RejeiÁ„o: Data de emiss„o NF-e posterior a data de recebimento"
            Case "212": FunVerifStatusNFe = "RejeiÁ„o: Data de emiss„o NF-e posterior a data de recebimento"
            Case "213": FunVerifStatusNFe = "RejeiÁ„o: CNPJ-Base do Emitente difere do CNPJ-Base do Certificado Digital"
            Case "214": FunVerifStatusNFe = "RejeiÁ„o: Tamanho da mensagem excedeu o limite estabelecido"
            Case "215": FunVerifStatusNFe = "RejeiÁ„o: Falha no schema XML"
            Case "216": FunVerifStatusNFe = "RejeiÁ„o: Chave de Acesso difere da cadastrada"
            Case "217": FunVerifStatusNFe = "RejeiÁ„o: NF-e n„o consta na base de dados da SEFAZ"
            Case "218": FunVerifStatusNFe = "RejeiÁ„o: NF-e j· esta cancelada na base de dados da SEFAZ"
            Case "219": FunVerifStatusNFe = "RejeiÁ„o: CirculaÁ„o da NF-e verificada"
            Case "220": FunVerifStatusNFe = "RejeiÁ„o: NF-e autorizada h· mais de 7 dias (168 horas)"
            Case "221": FunVerifStatusNFe = "RejeiÁ„o: Confirmado o recebimento da NF-e pelo destinat·rio"
            Case "222": FunVerifStatusNFe = "RejeiÁ„o: Protocolo de AutorizaÁ„o de Uso difere do cadastrado"
            Case "223": FunVerifStatusNFe = "RejeiÁ„o: CNPJ do transmissor do lote difere do CNPJ do transmissor da consulta"
            Case "224": FunVerifStatusNFe = "RejeiÁ„o: A faixa inicial È maior que a faixa final"
            Case "225": FunVerifStatusNFe = "RejeiÁ„o: Falha no Schema XML da NFe"
            Case "226": FunVerifStatusNFe = "RejeiÁ„o: CÛdigo da UF do Emitente diverge da UF autorizadora"
            Case "227": FunVerifStatusNFe = "RejeiÁ„o: Erro na Chave de Acesso - Campo ID"
            Case "228": FunVerifStatusNFe = "RejeiÁ„o: Data de Emiss„o muito atrasada"
            Case "229": FunVerifStatusNFe = "RejeiÁ„o: IE do emitente n„o informada"
            Case "230": FunVerifStatusNFe = "RejeiÁ„o: IE do emitente n„o cadastrada"
            Case "231": FunVerifStatusNFe = "RejeiÁ„o: IE do emitente n„o vinculada ao CNPJ"
            Case "232": FunVerifStatusNFe = "RejeiÁ„o: IE do destinat·rio n„o informada"
            Case "233": FunVerifStatusNFe = "RejeiÁ„o: IE do destinat·rio n„o cadastrada"
            Case "234": FunVerifStatusNFe = "RejeiÁ„o: IE do destinat·rio n„o vinculada ao CNPJ"
            Case "235": FunVerifStatusNFe = "RejeiÁ„o: InscriÁ„o SUFRAMA inv·lida"
            Case "236": FunVerifStatusNFe = "RejeiÁ„o: Chave de Acesso com dÌgito verificador inv·lido"
            Case "237": FunVerifStatusNFe = "RejeiÁ„o: CPF do destinat·rio inv·lido"
            Case "238": FunVerifStatusNFe = "RejeiÁ„o: CabeÁalho - Vers„o do arquivo XML superior a Vers„o vigente"
            Case "239": FunVerifStatusNFe = "RejeiÁ„o: CabeÁalho - Vers„o do arquivo XML n„o suportada"
            Case "240": FunVerifStatusNFe = "RejeiÁ„o: Cancelamento/InutilizaÁ„o - Irregularidade Fiscal do Emitente"
            Case "241": FunVerifStatusNFe = "RejeiÁ„o: Um n˙mero da faixa j· foi utilizado"
            Case "242": FunVerifStatusNFe = "RejeiÁ„o: CabeÁalho - Falha no Schema XML"
            Case "243": FunVerifStatusNFe = "RejeiÁ„o: XML Mal Formado"
            Case "244": FunVerifStatusNFe = "RejeiÁ„o: CNPJ do Certificado Digital difere do CNPJ da Matriz e do CNPJ do Emitente"
            Case "245": FunVerifStatusNFe = "RejeiÁ„o: CNPJ Emitente n„o cadastrado"
            Case "246": FunVerifStatusNFe = "RejeiÁ„o: CNPJ Destinat·rio n„o cadastrado"
            Case "247": FunVerifStatusNFe = "RejeiÁ„o: Sigla da UF do Emitente diverge da UF autorizadora"
            Case "248": FunVerifStatusNFe = "RejeiÁ„o: UF do Recibo diverge da UF autorizadora"
            Case "249": FunVerifStatusNFe = "RejeiÁ„o: UF da Chave de Acesso diverge da UF autorizadora"
            Case "250": FunVerifStatusNFe = "RejeiÁ„o: UF diverge da UF autorizadora"
            Case "251": FunVerifStatusNFe = "RejeiÁ„o: UF/MunicÌpio destinat·rio n„o pertence a SUFRAMA"
            Case "252": FunVerifStatusNFe = "RejeiÁ„o: Ambiente informado diverge do Ambiente de recebimento"
            Case "253": FunVerifStatusNFe = "RejeiÁ„o: Digito Verificador da chave de acesso composta inv·lida"
            Case "254": FunVerifStatusNFe = "RejeiÁ„o: NF-e referenciada n„o informada para NF-e complementar"
            Case "255": FunVerifStatusNFe = "RejeiÁ„o: Informada mais de uma NF-e referenciada para NF-e complementar"
            Case "256": FunVerifStatusNFe = "RejeiÁ„o: Uma NF-e da faixa j· est· inutilizada na Base de dados da SEFAZ"
            Case "257": FunVerifStatusNFe = "RejeiÁ„o: Solicitante n„o habilitado para emiss„o da NF-e"
            Case "258": FunVerifStatusNFe = "RejeiÁ„o: CNPJ da consulta inv·lido"
            Case "259": FunVerifStatusNFe = "RejeiÁ„o: CNPJ da consulta n„o cadastrado como contribuinte na UF"
            Case "260": FunVerifStatusNFe = "RejeiÁ„o: IE da consulta inv·lida"
            Case "261": FunVerifStatusNFe = "RejeiÁ„o: IE da consulta n„o cadastrada como contribuinte na UF"
            Case "262": FunVerifStatusNFe = "RejeiÁ„o: UF n„o fornece consulta por CPF"
            Case "263": FunVerifStatusNFe = "RejeiÁ„o: CPF da consulta inv·lido"
            Case "264": FunVerifStatusNFe = "RejeiÁ„o: CPF da consulta n„o cadastrado como contribuinte na UF"
            Case "265": FunVerifStatusNFe = "RejeiÁ„o: Sigla da UF da consulta difere da UF do Web Service"
            Case "266": FunVerifStatusNFe = "RejeiÁ„o: SÈrie utilizada n„o permitida no Web Service"
            Case "267": FunVerifStatusNFe = "RejeiÁ„o: NF Complementar referencia uma NF-e inexistente"
            Case "268": FunVerifStatusNFe = "RejeiÁ„o: NF Complementar referencia uma outra NF-e Complementar"
            Case "269": FunVerifStatusNFe = "RejeiÁ„o: CNPJ Emitente da NF Complementar difere do CNPJ da NF Referenciada"
            Case "270": FunVerifStatusNFe = "RejeiÁ„o: CÛdigo MunicÌpio do Fato Gerador: dÌgito inv·lido"
            Case "271": FunVerifStatusNFe = "RejeiÁ„o: CÛdigo MunicÌpio do Fato Gerador: difere da UF do emitente"
            Case "272": FunVerifStatusNFe = "RejeiÁ„o: CÛdigo MunicÌpio do Emitente: dÌgito inv·lido"
            Case "273": FunVerifStatusNFe = "RejeiÁ„o: CÛdigo MunicÌpio do Emitente: difere da UF do emitente"
            Case "274": FunVerifStatusNFe = "RejeiÁ„o: CÛdigo MunicÌpio do Destinat·rio: dÌgito inv·lido"
            Case "275": FunVerifStatusNFe = "RejeiÁ„o: CÛdigo MunicÌpio do Destinat·rio: difere da UF do Destinat·rio"
            Case "276": FunVerifStatusNFe = "RejeiÁ„o: CÛdigo MunicÌpio do Local de Retirada: dÌgito inv·lido"
            Case "277": FunVerifStatusNFe = "RejeiÁ„o: CÛdigo MunicÌpio do Local de Retirada: difere da UF do Local de Retirada"
            Case "278": FunVerifStatusNFe = "RejeiÁ„o: CÛdigo MunicÌpio do Local de Entrega: dÌgito inv·lido"
            Case "279": FunVerifStatusNFe = "RejeiÁ„o: CÛdigo MunicÌpio do Local de Entrega: difere da UF do Local de Entrega"
            Case "280": FunVerifStatusNFe = "RejeiÁ„o: Certificado Transmissor inv·lido"
            Case "281": FunVerifStatusNFe = "RejeiÁ„o: Certificado Transmissor Data Validade"
            Case "282": FunVerifStatusNFe = "RejeiÁ„o: Certificado Transmissor sem CNPJ"
            Case "283": FunVerifStatusNFe = "RejeiÁ„o: Certificado Transmissor - erro Cadeia de CertificaÁ„o"
            Case "284": FunVerifStatusNFe = "RejeiÁ„o: Certificado Transmissor revogado"
            Case "285": FunVerifStatusNFe = "RejeiÁ„o: Certificado Transmissor difere ICP-Brasil"
            Case "286": FunVerifStatusNFe = "RejeiÁ„o: Certificado Transmissor erro no acesso a LCR"
            Case "287": FunVerifStatusNFe = "RejeiÁ„o: CÛdigo MunicÌpio do FG - ISSQN: dÌgito inv·lido"
            Case "288": FunVerifStatusNFe = "RejeiÁ„o: CÛdigo MunicÌpio do FG - Transporte: dÌgito inv·lido"
            Case "289": FunVerifStatusNFe = "RejeiÁ„o: CÛdigo da UF informada diverge da UF solicitada"
            Case "290": FunVerifStatusNFe = "RejeiÁ„o: Certificado Assinatura inv·lido"
            Case "291": FunVerifStatusNFe = "RejeiÁ„o: Certificado Assinatura Data Validade"
            Case "292": FunVerifStatusNFe = "RejeiÁ„o: Certificado Assinatura sem CNPJ"
            Case "293": FunVerifStatusNFe = "RejeiÁ„o: Certificado Assinatura - erro Cadeia de CertificaÁ„o"
            Case "294": FunVerifStatusNFe = "RejeiÁ„o: Certificado Assinatura revogado"
            Case "295": FunVerifStatusNFe = "RejeiÁ„o: Certificado Assinatura difere ICP-Brasil"
            Case "296": FunVerifStatusNFe = "RejeiÁ„o: Certificado Assinatura erro no acesso a LCR"
            Case "297": FunVerifStatusNFe = "RejeiÁ„o: Assinatura difere do calculado"
            Case "298": FunVerifStatusNFe = "RejeiÁ„o: Assinatura difere do padr„o do Projeto"
            Case "299": FunVerifStatusNFe = "RejeiÁ„o: XML da ·rea de cabeÁalho com codificaÁ„o diferente de UTF-8"
            Case "301": FunVerifStatusNFe = "Uso Denegado : Irregularidade fiscal do emitente"
            Case "302": FunVerifStatusNFe = "Uso Denegado : Irregularidade fiscal do destinat·rio"
            Case "401": FunVerifStatusNFe = "RejeiÁ„o: CPF do remetente inv·lido"
            Case Is >= "402": FunVerifStatusNFe = "RejeiÁ„o"
        End Select
    Else
        Select Case TBTipo!status
            Case "": FunVerifStatusNFe = ""
            Case 0: FunVerifStatusNFe = "Autorizada"
            Case "1": FunVerifStatusNFe = "Lote em processamento"
            Case "2": FunVerifStatusNFe = "RejeiÁ„o de envio"
            Case "3": FunVerifStatusNFe = "Cancelada"
        End Select
    End If
End If
TBTipo.Close

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Function

Sub ProcVerificaTPNFe()
On Error GoTo tratar_erro

If IDempresa <> 0 Then
'Regime = 0
Set TBFIltro = CreateObject("adodb.recordset")
TBFIltro.Open "Select * from Empresa where Codigo = " & IDempresa, Conexao, adOpenKeyset, adLockOptimistic
If TBFIltro.EOF = False Then
    tpAmb = IIf(IsNull(TBFIltro!tpAmb) = False, TBFIltro!tpAmb, "2")
    TPCertificado = IIf(IsNull(TBFIltro!TPCertificado) = False, TBFIltro!TPCertificado, "A1")
    SerialCertificado = IIf(IsNull(TBFIltro!Certificadodigital) = False, TBFIltro!Certificadodigital, "Sem certificado")
    chCNPJ = ReturnNumbersOnly(TBFIltro!CNPJ)
    NF_Serie = IIf(IsNull(TBFIltro!NF_Serie), 0, TBFIltro!NF_Serie)
    ID_empresa = IIf(IsNull(TBFIltro!CODIGO), 0, TBFIltro!CODIGO)
    NomeEmpresa = IIf(IsNull(TBFIltro!Empresa), 0, TBFIltro!Empresa)
    UF_Empresa = IIf(IsNull(TBFIltro!UF), 0, TBFIltro!UF)
End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcVerificaRegime()
On Error GoTo tratar_erro

If IDempresa <> 0 Then
'Regime = 0
Set TBFIltro = CreateObject("adodb.recordset")
TBFIltro.Open "Select * from Empresa where Codigo = " & IDempresa, Conexao, adOpenKeyset, adLockOptimistic
If TBFIltro.EOF = False Then

    If TBFIltro!Simples = True Then
        RegimeEmpresa = 1
        
        Set TBRegime = CreateObject("adodb.recordset")
        TBRegime.Open "Select * from Impostos_TabelaDAS where ID_empresa = " & IDempresa, Conexao, adOpenKeyset, adLockOptimistic
        If TBRegime.EOF = False Then
        TabelaSN = TBRegime!Tabela
        End If
        TBRegime.Close
        
   ' If ReturnNumbersOnly(TBFIltro!CNPJ) = "09382448000191" Then
   '     AliquotaSN = TBFIltro!AliquotaSN
   ' End If
    End If
    
    If TBFIltro!Presumido = True Then RegimeEmpresa = 2
    If TBFIltro!Real = True Then RegimeEmpresa = 3
    If TBFIltro!Simples1 = True Then RegimeEmpresa = 4
End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcAbrirArquivo(caminho As String)
On Error GoTo tratar_erro

Call ShellExecute(0&, vbNullString, caminho, vbNullString, vbNullString, SW_SHOWDEFAULT)

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCriarRMOrdemPC(IDpedido As Long, IDempresa As Long)
On Error GoTo tratar_erro

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Empresa where Codigo = " & IDempresa & " and Gerar_RM_ordem_PC = 'True'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    Conexao.Execute "INSERT INTO Producaomaterial (Ordem, Codigo, Descricao, Requisitado, Unidade, Versao, Saida) Select Ordem, Desenho, Descricao, Quant_Comp, UN, 'A', 'N√O' from Compras_pedido_lista where IDPedido = " & IDpedido & " and Ordem IS NOT NULL and Ordem <> 0 and (Status_Item = 'APROVADO' or Status_Item = 'N_RECEBIDO')"
End If

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcExcluirRMOrdemPC(IDpedido As Long, IDempresa As Long)
On Error GoTo tratar_erro

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Empresa where Codigo = " & IDempresa & " and Gerar_RM_ordem_PC = 'True'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    Conexao.Execute "DELETE from PM from Producaomaterial PM INNER JOIN Compras_pedido_lista CPL ON CPL.Desenho = PM.Codigo and CPL.Ordem = PM.Ordem where CPL.IDpedido = " & IDpedido & " and CPL.Ordem IS NOT NULL and CPL.Ordem <> 0 and (CPL.Status_Item = 'APROVADO' or CPL.Status_Item = 'N_RECEBIDO')"
End If

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcDefinirPrazosOS(OF As Long, PrazoOP As Date, AltPrazofinal As Boolean)
On Error GoTo tratar_erro

If AltPrazofinal = True Then TextoFiltro = "IDproducao = " & OF Else TextoFiltro = "Ordem = " & OF
Set TBProducaoFases = CreateObject("adodb.recordset")
TBProducaoFases.Open "Select IDproducao, IDFase, Quantidade, prazofinal, prazofinalinicio, maquina, Custos from Ordemservico where " & TextoFiltro & " order by Idproducao desc", Conexao, adOpenKeyset, adLockOptimistic
If TBProducaoFases.EOF = False Then
    Do While TBProducaoFases.EOF = False
        Set TBFases = CreateObject("adodb.recordset")
        TBFases.Open "Select TESegundos, TPSegundos from Fases where IDFase = " & TBProducaoFases!IDFase, Conexao, adOpenKeyset, adLockOptimistic
        If TBFases.EOF = False Then
            Valor1 = (IIf(IsNull(TBFases!TESegundos), 0, TBFases!TESegundos) * TBProducaoFases!quantidade) + IIf(IsNull(TBFases!TPSegundos), 0, TBFases!TPSegundos)

            If AltPrazofinal = True Then
                PrazoOP = TBProducaoFases!PrazoFinal + Qtde
             
VoltaInicio:
                Permitido2 = False
                Select Case Weekday(PrazoOP)
                    Case 1
                        If FunVerificaFeriado(PrazoOP) = False Then ProcTotalCadMaqTurnos "Domingo", True Else Permitido2 = True
                    Case 2
                        If FunVerificaFeriado(PrazoOP) = False Then ProcTotalCadMaqTurnos "Segunda", True Else Permitido2 = True
                    Case 3
                        If FunVerificaFeriado(PrazoOP) = False Then ProcTotalCadMaqTurnos "TerÁa", True Else Permitido2 = True
                    Case 4
                        If FunVerificaFeriado(PrazoOP) = False Then ProcTotalCadMaqTurnos "Quarta", True Else Permitido2 = True
                    Case 5
                        If FunVerificaFeriado(PrazoOP) = False Then ProcTotalCadMaqTurnos "Quinta", True Else Permitido2 = True
                    Case 6
                        If FunVerificaFeriado(PrazoOP) = False Then ProcTotalCadMaqTurnos "Sexta", True Else Permitido2 = True
                    Case 7
                        If FunVerificaFeriado(PrazoOP) = False Then ProcTotalCadMaqTurnos "S·bado", True Else Permitido2 = True
                End Select
                
                If Permitido2 = True And Qtde > 0 Then
                    PrazoOP = PrazoOP + 1
                    Qtde = Qtde - 1
                    GoTo VoltaInicio
                End If
            End If
            
            Inicio = PrazoOP
            Qtd = 0
'            If IsNull(TBFases!Leadtime) = False And TBFases!Leadtime <> "" And TBProducaoFases!custos = False Then
'                'Essa rotina se usa quando for fase de terceiros
'                Qtd = TBFases!Leadtime
'                qtdeliberada = Qtd
'                Do While qtdeliberada > 0
'                    Inicio = Inicio - 1
'                    If Weekday(Inicio) = 1 Or Weekday(Inicio) = 7 Then
'                        Qtd = Qtd + 1
'                    Else
'                        If FunVerificaFeriado(Inicio) = True Then Qtd = Qtd + 1
'                    End If
'                    qtdeliberada = qtdeliberada - 1
'                Loop
'                Inicio = PrazoOP - Qtd
'
'Inicio1:
'                If Weekday(Inicio) = 1 Then
'                    Qtd = Qtd + 2
'                ElseIf Weekday(Inicio) = 7 Then
'                        Qtd = Qtd + 1
'                    ElseIf FunVerificaFeriado(Inicio) = True Then
'                            Inicio = Inicio + 1
'                            Qtd = Qtd + 1
'                            GoTo Inicio1
'                End If
'            Else
                'verifica se tem carga maquina, se n„o tiver n„o precisa calcular
                Set TBNivel13 = CreateObject("adodb.recordset")
                TBNivel13.Open "Select Codigo from cadmaqturnos where maquina = '" & TBProducaoFases!maquina & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBNivel13.EOF = False Then
                    'Encontrou = False 'se usa quando o prazo final È o mesmo do dia do computador
                    Do While Valor1 > 0
                        Permitido2 = False
                        If Inicio = Date Then
                           'Encontrou = True
                            'Se o prazo final for igual ao dia do computador ele verifica se È feriado, se for ele aumenta um dia
                            If FunVerificaFeriado(Inicio) = True Then
                                Qtd = Qtd - 1
                                Inicio = Inicio + 1
                            Else
                                GoTo Pula
                            End If
                        End If
                            
                        Permitido2 = False
                        Select Case Weekday(Inicio)
                            Case 1
                                If FunVerificaFeriado(Inicio) = False Then ProcTotalCadMaqTurnos "Domingo", False Else Permitido2 = True
                            Case 2
                                If FunVerificaFeriado(Inicio) = False Then ProcTotalCadMaqTurnos "Segunda", False Else Permitido2 = True
                            Case 3
                                If FunVerificaFeriado(Inicio) = False Then ProcTotalCadMaqTurnos "TerÁa", False Else Permitido2 = True
                            Case 4
                                If FunVerificaFeriado(Inicio) = False Then ProcTotalCadMaqTurnos "Quarta", False Else Permitido2 = True
                            Case 5
                                If FunVerificaFeriado(Inicio) = False Then ProcTotalCadMaqTurnos "Quinta", False Else Permitido2 = True
                            Case 6
                                If FunVerificaFeriado(Inicio) = False Then ProcTotalCadMaqTurnos "Sexta", False Else Permitido2 = True
                            Case 7
                                If FunVerificaFeriado(Inicio) = False Then ProcTotalCadMaqTurnos "S·bado", False Else Permitido2 = True
                        End Select
                        
                        If Permitido2 = True Then
                            'If Encontrou = False Then
                                Qtd = Qtd + 1
                                Inicio = Inicio - 1
                            'Else
                                'Qtd = Qtd - 1
                                'Inicio = Inicio + 1
                           ' End If
                        End If
                    Loop
                End If
                TBNivel13.Close
'            End If
            
Pula:
            TBProducaoFases!prazofinalinicio = Format(PrazoOP - Qtd, "dd/mm/yyyy")
            Dataini = Format(TBProducaoFases!prazofinalinicio, "dd/mm/yyyy")
            Dataini = Dataini + Qtd
            TBProducaoFases!PrazoFinal = Format(Dataini, "dd/mm/yyyy")
            PrazoOP = Format(TBProducaoFases!prazofinalinicio, "dd/mm/yyyy")
            TBProducaoFases.Update
        End If
        TBFases.Close
        TBProducaoFases.MoveNext
    Loop
End If
TBProducaoFases.Close

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcTotalCadMaqTurnos(Diasemana As String, AltPrazofinal As Boolean)
On Error GoTo tratar_erro

Execucao = 0
If AltPrazofinal = False Then
    Set TBNivel10 = CreateObject("adodb.recordset")
    TBNivel10.Open "Select Totalturno from cadmaqturnos where maquina = '" & TBProducaoFases!maquina & "' and diasemana = '" & Diasemana & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBNivel10.EOF = False Then
        Do While TBNivel10.EOF = False
            Execucao = Execucao + TBNivel10!TotalTurno
            TBNivel10.MoveNext
        Loop
        ElapsedTime (Execucao)
        
        TEUSEG = 0 'Tempo total utilizado por maquina
        'If Encontrou = False Then
            Set TBNivel11 = CreateObject("adodb.recordset")
            TBNivel11.Open "Select SUM(H.Segundos) AS TEUSEG from (ordemservico OS INNER JOIN Producao P ON P.Ordem = OS.Ordem) INNER JOIN Ordemservico_HoraUtilizadaporDia H ON OS.IDproducao = H.OS where P.Status <> 'Cancelada' and OS.maquina = '" & TBProducaoFases!maquina & "' AND H.Data = '" & Inicio & "' and OS.pronto = 'N„o' and OS.IDproducao <> " & TBProducaoFases!IDProducao, Conexao, adOpenKeyset, adLockOptimistic
            If TBNivel11.EOF = False Then
                TEUSEG = IIf(IsNull(TBNivel11!TEUSEG), 0, TBNivel11!TEUSEG)
            End If
            TBNivel11.Close
        'End If
        s = s - TEUSEG 'Segundos disponiveis em todos os turnos do posto e dia
    
        If s > 0 Then
            Set TBNivel12 = CreateObject("adodb.recordset")
            TBNivel12.Open "Select OS, Data, Segundos from Ordemservico_HoraUtilizadaporDia where OS = " & TBProducaoFases!IDProducao & " and Data = '" & Inicio & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBNivel12.EOF = True Then
                TBNivel12.AddNew
                TBNivel12!OS = TBProducaoFases!IDProducao
                TBNivel12!Data = Inicio
            End If
            'If Valor1 > S And Encontrou = True Then
            If Valor1 > s Then
                TBNivel12!Segundos = s
                Permitido2 = True
                Valor1 = Valor1 - s
            Else
                TBNivel12!Segundos = Valor1
                Valor1 = 0
            End If
            
            TBNivel12.Update
            TBNivel12.Close
        Else
            Permitido2 = True
        End If
    Else
        Permitido2 = True
    End If
    TBNivel10.Close
Else
    Permitido2 = True
End If

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Function FunDefinirPrazoPed(Prazo As Date) As Date
On Error GoTo tratar_erro

Inicio:
    If Weekday(Prazo) = 1 Then
        Prazo = Prazo + 1
        GoTo Inicio
    ElseIf Weekday(Prazo) = 7 Then
            Prazo = Prazo + 2
            GoTo Inicio
        ElseIf FunVerificaFeriado(Prazo) = True Then
                Prazo = Prazo + 1
                GoTo Inicio
    End If
    FunDefinirPrazoPed = Prazo

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Function FunVerificaFeriado(Data1 As Date) As Boolean
On Error GoTo tratar_erro

FunVerificaFeriado = False
Set TBCFOP = CreateObject("adodb.recordset")
TBCFOP.Open "Select ID from Feriados where Data_feriado = '" & Data & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBCFOP.EOF = False Then
    FunVerificaFeriado = True
End If
TBCFOP.Close

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Function FunArredondarPraCima(valor As Double)
On Error GoTo tratar_erro

v_1 = CSng(valor)
If valor > 0 And valor < 1 Then valor = 1
If valor > 0 Then
    v_2 = Format(CDbl(valor), "##,########")
    DIF1 = v_1 - v_2
    If DIF1 > 0 Then FunArredondarPraCima = v_2 + 1 Else FunArredondarPraCima = v_2
Else
    FunArredondarPraCima = valor
End If

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Function FunCalculaQtdePC(Codinterno As String, Qtde As Double, CalculaPC As Boolean, Unconversao As String) As Double
On Error GoTo tratar_erro

FunCalculaQtdePC = 0
If Codinterno = "" Then Exit Function
Set TBAliquota = CreateObject("adodb.recordset")
TBAliquota.Open "Select Un_Kg, PBruto, Unidade from projproduto where Desenho = '" & Codinterno & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBAliquota.EOF = False Then
    If TBAliquota!Unidade = "P«" Or TBAliquota!Unidade = "PC" Or TBAliquota!Unidade = "UN" Or TBAliquota!Unidade = "CJ" Then
        'Calcula quantidade se a unidade for diferente
        If TBAliquota!Unidade <> Unconversao And Unconversao <> "" Then
            If FunVerifUNConversao(TBAliquota!Unidade, Unconversao) = True Then
                FunCalculaQtdePC = FunConverteUN(TBAliquota!Unidade, Unconversao, Qtde, Codinterno)
            Else
                FunCalculaQtdePC = Qtde / FunVerificaTabelaConversaoUnidade(TBAliquota!Unidade, Unconversao)
            End If
        Else
            FunCalculaQtdePC = Qtde
        End If
    Else
        If Unconversao = "P«" Or Unconversao = "PC" Or Unconversao = "UN" Or Unconversao = "CJ" Then
            FunCalculaQtdePC = Qtde
        Else
            If CalculaPC = True Then
                If TBAliquota!Unidade = "KG" And IsNull(TBAliquota!PBruto) = False And TBAliquota!PBruto <> 0 And (TBAliquota!Un_Kg = "Mt≤" Or TBAliquota!Un_Kg = "Mt/L") Then FunCalculaQtdePC = Format(Qtde / TBAliquota!PBruto, "###,##0.0000")
            Else
                FunCalculaQtdePC = Format(Qtde * TBAliquota!PBruto, "###,##0.0000")
            End If
        End If
    End If
End If
TBAliquota.Close

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Sub ProcAtualizaCTTEROrdem(OS As Long)
On Error GoTo tratar_erro

Set TBOrdem = CreateObject("adodb.recordset")
TBOrdem.Open "Select CTServico, Ordem from Ordemservico where IDProducao = " & OS, Conexao, adOpenKeyset, adLockOptimistic
If TBOrdem.EOF = False Then
    Set TBproducao = CreateObject("adodb.recordset")
    'TBproducao.Open "Select Sum(ROUND(NFP.dbl_ValorUnitario * NFPP.Quantidade, 2)) as Valortotal from (tbl_Detalhes_Nota_pedidos NFPP INNER JOIN Compras_pedido_lista PP ON PP.IDlista = NFPP.ID_carteira and PP.Desenho = NFPP.Codinterno) INNER JOIN tbl_Detalhes_Nota NFP ON NFP.Int_codigo = NFPP.ID_prod_NF where PP.OS = " & OS, Conexao, adOpenKeyset, adLockOptimistic
    TBproducao.Open "Select Sum(ROUND(NFP.dbl_ValorUnitario * NFPP.Quantidade, 2)) as Valortotal from (Compras_pedido_lista CPL INNER JOIN tbl_Detalhes_Nota_pedidos NFPP ON NFPP.ID_carteira = CPL.IDlista and NFPP.Codinterno = CPL.Desenho) INNER JOIN tbl_Detalhes_Nota NFP ON NFP.Int_codigo = NFPP.ID_prod_NF where CPL.OS  = " & OS & " and CPL.Remessa = 'False' and (CPL.Status_Item = 'RECEBIDO' or CPL.Status_Item = 'PARCIAL')", Conexao, adOpenKeyset, adLockOptimistic
    If TBproducao.EOF = False Then
        TBOrdem!CTServico = IIf(IsNull(TBproducao!ValorTotal), 0, TBproducao!ValorTotal)
        TBOrdem.Update
    End If
    TBproducao.Close
    
    ProcAtualizaCTTerceirosOrdem TBOrdem!Ordem
End If

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Function FunCalculaQtdePCKG(Qtde_est_KG As Double, Qtde_est_PC As Double, Qtde_movimentacao As Double, CalculaPC As Boolean) As Double
On Error GoTo tratar_erro
Dim Kg_un As Double 'OK

FunCalculaQtdePCKG = 0
If Qtde_est_PC > 0 Then
    Kg_un = Format(Qtde_est_KG / Qtde_est_PC, "###,##0.0000000000")
    If CalculaPC = True Then
        If Kg_un > 0 Then FunCalculaQtdePCKG = Format(Qtde_movimentacao / Kg_un, "###,##0.0000")
    Else
        FunCalculaQtdePCKG = Format(Qtde_movimentacao * Kg_un, "###,##0.0000")
    End If
    'If CalculaPC = True Then FunCalculaQtdePCKG = Qtde_movimentacao / Kg_un Else FunCalculaQtdePCKG = Qtde_movimentacao * Kg_un
End If
    
Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Sub ProcAtualizaCTTerceirosOrdem(Ordem As Long)
On Error GoTo tratar_erro

Set TBOrdem = CreateObject("adodb.recordset")
TBOrdem.Open "Select Ordem, CTServico from producao where Ordem = " & Ordem, Conexao, adOpenKeyset, adLockOptimistic
If TBOrdem.EOF = False Then
    Set TBproducao = CreateObject("adodb.recordset")
    TBproducao.Open "Select Sum(CTServico) as Valortotal from ordemservico where Ordem = " & TBOrdem!Ordem, Conexao, adOpenKeyset, adLockOptimistic
    If TBproducao.EOF = False Then
        TBOrdem!CTServico = IIf(IsNull(TBproducao!ValorTotal), 0, TBproducao!ValorTotal)
        TBOrdem.Update
    End If
    TBproducao.Close
End If

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Function FunFormataCasasDecimais(Casa As Integer, valor As Double) As String
On Error GoTo tratar_erro

Select Case Casa
    Case "1": FunFormataCasasDecimais = Format(valor, "###,##0.0")
    Case "2": FunFormataCasasDecimais = Format(valor, "###,##0.00")
    Case "3": FunFormataCasasDecimais = Format(valor, "###,##0.000")
    Case "4": FunFormataCasasDecimais = Format(valor, "###,##0.0000")
    Case "5": FunFormataCasasDecimais = Format(valor, "###,##0.00000")
    Case "10": FunFormataCasasDecimais = Format(valor, "###,##0.0000000000")
End Select

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Function FunValidarCliente()
On Error GoTo tratar_erro

Set TBFIltro = CreateObject("adodb.recordset")
TBFIltro.Open "Select * from Empresa", Conexao, adOpenKeyset, adLockOptimistic
If TBFIltro.EOF = False Then
    Do While TBFIltro.EOF = False
        FunAbreBDSite
        If ConexaoMySql.State = 1 Then
            Set TBMySQL = New ADODB.Recordset
            TBMySQL.Open "Select * From Clientes Where CNPJ = '" & TBFIltro!CNPJ & "'", ConexaoMySql, adOpenKeyset, adLockOptimistic, adCmdText
            With TBMySQL
                If .EOF = False Then
                    .Fields!NomeRazao = TBFIltro!Razao
                    .Update
                    
                    TBFIltro!Licencas_caprind = IIf(IsNull(.Fields!Licencas), 0, .Fields!Licencas)
                    TBFIltro!Licencas_gerprod = IIf(IsNull(.Fields!Licencas_gerprod), 0, .Fields!Licencas_gerprod)
                    TBFIltro!Modulo = .Fields!Modulo
                    TBFIltro.Update
                    
                    Permitido = True
                    If .Fields!Liberado = "N√O" Then
                        Permitido = False
                        If IsNull(.Fields!Codigo_do_erro) = True Or .Fields!Codigo_do_erro = 0 Then
                            mensagemerro = "N„o È permitido efetuar o logon, pois o mesmo est· com o acesso bloqueado"
                        Else
                            Select Case .Fields!Codigo_do_erro
                                Case 1: mensagemerro = "Error Accessing the system registry"
                                Case 2: mensagemerro = "Invalid procedure call or argument"
                                Case 3: mensagemerro = "Out of memory"
                                Case 4: mensagemerro = "Many client applications trying to access the DLL at the same time"
                                Case 5: mensagemerro = "Server object was not properly registered or not found"
                                Case 6: mensagemerro = "ActiveX Control not found"
                                Case 7: mensagemerro = "License information for this component not found or not found DLL"
                                Case 8: mensagemerro = "Invalid number of arguments or invalid property assignment"
                                Case 9: mensagemerro = "Syntax error (missing operator)"
                            End Select
                        End If
                    End If
                    If Permitido = False Then
                        USMsgBox (mensagemerro & "."), vbCritical, "CAPRIND v5.0"
                        End
                    End If
                    
                    'Verifica n˙mero de licenÁas
VerifLicenca:
                    If IsNull(.Fields!Licencas) = False And .Fields!Licencas <> "" And pubUsuario <> "PROCAM" Then
                        Set TBLogon = CreateObject("adodb.recordset")
                        TBLogon.Open "Select IDLogon from Logon where Usuario <> 'PROCAM' and Data = '" & Date & "' and Tipo = 'C'", Conexao, adOpenKeyset, adLockOptimistic
                        If TBLogon.EOF = False Then
                            Contador = TBLogon.RecordCount + 1
                            Contador2 = .Fields!Licencas
                            If Contador > Contador2 Then
                                Formulario = "ConfiguraÁ„o do sistema/Usu·rios/Conectados"
                                ProcLiberaAcessos False
                                If Acessos = True Then
                                    If USMsgBox("N„o È permitido efetuar o logon, pois foi atingido o n˙mero m·ximo de usu·rios permitidos. Deseja desconectar algum usu·rio?", vbYesNo, "CAPRIND v5.0") = vbYes Then
                                        Permitido2 = True
                                        frmOpcoes_Lista_usuarios.Show (1)
                                        If Permitido2 = True Then GoTo VerifLicenca
                                    Else
                                        USMsgBox ("N„o È permitido efetuar o logon, pois foi atingido o n˙mero m·ximo de usu·rios permitidos. O sistema Caprind ser· encerrado."), vbCritical, "CAPRIND v5.0"
                                        End
                                    End If
                                Else
                                    USMsgBox ("N„o È permitido efetuar o logon, pois foi atingido o n˙mero m·ximo de usu·rios permitidos. O sistema Caprind ser· encerrado."), vbCritical, "CAPRIND v5.0"
                                    End
                                End If
                            End If
                        End If
                        TBLogon.Close
                    End If
                Else
                    .AddNew
                    .Fields!Responsavel = pubUsuario
                    .Fields!Cargo = pubSetor
                    .Fields!Liberado = "SIM"
                    .Fields!NomeRazao = TBFIltro!Razao
                    .Fields!CNPJ = TBFIltro!CNPJ
                    .Fields!telefone = TBFIltro!telefone
                    .Fields!Email = TBFIltro!Email
                    .Update
                End If
            End With
        End If
        TBFIltro.MoveNext
    Loop
End If
TBFIltro.Close
FunValidarUsuario

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Function FunValidarUsuario()
On Error GoTo tratar_erro

Set TBFIltro = CreateObject("adodb.recordset")
TBFIltro.Open "Select * from Empresa", Conexao, adOpenKeyset, adLockOptimistic
If TBFIltro.EOF = False Then
    Do While TBFIltro.EOF = False
        FunAbreBDSite
        If ConexaoMySql.State = 1 Then
            Set TBMySQL = New ADODB.Recordset
            TBMySQL.Open "Select * From usuarios Where CNPJ = '" & TBFIltro!CNPJ & "' and Usuario = '" & pubUsuario & "'", ConexaoMySql, adOpenKeyset, adLockOptimistic, adCmdText
            With TBMySQL
                If .EOF = True Then
                    .AddNew
                    .Fields!Liberado = "SIM"
                Else
                    Permitido = True
                    If .Fields!Liberado = "N√O" Then
                        Permitido = False
                        If IsNull(.Fields!Codigo_do_erro) = True Or .Fields!Codigo_do_erro = 0 Then
                            mensagemerro = "N„o È permitido o usu·rio " & pubUsuario & " efetuar o logon, pois o mesmo est· com o acesso bloqueado"
                        Else
                            Select Case .Fields!Codigo_do_erro
                                Case 1: mensagemerro = "Error Accessing the system registry"
                                Case 2: mensagemerro = "Invalid procedure call or argument."
                                Case 3: mensagemerro = "Out of memory"
                                Case 4: mensagemerro = "Many client applications trying to access the DLL at the same time"
                                Case 5: mensagemerro = "Server object was not properly registered or not found"
                                Case 6: mensagemerro = "ActiveX Control not found"
                                Case 7: mensagemerro = "License information for this component not found or not found DLL"
                                Case 8: mensagemerro = "Invalid number of arguments or invalid property assignment"
                                Case 9: mensagemerro = "Syntax error (missing operator)"
                            End Select
                        End If
                    End If
                    If Permitido = False Then
                        USMsgBox (mensagemerro & "."), vbCritical, "CAPRIND v5.0"
                        End
                    End If
                End If
                .Fields!CNPJ = TBFIltro!CNPJ
                .Fields!Usuario = pubUsuario
                .Fields!Nome = pubNome
                .Fields!Senha = pubSenha
                .Fields!Cargo = pubSetor
                If pubEmail <> "" Then .Fields!Email = pubEmail
                .Fields!Logado = "SIM"
                .Fields!Nivel = 2
                .Fields!ativo = 1
                .Update
                
                caminho = App.Path & "\CI.txt"
                Set GerArqPastas = CreateObject("Scripting.FileSystemObject")
                If GerArqPastas.FileExists(caminho) = True Then GerArqPastas.DeleteFile (caminho) 'Deleta o arquivo antigo
                
                arq = FreeFile 'Esta linha atribui um espaÁo livre na memÛria para armazenar o arquivo.
                Open caminho For Append As #arq 'Escreve em um arquivo j· existente
                
                'Verifica a data limite para utilizaÁ„o sem internet e converte (15 dias)
                Contador = 1
                Dataini = Date + 15
                DadosArquivo = "@%&*#"
                Do While Contador <= 8
                    Select Case Mid(Format(Dataini, "dd/mm/yy"), Contador, 1)
                        Case 0: p = "!"
                        Case 1: p = "#"
                        Case 2: p = "S"
                        Case 3: p = "&"
                        Case 4: p = "|"
                        Case 5: p = "Z"
                        Case 6: p = "@"
                        Case 7: p = "$"
                        Case 8: p = "T"
                        Case 9: p = "^"
                        Case "/": p = "?"
                    End Select
                    DadosArquivo = DadosArquivo & p
                    Contador = Contador + 1
                Loop
                DadosArquivo = DadosArquivo & "&*$@!"
                Print #arq, DadosArquivo; 'Escreve a data limite para utilizaÁ„o sem internet
                Close (arq) 'Fecha o arquivo
                GerArqPastas.GetFile(caminho).Attributes = Hidden
                
            End With
        End If
        TBFIltro.MoveNext
    Loop
End If

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Function FunValidarClienteSemInternet()
On Error GoTo tratar_erro

Set TBFIltro = CreateObject("adodb.recordset")
TBFIltro.Open "Select Licencas_caprind from Empresa where Licencas_caprind IS NOT NULL and Licencas_caprind <> N''", Conexao, adOpenKeyset, adLockOptimistic
If TBFIltro.EOF = False Then
    Do While TBFIltro.EOF = False
VerifLicenca:
        Set TBLogon = CreateObject("adodb.recordset")
        TBLogon.Open "Select IDLogon from Logon where Usuario <> 'PROCAM' and Data = '" & Date & "' and Tipo = 'C'", Conexao, adOpenKeyset, adLockOptimistic
        If TBLogon.EOF = False Then
            Contador = TBLogon.RecordCount + 1
            Contador2 = TBFIltro!Licencas_caprind
            If Contador > Contador2 Then
                Formulario = "ConfiguraÁ„o do sistema/Usu·rios/Conectados"
                ProcLiberaAcessos False
                If Acessos = True Then
                    If USMsgBox("N„o È permitido efetuar o logon, pois foi atingido o n˙mero m·ximo de usu·rios permitidos. Deseja desconectar algum usu·rio?", vbYesNo, "CAPRIND v5.0") = vbYes Then
                        Permitido2 = True
                        frmOpcoes_Lista_usuarios.Show (1)
                        If Permitido2 = True Then GoTo VerifLicenca
                    Else
                        USMsgBox ("N„o È permitido efetuar o logon, pois foi atingido o n˙mero m·ximo de usu·rios permitidos. O sistema Caprind ser· encerrado."), vbCritical, "CAPRIND v5.0"
                        End
                    End If
                Else
                    USMsgBox ("N„o È permitido efetuar o logon, pois foi atingido o n˙mero m·ximo de usu·rios permitidos. O sistema Caprind ser· encerrado."), vbCritical, "CAPRIND v5.0"
                    End
                End If
            End If
        End If
        TBLogon.Close
        TBFIltro.MoveNext
    Loop
End If
TBFIltro.Close

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Sub ProcVerificaInternet(Logon As Boolean, ChatLogon As Boolean)
On Error GoTo tratar_erro
Dim ip As String 'OK

TemInternet = DS.IsInternetOnline

If TemInternet = False Then
    With frmabertura
            If Logon = True Then
                .txtUsuario.Enabled = False
                .txtSenha.Enabled = False
                .cmbBanco.Enabled = False
                .Cmd_novo_local_bd.Enabled = False
                .Cmd_chat.Enabled = False
                .cmdAcessar.Enabled = False
            End If
    End With
    ProcVerifDiasUtilizadosSemInternet
Else
    If Logon = True Or ChatLogon = True Then
        With frmabertura
            If Logon = True Then
                .txtUsuario.Enabled = False
                .txtSenha.Enabled = False
                .cmbBanco.Enabled = False
                .Cmd_novo_local_bd.Enabled = False
                .Cmd_chat.Enabled = False
                .cmdAcessar.Enabled = False
            End If
        End With
    End If
End If

'If Logon = True Then
'    With frmabertura
'        .txtUsuario.Enabled = True
'        .txtsenha.Enabled = True
'        .cmbBanco.Enabled = True
'        .Cmd_novo_local_bd.Enabled = True
'        .Cmd_chat.Enabled = True
'        .cmdAcessar.Enabled = True
'    End With
'End If

If TemInternet = True Then
    Datainitexto = Date
    SaveSetting "{6F6CC9481C35-D412-2107-A1Z1-B6261BD8}", "Default", "Main", "DR1AG4HJ78" & FunTamanhoTextoZeroEsq(Day(Date), 2) & "CGT85910" & FunTamanhoTextoZeroEsq(Month(Date), 2) & "RGTOPU" & Year(Date) & "58963557"
End If

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcVerifDiasUtilizadosSemInternet()
On Error GoTo tratar_erro

caminho = App.Path & "\CI.txt"

'Verifica se o arquivo esta na pasta
Set GerArqPastas = CreateObject("Scripting.FileSystemObject")
If GerArqPastas.FileExists(caminho) = False Then
    USMsgBox ("Ocorreu um erro inesperado, o sistema ser· encerrado."), vbCritical, "CAPRIND v5.0"
    End
Else
    'Verifica se o arquivo n„o foi alterado
    arq = FreeFile 'Esta linha atribui um espaÁo livre na memÛria para armazenar o arquivo.
    Open caminho For Input As #arq 'Abre o arquivo
    Line Input #arq, Linha 'Ler uma linha
    
    Permitido = True
    If Len(Linha) <> 18 Then
        Permitido = False
        GoTo Parar
    End If
    
    'Verifica se existe algum caracter diferente
    Contador = 6
    Do While Contador <= 13
        If Mid(Linha, Contador, 1) <> "!" And Mid(Linha, Contador, 1) <> "#" And Mid(Linha, Contador, 1) <> "S" And Mid(Linha, Contador, 1) <> "&" And Mid(Linha, Contador, 1) <> "|" And Mid(Linha, Contador, 1) <> "Z" And Mid(Linha, Contador, 1) <> "@" And Mid(Linha, Contador, 1) <> "$" And Mid(Linha, Contador, 1) <> "T" And Mid(Linha, Contador, 1) <> "^" And Mid(Linha, Contador, 1) <> "?" Then Permitido = False
        Contador = Contador + 1
    Loop
    Close (arq) 'Fecha o arquivo

Parar:
    If Permitido = False Then
        USMsgBox ("Ocorreu um erro inesperado, o sistema ser· encerrado."), vbCritical, "CAPRIND v5.0"
        End
    End If
    
    'Verifica dados do arquivo e carrega nas variaveis
    arq = FreeFile 'Esta linha atribui um espaÁo livre na memÛria para armazenar o arquivo.
    Open caminho For Input As #arq 'Abre o arquivo
    Line Input #arq, Linha 'Ler uma linha
    
    DataTexto = ""
    'Verifica a data que esta no txt e converte
    Contador = 6
    Do While Contador <= 13
        Select Case Mid(Linha, Contador, 1)
            Case "!": p = 0
            Case "#": p = 1
            Case "S": p = 2
            Case "&": p = 3
            Case "|": p = 4
            Case "Z": p = 5
            Case "@": p = 6
            Case "$": p = 7
            Case "T": p = 8
            Case "^": p = 9
            Case "?": p = "/"
        End Select
        If DataTexto = "" Then DataTexto = p Else DataTexto = DataTexto & p
        Contador = Contador + 1
    Loop
    Dataini = DataTexto
    If Format(Date, "dd/mm/yy") > Dataini Then
        USMsgBox ("Venceu o prazo de 15 dias para utilizaÁ„o do Caprind sem internet, o sistema ser· encerrado. Favor entar em contato com o suporte atravÈs do e-mail suporte@caprind.com.br."), vbCritical, "CAPRIND v5.0"
        End
    End If
    Close (arq) 'Fecha o arquivo
    
    Datainitexto = GetSetting("{6F6CC9481C35-D412-2107-A1Z1-B6261BD8}", "Default", "Main") 'Verifica a data inicial sem internet
    If Datainitexto <> "" Then
        Datainitexto = GetSetting("{6F6CC9481C35-D412-2107-A1Z1-B6261BD8}", "Default", "Main") 'Carrega data incial sem internet
        DiaTexto = Mid(Datainitexto, 11, 2)
        MesTexto = Mid(Datainitexto, 21, 2)
        AnoTexto = Mid(Datainitexto, 29, 4)
        Datainitexto = DiaTexto & "/" & MesTexto & "/" & AnoTexto
    Else
        Datainitexto = Date
        SaveSetting "{6F6CC9481C35-D412-2107-A1Z1-B6261BD8}", "Default", "Main", "DR1AG4HJ78" & FunTamanhoTextoZeroEsq(Day(Date), 2) & "CGT85910" & FunTamanhoTextoZeroEsq(Month(Date), 2) & "RGTOPU" & Year(Date) & "58963557" 'Salva data incial sem internet
    End If
    Dataini = Datainitexto
    Dia = DateDiff("d", Dataini, Date)
    If Dia > 15 Then
        USMsgBox ("Venceu o prazo de 15 dias para utilizaÁ„o do Caprind sem internet, o sistema ser· encerrado. Favor entar em contato com o suporte atravÈs do e-mail suporte@caprind.com.br."), vbCritical, "CAPRIND v5.0"
        End
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Function FunVerifIDPlano(IDFase As Long) As Long
On Error GoTo tratar_erro

FunVerifIDPlano = 0
If IDFase <> 0 Then
    Set TBplano = CreateObject("adodb.recordset")
    TBplano.Open "Select * from Plano where IDfase = " & IDFase, Conexao, adOpenKeyset, adLockOptimistic
    If TBplano.EOF = False Then
        FunVerifIDPlano = TBplano!IDPlano
    End If
    TBplano.Close
End If

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Sub ProcVerifInstancia(Combo As ComboBox)
On Error GoTo tratar_erro

With Combo
    .Clear
    Screen.MousePointer = vbHourglass
    For Each vSrv In EnumSqlServers
        .AddItem vSrv
    Next
    m_bEnumSrv = True
    Screen.MousePointer = vbDefault
End With

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcAgregarProdutoForn(Codproduto As Long, IDforn As Long, ValorCusto As Double)
On Error GoTo tratar_erro

Permitido2 = False
Set TBTempo = CreateObject("adodb.recordset")
TBTempo.Open "Select * from Projproduto_fornecedor where codproduto = " & Codproduto & " and IDfornecedor = " & IDforn, Conexao, adOpenKeyset, adLockOptimistic
If TBTempo.EOF = True Then
    TBTempo.AddNew
    Permitido2 = True
Else
    If TBTempo!PCusto = 0 Then Permitido2 = True
End If
If Permitido2 = True Then
    TBTempo!Codproduto = Codproduto
    TBTempo!IDFornecedor = IDforn
    TBTempo!PCusto = ValorCusto
    TBTempo.Update
End If
TBTempo.Close

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcAgregarProdutoCli(Codproduto As Long, IDCli As Long, TipoCli As String, Unest As String, Uncom As String, valor As Double)
On Error GoTo tratar_erro

Permitido1 = False
Set TBTempo = CreateObject("adodb.recordset")
TBTempo.Open "Select * from Projproduto_clientes where codproduto = " & Codproduto & " and Idcliente = " & IDCli, Conexao, adOpenKeyset, adLockOptimistic
If TBTempo.EOF = True Then
    TBTempo.AddNew
    Permitido1 = True
Else
    If Right(TipoCli, 1) = "P" And TBTempo!PConsumo = 0 Or Right(TipoCli, 1) = "R" And TBTempo!PRevenda = 0 Then Permitido1 = True
End If
If Permitido1 = True Then
    TBTempo!Codproduto = Codproduto
    TBTempo!IDCliente = IDCli
    valor = valor * FunVerificaTabelaConversaoUnidade(Unest, Uncom)
    If Right(TipoCli, 1) = "P" Then TBTempo!PConsumo = valor Else TBTempo!PRevenda = valor
    TBTempo.Update
End If
TBTempo.Close

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Function FunVerifMovimentacaoEstPC(ID_empresa As Integer) As Boolean
On Error GoTo tratar_erro

FunVerifMovimentacaoEstPC = False
Set TBTempo = CreateObject("adodb.recordset")
TBTempo.Open "Select Codigo from Empresa where Codigo = " & ID_empresa & " and Movimentar_estoque_pc = 'True'", Conexao, adOpenKeyset, adLockOptimistic
If TBTempo.EOF = False Then
    FunVerifMovimentacaoEstPC = True
End If
TBTempo.Close

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Function FunVerifValidarAutomPropPI(ID_empresa As Integer) As Boolean
On Error GoTo tratar_erro

FunVerifValidarAutomPropPI = False
Set TBTempo = CreateObject("adodb.recordset")
TBTempo.Open "Select Codigo from Empresa where Codigo = " & ID_empresa & " and Validar_prop_pi_autom = 'True'", Conexao, adOpenKeyset, adLockOptimistic
If TBTempo.EOF = False Then
    FunVerifValidarAutomPropPI = True
End If
TBTempo.Close

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Function FunVerifProdSimiliar(ID_empresa As Integer) As Boolean
On Error GoTo tratar_erro

FunVerifProdSimiliar = False
Set TBTempo = CreateObject("adodb.recordset")
TBTempo.Open "Select Codigo from Empresa where Codigo = " & ID_empresa & " and Ativar_prod_similares = 'True'", Conexao, adOpenKeyset, adLockOptimistic
If TBTempo.EOF = False Then
    FunVerifProdSimiliar = True
End If
TBTempo.Close

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Function FunVerifCodRefFornSPED(ID_empresa As Long) As Boolean
On Error GoTo tratar_erro

FunVerifCodRefFornSPED = False
Set TBTempo = CreateObject("adodb.recordset")
TBTempo.Open "Select Codigo from Empresa where Codigo = " & ID_empresa & " and Codigo_ref_SPED_forn = 'True'", Conexao, adOpenKeyset, adLockOptimistic
If TBTempo.EOF = False Then
    FunVerifCodRefFornSPED = True
End If
TBTempo.Close

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Function FunVerifCodRefCliDANFE(ID_empresa As Integer) As Boolean
On Error GoTo tratar_erro

FunVerifCodRefCliDANFE = False
Set TBTempo = CreateObject("adodb.recordset")
TBTempo.Open "Select Codigo from Empresa where Codigo = " & ID_empresa & " and Codigo_ref_DANFE = 'True'", Conexao, adOpenKeyset, adLockOptimistic
If TBTempo.EOF = False Then
    FunVerifCodRefCliDANFE = True
End If
TBTempo.Close

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Function FunVerifStatusAprovadoPC(ID_empresa As Integer) As Boolean
On Error GoTo tratar_erro

FunVerifStatusAprovadoPC = False
Set TBTempo = CreateObject("adodb.recordset")
TBTempo.Open "Select Codigo from Empresa where Codigo = " & ID_empresa & " and Salvar_status_aprovado_PC = 'True'", Conexao, adOpenKeyset, adLockOptimistic
If TBTempo.EOF = False Then
    FunVerifStatusAprovadoPC = True
End If
TBTempo.Close

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Function FunVerifEnviarEmailOutlook(ID_empresa As Integer) As Boolean
On Error GoTo tratar_erro

FunVerifEnviarEmailOutlook = False
Set TBTempo = CreateObject("adodb.recordset")
TBTempo.Open "Select Codigo from Empresa where Codigo = " & ID_empresa & " and Enviar_email_outlook = 'True'", Conexao, adOpenKeyset, adLockOptimistic
If TBTempo.EOF = False Then
    FunVerifEnviarEmailOutlook = True
End If
TBTempo.Close

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Sub ProcAcertaOS(quantidade As Double, Alteracao_OS As Boolean)
On Error GoTo tratar_erro

TotalFaseSeg = 0
TotalFaseSegPrep = 0
CustoOrdem = 0
TotalOrdem = 0
TotalOS = 0

PcHora = 0
Set TBOSC = CreateObject("adodb.recordset")
TBOSC.Open "Select OS.*, P.oscontrolada As OPCont, P.Processo_controlado As OPPCont from ordemServico OS INNER JOIN Producao P ON P.Ordem = OS.Ordem where OS.Ordem = " & OF & " ORDER BY OS.fase, OS.retrabalho, OS.IDproducao", Conexao, adOpenKeyset, adLockOptimistic
If TBOSC.EOF = False Then
    TBOSC.MoveFirst
    Do Until TBOSC.EOF
        TOTALPECA = 0
        TotalOS = 0
        
        'QtdeAntiga = TBOSC!Quantidade
        If TBOSC!Retrabalho = False Then TBOSC!quantidade = quantidade
        
        'Verifica se a maquina agrega custos/eficiencia na ordem
        Set TBMaquinas = CreateObject("adodb.recordset")
        TBMaquinas.Open "Select * from cadmaquinas where maquina = '" & TBOSC!maquina & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBMaquinas.EOF = False Then
            If TBMaquinas!custos = True Then TBOSC!custos = True Else TBOSC!custos = False
            If IsNull(TBMaquinas!PrecoHora_Setup) = False And TBMaquinas!PrecoHora_Setup <> "" Then TBOSC!Valor_hs_prep = TBMaquinas!PrecoHora_Setup Else TBOSC!Valor_hs_prep = IIf(IsNull(TBMaquinas!PrecoHora), 0, TBMaquinas!PrecoHora)
            TBOSC!Valor_hs_exec = IIf(IsNull(TBMaquinas!PrecoHora), 0, TBMaquinas!PrecoHora)
        End If
        TBMaquinas.Close
        
        PcHora = IIf(IsNull(TBOSC!pc_te) = False, TBOSC!pc_te, 1)
        'Tempo total por peÁa
        ElapsedTime (IIf(IsNull(TBOSC!Execucao), 0, TBOSC!Execucao))
        If PcHora <> 0 Then
            TOTALPECA = TOTALPECA + (s / PcHora)
            TotalFaseSeg = s / PcHora
        Else
            TOTALPECA = TOTALPECA + s
            TotalFaseSeg = s
        End If
        
        'Tempo total do lote
        If PcHora <> 0 Then TotalOS = s / PcHora Else TotalOS = s
        ElapsedTime (IIf(IsNull(TBOSC!Preparacao), 0, TBOSC!Preparacao))
        TotalFaseSegPrep = s
        TotalOS = (TotalOS * TBOSC!quantidade) + s
        
        'Verifica custo de execucao por segundos * custo da hora maquina
        CustohoraSeg = TBOSC!Valor_hs_exec / 3600
        CustoFase = CustohoraSeg * TotalFaseSeg
        
        'Verifica custo de preparacao por segundos * custo da hora maquina
        If IsNull(TBOSC!Valor_hs_prep) = False And TBOSC!Valor_hs_prep <> "" Then CustohoraSeg = TBOSC!Valor_hs_prep / 3600
        CustopreparacaoSeg = CustohoraSeg * TotalFaseSegPrep
        
        'Custo por peÁa
        If TBOSC!quantidade <> 0 Then
        TBOSC!CPPECA = Format(CustoFase + (CustopreparacaoSeg / TBOSC!quantidade), "###,##0.0000000000")
        'Custo do lote
        TBOSC!CPLOTE = Format(TBOSC!CPPECA * TBOSC!quantidade, "###,##0.0000000000")
        CustoOrdem = CustoOrdem + TBOSC!CPLOTE
        End If
        
        If TBOSC!pecahora = True Then
            TBOSC!Pcshora = IIf(IsNull(TBOSC!pc_te) = False, TBOSC!pc_te, 1)
        Else
            If IsNull(TBOSC!Execucao) = False And TBOSC!Execucao <> "00:00:00" Then
                ElapsedTime (TBOSC!Execucao)
                TBOSC!Pcshora = 3600 / s
            End If
        End If
        'Tempo total por peÁa
        TBOSC!TempoExecucao = TOTALPECA
        TBOSC!TempoExecucao = FormataTempo(TBOSC!TempoExecucao)
        
        TBOSC!TTLPREVS = TotalOS 'Tempo total do lote previsto em segundos
        TBOSC!TempoTotalLote = FormataTempo(TBOSC!TTLPREVS) 'Tempo total do lote previsto
        
        TBOSC!TESegundos = FunCalculaSegPC(TBOSC!Execucao, TBOSC!pc_te)
        
        TotalOrdem = TotalOrdem + TotalOS
        
        If Alteracao_OS = False Then
            If TBOSC!Pronto = "N√O" Then
                TBOSC!OSControlada = TBOSC!OPCont
                TBOSC!Processo_controlado = TBOSC!OPPCont
            End If
        End If
                
        TBOSC.Update
        TBOSC.MoveNext
    Loop
End If
TBOSC.Close
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from producao where Ordem = " & OF, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    QuantSolicitado = TBAbrir!Quant
    
    'Custo por unidade
    If Int(QuantSolicitado) <> 0 Then TBAbrir!cpp = CustoOrdem / Int(QuantSolicitado) Else TBAbrir!cpp = CustoOrdem
    'Custo do lote
    TBAbrir!CTTPrev = CustoOrdem
    
    'Tempo total do lote
    TBAbrir!TTTPrev = TotalOrdem
    TBAbrir!TTTPrev = FormataTempo(TBAbrir!TTTPrev)
    'Tempo total do lote em segundos
    TBAbrir!TTTPREVSegundos = TotalOrdem
    
    'Tempo total por unidade
    If TotalOrdem <> 0 Then
        If QuantSolicitado >= 1 Then
        TotalOrdem = TotalOrdem / QuantSolicitado
        'Else
        TBAbrir!TPP = TotalOrdem
        End If
        
        If QuantSolicitado < 1 Then
        'TotalOrdem = TotalOrdem * QuantSolicitado
        'Else
        TBAbrir!TPP = TotalOrdem
        End If
        
        TBAbrir!TPP = FormataTempo(TBAbrir!TPP)
    Else
        TBAbrir!TPP = "00:00:00"
    End If
    
    TBAbrir.Update
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Function FunVerifApagaIDSimilar(ID_similar) As Integer
On Error GoTo tratar_erro

FunVerifApagaIDSimilar = ID_similar
Set TBCFOP = CreateObject("adodb.recordset")
TBCFOP.Open "Select Codproduto from Projproduto where ID_similar = " & ID_similar, Conexao, adOpenKeyset, adLockOptimistic
If TBCFOP.EOF = True Then
    Conexao.Execute "DELETE Projproduto_similar where ID = " & ID_similar
    FunVerifApagaIDSimilar = 0
End If
TBCFOP.Close

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Sub ProcSalvarViaOrdem(Ordem As Long, CarregarCampoImpresso As Boolean)
On Error GoTo tratar_erro

If NomeRel = "Pcp_ordem.rpt" Or NomeRel = "Pcp_ordem e rm.rpt" Or NomeRel = "Pcp_ordem e rm_resumido.rpt" Or NomeRel = "Pcp_ordem e rm_selecionadas.rpt" Or NomeRel = "Pcp_ordem_apontamento manual.rpt" Or NomeRel = "Pcp_rm.rpt" Then
    Set TBAliquota = CreateObject("adodb.recordset")
    TBAliquota.Open "Select Impof from Producao where Ordem = " & Ordem, Conexao, adOpenKeyset, adLockOptimistic
    If TBAliquota.EOF = False Then
        TBAliquota!Impof = IIf(IsNull(TBAliquota!Impof), 0, TBAliquota!Impof) + 1
        TBAliquota.Update
    End If
    TBAliquota.Close
    
    If CarregarCampoImpresso = True Then
        With frmprod
            If .txtof <> "" Then
                If Ordem = .txtof Then .Chk_impressa.Value = 1
            End If
        End With
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcINSERTINTO(NomeTabela As String, NomeCampos As String, Valores As String)
On Error GoTo tratar_erro
'Debug.print Valores
'Debug.print NomeCampos

Conexao.Execute "INSERT INTO " & NomeTabela & " (" & NomeCampos & ") VALUES (" & Valores & ")"

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCorrigeColunasForm(Lista As ListView, TextoFiltro As String, Numero_colunas As Integer, VerifCol1 As Boolean, TamCol1 As Integer, TamCol2 As Integer, TamCol3 As Integer, TamCol4 As Integer, TamCol5 As Integer, TamCol6 As Integer, TamCol7 As Integer, TamCol8 As Integer, TamCol9 As Integer, TamCol10 As Integer, TamCol11 As Integer, TamCol12 As Integer, TamCol13 As Integer, TamCol14 As Integer, TamCol15 As Integer, TamCol16 As Integer, TamCol17 As Integer, TamCol18 As Integer, TamCol19 As Integer, TamCol20 As Integer, TamCol21 As Integer, TamCol22 As Integer, TamCol23 As Integer, TamCol24 As Integer, TamCol25 As Integer, TamCol26 As Integer, TamCol27 As Integer, TamCol28 As Integer, TamCol29 As Integer, TamCol30 As Integer, TamCol31 As Integer)
On Error GoTo tratar_erro

With Lista
    Contador = 1
    Set TBAbrir = CreateObject("adodb.recordset")
    StrSql = "Select ID, Coluna1 from Usuarios_monitor_trabalho where ID_usuario = " & pubIDUsuario & " and Modulo = '" & TextoFiltro & "'"
    'Debug.print StrSql
    
    TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Do While Contador <= Numero_colunas
            If Contador = 1 Then
                If VerifCol1 = True Then
                    If TBAbrir!Coluna1 = False Then .ColumnHeaders(1).Width = 0 Else .ColumnHeaders(1).Width = TamCol1
                End If
            Else
                Select Case Contador
                    Case 2: TamCol = TamCol2
                    Case 3: TamCol = TamCol3
                    Case 4: TamCol = TamCol4
                    Case 5: TamCol = TamCol5
                    Case 6: TamCol = TamCol6
                    Case 7: TamCol = TamCol7
                    Case 8: TamCol = TamCol8
                    Case 9: TamCol = TamCol9
                    Case 10: TamCol = TamCol10
                    Case 11: TamCol = TamCol11
                    Case 12: TamCol = TamCol12
                    Case 13: TamCol = TamCol13
                    Case 14: TamCol = TamCol14
                    Case 15: TamCol = TamCol15
                    Case 16: TamCol = TamCol16
                    Case 17: TamCol = TamCol17
                    Case 18: TamCol = TamCol18
                    Case 19: TamCol = TamCol19
                    Case 20: TamCol = TamCol20
                    Case 21: TamCol = TamCol21
                    Case 22: TamCol = TamCol22
                    Case 23: TamCol = TamCol23
                    Case 24: TamCol = TamCol24
                    Case 25: TamCol = TamCol25
                    Case 26: TamCol = TamCol26
                    Case 27: TamCol = TamCol27
                    Case 28: TamCol = TamCol28
                    Case 29: TamCol = TamCol29
                    Case 30: TamCol = TamCol30
                    Case 31: TamCol = TamCol31
                End Select
                
                CampoFiltro = "Coluna" & Contador
                Set TBAbrir = CreateObject("adodb.recordset")
                StrSql = "Select ID from Usuarios_monitor_trabalho where ID_usuario = " & pubIDUsuario & " and Modulo = '" & TextoFiltro & "' and " & CampoFiltro & " = 'False'"
                'Debug.print StrSql
                
                TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    .ColumnHeaders(Contador).Width = 0
                Else
                    .ColumnHeaders(Contador).Width = TamCol
                End If
            End If
            Contador = Contador + 1
        Loop
    Else
        Do While Contador <= Numero_colunas
            Select Case Contador
                Case 1: TamCol = TamCol1
                Case 2: TamCol = TamCol2
                Case 3: TamCol = TamCol3
                Case 4: TamCol = TamCol4
                Case 5: TamCol = TamCol5
                Case 6: TamCol = TamCol6
                Case 7: TamCol = TamCol7
                Case 8: TamCol = TamCol8
                Case 9: TamCol = TamCol9
                Case 10: TamCol = TamCol10
                Case 11: TamCol = TamCol11
                Case 12: TamCol = TamCol12
                Case 13: TamCol = TamCol13
                Case 14: TamCol = TamCol14
                Case 15: TamCol = TamCol15
                Case 16: TamCol = TamCol16
                Case 17: TamCol = TamCol17
                Case 18: TamCol = TamCol18
                Case 19: TamCol = TamCol19
                Case 20: TamCol = TamCol20
                Case 21: TamCol = TamCol21
                Case 22: TamCol = TamCol22
                Case 23: TamCol = TamCol23
                Case 24: TamCol = TamCol24
                Case 25: TamCol = TamCol25
                Case 26: TamCol = TamCol26
                Case 27: TamCol = TamCol27
                Case 28: TamCol = TamCol28
                Case 29: TamCol = TamCol29
                Case 30: TamCol = TamCol30
                Case 31: TamCol = TamCol31
            End Select
            .ColumnHeaders(Contador).Width = TamCol
            Contador = Contador + 1
        Loop
    End If
    TBAbrir.Close
End With

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcGravarTotaisPC(ID As Long)
On Error GoTo tratar_erro

'PRODUTOS
TotalProduto = 0
valor = 0
SumIPI = 0
TotalICMSCST = 0
BASECALCULO = 0
TotalICMS = 0
TotalBCICMSCST = 0
Valor1 = 0
Valor2 = 0
Valor3 = 0

'Foi necessario dar while porque a funÁ„o round n„o considerava a formataÁ„o de 10 casas do valor unitario, alguns casos o valor ficava maior ou menor
Set TBTotaisnota = CreateObject("adodb.recordset")
TBTotaisnota.Open "Select preco_unitario, Quant_Comp from compras_pedido_lista where IDPedido = " & ID & " and Tipo = 'P' and Remessa = 'False' and status_item <> 'CANCELADO'", Conexao, adOpenKeyset, adLockReadOnly
Do While TBTotaisnota.EOF = False
    TotalProduto = TotalProduto + Format(TBTotaisnota!preco_unitario * TBTotaisnota!Quant_Comp, "0.00")
    TBTotaisnota.MoveNext
Loop

Set TBTotaisnota = CreateObject("adodb.recordset")
'TBTotaisnota.Open "Select Sum(preco_unitario) AS unit, Sum(Quant_Comp) AS quantcomp, Sum(preco_total) as Valor, Sum(vlripi) as SumIPI, Sum(Valor_ICMS_ST) as TotalICMSCST, Sum(BC_ICMS) as BASECALCULO, Sum(vlricms) as TotalICMS, Sum(BC_ICMS_ST) as TotalBCICMSCST, Sum(Frete) as SumFrete, Sum(Seguro) as SumSeguro, Sum(Acessorias) as SumAcessorias from compras_pedido_lista where IDPedido = " & ID & " and Tipo = 'P' and Remessa = 'False' and status_item <> 'CANCELADO'", Conexao, adOpenKeyset, adLockOptimistic
TBTotaisnota.Open "Select Sum(ValorDesconto * Quant_Comp) as TotalDesconto, Sum(preco_total) as Valor, Sum(vlripi) as SumIPI, Sum(Valor_ICMS_ST) as TotalICMSCST, Sum(BC_ICMS) as BASECALCULO, Sum(vlricms) as TotalICMS, Sum(BC_ICMS_ST) as TotalBCICMSCST, Sum(Frete) as SumFrete, Sum(Seguro) as SumSeguro, Sum(Acessorias) as SumAcessorias from compras_pedido_lista where IDPedido = " & ID & "  and Remessa = 'False' and status_item <> 'CANCELADO'", Conexao, adOpenKeyset, adLockOptimistic
If TBTotaisnota.EOF = False Then
'    If IsNull(TBTotaisnota!unit) = False And IsNull(TBTotaisnota!quantcomp) = False Then
'        TotalProduto = Format(TBTotaisnota!unit * TBTotaisnota!quantcomp, "###,##0.00")
'    Else
'        TotalProduto = 0
'    End If
    valor = IIf(IsNull(TBTotaisnota!valor), 0, Format(TBTotaisnota!valor, "###,##0.00"))
    SumIPI = IIf(IsNull(TBTotaisnota!SumIPI), 0, TBTotaisnota!SumIPI)
    TotalICMSCST = IIf(IsNull(TBTotaisnota!TotalICMSCST), 0, TBTotaisnota!TotalICMSCST)
    BASECALCULO = IIf(IsNull(TBTotaisnota!BASECALCULO), 0, TBTotaisnota!BASECALCULO)
    TotalICMS = IIf(IsNull(TBTotaisnota!TotalICMS), 0, TBTotaisnota!TotalICMS)
    TotalBCICMSCST = IIf(IsNull(TBTotaisnota!TotalBCICMSCST), 0, TBTotaisnota!TotalBCICMSCST)
    Valor1 = IIf(IsNull(TBTotaisnota!SumFrete), 0, TBTotaisnota!SumFrete)
    Valor2 = IIf(IsNull(TBTotaisnota!SumSeguro), 0, TBTotaisnota!SumSeguro)
    Valor3 = IIf(IsNull(TBTotaisnota!SumAcessorias), 0, TBTotaisnota!SumAcessorias)
    TotalDescontoProdutos = Round(TBTotaisnota!TotalDesconto, 2)
End If

'SERVI«OS
TotalServicos = 0
Qtde = 0
Set TBTotaisnota = CreateObject("adodb.recordset")
StrSql = "Select Sum(valordesconto * Quant_Comp)  as TotalDesconto,Sum(ROUND(preco_unitario * Quant_Comp, 2)) as TotalServicos, Sum(ROUND(preco_total, 2)) as Valor1 from compras_pedido_lista where IDPedido = " & ID & " and Tipo = 'S' and status_item <> 'CANCELADO'"
'Debug.print StrSql

TBTotaisnota.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
If TBTotaisnota.EOF = False Then
    TotalServicos = IIf(IsNull(TBTotaisnota!TotalServicos), 0, Format(TBTotaisnota!TotalServicos, "###,##0.00"))
    Qtde = IIf(IsNull(TBTotaisnota!Valor1), 0, Format(TBTotaisnota!Valor1, "###,##0.00"))
    TotalDescontoServico = IIf(IsNull(TBTotaisnota!TotalDesconto), 0, Format(TBTotaisnota!TotalDesconto, "###,##0.00"))
End If
TBTotaisnota.Close

TotalDesconto = IIf(IsNull(TotalDescontoProdutos) = True, 0, TotalDescontoProdutos) + IIf(IsNull(TotalDescontoServico) = True, 0, TotalDescontoServico)

Set TBCarteira = CreateObject("adodb.recordset")
TBCarteira.Open "Select * from compras_pedido where IDPedido = " & ID, Conexao, adOpenKeyset, adLockOptimistic
If TBCarteira.EOF = True Then TBCarteira.AddNew
TBCarteira!dbl_Base_ICMS = Format(BASECALCULO, "###,##0.00")
TBCarteira!dbl_Valor_ICMS = Format(TotalICMS, "###,##0.00")
TBCarteira!dbl_Base_ICMS_Subst = Format(TotalBCICMSCST, "###,##0.00")
TBCarteira!dbl_Valor_ICMS_Subst = Format(TotalICMSCST, "###,##0.00")
TBCarteira!dbl_Valor_Total_Produtos = Format(TotalProduto, "###,##0.00")
TBCarteira!dbl_valor_total_servicos = Format(TotalServicos, "###,##0.00")
TBCarteira!TotalDesconto = Format(TotalDesconto, "###,##0.0000000000")
If TBCarteira!TotalDesconto <= 0 Then TBCarteira!TotalDesconto = 0
TBCarteira!dbl_Valor_Total_IPI = Format(SumIPI, "###,##0.00")

SubTotal = Format(TotalProduto + TotalServicos, "###,##0.00")
TBCarteira!SubTotal = Format(SubTotal, "###,##0.00")

TBCarteira!Total_Frete = Format(Valor1, "###,##0.00")
TBCarteira!Total_Seguro = Format(Valor2, "###,##0.00")
TBCarteira!Total_Acessorias = Format(Valor3, "###,##0.00")
TBCarteira!dbl_valor_total = Format((SubTotal - TotalDesconto) + SumIPI + TotalICMSCST + Valor1 + Valor2 + Valor3, "###,##0.00")
TBCarteira.Update
TBCarteira.Close

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Function FunDefinirPrazoAntecipacao(Dias_antec As Integer, PrazoFinal As Date) As Date
On Error GoTo tratar_erro

'Verifica quantidade de sabados, domingos e feridos
Inicio = PrazoFinal - Dias_antec

Inicio:
    If Weekday(Inicio) = 1 Then
        Inicio = Inicio - 2
    ElseIf Weekday(Inicio) = 7 Then
            Inicio = Inicio - 1
        ElseIf FunVerificaFeriado(Inicio) = True Then
                Inicio = Inicio - 1
                GoTo Inicio
    End If
    FunDefinirPrazoAntecipacao = Inicio

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Sub ProcGavarPCJurosMulta(IDConta As Long, IDnota As Long, ValorJuros As Double, ValorMulta As Double, TipoConta As String, Pago_recebido As Boolean)
On Error GoTo tratar_erro

Set TBAliquota = CreateObject("adodb.recordset")
TBAliquota.Open "Select int_codfamilia from tbl_familia where Juros = 'True' and Multa = 'True' and Destino = '" & TipoConta & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBAliquota.EOF = False Then
    ProcEnviaDadosPCJurosMulta IDConta, IDnota, TBAliquota!int_codfamilia, Format(ValorJuros + ValorMulta, "###,##0.00"), TipoConta, Pago_recebido
Else
    Set TBAliquota = CreateObject("adodb.recordset")
    TBAliquota.Open "Select int_codfamilia from tbl_familia where Juros = 'True' and Destino = '" & TipoConta & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBAliquota.EOF = False Then
        ProcEnviaDadosPCJurosMulta IDConta, IDnota, TBAliquota!int_codfamilia, Format(ValorJuros, "###,##0.00"), TipoConta, Pago_recebido
    End If
    Set TBAliquota = CreateObject("adodb.recordset")
    TBAliquota.Open "Select int_codfamilia from tbl_familia where Multa = 'True' and Destino = '" & TipoConta & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBAliquota.EOF = False Then
        ProcEnviaDadosPCJurosMulta IDConta, IDnota, TBAliquota!int_codfamilia, Format(ValorMulta, "###,##0.00"), TipoConta, Pago_recebido
    End If
End If
TBAliquota.Close

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcEnviaDadosPCJurosMulta(IDConta As Long, IDnota As Long, IDPC As Long, valor As Double, TipoConta As String, Pago_recebido As Boolean)
On Error GoTo tratar_erro

If valor > 0 Then
    Set TBTempo = CreateObject("adodb.recordset")
    TBTempo.Open "Select * from Familia_financeiro where IDConta = " & IDConta & " and ID_PC = " & IDPC & " and TipoConta = '" & TipoConta & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBTempo.EOF = True Then TBTempo.AddNew
    TBTempo!IDConta = IDConta
    TBTempo!IDnota = IDnota
    TBTempo!TipoConta = TipoConta
    TBTempo!valor = valor
    TBTempo!Pago_recebido = Pago_recebido
    TBTempo!ID_PC = IDPC
    TBTempo.Update
    TBTempo.Close
Else
    If Pago_recebido = False Then Conexao.Execute "Delete from Familia_financeiro where IDConta = " & IDConta & " and ID_PC = " & IDPC & " and TipoConta = '" & TipoConta & "' and Deposito_transf = 'False'"
End If

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Function FunVerifTipoFiltroIMF(Inicio As OptionButton, Meio As OptionButton, FIM As OptionButton, Igual As OptionButton, Texto As String) As String
On Error GoTo tratar_erro

If Inicio.Value = True Then FunVerifTipoFiltroIMF = " like '" & Texto & "%'"
If Meio.Value = True Then FunVerifTipoFiltroIMF = " like '%" & Texto & "%'"
If FIM.Value = True Then FunVerifTipoFiltroIMF = " like '%" & Texto & "'"
If Igual.Value = True Then FunVerifTipoFiltroIMF = " = '" & Texto & "'"


 Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Function FunVerifTipoFiltroIMFRel(Inicio As OptionButton, Meio As OptionButton, FIM As OptionButton, Igual As OptionButton, Texto As String) As String
On Error GoTo tratar_erro

If Inicio.Value = True Then FunVerifTipoFiltroIMFRel = " like '" & Texto & "*'"
If Meio.Value = True Then FunVerifTipoFiltroIMFRel = " like '*" & Texto & "*'"
If FIM.Value = True Then FunVerifTipoFiltroIMFRel = " like '*" & Texto & "'"
If Igual.Value = True Then FunVerifTipoFiltroIMFRel = " = '" & Texto & "'"

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Function FunVerifTipoFiltroIMFLista(Frase As String, Texto As String) As String
On Error GoTo tratar_erro

Select Case Frase
    Case "InÌcio": FunVerifTipoFiltroIMFLista = " like '" & Texto & "%'"
    Case "Meio": FunVerifTipoFiltroIMFLista = " like '%" & Texto & "%'"
    Case "Fim": FunVerifTipoFiltroIMFLista = " like '%" & Texto & "'"
    Case "Igual": FunVerifTipoFiltroIMFLista = " = '" & Texto & "'"
End Select

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Function FunVerifTipoFiltroIMFListaRel(Frase As String, Texto As String) As String
On Error GoTo tratar_erro

Select Case Frase
    Case "InÌcio": FunVerifTipoFiltroIMFListaRel = " like '" & Texto & "*'"
    Case "Meio": FunVerifTipoFiltroIMFListaRel = " like '*" & Texto & "*'"
    Case "Fim": FunVerifTipoFiltroIMFListaRel = " like '*" & Texto & "'"
    Case "Igual": FunVerifTipoFiltroIMFListaRel = " = '" & Texto & "'"
End Select

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Sub ProcFiltroPadrao(Combo_FIltro As ComboBox, Meio_Filtro As OptionButton, Fim_Filtro As OptionButton, Igual_Filtro As OptionButton, ID_Empresa_Filtro As Integer, Tipo_Filtro As String, Aplicacao_Filtro As String, Empresa_Filtro As Boolean)
On Error GoTo tratar_erro

Permitido = True
Set TBAliquota = CreateObject("adodb.recordset")
TBAliquota.Open "Select Filtrarpor, Frase from Empresa_Filtros where " & IIf(Empresa_Filtro = True, "ID_empresa = " & ID_Empresa_Filtro & " and ", "") & "Tipo = '" & Tipo_Filtro & "' and Aplicacao = '" & Aplicacao_Filtro & "' and Filtrarpor IS NOT NULL", Conexao, adOpenKeyset, adLockOptimistic
If TBAliquota.EOF = False Then
    If TBAliquota!filtrarpor <> "" Then Combo_FIltro = TBAliquota!filtrarpor
Prosseguir:
    If TBAliquota!Frase = "Meio" Then
        Meio_Filtro.Value = True
    ElseIf TBAliquota!Frase = "Fim" Then
            Fim_Filtro.Value = True
        ElseIf TBAliquota!Frase = "Igual" Then
            Igual_Filtro.Value = True
    End If
Else
    Permitido = False
End If
TBAliquota.Close


Exit Sub
tratar_erro:
    If Err.Number = 383 And Tipo_Filtro = "Produtos/ServiÁos" And Aplicacao_Filtro = "C" Then
        Combo_FIltro = "DescriÁ„o"
        GoTo Prosseguir
    End If
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Function FunVerifFusoHorario(CCE As Boolean) As String
On Error GoTo tratar_erro
Const TIME_ZONE_ID_DAYLIGHT As Long = 2
Dim tzi As TIME_ZONE_INFORMATION
Dim dwBias As Long
Dim sZone As String
Dim tmp As String

If CCE = True Then
    Select Case GetTimeZoneInformation(tzi)
        Case TIME_ZONE_ID_DAYLIGHT
            dwBias = tzi.Bias + tzi.DaylightBias
            'sZone = " (" & Left$(tzi.DaylightName, 1) & "DT)"
        Case Else
            dwBias = tzi.Bias + tzi.StandardBias
            'sZone = " (" & Left$(tzi.StandardName, 1) & "ST)"
    End Select
Else
    dwBias = tzi.Bias + tzi.StandardBias
End If
    
'tmp = "  " & Right$("00" & CStr(dwBias \ 60), 2) & Right$("00" & CStr(dwBias Mod 60), 2) & sZone
tmp = "  " & Right$("00" & CStr(dwBias \ 60), 2) & Right$("00" & CStr(dwBias Mod 60), 2)
If dwBias > 0 Then Mid$(tmp, 2, 1) = "-" Else Mid$(tmp, 2, 2) = "+0"
'FunVerifFusoHorario = Format$(Now, "ddd, dd mmm yyyy Hh:Mm:Ss") & tmp
FunVerifFusoHorario = Mid(tmp, 2, 1) & Mid(tmp, 3, 2) & ":" & Mid(tmp, 5, 2)

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Public Function CalculaDV(chave43 As String)
  Dim indice, multiplicador As Integer
  Dim soma, Resto, digito_verificador As Integer
  
  'Zera a soma
  soma = 0
  'Multiplicador inicia com 9
  multiplicador = 2
  
  'Multiplica do 43∞ atÈ o 1∞ caractere da chave
  For indice = Len(chave43) To 1 Step -1
  'Multiplica cada digito da chave pelo multiplicador correspondente e soma
  soma = soma + (Mid(chave43, indice, 1) * multiplicador)
  multiplicador = multiplicador + 1
 
  'Se multiplicador chegou a 2, volta para 9
  If (multiplicador > 9) Then multiplicador = 2
  Next indice
 
  'Pega o resto da divis„o atravÈs da funÁ„o mod
  Resto = soma Mod 11
  
  'DÌgito verificador È o resultado da subtraÁ„o 11 - resto
  digito_verificador = 11 - Resto
  
  'Testa se o DV È maior = 10
  If (digito_verificador >= 10) Then digito_verificador = 0
  
  'Retorna o DV
  CalculaDV = Abs(digito_verificador)
End Function


Sub ProcCarregaCaminhoNomeArquivo(CommonDialog1 As CommonDialog, extensao As String, ExtensaoPermitida As String)
On Error GoTo tratar_erro

caminho = ""
Nome_anexo = ""
With CommonDialog1
    .filename = ""
    .Filter = ExtensaoPermitida
    .InitDir = App.Path
    .DefaultExt = extensao
    .ShowOpen
    caminho = .filename
    If caminho <> "" Then Nome_anexo = .FileTitle
End With

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub procCarregaCertificado(certificadoBOX As ComboBox)
On Error GoTo tratar_erro
Dim i As Integer
Dim vetor_ As Variant

'Utiliza MÈtodo do Componente para Listar Certificados instalado no SO
'Set spdNFe = New NFeX.spdNFeSCAN
'vetor_ = Split(spdNFe.ListarCertificados, "|")
'certificadoBOX.Clear
'
'For i = LBound(vetor_) To UBound(vetor_)
'    certificadoBOX.AddItem vetor_(i)
'Next
    
Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Function FunTextoMaiusculoDigitar(CaixaTexto As TextBox) As Boolean
On Error GoTo tratar_erro

With CaixaTexto
    If .SelStart > 0 Then x = .SelStart
    .Text = UCase(.Text)
    .SelStart = x
End With

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Function

Sub ProcVerifImpostosEmpresa(ID_empresa As Integer, retorno As Boolean, CodIntServ As String, ServExecCliente As Boolean, Valor_total As Double, Serv As Boolean, TabelaSN As Integer, VlrTotalFat12UltMesesSomado As Double)
On Error GoTo tratar_erro

DAS = 0
'Prod
PIS_Prod = 0
Cofins_Prod = 0
CSLL_Prod = 0
IRPJ_Prod = 0
ICMS_SN = 0
IPI_SN = 0
'Serv
PIS_Serv = 0
Cofins_Serv = 0
CSLL_Serv = 0
ISS_Serv = 0
IRPJ_Serv = 0
cpp = 0

Regime = 0
Set TBFIltro = CreateObject("adodb.recordset")
TBFIltro.Open "Select * from Empresa where Codigo = " & ID_empresa, Conexao, adOpenKeyset, adLockOptimistic
If TBFIltro.EOF = False Then
    If TBFIltro!Simples = True Then Regime = 1
    If TBFIltro!Presumido = True Then Regime = 2
    If TBFIltro!Real = True Then Regime = 3
    If TBFIltro!Simples1 = True Then Regime = 4
        
    If DestacaImpostos = "SIM" And retorno = False Then
        Set TBFI = CreateObject("adodb.recordset")
        If Regime = 1 Then
            ProcVerifImpostosSN ID_empresa, TabelaSN, VlrTotalFat12UltMesesSomado
        Else
            TBFI.Open "Select * from Impostos where Regime = " & Regime & " and ID_empresa = " & ID_empresa, Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False Then
                'Prod
                If TemPIS = True Then PIS_Prod = IIf(IsNull(TBFI!PIS_produtos), 0, TBFI!PIS_produtos)
                If TemCOFINS = True Then Cofins_Prod = IIf(IsNull(TBFI!Cofins_produtos), 0, TBFI!Cofins_produtos)
                CSLL_Prod = IIf(IsNull(TBFI!CSLL_produtos), 0, TBFI!CSLL_produtos)
                IRPJ_Prod = IIf(IsNull(TBFI!IRPJ_produtos), 0, TBFI!IRPJ_produtos)
                
                'Serv
                If Serv = True Then
                    PIS_Serv = IIf(IsNull(TBFI!PIS), 0, TBFI!PIS)
                    Cofins_Serv = IIf(IsNull(TBFI!Cofins), 0, TBFI!Cofins)
                    CSLL_Serv = IIf(IsNull(TBFI!CSLL), 0, TBFI!CSLL)
                    ISS_Serv = IIf(IsNull(TBFI!ISS), 0, TBFI!ISS)
                    If Valor_total > 29 Then
                        INSS_Serv = IIf(IsNull(TBFI!INSS), 0, TBFI!INSS)
                        If ServExecCliente = True Then
                            Set TBItem = CreateObject("adodb.recordset")
                            TBItem.Open "Select INSS from projproduto where desenho = '" & CodIntServ & "' and Servico_cliente = 'True'", Conexao, adOpenKeyset, adLockOptimistic
                            If TBItem.EOF = False Then
                                INSS_Serv = IIf(IsNull(TBItem!INSS), 0, TBItem!INSS)
                            End If
                            TBItem.Close
                        End If
                    End If
                    IRPJ_Serv = IIf(IsNull(TBFI!IRPJ_servicos), 0, TBFI!IRPJ_servicos)
                    IRRF_Serv = IIf(IsNull(TBFI!IRRF), 0, TBFI!IRRF)
                End If
            End If
            TBFI.Close
        End If
    End If
End If
TBFIltro.Close

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcVerifImpostosSN(ID_empresa As Integer, TabelaSN As Integer, VlrTotalFat12UltMesesSomado As Double)
On Error GoTo tratar_erro
Dim valorSimples As Double

If VlrTotalFat12UltMesesSomado = 0 Then ValorTotalPago = FunVerifVlrTotalFat12UltMesesSomado(ID_empresa) Else ValorTotalPago = VlrTotalFat12UltMesesSomado
NovoValor = Replace(ValorTotalPago, ",", ".")
Set TBFI = CreateObject("adodb.recordset")
TBFI.Open "Select IT.* from Impostos_TabelaDAS IT INNER JOIN empresa E ON E.Codigo = IT.ID_empresa where E.Codigo = " & ID_empresa & " and E.Simples = 'True' and IT.Tabela = " & TabelaSN & " and IT.De <= " & NovoValor & " and IT.Ate >= " & NovoValor, Conexao, adOpenKeyset, adLockOptimistic
If TBFI.EOF = True Then
    Set TBFI = CreateObject("adodb.recordset")
    StrSql = "Select IT.* from Impostos_TabelaDAS IT INNER JOIN empresa E ON E.Codigo = IT.ID_empresa where E.Codigo = " & ID_empresa & " and E.Simples = 'True' and IT.Tabela = " & TabelaSN & " order by IT.Ate desc"
    'Debug.print StrSql
    
    TBFI.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
End If
If TBFI.EOF = False Then
    valorSimples = Format((ValorTotalPago * TBFI!DAS) / 100, "0.0000")
    ValorConta = IIf(IsNull(TBFI!Valor_deduzir), 0, TBFI!Valor_deduzir)
    
    If ValorConta > 0 Then
        Valor2 = ValorConta
        ValorConta = Format(valorSimples - ValorConta, "0.0000")
        DAS = Format((ValorConta / ValorTotalPago) * 100, "0.0000")
    Else
        DAS = IIf(IsNull(TBFI!DAS), 0, TBFI!DAS)
    End If
    Var = "De: " & Format(TBFI!De, "###,##0.00") & " atÈ: " & Format(TBFI!Ate, "###,##0.00")
    
    'Prod
    PIS_Prod = Format((DAS * IIf(IsNull(TBFI!PIS), 0, TBFI!PIS)) / 100, "0.0000")
    Cofins_Prod = Format((DAS * IIf(IsNull(TBFI!Cofins), 0, TBFI!Cofins)) / 100, "0.0000")
    CSLL_Prod = Format((DAS * IIf(IsNull(TBFI!CSLL), 0, TBFI!CSLL)) / 100, "0.0000")
    IRPJ_Prod = Format((DAS * IIf(IsNull(TBFI!IRPJ), 0, TBFI!IRPJ)) / 100, "0.0000")
    CPP_Prod = Format((DAS * IIf(IsNull(TBFI!cpp), 0, TBFI!cpp)) / 100, "0.0000")
    IPI_SN = Format((DAS * IIf(IsNull(TBFI!IPI), 0, TBFI!IPI)) / 100, "0.0000")
    ICMS_SN = Format((DAS * IIf(IsNull(TBFI!ICMS), 0, TBFI!ICMS)) / 100, "0.0000")
'    AliquotaSN = Format((DAS * IIf(IsNull(TBFI!ICMS), 0, TBFI!ICMS)) / 100, "0.0000")
    
    'Serv
    PIS_Serv = IIf(IsNull(TBFI!PIS), 0, TBFI!PIS)
    Cofins_Serv = IIf(IsNull(TBFI!Cofins), 0, TBFI!Cofins)
    CSLL_Serv = IIf(IsNull(TBFI!CSLL), 0, TBFI!CSLL)
    IRPJ_Serv = IIf(IsNull(TBFI!IRPJ), 0, TBFI!IRPJ)
    CPP_Serv = IIf(IsNull(TBFI!cpp), 0, TBFI!cpp)
    ISS_Serv = Format((DAS * IIf(IsNull(TBFI!ISS), 0, TBFI!ISS)) / 100, "0.00")
    
    If ValorTotalPago > 3600000 Then
        'Calculo de ISS ou ICMS do simples quando chega na faixa 6 e o valor anual(Jan a Dezembro) ainda n„o passou 3.600.000,00
        Set TBControleNF = CreateObject("adodb.recordset")
        TBControleNF.Open "Select IT.DAS, IT.Valor_deduzir, IT.ISS, IT.ICMS from Impostos_TabelaDAS IT INNER JOIN empresa E ON E.Codigo = IT.ID_empresa where E.Codigo = " & ID_empresa & " and E.Simples = 'True' and IT.Tabela = " & TabelaSN & " AND IT.De <= 1800000.01 and IT.Ate >= 3600000 and IT.DAS IS NOT NULL and IT.Valor_deduzir IS NOT NULL", Conexao, adOpenKeyset, adLockReadOnly
        If TBControleNF.EOF = False Then
            Valor_Retencao_Servico = (ValorTotalPago * TBControleNF!DAS) / 100
            ValorConta = Format(Valor_Retencao_Servico - TBControleNF!Valor_deduzir, "0.00")
            Valor_CSLL_Serv = (ValorConta / ValorTotalPago) * 100
            If TBControleNF!ICMS <> 0 And IsNull(TBControleNF!ICMS) = False Then
                ICMS_SN = Format((Valor_CSLL_Serv * TBControleNF!ICMS) / 100, "0.00")
            Else
                ISS_Serv = Format((Valor_CSLL_Serv * TBControleNF!ISS) / 100, "0.00")
            End If
        End If
    End If
    
    If ISS_Serv > 5 Then
        valorSimples = ISS_Serv - 5
        ISS_Serv = 5
        Dim PIS_diluido As Double
        Dim Cofins_diluido As Double
        Dim CSLL_diluido As Double
        Dim IRPJ_diluido As Double
        Dim CPP_diluido As Double
        
        If TabelaSN = 3 Then
            IRPJ_diluido = 6.02
            CSLL_diluido = 5.26
            Cofins_diluido = 19.28
            PIS_diluido = 4.18
            CPP_diluido = 65.26
        ElseIf TabelaSN = 4 Then
            IRPJ_diluido = 31.33
            CSLL_diluido = 32
            Cofins_diluido = 30.13
            PIS_diluido = 6.54
            CPP_diluido = 0
        Else
            IRPJ_diluido = 30.07
            CSLL_diluido = 16.34
            Cofins_diluido = 18.43
            PIS_diluido = 3.99
            CPP_diluido = 31.17
        End If
        PIS_Serv = Format(PIS_Serv + ((valorSimples * PIS_diluido) / 100), "0.00")
        Cofins_Serv = Format(Cofins_Serv + ((valorSimples * Cofins_diluido) / 100), "0.00")
        CSLL_Serv = Format(CSLL_Serv + ((valorSimples * CSLL_diluido) / 100), "0.00")
        IRPJ_Serv = Format(IRPJ_Serv + ((valorSimples * IRPJ_diluido) / 100), "0.00")
        CPP_Serv = Format(CPP_Serv + ((valorSimples * CPP_diluido) / 100), "0.00")
    End If
End If
TBFI.Close

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Function FunVerifIDEmpresaSite(ID_empresa As Integer) As Long
On Error GoTo tratar_erro

Set TBFIltro = CreateObject("adodb.recordset")
TBFIltro.Open "Select CNPJ from Empresa where Codigo = " & ID_empresa, Conexao, adOpenKeyset, adLockOptimistic
If TBFIltro.EOF = False Then
    FunAbreBDSite
    Set TBMySQL = New ADODB.Recordset
    TBMySQL.Open "Select * From Clientes Where Cnpj = '" & TBFIltro!CNPJ & "'", ConexaoMySql, adOpenKeyset, adLockOptimistic
    If TBMySQL.EOF = False Then
        FunVerifIDEmpresaSite = TBMySQL!ID
    End If
    TBMySQL.Close
End If
TBFIltro.Close

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Public Function FunVerifVlrUnitUltCompra(Codinterno As String, IDlista As Long) As String
On Error GoTo tratar_erro

FunVerifVlrUnitUltCompra = ""
If IDlista > 0 Then TextoFiltro = "and IDlista < " & IDlista & "" Else TextoFiltro = ""
Set TBTempo = CreateObject("adodb.recordset")
TBTempo.Open "Select preco_unitario from Compras_pedido_lista where Desenho = '" & Codinterno & "' " & TextoFiltro & " and IDpedido IS NOT NULL and IDpedido <> 0 and Status_item <> 'CANCELADO' order by IDlista desc", Conexao, adOpenKeyset, adLockOptimistic
If TBTempo.EOF = False Then
    FunVerifVlrUnitUltCompra = Format(TBTempo!preco_unitario, "###,##0.0000000000")
End If
TBTempo.Close

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Public Function FunGerarVersaoEstrutura(UltVersao As String) As String
On Error GoTo tratar_erro

Select Case UltVersao
    Case "A": FunGerarVersaoEstrutura = "B"
    Case "B": FunGerarVersaoEstrutura = "C"
    Case "C": FunGerarVersaoEstrutura = "D"
    Case "D": FunGerarVersaoEstrutura = "E"
    Case "E": FunGerarVersaoEstrutura = "F"
    Case "F": FunGerarVersaoEstrutura = "G"
    Case "G": FunGerarVersaoEstrutura = "H"
    Case "H": FunGerarVersaoEstrutura = "I"
    Case "I": FunGerarVersaoEstrutura = "J"
    Case "J": FunGerarVersaoEstrutura = "K"
    Case "K": FunGerarVersaoEstrutura = "L"
    Case "L": FunGerarVersaoEstrutura = "M"
    Case "M": FunGerarVersaoEstrutura = "N"
    Case "N": FunGerarVersaoEstrutura = "O"
    Case "O": FunGerarVersaoEstrutura = "P"
    Case "P": FunGerarVersaoEstrutura = "Q"
    Case "Q": FunGerarVersaoEstrutura = "R"
    Case "R": FunGerarVersaoEstrutura = "S"
    Case "S": FunGerarVersaoEstrutura = "T"
    Case "T": FunGerarVersaoEstrutura = "U"
    Case "U": FunGerarVersaoEstrutura = "V"
    Case "V": FunGerarVersaoEstrutura = "W"
    Case "W": FunGerarVersaoEstrutura = "X"
    Case "X": FunGerarVersaoEstrutura = "Y"
    Case "Y": FunGerarVersaoEstrutura = "Z"
End Select

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Public Function FunValidarEstrutura(Codproduto As Long, DtValidacao As String, RespValidacao As String, versao As String)
On Error GoTo tratar_erro

Set TBNivel2 = CreateObject("adodb.recordset")
TBNivel2.Open "Select PP.Codproduto, PP.desenho from projconjunto PC LEFT JOIN Projproduto PP on PC.desenho = PP.desenho where PC.codproduto = " & Codproduto & " and PP.Subtipoitem <> '0' and PC.Versao = '" & versao & "' or PC.Codproduto IS NULL order by PC.Codproduto", Conexao, adOpenKeyset, adLockOptimistic
If TBNivel2.EOF = False Then
    Do While TBNivel2.EOF = False
        If FunVerifTemEstrutura(TBNivel2!Desenho) = True Then
            ProcGravarValidEstruturaVersao TBNivel2!Codproduto, DtValidacao, RespValidacao, versao
            ProcValidarEngenhariaEstruturaProd TBNivel2!Codproduto
        End If
        TBNivel2.MoveNext
    Loop
End If
TBNivel2.Close

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Function FunCriarNovoNumeroProcesso() As String
On Error GoTo tratar_erro

Set TBTempo = CreateObject("adodb.recordset")
TBTempo.Open "Select Nprocesso from Processos where Nprocesso IS NOT NULL and Nprocesso <> N'' and Year (DtImplantacao) = '" & Year(Date) & "' order by Ordenarprocesso desc", Conexao, adOpenKeyset, adLockOptimistic
If TBTempo.EOF = False Then
    Numero = Left(TBTempo!Nprocesso, Len(TBTempo!Nprocesso) - 3)
    Numero = Right(Numero, 5) + 1
Else
    Numero = 1
End If
TBTempo.Close

a = Numero
Ano = Right(Year(Date), 2)
Select Case Len(a)
    Case 1: FunCriarNovoNumeroProcesso = "PRO-0000" & Numero & "/" & Ano
    Case 2: FunCriarNovoNumeroProcesso = "PRO-000" & Numero & "/" & Ano
    Case 3: FunCriarNovoNumeroProcesso = "PRO-00" & Numero & "/" & Ano
    Case 4: FunCriarNovoNumeroProcesso = "PRO-0" & Numero & "/" & Ano
    Case 5: FunCriarNovoNumeroProcesso = "PRO-" & Numero & "/" & Ano
End Select

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Function FunVerifiCodRefUtilizado(Codinterno As String, CodRef As String) As Boolean
On Error GoTo tratar_erro
    
FunVerifiCodRefUtilizado = False
Set TBTempo = CreateObject("adodb.recordset")
TBTempo.Open "Select P.Codproduto from Projproduto P INNER JOIN item_aplicacoes IA ON IA.Codproduto = P.Codproduto where P.Desenho <> '" & Codinterno & "' and IA.n_referencia = '" & CodRef & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBTempo.EOF = False Then
    USMsgBox ("N„o È permitido utilizar este cÛdigo de referÍncia, pois o mesmo pertence a outro produto."), vbExclamation, "CAPRIND v5.0"
    FunVerifiCodRefUtilizado = True
    TBTempo.Close
    Exit Function
End If
TBTempo.Close

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Function FunVerifCotacaoValida(ID_empresa As Integer, Codinterno As String, Produto As Boolean, MostrarMsg As Boolean, Acao As String, IDforn As Long) As Boolean
On Error GoTo tratar_erro

FunVerifCotacaoValida = True
Set TBTempo = CreateObject("adodb.recordset")
TBTempo.Open "Select Codigo from Empresa where Codigo = " & ID_empresa & " and Bloq_compra_cot_valida = 'True'", Conexao, adOpenKeyset, adLockOptimistic
If TBTempo.EOF = False Then
    Set TBFI = CreateObject("adodb.recordset")
    TBFI.Open "Select CC.ID_cotacao from (Cotacao_item CI INNER JOIN Compras_Cotacao CC ON CC.ID_cotacao = CI.idcot) INNER JOIN Cotacao_fornecedor CF ON CF.IDitem = CI.ID where CC.ID_empresa = " & ID_empresa & " and CI.coditem = '" & Codinterno & "' and CF.IDforn = " & IDforn & " and CF.aprovadoforn = 1 and CC.statuscotacao = 'APROVADA' and CC.Data_validade >= '" & Format(Date, "Short Date") & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBFI.EOF = True Then
        If MostrarMsg = True Then USMsgBox ("N„o È permitido " & Acao & " este " & IIf(Produto = True, "produto", "serviÁo") & ", pois o mesmo n„o tem uma cotaÁ„o valida para o fornecedor deste pedido."), vbExclamation, "CAPRIND v5.0"
        TBFI.Close
        FunVerifCotacaoValida = False
        Exit Function
    End If
    TBFI.Close
End If
TBTempo.Close

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Function FunAlterarProdSimiliarOrdem(ID_empresa As Integer, Codint_similiar As String, Ordem As Long, Codint As String, Qtde_similar As Double, Qtde_similarPC As Double, Req_PI As Boolean) As Boolean
On Error GoTo tratar_erro

FunAlterarProdSimiliarOrdem = False
If Req_PI = True Then
    CamposFiltro = "VC.Qtde_produzir"
    INNERJOINTEXTO = "INNER JOIN vendas_carteira VC ON VC.Codigo = PM.ID_carteira"
    TextoFiltro = "ID_carteira"
Else
    CamposFiltro = "PR.Quant"
    INNERJOINTEXTO = "INNER JOIN Producao PR ON PR.Ordem = PM.Ordem"
    TextoFiltro = "Ordem"
End If

If FunVerifProdSimiliar(ID_empresa) = True Then
    Set TBFIltro = CreateObject("adodb.recordset")
    TBFIltro.Open "Select Descricao, desenho, peso_metro, Largura, Comprimento, PBruto, un_kg from projproduto where desenho = '" & Codint_similiar & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBFIltro.EOF = False Then
        Set TBMaterial = CreateObject("adodb.recordset")
        TBMaterial.Open "Select PM.*, P.Un_kg, P.SubTipoItem, P.PBruto, " & CamposFiltro & " from Producaomaterial PM " & INNERJOINTEXTO & " INNER JOIN projproduto P ON PM.Codigo = P.Desenho where PM." & TextoFiltro & " = " & Ordem & " and P.Desenho = '" & Codint & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBMaterial.EOF = False Then
            Set TBGravar = CreateObject("adodb.recordset")
            TBGravar.Open "Select * from Producaomaterial where " & TextoFiltro & " = " & Ordem & " and Codigo = '" & TBFIltro!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBGravar.EOF = True Then
                TBGravar.AddNew
                valor = 0
                Valor1 = 0
            Else
                valor = TBGravar!Requisitado
                Valor1 = TBGravar!DimensaoTotal
            End If
            
            TBGravar!Posicao = TBMaterial!Posicao
            TBGravar!quantidade = TBMaterial!quantidade
            TBGravar!Unidade = TBMaterial!Unidade
            TBGravar!CODIGO = TBFIltro!Desenho
            TBGravar!Descricao = TBFIltro!Descricao
            TBGravar!Ordem = TBMaterial!Ordem
            TBGravar!ID_carteira = TBMaterial!ID_carteira
            TBGravar!PesoMetro = TBFIltro!peso_metro
            
            If TBFIltro!Un_Kg = "Mt≤" Then
                TBGravar!Dimensao = IIf(IsNull(TBFIltro!Comprimento), 0, TBFIltro!Comprimento) * IIf(IsNull(TBFIltro!Largura), 0, TBFIltro!Largura)
                TBGravar!DimensaoTotal = ((TBGravar!Dimensao / 1000) / 1000) * Qtde_similar
            Else
                TBGravar!Dimensao = TBMaterial!Dimensao
                TBGravar!DimensaoTotal = TBMaterial!DimensaoTotal
            End If
            
            'Peso bruto
            If IsNull(TBFIltro!PBruto) = False And TBFIltro!PBruto <> "0" Then
                TBGravar!pesounidade = TBFIltro!PBruto
            Else
                Select Case TBFIltro!Un_Kg
                    Case "Mt/L": TBGravar!pesounidade = Format(TBFIltro!peso_metro / 1000 * TBGravar!Dimensao, "###,##0.0000000000")
                    Case "PÁ": TBGravar!pesounidade = Format(TBFIltro!peso_metro, "###,##0.0000000000")
                    Case "Mt≤": TBGravar!pesounidade = Format(((TBGravar!Dimensao * TBFIltro!peso_metro) / 1000) / 1000, "###,##0.0000000000")
                    Case "N/a": TBGravar!pesounidade = Format(0, "###,##0.0000000000")
                End Select
            End If
            
            TBGravar!PesoTotal = Format(TBGravar!pesounidade * (TBMaterial!quantidade / TBMaterial!Quant), "###,##0.0000000000")
            TBGravar!Requisitado = Format(Qtde_similar, "###,##0.0000")
            TBGravar!Total_pc = Qtde_similarPC
            TBGravar!versao = TBMaterial!versao
            TBGravar!Saida = "N√O"
            TBGravar.Update
                                
            If Qtde_similar < (TBMaterial!Requisitado + valor) Then
                TBMaterial!Requisitado = Format((TBMaterial!Requisitado + valor) - Qtde_similar, "###,##0.0000")
                If TBFIltro!Un_Kg = "Mt≤" Then TBMaterial!DimensaoTotal = Format((TBMaterial!DimensaoTotal + Valor1) - TBGravar!DimensaoTotal, "###,##0.0000")
                
                If TBMaterial!Unidade = "KG" Or TBMaterial!SubTipoItem = 1 Or TBMaterial!SubTipoItem = 2 Or TBMaterial!SubTipoItem = 3 Then
                    If TBMaterial!Unidade = "KG" And (TBMaterial!Un_Kg = "Mt≤" Or TBMaterial!Un_Kg = "Mt/L") Then
                        If IsNull(TBMaterial!PBruto) = False And TBMaterial!PBruto > 0 And TBMaterial!PBruto <> "" Then TBMaterial!Total_pc = Format(TBMaterial!Requisitado / TBMaterial!PBruto, "###,##0.0000") Else TBMaterial!Total_pc = Null
                    Else
                        If TBMaterial!SubTipoItem = 1 Or TBMaterial!SubTipoItem = 2 Or TBMaterial!SubTipoItem = 3 Then TBMaterial!Total_pc = TBMaterial!Requisitado Else TBMaterial!Total_pc = Null
                    End If
                    TBMaterial.Update
                End If
            Else
                Conexao.Execute "DELETE Producaomaterial where Idmateriaprima = " & TBMaterial!IdMateriaPrima
            End If
            
            TBGravar.Close
        End If
        
        FunAlterarProdSimiliarOrdem = True
    End If
    TBFIltro.Close
End If

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Public Sub ProcCalculaQtdesSimilar(Codint_similar As String, Qtde_similar As TextBox, Qtde_similarPC As TextBox)
On Error GoTo tratar_erro

'Calcula quantidade P« ou quantidade
Set TBMaterial = CreateObject("adodb.recordset")
TBMaterial.Open "select un_kg, Unidade, SubTipoItem, PBruto from projproduto where desenho = '" & Codint_similar & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBMaterial.EOF = False Then
    Valor3 = IIf(TBMaterial!PBruto = "", 0, TBMaterial!PBruto)
    If TBMaterial!Unidade = "KG" Or TBMaterial!SubTipoItem = 1 Or TBMaterial!SubTipoItem = 2 Or TBMaterial!SubTipoItem = 3 Then
        If TBMaterial!Unidade = "KG" And (TBMaterial!Un_Kg = "Mt≤" Or TBMaterial!Un_Kg = "Mt/L") Then
            If Valor3 > 0 Then
                If Qtde_similar.Locked = False Then
                    Qtde_similarPC = 0
                Else
                    If Qtde_similarPC <> "" Then Qtde_similar = Format(Qtde_similarPC * Valor3, "###,##0.0000") Else Qtde_similar = 0
                End If
            Else
                If Qtde_similar.Locked = False Then Qtde_similarPC = 0 Else Qtde_similar = 0
            End If
        Else
            If TBMaterial!SubTipoItem = 1 Or TBMaterial!SubTipoItem = 2 Or TBMaterial!SubTipoItem = 3 Then
                If Qtde_similar.Locked = False Then
                    Qtde_similarPC = 0
                Else
                    Qtde_similar = IIf(Qtde_similarPC = "", 0, Qtde_similarPC)
                End If
            Else
                If Qtde_similar.Locked = False Then Qtde_similarPC = 0 Else Qtde_similar = 0
            End If
        End If
    End If
End If
TBMaterial.Close

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcAcertaRequisicao(Desenho As String, Ordem As Long, IDcarteira As Long, Versao_estrutura As String, QuantSolicitado1 As Double, EmpenhoVendas As Boolean)
On Error GoTo tratar_erro

'Verifica vers„o do material adicionado na ordem
Versao_estruturaTexto = ""
Set TBMaterial = CreateObject("adodb.recordset")
TBMaterial.Open "Select Versao from producaomaterial where " & IIf(Ordem <> 0, "Ordem = " & Ordem, "ID_carteira = " & IDcarteira), Conexao, adOpenKeyset, adLockOptimistic
If TBMaterial.EOF = False Then
    Versao_estruturaTexto = TBMaterial!versao
End If

Set TBItem = CreateObject("adodb.recordset")
TBItem.Open "Select * from projproduto where desenho = '" & Desenho & "' and DtValidacaoConj IS NOT NULL", Conexao, adOpenKeyset, adLockOptimistic
If TBItem.EOF = False Then
StrSql = "Select PC.*, P.PBruto, P.SubTipoItem, P.Desenho, P.Producao, P.Compras from projconjunto PC INNER JOIN projproduto P ON P.Desenho = PC.Desenho where PC.codproduto = " & TBItem!Codproduto & " and PC.Versao = '" & IIf(Versao_estruturaTexto = "", Versao_estrutura, Versao_estruturaTexto) & "' and P.bloqueado = 'False' order by PC.Posicao, PC.codigo"
'Debug.print StrSql

    Set TBProcessos = CreateObject("adodb.recordset")
    TBProcessos.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
    If TBProcessos.EOF = False Then
        Do While TBProcessos.EOF = False
            
            Set TBMaterial = CreateObject("adodb.recordset")
            If TBProcessos!Unidade = "M≥" Then TextoFiltro = " and dimensao = " & Replace(TBProcessos!Dimensoes, ",", ".") Else TextoFiltro = ""
            TBMaterial.Open "Select * from producaomaterial where " & IIf(Ordem <> 0, "Ordem = " & Ordem, "ID_carteira = " & IDcarteira) & " and codigo = '" & TBProcessos!Desenho & "'" & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
            
            If TBMaterial.EOF = True Then
            TBMaterial.AddNew
            End If
            
            ProcEnviaDadosRequisicao Ordem, IDcarteira, QuantSolicitado1
            Peso = TBProcessos!quantidade
            If TBProcessos!Un_Kg <> "N/a" And TBProcessos!Un_Kg <> "" And (TBProcessos!Unidade = "KG" Or TBProcessos!Unidade = "MT" Or TBProcessos!Unidade = "MM" Or TBProcessos!Unidade = "M≥") Then
                Select Case TBProcessos!Unidade
                    Case "KG": Peso = TBProcessos!PesoTotal
                    Case "MT": Peso = (TBProcessos!Dimensoes / 1000) * TBProcessos!quantidade
                    Case "MM": Peso = TBProcessos!Dimensoes * TBProcessos!quantidade
                    Case "M≥": Peso = TBProcessos!PesoTotal
                End Select
            End If
            
            TBMaterial!Percentual_perda = TBProcessos!Percentual_perda
            
            If TBProcessos!Unidade = "M≥" Then
                TBMaterial!Requisitado = Peso * QuantSolicitado1
                TBMaterial!DimensaoTotal = TBProcessos!Dimensoes
                TBMaterial!Total_pc = TBProcessos!quantidade
            Else
                TBMaterial!Requisitado = Format(TBMaterial!Requisitado + (Peso * QuantSolicitado1), "###,##0.0000")
                If TBProcessos!Unidade = "KG" Or TBProcessos!SubTipoItem = 1 Or TBProcessos!SubTipoItem = 2 Or TBProcessos!SubTipoItem = 3 Then
                    If TBProcessos!Unidade = "KG" And (TBProcessos!Un_Kg = "Mt≤" Or TBProcessos!Un_Kg = "Mt/L") Then
                        If IsNull(TBProcessos!PBruto) = False And TBProcessos!PBruto > 0 And TBProcessos!PBruto <> "" Then TBMaterial!Total_pc = Format(TBMaterial!Requisitado / TBProcessos!PBruto, "###,##0.0000") Else TBMaterial!Total_pc = Null
                    Else
                        If TBProcessos!Unidade = "P«" Or TBProcessos!Unidade = "PC" Or TBProcessos!Unidade = "UN" Or TBProcessos!Unidade = "CJ" Then TBMaterial!Total_pc = TBMaterial!Requisitado Else TBMaterial!Total_pc = Null
                    End If
                End If
            End If
            'TBMaterial!Tipo_Item = TBProcessos!Tipo
            TBMaterial.Update
            
            If EmpenhoVendas = True Then
                Quant = QuantSolicitado1
                ProcEmpenhoPIN1 TBProcessos!Desenho, IDcarteira, Versao_estrutura, TBMaterial!Requisitado
            Else
                'Empenha do estoque (componentes que s„o comprados)
                If frmprod.MRP_considerar_estoque = True And TBProcessos!SubTipoItem = 3 And TBProcessos!Compras = True And TBProcessos!Producao = False Then
                    QuantSolicitado = TBMaterial!Requisitado
                    If frmprod.Consignacao = False Then Familiatext = "and Consignacao = 'False'" Else Familiatext = "and (Consignacao = 'False' or Consignacao = 'True' and id_cliente = " & IIf(frmprod.Txt_ID_cliente = "", 0, frmprod.Txt_ID_cliente) & " and Tipodest_NFcons = 'C' or Consignacao = 'True' and Tipodest_NFcons = 'F')"
                    Set TBEstoque = CreateObject("adodb.recordset")
                    TBEstoque.Open "Select IDestoque, Estoque_disponivel, Estoque_real_PC from Estoque_produtos where ID_empresa = " & TBCarteira!ID_empresa & " and Desenho = '" & TBMaterial!CODIGO & "' and Liberado = 'SIM' and Estoque_disponivel > 0 " & Familiatext & " order by Consignacao desc, Data, IDestoque", Conexao, adOpenKeyset, adLockOptimistic
                    If TBEstoque.EOF = False Then
                        Do While TBEstoque.EOF = False And QuantSolicitado > 0
                            ProcEmpenharREAutomOrdemMRP Ordem, TBEstoque!IDEstoque, TBEstoque!Estoque_Disponivel, TBMaterial!CODIGO
                            TBEstoque.MoveNext
                        Loop
                    End If
                End If
            End If
            
            TBMaterial.Close
            TBProcessos.MoveNext
        Loop
    End If
    TBProcessos.Close
End If
TBItem.Close

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviaDadosRequisicao(Ordem As Long, IDcarteira As Long, QuantSolicitado1 As Double)
On Error GoTo tratar_erro

TBMaterial!quantidade = TBMaterial!quantidade + (TBProcessos!quantidade * QuantSolicitado1)
TBMaterial!Unidade = TBProcessos!Unidade
TBMaterial!CODIGO = TBProcessos!Desenho
TBMaterial!ID_partnumber_fabricante = TBProcessos!ID_partnumber_fabricante
TBMaterial!Descricao = TBProcessos!Descricao
If Ordem <> 0 Then
    TBMaterial!Posicao = TBProcessos!Posicao
    TBMaterial!Ordem = Ordem
    TBMaterial!Saida = "N√O"
Else
    TBMaterial!ID_carteira = IDcarteira
End If
TBMaterial!PesoMetro = TBProcessos!PesoMetro
TBMaterial!pesounidade = TBProcessos!Peso
TBMaterial!PesoTotal = TBMaterial!PesoTotal + (TBProcessos!PesoTotal * QuantSolicitado1)
TBMaterial!Dimensao = TBProcessos!Dimensoes
If TBProcessos!Un_Kg = "Mt≤" Then TBMaterial!DimensaoTotal = ((TBProcessos!Dimensoes / 1000) / 1000) * TBMaterial!quantidade Else TBMaterial!DimensaoTotal = (TBProcessos!Dimensoes / 1000) * TBMaterial!quantidade
TBMaterial!versao = TBProcessos!versao
TBMaterial!Obs = TBProcessos!Obs
TBMaterial!Tipo_ITEM = 1

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcEmpenharREAutomOrdemMRP(Ordem As Long, IDEstoque As Long, Qtde_estoque As Double, Codinterno As String)
On Error GoTo tratar_erro

Set TBTempo = CreateObject("adodb.recordset")
TBTempo.Open "Select P.ID_empresa, PM.Ordem, PM.Requisitado, PM.Requisitado - (ISNULL(QSEP.Saida, 0) + ISNULL(QEPD.Qtde_empenhar, 0)) AS RequisitadoSaldo from ((Producao P INNER JOIN Producaomaterial PM ON PM.Ordem = P.Ordem) LEFT JOIN Qtde_saida_estoque_produto QSEP ON QSEP.Ordem = PM.Ordem and QSEP.Desenho = PM.Codigo) LEFT JOIN Qtde_empenhada_produto_detalhado QEPD ON QEPD.Ordem = PM.Ordem and QEPD.Codinterno = PM.Codigo where P.Ordem = " & Ordem & " and PM.Codigo = '" & Codinterno & "' and PM.Saida <> 'SIM' and PM.Requisitado - (ISNULL(QSEP.Saida, 0) + ISNULL(QEPD.Qtde_empenhar, 0)) > 0", Conexao, adOpenKeyset, adLockOptimistic
If TBTempo.EOF = False Then
    QuantSolicitado = TBTempo!RequisitadoSaldo
    Do While TBTempo.EOF = False And Qtde_estoque > 0 And QuantSolicitado > 0
        Set TBComponente = CreateObject("adodb.recordset")
        TBComponente.Open "Select * from Producao_NF_Consignada where Data = '" & Date & "' and Responsavel = '" & pubUsuario & "' and Ordem = " & TBTempo!Ordem & " and IDestoque = " & IDEstoque, Conexao, adOpenKeyset, adLockOptimistic
        If TBComponente.EOF = True Then TBComponente.AddNew
        TBComponente!Data = Date
        TBComponente!Responsavel = pubUsuario
        TBComponente!Ordem = Ordem
        TBComponente!Codinterno = Codinterno
        TBComponente!IDEstoque = IDEstoque
        If Qtde_estoque > TBTempo!RequisitadoSaldo Then
            TBComponente!quantidade = TBComponente!quantidade + TBTempo!RequisitadoSaldo
            If TBComponente!quantidade > TBTempo!Requisitado Then
                TBComponente!quantidade = TBTempo!Requisitado
                QuantSolicitado = 0
            Else
                QuantSolicitado = QuantSolicitado - TBTempo!RequisitadoSaldo
            End If
            Qtde_estoque = Qtde_estoque - TBTempo!RequisitadoSaldo
        Else
            TBComponente!quantidade = TBComponente!quantidade + Qtde_estoque
            If TBComponente!quantidade > TBTempo!Requisitado Then
                TBComponente!quantidade = TBTempo!Requisitado
                QuantSolicitado = 0
            Else
                QuantSolicitado = QuantSolicitado - Qtde_estoque
            End If
            Qtde_estoque = 0
        End If
        
        Set TBCFOP = CreateObject("adodb.recordset")
        TBCFOP.Open "Select Movimentar_estoque_pc from Empresa where Codigo = " & TBTempo!ID_empresa & " and Movimentar_estoque_pc = 'True'", Conexao, adOpenKeyset, adLockOptimistic
        If TBCFOP.EOF = False And IsNull(TBEstoque!estoque_real_PC) = False And TBEstoque!estoque_real_PC > 0 Then TBComponente!Quantidade_PC = TBComponente!quantidade
        TBCFOP.Close
        
        TBComponente.Update
        TBTempo.MoveNext
    Loop
End If
TBTempo.Close

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Public Sub ProcAtualizaQtdeExpProdPed(ID_prod_NF_ExpProd As Long, Codinterno_ExpProd As String, Qtde_ExpProd As Double, Ordem_ExpProd As String, IDestoque_ExpProd As Long, Data_ExpProd As Date)
On Error GoTo tratar_erro

QTLOTE = 0

Set TBComponente = CreateObject("adodb.recordset")
StrSql = "Select NFPP.*, VC.* from tbl_Detalhes_Nota_pedidos NFPP INNER JOIN vendas_carteira VC ON VC.Codigo = NFPP.ID_Carteira and VC.Desenho = NFPP.Codinterno where ID_prod_NF = " & ID_prod_NF_ExpProd & " and Codinterno = '" & Codinterno_ExpProd & "' order by VC.Prazofinal"

'Debug.print StrSql

TBComponente.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
If TBComponente.EOF = False Then
    Do While TBComponente.EOF = False
        TBComponente!dataexpedicao = Data_ExpProd
        QTLOTE = Format(TBComponente!qtdeexpedida / FunVerificaTabelaConversaoUnidade(TBComponente!Unidade, TBComponente!Unidade_com), "###,##0.0000")
        If Qtde_ExpProd > (TBComponente!Qtde_produzir - QTLOTE) Then qt = TBComponente!Qtde_produzir - QTLOTE Else qt = Qtde_ExpProd
        TBComponente!qtdeexpedida = Format(TBComponente!qtdeexpedida + (qt * FunVerificaTabelaConversaoUnidade(TBComponente!Unidade, TBComponente!Unidade_com)), "###,##0.0000")
        TBComponente.Update
        Qtde_ExpProd = Qtde_ExpProd - qt
        
        'Vincula pedido na ordem para estoque
        If qt > 0 And IsNumeric(Ordem_ExpProd) = True And Ordem_ExpProd <> "0" And Left(TBEstoque!status, 13) = "ENTRADA_ORDEM" Then
            Permitido2 = False
            Set TBGravar = CreateObject("adodb.recordset")
            TBGravar.Open "Select * from Producao_pedidos where IDcarteira = " & TBComponente!ID_carteira & " and Ordem = " & Ordem_ExpProd, Conexao, adOpenKeyset, adLockOptimistic
            If TBGravar.EOF = True Then
                TBGravar.AddNew
                Permitido2 = True
            Else
                If TBGravar!Expedicao = True Then Permitido2 = True
            End If
            If Permitido2 = True Then
                TBGravar!Data = Data_ExpProd
                TBGravar!Responsavel = pubUsuario
                TBGravar!IDcarteira = TBComponente!ID_carteira
                TBGravar!Ordem = Ordem_ExpProd
                TBGravar!Qtde_empenho = TBGravar!Qtde_empenho + qt
                TBGravar!Qtde_entrada = TBGravar!Qtde_empenho
                TBGravar!Expedicao = True
                TBGravar.Update
            End If
        End If
        
        'Atualiza qtde. de saida no empenho
        Do While qt > 0
            Set TBGravar = CreateObject("adodb.recordset")
            'TBGravar.Open "Select * from Estoque_Controle_Empenho_Vendas where ID_carteira = " & TBComponente!ID_carteira & " and Qtde_empenhada - Qtde_saida > 0", Conexao, adOpenKeyset, adLockOptimistic
            TBGravar.Open "Select * from Estoque_Controle_Empenho_Vendas where ID_carteira = " & TBComponente!ID_carteira & " and ID_estoque = " & IDestoque_ExpProd & " and Qtde_empenhada - Qtde_saida > 0", Conexao, adOpenKeyset, adLockOptimistic
            If TBGravar.EOF = False Then
                If (TBGravar!Qtde_empenhada - TBGravar!Qtde_saida) >= qt Then
                    TBGravar!Qtde_saida = TBGravar!Qtde_saida + qt
                    qt = 0
                Else
                    TBGravar!Qtde_saida = TBGravar!Qtde_saida + (TBGravar!Qtde_empenhada - TBGravar!Qtde_saida)
                    qt = qt - (TBGravar!Qtde_empenhada - TBGravar!Qtde_saida)
                End If
                TBGravar.Update
            Else
                GoTo Prosseguir
            End If
            TBGravar.Close
        Loop
Prosseguir:
        
        If Qtde_ExpProd <= 0 Then Exit Sub
        
        TBComponente.MoveNext
    Loop
End If
TBComponente.Close

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcCriaCreditoCCProdutoItemSelecionada(Codproduto_CC As Long, Data_CC As Date, Empresa_CC As Integer, IDoperacao_CC As Long, ValorTotal_CC As Double)
On Error GoTo tratar_erro

Set TBFIltro = CreateObject("adodb.recordset")
TBFIltro.Open "Select * from projproduto where Codproduto = " & Codproduto_CC & " and ID_CC is not null", Conexao, adOpenKeyset, adLockOptimistic
If TBFIltro.EOF = False Then
    If TBFIltro!ID_CC <> "" Then
        Set TBFI = CreateObject("adodb.recordset")
        TBFI.Open "Select * from CC_realizado", Conexao, adOpenKeyset, adLockOptimistic
        TBFI.AddNew
        TBFI!ID_CC = TBFI!ID_CC
        TBFI!valor = ValorTotal_CC
        TBFI!Data = Data_CC
        TBFI!Responsavel = pubUsuario
        TBFI!ID_empresa = Empresa_CC
        TBFI!Operacao = "CrÈdito"
        TBFI!ID_estoque = IDoperacao_CC
        TBFI!Cod_produto = TBFIltro!Codproduto
        TBFI!ID_PC = TBFIltro!ID_PC
        TBFI!Bloqueado = False
        TBFI.Update
        
        'Grava movimentaÁ„o no centro consolidado
        Set TBAfericao = CreateObject("adodb.recordset")
        TBAfericao.Open "Select * from Usuarios_Setor_Consolidacao where ID_CC_consolidado = " & TBFIltro!ID_CC, Conexao, adOpenKeyset, adLockOptimistic
        If TBAfericao.EOF = False Then
            Do While TBAfericao.EOF = False
                Set TBFI = CreateObject("adodb.recordset")
                TBFI.Open "Select * from CC_realizado", Conexao, adOpenKeyset, adLockOptimistic
                TBFI.AddNew
                TBFI!ID_CC = TBAfericao!ID_CC
                TBFI!valor = ValorTotal_CC
                TBFI!Data = Data_CC
                TBFI!Responsavel = pubUsuario
                TBFI!ID_empresa = Empresa_CC
                TBFI!Operacao = "CrÈdito"
                TBFI!ID_estoque = IDoperacao_CC
                TBFI!Cod_produto = TBFIltro!Codproduto
                TBFI!ID_PC = TBFIltro!ID_PC
                TBFI!Bloqueado = False
                TBFI.Update
                
                Set TBCiclo = CreateObject("adodb.recordset")
                TBCiclo.Open "Select * from Usuarios_Setor_Consolidacao where ID_CC_consolidado = " & TBAfericao!ID_CC, Conexao, adOpenKeyset, adLockOptimistic
                If TBCiclo.EOF = False Then
                    Do While TBCiclo.EOF = False
                        Set TBFI = CreateObject("adodb.recordset")
                        TBFI.Open "Select * from CC_realizado", Conexao, adOpenKeyset, adLockOptimistic
                        TBFI.AddNew
                        TBFI!ID_CC = TBCiclo!ID_CC
                        TBFI!valor = ValorTotal_CC
                        TBFI!Data = Data_CC
                        TBFI!Responsavel = pubUsuario
                        TBFI!ID_empresa = Empresa_CC
                        TBFI!Operacao = "CrÈdito"
                        TBFI!ID_estoque = IDoperacao_CC
                        TBFI!Cod_produto = TBFIltro!Codproduto
                        TBFI!ID_PC = TBFIltro!ID_PC
                        TBFI!Bloqueado = False
                        TBFI.Update
                        TBCiclo.MoveNext
                    Loop
                End If
                TBCiclo.Close
                
                TBAfericao.MoveNext
            Loop
        End If
        TBAfericao.Close
    End If
End If
TBFIltro.Close

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub FunAlterarProdSimiliarOrdemPC(ID_empresa As Integer, ID_pedido As Long)
On Error GoTo tratar_erro

If FunVerifProdSimiliar(ID_empresa) = True Then
    Set TBFIltro = CreateObject("adodb.recordset")
    TBFIltro.Open "Select CPL.*, P.Peso_metro, P.Un_kg, P.SubTipoItem, P.Unidade, P.PBruto, P.Comprimento, P.Largura from Compras_pedido_lista CPL INNER JOIN Projproduto P ON P.Desenho = CPL.Desenho where CPL.IDpedido = " & ID_pedido & " and CPL.Ordem IS NOT NULL and CPL.Ordem <> 0 and (CPL.Status_Item = 'APROVADO' or CPL.Status_Item = 'N_RECEBIDO')", Conexao, adOpenKeyset, adLockOptimistic
    If TBFIltro.EOF = False Then
        Do While TBFIltro.EOF = False
            Set TBMaterial = CreateObject("adodb.recordset")
            TBMaterial.Open "Select PM.*, PR.Quant, P.Un_kg, P.SubTipoItem, P.PBruto from Producaomaterial PM INNER JOIN Producao PR ON PR.Ordem = PM.Ordem INNER JOIN projproduto P ON PM.Codigo = P.Desenho INNER JOIN Projproduto P1 ON P1.ID_similar = P.ID_similar where PM.Ordem = " & TBFIltro!Ordem & " and P1.Desenho = '" & TBFIltro!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBMaterial.EOF = False Then
                Set TBGravar = CreateObject("adodb.recordset")
                TBGravar.Open "Select * from Producaomaterial where Ordem = " & TBFIltro!Ordem & " and Codigo = '" & TBFIltro!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBGravar.EOF = True Then
                    TBGravar.AddNew
                    valor = 0
                    Valor1 = 0
                Else
                    valor = TBGravar!Requisitado
                    Valor1 = TBGravar!DimensaoTotal
                End If
                
                TBGravar!quantidade = TBMaterial!quantidade
                TBGravar!Unidade = TBMaterial!Unidade
                TBGravar!CODIGO = TBFIltro!Desenho
                TBGravar!Descricao = TBFIltro!Descricao
                TBGravar!Ordem = TBMaterial!Ordem
                TBGravar!PesoMetro = TBFIltro!peso_metro
                
                If TBFIltro!Un_Kg = "Mt≤" Then
                    TBGravar!Dimensao = TBFIltro!Comprimento * TBFIltro!Largura
                    TBGravar!DimensaoTotal = ((TBGravar!Dimensao / 1000) / 1000) * TBFIltro!Quant_Comp_PC
                Else
                    TBGravar!Dimensao = TBMaterial!Dimensao
                    TBGravar!DimensaoTotal = TBMaterial!DimensaoTotal
                End If
                
                'Peso bruto
                If IsNull(TBFIltro!PBruto) = False And TBFIltro!PBruto <> "0" Then
                    TBGravar!pesounidade = TBFIltro!PBruto
                Else
                    Select Case TBFIltro!Un_Kg
                        Case "Mt/L": TBGravar!pesounidade = Format(TBFIltro!peso_metro / 1000 * TBGravar!Dimensao, "###,##0.0000000000")
                        Case "PÁ": TBGravar!pesounidade = Format(TBFIltro!peso_metro, "###,##0.0000000000")
                        Case "Mt≤": TBGravar!pesounidade = Format(((TBGravar!Dimensao * TBFIltro!peso_metro) / 1000) / 1000, "###,##0.0000000000")
                        Case "N/a": TBGravar!pesounidade = Format(0, "###,##0.0000000000")
                    End Select
                End If
                
                TBGravar!PesoTotal = Format(TBGravar!pesounidade * (TBMaterial!quantidade / TBMaterial!Quant), "###,##0.0000000000")
                TBGravar!Requisitado = Format(TBFIltro!Quant_Comp, "###,##0.0000")
                TBGravar!Total_pc = TBFIltro!Quant_Comp_PC
                TBGravar!versao = TBMaterial!versao
                TBGravar!Saida = "N√O"
                TBGravar.Update
                                    
                If TBFIltro!Quant_Comp < (TBMaterial!Requisitado + valor) Then
                    TBMaterial!Requisitado = Format((TBMaterial!Requisitado + valor) - TBFIltro!Quant_Comp, "###,##0.0000")
                    If TBFIltro!Un_Kg = "Mt≤" Then TBMaterial!DimensaoTotal = Format((TBMaterial!DimensaoTotal + Valor1) - TBGravar!DimensaoTotal, "###,##0.0000")
                    
                    If TBMaterial!Unidade = "KG" Or TBMaterial!SubTipoItem = 1 Or TBMaterial!SubTipoItem = 2 Or TBMaterial!SubTipoItem = 3 Then
                        If TBMaterial!Unidade = "KG" And (TBMaterial!Un_Kg = "Mt≤" Or TBMaterial!Un_Kg = "Mt/L") Then
                            If IsNull(TBMaterial!PBruto) = False And TBMaterial!PBruto > 0 And TBMaterial!PBruto <> "" Then TBMaterial!Total_pc = Format(TBMaterial!Requisitado / TBMaterial!PBruto, "###,##0.0000") Else TBMaterial!Total_pc = Null
                        Else
                            If TBMaterial!SubTipoItem = 1 Or TBMaterial!SubTipoItem = 2 Or TBMaterial!SubTipoItem = 3 Then TBMaterial!Total_pc = TBMaterial!Requisitado Else TBMaterial!Total_pc = Null
                        End If
                        TBMaterial.Update
                    End If
                Else
                    Conexao.Execute "DELETE Producaomaterial where Idmateriaprima = " & TBMaterial!IdMateriaPrima
                End If
                
                TBGravar.Close
            End If
            TBFIltro.MoveNext
        Loop
    End If
    TBFIltro.Close
End If

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Function FunVerifOPCarregaOS(Combo As ComboBox, Ordem As String, Novo As Boolean, CarregarOS As Boolean) As Boolean
On Error GoTo tratar_erro

FunVerifOPCarregaOS = False
With Combo
    If CarregarOS = True Then
        .Clear
        .Locked = True
        .TabStop = False
    End If
    
    If Ordem = "" Or Ordem = "0" Then Exit Function
    If Novo = True Then TextoFiltro = " and DtValidacao IS NOT NULL and DtValidacao_custo IS NULL" Else TextoFiltro = ""
    Set TBproducao = CreateObject("adodb.recordset")
    TBproducao.Open "Select * from Producao where Ordem = " & Ordem & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
    If TBproducao.EOF = False Then
        FunVerifOPCarregaOS = True
        
        If CarregarOS = True Then
            Set TBOrdem = CreateObject("adodb.recordset")
            TBOrdem.Open "Select IDProducao, custos from ordemservico where Ordem = " & TBproducao!Ordem & " order by fase, retrabalho, IDproducao", Conexao, adOpenKeyset, adLockOptimistic
            If TBOrdem.EOF = False Then
                .AddItem ""
                .Locked = False
                .TabStop = True
                Do While TBOrdem.EOF = False
                    IDFase = 0
                    Set TBFI = CreateObject("adodb.recordset")
                    TBFI.Open "Select OSMU.OS from Ordemservico_maq_utilizadas OSMU INNER JOIN CadMaquinas CM ON OSMU.Maquina = CM.Maquina where OSMU.Ordem = " & TBOrdem!IDProducao & " and CM.Custos = 'False'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBFI.EOF = False Then
                        Do While TBFI.EOF = False
                            If IDFase <> TBFI!OS Then .AddItem TBFI!OS
                            IDFase = TBFI!OS
                            TBFI.MoveNext
                        Loop
                    Else
                        If TBOrdem!custos = False Then .AddItem TBOrdem!IDProducao
                    End If
                    TBFI.Close
                    TBOrdem.MoveNext
                Loop
            End If
            TBOrdem.Close
        End If
    Else
        If Novo = True Then MsgTexto = "validada com este n˙mero, ou a mesma est· com o resultado validado" Else MsgTexto = "com este n˙mero"
        USMsgBox ("N„o foi encontrado nenhuma ordem de produÁ„o " & MsgTexto & "."), vbExclamation, "CAPRIND v5.0"
    End If
End With

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Function FunVerifOP(Ordem As String, Novo As Boolean) As Boolean
On Error GoTo tratar_erro

FunVerifOP = False
If Ordem = "" Or Ordem = "0" Then Exit Function
If Novo = True Then TextoFiltro = " and DtValidacao IS NOT NULL and DtValidacao_custo IS NULL" Else TextoFiltro = ""
Set TBproducao = CreateObject("adodb.recordset")
TBproducao.Open "Select * from Producao where Ordem = " & Ordem & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
If TBproducao.EOF = False Then
    FunVerifOP = True
Else
    USMsgBox ("N„o foi encontrado nenhuma ordem de produÁ„o validada com este n˙mero."), vbExclamation, "CAPRIND v5.0"
End If

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Sub ProcEmpenharREAutomOrdem(IDEstoque As Long, Qtde_entrada As Double, LOTE As String, Data1 As Date, Responsavel As String, Codinterno As String, Excluir As Boolean)
On Error GoTo tratar_erro

'Empenha para a ordem vinculada ao mesmo pedido interno que requisita o material
Set TBCFOP = CreateObject("adodb.recordset")
TBCFOP.Open "Select PP.OrdemEmpenho, PP.Qtde_empenho from Producao_pedidos PP INNER JOIN vendas_carteira VC ON VC.Codigo = PP.IDcarteira where PP.Ordem = " & LOTE & " and PP.OrdemEmpenho IS NOT NULL and PP.OrdemEmpenho <> 0 order by VC.Prazofinal", Conexao, adOpenKeyset, adLockOptimistic
If TBCFOP.EOF = False Then
    Do While TBCFOP.EOF = False And Qtde_entrada > 0
        Set TBComponente = CreateObject("adodb.recordset")
        TBComponente.Open "Select * from Producao_NF_Consignada where Data = '" & Data & "' and Responsavel = '" & Responsavel & "' and Ordem = " & TBCFOP!OrdemEmpenho & " and IDestoque = " & IDEstoque, Conexao, adOpenKeyset, adLockOptimistic
        If Excluir = False Then
            'Verifica quantidade empenhada
            Qtde = TBCFOP!Qtde_empenho
            
            If TBComponente.EOF = True Then TBComponente.AddNew
            TBComponente!Data = Data
            TBComponente!Responsavel = Responsavel
            TBComponente!Ordem = TBCFOP!OrdemEmpenho
            TBComponente!Codinterno = Codinterno
            TBComponente!IDEstoque = IDEstoque
            
            If Qtde_entrada > Qtde Then
                TBComponente!quantidade = TBComponente!quantidade + Qtde
                Qtde_entrada = Qtde_entrada - Qtde
            Else
                TBComponente!quantidade = TBComponente!quantidade + Qtde_entrada
                Qtde_entrada = 0
            End If
            TBComponente!Quantidade_PC = TBComponente!quantidade
            TBComponente.Update
        Else
            If TBComponente.EOF = False Then
                If TBComponente!quantidade - Qtde_entrada <= 0 Then
                    Qtde_entrada = Qtde_entrada - TBComponente!quantidade
                    TBComponente.Delete
                Else
                    If Qtde_entrada > TBComponente!quantidade Then
                        Qtde_entrada = Qtde_entrada - TBComponente!quantidade
                        TBComponente!quantidade = 0
                    Else
                        TBComponente!quantidade = TBComponente!quantidade - Qtde_entrada
                        Qtde_entrada = 0
                    End If
                    TBComponente!Quantidade_PC = TBComponente!quantidade
                    TBComponente.Update
                End If
            End If
        End If
        TBComponente.Close
Proximo:
        TBCFOP.MoveNext
    Loop
End If
TBCFOP.Close

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLibBlocTxt(Texto As TextBox, Libera As Boolean)
On Error GoTo tratar_erro

With Texto
    If Libera = True Then
        .Locked = False
        .TabStop = True
    Else
        .Locked = True
        .TabStop = False
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLibBlocCmb(Combo As ComboBox, Libera As Boolean)
On Error GoTo tratar_erro

With Combo
    If Libera = True Then
        .Locked = False
        .TabStop = True
    Else
        .Locked = True
        .TabStop = False
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcFunAbreBD_Configuracao(ServidorConf As String, BancoConf As String, UsuarioConf As String, SenhaConf As String)
On Error GoTo tratar_erro

Set Conexao_Configuracao = New ADODB.Connection
With Conexao_Configuracao
    .Provider = "SQLOLEDB"
    .Properties("Data Source").Value = ServidorConf
    .Properties("Initial catalog").Value = BancoConf
    .Properties("User ID").Value = UsuarioConf
    .Properties("Password").Value = SenhaConf
    .Properties("Persist Security Info") = "False"
    .Open
    .Close
End With

Exit Sub
tratar_erro:
    If Err.Number = "-2147467259" Then
        Permitido = False
        Exit Sub
    End If
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Function FunVerifTemEstrutura(Codinterno As String) As Boolean
On Error GoTo tratar_erro
   
FunVerifTemEstrutura = True
Set TBTempo = CreateObject("adodb.recordset")
TBTempo.Open "Select P.Codproduto from Projproduto P LEFT JOIN ProjConjunto PC ON PC.Codproduto = P.Codproduto where P.Desenho = '" & Codinterno & "'", Conexao, adOpenKeyset, adLockOptimistic
'TBTempo.Open "Select P.Codproduto from Projproduto P LEFT JOIN ProjConjunto PC ON PC.Codproduto = P.Codproduto where P.Desenho = '" & Codinterno & "' and P.Subtipoitem = 0", Conexao, adOpenKeyset, adLockOptimistic
If TBTempo.EOF = True Then FunVerifTemEstrutura = False
TBTempo.Close

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Sub ProcRemoveObjetosResize(Formulario As Form)
On Error GoTo tratar_erro
Dim Controle As Control

For i = 0 To Formulario.Controls.Count - 1
    Set Controle = Formulario.Controls(i)
    If TypeOf Controle Is ListView Then Controle.Tag = "Ex_Columns,Ex_Font"
    If TypeOf Controle Is MSFlexGrid Then Controle.Tag = "Ex_Font"
    If TypeOf Controle Is USToolBar Then Controle.Tag = "Ex_Font"
    If TypeOf Controle Is USTab Then Controle.Tag = "Ex_Font"
    
'    If TypeOf Controle Is USTreeView Then Controle.Tag = "Ex_Font"
    
'    If TypeOf Controle Is Label Then Controle.Tag = "Ex_Font"
'    If TypeOf Controle Is TextBox Then Controle.Tag = "Ex_Height"
'    If TypeOf Controle Is ComboBox Then Controle.Tag = "Ex_Height, Ex_Font"
'    If TypeOf Controle Is USButton Then Controle.Tag = "Ex_Height, Ex_Font"
'    If TypeOf Controle Is MaskEdBox Then Controle.Tag = "Ex_Height"
'    If TypeOf Controle Is frame Then Controle.Tag = "Ex_Font"
'    If TypeOf Controle Is Button Then Controle.Tag = "Ex_Height, Ex_Font"
'    If TypeOf Controle Is DTPicker Then Controle.Tag = "Ex_Height, Ex_Font"
'    If TypeOf Controle Is SSTab Then Controle.Tag = "Ex_Font"
'    If TypeOf Controle Is CommandButton Then Controle.Tag = "Ex_Height, Ex_Font"
'    If TypeOf Controle Is Image Then Controle.Tag = "Ex_Height"
'    If TypeOf Controle Is USTextBox Then Controle.Tag = "Ex_Height, Ex_Font"
'    If TypeOf Controle Is OptionButton Then Controle.Tag = "Ex_Height, Ex_Font"
'    If TypeOf Controle Is CheckBox Then Controle.Tag = "Ex_Height, Ex_Font"
Next i


Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcNivel2Consignacao()
On Error GoTo tratar_erro

Set TBNivel2 = CreateObject("adodb.recordset")
TBNivel2.Open "Select * from Estoque_consignado_saida where ID = '" & TBAbrir!ID_nota & "' and int_Codigo = '" & TBAbrir!ID_Produto & "' order by Id", Conexao, adOpenKeyset, adLockOptimistic
If TBNivel2.EOF = False Then
    Do While TBNivel2.EOF = False
            Saldo = Saldo - TBAbrir!Qtde
            Contador1 = Contador1 + 1
            arrNodes(Contador1).Level = 1
            arrNodes(Contador1).Text = TBNivel2!int_NotaFiscal & vbTab & "" & vbTab & TBNivel2!dt_DataEmissao & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & Format(TBAbrir!Qtde, "###,##0.00") & vbTab & Format(Saldo, "###,##0.00")
        TBNivel2.MoveNext
    Loop
End If

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcNivel3Consignacao()
On Error GoTo tratar_erro

Set TBNivel3 = CreateObject("adodb.recordset")
TBNivel3.Open "Select * from projconjunto where codproduto = " & Codproduto & TextoFiltro & " order by Posicao, Codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBNivel2.EOF = False Then
    Do While TBNivel2.EOF = False
            
            Contador1 = Contador1 + 1
            arrNodes(Contador1).Level = 2
            arrNodes(Contador1).Text = TBNivel3!Desenho & vbTab & PosicaoTexto & vbTab & Codproduto & vbTab & CodRef & vbTab & TBNivel2!Descricao & vbTab & TBNivel2!Unidade & vbTab & TBNivel2!Versao_desenho & vbTab & Kg_un & vbTab & Un_Kg & vbTab & Dim_mm & vbTab & Kg_pc & vbTab & IIf(IsNull(TBAbrir!Largura), 0, Format(TBAbrir!Largura, "###,##0.00")) & vbTab & IIf(IsNull(TBAbrir!Comprimento), "", Format(TBAbrir!Comprimento, "###,##0.00")) & vbTab & QtTexto & vbTab & Peso_total & vbTab & Format(valor, "###,##0.00000000") & vbTab & DataValidacao & vbTab & RespValidacao & vbTab & TBNivel2!CODIGO & vbTab & PartNumber & vbTab & TBNivel2!Obs
        TBNivel2.MoveNext
    Loop
End If
Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub



'Sub ProcCorrigeTamColLista(Lista As ListView)
'On Error GoTo tratar_erro
'Dim Column As Long
'Dim Counter As Long
'
'Counter = 0
'For Column = 1 To Lista.ColumnHeaders.Count - 2
'    'SendMessage Lista.hwnd, LVM_SETCOLUMNWIDTH, Column, LVSCW_AUTOSIZE_USEHEADER
'    With Lista
'        frmMenuCaprind.Label_ListView.Font.size = Lista.Font.size & "  "
'        frmMenuCaprind.Label_ListView1.Font.size = Lista.Font.size & "  "
'        If Lista.ColumnHeaders.Item(Column) <> "" Then
'        For InitFor = 1 To .ListItems.Count
'            frmMenuCaprind.Label_ListView = Lista.ColumnHeaders.Item(Column)
'            If Counter = 0 Then frmMenuCaprind.Label_ListView1 = .ListItems.Item(InitFor) Else frmMenuCaprind.Label_ListView1 = .ListItems.Item(InitFor).ListSubItems(Counter)
'            If Lista.ColumnHeaders(Column).Width < frmMenuCaprind.Label_ListView.Width Then
'                Lista.ColumnHeaders(Column).Width = frmMenuCaprind.Label_ListView.Width
'            ElseIf Lista.ColumnHeaders(Column).Width < frmMenuCaprind.Label_ListView1.Width Then
'                    Lista.ColumnHeaders(Column).Width = frmMenuCaprind.Label_ListView1.Width
'            End If
'        Next InitFor
'        End If
'    End With
'    Counter = Counter + 1
'Next
'
'Exit Sub
'tratar_erro:
'    usMsgbox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
'    Exit Sub
'End Sub

Function FunBuscaDescPadraoFamilia(Familia As String, Codinterno As String, Descricao As String) As String
On Error GoTo tratar_erro

If VerifDadosPadraoFamilia = False Then Exit Function
Permitido = True
If Codinterno <> "" Then
    Set TBFamilia = CreateObject("adodb.recordset")
    TBFamilia.Open "Select Descricao from projproduto where Desenho = '" & Codinterno & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBFamilia.EOF = False Then
        FunBuscaDescPadraoFamilia = IIf(IsNull(TBFamilia!Descricao), "", (TBFamilia!Descricao))
        Permitido = False
    End If
End If
If Permitido = True Then
    FunBuscaDescPadraoFamilia = Descricao
    Set TBFamilia = CreateObject("adodb.recordset")
    TBFamilia.Open "Select Descinicio from projfamilia where Familia = '" & Familia & "' and Descinicio IS NOT NULL and Descinicio <> N''", Conexao, adOpenKeyset, adLockOptimistic
    If TBFamilia.EOF = False Then
        If USMsgBox("Existe descriÁ„o padr„o cadastrada para essa famÌlia, deseja carregar essa descriÁ„o?", vbYesNo, "CAPRIND v5.0") = vbYes Then FunBuscaDescPadraoFamilia = IIf(IsNull(TBFamilia!Descinicio), "", TBFamilia!Descinicio)
    'ElseIf Descricao <> "" Then
            'If usMsgbox("N„o existe descriÁ„o padr„o cadastrada para essa famÌlia, deseja apagar a descriÁ„o informada?", vbyesno, "CAPRIND v5.0") = vbYes Then FunBuscaDescPadraoFamilia = ""
    End If
End If
TBFamilia.Close

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Function FunBuscaIDCFPadraoFamilia(Familia As String, Codinterno As String, ID_CF As Long) As Long
On Error GoTo tratar_erro

Permitido = True
If Codinterno <> "" Then
    Set TBFamilia = CreateObject("adodb.recordset")
    TBFamilia.Open "Select ID_CF from projproduto where Desenho = '" & Codinterno & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBFamilia.EOF = False Then
        FunBuscaIDCFPadraoFamilia = IIf(IsNull(TBFamilia!ID_CF), 0, (TBFamilia!ID_CF))
        Permitido = False
    End If
End If
If Permitido = True Then
    FunBuscaIDCFPadraoFamilia = ID_CF
    Set TBFamilia = CreateObject("adodb.recordset")
    TBFamilia.Open "Select ID_CF from projfamilia where Familia = '" & Familia & "' and ID_CF IS NOT NULL and ID_CF <> 0 and ID_CF <> N''", Conexao, adOpenKeyset, adLockOptimistic
    If TBFamilia.EOF = False Then
        FunBuscaIDCFPadraoFamilia = IIf(IsNull(TBFamilia!ID_CF), "", TBFamilia!ID_CF)
    End If
End If
TBFamilia.Close

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Public Function FunProgBarKFTP(prog_bar As PictureBox, prog_hst As PictureBox, pc As Byte)
On Error GoTo tratar_erro

prog_bar.Width = Int((pc / 100) * prog_hst.Width)
prog_bar.Visible = (pc <> 0)
DoEvents
prog_bar.Refresh

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Public Function FunConectaKFTP(kftp As kftp, PastaServidor As String, MostrarMsg As Boolean) As Boolean
On Error GoTo tratar_erro

FunConectaKFTP = True
'Antiga senha"Cap0902loc@@62"

If Not kftp.Connect("ftp.caprind.com.br", "caprind1", "C@prind0902loc$", Val("")) Then
    If kftp.LastError = "Please disconnect" Then kftp.Disconnect
    If MostrarMsg = True Then USMsgBox ("N„o foi possÌvel conectar com o servidor de atualizaÁ„o, a atualizaÁ„o ser· encerrada."), vbExclamation, "CAPRIND v5.0"
    FunConectaKFTP = False
Else
    If Not kftp.ChangeWorkingDir(PastaServidor) Then
        If MostrarMsg = True Then USMsgBox ("N„o foi localizado a pasta no servidor de atualizaÁ„o, a atualizaÁ„o ser· encerrada."), vbExclamation, "CAPRIND v5.0"
        FunConectaKFTP = False
    End If
End If

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Public Function FunDownloadKFTP(kftp As kftp, NomeArquivoBaixar As String, CaminhoSalvar As String) As Boolean
On Error GoTo tratar_erro

FunDownloadKFTP = True
kftp.ChangeTransfertMode (Binary)
If Not kftp.DownloadFile(NomeArquivoBaixar, CaminhoSalvar, 0, True) Then
    USMsgBox ("N„o foi possÌvel baixar a atualizaÁ„o, a atualizaÁ„o ser· encerrada."), vbExclamation, "CAPRIND v5.0"
    FunDownloadKFTP = False
End If

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
     Exit Function
End Function

Public Function FunVerifTabelaCampoExiste(NomeTabela As String, NomeCampo As String, VerifTabela As Boolean) As Boolean
On Error GoTo tratar_erro

FunVerifTabelaCampoExiste = False
Set TBAbrir = CreateObject("adodb.recordset")
If VerifTabela = True Then
    TBAbrir.Open "SELECT * FROM SYSOBJECTS OBJ INNER JOIN SYSCOLUMNS COL ON OBJ.ID = COL.ID WHERE OBJ.NAME LIKE '" & NomeTabela & "'", Conexao_NFe, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        FunVerifTabelaCampoExiste = True
    End If
Else
    TBAbrir.Open "SELECT * FROM SYSOBJECTS OBJ INNER JOIN SYSCOLUMNS COL ON OBJ.ID = COL.ID WHERE OBJ.NAME LIKE '" & NomeTabela & "'", Conexao_NFe, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Do While TBAbrir.EOF = False
            If TBAbrir!Name = NomeCampo Then
                FunVerifTabelaCampoExiste = True
                TBAbrir.Close
                Exit Function
            End If
            TBAbrir.MoveNext
        Loop
    End If
End If
TBAbrir.Close

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Public Function FunVerifTabelaCampoExisteGNFe(NomeTabela As String, NomeCampo As String, VerifTabela As Boolean) As Boolean
On Error GoTo tratar_erro

FunVerifTabelaCampoExisteGNFe = False
Set TBAbrir = CreateObject("adodb.recordset")
If VerifTabela = True Then
    TBAbrir.Open "SELECT * FROM SYSOBJECTS OBJ INNER JOIN SYSCOLUMNS COL ON OBJ.ID = COL.ID WHERE OBJ.NAME LIKE '" & NomeTabela & "'", Conexao_NFe, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        FunVerifTabelaCampoExisteGNFe = True
    End If
Else
    TBAbrir.Open "SELECT * FROM SYSOBJECTS OBJ INNER JOIN SYSCOLUMNS COL ON OBJ.ID = COL.ID WHERE OBJ.NAME LIKE '" & NomeTabela & "'", Conexao_NFe, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Do While TBAbrir.EOF = False
            If TBAbrir!Name = NomeCampo Then
                FunVerifTabelaCampoExisteGNFe = True
                TBAbrir.Close
                Exit Function
            End If
            TBAbrir.MoveNext
        Loop
    End If
End If
TBAbrir.Close

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Public Function FunVerifConfigEmail(TextoFiltro As String, Aplicacao As String, Usuario As String) As Boolean
On Error GoTo tratar_erro

FunVerifConfigEmail = True
Set TBFI = CreateObject("adodb.recordset")
TBFI.Open "Select * from Empresa_email where " & TextoFiltro & " and Aplicacao = '" & Aplicacao & "' and Usuario_caprind = '" & Usuario & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBFI.EOF = False Then
    Email = TBFI!Email
    Nome_email = TBFI!Nome
    Usuario_email = TBFI!Usuario
    Senha_email = TBFI!Senha
    Servidor_SMTP = TBFI!Servidor_SMTP
    Porta_email = TBFI!Porta
Else
    Select Case Aplicacao
        Case "C": MsgTexto = "Compras"
        Case "CU": MsgTexto = "Custos"
        Case "F": MsgTexto = "Financeiro"
        Case "V": MsgTexto = "Vendas"
    End Select
    USMsgBox ("… necess·rio cadastrar as configuraÁıes de e-mail para a aplicaÁ„o " & MsgTexto & " e usu·rio " & Usuario & " antes de enviar."), vbExclamation, "CAPRIND v5.0"
    FunVerifConfigEmail = False
End If

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Public Function FunVerifHorarioChat() As Boolean
On Error GoTo tratar_erro

'Verifica hor·rio do servidor e se o chat est· disponÌvel neste hor·rio
FunVerifHorarioChat = True
Texto = ""
Numero = 0
Numero1 = Len(NomeServidor)
Hora = 0
If Numero1 <> 1 Then
    Do While Numero1 <> 0
        If Texto = "\" Then GoTo Pula
        Texto = Left(NomeServidor, (Numero + 1))
        Texto = Right(Texto, Len(Texto) - Numero)
        Numero = Numero + 1
        Numero1 = Numero1 - 1
    Loop
End If
Pula:
    Familiatext = Left(NomeServidor, Numero - 1)
    Dataini = Format(FunHoraServidor("\\" & Familiatext), "hh:mm:ss")
    If Dataini < "09:00:00" Or Dataini >= "12:00:00" And Dataini <= "13:00:00" Or Dataini > "17:00:00" Then
        USMsgBox ("O chat (online) sÛ est· disponÌvel nos hor·rios das 09hs as 12hs e das 13hs as 17hs."), vbInformation, "CAPRIND v5.0"
        FunVerifHorarioChat = False
    End If

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Public Function FunHoraServidor(ByVal pNomeServidor As String) As Variant
On Error GoTo tratar_erro
Dim t As TIME_OF_DAY
Dim tPtr As Long
Dim Resultado As Long
Dim szServer As String
Dim dataServidor As Date

If Left(pNomeServidor, 2) = "\\" Then szServer = StrConv(pNomeServidor, vbUnicode) Else szServer = StrConv("\\" & pNomeServidor, vbUnicode)
Resultado = NetRemoteTOD(szServer, tPtr)

If Resultado = 0 Then
    Call CopyMemory(t, ByVal tPtr, Len(t))
    dataServidor = DateSerial(70, 1, 1) + (t.t_elapsedt / 60 / 60 / 24)
    dataServidor = dataServidor - (t.t_timezone / 60 / 24)
    NetApiBufferFree (tPtr)
    FunHoraServidor = dataServidor
Else
    FunHoraServidor = Now
End If

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Public Function FunVerifVlrTotalFat12UltMesesSomado(ID_empresa As Integer) As Double
On Error GoTo tratar_erro

ProcAcertaFaturamento12Meses (ID_empresa)

FunVerifVlrTotalFat12UltMesesSomado = 0

Set TBTempo = CreateObject("adodb.recordset")
'=============================================================
' Primeiro soma o valor total faturado nos ultimos doze meses informado
'=============================================================
'StrSql = "Select Sum(IFM.Valor-ISNULL(FDM.VlrTotalFat,0)) as Valor from Impostos_FaturamentoMes IFM LEFT OUTER JOIN Faturamento_12ultimos_meses_devolucao_mensal FDM on IFM.Ano = FDM.Ano and IFM.Mes = FDM.Mes and IFM.ID_empresa = FDM.ID_empresa where IFM.ID_empresa = " & IDempresa
StrSql = "Select sum(Valor) as Valor from Impostos_FaturamentoMes where ID_empresa = " & IDempresa

'Debug.print StrSql

TBTempo.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic

'TBTempo.Open "Select ROUND(SUM(ISNULL(Valor, 0)), 2) as Valor from Impostos_FaturamentoMes where ID_empresa = " & ID_empresa & " and Mes >= " & Month(Date) & " and Ano = " & Year(Date) - 1 & " or Mes < " & Month(Date) & " and Ano = " & Year(Date), Conexao, adOpenKeyset, adLockReadOnly

If TBTempo.EOF = False Then
    FunVerifVlrTotalFat12UltMesesSomado = IIf(IsNull(TBTempo!valor), 0, TBTempo!valor)
End If

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Public Sub ProcAcertaFaturamento12Meses(ID_empresa As Integer)
On Error GoTo tratar_erro
Dim Ano As Long

Set TBTempo = CreateObject("adodb.recordset")
TBTempo.Open "Select * from Faturamento_12ultimos_meses_mensal_calculado where id_empresa  = " & IDempresa & " Order BY Ano desc, Mes asc", Conexao, adOpenKeyset, adLockReadOnly
If TBTempo.EOF = False Then
'Conexao.Execute ("Delete from Impostos_FaturamentoMes")

If TBTempo!Ano = Year(Date) Then
Ano = TBTempo!Ano - 1
Else
Ano = TBTempo!Ano
End If

Do While TBTempo.EOF = False
Conexao.Execute ("Delete from Impostos_FaturamentoMes where Mes = " & TBTempo!Mes & " and Ano = " & Ano & "")


Set TBGravar_NFe = CreateObject("adodb.recordset")
TBGravar_NFe.Open "Select * from Impostos_FaturamentoMes where ID_empresa = " & ID_empresa & " and Mes = " & TBTempo!Mes & " and Ano = " & TBTempo!Ano & "", Conexao, adOpenKeyset, adLockOptimistic
  If TBGravar_NFe.EOF = True Then
      TBGravar_NFe.AddNew
End If
      TBGravar_NFe!ID_empresa = ID_empresa
      TBGravar_NFe!Mes = TBTempo!Mes
      TBGravar_NFe!Ano = TBTempo!Ano
      TBGravar_NFe!valor = TBTempo!Total
      TBGravar_NFe.Update
      TBGravar_NFe.Close

TBTempo.MoveNext
Loop
End If

TBTempo.Close

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Sub ProcCarregaMenu(TreeviewMnu As TreeView, ImageListMnu As ImageList)
On Error GoTo tratar_erro
Dim Formulario1 As String
Dim Formulario2 As String
Dim Formulario3 As String
Dim Formulario4 As String
Dim Formulario5 As String

Set TreeviewMnu.ImageList = ImageListMnu
'limpa qualquer nÛ criado
TreeviewMnu.Nodes.Clear
'exibe linhas
TreeviewMnu.LineStyle = tvwTreeLines
'Exibindo caixa de verificacao
TreeviewMnu.CheckBoxes = False
'Inclui itens
Set nodmenu = TreeviewMnu.Nodes.Add(, , , "Bem vindo, " & pubUsuario, "Menu")
    nodmenu.Expanded = True
    'ConfiguraÁ„o do sistema
    Set nodA = TreeviewMnu.Nodes.Add(nodmenu, tvwChild, , "ConfiguraÁ„o do sistema", "A")
    
        Formulario = "ConfiguraÁ„o do sistema/OpÁıes gerais/ConfiguraÁ„o do sistema"
        Formulario1 = "ConfiguraÁ„o do sistema/OpÁıes gerais/Cadastro de empresa"
        Formulario2 = "ConfiguraÁ„o do sistema/OpÁıes gerais/Cadastro de moedas"
        Formulario3 = "ConfiguraÁ„o do sistema/OpÁıes gerais/Cadastro de unidades"
        Formulario4 = "ConfiguraÁ„o do sistema/OpÁıes gerais/Cadastro de condiÁ„o de pagamento/recebimento"
        Formulario5 = "ConfiguraÁ„o do sistema/OpÁıes gerais/Cadastro de feriados"
        
        Set TBAcessos = CreateObject("adodb.recordset")
        TBAcessos.Open "Select * FROM Acessos WHERE IDUsuario = " & pubIDUsuario & " AND (Acesso = '" & Formulario & "' or Acesso = '" & Formulario1 & "' or Acesso = '" & Formulario2 & "' or Acesso = '" & Formulario3 & "' or Acesso = '" & Formulario4 & "' or Acesso = '" & Formulario5 & "')", Conexao, adOpenKeyset, adLockOptimistic
        If TBAcessos.EOF = False Then Set nodA1 = TreeviewMnu.Nodes.Add(nodA, tvwChild, "A", "OpÁıes gerais", "B")
        
        'Usu·rios
        Set nodA2 = TreeviewMnu.Nodes.Add(nodA, tvwChild, , "Usu·rios", "A")
        
            Formulario = "ConfiguraÁ„o do sistema/Usu·rios"
            ProcLiberaAcessos False
            If Acessos = True Then Set nodA2A = TreeviewMnu.Nodes.Add(nodA2, tvwChild, "B", "Usu·rios", "B")
            
            Formulario = "ConfiguraÁ„o do sistema/Usu·rios/Conectados"
            ProcLiberaAcessos False
            If Acessos = True Then Set nodA2B = TreeviewMnu.Nodes.Add(nodA2, tvwChild, "DS", "Conectados", "B")
            
            Formulario = "ConfiguraÁ„o do sistema/Usu·rios/Eventos realizados"
            ProcLiberaAcessos False
            If Acessos = True Then Set nodA2C = TreeviewMnu.Nodes.Add(nodA2, tvwChild, "DT", "Eventos realizados", "B")
        
        'Criar backup
        Set nodA3 = TreeviewMnu.Nodes.Add(nodA, tvwChild, , "Criar backup", "A")
            
            Formulario = "ConfiguraÁ„o do sistema/Criar backup/ConfiguraÁıes"
            ProcLiberaAcessos False
            If Acessos = True Then Set nodA3D = TreeviewMnu.Nodes.Add(nodA3, tvwChild, "ET", "ConfiguraÁıes", "B")
            
            Formulario = "ConfiguraÁ„o do sistema/Criar backup/Apontamentos"
            ProcLiberaAcessos False
            If Acessos = True Then Set nodA3B = TreeviewMnu.Nodes.Add(nodA3, tvwChild, "DI", "Apontamentos", "B")
            
            Formulario = "ConfiguraÁ„o do sistema/Criar backup/Eventos"
            ProcLiberaAcessos False
            If Acessos = True Then Set nodA3B = TreeviewMnu.Nodes.Add(nodA3, tvwChild, "DK", "Eventos realizados por usu·rio", "B")
        
        'Reindexar
        Set nodA4 = TreeviewMnu.Nodes.Add(nodA, tvwChild, , "Reindexar BD", "A")
        
            Formulario = "ConfiguraÁ„o do sistema/Reindexar BD/Caprind e Gerprod"
            ProcLiberaAcessos False
            If Acessos = True Then Set nodA4A = TreeviewMnu.Nodes.Add(nodA4, tvwChild, "EB", "Caprind e Gerprod", "B")
            
            Formulario = "ConfiguraÁ„o do sistema/Reindexar BD/GNFe"
            ProcLiberaAcessos False
            If Acessos = True Then Set nodA4B = TreeviewMnu.Nodes.Add(nodA4, tvwChild, "EC", "GNFe", "B")
    
    'Administrativo
    Set nodB = TreeviewMnu.Nodes.Add(nodmenu, tvwChild, , "Administrativo", "A")
        
        'RH
        Set nodB1 = TreeviewMnu.Nodes.Add(nodB, tvwChild, , "RH", "A")
            
            Formulario = "RH/Funcion·rios"
            ProcLiberaAcessos False
            If Acessos = True Then Set nodB1A = TreeviewMnu.Nodes.Add(nodB1, tvwChild, "C", "Funcion·rios", "B")
            
            Set nodB1B = TreeviewMnu.Nodes.Add(nodB1, tvwChild, , "RelatÛrios", "A")
                
                Formulario = "RH/RelatÛrios/DesoneraÁ„o da folha de pagamento"
                ProcLiberaAcessos False
                If Acessos = True Then Set nodB1BA = TreeviewMnu.Nodes.Add(nodB1B, tvwChild, "EE", "DesoneraÁ„o da folha de pagamento", "B")
        
        'Compras
        Set nodB2 = TreeviewMnu.Nodes.Add(nodB, tvwChild, , "Compras", "A")
            Formulario = "Compras/FamÌlias"
            ProcLiberaAcessos False
            If Acessos = True Then Set nodB2A = TreeviewMnu.Nodes.Add(nodB2, tvwChild, "D", "FamÌlias", "B")
            
            Formulario = "Compras/Produtos e serviÁos"
            ProcLiberaAcessos False
            If Acessos = True Then Set nodB2B = TreeviewMnu.Nodes.Add(nodB2, tvwChild, "E", "Produtos e serviÁos", "B")
            
            Formulario = "Compras/Fornecedores"
            ProcLiberaAcessos False
            If Acessos = True Then Set nodB2C = TreeviewMnu.Nodes.Add(nodB2, tvwChild, "F", "Fornecedores", "B")
            
            Formulario = "Compras/ProgramaÁ„o"
            ProcLiberaAcessos False
            If Acessos = True Then Set nodB2F = TreeviewMnu.Nodes.Add(nodB2, tvwChild, "G", "ProgramaÁ„o", "B")
            
            Formulario = "Compras/CotaÁ„o"
            ProcLiberaAcessos False
            If Acessos = True Then Set nodB2G = TreeviewMnu.Nodes.Add(nodB2, tvwChild, "H", "CotaÁ„o", "B")
                                
            Set nodB2H = TreeviewMnu.Nodes.Add(nodB2, tvwChild, , "Pedido", "A")
                
                Formulario = "Compras/Pedido"
                ProcLiberaAcessos False
                If Acessos = True Then Set nodB2HA = TreeviewMnu.Nodes.Add(nodB2H, tvwChild, "I", "Gerenciar", "B")
                
                Formulario = "Compras/Pedido/Aprovar"
                ProcLiberaAcessos False
                If Acessos = True Then Set nodB2HB = TreeviewMnu.Nodes.Add(nodB2H, tvwChild, "DQ", "Aprovar", "B")
                
            Formulario = "Compras/Necessidade"
            ProcLiberaAcessos False
            If Acessos = True Then Set nodB2K = TreeviewMnu.Nodes.Add(nodB2, tvwChild, "K", "Necessidade", "B")
            
            Formulario = "Compras/N„o conformidade"
            ProcLiberaAcessos False
            If Acessos = True Then Set nodB2L = TreeviewMnu.Nodes.Add(nodB2, tvwChild, "L", "N„o conformidade", "B")
            
            Formulario = "Compras/AtualizaÁ„o de valores"
            ProcLiberaAcessos False
            If Acessos = True Then Set nodB2M = TreeviewMnu.Nodes.Add(nodB2, tvwChild, "EL", "AtualizaÁ„o de valores", "B")
            
            Set nodB2J = TreeviewMnu.Nodes.Add(nodB2, tvwChild, , "RelatÛrios", "A")
               
                Formulario = "Compras/RelatÛrios/HistÛrico"
                ProcLiberaAcessos False
                If Acessos = True Then Set nodB2JA = TreeviewMnu.Nodes.Add(nodB2J, tvwChild, "M", "HistÛrico", "B")
                
                Formulario = "Compras/RelatÛrios/Õndice de atraso"
                ProcLiberaAcessos False
                If Acessos = True Then Set nodB2JB = TreeviewMnu.Nodes.Add(nodB2J, tvwChild, "EN", "Õndice de atraso", "B")
        
        'Vendas
        Set nodB3 = TreeviewMnu.Nodes.Add(nodB, tvwChild, , "Vendas", "A")
            Formulario = "Vendas/FamÌlias"
            ProcLiberaAcessos False
            If Acessos = True Then Set nodB3A = TreeviewMnu.Nodes.Add(nodB3, tvwChild, "N", "FamÌlias", "B")
            
            Formulario = "Vendas/Produtos e serviÁos"
            ProcLiberaAcessos False
            If Acessos = True Then Set nodB3B = TreeviewMnu.Nodes.Add(nodB3, tvwChild, "O", "Produtos e serviÁos", "B")
            
            Formulario = "Vendas/Clientes"
            ProcLiberaAcessos False
            If Acessos = True Then Set nodB3C = TreeviewMnu.Nodes.Add(nodB3, tvwChild, "P", "Clientes", "B")
            
            Formulario = "Vendas/Vendedores"
            ProcLiberaAcessos False
            If Acessos = True Then Set nodB3E = TreeviewMnu.Nodes.Add(nodB3, tvwChild, "Q", "Vendedores", "B")
            
            Formulario = "Vendas/SimulaÁ„o"
            ProcLiberaAcessos False
            If Acessos = True Then Set nodB3Q = TreeviewMnu.Nodes.Add(nodB3, tvwChild, "EY", "SimulaÁ„o", "B")
            
            Formulario = "Vendas/Empenho"
            ProcLiberaAcessos False
            If Acessos = True Then Set nodB3L = TreeviewMnu.Nodes.Add(nodB3, tvwChild, "EH", "Empenho", "B")
            
            Formulario = "Vendas/Telemarketing"
            ProcLiberaAcessos False
            If Acessos = True Then Set nodB3F = TreeviewMnu.Nodes.Add(nodB3, tvwChild, "R", "Telemarketing", "B")
            
            Formulario = "Vendas/ProgramaÁ„o"
            ProcLiberaAcessos False
            If Acessos = True Then Set nodB3P = TreeviewMnu.Nodes.Add(nodB3, tvwChild, "DF", "ProgramaÁ„o", "B")
            
            Formulario = "Vendas/Proposta comercial"
            ProcLiberaAcessos False
            If Acessos = True Then Set nodB3G = TreeviewMnu.Nodes.Add(nodB3, tvwChild, "S", "Proposta comercial", "B")
            
            Formulario = "Vendas/Pedido interno"
            ProcLiberaAcessos False
            If Acessos = True Then Set nodB3H = TreeviewMnu.Nodes.Add(nodB3, tvwChild, "T", "Pedido interno", "B")
            
            Formulario = "Vendas/Follow up"
            ProcLiberaAcessos False
            If Acessos = True Then Set nodB3J = TreeviewMnu.Nodes.Add(nodB3, tvwChild, "U", "Carteira de vendas", "B")
            
            Formulario = "Vendas/SituaÁ„o da produÁ„o"
            ProcLiberaAcessos False
            If Acessos = True Then Set nodB3K = TreeviewMnu.Nodes.Add(nodB3, tvwChild, "W", "SituaÁ„o da produÁ„o", "B")
            
            Formulario = "Vendas/InformaÁıes faturamento"
            ProcLiberaAcessos False
            If Acessos = True Then Set nodB3M = TreeviewMnu.Nodes.Add(nodB3, tvwChild, "X", "InformaÁıes faturamento", "B")
            
            Formulario = "Vendas/AtualizaÁ„o de valores"
            ProcLiberaAcessos False
            If Acessos = True Then Set nodB3N = TreeviewMnu.Nodes.Add(nodB3, tvwChild, "EM", "AtualizaÁ„o de valores", "B")
            
            Set nodB3O = TreeviewMnu.Nodes.Add(nodB3, tvwChild, , "RelatÛrios", "A")
            
                Formulario = "Vendas/RelatÛrios/HistÛrico"
                ProcLiberaAcessos False
                If Acessos = True Then Set nodB3OA = TreeviewMnu.Nodes.Add(nodB3O, tvwChild, "Z", "HistÛrico", "B")
                
                Formulario = "Vendas/RelatÛrios/HistÛrico"
                ProcLiberaAcessos False
                If Acessos = True Then Set nodB3OA = TreeviewMnu.Nodes.Add(nodB3O, tvwChild, "AW", "Desempenho", "B")
                
                
                Formulario = "Vendas/RelatÛrios/Õndice de atraso"
                ProcLiberaAcessos False
                If Acessos = True Then Set nodb3OD = TreeviewMnu.Nodes.Add(nodB3O, tvwChild, "DJ", "Õndice de atraso", "B")
                
                Formulario = "Vendas/RelatÛrios/Comiss„o"
                ProcLiberaAcessos False
                If Acessos = True Then Set nodb3OE = TreeviewMnu.Nodes.Add(nodB3O, tvwChild, "DM", "Comiss„o", "B")
                                        
                Formulario = "Vendas/RelatÛrios/Comiss„o"
                ProcLiberaAcessos False
                If Acessos = True Then Set nodb3OE = TreeviewMnu.Nodes.Add(nodB3O, tvwChild, "DN", "Comissıes x meta", "B")
                                        
        'Financeiro
        Set nodB4 = TreeviewMnu.Nodes.Add(nodB, tvwChild, , "Financeiro", "A")
            
            Formulario = "Financeiro/Plano de contas"
            ProcLiberaAcessos False
            If Acessos = True Then Set nodB4AA = TreeviewMnu.Nodes.Add(nodB4, tvwChild, "AB", "Plano de contas", "B")
            
            Formulario = "Financeiro/InstituiÁıes"
            ProcLiberaAcessos False
            If Acessos = True Then Set nodB4AB = TreeviewMnu.Nodes.Add(nodB4, tvwChild, "AC", "InstituiÁıes Financeiras", "B")
            
            Formulario = "Financeiro/Contas a pagar"
            ProcLiberaAcessos False
            If Acessos = True Then Set nodB4C = TreeviewMnu.Nodes.Add(nodB4, tvwChild, "AE", "Contas a pagar", "B")
            
            Formulario = "Financeiro/Contas pagas"
            ProcLiberaAcessos False
            If Acessos = True Then Set nodB4D = TreeviewMnu.Nodes.Add(nodB4, tvwChild, "AF", "Contas pagas", "B")
            
            Formulario = "Financeiro/Contas a receber"
            ProcLiberaAcessos False
            If Acessos = True Then Set nodB4E = TreeviewMnu.Nodes.Add(nodB4, tvwChild, "AG", "Contas a receber", "B")
                        
            Formulario = "Financeiro/Contas recebidas"
            ProcLiberaAcessos False
            If Acessos = True Then Set nodB4F = TreeviewMnu.Nodes.Add(nodB4, tvwChild, "AH", "Contas recebidas", "B")
            
            Formulario = "Financeiro/Desconto de duplicata"
            ProcLiberaAcessos False
            If Acessos = True Then Set nodB4H = TreeviewMnu.Nodes.Add(nodB4, tvwChild, "AI", "Desconto de duplicata", "B")
            
            Formulario = "Financeiro/Fluxo de caixa"
            ProcLiberaAcessos False
            If Acessos = True Then
              Set nodB4G = TreeviewMnu.Nodes.Add(nodB4, tvwChild, "AJ", "Fluxo de caixa", "A")
'===================================================================================================================
                   Set nodB4G1 = TreeviewMnu.Nodes.Add(nodB4G, tvwChild, "AJ1", "Resumido gr·fico", "B")
                   Set nodB4G2 = TreeviewMnu.Nodes.Add(nodB4G, tvwChild, "AJ2", "Detalhado", "B")
'===================================================================================================================
            End If
            
            Set nodB4I = TreeviewMnu.Nodes.Add(nodB4, tvwChild, , "RelatÛrios", "A")
            
                Formulario = "Financeiro/RelatÛrios/HistÛrico"
                ProcLiberaAcessos False
                If Acessos = True Then Set nodB4IA = TreeviewMnu.Nodes.Add(nodB4I, tvwChild, "DE", "HistÛrico", "B")
                
                Formulario = "Financeiro/RelatÛrios/Raz„o"
                ProcLiberaAcessos False
                If Acessos = True Then
                Set nodB4IB = TreeviewMnu.Nodes.Add(nodB4I, tvwChild, "EF", "Raz„o", "B")
'===================================================================================================================
                   Set nodB4IB1 = TreeviewMnu.Nodes.Add(nodB4IB, tvwChild, "EF1", "Detalhado", "B")
                   Set nodB4IB2 = TreeviewMnu.Nodes.Add(nodB4IB, tvwChild, "EF2", "Resumido", "B")
'===================================================================================================================
                End If
        'Faturamento
        Set nodB5 = TreeviewMnu.Nodes.Add(nodB, tvwChild, , "Faturamento", "A")
              
           'Fiscal
           Set nodB5A = TreeviewMnu.Nodes.Add(nodB5, tvwChild, , "Fiscal", "A")
                
                Formulario = "Faturamento/Fiscal/ClassificaÁ„o fiscal"
                ProcLiberaAcessos False
                If Acessos = True Then Set nodB5AA = TreeviewMnu.Nodes.Add(nodB5A, tvwChild, "AK", "ClassificaÁ„o fiscal", "B")
                
                Formulario = "Faturamento/Fiscal/Natureza de operaÁ„o"
                ProcLiberaAcessos False
                If Acessos = True Then Set nodB5AB = TreeviewMnu.Nodes.Add(nodB5A, tvwChild, "AL", "Natureza de operaÁ„o", "B")
                          
                Formulario = "Faturamento/Fiscal/Natureza de operaÁ„o"
                ProcLiberaAcessos False
                If Acessos = True Then Set nodB5AC = TreeviewMnu.Nodes.Add(nodB5A, tvwChild, "KK", "Regiıes e SubstituiÁ„o tribut·ria", "B")
                          
            'Nota fiscal
            Set nodB5B = TreeviewMnu.Nodes.Add(nodB5, tvwChild, , "Nota fiscal", "A")
                
                Formulario = "Faturamento/Nota fiscal/Terceiros"
                ProcLiberaAcessos False
                If Acessos = True Then Set nodB5BA = TreeviewMnu.Nodes.Add(nodB5B, tvwChild, "AM", "Terceiros", "B")
                
                Formulario = "Faturamento/Nota fiscal/PrÛpria"
                ProcLiberaAcessos False
                If Acessos = True Then Set nodB5BB = TreeviewMnu.Nodes.Add(nodB5B, tvwChild, "AN", "PrÛpria", "B")
                
                Formulario = "Faturamento/Nota fiscal/SPED"
                ProcLiberaAcessos False
                If Acessos = True Then Set nodB5BC = TreeviewMnu.Nodes.Add(nodB5B, tvwChild, "DZ", "SPED", "B")
                
                Formulario = "Faturamento/Nota fiscal/Exportar"
                ProcLiberaAcessos False
                If Acessos = True Then Set nodB5BD = TreeviewMnu.Nodes.Add(nodB5B, tvwChild, "EA", "Exportar (.txt)", "B")
            
            Formulario = "Faturamento/Carta de correÁ„o"
            ProcLiberaAcessos False
            If Acessos = True Then Set nodB5C = TreeviewMnu.Nodes.Add(nodB5, tvwChild, "AO", "Carta de correÁ„o", "B")
            
            Formulario = "Faturamento/Minuta de despacho"
            ProcLiberaAcessos False
            If Acessos = True Then Set nodB5CA = TreeviewMnu.Nodes.Add(nodB5, tvwChild, "AP", "Minuta de despacho", "B")
            
            'RelatÛrios
            Set nodB5D = TreeviewMnu.Nodes.Add(nodB5, tvwChild, , "RelatÛrios", "A")
            
                Formulario = "Faturamento/RelatÛrios/HistÛrico"
                ProcLiberaAcessos False
                If Acessos = True Then Set nodB5DA = TreeviewMnu.Nodes.Add(nodB5D, tvwChild, "AQ", "HistÛrico", "B")
                
                Formulario = "Faturamento/RelatÛrios/Relacionamento de notas fiscais"
                ProcLiberaAcessos False
                If Acessos = True Then Set nodB5DB = TreeviewMnu.Nodes.Add(nodB5D, tvwChild, "AR", "Relacionamento de notas fiscais", "B")
                
                Formulario = "Faturamento/RelatÛrios/Impostos"
                ProcLiberaAcessos False
                If Acessos = True Then Set nodB5DC = TreeviewMnu.Nodes.Add(nodB5D, tvwChild, "AS", "Impostos", "B")
                
                Formulario = "Faturamento/RelatÛrios/Doze ˙ltimos meses"
                ProcLiberaAcessos False
                If Acessos = True Then Set nodB5DD = TreeviewMnu.Nodes.Add(nodB5D, tvwChild, "FB", "Doze ˙ltimos meses", "B")
                                
        'Custos
        Set nodB6 = TreeviewMnu.Nodes.Add(nodB, tvwChild, , "Custos", "A")
            
            Formulario = "Custos/Centro de custo"
            ProcLiberaAcessos False
            If Acessos = True Then Set nodB6C = TreeviewMnu.Nodes.Add(nodB6, tvwChild, "DO", "Centro de custo", "B")
            
            'RelatÛrios
            Set nodB6A = TreeviewMnu.Nodes.Add(nodB6, tvwChild, , "RelatÛrios", "A")
            
                Formulario = "Custos/RelatÛrios/Detalhado"
                ProcLiberaAcessos False
                If Acessos = True Then Set nodB6AA = TreeviewMnu.Nodes.Add(nodB6A, tvwChild, "AT", "Detalhado", "B")
                
                Formulario = "Custos/RelatÛrios/Resumido"
                ProcLiberaAcessos False
                If Acessos = True Then Set nodB6AB = TreeviewMnu.Nodes.Add(nodB6A, tvwChild, "AU", "Resumido", "B")
                
                Formulario = "Custos/RelatÛrios/Previsto x Realizado"
                ProcLiberaAcessos False
                If Acessos = True Then Set nodB6AC = TreeviewMnu.Nodes.Add(nodB6A, tvwChild, "DP", "Previsto x Realizado", "B")
                                                
    'Engenharia
    Set nodC = TreeviewMnu.Nodes.Add(nodmenu, tvwChild, , "Engenharia", "A")
        
        Formulario = "Engenharia/FamÌlias"
        ProcLiberaAcessos False
        If Acessos = True Then Set nodC2 = TreeviewMnu.Nodes.Add(nodC, tvwChild, "AV", "FamÌlias", "B")
        
        Formulario = "Engenharia/Produtos e serviÁos"
        ProcLiberaAcessos False
        If Acessos = True Then Set nodC3 = TreeviewMnu.Nodes.Add(nodC, tvwChild, "AX", "Produtos e serviÁos", "B")
        
        Formulario = "Engenharia/Conjuntos"
        ProcLiberaAcessos False
        If Acessos = True Then Set nodC5 = TreeviewMnu.Nodes.Add(nodC, tvwChild, "AZ", "Conjuntos", "B")
        
        Formulario = "Engenharia/Estrutura"
        ProcLiberaAcessos False
        
        If Acessos = True Then
        Set nodC6 = TreeviewMnu.Nodes.Add(nodC, tvwChild, "BA", "Estrutura", "A")
                   Set nodC61 = TreeviewMnu.Nodes.Add(nodC6, tvwChild, "BA1", "Completa", "B")
                   Set nodC62 = TreeviewMnu.Nodes.Add(nodC6, tvwChild, "BA2", "Resumida", "B")
        End If
        
        Formulario = "Engenharia/Controle de projetos"
        ProcLiberaAcessos False
        If Acessos = True Then Set nodC7 = TreeviewMnu.Nodes.Add(nodC, tvwChild, "BB", "Controle de projetos", "B")
                
        Formulario = "Engenharia/Processos"
        ProcLiberaAcessos False
        If Acessos = True Then Set nodC8 = TreeviewMnu.Nodes.Add(nodC, tvwChild, "BC", "Processos", "B")
            
        Formulario = "Engenharia/Normas"
        ProcLiberaAcessos False
        If Acessos = True Then Set nodC10 = TreeviewMnu.Nodes.Add(nodC, tvwChild, "DA", "Normas", "B")
    
    'PCP
    Set nodD = TreeviewMnu.Nodes.Add(nodmenu, tvwChild, , "PCP", "A")
        
        Formulario = "PCP/Postos de trabalho"
        ProcLiberaAcessos False
        If Acessos = True Then Set nodD1 = TreeviewMnu.Nodes.Add(nodD, tvwChild, "BE", "Postos de trabalho", "B")
        
        Formulario = "PCP/CÛdigos de trabalho"
        ProcLiberaAcessos False
        If Acessos = True Then Set nodD2 = TreeviewMnu.Nodes.Add(nodD, tvwChild, "BF", "CÛdigos de trabalho", "B")
        
        Formulario = "PCP/Carga de posto de trabalho"
        ProcLiberaAcessos False
        If Acessos = True Then Set nodD3 = TreeviewMnu.Nodes.Add(nodD, tvwChild, "BG", "Carga de posto de trabalho", "B")
        
        Formulario = "PCP/Gerenciamento de ordem"
        ProcLiberaAcessos False
        If Acessos = True Then Set nodD4 = TreeviewMnu.Nodes.Add(nodD, tvwChild, "BH", "Gerenciamento de ordem", "B")
                
        Formulario = "PCP/Monitor de trabalho"
        ProcLiberaAcessos False
        If Acessos = True Then Set nodD6 = TreeviewMnu.Nodes.Add(nodD, tvwChild, "BI", "Monitor de trabalho", "B")
        
        Formulario = "PCP/SituaÁ„o da produÁ„o"
        ProcLiberaAcessos False
        If Acessos = True Then Set nodD7 = TreeviewMnu.Nodes.Add(nodD, tvwChild, "BJ", "SituaÁ„o da produÁ„o", "B")
        
        Formulario = "PCP/Programas CNC"
        ProcLiberaAcessos False
        If Acessos = True Then Set nodD8 = TreeviewMnu.Nodes.Add(nodD, tvwChild, "BK", "Programas CNC", "B")
        
        Formulario = "PCP/Necessidade"
        ProcLiberaAcessos False
        If Acessos = True Then Set nodD12 = TreeviewMnu.Nodes.Add(nodD, tvwChild, "DU", "Necessidade", "B")
        
        Formulario = "PCP/N„o conformidade"
        ProcLiberaAcessos False
        If Acessos = True Then Set nodD10 = TreeviewMnu.Nodes.Add(nodD, tvwChild, "BL", "N„o conformidade", "B")
        
        Formulario = "PCP/ProgramaÁ„o da produÁ„o"
        ProcLiberaAcessos False
        If Acessos = True Then Set nodD11 = TreeviewMnu.Nodes.Add(nodD, tvwChild, "DL", "ProgramaÁ„o da produÁ„o", "B")
                
        Formulario = "PCP/Plano da produÁ„o"
        ProcLiberaAcessos False
        If Acessos = True Then Set nodD13 = TreeviewMnu.Nodes.Add(nodD, tvwChild, "EO", "Plano da produÁ„o", "B")
        
        Formulario = "PCP/RequisiÁ„o da ordem"
        ProcLiberaAcessos False
        If Acessos = True Then Set nodD14 = TreeviewMnu.Nodes.Add(nodD, tvwChild, "ER", "RequisiÁ„o da ordem", "B")
        
        'RelatÛrios
        Set nodD9 = TreeviewMnu.Nodes.Add(nodD, tvwChild, , "RelatÛrios", "A")
            
            Formulario = "PCP/RelatÛrios/Produtividade"
            ProcLiberaAcessos False
            If Acessos = True Then Set nodD9A = TreeviewMnu.Nodes.Add(nodD9, tvwChild, "BM", "Produtividade", "B")
            
            Formulario = "PCP/RelatÛrios/N„o conformidade"
            ProcLiberaAcessos False
            If Acessos = True Then Set nodD9E = TreeviewMnu.Nodes.Add(nodD9, tvwChild, "BN", "N„o conformidade", "B")
               
            Formulario = "PCP/RelatÛrios/Monitor de eventos"
            ProcLiberaAcessos False
            If Acessos = True Then Set nodD9F = TreeviewMnu.Nodes.Add(nodD9, tvwChild, "BO", "Monitor de eventos", "B")
            
            Formulario = "PCP/RelatÛrios/Õndice de atraso"
            ProcLiberaAcessos False
            If Acessos = True Then Set nodD9G = TreeviewMnu.Nodes.Add(nodD9, tvwChild, "BP", "Õndice de atraso", "B")
            
            Formulario = "PCP/RelatÛrios/Resultados da ordem"
            ProcLiberaAcessos False
            If Acessos = True Then Set nodD9H = TreeviewMnu.Nodes.Add(nodD9, tvwChild, "CX", "Resultados da ordem", "B")
            
    'Qualidade
    Set Node = TreeviewMnu.Nodes.Add(nodmenu, tvwChild, , "Qualidade", "A")
        
        Formulario = "Qualidade/FamÌlias"
        ProcLiberaAcessos False
        If Acessos = True Then Set nodE1 = TreeviewMnu.Nodes.Add(Node, tvwChild, "BQ", "FamÌlias", "B")
        
        Formulario = "Qualidade/Instrumentos"
        ProcLiberaAcessos False
        If Acessos = True Then Set nodE1 = TreeviewMnu.Nodes.Add(Node, tvwChild, "BR", "Instrumentos", "B")
        
        Formulario = "Qualidade/Almoxarifado"
        ProcLiberaAcessos False
        If Acessos = True Then Set nodE15 = TreeviewMnu.Nodes.Add(Node, tvwChild, "EI", "Almoxarifado", "B")
        
        Formulario = "Qualidade/Plano de inspeÁ„o"
        ProcLiberaAcessos False
        If Acessos = True Then Set nodE2 = TreeviewMnu.Nodes.Add(Node, tvwChild, "BS", "Plano de inspeÁ„o", "B")
        
        Formulario = "Qualidade/Controle de mediÁ„o"
        ProcLiberaAcessos False
        If Acessos = True Then Set nodE3 = TreeviewMnu.Nodes.Add(Node, tvwChild, "BT", "Controle de mediÁ„o", "B")
        
        Formulario = "Qualidade/InspeÁ„o de recebimento"
        ProcLiberaAcessos False
        If Acessos = True Then Set nodE4 = TreeviewMnu.Nodes.Add(Node, tvwChild, "BU", "InspeÁ„o de recebimento", "B")
        
        'Ensaios
        Set nodE4A = TreeviewMnu.Nodes.Add(Node, tvwChild, , "Ensaios", "A")
            
            Formulario = "Qualidade/Ensaios/Ultra-som"
            ProcLiberaAcessos False
            If Acessos = True Then Set nodE4AA = TreeviewMnu.Nodes.Add(nodE4A, tvwChild, "BW", "Ultra-som (Em manutenÁ„o)", "B")
            
            Formulario = "Qualidade/Ensaios/LÌquido penetrante"
            ProcLiberaAcessos False
            If Acessos = True Then Set nodE4AB = TreeviewMnu.Nodes.Add(nodE4A, tvwChild, "BV", "LÌquido penetrante (Em manutenÁ„o)", "B")
            
            Formulario = "Qualidade/Ensaios/Controle de certificados"
            ProcLiberaAcessos False
            If Acessos = True Then Set nodE4AC = TreeviewMnu.Nodes.Add(nodE4A, tvwChild, "BX", "Controle de certificados (Em manutenÁ„o)", "B")
        
        Formulario = "Qualidade/Controle de certificados"
        ProcLiberaAcessos False
        If Acessos = True Then Set nodE5 = TreeviewMnu.Nodes.Add(Node, tvwChild, "BZ", "Controle de certificados", "B")
        
        Formulario = "Qualidade/Controle de documentos e dados"
        ProcLiberaAcessos False
        If Acessos = True Then Set nodE14 = TreeviewMnu.Nodes.Add(Node, tvwChild, "DG", "Controle de documentos e dados", "B")
        
        'N„o conformidade
        Set nodE6 = TreeviewMnu.Nodes.Add(Node, tvwChild, , "N„o conformidades", "A")
            
            Formulario = "Qualidade/N„o conformidade/DescriÁ„o da n„o conformidade"
            ProcLiberaAcessos False
            If Acessos = True Then Set nodE6A = TreeviewMnu.Nodes.Add(nodE6, tvwChild, "FA", "DescriÁ„o da n„o conformidade", "B")
            
            Formulario = "Qualidade/N„o conformidade"
            ProcLiberaAcessos False
            If Acessos = True Then Set nodE6B = TreeviewMnu.Nodes.Add(nodE6, tvwChild, "CA", "Gerenciamento", "B")
        
        Formulario = "Qualidade/SolicitaÁ„o de aÁ„o"
        ProcLiberaAcessos False
        If Acessos = True Then Set nodE7 = TreeviewMnu.Nodes.Add(Node, tvwChild, "CB", "SolicitaÁ„o de aÁ„o", "B")
        
        Formulario = "Qualidade/solicitaÁ„o de desvio"
        ProcLiberaAcessos False
        If Acessos = True Then Set nodE8 = TreeviewMnu.Nodes.Add(Node, tvwChild, "CC", "SolicitaÁ„o de desvio", "B")
        
        Formulario = "Qualidade/RNC"
        ProcLiberaAcessos False
        If Acessos = True Then Set nodE9 = TreeviewMnu.Nodes.Add(Node, tvwChild, "CD", "RNC", "B")
        
        'PPAP
        Set nodE411 = TreeviewMnu.Nodes.Add(Node, tvwChild, , "PPAP", "A")
            Formulario = "Qualidade/PPAP/PSW"
            ProcLiberaAcessos False
            If Acessos = True Then Set nodE11A = TreeviewMnu.Nodes.Add(nodE411, tvwChild, "CE", "PSW (Em manutenÁ„o)", "B")
            
            Formulario = "Qualidade/PPAP/FMEA"
            ProcLiberaAcessos False
            If Acessos = True Then Set nodE11B = TreeviewMnu.Nodes.Add(nodE411, tvwChild, "DB", "FMEA (Em manutenÁ„o)", "B")
            
            Formulario = "Qualidade/PPAP/Plano de controle"
            ProcLiberaAcessos False
            If Acessos = True Then Set nodE411C = TreeviewMnu.Nodes.Add(nodE411, tvwChild, "DC", "Plano de controle (Em manutenÁ„o)", "B")
            
        Formulario = "Qualidade/HistÛrico de revis„o dos relatÛrios"
        ProcLiberaAcessos False
        If Acessos = True Then Set nodE13 = TreeviewMnu.Nodes.Add(Node, tvwChild, "DD", "HistÛrico de revis„o dos relatÛrios", "B")
            
        'RelatÛrios
        Set nodE12 = TreeviewMnu.Nodes.Add(Node, tvwChild, , "RelatÛrios", "A")
            
            Formulario = "Qualidade/RelatÛrios/N„o conformidade"
            ProcLiberaAcessos False
            If Acessos = True Then Set nodE12A = TreeviewMnu.Nodes.Add(nodE12, tvwChild, "CF", "N„o conformidade", "B")
            
            Formulario = "Qualidade/RelatÛrios/HistÛrico"
            ProcLiberaAcessos False
            If Acessos = True Then Set nodE12B = TreeviewMnu.Nodes.Add(nodE12, tvwChild, "ED", "HistÛrico", "B")
            
    'Estoque
    Set nodF = TreeviewMnu.Nodes.Add(nodmenu, tvwChild, , "Estoque", "A")
    
        Formulario = "Estoque/Almoxarifado"
        ProcLiberaAcessos False
        If Acessos = True Then Set nodF7 = TreeviewMnu.Nodes.Add(nodF, tvwChild, "CG", "Almoxarifado", "B")
                        
        Formulario = "Estoque/Local de armazenamento"
        ProcLiberaAcessos False
        If Acessos = True Then Set nodF1 = TreeviewMnu.Nodes.Add(nodF, tvwChild, "CH", "Local de armazenamento", "B")
        
        Formulario = "Estoque/RequisiÁ„o de materiais"
        ProcLiberaAcessos False
        If Acessos = True Then Set nodF5 = TreeviewMnu.Nodes.Add(nodF, tvwChild, "CI", "RequisiÁ„o de materiais", "B")
        
        'Recebimento
        Set nodF2 = TreeviewMnu.Nodes.Add(nodF, tvwChild, , "Recebimento", "A")
            
            Formulario = "Estoque/Recebimento/Pedido de compra"
            ProcLiberaAcessos False
            If Acessos = True Then Set nodF2A = TreeviewMnu.Nodes.Add(nodF2, tvwChild, "CJ", "Pedido de compra", "B")
            
            Formulario = "Estoque/Recebimento/ConsignaÁ„o"
            ProcLiberaAcessos False
            If Acessos = True Then Set nodF2B = TreeviewMnu.Nodes.Add(nodF2, tvwChild, "CK", "ConsignaÁ„o", "B")
            
            Formulario = "Estoque/Recebimento/Importar nota de terceiros"
            ProcLiberaAcessos False
            If Acessos = True Then Set nodF2C = TreeviewMnu.Nodes.Add(nodF2, tvwChild, "ES", "Importar nota de terceiros", "B")
        
        Formulario = "Estoque/Invent·rio"
        ProcLiberaAcessos False
        If Acessos = True Then Set nodF8 = TreeviewMnu.Nodes.Add(nodF, tvwChild, "EJ", "Invent·rio", "B")
        
        Formulario = "Estoque/MovimentaÁ„o"
        ProcLiberaAcessos False
        If Acessos = True Then Set nodF3 = TreeviewMnu.Nodes.Add(nodF, tvwChild, "CM", "MovimentaÁ„o", "B")
        
        Formulario = "Estoque/Necessidade"
        ProcLiberaAcessos False
        If Acessos = True Then Set nodF9 = TreeviewMnu.Nodes.Add(nodF, tvwChild, "EK", "Necessidade", "B")
        
        Formulario = "Estoque/Ordem de faturamento"
        ProcLiberaAcessos False
        If Acessos = True Then Set nodF6 = TreeviewMnu.Nodes.Add(nodF, tvwChild, "EG", "Ordem de faturamento", "B")
        
        Formulario = "Estoque/Nota fiscal"
        ProcLiberaAcessos False
        If Acessos = True Then Set nodF4 = TreeviewMnu.Nodes.Add(nodF, tvwChild, "CN", "Nota fiscal", "B")
        
'        Formulario = "Estoque/Material terceiro"
'        ProcLiberaAcessos False
'        If Acessos = False Then Set nodF10 = TreeviewMnu.Nodes.Add(nodF, tvwChild, "MT", "Material Terceiro", "B")
 'Estoque relatÛrios
'        Formulario = "Estoque/RelatÛrios"
'        'ProcLiberaAcessos False
'        If Acessos = False Then Set nodF11 = TreeviewMnu.Nodes.Add(nodF, tvwChild, "MI", "RelatÛrios", "B")
    
    'ManutenÁ„o
    Set nodH = TreeviewMnu.Nodes.Add(nodmenu, tvwChild, , "ManutenÁ„o", "A")
    
        Formulario = "ManutenÁ„o/Equipamentos"
        ProcLiberaAcessos False
        If Acessos = True Then Set nodH1 = TreeviewMnu.Nodes.Add(nodH, tvwChild, "CO", "Gerenciamento", "B")
        
        'RelatÛrios
        Set nodH3 = TreeviewMnu.Nodes.Add(nodH, tvwChild, , "RelatÛrios", "A")
        
            Formulario = "ManutenÁ„o/RelatÛrios/HistÛrico"
            ProcLiberaAcessos False
            If Acessos = True Then Set nodH4 = TreeviewMnu.Nodes.Add(nodH3, tvwChild, "CZ", "HistÛrico", "B")
    
    'Outros
    Set nodG = TreeviewMnu.Nodes.Add(nodmenu, tvwChild, , "Outros", "A")
        
        Formulario = "Outros/SolicitaÁ„o"
        ProcLiberaAcessos False
        If Acessos = True Then Set nodGA = TreeviewMnu.Nodes.Add(nodG, tvwChild, "CQ", "SolicitaÁ„o de compra", "B")
        
        Formulario = "Outros/SolicitaÁ„o de produÁ„o"
        ProcLiberaAcessos False
        If Acessos = True Then Set nodGA = TreeviewMnu.Nodes.Add(nodG, tvwChild, "EQ", "SolicitaÁ„o de produÁ„o", "B")
            
        Formulario = "Outros/Follow up de compras"
        ProcLiberaAcessos False
        If Acessos = True Then Set nodGB = TreeviewMnu.Nodes.Add(nodG, tvwChild, "J", "Follow up de compras", "B")
        
        Formulario = "Outros/ValidaÁ„o de procedimentos"
        ProcLiberaAcessos False
        If Acessos = True Then Set nodGC = TreeviewMnu.Nodes.Add(nodG, tvwChild, "CS", "ValidaÁ„o de procedimentos", "B")
        
        Formulario = "Outros/An·lise crÌtica"
        ProcLiberaAcessos False
        If Acessos = True Then Set nodGD = TreeviewMnu.Nodes.Add(nodG, tvwChild, "CT", "An·lise crÌtica", "B")
        
        Set nodGE = TreeviewMnu.Nodes.Add(nodG, tvwChild, "EU", "Calculadora", "B")
        
        'Downloads
        Set nodGF = TreeviewMnu.Nodes.Add(nodG, tvwChild, , "Downloads", "A")
            Formulario = "Outros/Downloads/Nota fiscal"
            ProcLiberaAcessos False
            If Acessos = True Then Set nodGFA = TreeviewMnu.Nodes.Add(nodGF, tvwChild, "EV", "Nota fiscal", "B")
            
            Formulario = "Outros/Downloads/Boleto"
            ProcLiberaAcessos False
            If Acessos = True Then Set nodGFB = TreeviewMnu.Nodes.Add(nodGF, tvwChild, "EX", "Boleto", "B")
        
    'Suporte
    Set nodI = TreeviewMnu.Nodes.Add(nodmenu, tvwChild, , "Suporte", "A")
        
        Formulario = "Suporte/Chamado"
        ProcLiberaAcessos False
        If Acessos = True Then Set nodIA = TreeviewMnu.Nodes.Add(nodI, tvwChild, "CR", "Chamado", "B")
        
        Formulario = "Suporte/Chat (online)"
        ProcLiberaAcessos False
        If Acessos = True Then Set nodIB = TreeviewMnu.Nodes.Add(nodI, tvwChild, "DR", "Chat (online)", "B")
        
        Formulario = "Suporte/SolicitaÁ„o de atendimento"
        ProcLiberaAcessos False
        If Acessos = True Then Set nodIC = TreeviewMnu.Nodes.Add(nodI, tvwChild, "EP", "SolicitaÁ„o de atendimento (Em desenvolvimento)", "B")
        
        nodID = TreeviewMnu.Nodes.Add(nodI, tvwChild, "EZ", "Download Team Viewer 9", "B")
        
        'AtualizaÁ„o
        Set nodK = TreeviewMnu.Nodes.Add(nodI, tvwChild, , "AtualizaÁ„o", "A")
            
            Formulario = "Suporte/AtualizaÁ„o/Caprind e Gerprod"
            ProcLiberaAcessos False
            If Acessos = True Then Set nodKA = TreeviewMnu.Nodes.Add(nodK, tvwChild, "DV", "Caprind e Gerprod", "B")
                    
    'Finalizar
    Set nodJ = TreeviewMnu.Nodes.Add(nodmenu, tvwChild, "", "Finalizar", "A")
        
        Set nodJB = TreeviewMnu.Nodes.Add(nodJ, tvwChild, "CW", "Fazer logoff de " & pubUsuario, "B")
        
        Set nodJC = TreeviewMnu.Nodes.Add(nodJ, tvwChild, "CV", "Sair", "B")

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcAbreModuloMenuTreeView(Node As String)
On Error GoTo tratar_erro

'=======================================================================
'Ultimo È "FB" Faturamento/RelatÛrios/Doze ˙ltimos meses
'=======================================================================

Select Case Node
    'ConfiguraÁ„o do sistema
    Case "A": frmOpcoesGeral.Show
    Case "B": frmUsuarios.Show
    Case "DS": frmOpcoes_Lista_usuarios.Show
    Case "DT": frmOpcoes_Lista_eventos.Show
    Case "ET": Frm_configuracao_backup.Show 1
    Case "DI": frmBackup_apontamentos.Show 1
    Case "DK": frmMDI.ProcCriarBackupEventos
    Case "EB":
        If USMsgBox("Deseja realmente reindexar o BD do Caprind e Gerprod?", vbYesNo, "CAPRIND v5.0") = vbYes Then
            If FunVerificaUsuariosConectados(pubUsuario) = False Then
                USMsgBox ("N„o È permitido reindexar o BD, pois outros usu·rios est„o utilizando o sistema."), vbExclamation, "CAPRIND v5.0"
            Else
                ProcReindexarBDCaprindeGerprod
            End If
        End If
    Case "EC":
        If USMsgBox("Deseja realmente reindexar o BD do GNFe?", vbYesNo, "CAPRIND v5.0") = vbYes Then
            If FunVerificaUsuariosConectados(pubUsuario) = False Then
                USMsgBox ("N„o È permitido reindexar o BD, pois outros usu·rios est„o utilizando o sistema."), vbExclamation, "CAPRIND v5.0"
            Else
                ProcReindexarBDGNFe
            End If
        End If
    
    'RH
    Case "C":
        If FunVefificaModuloLocacao(True, False, False) = False Then Exit Sub
        frmRH_Funcionarios.Show
    Case "EE":
        If FunVefificaModuloLocacao(True, False, False) = False Then Exit Sub
        FrmRH_Relatorios_Desoneracao.Show 1
    
    'Compras
    Case "D":
        If FunVefificaModuloLocacao(True, False, True) = False Then Exit Sub
        If Formulario_familia = "Compras/FamÌlias" Or Formulario_familia = "Vendas/FamÌlias" Or Formulario_familia = "Engenharia/FamÌlias" Or Formulario_familia = "Qualidade/FamÌlias" Then
            If FunVerifFormAberto(frmproj_familia) = True Then
                If USMsgBox("O mÛdulo " & Formulario_familia & " est· aberto, deseja fech·-lo para prosseguir?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub Else Unload frmproj_familia
            End If
        End If
        Compras_Familia = True
        Vendas_Familia = False
        Qualidade_Familia = False
        frmproj_familia.Show
    Case "E":
        If FunVefificaModuloLocacao(True, False, True) = False Then Exit Sub
        If Formulario_produto = "Compras/Produtos e serviÁos" Or Formulario_produto = "Vendas/Produtos e serviÁos" Or Formulario_produto = "Engenharia/Produtos e serviÁos" Then
            If FunVerifFormAberto(frmproj_produto) = True Then
                If USMsgBox("O mÛdulo " & Formulario_produto & " est· aberto, deseja fech·-lo para prosseguir?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub Else Unload frmproj_produto
            End If
        End If
        Engenharia_Produtos = False
        Compras_Produtos = True
        Vendas_Produtos = False
        frmproj_produto.Show
    Case "F":
        If FunVefificaModuloLocacao(True, False, True) = False Then Exit Sub
        frmCompras_fornecedores.Show
    Case "G":
        If FunVefificaModuloLocacao(True, False, True) = False Then Exit Sub
        Programacao = True
        frmCompras_programacao.Show
    Case "H":
        If FunVefificaModuloLocacao(True, False, True) = False Then Exit Sub
        frmcompras_reqcot.Show
    Case "I":
        If FunVefificaModuloLocacao(True, False, True) = False Then Exit Sub
        Programacao = False
        frmCompras_Pedido.Show
    Case "K":
        If FunVefificaModuloLocacao(True, False, True) = False Then Exit Sub
        If Formulario_necessidade = "Compras/Necessidade" Or Formulario_necessidade = "PCP/Necessidade" Or Formulario_necessidade = "Estoque/Necessidade" Then
            If FunVerifFormAberto(Frm_necessidade) = True Then
                If USMsgBox("O mÛdulo " & Formulario_necessidade & " est· aberto, deseja fech·-lo para prosseguir?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub Else Unload Frm_necessidade
            End If
        End If
        Compras_Necessidade = True
        PCP_Necessidade = False
        Frm_necessidade.Show
    Case "DQ":
        If FunVefificaModuloLocacao(True, False, True) = False Then Exit Sub
        frmCompras_Aprovar_Pedido.Show
    Case "L":
        If FunVefificaModuloLocacao(True, False, True) = False Then Exit Sub
        frmCompras_NaoConformidade.Show
    Case "EL":
        If FunVefificaModuloLocacao(True, False, True) = False Then Exit Sub
        Vendas_AtualizaÁ„o_Valores = False
        Frm_atualizacao_valores.Show
    Case "M":
        If FunVefificaModuloLocacao(True, False, True) = False Then Exit Sub
        frmCompras_Relatorios_Historico2.Show
    Case "EN":
        If FunVefificaModuloLocacao(True, False, True) = False Then Exit Sub
        frmCompras_Relatorios_Indice_Atraso.Show
    
    'Vendas
    Case "N":
        If FunVefificaModuloLocacao(True, True, True) = False Then Exit Sub
        If Formulario_familia = "Compras/FamÌlias" Or Formulario_familia = "Vendas/FamÌlias" Or Formulario_familia = "Engenharia/FamÌlias" Or Formulario_familia = "Qualidade/FamÌlias" Then
            If FunVerifFormAberto(frmproj_familia) = True Then
                If USMsgBox("O mÛdulo " & Formulario_familia & " est· aberto, deseja fech·-lo para prosseguir?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub Else Unload frmproj_familia
            End If
        End If
        Compras_Familia = False
        Vendas_Familia = True
        Qualidade_Familia = False
        frmproj_familia.Show
    Case "O":
        If FunVefificaModuloLocacao(True, True, True) = False Then Exit Sub
        If Formulario_produto = "Compras/Produtos e serviÁos" Or Formulario_produto = "Vendas/Produtos e serviÁos" Or Formulario_produto = "Engenharia/Produtos e serviÁos" Then
            If FunVerifFormAberto(frmproj_produto) = True Then
                If USMsgBox("O mÛdulo " & Formulario_produto & " est· aberto, deseja fech·-lo para prosseguir?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub Else Unload frmproj_produto
            End If
        End If
        Engenharia_Produtos = False
        Compras_Produtos = False
        Vendas_Produtos = True
        frmproj_produto.Show
    Case "P":
        If FunVefificaModuloLocacao(True, True, True) = False Then Exit Sub
        frmVendas_cliente.Show
    Case "Q":
        If FunVefificaModuloLocacao(True, True, True) = False Then Exit Sub
        frmVendas_Vendedores.Show
    Case "EY":
        If FunVefificaModuloLocacao(True, True, True) = False Then Exit Sub
        If FunVerifAtualizacaoObrigatoria(True, False) = False Then frmvendas_simulacao.Show
    Case "EH":
        If FunVefificaModuloLocacao(True, True, True) = False Then Exit Sub
        frmVendas_Empenho.Show
    Case "R":
        If FunVefificaModuloLocacao(True, True, True) = False Then Exit Sub
        frmVendas_Tele_Clientes.Show
    Case "DF":
        If FunVefificaModuloLocacao(True, True, True) = False Then Exit Sub
        frmVendas_programacao.Show
    Case "S":
        If FunVefificaModuloLocacao(True, True, True) = False Then Exit Sub
        If FunVerifAtualizacaoObrigatoria(True, False) = False Then
            Vendas_PI = False
            Vendas_Proposta = True
            frmVendas_proposta.Show
        End If
    Case "T":
        If FunVefificaModuloLocacao(True, True, True) = False Then Exit Sub
        If FunVerifAtualizacaoObrigatoria(True, False) = False Then
            Vendas_PI = True
            Vendas_Proposta = False
            frmVendas_PI.Show
        End If
    Case "U":
        If FunVefificaModuloLocacao(True, True, True) = False Then Exit Sub
        frmVendas_carteira.Show
    Case "W":
        If FunVefificaModuloLocacao(True, True, True) = False Then Exit Sub
        Vendas = True
        FrmSituacao_pedido_producao.Show
    Case "EM":
        If FunVefificaModuloLocacao(True, True, True) = False Then Exit Sub
        Vendas_AtualizaÁ„o_Valores = True
        Frm_atualizacao_valores.Show
    'Case "DN":
    Case "X":
        If FunVefificaModuloLocacao(True, True, True) = False Then Exit Sub
        Faturamento = False
        frmFaturamento_Relatorios.Show
    Case "Z":
        If FunVefificaModuloLocacao(True, True, True) = False Then Exit Sub
        frmVendas_Relatorios_Historico.Show
    Case "AW":
        If FunVefificaModuloLocacao(True, True, True) = False Then Exit Sub
        frmVendas_Desempenho.Show
    Case "DJ":
        If FunVefificaModuloLocacao(True, True, True) = False Then Exit Sub
        frmVendas_Relatorios_Indice_Atraso.Show
    Case "DM":
        If FunVefificaModuloLocacao(True, True, True) = False Then Exit Sub
        frmVendas_comissao.Show
    Case "DN":
        If FunVefificaModuloLocacao(True, True, True) = False Then Exit Sub
        frmVendas_Comissoes_Metas.Show
        
    
    'Financeiro
    'Case "AA":
    Case "AB":
        If FunVefificaModuloLocacao(True, False, False) = False Then Exit Sub
        frmFinanceiro_familia.Show
    Case "AC":
        If FunVefificaModuloLocacao(True, False, False) = False Then Exit Sub
        frm_Instituicoes.Show
    Case "AE":
        If FunVefificaModuloLocacao(True, False, False) = False Then Exit Sub
        frmContas_Pagar.Show
    Case "AF":
        If FunVefificaModuloLocacao(True, False, False) = False Then Exit Sub
        frmContas_Pagas.Show
    Case "AG":
        If FunVefificaModuloLocacao(True, False, False) = False Then Exit Sub
        frmContas_Receber.Show
    Case "AH":
        If FunVefificaModuloLocacao(True, False, False) = False Then Exit Sub
        frmContas_recebidas.Show
    Case "AI":
        If FunVefificaModuloLocacao(True, False, False) = False Then Exit Sub
        frm_trocaduplicata.Show
    Case "AJ1":
        If FunVefificaModuloLocacao(True, False, False) = False Then Exit Sub
        frmFluxo_Caixa_Grafico.Show
    Case "AJ2":
        If FunVefificaModuloLocacao(True, False, False) = False Then Exit Sub
        frmFluxodecaixa.Show
    Case "DE":
        If FunVefificaModuloLocacao(True, False, False) = False Then Exit Sub
        frmFinanceiro_Relatorios.Show
    Case "EF1":
        If FunVefificaModuloLocacao(True, False, False) = False Then Exit Sub
        frmFinanceiro_Relatorios_Razao_Detalhado.Show
    Case "EF2":
        If FunVefificaModuloLocacao(True, False, False) = False Then Exit Sub
        frmFinanceiro_Relatorios_Razao.Show
        
    'Faturamento
    Case "AK":
        If FunVefificaModuloLocacao(True, False, True) = False Then Exit Sub
        frm_Classificacao_Fiscal.Show
    Case "AL":
        If FunVefificaModuloLocacao(True, False, True) = False Then Exit Sub
        frm_Natureza_OP.Show
    Case "KK":
        If FunVefificaModuloLocacao(True, False, True) = False Then Exit Sub
        frmSubstituicao_Tributaria.Show 1
        
    Case "AM":
        'If FunVefificaModuloLocacao(True, False, True) = False Then Exit Sub
        If Formulario_nota = "Faturamento/Nota fiscal/PrÛpria" Or Formulario_nota = "Estoque/Ordem de faturamento" Or Formulario_nota = "Estoque/Nota fiscal" Then
            If FunVerifFormAberto(frmFaturamento_Prod_Serv) = True Then
                If USMsgBox("O mÛdulo " & Formulario_nota & " est· aberto, deseja fech·-lo para prosseguir?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub Else Unload frmFaturamento_Prod_Serv
            End If
        End If
        Formulario = "Faturamento/Nota fiscal/Terceiros"
        Faturamento_NF_Saida = False
        '======================================================
        Faturamento_NF_Terceiro = True
        Faturamento_NF_Propria = False
        '======================================================
        frmFaturamento_Prod_Serv.Show
    Case "AN"
        'If FunVefificaModuloLocacao(True, False, True) = False Then Exit Sub
        If Formulario_nota = "Faturamento/Nota fiscal/Terceiros" Or Formulario_nota = "Estoque/Ordem de faturamento" Or Formulario_nota = "Estoque/Nota fiscal" Then
            If FunVerifFormAberto(frmFaturamento_Prod_Serv) = True Then
                If USMsgBox("O mÛdulo " & Formulario_nota & " est· aberto, deseja fech·-lo para prosseguir?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub Else Unload frmFaturamento_Prod_Serv
            End If
        End If
        Formulario = "Faturamento/Nota fiscal/PrÛpria"
        Faturamento_NF_Saida = True
        '======================================================
        Faturamento_NF_Terceiro = False
        Faturamento_NF_Propria = True
        '======================================================
        frmFaturamento_Prod_Serv.Show
    Case "DZ":
        'If FunVefificaModuloLocacao(True, False, True) = False Then Exit Sub
        frmFaturamento_Prod_serv_SPED.Show
    Case "EA":
        'If FunVefificaModuloLocacao(True, False, True) = False Then Exit Sub
        frmFaturamento_Prod_serv_exportar.Show 1
    Case "AO":
        'If FunVefificaModuloLocacao(True, False, True) = False Then Exit Sub
        frmFaturamento_CartaCorrecao_NS.Show
    Case "AP":
        If FunVefificaModuloLocacao(True, False, True) = False Then Exit Sub
        frmMinuta.Show
    Case "AQ":
        If FunVefificaModuloLocacao(True, False, True) = False Then Exit Sub
        Faturamento = True
        frmFaturamento_Relatorios.Show
    Case "AR":
        If FunVefificaModuloLocacao(True, False, True) = False Then Exit Sub
        frmFaturamento_Relatorios_Relacionamento.Show
    Case "AS":
        If FunVefificaModuloLocacao(True, False, True) = False Then Exit Sub
        frmFaturamento_Relatorios_Impostos.Show
    Case "FB":
        If FunVefificaModuloLocacao(True, False, True) = False Then Exit Sub
        frmFaturamento_12ultimos_meses.Show 1
    
    'Custos
    Case "DO":
        If FunVefificaModuloLocacao(True, False, False) = False Then Exit Sub
        Frm_centro_de_custo.Show
    Case "AT":
        If FunVefificaModuloLocacao(True, False, False) = False Then Exit Sub
        frmVendas_custos_relatorios.Show
    Case "AU":
        If FunVefificaModuloLocacao(True, False, False) = False Then Exit Sub
        frmRelatorios_Custos.Show
    Case "DP":
        If FunVefificaModuloLocacao(True, False, False) = False Then Exit Sub
        frmRelatorios_Custos_Prev_Real.Show
    
    'Engenharia
    Case "AV":
        If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub
        If Formulario_familia = "Compras/FamÌlias" Or Formulario_familia = "Vendas/FamÌlias" Or Formulario_familia = "Engenharia/FamÌlias" Or Formulario_familia = "Qualidade/FamÌlias" Then
            If FunVerifFormAberto(frmproj_familia) = True Then
                If USMsgBox("O mÛdulo " & Formulario_familia & " est· aberto, deseja fech·-lo para prosseguir?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub Else Unload frmproj_familia
            End If
        End If
        Compras_Familia = False
        Vendas_Familia = False
        Qualidade_Familia = False
        frmproj_familia.Show
    Case "AX":
        If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub
        If Formulario_produto = "Compras/Produtos e serviÁos" Or Formulario_produto = "Vendas/Produtos e serviÁos" Or Formulario_produto = "Engenharia/Produtos e serviÁos" Then
            If FunVerifFormAberto(frmproj_produto) = True Then
                If USMsgBox("O mÛdulo " & Formulario_produto & " est· aberto, deseja fech·-lo para prosseguir?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub Else Unload frmproj_produto
            End If
        End If
        Engenharia_Produtos = True
        Compras_Produtos = False
        Vendas_Produtos = False
        frmproj_produto.Show
    Case "AZ":
        If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub
        frmproj_conjunto.Show
        
 
    Case "BA1": ' Estrutura completa
        If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub
        If FunVerifAtualizacaoObrigatoria(True, False) = False Then frmproj_produto_estrutura.Show
    Case "BA2": ' Estrutura Resumida
        If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub
        If FunVerifAtualizacaoObrigatoria(True, False) = False Then frmproj_produto_estrutura_Resumida.Show
        
    Case "BB":
        If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub
        frmControle_projetos.Show
    Case "DA":
        If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub
        frmNorma.Show
    
    'Processos
    Case "BC":
        If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub
        If FunVerifAtualizacaoObrigatoria(True, False) = False Then frmProcessos.Show
    
    'PCP
    Case "BE":
        If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub
        frmGermaqfer.Show
    Case "BF":
        If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub
        frmCodigoDesc.Show
    Case "BG":
        If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub
        frmCargaMaq.Show
    Case "BH":
        If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub
        If FunVerifAtualizacaoObrigatoria(True, False) = False Then frmprod.Show
    Case "BI":
        If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub
        frmMonitorTrab.Show
    Case "BJ":
        If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub
        Vendas = False
        FrmSituacao_pedido_producao.Show
    Case "BK":
        If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub
        frmprogramacao.Show
    Case "BL":
        If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub
        PCP_Ordem = True
        frmcqnc.Show
    Case "DU":
        If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub
        If Formulario_necessidade = "Compras/Necessidade" Or Formulario_necessidade = "PCP/Necessidade" Or Formulario_necessidade = "Estoque/Necessidade" Then
            If FunVerifFormAberto(Frm_necessidade) = True Then
                If USMsgBox("O mÛdulo " & Formulario_necessidade & " est· aberto, deseja fech·-lo para prosseguir?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub Else Unload Frm_necessidade
            End If
        End If
        Compras_Necessidade = False
        PCP_Necessidade = True
        Frm_necessidade.Show
    Case "DL":
        If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub
        frmProd_programacao.Show
    Case "EO":
        If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub
        frmPlano_producao.Show
    Case "ER":
        If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub
        frmprod_alterarRM.Show
    Case "BM":
        If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub
        frmProd_Relatorios_Produtividade.Show
    Case "BN":
        If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub
        frmPCP_Relatorios_NC.Show
    Case "BO":
        If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub
        frmProd_Relatorios_Monitor_Eventos.Show
    Case "BP":
        If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub
        frmRelatorios_indice_atraso.Show
    Case "CX":
        If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub
        frmPCP_relatorios_resultados.Show
    
    'Qualidade
    Case "BQ"
        If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub
        If Formulario_familia = "Compras/FamÌlias" Or Formulario_familia = "Vendas/FamÌlias" Or Formulario_familia = "Engenharia/FamÌlias" Or Formulario_familia = "Qualidade/FamÌlias" Then
            If FunVerifFormAberto(frmproj_familia) = True Then
                If USMsgBox("O mÛdulo " & Formulario_familia & " est· aberto, deseja fech·-lo para prosseguir?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub Else Unload frmproj_familia
            End If
        End If
        Compras_Familia = False
        Vendas_Familia = False
        Qualidade_Familia = True
        frmproj_familia.Show
    Case "BR":
        If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub
        frmInstrumentos.Show
    Case "EI":
        If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub
        Qualidade_Almox = True
        frmCFI.Show
    Case "BS":
        If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub
        frmPlanoinspecao.Show
    Case "BT":
        If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub
        frmPlanomedicao.Show
    Case "BU":
        If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub
        frmCompras_recebimento.Show
    'Case "BW":
        'If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub
        'frmUltraSom.Show
    'Case "BV":
        'If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub
        'frmLiquido.Show
    'Case "BX":
        'If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub
        'frmCertificado_qualidade.Show
    Case "BZ":
        If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub
        frmCertificado.Show
    Case "DG":
        If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub
        frmCQ_sistema.Show
    Case "FA":
        If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub
        frmcqnc_descricaoNC.Show
    Case "CA":
        If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub
        PCP_Ordem = False
        frmcqnc.Show
    Case "CB":
        If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub
        frmCQ_SA.Show
    Case "CC":
        If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub
        RNC_Nao_Conformidade = False
        frmCQ_SD.Show
    Case "CD":
        If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub
        RNC_Inspecao_Recebimento = False
        RNC_Controle_Medicao = False
        RNC_Nao_Conformidade = False
        RNC_Solicitacao_Desvio = False
        frmQualidade_RNC.Show
    'Case "CE":
        'If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub
        'frmQualidadePPAP.Show
    'Case "DB":
        'If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub
        'frmQualidadePPAP_FMEA.Show
    'Case "DC":
        'If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub
        'frmQualidadePPAP_PlanoControle.Show
    Case "DD":
        'If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub
        frmQualidade_Revisao_Relatorios.Show
    Case "CF":
        If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub
        Qualidade_NC = True
        frmQualidade_Relatorios_NC.Show
    Case "ED":
        If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub
        frmQualidade_Relatorios_historico.Show (1)
    
    'Estoque
    Case "CG":
        If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub
        Qualidade_Almox = False
        frmCFI.Show
    Case "CH":
        If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub
        frmEstoque_Localarmaz.Show
    Case "CI":
        If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub
        frmRequisicao_materiais.Show
    Case "CJ":
        If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub
        Programacao = False
        frmEstoque_Recebimento.Show
    Case "CK":
        If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub
        frmEstoque_Recebimento_consignacao.Show
    Case "ES":
        If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub
        frmEstoque_ImportarNFe.Show (1)
    Case "EJ":
        If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub
        frmestoque_fisico.Show
    Case "CM":
        If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub
        frmestoque_item.Show
    Case "EK":
        If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub
        If Formulario_necessidade = "Compras/Necessidade" Or Formulario_necessidade = "PCP/Necessidade" Or Formulario_necessidade = "Estoque/Necessidade" Then
            If FunVerifFormAberto(Frm_necessidade) = True Then
                If USMsgBox("O mÛdulo " & Formulario_necessidade & " est· aberto, deseja fech·-lo para prosseguir?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub Else Unload Frm_necessidade
            End If
        End If
        Compras_Necessidade = False
        PCP_Necessidade = False
        Frm_necessidade.Show
    Case "EJ":
        If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub
        frmestoque_fisico.Show
        
    Case "EG":
        If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub
        If Formulario_nota = "Faturamento/Nota fiscal/Terceiros" Or Formulario_nota = "Faturamento/Nota fiscal/PrÛpria" Or Formulario_nota = "Estoque/Nota fiscal" Then
            If FunVerifFormAberto(frmFaturamento_Prod_Serv) = True Then
                If USMsgBox("O mÛdulo " & Formulario_nota & " est· aberto, deseja fech·-lo para prosseguir?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub Else Unload frmFaturamento_Prod_Serv
            End If
        End If
        Formulario = "Estoque/Ordem de faturamento"
        Faturamento_NF_Saida = True
        frmEstoque_Ordem_Faturamento.Show
    Case "CN":
        If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub
        If Formulario_nota = "Faturamento/Nota fiscal/Terceiros" Or Formulario_nota = "Faturamento/Nota fiscal/PrÛpria" Or Formulario_nota = "Estoque/Ordem de faturamento" Then
            If FunVerifFormAberto(frmFaturamento_Prod_Serv) = True Then
                If USMsgBox("O mÛdulo " & Formulario_nota & " est· aberto, deseja fech·-lo para prosseguir?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub Else Unload frmFaturamento_Prod_Serv
            End If
        End If
        Formulario = "Estoque/Nota fiscal"
        Faturamento_NF_Saida = False
        frmFaturamento_Prod_Serv.Show
        
    Case "MT":
'        If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub
'        If Formulario_nota = "Faturamento/Nota fiscal/Terceiros" Or Formulario_nota = "Faturamento/Nota fiscal/PrÛpria" Or Formulario_nota = "Estoque/Ordem de faturamento" Then
'            If FunVerifFormAberto(frmFaturamento_Prod_Serv) = True Then
'                If USMsgBox("O mÛdulo " & Formulario_nota & " est· aberto, deseja fech·-lo para prosseguir?", vbyesno, "CAPRIND v5.0") = vbNo Then Exit Sub Else Unload frmFaturamento_Prod_Serv
'            End If
'        End If
'        Formulario = "Estoque/Nota fiscal"
'        Faturamento_NF_Saida = False
        frmEstoque_consignado.Show
     Case "MI":
'        If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub
'        If Formulario_nota = "Faturamento/Nota fiscal/Terceiros" Or Formulario_nota = "Faturamento/Nota fiscal/PrÛpria" Or Formulario_nota = "Estoque/Ordem de faturamento" Then
'            If FunVerifFormAberto(frmFaturamento_Prod_Serv) = True Then
'                If USMsgBox("O mÛdulo " & Formulario_nota & " est· aberto, deseja fech·-lo para prosseguir?", vbyesno, "CAPRIND v5.0") = vbNo Then Exit Sub Else Unload frmFaturamento_Prod_Serv
'            End If
'        End If
'        Formulario = "Estoque/Nota fiscal"
'        Faturamento_NF_Saida = False
        frmestoque_item_relat.Show
       
    
    'ManutenÁ„o
    Case "CO":
        If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub
        frmManutencao.Show
    'Case "CP":
    Case "CZ":
        If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub
        frmManutencao_relatorios.Show
    
    'Outros
    Case "CQ":
        If FunVefificaModuloLocacao(True, True, True) = False Then Exit Sub
        frmCompras_Requisicao.Show
    Case "EQ":
        If FunVefificaModuloLocacao(True, True, True) = False Then Exit Sub
        frmOutros_Solicitacao_PCP.Show
    Case "J":
        If FunVefificaModuloLocacao(True, True, True) = False Then Exit Sub
        frmCompras_Requisicao_Lista.Show
    Case "CS":
        If FunVefificaModuloLocacao(True, True, True) = False Then Exit Sub
        frmValidacao.Show
    Case "CT":
        If FunVefificaModuloLocacao(True, True, True) = False Then Exit Sub
        frmVendas_analise.Show
    Case "EU":
        If FunVefificaModuloLocacao(True, True, True) = False Then Exit Sub
        Shell "Calc.exe", vbNormalFocus
    Case "EV":
        If FunVefificaModuloLocacao(True, True, True) = False Then Exit Sub
        If TemInternet = True Then
            Downloads_NF = True
            frmMDI_downloads.Show 1
        Else
            USMsgBox ("N„o È permitido abrir este mÛdulo, pois n„o foi encontrado conex„o com a internet."), vbExclamation, "CAPRIND v5.0"
        End If
    Case "EX":
        If FunVefificaModuloLocacao(True, True, True) = False Then Exit Sub
        If TemInternet = True Then
            Downloads_NF = False
            frmMDI_downloads.Show 1
        Else
            USMsgBox ("N„o È permitido abrir este mÛdulo, pois n„o foi encontrado conex„o com a internet."), vbExclamation, "CAPRIND v5.0"
        End If
    
    'Suporte
    Case "CR":
        If FunVefificaModuloLocacao(True, True, True) = False Then Exit Sub
        If TemInternet = True And ErroDriverMYSQL = False Then
            If FunVerificaManutencaoAtiva = False Then Exit Sub
            frmSuporte.Show
        Else
            If TemInternet = False Then MsgTexto = "n„o foi encontrado conex„o com a internet" Else MsgTexto = "no momento estamos sem conex„o com nosso servidor, favor tentar mais tarde"
            USMsgBox ("N„o È permitido abrir este mÛdulo, pois " & MsgTexto & "."), vbExclamation, "CAPRIND v5.0"
        End If
    Case "EP":
        'If FunVefificaModuloLocacao(True, True, True) = False Then Exit Sub
'        If TemInternet = True And ErroDriverMYSQL = False Then
'            If FunVerificaManutencaoAtiva = False Then Exit Sub
'            FrmAtendimento.Show
'        Else
'            If TemInternet = False Then MsgTexto = "n„o foi encontrado conex„o com a internet" Else MsgTexto = "no momento estamos sem conex„o com nosso servidor, favor tentar mais tarde"
'            usMsgbox ("N„o È permitido abrir este mÛdulo, pois " & MsgTexto & "."), vbExclamation, "CAPRIND v5.0"
'        End If
    Case "DR":
        If FunVefificaModuloLocacao(True, True, True) = False Then Exit Sub
        If TemInternet = True And ErroDriverMYSQL = False Then
            If FunVerificaManutencaoAtiva = False Then Exit Sub
'            Chat = True
'            Video_ajuda = False
'            With Frm_web
'                .WindowState = frmMDI.WindowState
'                .Web.Visible = False
'                .Web.Navigate "http://www.caprind.com.br"
'                .Show
'            End With
            Set IE = New InternetExplorer
            IE.Navigate "http://www.caprind.com.br/Suporte/chat.php"
            IE.Visible = True
            'frmMDI.Timer_chat.Enabled = True
        Else
            If TemInternet = False Then MsgTexto = "n„o foi encontrado conex„o com a internet" Else MsgTexto = "no momento estamos sem conex„o com nosso servidor, favor tentar mais tarde"
            USMsgBox ("N„o È permitido abrir este mÛdulo, pois " & MsgTexto & "."), vbExclamation, "CAPRIND v5.0"
        End If
    Case "EZ":
        If FunVefificaModuloLocacao(True, True, True) = False Then Exit Sub
        If GerArqPastas.FileExists(Left(Localrel, 3) & "Programas\Acesso remoto\TeamViewer_Setup_pt 9.0.exe") = True Then
            DS.FileExecute (Left(Localrel, 3) & "Programas\Acesso remoto\TeamViewer_Setup_pt 9.0.exe")
            Exit Sub
        End If
        If TemInternet = True Then
            If FunVerificaManutencaoAtiva = False Then Exit Sub
            Atualizacao_GNFe = False
            Atualizacao_GMRE = False
            Atualizacao_versao = False
            Atualizacao_TeamViewer = True
            Frm_atualizacao_sistema.Show 1
        Else
            USMsgBox ("N„o È permitido baixar o Team Viewer, pois n„o foi encontrado conex„o com a internet."), vbExclamation, "CAPRIND v5.0"
        End If
    Case "DV":
        If FunVefificaModuloLocacao(True, True, True) = False Then Exit Sub
        If TemInternet = True And ErroDriverMYSQL = False Then
            If FunVerificaManutencaoAtiva = False Then Exit Sub
            If GerArqPastas.FileExists(Left(Localrel, 3) & "Caprind.zip") = True Then GerArqPastas.DeleteFile (Left(Localrel, 3) & "Caprind.zip")
            Atualizacao_GNFe = False
            Atualizacao_GMRE = False
            Atualizacao_versao = False
            Frm_atualizacao_sistema.Show 1
        Else
            If TemInternet = False Then MsgTexto = "n„o foi encontrado conex„o com a internet" Else MsgTexto = "no momento estamos sem conex„o com nosso servidor, favor tentar mais tarde"
            USMsgBox ("N„o È permitido baixar a atualizaÁ„o, pois " & MsgTexto & "."), vbExclamation, "CAPRIND v5.0"
        End If
    Case "DX":
        If FunVefificaModuloLocacao(True, True, True) = False Then Exit Sub
        If TemInternet = True And ErroDriverMYSQL = False Then
            If FunVerificaManutencaoAtiva = False Then Exit Sub
            If GerArqPastas.FileExists(Left(Localrel, 3) & "Programas\GNFe\GNFe.zip") = True Then GerArqPastas.DeleteFile (Left(Localrel, 3) & "Programas\GNFe\GNFe.zip")
            Atualizacao_GNFe = True
            Atualizacao_GMRE = False
            Atualizacao_versao = False
            Frm_atualizacao_sistema.Show 1
        Else
            If TemInternet = False Then MsgTexto = "n„o foi encontrado conex„o com a internet" Else MsgTexto = "no momento estamos sem conex„o com nosso servidor, favor tentar mais tarde"
            USMsgBox ("N„o È permitido baixar a atualizaÁ„o, pois " & MsgTexto & "."), vbExclamation, "CAPRIND v5.0"
        End If
    Case "DY":
        If FunVefificaModuloLocacao(True, True, True) = False Then Exit Sub
        If TemInternet = True And ErroDriverMYSQL = False Then
            If FunVerificaManutencaoAtiva = False Then Exit Sub
            If GerArqPastas.FileExists(Left(Localrel, 3) & "Programas\GMRE\GMRE.zip") = True Then GerArqPastas.DeleteFile (Left(Localrel, 3) & "Programas\GMRE\GMRE.zip")
            Atualizacao_GNFe = False
            Atualizacao_GMRE = True
            Atualizacao_versao = False
            Frm_atualizacao_sistema.Show 1
        Else
            If TemInternet = False Then MsgTexto = "n„o foi encontrado conex„o com a internet" Else MsgTexto = "no momento estamos sem conex„o com nosso servidor, favor tentar mais tarde"
            USMsgBox ("N„o È permitido baixar a atualizaÁ„o, pois " & MsgTexto & "."), vbExclamation, "CAPRIND v5.0"
        End If
    
    'Sair
    Case "CW":
        If USMsgBox("Deseja realmente fazer logoff de " & pubUsuario & "?", vbYesNo, "Caprind") = vbYes Then
            ProcLogonOut
            frmLogoff.Show 1
        End If
    Case "CV":
        If USMsgBox("Deseja realmente abandonar o Caprind?", vbYesNo, "Caprind") = vbYes Then
            Call ProcLogonOut
            If Time$ > "19:00:00" Then
                USMsgBox "Boa noite " & pubUsuario & ", Obrigado por utilizar o Sistema Caprind... !!!", _
                        vbInformation, "CAPRIND v5.0"
                ElseIf Time$ > "12:00:00" And Time$ < "18:59:59" Then
                    USMsgBox "Boa Tarde " & pubUsuario & ", Obrigado por utilizar o Sistema Caprind... !!!", _
                            vbInformation, "CAPRIND v5.0"
                ElseIf Time$ > "00:00:00" And Time$ < "11:59:59" Then
                    USMsgBox "Bom dia " & pubUsuario & ", Obrigado por utilizar o Sistema Caprind... !!!", _
                        vbInformation, "CAPRIND v5.0"
            End If
            FunFechaBD
            End
        End If
End Select

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcRecarregaMenu()
On Error GoTo tratar_erro

ProcCarregaMenu frmMDI.TreeView1, frmMDI.ImageList1

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcExcluirArquivosRemessa(ID_nota As Long)
On Error GoTo tratar_erro

'Excluir arquivo remessa
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from tbl_Detalhes_Recebimento WHERE id_nota = " & ID_nota & " and txt_Portador_Banco <> 'Null' and Seq_remessa is not null", Conexao, adOpenKeyset, adLockReadOnly
If TBAbrir.EOF = False Then
    Do While TBAbrir.EOF = False
        If TBAbrir!txt_Portador_Banco <> "" And TBAbrir!Seq_remessa <> "" Then ProcExcluirArquivosRemessa1
        TBAbrir.MoveNext
    Loop
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Public Sub ProcExcluirArquivosRemessa1()
On Error GoTo tratar_erro

Dia = Format(TBAbrir!Data_emissao, "dd")
Mes = Format(TBAbrir!Data_emissao, "mm")
Ano = Format(TBAbrir!Data_emissao, "yyyy")
Set TBFI = CreateObject("adodb.recordset")
TBFI.Open "Select * from tbl_Instituicoes where txt_Descricao = '" & TBAbrir!txt_Portador_Banco & "'", Conexao, adOpenKeyset, adLockReadOnly
If TBFI.EOF = False Then
    If IsNull(TBFI!int_NBanco) = False And TBFI!int_NBanco <> "" Then
        Seq1 = TBAbrir!ID
        Select Case TBFI!int_NBanco
            Case "001": 'Banco do brasil
                Arquivo = "CBR" & Dia & Mes & "." & Seq1
                Diretorio = Localrel & "\Boletos\Arquivos remessa\Banco do brasil"
            Case "033": 'Santander
                Arquivo = "DB" & Dia & Mes & Right(Ano, 2) & "." & Seq1
                Diretorio = Localrel & "\Boletos\Arquivos remessa\Santander"
            Case "104":
                Arquivo = "CB" & Dia & Mes & "." & Seq1 'Caixa
                Diretorio = Localrel & "\Boletos\Arquivos remessa\Caixa"
            Case "237": 'Bradesco
                Arquivo = "CB" & Dia & Mes & "." & Seq1
                Diretorio = Localrel & "\Boletos\Arquivos remessa\Bradesco"
            Case "341": 'Ita˙
                If TBAbrir!Seq_remessa < 10 Then Seq1 = "0" & TBAbrir!Seq_remessa Else Seq1 = TBAbrir!Seq_remessa
                Arquivo = Dia & Mes & Right(Ano, 2) & Seq1
                Diretorio = Localrel & "\Boletos\Arquivos remessa\Ita˙"
            Case "399": 'HSBC
                Arquivo = "D" & Dia & Mes & Ano & "." & Seq1
                Diretorio = Localrel & "\Boletos\Arquivos remessa\HSBC"
            Case "409": 'Unibanco
                Arquivo = "CBR" & Dia & Mes & "." & Seq1
                Diretorio = Localrel & "\Boletos\Arquivos remessa\Unibanco"
        End Select
        TBAbrir!Seq_remessa = Null
        TBAbrir.Update
        Set GerArqPastas = CreateObject("Scripting.FileSystemObject")
        If GerArqPastas.FileExists(Diretorio & "\" & Arquivo & ".txt") = True Then GerArqPastas.DeleteFile (Diretorio & "\" & Arquivo & ".txt")
    End If
End If

tratar_erro:
    If Err.Number = 53 Then Exit Sub
    'usMsgbox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Public Sub procAbrirNotaPDF(TipoNF As String, cnpjXML As String, NotaXML As String, SerieXML As String, DiretorioXMLDanfe As String, CCE As Boolean)
On Error GoTo tratar_erro

Dim diretorioNota As String

cnpjXML = ReturnNumbersOnly(cnpjXML)
NotaXML = FunTamanhoTextoZeroEsq(NotaXML, 9)
SerieXML = FunTamanhoTextoZeroEsq(SerieXML, 5)

diretorioNota = DiretorioXMLDanfe & "\PDF\" & TipoNF & "\" & cnpjXML & NotaXML & SerieXML & IIf(CCE = True, "11011001", "00000000") & ".pdf"

If Dir(diretorioNota) <> "" Then
    ProcAbrirArquivo (diretorioNota)
Else
    USMsgBox "N„o foi possÌvel localizar a Danfe, favor consultar status da nota.", vbExclamation, "CAPRIND v5.0"
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Function FunBuscaEndereco(CEP As String) As Boolean
On Error GoTo tratar_erro

MousePointer = 11
Set myXML = New DOMDocument50
myXML.resolveExternals = True
myXML.validateOnParse = True
myXML.async = False

CEP = ReturnNumbersOnly(CEP)

StrSql = "http://viacep.com.br/ws/" & CEP & "/xml/ "

myXML.Load (StrSql)
myXML.Save ("C:/test.xml")
For Each x In myXML.documentElement.childNodes
    Select Case x.nodeName
        Case Is = "logradouro": If x.childNodes(0).Text <> "" Then Endereco = x.childNodes(0).Text
        Case Is = "uf": If x.childNodes(0).Text <> "" Then UF = x.childNodes(0).Text
        Case Is = "localidade": If x.childNodes(0).Text <> "" Then Cidade = x.childNodes(0).Text
        Case Is = "bairro": If x.childNodes(0).Text <> "" Then Bairro = x.childNodes(0).Text
        
        If Endereco <> "" Or UF <> "" Or Cidade <> "" Or Bairro <> "" Then
                FunBuscaEndereco = True
        End If
        
        If Endereco = "" And UF = "" And Cidade = "" And Bairro <> "" Then
                FunBuscaEndereco = False
        End If
        
    End Select
Next

Exit Function
tratar_erro:
    MousePointer = 0
    If Err.Number = 91 Then
        USMsgBox ("N„o foi possÌvel carregar todos os dados referentes a este CEP."), vbInformation, "CAPRIND v5.0"
        Exit Function
    End If
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Function FunBuscaValorAproximadoTributos() As Boolean
On Error GoTo tratar_erro

Dim resposta As String
Dim obj As MSXML2.ServerXMLHTTP50
Set obj = New MSXML2.ServerXMLHTTP50
obj.Open "GET", "https://apidoni.ibpt.org.br/api/v1/produtos?token=MDvFlwOEEsTwHkSfpvx4rqvwP2U_6S6m7WpD6kclmkKMqOG87j9FtXQyf9BWru0B&cnpj=16740838000151&codigo=87149500&uf=SP&ex=0&descricao=X&unidadeMedida=X&valor=1&gtin=0"

conteudo = "codigo=87149500&uf=SP&ex=0&descricao=X&unidadeMedida=X&valor=1&gtin=0"
obj.send conteudo

resposta = obj.responseText
'Debug.print resposta

CODIGO = LerDadosJSON(resposta, "Codigo", "", "")
UF = LerDadosJSON(resposta, "UF", "", "")
EX = LerDadosJSON(resposta, "EX", "", "")
Descricao = LerDadosJSON(resposta, "Descricao", "", "")
nacional = LerDadosJSON(resposta, "Nacional", "", "")

USMsgBox "Codigo = " & CODIGO & vbCrLf & "UF: " & UF & vbCrLf & "Descricao: " & Descricao & vbCrLf & "Nacional: " & nacional & vbCrLf

Exit Function
tratar_erro:
    MousePointer = 0
    If Err.Number = 91 Then
        USMsgBox ("N„o foi possÌvel carregar todos os dados referentes a este CEP."), vbInformation, "CAPRIND v5.0"
        Exit Function
    End If
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Sub ProcAtualizaVlrEntradaEstoque(AtualizacaoRecebEst As Boolean)
On Error GoTo tratar_erro

If AtualizacaoRecebEst = False Then
    INNERJOINTEXTO = "INNER JOIN tbl_Detalhes_Nota_pedidos NFPP ON NFPP.ID_carteira = EM.IDLista_recebimento and NFPP.Codinterno = EM.Desenho) INNER JOIN tbl_Detalhes_Nota NFP ON NFP.Int_codigo = NFPP.ID_prod_NF where NFP.ID_nota = " & TBVendas!ID & " and NFP.Remessa = 'False'"
Else
    If Left(Formulario, 11) = "Faturamento" Then
        IDFase = TBVendas!ID_empresa
        IDclienteTexto = TBVendas!Id_Int_Cliente
        TextoFiltro = "NF.int_NotaFiscal = '" & TBVendas!int_NotaFiscal & "' and NF.Serie = '" & TBVendas!Serie & "'"
    Else
        IDFase = frmEstoque_Recebimento.txtID_empresa
        IDclienteTexto = frmEstoque_Recebimento.Txt_ID_forn
        TextoFiltro = "NF.int_NotaFiscal = '" & frmEstoque_Recebimento.txtnotafiscal & "' and NF.Serie = '" & frmEstoque_Recebimento.txtSerie & "'"
    End If
    INNERJOINTEXTO = "INNER JOIN tbl_Dados_Nota_Fiscal NF ON NF.int_NotaFiscal = ECR.Nota_fiscal and NF.Serie = ECR.Serie and NF.ID_empresa = ECR.ID_empresa and NF.Id_Int_Cliente = " & IDclienteTexto & ") INNER JOIN tbl_Detalhes_Nota NFP ON NFP.ID_nota = NF.ID and NFP.int_Cod_Produto = ECR.Desenho where " & TextoFiltro & " and NF.Aplicacao = 'P' and NF.int_TipoNota = 2 and NFP.Remessa = 'False'"
End If
'Atualiza valor do estoque
Set TBFIltro = CreateObject("adodb.recordset")
StrSql = "Select ECR.ID_empresa, NFP.*, EM.Entrada, EM.VlrUnit, EM.VlrTotal, EM.IDestoque, CP.Estado, ISNULL(CPL.Qtde_estoque, 0) AS Qtde_estoque from ((((estoque_movimentacao EM INNER JOIN Estoque_controle_recebimento ECR ON EM.IDEstoque_recebimento = ECR.Id) INNER JOIN Compras_pedido_lista CPL ON ECR.IDPedido = CPL.IDPedido and ECR.IdLista = CPL.IdLista and ECR.Desenho = CPL.Desenho) INNER JOIN Compras_pedido CP ON CP.IDpedido = CPL.IDPedido) " & INNERJOINTEXTO & " and CPL.remessa = 'False'"

'Debug.print StrSql

TBFIltro.Open "Select ECR.ID_empresa, NFP.*, EM.Entrada, EM.VlrUnit, EM.VlrTotal, EM.IDestoque, CP.Estado, ISNULL(CPL.Qtde_estoque, 0) AS Qtde_estoque from ((((estoque_movimentacao EM INNER JOIN Estoque_controle_recebimento ECR ON EM.IDEstoque_recebimento = ECR.Id) INNER JOIN Compras_pedido_lista CPL ON ECR.IDPedido = CPL.IDPedido and ECR.IdLista = CPL.IdLista and ECR.Desenho = CPL.Desenho) INNER JOIN Compras_pedido CP ON CP.IDpedido = CPL.IDPedido) " & INNERJOINTEXTO & " and CPL.remessa = 'False'", Conexao, adOpenKeyset, adLockOptimistic
If TBFIltro.EOF = False Then
    Do While TBFIltro.EOF = False
        qt = 1
        If TBFIltro!txt_Unid <> TBFIltro!Unidade_com And TBFIltro!Qtde_estoque > 0 Then qt = TBFIltro!int_Qtd / TBFIltro!Entrada
                
        'Verifica valor do ICMS
        ValorICMS = 0
        Valor1 = 0
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from tbl_Detalhes_Nota_CST_ICMS where ID_item = " & TBFIltro!Int_codigo, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            If IsNull(TBAbrir!Valor_ICMS) = False And TBAbrir!Valor_ICMS <> 0 Then
                ValorICMS = TBAbrir!Valor_ICMS
            ElseIf IsNull(TBAbrir!Valor_ICMS_ST) = False And TBAbrir!Valor_ICMS_ST <> 0 Then
                    ValorICMS = TBAbrir!Valor_ICMS_ST
                ElseIf IsNull(TBAbrir!Valor_ICMS_SN) = False And TBAbrir!Valor_ICMS_SN <> 0 Then
                        ValorICMS = TBAbrir!Valor_ICMS_SN
            End If
        End If
        If ValorICMS <> 0 Then Valor1 = Format(qt * (ValorICMS / TBFIltro!int_Qtd), "0.0000000000") 'Valor unit·rio de ICMS
        
        QuantsolicitadoN2 = Format(qt * (IIf(IsNull(TBFIltro!Valor_desconto), 0, TBFIltro!Valor_desconto) / TBFIltro!int_Qtd), "0.0000000000") 'Valor unit·rio de desconto
        Valor2 = Format(qt * (IIf(IsNull(TBFIltro!Valor_frete), 0, TBFIltro!Valor_frete) / TBFIltro!int_Qtd), "0.0000000000")
        Valor3 = Format(qt * (IIf(IsNull(TBFIltro!Valor_seguro), 0, TBFIltro!Valor_seguro) / TBFIltro!int_Qtd), "0.0000000000")
        ValorPagar = Format(qt * (IIf(IsNull(TBFIltro!Valor_acessorias), 0, TBFIltro!Valor_acessorias) / TBFIltro!int_Qtd), "0.0000000000")
        
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select Simples, Real from Empresa where Codigo = " & TBFIltro!ID_empresa, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            If TBAbrir!Simples = True Then
                If TBFIltro!Estado = "EX" Then
                    'Quando È nota de importaÁ„o os valores de PIS e Cofins j· est„o inclusos nas despesas acessorias
                    Valor_PIS_Prod = 0
                    Valor_Cofins_Prod = 0
                Else
                    Valor_PIS_Prod = Format(qt * (TBFIltro!Total_PIS_prod / TBFIltro!int_Qtd), "0.0000000000")
                    Valor_Cofins_Prod = Format(qt * (TBFIltro!Total_Cofins_prod / TBFIltro!int_Qtd), "0.0000000000")
                End If
                Valor_CSLL_Prod = Format(qt * (TBFIltro!Total_CSLL_prod / TBFIltro!int_Qtd), "0.0000000000")
                Valor_IRPJ_Prod = Format(qt * (TBFIltro!Total_IRPJ_prod / TBFIltro!int_Qtd), "0.0000000000")
                'VALOR UNIT¡RIO DO ESTOQUE = (Valor unit·rio - Valor desc.) + (Valor ICMS + Frete + Seguro + Valor despesas + Valor PIS + Valor Cofins + Valor CSLL + Valor IRPJ)
                Qtd = Format(qt * (IIf(IsNull(TBFIltro!dbl_ValorUnitario), "0", TBFIltro!dbl_ValorUnitario) - QuantsolicitadoN2) + (Valor1 + Valor2 + Valor3 + ValorPagar + Valor_PIS_Prod + Valor_Cofins_Prod + Valor_CSLL_Prod + Valor_IRPJ_Prod), "0.0000000000")
            ElseIf TBAbrir!Real = True Then
                    Valor_PIS_Prod = Format(qt * (TBFIltro!Total_PIS_prod / TBFIltro!int_Qtd), "0.0000000000")
                    Valor_Cofins_Prod = Format(qt * (TBFIltro!Total_Cofins_prod / TBFIltro!int_Qtd), "0.0000000000")
                    'VALOR UNIT¡RIO DO ESTOQUE = (Valor unit·rio + Valor do frete + Valor Seguro + Valor despesas) - (Valor desc. + Valor ICMS + Valor PIS + Valor Cofins)
                    Qtd = (qt * IIf(IsNull(TBFIltro!dbl_ValorUnitario), "0", TBFIltro!dbl_ValorUnitario)) + Valor2 + Valor3 + ValorPagar
                    Qtd = Format(Qtd - (QuantsolicitadoN2 + Valor1 + Valor_PIS_Prod + Valor_Cofins_Prod), "0.0000000000")
                Else
                    'VALOR UNIT¡RIO DO ESTOQUE = (Valor unit·rio + Valor do frete + Valor Seguro + Valor despesas) - (Valor desc. + Valor ICMS)
                    Qtd = (qt * IIf(IsNull(TBFIltro!dbl_ValorUnitario), "0", TBFIltro!dbl_ValorUnitario)) + Valor2 + Valor3 + ValorPagar
                    Qtd = Format(Qtd - (QuantsolicitadoN2 + Valor1), "0.0000000000")
                End If
        End If
        Qtde = TBFIltro!Entrada
        TBFIltro!VlrUnit = Qtd
        TBFIltro!vlrTotal = Format(Qtd * Qtde, "0.00")
        
        Set TBFI = CreateObject("adodb.recordset")
        TBFI.Open "Select estoque_real, valor_unitario, Valor_total from estoque_controle where IdEstoque = " & TBFIltro!IDEstoque & " and estoque_real IS NOT NULL", Conexao, adOpenKeyset, adLockOptimistic
        If TBFI.EOF = False Then
            Qtde = TBFI!estoque_real
            TBFI!valor_unitario = Qtd
            TBFI!Valor_total = Format(Qtd * Qtde, "###,##0.00")
            TBFI.Update
        End If
        TBFI.Close
        
        TBFIltro.Update
        TBFIltro.MoveNext
    Loop
End If
TBFIltro.Close

'Atualiza cÛdigo de refÍrencia no estoque
If AtualizacaoRecebEst = False Then
    If FunVerifCodRefFornSPED(TBVendas!ID_empresa) = True Then
        Conexao.Execute "Update EC set EC.ref = NFP.N_Referencia from (((tbl_Detalhes_Nota_pedidos NFPP INNER JOIN Estoque_controle_recebimento ECR ON ECR.IdLista = NFPP.ID_carteira AND ECR.Desenho = NFPP.Codinterno) INNER JOIN Estoque_movimentacao EM ON EM.IDEstoque_recebimento = ECR.ID) INNER JOIN Estoque_Controle EC ON EC.IDestoque = EM.Idestoque) INNER JOIN tbl_Detalhes_Nota NFP ON NFP.Int_codigo = NFPP.ID_prod_NF where NFPP.ID_nota = " & TBVendas!ID
        
        'NF consignada
        Conexao.Execute "Update EC set EC.ref = NFP.N_Referencia from Estoque_Controle EC INNER JOIN tbl_Detalhes_Nota NFP ON NFP.int_Cod_Produto = EC.Desenho and NFP.ID_nota = " & TBVendas!ID & " where EC.Consignacao = 'True' and EC.Lote = '" & TBVendas!int_NotaFiscal & "' and EC.Cliente = '" & TBVendas!txt_Razao_Nome & "' and EC.id_cliente = " & TBVendas!Id_Int_Cliente
    End If
Else
    If FunVerifCodRefFornSPED(IDFase) = True Then
        Conexao.Execute "Update EC set EC.ref = NFP.N_Referencia from (((estoque_movimentacao EM INNER JOIN Estoque_controle_recebimento ECR ON EM.IDEstoque_recebimento = ECR.Id) INNER JOIN tbl_Dados_Nota_Fiscal NF ON NF.int_NotaFiscal = ECR.Nota_fiscal and NF.Serie = ECR.Serie and NF.ID_empresa = ECR.ID_empresa and NF.Id_Int_Cliente = " & IDclienteTexto & ") INNER JOIN tbl_Detalhes_Nota NFP ON NFP.ID_nota = NF.ID and NFP.int_Cod_Produto = ECR.Desenho) INNER JOIN Estoque_Controle EC ON EC.IDestoque = EM.Idestoque where " & TextoFiltro & " and NF.Aplicacao = 'P' and NF.int_TipoNota = 2"
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Function FunCarregaCodRef(Codinterno As String) As String
On Error GoTo tratar_erro

FunCarregaCodRef = ""
Set TBCarregarCombo = CreateObject("adodb.recordset")
TBCarregarCombo.Open "Select IA.N_referencia from Projproduto P INNER JOIN item_aplicacoes IA ON IA.Codproduto = P.Codproduto where P.Desenho = '" & Codinterno & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBCarregarCombo.EOF = False Then
    FunCarregaCodRef = TBCarregarCombo!N_referencia
End If
TBCarregarCombo.Close

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Sub ProcVerifQtdeLicencasModulo()
On Error GoTo tratar_erro

Qtlicencas_caprind = 0
Qtlicencas_gerprod = 0
Modulo_caprind = ""
If TemInternet = True And ErroDriverMYSQL = False Then
    Set TBComponente = CreateObject("adodb.recordset")
    TBComponente.Open "Select CNPJ from Empresa", Conexao, adOpenKeyset, adLockOptimistic
    If TBComponente.EOF = False Then
        FunAbreBDSite
        If ConexaoMySql.State = 1 Then
            Set TBMySQL = New ADODB.Recordset
            TBMySQL.Open "Select Licencas, Licencas_gerprod, Modulo From Clientes Where CNPJ = '" & TBComponente!CNPJ & "'", ConexaoMySql, adOpenKeyset, adLockOptimistic, adCmdText
            With TBMySQL
                If .EOF = False Then
                    'Verifica n˙mero de licenÁas
                    If IsNull(.Fields!Licencas) = False And .Fields!Licencas <> "" Then Qtlicencas_caprind = .Fields!Licencas
                    If IsNull(.Fields!Licencas_gerprod) = False And .Fields!Licencas_gerprod <> "" Then Qtlicencas_gerprod = .Fields!Licencas_gerprod
                    Modulo_caprind = IIf(IsNull(.Fields!Modulo), "", .Fields!Modulo)
                End If
            End With
        End If
    End If
    TBComponente.Close
Else
    Set TBComponente = CreateObject("adodb.recordset")
    TBComponente.Open "Select Licencas_caprind, Licencas_gerprod, Modulo from Empresa", Conexao, adOpenKeyset, adLockOptimistic
    If TBComponente.EOF = False Then
        If IsNull(TBComponente!Licencas_caprind) = False And TBComponente!Licencas_caprind <> "" Then Qtlicencas_caprind = TBComponente!Licencas_caprind
        If IsNull(TBComponente!Licencas_gerprod) = False And TBComponente!Licencas_gerprod <> "" Then Qtlicencas_gerprod = TBComponente!Licencas_gerprod
        Modulo_caprind = IIf(IsNull(TBComponente!Modulo), "", TBComponente!Modulo)
    End If
    TBComponente.Close
End If

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Function FunVerifMaxPBListaPaginacao(QtdeRegBD As Long, QtdeRegPag As Long) As Long
On Error GoTo tratar_erro

FunVerifMaxPBListaPaginacao = IIf(QtdeRegBD > QtdeRegPag, QtdeRegPag, QtdeRegBD)

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Sub ProcValidarEngenhariaEstruturaProd(Codproduto As Long)
On Error GoTo tratar_erro

Set TBTempo = CreateObject("adodb.recordset")
TBTempo.Open "Select dtValidacaoConj, RespValidacaoConj from Projproduto where Codproduto = " & Codproduto, Conexao, adOpenKeyset, adLockOptimistic
If TBTempo.EOF = False Then
    Set TBCFOP = CreateObject("adodb.recordset")
    TBCFOP.Open "Select DtValidacao, RespValidacao from Projconjunto_desc_versao where Codproduto = " & Codproduto & " and DtValidacao IS NOT NULL order by DtValidacao", Conexao, adOpenKeyset, adLockOptimistic
    If TBCFOP.EOF = False Then
        TBTempo!DtValidacaoConj = TBCFOP!DtValidacao
        TBTempo!RespValidacaoConj = TBCFOP!RespValidacao
    Else
        TBTempo!DtValidacaoConj = Null
        TBTempo!RespValidacaoConj = Null
    End If
    TBTempo.Update
End If
TBTempo.Close

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcGravarValidEstruturaVersao(Codproduto As Long, DtValidacao As String, RespValidacao As String, versao As String)
On Error GoTo tratar_erro

Set TBGravar = CreateObject("adodb.recordset")

StrSql = "Select * from Projconjunto_desc_versao where Codproduto = " & Codproduto & " and Versao = '" & versao & "'"
'Debug.print StrSql

TBGravar.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then TBGravar.AddNew
TBGravar!Codproduto = Codproduto
TBGravar!versao = versao
Set TBAbrir = CreateObject("adodb.recordset")
StrSql = "Select Codigo from Projconjunto where codproduto = " & Codproduto & " and Versao = '" & versao & "'"
'Debug.print StrSql

TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    TBGravar!DtValidacao = IIf(DtValidacao = "", Null, DtValidacao)
    TBGravar!RespValidacao = IIf(RespValidacao = "", Null, RespValidacao)
End If
TBAbrir.Close
TBGravar.Update
TBGravar.Close

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCopiarDescVersaoEstrutura(CodprodutoAnt As Long, VersaoAnt As String, CodprodutoNovo As Long, VersaoNova As String)
On Error GoTo tratar_erro

Set TBTempo = CreateObject("adodb.recordset")
TBTempo.Open "Select * from Projconjunto_desc_versao where codproduto = " & CodprodutoAnt & " and Versao = '" & VersaoAnt & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBTempo.EOF = False Then
    Set TBGravar = CreateObject("adodb.recordset")
    TBGravar.Open "Select * from Projconjunto_desc_versao where codProduto = " & CodprodutoNovo & " and Versao = '" & VersaoNova & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBGravar.EOF = True Then TBGravar.AddNew
    TBGravar!Codproduto = CodprodutoNovo
    TBGravar!versao = VersaoNova
    TBGravar!Descricao = TBTempo!Descricao
    TBGravar.Update
End If
TBTempo.Close

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Function FunVerifVersaoCriadaEstrutura(versao As String, Codproduto As Long) As Boolean
On Error GoTo tratar_erro

FunVerifVersaoCriadaEstrutura = True
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select versao from Projconjunto WHERE Codproduto = " & Codproduto & " and Versao = '" & versao & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = True Then FunVerifVersaoCriadaEstrutura = False
TBAbrir.Close

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Sub ProcAtualizaCTMaterialOrdem(ID_empresa As Integer, Ordem As Long)
On Error GoTo tratar_erro

Set TBCiclo = CreateObject("adodb.recordset")
TBCiclo.Open "Select Ordem, CTMaterial from producao where ID_empresa = " & ID_empresa & " and Ordem = " & Ordem, Conexao, adOpenKeyset, adLockOptimistic
If TBCiclo.EOF = False Then
    valor = 0
    Set TBComponente = CreateObject("adodb.recordset")
    TBComponente.Open "Select Sum(EM.VlrTotal) as Valor from Estoque_movimentacao EM INNER JOIN Estoque_controle EC ON EC.IDestoque = EM.IDestoque where EM.Documento = '" & TBCiclo!Ordem & "' and EC.Consignacao = 'False' and (EM.Operacao = 'SAIDA_ORDEM' or EM.Operacao = 'SAIDA_ORDEM_PARCIAL')", Conexao, adOpenKeyset, adLockOptimistic
    If TBComponente.EOF = False Then
        valor = IIf(IsNull(TBComponente!valor), 0, TBComponente!valor)
    End If
    TBComponente.Close
    TBCiclo!CTMaterial = Format(valor, "###,##0.00")
    TBCiclo.Update
End If
TBCiclo.Close

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Function FunVerificaVlrUltCompra(Codinterno As String) As Double
On Error GoTo tratar_erro

FunVerificaVlrUltCompra = 0
Set TBComponente = CreateObject("adodb.recordset")
TBComponente.Open "Select ROUND(CPL.preco_unitario * ISNULL(CC.Valor_moeda, 1), 10) AS preco_unitario from Compras_pedido_lista CPL LEFT JOIN Compras_comercial CC ON CC.IDpedido = CPL.IDpedido where CPL.Desenho = '" & Codinterno & "' and CPL.IDpedido <> 0 order by CPL.Idlista desc", Conexao, adOpenKeyset, adLockOptimistic
If TBComponente.EOF = False Then FunVerificaVlrUltCompra = TBComponente!preco_unitario
TBComponente.Close

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

'ESTRUTURA NOVA=============================================================
Sub ProcNivel2Estrutura(NomeFormulario As Form, versao As String, MostrarValores As Boolean, Carteira_pedidos As Boolean, CarregarNivelAbaixo As Boolean, Simulacao_vendas As Boolean)
On Error GoTo tratar_erro

If versao = "" Then TextoFiltro = " " Else TextoFiltro = " and Versao = '" & versao & "'"
Set TBNivel2 = CreateObject("adodb.recordset")
TBNivel2.Open "Select * from projconjunto where codproduto = " & Codproduto & TextoFiltro & " order by Posicao, Codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBNivel2.EOF = False Then
    Do While TBNivel2.EOF = False
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select P.codproduto, P.largura, P.comprimento, P.Subtipoitem, P.Producao, P.PCusto, PCDV.DtValidacao, PCDV.RespValidacao from projproduto P LEFT JOIN Projconjunto_desc_versao PCDV ON PCDV.Codproduto = P.Codproduto " & IIf(versao <> "", "and PCDV.Versao = '" & versao & "'", "") & " where P.desenho = '" & TBNivel2!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Codproduto = TBAbrir!Codproduto
            If TBNivel2!Unidade = "KG" And TBNivel2!PesoTotal > 0 Or TBNivel2!Unidade = "MT" And TBNivel2!Dimensoes > 0 Or TBNivel2!Unidade = "MM" And TBNivel2!Dimensoes > 0 Then
                Select Case TBNivel2!Unidade
                    Case "KG":  Peso = TBNivel2!PesoTotal
                    Case "MT":  Peso = (TBNivel2!Dimensoes * TBNivel2!quantidade) / 1000
                    Case "MM":  Peso = TBNivel2!Dimensoes * TBNivel2!quantidade
                End Select
            Else
                Peso = TBNivel2!quantidade
            End If
            
            QuantsolicitadoN1 = Peso * quantidade
            FunCarregaValorEstrutura TBNivel2!CODIGO, TBNivel2!Desenho, MostrarValores, Simulacao_vendas, QuantsolicitadoN1, True
            
            If Carteira_pedidos = True Then
                Processos = "N√O"
                Set TBProcessos = CreateObject("adodb.recordset")
                TBProcessos.Open "Select PR.IDProcesso FROM Processos PR INNER JOIN Projproduto P ON PR.Codproduto = P.Codproduto WHERE P.Desenho = '" & TBNivel2!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBProcessos.EOF = False Then
                    Processos = "SIM"
                End If
                OrdemCarteira = ""
                Set TBOrdem = CreateObject("adodb.recordset")
                TBOrdem.Open "Select PR.Ordem FROM Producao PR INNER JOIN Producao_pedidos PP ON PR.Ordem = PP.Ordem where PP.IDCarteira = " & Ordem & " and PR.Desenho = '" & TBNivel2!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBOrdem.EOF = False Then
                    OrdemCarteira = TBOrdem!Ordem
                End If
            End If
                        
            PosicaoTexto = FunTamanhoTextoZeroEsq(TBNivel2!Posicao, 3)
            
            CodRef = ""
            Set TBItem = CreateObject("adodb.recordset")
            TBItem.Open "Select n_referencia from item_aplicacoes where codproduto = " & TBAbrir!Codproduto, Conexao, adOpenKeyset, adLockOptimistic
            If TBItem.EOF = False Then
                CodRef = TBItem!N_referencia
            End If
            TBItem.Close
            
            If TBAbrir!SubTipoItem = 1 Or TBAbrir!SubTipoItem = 2 Then
                DataValidacao = IIf(IsNull(TBAbrir!DtValidacao), "", TBAbrir!DtValidacao)
                RespValidacao = IIf(IsNull(TBAbrir!RespValidacao), "", TBAbrir!RespValidacao)
            Else
                DataValidacao = ""
                RespValidacao = ""
            End If
            
            PartNumber = ""
            If IsNull(TBNivel2!ID_partnumber_fabricante) = False Then
                Set TBProduto = CreateObject("adodb.recordset")
                TBProduto.Open "Select Part_number from Projproduto_fabricante where ID = " & TBNivel2!ID_partnumber_fabricante, Conexao, adOpenKeyset, adLockOptimistic
                If TBProduto.EOF = False Then PartNumber = TBProduto!Part_number
                TBProduto.Close
            End If
            
            Contador1 = Contador1 + 1
            arrNodes(Contador1).Level = 1
            
            
            If Carteira_pedidos = True Or Simulacao_vendas = True Then
                QtTexto = Format(QuantsolicitadoN1, "###,##0.0000")
                If Carteira_pedidos = True Then
                    arrNodes(Contador1).Text = TBNivel2!Desenho & vbTab & PosicaoTexto & vbTab & Codproduto & vbTab & Codproduto & vbTab & CodRef & vbTab & TBNivel2!Descricao & vbTab & TBNivel2!Unidade & vbTab & TBNivel2!Versao_desenho & vbTab & "" & vbTab & "" & vbTab & QtTexto & vbTab & "" & vbTab & "" & vbTab & Processos & vbTab & OrdemCarteira & vbTab & DataValidacao & vbTab & RespValidacao & vbTab & PartNumber & vbTab & TBNivel2!Obs
                Else
                    arrNodes(Contador1).Text = TBNivel2!Desenho & vbTab & PosicaoTexto & vbTab & Codproduto & vbTab & CodRef & vbTab & TBNivel2!Descricao & vbTab & TBNivel2!Unidade & vbTab & TBNivel2!Versao_desenho & vbTab & QtTexto & vbTab & Format(QuantComprado, "###,##0.0000") & vbTab & Format(IIf(QuantComprado - QuantsolicitadoN1 < 0, (QuantComprado - QuantsolicitadoN1) * -1, 0), "###,##0.0000") & vbTab & Format(Valor_Cofins_Serv, "###,##0.00000000") & vbTab & Format(valor, "###,##0.00000000") & vbTab & Format(Valor_Cofins_Prod, "###,##0.00000000") & vbTab & DataValidacao & vbTab & RespValidacao & vbTab & PartNumber & vbTab & TBNivel2!Obs & IIf(QuantComprado - QuantsolicitadoN1 < 0, 1, 0)
                End If
            Else
                Kg_un = Format(TBNivel2!PesoMetro, "###,##0.0000000000")
                Un_Kg = TBNivel2!Un_Kg
                Dim_mm = Format(TBNivel2!Dimensoes, "###,##0.0000000000")
                Kg_pc = Format(TBNivel2!Peso, "###,##0.0000000000")
                QtTexto = Format(TBNivel2!quantidade, "0.0000")
                Peso_total = Format(TBNivel2!PesoTotal, "###,##0.0000000000")
                If Formulario = "Engenharia/Processos" Then
                    arrNodes(Contador1).Text = TBNivel2!Desenho & vbTab & PosicaoTexto & vbTab & Codproduto & vbTab & CodRef & vbTab & TBNivel2!Descricao & vbTab & TBNivel2!Unidade & vbTab & TBNivel2!Versao_desenho & vbTab & Kg_un & vbTab & Un_Kg & vbTab & Dim_mm & vbTab & Kg_pc & vbTab & QtTexto & vbTab & Peso_total & vbTab & DataValidacao & vbTab & RespValidacao & vbTab & PartNumber & vbTab & TBNivel2!Obs
                ElseIf Formulario = "Engenharia/Estrutura/Resumida" Then
                  '  If PosicaoTexto = "031" Then MsgBox "Aqui"
                    arrNodes(Contador1).Text = TBNivel2!Desenho & vbTab & PosicaoTexto & vbTab & IIf(TBNivel2!Tipo = 1, "Titular", "Alternativo") & vbTab & Codproduto & vbTab & Replace(CodRef, vbTab, "") & vbTab & PartNumber & vbTab & TBNivel2!Descricao & vbTab & TBNivel2!Obs & vbTab & TBNivel2!Unidade & vbTab & TBNivel2!Versao_desenho & vbTab & Kg_un & vbTab & Un_Kg & vbTab & Dim_mm & vbTab & Kg_pc & vbTab & IIf(IsNull(TBAbrir!Largura), 0, Format(TBAbrir!Largura, "###,##0.00")) & vbTab & IIf(IsNull(TBAbrir!Comprimento), "", Format(TBAbrir!Comprimento, "###,##0.00")) & vbTab & QtTexto & vbTab & Peso_total & vbTab & Format(valor, "###,##0.00000000") & vbTab & Format(DataValidacao, "dd/mm/yy") & vbTab & RespValidacao & vbTab & TBNivel2!CODIGO
                ElseIf Formulario = "Engenharia/Estrutura/Detalhada" Then
                    arrNodes(Contador1).Text = TBNivel2!Desenho & vbTab & PosicaoTexto & vbTab & IIf(TBNivel2!Tipo = 1, "Titular", "Alternativo") & vbTab & Codproduto & vbTab & Replace(CodRef, vbTab, "") & vbTab & TBNivel2!Descricao & vbTab & TBNivel2!Unidade & vbTab & TBNivel2!Versao_desenho & vbTab & Kg_un & vbTab & Un_Kg & vbTab & Dim_mm & vbTab & Kg_pc & vbTab & IIf(IsNull(TBAbrir!Largura), 0, Format(TBAbrir!Largura, "###,##0.00")) & vbTab & IIf(IsNull(TBAbrir!Comprimento), "", Format(TBAbrir!Comprimento, "###,##0.00")) & vbTab & QtTexto & vbTab & Peso_total & vbTab & Format(valor, "###,##0.00000000") & vbTab & DataValidacao & vbTab & RespValidacao & vbTab & TBNivel2!CODIGO & vbTab & PartNumber & vbTab & TBNivel2!Obs
                ElseIf Formulario = "Vendas/Pedido interno" Then
                    arrNodes(Contador1).Text = TBNivel2!Desenho & vbTab & TBNivel2!Descricao & vbTab & TBNivel2!Unidade & vbTab & QtTexto & vbTab & TBNivel2!Obs
                Else
                    arrNodes(Contador1).Text = TBNivel2!Desenho & vbTab & PosicaoTexto & vbTab & Codproduto & vbTab & CodRef & vbTab & TBNivel2!Descricao & vbTab & TBNivel2!Unidade & vbTab & TBNivel2!Versao_desenho & vbTab & Kg_un & vbTab & Un_Kg & vbTab & Dim_mm & vbTab & Kg_pc & vbTab & QtTexto & vbTab & Peso_total & vbTab & Format(valor, "###,##0.00000000") & vbTab & DataValidacao & vbTab & RespValidacao & vbTab & TBNivel2!CODIGO & vbTab & PartNumber & vbTab & TBNivel2!Obs
                End If
            End If
            If CarregarNivelAbaixo = True Then ProcNivel3Estrutura NomeFormulario, TBNivel2!Versao_desenho, MostrarValores, Carteira_pedidos, CarregarNivelAbaixo, Simulacao_vendas
        End If
        TBNivel2.MoveNext
    Loop
End If

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcNivel3Estrutura(NomeFormulario As Form, versao As String, MostrarValores As Boolean, Carteira_pedidos As Boolean, CarregarNivelAbaixo As Boolean, Simulacao_vendas As Boolean)
On Error GoTo tratar_erro

If versao = "" Then TextoFiltro = " " Else TextoFiltro = " and Versao = '" & versao & "'"
Set TBNivel3 = CreateObject("adodb.recordset")
TBNivel3.Open "Select * from projconjunto where codproduto = " & Codproduto & TextoFiltro & " order by Posicao, Codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBNivel3.EOF = False Then
    Do While TBNivel3.EOF = False
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select P.codproduto, P.largura, P.comprimento, P.Subtipoitem, P.Producao, P.PCusto, PCDV.DtValidacao, PCDV.RespValidacao from projproduto P LEFT JOIN Projconjunto_desc_versao PCDV ON PCDV.Codproduto = P.Codproduto " & IIf(versao <> "", "and PCDV.Versao = '" & versao & "'", "") & " where P.desenho = '" & TBNivel3!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Codproduto = TBAbrir!Codproduto
            If TBNivel3!Unidade = "KG" And TBNivel3!PesoTotal > 0 Or TBNivel3!Unidade = "MT" And TBNivel3!Dimensoes > 0 Or TBNivel3!Unidade = "MM" And TBNivel3!Dimensoes > 0 Then
                Select Case TBNivel3!Unidade
                    Case "KG":  Peso = TBNivel3!PesoTotal
                    Case "MT":  Peso = (TBNivel3!Dimensoes * TBNivel3!quantidade) / 1000
                    Case "MM":  Peso = TBNivel3!Dimensoes * TBNivel3!quantidade
                End Select
            Else
                Peso = TBNivel3!quantidade
            End If
            
            QuantsolicitadoN2 = Peso * QuantsolicitadoN1
            FunCarregaValorEstrutura TBNivel3!CODIGO, TBNivel3!Desenho, MostrarValores, Simulacao_vendas, QuantsolicitadoN2, False
            
            If Carteira_pedidos = True Then
                Processos = "N√O"
                Set TBProcessos = CreateObject("adodb.recordset")
                TBProcessos.Open "Select PR.IDProcesso FROM Processos PR INNER JOIN Projproduto P ON PR.Codproduto = P.Codproduto WHERE P.Desenho = '" & TBNivel3!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBProcessos.EOF = False Then
                    Processos = "SIM"
                End If
                OrdemCarteira = ""
                Set TBOrdem = CreateObject("adodb.recordset")
                TBOrdem.Open "Select PR.Ordem FROM Producao PR INNER JOIN Producao_pedidos PP ON PR.Ordem = PP.Ordem where PP.IDCarteira = " & Ordem & " and PR.Desenho = '" & TBNivel3!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBOrdem.EOF = False Then
                    OrdemCarteira = TBOrdem!Ordem
                End If
            End If
            
            PosicaoTexto = FunTamanhoTextoZeroEsq(TBNivel3!Posicao, 3)
            
            CodRef = ""
            Set TBItem = CreateObject("adodb.recordset")
            TBItem.Open "Select n_referencia from item_aplicacoes where codproduto = " & TBAbrir!Codproduto, Conexao, adOpenKeyset, adLockOptimistic
            If TBItem.EOF = False Then
                CodRef = TBItem!N_referencia
            End If
            TBItem.Close
            
            If TBAbrir!SubTipoItem = 1 Or TBAbrir!SubTipoItem = 2 Then
                DataValidacao = IIf(IsNull(TBAbrir!DtValidacao), "", TBAbrir!DtValidacao)
                RespValidacao = IIf(IsNull(TBAbrir!RespValidacao), "", TBAbrir!RespValidacao)
            Else
                DataValidacao = ""
                RespValidacao = ""
            End If
            
            PartNumber = ""
            If IsNull(TBNivel3!ID_partnumber_fabricante) = False Then
                Set TBProduto = CreateObject("adodb.recordset")
                TBProduto.Open "Select Part_number from Projproduto_fabricante where ID = " & TBNivel3!ID_partnumber_fabricante, Conexao, adOpenKeyset, adLockOptimistic
                If TBProduto.EOF = False Then PartNumber = TBProduto!Part_number
                TBProduto.Close
            End If
            
            Contador1 = Contador1 + 1
            arrNodes(Contador1).Level = 2
            If Carteira_pedidos = True Or Simulacao_vendas = True Then
                QtTexto = Format(QuantsolicitadoN2, "###,##0.0000")
                If Carteira_pedidos = True Then
                    arrNodes(Contador1).Text = TBNivel3!Desenho & vbTab & PosicaoTexto & vbTab & Codproduto & vbTab & CodRef & vbTab & TBNivel3!Descricao & vbTab & TBNivel3!Unidade & vbTab & TBNivel3!Versao_desenho & vbTab & "" & vbTab & "" & vbTab & QtTexto & vbTab & "" & vbTab & "" & vbTab & Processos & vbTab & OrdemCarteira & vbTab & DataValidacao & vbTab & RespValidacao & vbTab & PartNumber & vbTab & TBNivel3!Obs
                Else
                    arrNodes(Contador1).Text = TBNivel3!Desenho & vbTab & PosicaoTexto & vbTab & Codproduto & vbTab & CodRef & vbTab & TBNivel3!Descricao & vbTab & TBNivel3!Unidade & vbTab & TBNivel3!Versao_desenho & vbTab & QtTexto & vbTab & Format(QuantComprado, "###,##0.0000") & vbTab & Format(IIf(QuantComprado - QuantsolicitadoN2 < 0, (QuantComprado - QuantsolicitadoN2) * -1, 0), "###,##0.0000") & vbTab & Format(Valor_Cofins_Serv, "###,##0.00000000") & vbTab & Format(valor, "###,##0.00000000") & vbTab & Format(Valor_Cofins_Prod, "###,##0.00000000") & vbTab & DataValidacao & vbTab & RespValidacao & vbTab & PartNumber & vbTab & TBNivel3!Obs & IIf(QuantComprado - QuantsolicitadoN2 < 0, 1, 0)
                End If
            Else
                Kg_un = Format(TBNivel3!PesoMetro, "###,##0.0000000000")
                Un_Kg = TBNivel3!Un_Kg
                Dim_mm = Format(TBNivel3!Dimensoes, "###,##0.0000000000")
                Kg_pc = Format(TBNivel3!Peso, "###,##0.0000000000")
                QtTexto = Format(TBNivel3!quantidade, "0.0000")
                Peso_total = Format(TBNivel3!PesoTotal, "###,##0.0000000000")
                If Formulario = "Engenharia/Processos" Then
                    arrNodes(Contador1).Text = TBNivel3!Desenho & vbTab & PosicaoTexto & vbTab & Codproduto & vbTab & CodRef & vbTab & TBNivel3!Descricao & vbTab & TBNivel3!Unidade & vbTab & TBNivel3!Versao_desenho & vbTab & Kg_un & vbTab & Un_Kg & vbTab & Dim_mm & vbTab & Kg_pc & vbTab & QtTexto & vbTab & Peso_total & vbTab & DataValidacao & vbTab & RespValidacao & vbTab & PartNumber & vbTab & TBNivel3!Obs
                ElseIf Formulario = "Engenharia/Estrutura/Resumida" Then
                    arrNodes(Contador1).Text = TBNivel3!Desenho & vbTab & PosicaoTexto & vbTab & IIf(TBNivel3!Tipo = 1, "Titular", "Alternativo") & vbTab & Codproduto & vbTab & Replace(CodRef, vbTab, "") & vbTab & PartNumber & vbTab & TBNivel3!Descricao & vbTab & TBNivel3!Obs & vbTab & TBNivel3!Unidade & vbTab & TBNivel3!Versao_desenho & vbTab & Kg_un & vbTab & Un_Kg & vbTab & Dim_mm & vbTab & Kg_pc & vbTab & IIf(IsNull(TBAbrir!Largura), 0, Format(TBAbrir!Largura, "###,##0.00")) & vbTab & IIf(IsNull(TBAbrir!Comprimento), "", Format(TBAbrir!Comprimento, "###,##0.00")) & vbTab & QtTexto & vbTab & Peso_total & vbTab & Format(valor, "###,##0.00000000") & vbTab & Format(DataValidacao, "dd/mm/yy") & vbTab & RespValidacao & vbTab & TBNivel3!CODIGO
                ElseIf Formulario = "Engenharia/Estrutura/Detalhada" Then
                    arrNodes(Contador1).Text = TBNivel3!Desenho & vbTab & PosicaoTexto & vbTab & IIf(TBNivel3!Tipo = 1, "Titular", "Alternativo") & vbTab & Codproduto & vbTab & CodRef & vbTab & TBNivel3!Descricao & vbTab & TBNivel3!Unidade & vbTab & TBNivel3!Versao_desenho & vbTab & Kg_un & vbTab & Un_Kg & vbTab & Dim_mm & vbTab & Kg_pc & vbTab & IIf(IsNull(TBAbrir!Largura), 0, Format(TBAbrir!Largura, "###,##0.00")) & vbTab & IIf(IsNull(TBAbrir!Comprimento), "", Format(TBAbrir!Comprimento, "###,##0.00")) & vbTab & QtTexto & vbTab & Peso_total & vbTab & Format(valor, "###,##0.00000000") & vbTab & DataValidacao & vbTab & RespValidacao & vbTab & TBNivel3!CODIGO & vbTab & PartNumber & vbTab & TBNivel3!Obs
                ElseIf Formulario = "Vendas/Pedido interno" Then
                    arrNodes(Contador1).Text = TBNivel3!Desenho & vbTab & TBNivel3!Descricao & vbTab & TBNivel3!Unidade & vbTab & QtTexto & vbTab & TBNivel3!Obs
                Else
                    arrNodes(Contador1).Text = TBNivel3!Desenho & vbTab & PosicaoTexto & vbTab & Codproduto & vbTab & CodRef & vbTab & TBNivel3!Descricao & vbTab & TBNivel3!Unidade & vbTab & TBNivel3!Versao_desenho & vbTab & Kg_un & vbTab & Un_Kg & vbTab & Dim_mm & vbTab & Kg_pc & vbTab & QtTexto & vbTab & Peso_total & vbTab & Format(valor, "###,##0.00000000") & vbTab & DataValidacao & vbTab & RespValidacao & vbTab & TBNivel3!CODIGO & vbTab & PartNumber & vbTab & TBNivel3!Obs
                End If
            End If
            If CarregarNivelAbaixo = True Then ProcNivel4Estrutura NomeFormulario, TBNivel3!Versao_desenho, MostrarValores, Carteira_pedidos, CarregarNivelAbaixo, Simulacao_vendas
        End If
        TBNivel3.MoveNext
    Loop
End If

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcNivel4Estrutura(NomeFormulario As Form, versao As String, MostrarValores As Boolean, Carteira_pedidos As Boolean, CarregarNivelAbaixo As Boolean, Simulacao_vendas As Boolean)
On Error GoTo tratar_erro

If versao = "" Then TextoFiltro = " " Else TextoFiltro = " and Versao = '" & versao & "'"
Set TBNivel4 = CreateObject("adodb.recordset")
TBNivel4.Open "Select * from projconjunto where codproduto = " & Codproduto & TextoFiltro & " order by Posicao, Codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBNivel4.EOF = False Then
    Do While TBNivel4.EOF = False
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select P.codproduto, P.largura, P.comprimento, P.Subtipoitem, P.Producao, P.PCusto, PCDV.DtValidacao, PCDV.RespValidacao from projproduto P LEFT JOIN Projconjunto_desc_versao PCDV ON PCDV.Codproduto = P.Codproduto " & IIf(versao <> "", "and PCDV.Versao = '" & versao & "'", "") & " where P.desenho = '" & TBNivel4!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Codproduto = TBAbrir!Codproduto
            If TBNivel4!Unidade = "KG" And TBNivel4!PesoTotal > 0 Or TBNivel4!Unidade = "MT" And TBNivel4!Dimensoes > 0 Or TBNivel4!Unidade = "MM" And TBNivel4!Dimensoes > 0 Then
                Select Case TBNivel4!Unidade
                    Case "KG":  Peso = TBNivel4!PesoTotal
                    Case "MT":  Peso = (TBNivel4!Dimensoes * TBNivel4!quantidade) / 1000
                    Case "MM":  Peso = TBNivel4!Dimensoes * TBNivel4!quantidade
                End Select
            Else
                Peso = TBNivel4!quantidade
            End If
            
            QuantsolicitadoN3 = Peso * QuantsolicitadoN2
            FunCarregaValorEstrutura TBNivel4!CODIGO, TBNivel4!Desenho, MostrarValores, Simulacao_vendas, QuantsolicitadoN3, False
            
            If Carteira_pedidos = True Then
                Processos = "N√O"
                Set TBProcessos = CreateObject("adodb.recordset")
                TBProcessos.Open "Select PR.IDProcesso FROM Processos PR INNER JOIN Projproduto P ON PR.Codproduto = P.Codproduto WHERE P.Desenho = '" & TBNivel4!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBProcessos.EOF = False Then
                    Processos = "SIM"
                End If
                OrdemCarteira = ""
                Set TBOrdem = CreateObject("adodb.recordset")
                TBOrdem.Open "Select PR.Ordem FROM Producao PR INNER JOIN Producao_pedidos PP ON PR.Ordem = PP.Ordem where PP.IDCarteira = " & Ordem & " and PR.Desenho = '" & TBNivel4!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBOrdem.EOF = False Then
                    OrdemCarteira = TBOrdem!Ordem
                End If
            End If
            
            PosicaoTexto = FunTamanhoTextoZeroEsq(TBNivel4!Posicao, 3)
            
            CodRef = ""
            Set TBItem = CreateObject("adodb.recordset")
            TBItem.Open "Select n_referencia from item_aplicacoes where codproduto = " & TBAbrir!Codproduto, Conexao, adOpenKeyset, adLockOptimistic
            If TBItem.EOF = False Then
                CodRef = TBItem!N_referencia
            End If
            TBItem.Close
            
            If TBAbrir!SubTipoItem = 1 Or TBAbrir!SubTipoItem = 2 Then
                DataValidacao = IIf(IsNull(TBAbrir!DtValidacao), "", TBAbrir!DtValidacao)
                RespValidacao = IIf(IsNull(TBAbrir!RespValidacao), "", TBAbrir!RespValidacao)
            Else
                DataValidacao = ""
                RespValidacao = ""
            End If
            
            PartNumber = ""
            If IsNull(TBNivel4!ID_partnumber_fabricante) = False Then
                Set TBProduto = CreateObject("adodb.recordset")
                TBProduto.Open "Select Part_number from Projproduto_fabricante where ID = " & TBNivel4!ID_partnumber_fabricante, Conexao, adOpenKeyset, adLockOptimistic
                If TBProduto.EOF = False Then PartNumber = TBProduto!Part_number
                TBProduto.Close
            End If
            
            Contador1 = Contador1 + 1
            arrNodes(Contador1).Level = 3
            If Carteira_pedidos = True Or Simulacao_vendas = True Then
                QtTexto = Format(QuantsolicitadoN3, "###,##0.0000")
                If Carteira_pedidos = True Then
                    arrNodes(Contador1).Text = TBNivel4!Desenho & vbTab & PosicaoTexto & vbTab & Codproduto & vbTab & CodRef & vbTab & TBNivel4!Descricao & vbTab & TBNivel4!Unidade & vbTab & TBNivel4!Versao_desenho & vbTab & "" & vbTab & "" & vbTab & QtTexto & vbTab & "" & vbTab & "" & vbTab & Processos & vbTab & OrdemCarteira & vbTab & DataValidacao & vbTab & RespValidacao & vbTab & PartNumber & vbTab & TBNivel4!Obs
                Else
                    arrNodes(Contador1).Text = TBNivel4!Desenho & vbTab & PosicaoTexto & vbTab & Codproduto & vbTab & CodRef & vbTab & TBNivel4!Descricao & vbTab & TBNivel4!Unidade & vbTab & TBNivel4!Versao_desenho & vbTab & QtTexto & vbTab & Format(QuantComprado, "###,##0.0000") & vbTab & Format(IIf(QuantComprado - QuantsolicitadoN3 < 0, (QuantComprado - QuantsolicitadoN3) * -1, 0), "###,##0.0000") & vbTab & Format(Valor_Cofins_Serv, "###,##0.00000000") & vbTab & Format(valor, "###,##0.00000000") & vbTab & Format(Valor_Cofins_Prod, "###,##0.00000000") & vbTab & DataValidacao & vbTab & RespValidacao & vbTab & PartNumber & vbTab & TBNivel4!Obs & IIf(QuantComprado - QuantsolicitadoN3 < 0, 1, 0)
                End If
            Else
                Kg_un = Format(TBNivel4!PesoMetro, "###,##0.0000000000")
                Un_Kg = TBNivel4!Un_Kg
                Dim_mm = Format(TBNivel4!Dimensoes, "###,##0.0000000000")
                Kg_pc = Format(TBNivel4!Peso, "###,##0.0000000000")
                QtTexto = Format(TBNivel4!quantidade, "0.0000")
                Peso_total = Format(TBNivel4!PesoTotal, "###,##0.0000000000")
                If Formulario = "Engenharia/Processos" Then
                    arrNodes(Contador1).Text = TBNivel4!Desenho & vbTab & PosicaoTexto & vbTab & Codproduto & vbTab & CodRef & vbTab & TBNivel4!Descricao & vbTab & TBNivel4!Unidade & vbTab & TBNivel4!Versao_desenho & vbTab & Kg_un & vbTab & Un_Kg & vbTab & Dim_mm & vbTab & Kg_pc & vbTab & QtTexto & vbTab & Peso_total & vbTab & DataValidacao & vbTab & RespValidacao & vbTab & PartNumber & vbTab & TBNivel4!Obs
                ElseIf Formulario = "Engenharia/Estrutura/Resumida" Then
                    arrNodes(Contador1).Text = TBNivel4!Desenho & vbTab & PosicaoTexto & vbTab & IIf(TBNivel4!Tipo = 1, "Titular", "Alternativo") & vbTab & Codproduto & vbTab & Replace(CodRef, vbTab, "") & vbTab & PartNumber & vbTab & TBNivel4!Descricao & vbTab & TBNivel4!Obs & vbTab & TBNivel2!Unidade & vbTab & TBNivel4!Versao_desenho & vbTab & Kg_un & vbTab & Un_Kg & vbTab & Dim_mm & vbTab & Kg_pc & vbTab & IIf(IsNull(TBAbrir!Largura), 0, Format(TBAbrir!Largura, "###,##0.00")) & vbTab & IIf(IsNull(TBAbrir!Comprimento), "", Format(TBAbrir!Comprimento, "###,##0.00")) & vbTab & QtTexto & vbTab & Peso_total & vbTab & Format(valor, "###,##0.00000000") & vbTab & Format(DataValidacao, "dd/mm/yy") & vbTab & RespValidacao & vbTab & TBNivel4!CODIGO
                ElseIf Formulario = "Engenharia/Estrutura/Detalhada" Then
                    arrNodes(Contador1).Text = TBNivel4!Desenho & vbTab & PosicaoTexto & vbTab & IIf(TBNivel4!Tipo = 1, "Titular", "Alternativo") & vbTab & Codproduto & vbTab & CodRef & vbTab & TBNivel3!Descricao & vbTab & TBNivel3!Unidade & vbTab & TBNivel3!Versao_desenho & vbTab & Kg_un & vbTab & Un_Kg & vbTab & Dim_mm & vbTab & Kg_pc & vbTab & IIf(IsNull(TBAbrir!Largura), 0, Format(TBAbrir!Largura, "###,##0.00")) & vbTab & IIf(IsNull(TBAbrir!Comprimento), "", Format(TBAbrir!Comprimento, "###,##0.00")) & vbTab & QtTexto & vbTab & Peso_total & vbTab & Format(valor, "###,##0.00000000") & vbTab & DataValidacao & vbTab & RespValidacao & vbTab & TBNivel3!CODIGO & vbTab & PartNumber & vbTab & TBNivel3!Obs
                Else
                    arrNodes(Contador1).Text = TBNivel4!Desenho & vbTab & PosicaoTexto & vbTab & Codproduto & vbTab & CodRef & vbTab & TBNivel4!Descricao & vbTab & TBNivel4!Unidade & vbTab & TBNivel4!Versao_desenho & vbTab & Kg_un & vbTab & Un_Kg & vbTab & Dim_mm & vbTab & Kg_pc & vbTab & QtTexto & vbTab & Peso_total & vbTab & Format(valor, "###,##0.00000000") & vbTab & DataValidacao & vbTab & RespValidacao & vbTab & TBNivel4!CODIGO & vbTab & PartNumber & vbTab & TBNivel4!Obs
                End If
            End If
            If CarregarNivelAbaixo = True Then ProcNivel5Estrutura NomeFormulario, TBNivel4!Versao_desenho, MostrarValores, Carteira_pedidos, CarregarNivelAbaixo, Simulacao_vendas
        End If
        TBNivel4.MoveNext
    Loop
End If

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcNivel5Estrutura(NomeFormulario As Form, versao As String, MostrarValores As Boolean, Carteira_pedidos As Boolean, CarregarNivelAbaixo As Boolean, Simulacao_vendas As Boolean)
On Error GoTo tratar_erro

If versao = "" Then TextoFiltro = " " Else TextoFiltro = " and Versao = '" & versao & "'"
Set TBNivel5 = CreateObject("adodb.recordset")
TBNivel5.Open "Select * from projconjunto where codproduto = " & Codproduto & TextoFiltro & " order by Posicao, Codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBNivel5.EOF = False Then
    Do While TBNivel5.EOF = False
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select P.codproduto, P.largura, P.comprimento, P.Subtipoitem, P.Producao, P.PCusto, PCDV.DtValidacao, PCDV.RespValidacao from projproduto P LEFT JOIN Projconjunto_desc_versao PCDV ON PCDV.Codproduto = P.Codproduto " & IIf(versao <> "", "and PCDV.Versao = '" & versao & "'", "") & " where P.desenho = '" & TBNivel5!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Codproduto = TBAbrir!Codproduto
            If TBNivel5!Unidade = "KG" And TBNivel5!PesoTotal > 0 Or TBNivel5!Unidade = "MT" And TBNivel5!Dimensoes > 0 Or TBNivel5!Unidade = "MM" And TBNivel5!Dimensoes > 0 Then
                Select Case TBNivel5!Unidade
                    Case "KG":  Peso = TBNivel5!PesoTotal
                    Case "MT":  Peso = (TBNivel5!Dimensoes * TBNivel5!quantidade) / 1000
                    Case "MM":  Peso = TBNivel5!Dimensoes * TBNivel5!quantidade
                End Select
            Else
                Peso = TBNivel5!quantidade
            End If
            
            QuantsolicitadoN4 = Peso * QuantsolicitadoN3
            FunCarregaValorEstrutura TBNivel5!CODIGO, TBNivel5!Desenho, MostrarValores, Simulacao_vendas, QuantsolicitadoN4, False
            
            If Carteira_pedidos = True Then
                Processos = "N√O"
                Set TBProcessos = CreateObject("adodb.recordset")
                TBProcessos.Open "Select PR.IDProcesso FROM Processos PR INNER JOIN Projproduto P ON PR.Codproduto = P.Codproduto WHERE P.Desenho = '" & TBNivel5!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBProcessos.EOF = False Then
                    Processos = "SIM"
                End If
                OrdemCarteira = ""
                Set TBOrdem = CreateObject("adodb.recordset")
                TBOrdem.Open "Select PR.Ordem FROM Producao PR INNER JOIN Producao_pedidos PP ON PR.Ordem = PP.Ordem where PP.IDCarteira = " & Ordem & " and PR.Desenho = '" & TBNivel5!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBOrdem.EOF = False Then
                    OrdemCarteira = TBOrdem!Ordem
                End If
            End If
            
            PosicaoTexto = FunTamanhoTextoZeroEsq(TBNivel5!Posicao, 3)
            
            CodRef = ""
            Set TBItem = CreateObject("adodb.recordset")
            TBItem.Open "Select n_referencia from item_aplicacoes where codproduto = " & TBAbrir!Codproduto, Conexao, adOpenKeyset, adLockOptimistic
            If TBItem.EOF = False Then
                CodRef = TBItem!N_referencia
            End If
            TBItem.Close
            
            If TBAbrir!SubTipoItem <> 0 Then
                DataValidacao = IIf(IsNull(TBAbrir!DtValidacao), "", TBAbrir!DtValidacao)
                RespValidacao = IIf(IsNull(TBAbrir!RespValidacao), "", TBAbrir!RespValidacao)
            Else
                DataValidacao = ""
                RespValidacao = ""
            End If
            
            PartNumber = ""
            If IsNull(TBNivel5!ID_partnumber_fabricante) = False Then
                Set TBProduto = CreateObject("adodb.recordset")
                TBProduto.Open "Select Part_number from Projproduto_fabricante where ID = " & TBNivel5!ID_partnumber_fabricante, Conexao, adOpenKeyset, adLockOptimistic
                If TBProduto.EOF = False Then PartNumber = TBProduto!Part_number
                TBProduto.Close
            End If
            
            Contador1 = Contador1 + 1
            arrNodes(Contador1).Level = 4
            If Carteira_pedidos = True Or Simulacao_vendas = True Then
                QtTexto = Format(QuantsolicitadoN4, "###,##0.0000")
                If Carteira_pedidos = True Then
                    arrNodes(Contador1).Text = TBNivel5!Desenho & vbTab & PosicaoTexto & vbTab & Codproduto & vbTab & CodRef & vbTab & TBNivel5!Descricao & vbTab & TBNivel5!Unidade & vbTab & TBNivel5!Versao_desenho & vbTab & "" & vbTab & "" & vbTab & QtTexto & vbTab & "" & vbTab & "" & vbTab & Processos & vbTab & OrdemCarteira & vbTab & DataValidacao & vbTab & RespValidacao & vbTab & PartNumber & vbTab & TBNivel5!Obs
                Else
                    arrNodes(Contador1).Text = TBNivel5!Desenho & vbTab & PosicaoTexto & vbTab & Codproduto & vbTab & CodRef & vbTab & TBNivel5!Descricao & vbTab & TBNivel5!Unidade & vbTab & TBNivel5!Versao_desenho & vbTab & QtTexto & vbTab & Format(QuantComprado, "###,##0.0000") & vbTab & Format(IIf(QuantComprado - QuantsolicitadoN4 < 0, (QuantComprado - QuantsolicitadoN4) * -1, 0), "###,##0.0000") & vbTab & Format(Valor_Cofins_Serv, "###,##0.00000000") & vbTab & Format(valor, "###,##0.00000000") & vbTab & Format(Valor_Cofins_Prod, "###,##0.00000000") & vbTab & DataValidacao & vbTab & RespValidacao & vbTab & PartNumber & vbTab & TBNivel5!Obs & IIf(QuantComprado - QuantsolicitadoN4 < 0, 1, 0)
                End If
            Else
                Kg_un = Format(TBNivel5!PesoMetro, "###,##0.0000000000")
                Un_Kg = TBNivel5!Un_Kg
                Dim_mm = Format(TBNivel5!Dimensoes, "###,##0.0000000000")
                Kg_pc = Format(TBNivel5!Peso, "###,##0.0000000000")
                QtTexto = Format(TBNivel5!quantidade, "0.0000")
                Peso_total = Format(TBNivel5!PesoTotal, "###,##0.0000000000")
                If Formulario = "Engenharia/Processos" Then
                    arrNodes(Contador1).Text = TBNivel5!Desenho & vbTab & PosicaoTexto & vbTab & Codproduto & vbTab & CodRef & vbTab & TBNivel5!Descricao & vbTab & TBNivel5!Unidade & vbTab & TBNivel5!Versao_desenho & vbTab & Kg_un & vbTab & Un_Kg & vbTab & Dim_mm & vbTab & Kg_pc & vbTab & QtTexto & vbTab & Peso_total & vbTab & DataValidacao & vbTab & RespValidacao & vbTab & PartNumber & vbTab & TBNivel5!Obs
                ElseIf Formulario = "Engenharia/Estrutura/Resumida" Then
                    arrNodes(Contador1).Text = TBNivel5!Desenho & vbTab & PosicaoTexto & vbTab & IIf(TBNivel5!Tipo = 1, "Titular", "Alternativo") & vbTab & Codproduto & vbTab & Replace(CodRef, vbTab, "") & vbTab & PartNumber & vbTab & TBNivel5!Descricao & vbTab & TBNivel5!Obs & vbTab & TBNivel2!Unidade & vbTab & TBNivel5!Versao_desenho & vbTab & Kg_un & vbTab & Un_Kg & vbTab & Dim_mm & vbTab & Kg_pc & vbTab & IIf(IsNull(TBAbrir!Largura), 0, Format(TBAbrir!Largura, "###,##0.00")) & vbTab & IIf(IsNull(TBAbrir!Comprimento), "", Format(TBAbrir!Comprimento, "###,##0.00")) & vbTab & QtTexto & vbTab & Peso_total & vbTab & Format(valor, "###,##0.00000000") & vbTab & Format(DataValidacao, "dd/mm/yy") & vbTab & RespValidacao & vbTab & TBNivel5!CODIGO
                ElseIf Formulario = "Engenharia/Estrutura/Detalhada" Then
                    arrNodes(Contador1).Text = TBNivel5!Desenho & vbTab & PosicaoTexto & vbTab & IIf(TBNivel5!Tipo = 1, "Titular", "Alternativo") & vbTab & Codproduto & vbTab & CodRef & vbTab & TBNivel3!Descricao & vbTab & TBNivel3!Unidade & vbTab & TBNivel3!Versao_desenho & vbTab & Kg_un & vbTab & Un_Kg & vbTab & Dim_mm & vbTab & Kg_pc & vbTab & IIf(IsNull(TBAbrir!Largura), 0, Format(TBAbrir!Largura, "###,##0.00")) & vbTab & IIf(IsNull(TBAbrir!Comprimento), "", Format(TBAbrir!Comprimento, "###,##0.00")) & vbTab & QtTexto & vbTab & Peso_total & vbTab & Format(valor, "###,##0.00000000") & vbTab & DataValidacao & vbTab & RespValidacao & vbTab & TBNivel3!CODIGO & vbTab & PartNumber & vbTab & TBNivel3!Obs
                Else
                    arrNodes(Contador1).Text = TBNivel5!Desenho & vbTab & PosicaoTexto & vbTab & Codproduto & vbTab & CodRef & vbTab & TBNivel5!Descricao & vbTab & TBNivel5!Unidade & vbTab & TBNivel5!Versao_desenho & vbTab & Kg_un & vbTab & Un_Kg & vbTab & Dim_mm & vbTab & Kg_pc & vbTab & QtTexto & vbTab & Peso_total & vbTab & Format(valor, "###,##0.00000000") & vbTab & DataValidacao & vbTab & RespValidacao & vbTab & TBNivel5!CODIGO & vbTab & PartNumber & vbTab & TBNivel5!Obs
                End If
            End If
            If CarregarNivelAbaixo = True Then ProcNivel6Estrutura NomeFormulario, TBNivel5!Versao_desenho, MostrarValores, Carteira_pedidos, CarregarNivelAbaixo, Simulacao_vendas
        End If
        TBNivel5.MoveNext
    Loop
End If

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcNivel6Estrutura(NomeFormulario As Form, versao As String, MostrarValores As Boolean, Carteira_pedidos As Boolean, CarregarNivelAbaixo As Boolean, Simulacao_vendas As Boolean)
On Error GoTo tratar_erro

If versao = "" Then TextoFiltro = " " Else TextoFiltro = " and Versao = '" & versao & "'"
Set TBNivel6 = CreateObject("adodb.recordset")
TBNivel6.Open "Select * from projconjunto where codproduto = " & Codproduto & TextoFiltro & " order by Posicao, Codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBNivel6.EOF = False Then
    Do While TBNivel6.EOF = False
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select P.codproduto, P.largura, P.comprimento, P.Subtipoitem, P.Producao, P.PCusto, PCDV.DtValidacao, PCDV.RespValidacao from projproduto P LEFT JOIN Projconjunto_desc_versao PCDV ON PCDV.Codproduto = P.Codproduto " & IIf(versao <> "", "and PCDV.Versao = '" & versao & "'", "") & " where P.desenho = '" & TBNivel6!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Codproduto = TBAbrir!Codproduto
            If TBNivel6!Unidade = "KG" And TBNivel6!PesoTotal > 0 Or TBNivel6!Unidade = "MT" And TBNivel6!Dimensoes > 0 Or TBNivel6!Unidade = "MM" And TBNivel6!Dimensoes > 0 Then
                Select Case TBNivel6!Unidade
                    Case "KG":  Peso = TBNivel6!PesoTotal
                    Case "MT":  Peso = (TBNivel6!Dimensoes * TBNivel6!quantidade) / 1000
                    Case "MM":  Peso = TBNivel6!Dimensoes * TBNivel6!quantidade
                End Select
            Else
                Peso = TBNivel6!quantidade
            End If
            
            QuantsolicitadoN5 = Peso * QuantsolicitadoN4
            FunCarregaValorEstrutura TBNivel6!CODIGO, TBNivel6!Desenho, MostrarValores, Simulacao_vendas, QuantsolicitadoN5, False
            
            If Carteira_pedidos = True Then
                Processos = "N√O"
                Set TBProcessos = CreateObject("adodb.recordset")
                TBProcessos.Open "Select PR.IDProcesso FROM Processos PR INNER JOIN Projproduto P ON PR.Codproduto = P.Codproduto WHERE P.Desenho = '" & TBNivel6!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBProcessos.EOF = False Then
                    Processos = "SIM"
                End If
                OrdemCarteira = ""
                Set TBOrdem = CreateObject("adodb.recordset")
                TBOrdem.Open "Select PR.Ordem FROM Producao PR INNER JOIN Producao_pedidos PP ON PR.Ordem = PP.Ordem where PP.IDCarteira = " & Ordem & " and PR.Desenho = '" & TBNivel6!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBOrdem.EOF = False Then
                    OrdemCarteira = TBOrdem!Ordem
                End If
            End If
            
            PosicaoTexto = FunTamanhoTextoZeroEsq(TBNivel6!Posicao, 3)
            
            CodRef = ""
            Set TBItem = CreateObject("adodb.recordset")
            TBItem.Open "Select n_referencia from item_aplicacoes where codproduto = " & TBAbrir!Codproduto, Conexao, adOpenKeyset, adLockOptimistic
            If TBItem.EOF = False Then
                CodRef = TBItem!N_referencia
            End If
            TBItem.Close
            
            If TBAbrir!SubTipoItem <> 0 Then
                DataValidacao = IIf(IsNull(TBAbrir!DtValidacao), "", TBAbrir!DtValidacao)
                RespValidacao = IIf(IsNull(TBAbrir!RespValidacao), "", TBAbrir!RespValidacao)
            Else
                DataValidacao = ""
                RespValidacao = ""
            End If
            
            PartNumber = ""
            If IsNull(TBNivel6!ID_partnumber_fabricante) = False Then
                Set TBProduto = CreateObject("adodb.recordset")
                TBProduto.Open "Select Part_number from Projproduto_fabricante where ID = " & TBNivel6!ID_partnumber_fabricante, Conexao, adOpenKeyset, adLockOptimistic
                If TBProduto.EOF = False Then PartNumber = TBProduto!Part_number
                TBProduto.Close
            End If
            
            Contador1 = Contador1 + 1
            arrNodes(Contador1).Level = 5
            If Carteira_pedidos = True Or Simulacao_vendas = True Then
                QtTexto = Format(QuantsolicitadoN5, "###,##0.0000")
                If Carteira_pedidos = True Then
                    arrNodes(Contador1).Text = TBNivel6!Desenho & vbTab & PosicaoTexto & vbTab & Codproduto & vbTab & CodRef & vbTab & TBNivel6!Descricao & vbTab & TBNivel6!Unidade & vbTab & TBNivel6!Versao_desenho & vbTab & "" & vbTab & "" & vbTab & QtTexto & vbTab & "" & vbTab & "" & vbTab & Processos & vbTab & OrdemCarteira & vbTab & DataValidacao & vbTab & RespValidacao & vbTab & PartNumber & vbTab & TBNivel6!Obs
                Else
                    arrNodes(Contador1).Text = TBNivel6!Desenho & vbTab & PosicaoTexto & vbTab & Codproduto & vbTab & CodRef & vbTab & TBNivel6!Descricao & vbTab & TBNivel6!Unidade & vbTab & TBNivel6!Versao_desenho & vbTab & QtTexto & vbTab & Format(QuantComprado, "###,##0.0000") & vbTab & Format(IIf(QuantComprado - QuantsolicitadoN5 < 0, (QuantComprado - QuantsolicitadoN5) * -1, 0), "###,##0.0000") & vbTab & Format(Valor_Cofins_Serv, "###,##0.00000000") & vbTab & Format(valor, "###,##0.00000000") & vbTab & Format(Valor_Cofins_Prod, "###,##0.00000000") & vbTab & DataValidacao & vbTab & RespValidacao & vbTab & PartNumber & vbTab & TBNivel6!Obs & IIf(QuantComprado - QuantsolicitadoN5 < 0, 1, 0)
                End If
            Else
                Kg_un = Format(TBNivel6!PesoMetro, "###,##0.0000000000")
                Un_Kg = TBNivel6!Un_Kg
                Dim_mm = Format(TBNivel6!Dimensoes, "###,##0.0000000000")
                Kg_pc = Format(TBNivel6!Peso, "###,##0.0000000000")
                QtTexto = Format(TBNivel6!quantidade, "0.000")
                Peso_total = Format(TBNivel6!PesoTotal, "###,##0.0000000000")
                If Formulario = "Engenharia/Processos" Then
                    arrNodes(Contador1).Text = TBNivel6!Desenho & vbTab & PosicaoTexto & vbTab & Codproduto & vbTab & CodRef & vbTab & TBNivel6!Descricao & vbTab & TBNivel6!Unidade & vbTab & TBNivel6!Versao_desenho & vbTab & Kg_un & vbTab & Un_Kg & vbTab & Dim_mm & vbTab & Kg_pc & vbTab & QtTexto & vbTab & Peso_total & vbTab & DataValidacao & vbTab & RespValidacao & vbTab & PartNumber & vbTab & TBNivel6!Obs
                ElseIf Formulario = "Engenharia/Estrutura" Then
                    arrNodes(Contador1).Text = TBNivel6!Desenho & vbTab & PosicaoTexto & vbTab & Codproduto & vbTab & CodRef & vbTab & TBNivel6!Descricao & vbTab & TBNivel6!Unidade & vbTab & TBNivel6!Versao_desenho & vbTab & Kg_un & vbTab & Un_Kg & vbTab & Dim_mm & vbTab & Kg_pc & vbTab & IIf(IsNull(TBAbrir!Largura), 0, Format(TBAbrir!Largura, "###,##0.00")) & vbTab & IIf(IsNull(TBAbrir!Comprimento), "", Format(TBAbrir!Comprimento, "###,##0.00")) & vbTab & QtTexto & vbTab & Peso_total & vbTab & Format(valor, "###,##0.00000000") & vbTab & DataValidacao & vbTab & RespValidacao & vbTab & TBNivel6!CODIGO & vbTab & PartNumber & vbTab & TBNivel6!Obs
                Else
                    arrNodes(Contador1).Text = TBNivel6!Desenho & vbTab & PosicaoTexto & vbTab & Codproduto & vbTab & CodRef & vbTab & TBNivel6!Descricao & vbTab & TBNivel6!Unidade & vbTab & TBNivel6!Versao_desenho & vbTab & Kg_un & vbTab & Un_Kg & vbTab & Dim_mm & vbTab & Kg_pc & vbTab & QtTexto & vbTab & Peso_total & vbTab & Format(valor, "###,##0.00000000") & vbTab & DataValidacao & vbTab & RespValidacao & vbTab & TBNivel6!CODIGO & vbTab & PartNumber & vbTab & TBNivel6!Obs
                End If
            End If
            If CarregarNivelAbaixo = True Then ProcNivel7Estrutura NomeFormulario, TBNivel6!Versao_desenho, MostrarValores, Carteira_pedidos, CarregarNivelAbaixo, Simulacao_vendas
        End If
        TBNivel6.MoveNext
    Loop
End If

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcNivel7Estrutura(NomeFormulario As Form, versao As String, MostrarValores As Boolean, Carteira_pedidos As Boolean, CarregarNivelAbaixo As Boolean, Simulacao_vendas As Boolean)
On Error GoTo tratar_erro

If versao = "" Then TextoFiltro = " " Else TextoFiltro = " and Versao = '" & versao & "'"
Set TBNivel7 = CreateObject("adodb.recordset")
TBNivel7.Open "Select * from projconjunto where codproduto = " & Codproduto & TextoFiltro & " order by Posicao, Codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBNivel7.EOF = False Then
    Do While TBNivel7.EOF = False
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select P.codproduto, P.largura, P.comprimento, P.Subtipoitem, P.Producao, P.PCusto, PCDV.DtValidacao, PCDV.RespValidacao from projproduto P LEFT JOIN Projconjunto_desc_versao PCDV ON PCDV.Codproduto = P.Codproduto " & IIf(versao <> "", "and PCDV.Versao = '" & versao & "'", "") & " where P.desenho = '" & TBNivel7!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Codproduto = TBAbrir!Codproduto
            If TBNivel7!Unidade = "KG" And TBNivel7!PesoTotal > 0 Or TBNivel7!Unidade = "MT" And TBNivel7!Dimensoes > 0 Or TBNivel7!Unidade = "MM" And TBNivel7!Dimensoes > 0 Then
                Select Case TBNivel7!Unidade
                    Case "KG":  Peso = TBNivel7!PesoTotal
                    Case "MT":  Peso = (TBNivel7!Dimensoes * TBNivel7!quantidade) / 1000
                    Case "MM":  Peso = TBNivel7!Dimensoes * TBNivel7!quantidade
                End Select
            Else
                Peso = TBNivel7!quantidade
            End If
            
            QuantsolicitadoN6 = Peso * QuantsolicitadoN5
            FunCarregaValorEstrutura TBNivel7!CODIGO, TBNivel7!Desenho, MostrarValores, Simulacao_vendas, QuantsolicitadoN6, False
            
            If Carteira_pedidos = True Then
                Processos = "N√O"
                Set TBProcessos = CreateObject("adodb.recordset")
                TBProcessos.Open "Select PR.IDProcesso FROM Processos PR INNER JOIN Projproduto P ON PR.Codproduto = P.Codproduto WHERE P.Desenho = '" & TBNivel7!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBProcessos.EOF = False Then
                    Processos = "SIM"
                End If
                OrdemCarteira = ""
                Set TBOrdem = CreateObject("adodb.recordset")
                TBOrdem.Open "Select PR.Ordem FROM Producao PR INNER JOIN Producao_pedidos PP ON PR.Ordem = PP.Ordem where PP.IDCarteira = " & Ordem & " and PR.Desenho = '" & TBNivel7!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBOrdem.EOF = False Then
                    OrdemCarteira = TBOrdem!Ordem
                End If
            End If
            
            PosicaoTexto = FunTamanhoTextoZeroEsq(TBNivel7!Posicao, 3)
            
            CodRef = ""
            Set TBItem = CreateObject("adodb.recordset")
            TBItem.Open "Select n_referencia from item_aplicacoes where codproduto = " & TBAbrir!Codproduto, Conexao, adOpenKeyset, adLockOptimistic
            If TBItem.EOF = False Then
                CodRef = TBItem!N_referencia
            End If
            TBItem.Close
            
            If TBAbrir!SubTipoItem <> 0 Then
                DataValidacao = IIf(IsNull(TBAbrir!DtValidacao), "", TBAbrir!DtValidacao)
                RespValidacao = IIf(IsNull(TBAbrir!RespValidacao), "", TBAbrir!RespValidacao)
            Else
                DataValidacao = ""
                RespValidacao = ""
            End If
            
            PartNumber = ""
            If IsNull(TBNivel7!ID_partnumber_fabricante) = False Then
                Set TBProduto = CreateObject("adodb.recordset")
                TBProduto.Open "Select Part_number from Projproduto_fabricante where ID = " & TBNivel7!ID_partnumber_fabricante, Conexao, adOpenKeyset, adLockOptimistic
                If TBProduto.EOF = False Then PartNumber = TBProduto!Part_number
                TBProduto.Close
            End If
            
            Contador1 = Contador1 + 1
            arrNodes(Contador1).Level = 6
            If Carteira_pedidos = True Or Simulacao_vendas = True Then
                QtTexto = Format(QuantsolicitadoN6, "###,##0.0000")
                If Carteira_pedidos = True Then
                    arrNodes(Contador1).Text = TBNivel7!Desenho & vbTab & PosicaoTexto & vbTab & Codproduto & vbTab & CodRef & vbTab & TBNivel7!Descricao & vbTab & TBNivel7!Unidade & vbTab & TBNivel7!Versao_desenho & vbTab & "" & vbTab & "" & vbTab & QtTexto & vbTab & "" & vbTab & "" & vbTab & Processos & vbTab & OrdemCarteira & vbTab & DataValidacao & vbTab & RespValidacao & vbTab & PartNumber & vbTab & TBNivel7!Obs
                Else
                    arrNodes(Contador1).Text = TBNivel7!Desenho & vbTab & PosicaoTexto & vbTab & Codproduto & vbTab & CodRef & vbTab & TBNivel7!Descricao & vbTab & TBNivel7!Unidade & vbTab & TBNivel7!Versao_desenho & vbTab & QtTexto & vbTab & Format(QuantComprado, "###,##0.0000") & vbTab & Format(IIf(QuantComprado - QuantsolicitadoN6 < 0, (QuantComprado - QuantsolicitadoN6) * -1, 0), "###,##0.0000") & vbTab & Format(Valor_Cofins_Serv, "###,##0.00000000") & vbTab & Format(valor, "###,##0.00000000") & vbTab & Format(Valor_Cofins_Prod, "###,##0.00000000") & vbTab & DataValidacao & vbTab & RespValidacao & vbTab & PartNumber & vbTab & TBNivel7!Obs & IIf(QuantComprado - QuantsolicitadoN6 < 0, 1, 0)
                End If
            Else
                Kg_un = Format(TBNivel7!PesoMetro, "###,##0.0000000000")
                Un_Kg = TBNivel7!Un_Kg
                Dim_mm = Format(TBNivel7!Dimensoes, "###,##0.0000000000")
                Kg_pc = Format(TBNivel7!Peso, "###,##0.0000000000")
                QtTexto = Format(TBNivel7!quantidade, "0.000")
                Peso_total = Format(TBNivel7!PesoTotal, "###,##0.0000000000")
                If Formulario = "Engenharia/Processos" Then
                    arrNodes(Contador1).Text = TBNivel7!Desenho & vbTab & PosicaoTexto & vbTab & Codproduto & vbTab & CodRef & vbTab & TBNivel7!Descricao & vbTab & TBNivel7!Unidade & vbTab & TBNivel7!Versao_desenho & vbTab & Kg_un & vbTab & Un_Kg & vbTab & Dim_mm & vbTab & Kg_pc & vbTab & QtTexto & vbTab & Peso_total & vbTab & DataValidacao & vbTab & RespValidacao & vbTab & PartNumber & vbTab & TBNivel7!Obs
                ElseIf Formulario = "Engenharia/Estrutura" Then
                    arrNodes(Contador1).Text = TBNivel7!Desenho & vbTab & PosicaoTexto & vbTab & Codproduto & vbTab & CodRef & vbTab & TBNivel7!Descricao & vbTab & TBNivel7!Unidade & vbTab & TBNivel7!Versao_desenho & vbTab & Kg_un & vbTab & Un_Kg & vbTab & Dim_mm & vbTab & Kg_pc & vbTab & IIf(IsNull(TBAbrir!Largura), 0, Format(TBAbrir!Largura, "###,##0.00")) & vbTab & IIf(IsNull(TBAbrir!Comprimento), "", Format(TBAbrir!Comprimento, "###,##0.00")) & vbTab & QtTexto & vbTab & Peso_total & vbTab & Format(valor, "###,##0.00000000") & vbTab & DataValidacao & vbTab & RespValidacao & vbTab & TBNivel7!CODIGO & vbTab & PartNumber & vbTab & TBNivel7!Obs
                Else
                    arrNodes(Contador1).Text = TBNivel7!Desenho & vbTab & PosicaoTexto & vbTab & Codproduto & vbTab & CodRef & vbTab & TBNivel7!Descricao & vbTab & TBNivel7!Unidade & vbTab & TBNivel7!Versao_desenho & vbTab & Kg_un & vbTab & Un_Kg & vbTab & Dim_mm & vbTab & Kg_pc & vbTab & QtTexto & vbTab & Peso_total & vbTab & Format(valor, "###,##0.00000000") & vbTab & DataValidacao & vbTab & RespValidacao & vbTab & TBNivel7!CODIGO & vbTab & PartNumber & vbTab & TBNivel7!Obs
                End If
            End If
            If CarregarNivelAbaixo = True Then ProcNivel8Estrutura NomeFormulario, TBNivel7!Versao_desenho, MostrarValores, Carteira_pedidos, CarregarNivelAbaixo, Simulacao_vendas
        End If
        TBNivel7.MoveNext
    Loop
End If

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcNivel8Estrutura(NomeFormulario As Form, versao As String, MostrarValores As Boolean, Carteira_pedidos As Boolean, CarregarNivelAbaixo As Boolean, Simulacao_vendas As Boolean)
On Error GoTo tratar_erro

If versao = "" Then TextoFiltro = " " Else TextoFiltro = " and Versao = '" & versao & "'"
Set TBNivel8 = CreateObject("adodb.recordset")
TBNivel8.Open "Select * from projconjunto where codproduto = " & Codproduto & TextoFiltro & " order by Posicao, Codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBNivel8.EOF = False Then
    Do While TBNivel8.EOF = False
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select P.codproduto, P.largura, P.comprimento, P.Subtipoitem, P.Producao, P.PCusto, PCDV.DtValidacao, PCDV.RespValidacao from projproduto P LEFT JOIN Projconjunto_desc_versao PCDV ON PCDV.Codproduto = P.Codproduto " & IIf(versao <> "", "and PCDV.Versao = '" & versao & "'", "") & " where P.desenho = '" & TBNivel8!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Codproduto = TBAbrir!Codproduto
            If TBNivel8!Unidade = "KG" And TBNivel8!PesoTotal > 0 Or TBNivel8!Unidade = "MT" And TBNivel8!Dimensoes > 0 Or TBNivel8!Unidade = "MM" And TBNivel8!Dimensoes > 0 Then
                Select Case TBNivel8!Unidade
                    Case "KG":  Peso = TBNivel8!PesoTotal
                    Case "MT":  Peso = (TBNivel8!Dimensoes * TBNivel8!quantidade) / 1000
                    Case "MM":  Peso = TBNivel8!Dimensoes * TBNivel8!quantidade
                End Select
            Else
                Peso = TBNivel8!quantidade
            End If
            
            QuantsolicitadoN7 = Peso * QuantsolicitadoN6
            FunCarregaValorEstrutura TBNivel8!CODIGO, TBNivel8!Desenho, MostrarValores, Simulacao_vendas, QuantsolicitadoN7, False
            
            If Carteira_pedidos = True Then
                Processos = "N√O"
                Set TBProcessos = CreateObject("adodb.recordset")
                TBProcessos.Open "Select PR.IDProcesso FROM Processos PR INNER JOIN Projproduto P ON PR.Codproduto = P.Codproduto WHERE P.Desenho = '" & TBNivel8!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBProcessos.EOF = False Then
                    Processos = "SIM"
                End If
                OrdemCarteira = ""
                Set TBOrdem = CreateObject("adodb.recordset")
                TBOrdem.Open "Select PR.Ordem FROM Producao PR INNER JOIN Producao_pedidos PP ON PR.Ordem = PP.Ordem where PP.IDCarteira = " & Ordem & " and PR.Desenho = '" & TBNivel8!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBOrdem.EOF = False Then
                    OrdemCarteira = TBOrdem!Ordem
                End If
            End If
            
            PosicaoTexto = FunTamanhoTextoZeroEsq(TBNivel8!Posicao, 3)
            
            CodRef = ""
            Set TBItem = CreateObject("adodb.recordset")
            TBItem.Open "Select n_referencia from item_aplicacoes where codproduto = " & TBAbrir!Codproduto, Conexao, adOpenKeyset, adLockOptimistic
            If TBItem.EOF = False Then
                CodRef = TBItem!N_referencia
            End If
            TBItem.Close
            
            If TBAbrir!SubTipoItem <> 0 Then
                DataValidacao = IIf(IsNull(TBAbrir!DtValidacao), "", TBAbrir!DtValidacao)
                RespValidacao = IIf(IsNull(TBAbrir!RespValidacao), "", TBAbrir!RespValidacao)
            Else
                DataValidacao = ""
                RespValidacao = ""
            End If
            
            PartNumber = ""
            If IsNull(TBNivel8!ID_partnumber_fabricante) = False Then
                Set TBProduto = CreateObject("adodb.recordset")
                TBProduto.Open "Select Part_number from Projproduto_fabricante where ID = " & TBNivel8!ID_partnumber_fabricante, Conexao, adOpenKeyset, adLockOptimistic
                If TBProduto.EOF = False Then PartNumber = TBProduto!Part_number
                TBProduto.Close
            End If
            
            Contador1 = Contador1 + 1
            arrNodes(Contador1).Level = 7
            If Carteira_pedidos = True Or Simulacao_vendas = True Then
                QtTexto = Format(QuantsolicitadoN7, "###,##0.0000")
                If Carteira_pedidos = True Then
                    arrNodes(Contador1).Text = TBNivel8!Desenho & vbTab & PosicaoTexto & vbTab & Codproduto & vbTab & CodRef & vbTab & TBNivel8!Descricao & vbTab & TBNivel8!Unidade & vbTab & TBNivel8!Versao_desenho & vbTab & "" & vbTab & "" & vbTab & QtTexto & vbTab & "" & vbTab & "" & vbTab & Processos & vbTab & OrdemCarteira & vbTab & DataValidacao & vbTab & RespValidacao & vbTab & PartNumber & vbTab & TBNivel8!Obs
                Else
                    arrNodes(Contador1).Text = TBNivel8!Desenho & vbTab & PosicaoTexto & vbTab & Codproduto & vbTab & CodRef & vbTab & TBNivel8!Descricao & vbTab & TBNivel8!Unidade & vbTab & TBNivel8!Versao_desenho & vbTab & QtTexto & vbTab & Format(QuantComprado, "###,##0.0000") & vbTab & Format(IIf(QuantComprado - QuantsolicitadoN7 < 0, (QuantComprado - QuantsolicitadoN7) * -1, 0), "###,##0.0000") & vbTab & Format(Valor_Cofins_Serv, "###,##0.00000000") & vbTab & Format(valor, "###,##0.00000000") & vbTab & Format(Valor_Cofins_Prod, "###,##0.00000000") & vbTab & DataValidacao & vbTab & RespValidacao & vbTab & PartNumber & vbTab & TBNivel8!Obs & IIf(QuantComprado - QuantsolicitadoN7 < 0, 1, 0)
                End If
            Else
                Kg_un = Format(TBNivel8!PesoMetro, "###,##0.0000000000")
                Un_Kg = TBNivel8!Un_Kg
                Dim_mm = Format(TBNivel8!Dimensoes, "###,##0.0000000000")
                Kg_pc = Format(TBNivel8!Peso, "###,##0.0000000000")
                QtTexto = Format(TBNivel8!quantidade, "0.000")
                Peso_total = Format(TBNivel8!PesoTotal, "###,##0.0000000000")
                If Formulario = "Engenharia/Processos" Then
                    arrNodes(Contador1).Text = TBNivel8!Desenho & vbTab & PosicaoTexto & vbTab & Codproduto & vbTab & CodRef & vbTab & TBNivel8!Descricao & vbTab & TBNivel8!Unidade & vbTab & TBNivel8!Versao_desenho & vbTab & Kg_un & vbTab & Un_Kg & vbTab & Dim_mm & vbTab & Kg_pc & vbTab & QtTexto & vbTab & Peso_total & vbTab & DataValidacao & vbTab & RespValidacao & vbTab & PartNumber & vbTab & TBNivel8!Obs
                ElseIf Formulario = "Engenharia/Estrutura" Then
                    arrNodes(Contador1).Text = TBNivel8!Desenho & vbTab & PosicaoTexto & vbTab & Codproduto & vbTab & CodRef & vbTab & TBNivel8!Descricao & vbTab & TBNivel8!Unidade & vbTab & TBNivel8!Versao_desenho & vbTab & Kg_un & vbTab & Un_Kg & vbTab & Dim_mm & vbTab & Kg_pc & vbTab & IIf(IsNull(TBAbrir!Largura), 0, Format(TBAbrir!Largura, "###,##0.00")) & vbTab & IIf(IsNull(TBAbrir!Comprimento), "", Format(TBAbrir!Comprimento, "###,##0.00")) & vbTab & QtTexto & vbTab & Peso_total & vbTab & Format(valor, "###,##0.00000000") & vbTab & DataValidacao & vbTab & RespValidacao & vbTab & TBNivel8!CODIGO & vbTab & PartNumber & vbTab & TBNivel8!Obs
                Else
                    arrNodes(Contador1).Text = TBNivel8!Desenho & vbTab & PosicaoTexto & vbTab & Codproduto & vbTab & CodRef & vbTab & TBNivel8!Descricao & vbTab & TBNivel8!Unidade & vbTab & TBNivel8!Versao_desenho & vbTab & Kg_un & vbTab & Un_Kg & vbTab & Dim_mm & vbTab & Kg_pc & vbTab & QtTexto & vbTab & Peso_total & vbTab & Format(valor, "###,##0.00000000") & vbTab & DataValidacao & vbTab & RespValidacao & vbTab & TBNivel8!CODIGO & vbTab & PartNumber & vbTab & TBNivel8!Obs
                End If
            End If
            If CarregarNivelAbaixo = True Then ProcNivel9Estrutura NomeFormulario, TBNivel8!Versao_desenho, MostrarValores, Carteira_pedidos, CarregarNivelAbaixo, Simulacao_vendas
        End If
        TBNivel8.MoveNext
    Loop
End If

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcNivel9Estrutura(NomeFormulario As Form, versao As String, MostrarValores As Boolean, Carteira_pedidos As Boolean, CarregarNivelAbaixo As Boolean, Simulacao_vendas As Boolean)
On Error GoTo tratar_erro

If versao = "" Then TextoFiltro = " " Else TextoFiltro = " and Versao = '" & versao & "'"
Set TBNivel9 = CreateObject("adodb.recordset")
TBNivel9.Open "Select * from projconjunto where codproduto = " & Codproduto & TextoFiltro & " order by Posicao, Codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBNivel9.EOF = False Then
    Do While TBNivel9.EOF = False
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select P.codproduto, P.largura, P.comprimento, P.Subtipoitem, P.Producao, P.PCusto, PCDV.DtValidacao, PCDV.RespValidacao from projproduto P LEFT JOIN Projconjunto_desc_versao PCDV ON PCDV.Codproduto = P.Codproduto " & IIf(versao <> "", "and PCDV.Versao = '" & versao & "'", "") & " where P.desenho = '" & TBNivel9!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Codproduto = TBAbrir!Codproduto
            If TBNivel9!Unidade = "KG" And TBNivel9!PesoTotal > 0 Or TBNivel9!Unidade = "MT" And TBNivel9!Dimensoes > 0 Or TBNivel9!Unidade = "MM" And TBNivel9!Dimensoes > 0 Then
                Select Case TBNivel9!Unidade
                    Case "KG":  Peso = TBNivel9!PesoTotal
                    Case "MT":  Peso = (TBNivel9!Dimensoes * TBNivel9!quantidade) / 1000
                    Case "MM":  Peso = TBNivel9!Dimensoes * TBNivel9!quantidade
                End Select
            Else
                Peso = TBNivel9!quantidade
            End If
            
            QuantsolicitadoN8 = Peso * QuantsolicitadoN7
            FunCarregaValorEstrutura TBNivel9!CODIGO, TBNivel9!Desenho, MostrarValores, Simulacao_vendas, QuantsolicitadoN8, False
            
            If Carteira_pedidos = True Then
                Processos = "N√O"
                Set TBProcessos = CreateObject("adodb.recordset")
                TBProcessos.Open "Select PR.IDProcesso FROM Processos PR INNER JOIN Projproduto P ON PR.Codproduto = P.Codproduto WHERE P.Desenho = '" & TBNivel9!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBProcessos.EOF = False Then
                    Processos = "SIM"
                End If
                OrdemCarteira = ""
                Set TBOrdem = CreateObject("adodb.recordset")
                TBOrdem.Open "Select PR.Ordem FROM Producao PR INNER JOIN Producao_pedidos PP ON PR.Ordem = PP.Ordem where PP.IDCarteira = " & Ordem & " and PR.Desenho = '" & TBNivel9!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBOrdem.EOF = False Then
                    OrdemCarteira = TBOrdem!Ordem
                End If
            End If
                        
            PosicaoTexto = FunTamanhoTextoZeroEsq(TBNivel9!Posicao, 3)
            
            CodRef = ""
            Set TBItem = CreateObject("adodb.recordset")
            TBItem.Open "Select n_referencia from item_aplicacoes where codproduto = " & TBAbrir!Codproduto, Conexao, adOpenKeyset, adLockOptimistic
            If TBItem.EOF = False Then
                CodRef = TBItem!N_referencia
            End If
            TBItem.Close
            
            If TBAbrir!SubTipoItem <> 0 Then
                DataValidacao = IIf(IsNull(TBAbrir!DtValidacao), "", TBAbrir!DtValidacao)
                RespValidacao = IIf(IsNull(TBAbrir!RespValidacao), "", TBAbrir!RespValidacao)
            Else
                DataValidacao = ""
                RespValidacao = ""
            End If
            
            PartNumber = ""
            If IsNull(TBNivel9!ID_partnumber_fabricante) = False Then
                Set TBProduto = CreateObject("adodb.recordset")
                TBProduto.Open "Select Part_number from Projproduto_fabricante where ID = " & TBNivel9!ID_partnumber_fabricante, Conexao, adOpenKeyset, adLockOptimistic
                If TBProduto.EOF = False Then PartNumber = TBProduto!Part_number
                TBProduto.Close
            End If
            
            Contador1 = Contador1 + 1
            arrNodes(Contador1).Level = 8
            If Carteira_pedidos = True Or Simulacao_vendas = True Then
                QtTexto = Format(QuantsolicitadoN8, "###,##0.0000")
                If Carteira_pedidos = True Then
                    arrNodes(Contador1).Text = TBNivel9!Desenho & vbTab & PosicaoTexto & vbTab & Codproduto & vbTab & CodRef & vbTab & TBNivel9!Descricao & vbTab & TBNivel9!Unidade & vbTab & TBNivel9!Versao_desenho & vbTab & "" & vbTab & "" & vbTab & QtTexto & vbTab & "" & vbTab & "" & vbTab & Processos & vbTab & OrdemCarteira & vbTab & DataValidacao & vbTab & RespValidacao & vbTab & PartNumber & vbTab & TBNivel9!Obs
                Else
                    arrNodes(Contador1).Text = TBNivel9!Desenho & vbTab & PosicaoTexto & vbTab & Codproduto & vbTab & CodRef & vbTab & TBNivel9!Descricao & vbTab & TBNivel9!Unidade & vbTab & TBNivel9!Versao_desenho & vbTab & QtTexto & vbTab & Format(QuantComprado, "###,##0.0000") & vbTab & Format(IIf(QuantComprado - QuantsolicitadoN8 < 0, (QuantComprado - QuantsolicitadoN8) * -1, 0), "###,##0.0000") & vbTab & Format(Valor_Cofins_Serv, "###,##0.00000000") & vbTab & Format(valor, "###,##0.00000000") & vbTab & Format(Valor_Cofins_Prod, "###,##0.00000000") & vbTab & DataValidacao & vbTab & RespValidacao & vbTab & PartNumber & vbTab & TBNivel9!Obs & IIf(QuantComprado - QuantsolicitadoN8 < 0, 1, 0)
                End If
            Else
                Kg_un = Format(TBNivel9!PesoMetro, "###,##0.0000000000")
                Un_Kg = TBNivel9!Un_Kg
                Dim_mm = Format(TBNivel9!Dimensoes, "###,##0.0000000000")
                Kg_pc = Format(TBNivel9!Peso, "###,##0.0000000000")
                QtTexto = Format(TBNivel9!quantidade, "0.000")
                Peso_total = Format(TBNivel9!PesoTotal, "###,##0.0000000000")
                If Formulario = "Engenharia/Processos" Then
                    arrNodes(Contador1).Text = TBNivel9!Desenho & vbTab & PosicaoTexto & vbTab & Codproduto & vbTab & CodRef & vbTab & TBNivel9!Descricao & vbTab & TBNivel9!Unidade & vbTab & TBNivel9!Versao_desenho & vbTab & Kg_un & vbTab & Un_Kg & vbTab & Dim_mm & vbTab & Kg_pc & vbTab & QtTexto & vbTab & Peso_total & vbTab & DataValidacao & vbTab & RespValidacao & vbTab & PartNumber & vbTab & TBNivel9!Obs
                ElseIf Formulario = "Engenharia/Estrutura" Then
                    arrNodes(Contador1).Text = TBNivel9!Desenho & vbTab & PosicaoTexto & vbTab & Codproduto & vbTab & CodRef & vbTab & TBNivel9!Descricao & vbTab & TBNivel9!Unidade & vbTab & TBNivel9!Versao_desenho & vbTab & Kg_un & vbTab & Un_Kg & vbTab & Dim_mm & vbTab & Kg_pc & vbTab & IIf(IsNull(TBAbrir!Largura), 0, Format(TBAbrir!Largura, "###,##0.00")) & vbTab & IIf(IsNull(TBAbrir!Comprimento), "", Format(TBAbrir!Comprimento, "###,##0.00")) & vbTab & QtTexto & vbTab & Peso_total & vbTab & Format(valor, "###,##0.00000000") & vbTab & DataValidacao & vbTab & RespValidacao & vbTab & TBNivel9!CODIGO & vbTab & PartNumber & vbTab & TBNivel9!Obs
                Else
                    arrNodes(Contador1).Text = TBNivel9!Desenho & vbTab & PosicaoTexto & vbTab & Codproduto & vbTab & CodRef & vbTab & TBNivel9!Descricao & vbTab & TBNivel9!Unidade & vbTab & TBNivel9!Versao_desenho & vbTab & Kg_un & vbTab & Un_Kg & vbTab & Dim_mm & vbTab & Kg_pc & vbTab & QtTexto & vbTab & Peso_total & vbTab & Format(valor, "###,##0.00000000") & vbTab & DataValidacao & vbTab & RespValidacao & vbTab & TBNivel9!CODIGO & vbTab & PartNumber & vbTab & TBNivel9!Obs
                End If
            End If
            If CarregarNivelAbaixo = True Then ProcNivel10Estrutura NomeFormulario, TBNivel9!Versao_desenho, MostrarValores, Carteira_pedidos, CarregarNivelAbaixo, Simulacao_vendas
        End If
        TBNivel9.MoveNext
    Loop
End If

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcNivel10Estrutura(NomeFormulario As Form, versao As String, MostrarValores As Boolean, Carteira_pedidos As Boolean, CarregarNivelAbaixo As Boolean, Simulacao_vendas As Boolean)
On Error GoTo tratar_erro

If versao = "" Then TextoFiltro = " " Else TextoFiltro = " and Versao = '" & versao & "'"
Set TBNivel10 = CreateObject("adodb.recordset")
TBNivel10.Open "Select * from projconjunto where codproduto = " & Codproduto & TextoFiltro & " order by Posicao, Codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBNivel10.EOF = False Then
    Do While TBNivel10.EOF = False
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select P.codproduto, P.largura, P.comprimento, P.Subtipoitem, P.Producao, P.PCusto, PCDV.DtValidacao, PCDV.RespValidacao from projproduto P LEFT JOIN Projconjunto_desc_versao PCDV ON PCDV.Codproduto = P.Codproduto " & IIf(versao <> "", "and PCDV.Versao = '" & versao & "'", "") & " where P.desenho = '" & TBNivel10!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Codproduto = TBAbrir!Codproduto
            If TBNivel10!Unidade = "KG" And TBNivel10!PesoTotal > 0 Or TBNivel10!Unidade = "MT" And TBNivel10!Dimensoes > 0 Or TBNivel10!Unidade = "MM" And TBNivel10!Dimensoes > 0 Then
                Select Case TBNivel10!Unidade
                    Case "KG":  Peso = TBNivel10!PesoTotal
                    Case "MT":  Peso = (TBNivel10!Dimensoes * TBNivel10!quantidade) / 1000
                    Case "MM":  Peso = TBNivel10!Dimensoes * TBNivel10!quantidade
                End Select
            Else
                Peso = TBNivel10!quantidade
            End If
            
            QuantsolicitadoN9 = Peso * QuantsolicitadoN8
            FunCarregaValorEstrutura TBNivel10!CODIGO, TBNivel10!Desenho, MostrarValores, Simulacao_vendas, QuantsolicitadoN9, False
            
            If Carteira_pedidos = True Then
                Processos = "N√O"
                Set TBProcessos = CreateObject("adodb.recordset")
                TBProcessos.Open "Select PR.IDProcesso FROM Processos PR INNER JOIN Projproduto P ON PR.Codproduto = P.Codproduto WHERE P.Desenho = '" & TBNivel10!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBProcessos.EOF = False Then
                    Processos = "SIM"
                End If
                OrdemCarteira = ""
                Set TBOrdem = CreateObject("adodb.recordset")
                TBOrdem.Open "Select PR.Ordem FROM Producao PR INNER JOIN Producao_pedidos PP ON PR.Ordem = PP.Ordem where PP.IDCarteira = " & Ordem & " and PR.Desenho = '" & TBNivel10!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBOrdem.EOF = False Then
                    OrdemCarteira = TBOrdem!Ordem
                End If
            End If
            
            PosicaoTexto = FunTamanhoTextoZeroEsq(TBNivel10!Posicao, 3)
            
            CodRef = ""
            Set TBItem = CreateObject("adodb.recordset")
            TBItem.Open "Select n_referencia from item_aplicacoes where codproduto = " & TBAbrir!Codproduto, Conexao, adOpenKeyset, adLockOptimistic
            If TBItem.EOF = False Then
                CodRef = TBItem!N_referencia
            End If
            TBItem.Close
            
            If TBAbrir!SubTipoItem <> 0 Then
                DataValidacao = IIf(IsNull(TBAbrir!DtValidacao), "", TBAbrir!DtValidacao)
                RespValidacao = IIf(IsNull(TBAbrir!RespValidacao), "", TBAbrir!RespValidacao)
            Else
                DataValidacao = ""
                RespValidacao = ""
            End If
            
            PartNumber = ""
            If IsNull(TBNivel10!ID_partnumber_fabricante) = False Then
                Set TBProduto = CreateObject("adodb.recordset")
                TBProduto.Open "Select Part_number from Projproduto_fabricante where ID = " & TBNivel10!ID_partnumber_fabricante, Conexao, adOpenKeyset, adLockOptimistic
                If TBProduto.EOF = False Then PartNumber = TBProduto!Part_number
                TBProduto.Close
            End If
            
            Contador1 = Contador1 + 1
            arrNodes(Contador1).Level = 9
            If Carteira_pedidos = True Or Simulacao_vendas = True Then
                QtTexto = Format(QuantsolicitadoN9, "###,##0.0000")
                If Carteira_pedidos = True Then
                    arrNodes(Contador1).Text = TBNivel10!Desenho & vbTab & PosicaoTexto & vbTab & Codproduto & vbTab & CodRef & vbTab & TBNivel10!Descricao & vbTab & TBNivel10!Unidade & vbTab & TBNivel10!Versao_desenho & vbTab & "" & vbTab & "" & vbTab & QtTexto & vbTab & "" & vbTab & "" & vbTab & Processos & vbTab & OrdemCarteira & vbTab & DataValidacao & vbTab & RespValidacao & vbTab & PartNumber & vbTab & TBNivel10!Obs
                Else
                    arrNodes(Contador1).Text = TBNivel10!Desenho & vbTab & PosicaoTexto & vbTab & Codproduto & vbTab & CodRef & vbTab & TBNivel10!Descricao & vbTab & TBNivel10!Unidade & vbTab & TBNivel10!Versao_desenho & vbTab & QtTexto & vbTab & Format(QuantComprado, "###,##0.0000") & vbTab & Format(IIf(QuantComprado - QuantsolicitadoN9 < 0, (QuantComprado - QuantsolicitadoN9) * -1, 0), "###,##0.0000") & vbTab & Format(Valor_Cofins_Serv, "###,##0.00000000") & vbTab & Format(valor, "###,##0.00000000") & vbTab & Format(Valor_Cofins_Prod, "###,##0.00000000") & vbTab & DataValidacao & vbTab & RespValidacao & vbTab & PartNumber & vbTab & TBNivel10!Obs & IIf(QuantComprado - QuantsolicitadoN9 < 0, 1, 0)
                End If
            Else
                Kg_un = Format(TBNivel10!PesoMetro, "###,##0.0000000000")
                Un_Kg = TBNivel10!Un_Kg
                Dim_mm = Format(TBNivel10!Dimensoes, "###,##0.0000000000")
                Kg_pc = Format(TBNivel10!Peso, "###,##0.0000000000")
                QtTexto = Format(TBNivel10!quantidade, "0.000")
                Peso_total = Format(TBNivel10!PesoTotal, "###,##0.0000000000")
                If Formulario = "Engenharia/Processos" Then
                    arrNodes(Contador1).Text = TBNivel10!Desenho & vbTab & PosicaoTexto & vbTab & Codproduto & vbTab & CodRef & vbTab & TBNivel10!Descricao & vbTab & TBNivel10!Unidade & vbTab & TBNivel10!Versao_desenho & vbTab & Kg_un & vbTab & Un_Kg & vbTab & Dim_mm & vbTab & Kg_pc & vbTab & QtTexto & vbTab & Peso_total & vbTab & DataValidacao & vbTab & RespValidacao & vbTab & PartNumber & vbTab & TBNivel10!Obs
                ElseIf Formulario = "Engenharia/Estrutura" Then
                    arrNodes(Contador1).Text = TBNivel10!Desenho & vbTab & PosicaoTexto & vbTab & Codproduto & vbTab & CodRef & vbTab & TBNivel10!Descricao & vbTab & TBNivel10!Unidade & vbTab & TBNivel10!Versao_desenho & vbTab & Kg_un & vbTab & Un_Kg & vbTab & Dim_mm & vbTab & Kg_pc & vbTab & IIf(IsNull(TBAbrir!Largura), 0, Format(TBAbrir!Largura, "###,##0.00")) & vbTab & IIf(IsNull(TBAbrir!Comprimento), "", Format(TBAbrir!Comprimento, "###,##0.00")) & vbTab & QtTexto & vbTab & Peso_total & vbTab & Format(valor, "###,##0.00000000") & vbTab & DataValidacao & vbTab & RespValidacao & vbTab & TBNivel10!CODIGO & vbTab & PartNumber & vbTab & TBNivel10!Obs
                Else
                    arrNodes(Contador1).Text = TBNivel10!Desenho & vbTab & PosicaoTexto & vbTab & Codproduto & vbTab & CodRef & vbTab & TBNivel10!Descricao & vbTab & TBNivel10!Unidade & vbTab & TBNivel10!Versao_desenho & vbTab & Kg_un & vbTab & Un_Kg & vbTab & Dim_mm & vbTab & Kg_pc & vbTab & QtTexto & vbTab & Peso_total & vbTab & Format(valor, "###,##0.00000000") & vbTab & DataValidacao & vbTab & RespValidacao & vbTab & TBNivel10!CODIGO & vbTab & PartNumber & vbTab & TBNivel10!Obs
                End If
            End If
            If CarregarNivelAbaixo = True Then ProcNivel11Estrutura NomeFormulario, TBNivel10!Versao_desenho, MostrarValores, Carteira_pedidos, CarregarNivelAbaixo, Simulacao_vendas
        End If
        TBNivel10.MoveNext
    Loop
End If

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcNivel11Estrutura(NomeFormulario As Form, versao As String, MostrarValores As Boolean, Carteira_pedidos As Boolean, CarregarNivelAbaixo As Boolean, Simulacao_vendas As Boolean)
On Error GoTo tratar_erro

If versao = "" Then TextoFiltro = " " Else TextoFiltro = " and Versao = '" & versao & "'"
Set TBNivel11 = CreateObject("adodb.recordset")
TBNivel11.Open "Select * from projconjunto where codproduto = " & Codproduto & TextoFiltro & " order by Posicao, Codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBNivel11.EOF = False Then
    Do While TBNivel11.EOF = False
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select P.codproduto, P.largura, P.comprimento, P.Subtipoitem, P.Producao, P.PCusto, PCDV.DtValidacao, PCDV.RespValidacao from projproduto P LEFT JOIN Projconjunto_desc_versao PCDV ON PCDV.Codproduto = P.Codproduto " & IIf(versao <> "", "and PCDV.Versao = '" & versao & "'", "") & " where P.desenho = '" & TBNivel11!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Codproduto = TBAbrir!Codproduto
            If TBNivel11!Unidade = "KG" And TBNivel11!PesoTotal > 0 Or TBNivel11!Unidade = "MT" And TBNivel11!Dimensoes > 0 Or TBNivel11!Unidade = "MM" And TBNivel11!Dimensoes > 0 Then
                Select Case TBNivel11!Unidade
                    Case "KG":  Peso = TBNivel11!PesoTotal
                    Case "MT":  Peso = (TBNivel11!Dimensoes * TBNivel11!quantidade) / 1000
                    Case "MM":  Peso = TBNivel11!Dimensoes * TBNivel11!quantidade
                End Select
            Else
                Peso = TBNivel11!quantidade
            End If
                        
            QuantsolicitadoN10 = Peso * QuantsolicitadoN9
            FunCarregaValorEstrutura TBNivel11!CODIGO, TBNivel11!Desenho, MostrarValores, Simulacao_vendas, QuantsolicitadoN10, False
            
            If Carteira_pedidos = True Then
                Processos = "N√O"
                Set TBProcessos = CreateObject("adodb.recordset")
                TBProcessos.Open "Select PR.IDProcesso FROM Processos PR INNER JOIN Projproduto P ON PR.Codproduto = P.Codproduto WHERE P.Desenho = '" & TBNivel11!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBProcessos.EOF = False Then
                    Processos = "SIM"
                End If
                OrdemCarteira = ""
                Set TBOrdem = CreateObject("adodb.recordset")
                TBOrdem.Open "Select PR.Ordem FROM Producao PR INNER JOIN Producao_pedidos PP ON PR.Ordem = PP.Ordem where PP.IDCarteira = " & Ordem & " and PR.Desenho = '" & TBNivel11!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBOrdem.EOF = False Then
                    OrdemCarteira = TBOrdem!Ordem
                End If
            End If
            
            PosicaoTexto = FunTamanhoTextoZeroEsq(TBNivel11!Posicao, 3)
            
            CodRef = ""
            Set TBItem = CreateObject("adodb.recordset")
            TBItem.Open "Select n_referencia from item_aplicacoes where codproduto = " & TBAbrir!Codproduto, Conexao, adOpenKeyset, adLockOptimistic
            If TBItem.EOF = False Then
                CodRef = TBItem!N_referencia
            End If
            TBItem.Close
            
            If TBAbrir!SubTipoItem <> 0 Then
                DataValidacao = IIf(IsNull(TBAbrir!DtValidacao), "", TBAbrir!DtValidacao)
                RespValidacao = IIf(IsNull(TBAbrir!RespValidacao), "", TBAbrir!RespValidacao)
            Else
                DataValidacao = ""
                RespValidacao = ""
            End If
            
            PartNumber = ""
            If IsNull(TBNivel11!ID_partnumber_fabricante) = False Then
                Set TBProduto = CreateObject("adodb.recordset")
                TBProduto.Open "Select Part_number from Projproduto_fabricante where ID = " & TBNivel11!ID_partnumber_fabricante, Conexao, adOpenKeyset, adLockOptimistic
                If TBProduto.EOF = False Then PartNumber = TBProduto!Part_number
                TBProduto.Close
            End If
            
            Contador1 = Contador1 + 1
            arrNodes(Contador1).Level = 10
            If Carteira_pedidos = True Or Simulacao_vendas = True Then
                QtTexto = Format(QuantsolicitadoN10, "###,##0.0000")
                If Carteira_pedidos = True Then
                    arrNodes(Contador1).Text = TBNivel11!Desenho & vbTab & PosicaoTexto & vbTab & Codproduto & vbTab & CodRef & vbTab & TBNivel11!Descricao & vbTab & TBNivel11!Unidade & vbTab & TBNivel11!Versao_desenho & vbTab & "" & vbTab & "" & vbTab & QtTexto & vbTab & "" & vbTab & "" & vbTab & Processos & vbTab & OrdemCarteira & vbTab & DataValidacao & vbTab & RespValidacao & vbTab & PartNumber & vbTab & TBNivel11!Obs
                Else
                    arrNodes(Contador1).Text = TBNivel11!Desenho & vbTab & PosicaoTexto & vbTab & Codproduto & vbTab & CodRef & vbTab & TBNivel11!Descricao & vbTab & TBNivel11!Unidade & vbTab & TBNivel11!Versao_desenho & vbTab & QtTexto & vbTab & Format(QuantComprado, "###,##0.0000") & vbTab & Format(IIf(QuantComprado - QuantsolicitadoN10 < 0, (QuantComprado - QuantsolicitadoN10) * -1, 0), "###,##0.0000") & vbTab & Format(Valor_Cofins_Serv, "###,##0.00000000") & vbTab & Format(valor, "###,##0.00000000") & vbTab & Format(Valor_Cofins_Prod, "###,##0.00000000") & vbTab & DataValidacao & vbTab & RespValidacao & vbTab & PartNumber & vbTab & TBNivel11!Obs & IIf(QuantComprado - QuantsolicitadoN10 < 0, 1, 0)
                End If
            Else
                Kg_un = Format(TBNivel11!PesoMetro, "###,##0.0000000000")
                Un_Kg = TBNivel11!Un_Kg
                Dim_mm = Format(TBNivel11!Dimensoes, "###,##0.0000000000")
                Kg_pc = Format(TBNivel11!Peso, "###,##0.0000000000")
                QtTexto = Format(TBNivel11!quantidade, "0.000")
                Peso_total = Format(TBNivel11!PesoTotal, "###,##0.0000000000")
                If Formulario = "Engenharia/Processos" Then
                    arrNodes(Contador1).Text = TBNivel11!Desenho & vbTab & PosicaoTexto & vbTab & Codproduto & vbTab & CodRef & vbTab & TBNivel11!Descricao & vbTab & TBNivel11!Unidade & vbTab & TBNivel11!Versao_desenho & vbTab & Kg_un & vbTab & Un_Kg & vbTab & Dim_mm & vbTab & Kg_pc & vbTab & QtTexto & vbTab & Peso_total & vbTab & DataValidacao & vbTab & RespValidacao & vbTab & PartNumber & vbTab & TBNivel11!Obs
                ElseIf Formulario = "Engenharia/Estrutura" Then
                    arrNodes(Contador1).Text = TBNivel11!Desenho & vbTab & PosicaoTexto & vbTab & Codproduto & vbTab & CodRef & vbTab & TBNivel11!Descricao & vbTab & TBNivel11!Unidade & vbTab & TBNivel11!Versao_desenho & vbTab & Kg_un & vbTab & Un_Kg & vbTab & Dim_mm & vbTab & Kg_pc & vbTab & IIf(IsNull(TBAbrir!Largura), 0, Format(TBAbrir!Largura, "###,##0.00")) & vbTab & IIf(IsNull(TBAbrir!Comprimento), "", Format(TBAbrir!Comprimento, "###,##0.00")) & vbTab & QtTexto & vbTab & Peso_total & vbTab & Format(valor, "###,##0.00000000") & vbTab & DataValidacao & vbTab & RespValidacao & vbTab & TBNivel11!CODIGO & vbTab & PartNumber & vbTab & TBNivel11!Obs
                Else
                    arrNodes(Contador1).Text = TBNivel11!Desenho & vbTab & PosicaoTexto & vbTab & Codproduto & vbTab & CodRef & vbTab & TBNivel11!Descricao & vbTab & TBNivel11!Unidade & vbTab & TBNivel11!Versao_desenho & vbTab & Kg_un & vbTab & Un_Kg & vbTab & Dim_mm & vbTab & Kg_pc & vbTab & QtTexto & vbTab & Peso_total & vbTab & Format(valor, "###,##0.00000000") & vbTab & DataValidacao & vbTab & RespValidacao & vbTab & TBNivel11!CODIGO & vbTab & PartNumber & vbTab & TBNivel11!Obs
                End If
            End If
            If CarregarNivelAbaixo = True Then ProcNivel12Estrutura NomeFormulario, TBNivel11!Versao_desenho, MostrarValores, Carteira_pedidos, CarregarNivelAbaixo, Simulacao_vendas
        End If
        TBNivel11.MoveNext
    Loop
End If

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcNivel12Estrutura(NomeFormulario As Form, versao As String, MostrarValores As Boolean, Carteira_pedidos As Boolean, CarregarNivelAbaixo As Boolean, Simulacao_vendas As Boolean)
On Error GoTo tratar_erro

If versao = "" Then TextoFiltro = " " Else TextoFiltro = " and Versao = '" & versao & "'"
Set TBNivel12 = CreateObject("adodb.recordset")
TBNivel12.Open "Select * from projconjunto where codproduto = " & Codproduto & TextoFiltro & " order by Posicao, Codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBNivel12.EOF = False Then
    Do While TBNivel12.EOF = False
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select P.codproduto, P.largura, P.comprimento, P.Subtipoitem, P.Producao, P.PCusto, PCDV.DtValidacao, PCDV.RespValidacao from projproduto P LEFT JOIN Projconjunto_desc_versao PCDV ON PCDV.Codproduto = P.Codproduto " & IIf(versao <> "", "and PCDV.Versao = '" & versao & "'", "") & " where P.desenho = '" & TBNivel12!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Codproduto = TBAbrir!Codproduto
            If TBNivel12!Unidade = "KG" And TBNivel12!PesoTotal > 0 Or TBNivel12!Unidade = "MT" And TBNivel12!Dimensoes > 0 Or TBNivel12!Unidade = "MM" And TBNivel12!Dimensoes > 0 Then
                Select Case TBNivel12!Unidade
                    Case "KG":  Peso = TBNivel12!PesoTotal
                    Case "MT":  Peso = (TBNivel12!Dimensoes * TBNivel12!quantidade) / 1000
                    Case "MM":  Peso = TBNivel12!Dimensoes * TBNivel12!quantidade
                End Select
            Else
                Peso = TBNivel12!quantidade
            End If
            
            QuantsolicitadoN11 = Peso * QuantsolicitadoN10
            FunCarregaValorEstrutura TBNivel12!CODIGO, TBNivel12!Desenho, MostrarValores, Simulacao_vendas, QuantsolicitadoN11, False
            
            If Carteira_pedidos = True Then
                Processos = "N√O"
                Set TBProcessos = CreateObject("adodb.recordset")
                TBProcessos.Open "Select PR.IDProcesso FROM Processos PR INNER JOIN Projproduto P ON PR.Codproduto = P.Codproduto WHERE P.Desenho = '" & TBNivel12!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBProcessos.EOF = False Then
                    Processos = "SIM"
                End If
                OrdemCarteira = ""
                Set TBOrdem = CreateObject("adodb.recordset")
                TBOrdem.Open "Select PR.Ordem FROM Producao PR INNER JOIN Producao_pedidos PP ON PR.Ordem = PP.Ordem where PP.IDCarteira = " & Ordem & " and PR.Desenho = '" & TBNivel12!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBOrdem.EOF = False Then
                    OrdemCarteira = TBOrdem!Ordem
                End If
            End If
            
            PosicaoTexto = FunTamanhoTextoZeroEsq(TBNivel12!Posicao, 3)
            
            CodRef = ""
            Set TBItem = CreateObject("adodb.recordset")
            TBItem.Open "Select n_referencia from item_aplicacoes where codproduto = " & TBAbrir!Codproduto, Conexao, adOpenKeyset, adLockOptimistic
            If TBItem.EOF = False Then
                CodRef = TBItem!N_referencia
            End If
            TBItem.Close
            
            If TBAbrir!SubTipoItem <> 0 Then
                DataValidacao = IIf(IsNull(TBAbrir!DtValidacao), "", TBAbrir!DtValidacao)
                RespValidacao = IIf(IsNull(TBAbrir!RespValidacao), "", TBAbrir!RespValidacao)
            Else
                DataValidacao = ""
                RespValidacao = ""
            End If
            
            PartNumber = ""
            If IsNull(TBNivel12!ID_partnumber_fabricante) = False Then
                Set TBProduto = CreateObject("adodb.recordset")
                TBProduto.Open "Select Part_number from Projproduto_fabricante where ID = " & TBNivel12!ID_partnumber_fabricante, Conexao, adOpenKeyset, adLockOptimistic
                If TBProduto.EOF = False Then PartNumber = TBProduto!Part_number
                TBProduto.Close
            End If
            
            Contador1 = Contador1 + 1
            arrNodes(Contador1).Level = 11
            If Carteira_pedidos = True Or Simulacao_vendas = True Then
                QtTexto = Format(QuantsolicitadoN11, "###,##0.0000")
                If Carteira_pedidos = True Then
                    arrNodes(Contador1).Text = TBNivel12!Desenho & vbTab & PosicaoTexto & vbTab & Codproduto & vbTab & CodRef & vbTab & TBNivel12!Descricao & vbTab & TBNivel12!Unidade & vbTab & TBNivel12!Versao_desenho & vbTab & "" & vbTab & "" & vbTab & QtTexto & vbTab & "" & vbTab & "" & vbTab & Processos & vbTab & OrdemCarteira & vbTab & DataValidacao & vbTab & RespValidacao & vbTab & PartNumber & vbTab & TBNivel12!Obs
                Else
                    arrNodes(Contador1).Text = TBNivel12!Desenho & vbTab & PosicaoTexto & vbTab & Codproduto & vbTab & CodRef & vbTab & TBNivel12!Descricao & vbTab & TBNivel12!Unidade & vbTab & TBNivel12!Versao_desenho & vbTab & QtTexto & vbTab & Format(QuantComprado, "###,##0.0000") & vbTab & Format(IIf(QuantComprado - QuantsolicitadoN11 < 0, (QuantComprado - QuantsolicitadoN11) * -1, 0), "###,##0.0000") & vbTab & Format(Valor_Cofins_Serv, "###,##0.00000000") & vbTab & Format(valor, "###,##0.00000000") & vbTab & Format(Valor_Cofins_Prod, "###,##0.00000000") & vbTab & DataValidacao & vbTab & RespValidacao & vbTab & PartNumber & vbTab & TBNivel12!Obs & IIf(QuantComprado - QuantsolicitadoN11 < 0, 1, 0)
                End If
            Else
                Kg_un = Format(TBNivel12!PesoMetro, "###,##0.0000000000")
                Un_Kg = TBNivel12!Un_Kg
                Dim_mm = Format(TBNivel12!Dimensoes, "###,##0.0000000000")
                Kg_pc = Format(TBNivel12!Peso, "###,##0.0000000000")
                QtTexto = Format(TBNivel12!quantidade, "0.000")
                Peso_total = Format(TBNivel12!PesoTotal, "###,##0.0000000000")
                If Formulario = "Engenharia/Processos" Then
                    arrNodes(Contador1).Text = TBNivel12!Desenho & vbTab & PosicaoTexto & vbTab & Codproduto & vbTab & CodRef & vbTab & TBNivel12!Descricao & vbTab & TBNivel12!Unidade & vbTab & TBNivel12!Versao_desenho & vbTab & Kg_un & vbTab & Un_Kg & vbTab & Dim_mm & vbTab & Kg_pc & vbTab & QtTexto & vbTab & Peso_total & vbTab & DataValidacao & vbTab & RespValidacao & vbTab & PartNumber & vbTab & TBNivel12!Obs
                ElseIf Formulario = "Engenharia/Estrutura" Then
                    arrNodes(Contador1).Text = TBNivel12!Desenho & vbTab & PosicaoTexto & vbTab & Codproduto & vbTab & CodRef & vbTab & TBNivel12!Descricao & vbTab & TBNivel12!Unidade & vbTab & TBNivel12!Versao_desenho & vbTab & Kg_un & vbTab & Un_Kg & vbTab & Dim_mm & vbTab & Kg_pc & vbTab & IIf(IsNull(TBAbrir!Largura), 0, Format(TBAbrir!Largura, "###,##0.00")) & vbTab & IIf(IsNull(TBAbrir!Comprimento), "", Format(TBAbrir!Comprimento, "###,##0.00")) & vbTab & QtTexto & vbTab & Peso_total & vbTab & Format(valor, "###,##0.00000000") & vbTab & DataValidacao & vbTab & RespValidacao & vbTab & TBNivel12!CODIGO & vbTab & PartNumber & vbTab & TBNivel12!Obs
                Else
                    arrNodes(Contador1).Text = TBNivel12!Desenho & vbTab & PosicaoTexto & vbTab & Codproduto & vbTab & CodRef & vbTab & TBNivel12!Descricao & vbTab & TBNivel12!Unidade & vbTab & TBNivel12!Versao_desenho & vbTab & Kg_un & vbTab & Un_Kg & vbTab & Dim_mm & vbTab & Kg_pc & vbTab & QtTexto & vbTab & Peso_total & vbTab & Format(valor, "###,##0.00000000") & vbTab & DataValidacao & vbTab & RespValidacao & vbTab & TBNivel12!CODIGO & vbTab & PartNumber & vbTab & TBNivel12!Obs
                End If
            End If
            If CarregarNivelAbaixo = True Then ProcNivel13Estrutura NomeFormulario, TBNivel12!Versao_desenho, MostrarValores, Carteira_pedidos, CarregarNivelAbaixo, Simulacao_vendas
        End If
        TBNivel12.MoveNext
    Loop
End If

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcNivel13Estrutura(NomeFormulario As Form, versao As String, MostrarValores As Boolean, Carteira_pedidos As Boolean, CarregarNivelAbaixo As Boolean, Simulacao_vendas As Boolean)
On Error GoTo tratar_erro

If versao = "" Then TextoFiltro = " " Else TextoFiltro = " and Versao = '" & versao & "'"
Set TBNivel13 = CreateObject("adodb.recordset")
TBNivel13.Open "Select * from projconjunto where codproduto = " & Codproduto & TextoFiltro & " order by Posicao, Codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBNivel13.EOF = False Then
    Do While TBNivel13.EOF = False
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select P.codproduto, P.largura, P.comprimento, P.Subtipoitem, P.Producao, P.PCusto, PCDV.DtValidacao, PCDV.RespValidacao from projproduto P LEFT JOIN Projconjunto_desc_versao PCDV ON PCDV.Codproduto = P.Codproduto " & IIf(versao <> "", "and PCDV.Versao = '" & versao & "'", "") & " where P.desenho = '" & TBNivel13!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Codproduto = TBAbrir!Codproduto
            If TBNivel13!Unidade = "KG" And TBNivel13!PesoTotal > 0 Or TBNivel13!Unidade = "MT" And TBNivel13!Dimensoes > 0 Or TBNivel13!Unidade = "MM" And TBNivel13!Dimensoes > 0 Then
                Select Case TBNivel13!Unidade
                    Case "KG":  Peso = TBNivel13!PesoTotal
                    Case "MT":  Peso = (TBNivel13!Dimensoes * TBNivel13!quantidade) / 1000
                    Case "MM":  Peso = TBNivel13!Dimensoes * TBNivel13!quantidade
                End Select
            Else
                Peso = TBNivel13!quantidade
            End If
            
            QuantsolicitadoN12 = Peso * QuantsolicitadoN11
            FunCarregaValorEstrutura TBNivel13!CODIGO, TBNivel13!Desenho, MostrarValores, Simulacao_vendas, QuantsolicitadoN12, False
            
            If Carteira_pedidos = True Then
                Processos = "N√O"
                Set TBProcessos = CreateObject("adodb.recordset")
                TBProcessos.Open "Select PR.IDProcesso FROM Processos PR INNER JOIN Projproduto P ON PR.Codproduto = P.Codproduto WHERE P.Desenho = '" & TBNivel13!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBProcessos.EOF = False Then
                    Processos = "SIM"
                End If
                OrdemCarteira = ""
                Set TBOrdem = CreateObject("adodb.recordset")
                TBOrdem.Open "Select PR.Ordem FROM Producao PR INNER JOIN Producao_pedidos PP ON PR.Ordem = PP.Ordem where PP.IDCarteira = " & Ordem & " and PR.Desenho = '" & TBNivel13!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBOrdem.EOF = False Then
                    OrdemCarteira = TBOrdem!Ordem
                End If
            End If
            
            PosicaoTexto = FunTamanhoTextoZeroEsq(TBNivel13!Posicao, 3)
            
            CodRef = ""
            Set TBItem = CreateObject("adodb.recordset")
            TBItem.Open "Select n_referencia from item_aplicacoes where codproduto = " & TBAbrir!Codproduto, Conexao, adOpenKeyset, adLockOptimistic
            If TBItem.EOF = False Then
                CodRef = TBItem!N_referencia
            End If
            TBItem.Close
            
            If TBAbrir!SubTipoItem <> 0 Then
                DataValidacao = IIf(IsNull(TBAbrir!DtValidacao), "", TBAbrir!DtValidacao)
                RespValidacao = IIf(IsNull(TBAbrir!RespValidacao), "", TBAbrir!RespValidacao)
            Else
                DataValidacao = ""
                RespValidacao = ""
            End If
            
            PartNumber = ""
            If IsNull(TBNivel13!ID_partnumber_fabricante) = False Then
                Set TBProduto = CreateObject("adodb.recordset")
                TBProduto.Open "Select Part_number from Projproduto_fabricante where ID = " & TBNivel13!ID_partnumber_fabricante, Conexao, adOpenKeyset, adLockOptimistic
                If TBProduto.EOF = False Then PartNumber = TBProduto!Part_number
                TBProduto.Close
            End If
            
            Contador1 = Contador1 + 1
            arrNodes(Contador1).Level = 12
            If Carteira_pedidos = True Or Simulacao_vendas = True Then
                QtTexto = Format(QuantsolicitadoN12, "###,##0.0000")
                If Carteira_pedidos = True Then
                    arrNodes(Contador1).Text = TBNivel13!Desenho & vbTab & PosicaoTexto & vbTab & Codproduto & vbTab & CodRef & vbTab & TBNivel13!Descricao & vbTab & TBNivel13!Unidade & vbTab & TBNivel13!Versao_desenho & vbTab & "" & vbTab & "" & vbTab & QtTexto & vbTab & "" & vbTab & "" & vbTab & Processos & vbTab & OrdemCarteira & vbTab & DataValidacao & vbTab & RespValidacao & vbTab & PartNumber & vbTab & TBNivel13!Obs
                Else
                    arrNodes(Contador1).Text = TBNivel13!Desenho & vbTab & PosicaoTexto & vbTab & Codproduto & vbTab & CodRef & vbTab & TBNivel13!Descricao & vbTab & TBNivel13!Unidade & vbTab & TBNivel13!Versao_desenho & vbTab & QtTexto & vbTab & Format(QuantComprado, "###,##0.0000") & vbTab & Format(IIf(QuantComprado - QuantsolicitadoN12 < 0, (QuantComprado - QuantsolicitadoN12) * -1, 0), "###,##0.0000") & vbTab & Format(Valor_Cofins_Serv, "###,##0.00000000") & vbTab & Format(valor, "###,##0.00000000") & vbTab & Format(Valor_Cofins_Prod, "###,##0.00000000") & vbTab & DataValidacao & vbTab & RespValidacao & vbTab & PartNumber & vbTab & TBNivel13!Obs & IIf(QuantComprado - QuantsolicitadoN12 < 0, 1, 0)
                End If
            Else
                Kg_un = Format(TBNivel13!PesoMetro, "###,##0.0000000000")
                Un_Kg = TBNivel13!Un_Kg
                Dim_mm = Format(TBNivel13!Dimensoes, "###,##0.0000000000")
                Kg_pc = Format(TBNivel13!Peso, "###,##0.0000000000")
                QtTexto = Format(TBNivel13!quantidade, "0.000")
                Peso_total = Format(TBNivel13!PesoTotal, "###,##0.0000000000")
                If Formulario = "Engenharia/Processos" Then
                    arrNodes(Contador1).Text = TBNivel13!Desenho & vbTab & PosicaoTexto & vbTab & Codproduto & vbTab & CodRef & vbTab & TBNivel13!Descricao & vbTab & TBNivel13!Unidade & vbTab & TBNivel13!Versao_desenho & vbTab & Kg_un & vbTab & Un_Kg & vbTab & Dim_mm & vbTab & Kg_pc & vbTab & QtTexto & vbTab & Peso_total & vbTab & DataValidacao & vbTab & RespValidacao & vbTab & PartNumber & vbTab & TBNivel13!Obs
                ElseIf Formulario = "Engenharia/Estrutura" Then
                    arrNodes(Contador1).Text = TBNivel13!Desenho & vbTab & PosicaoTexto & vbTab & Codproduto & vbTab & CodRef & vbTab & TBNivel13!Descricao & vbTab & TBNivel13!Unidade & vbTab & TBNivel13!Versao_desenho & vbTab & Kg_un & vbTab & Un_Kg & vbTab & Dim_mm & vbTab & Kg_pc & vbTab & IIf(IsNull(TBAbrir!Largura), 0, Format(TBAbrir!Largura, "###,##0.00")) & vbTab & IIf(IsNull(TBAbrir!Comprimento), "", Format(TBAbrir!Comprimento, "###,##0.00")) & vbTab & QtTexto & vbTab & Peso_total & vbTab & Format(valor, "###,##0.00000000") & vbTab & DataValidacao & vbTab & RespValidacao & vbTab & TBNivel13!CODIGO & vbTab & PartNumber & vbTab & TBNivel13!Obs
                Else
                    arrNodes(Contador1).Text = TBNivel13!Desenho & vbTab & PosicaoTexto & vbTab & Codproduto & vbTab & CodRef & vbTab & TBNivel13!Descricao & vbTab & TBNivel13!Unidade & vbTab & TBNivel13!Versao_desenho & vbTab & Kg_un & vbTab & Un_Kg & vbTab & Dim_mm & vbTab & Kg_pc & vbTab & QtTexto & vbTab & Peso_total & vbTab & Format(valor, "###,##0.00000000") & vbTab & DataValidacao & vbTab & RespValidacao & vbTab & TBNivel13!CODIGO & vbTab & PartNumber & vbTab & TBNivel13!Obs
                End If
            End If
            If CarregarNivelAbaixo = True Then ProcNivel14Estrutura NomeFormulario, TBNivel13!Versao_desenho, MostrarValores, Carteira_pedidos, CarregarNivelAbaixo, Simulacao_vendas
        End If
        TBNivel13.MoveNext
    Loop
End If

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcNivel14Estrutura(NomeFormulario As Form, versao As String, MostrarValores As Boolean, Carteira_pedidos As Boolean, CarregarNivelAbaixo As Boolean, Simulacao_vendas As Boolean)
On Error GoTo tratar_erro

If versao = "" Then TextoFiltro = " " Else TextoFiltro = " and Versao = '" & versao & "'"
Set TBNivel14 = CreateObject("adodb.recordset")
TBNivel14.Open "Select * from projconjunto where codproduto = " & Codproduto & TextoFiltro & " order by Posicao, Codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBNivel14.EOF = False Then
    Do While TBNivel14.EOF = False
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select P.codproduto, P.largura, P.comprimento, P.Subtipoitem, P.Producao, P.PCusto, PCDV.DtValidacao, PCDV.RespValidacao from projproduto P LEFT JOIN Projconjunto_desc_versao PCDV ON PCDV.Codproduto = P.Codproduto " & IIf(versao <> "", "and PCDV.Versao = '" & versao & "'", "") & " where P.desenho = '" & TBNivel14!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Codproduto = TBAbrir!Codproduto
            If TBNivel14!Unidade = "KG" And TBNivel14!PesoTotal > 0 Or TBNivel14!Unidade = "MT" And TBNivel14!Dimensoes > 0 Or TBNivel14!Unidade = "MM" And TBNivel14!Dimensoes > 0 Then
                Select Case TBNivel14!Unidade
                    Case "KG":  Peso = TBNivel14!PesoTotal
                    Case "MT":  Peso = (TBNivel14!Dimensoes * TBNivel14!quantidade) / 1000
                    Case "MM":  Peso = TBNivel14!Dimensoes * TBNivel14!quantidade
                End Select
            Else
                Peso = TBNivel14!quantidade
            End If
            
            QuantsolicitadoN13 = Peso * QuantsolicitadoN12
            FunCarregaValorEstrutura TBNivel14!CODIGO, TBNivel14!Desenho, MostrarValores, Simulacao_vendas, QuantsolicitadoN13, False
            
            If Carteira_pedidos = True Then
                Processos = "N√O"
                Set TBProcessos = CreateObject("adodb.recordset")
                TBProcessos.Open "Select PR.IDProcesso FROM Processos PR INNER JOIN Projproduto P ON PR.Codproduto = P.Codproduto WHERE P.Desenho = '" & TBNivel14!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBProcessos.EOF = False Then
                    Processos = "SIM"
                End If
                OrdemCarteira = ""
                Set TBOrdem = CreateObject("adodb.recordset")
                TBOrdem.Open "Select PR.Ordem FROM Producao PR INNER JOIN Producao_pedidos PP ON PR.Ordem = PP.Ordem where PP.IDCarteira = " & Ordem & " and PR.Desenho = '" & TBNivel14!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBOrdem.EOF = False Then
                    OrdemCarteira = TBOrdem!Ordem
                End If
            End If
            
            PosicaoTexto = FunTamanhoTextoZeroEsq(TBNivel14!Posicao, 3)
            
            CodRef = ""
            Set TBItem = CreateObject("adodb.recordset")
            TBItem.Open "Select n_referencia from item_aplicacoes where codproduto = " & TBAbrir!Codproduto, Conexao, adOpenKeyset, adLockOptimistic
            If TBItem.EOF = False Then
                CodRef = TBItem!N_referencia
            End If
            TBItem.Close
            
            If TBAbrir!SubTipoItem <> 0 Then
                DataValidacao = IIf(IsNull(TBAbrir!DtValidacao), "", TBAbrir!DtValidacao)
                RespValidacao = IIf(IsNull(TBAbrir!RespValidacao), "", TBAbrir!RespValidacao)
            Else
                DataValidacao = ""
                RespValidacao = ""
            End If
            
            PartNumber = ""
            If IsNull(TBNivel14!ID_partnumber_fabricante) = False Then
                Set TBProduto = CreateObject("adodb.recordset")
                TBProduto.Open "Select Part_number from Projproduto_fabricante where ID = " & TBNivel14!ID_partnumber_fabricante, Conexao, adOpenKeyset, adLockOptimistic
                If TBProduto.EOF = False Then PartNumber = TBProduto!Part_number
                TBProduto.Close
            End If
            
            Contador1 = Contador1 + 1
            arrNodes(Contador1).Level = 13
            If Carteira_pedidos = True Or Simulacao_vendas = True Then
                QtTexto = Format(QuantsolicitadoN13, "###,##0.0000")
                If Carteira_pedidos = True Then
                    arrNodes(Contador1).Text = TBNivel14!Desenho & vbTab & PosicaoTexto & vbTab & Codproduto & vbTab & CodRef & vbTab & TBNivel14!Descricao & vbTab & TBNivel14!Unidade & vbTab & TBNivel14!Versao_desenho & vbTab & "" & vbTab & "" & vbTab & QtTexto & vbTab & "" & vbTab & "" & vbTab & Processos & vbTab & OrdemCarteira & vbTab & DataValidacao & vbTab & RespValidacao & vbTab & PartNumber & vbTab & TBNivel14!Obs
                Else
                    arrNodes(Contador1).Text = TBNivel14!Desenho & vbTab & PosicaoTexto & vbTab & Codproduto & vbTab & CodRef & vbTab & TBNivel14!Descricao & vbTab & TBNivel14!Unidade & vbTab & TBNivel14!Versao_desenho & vbTab & QtTexto & vbTab & Format(QuantComprado, "###,##0.0000") & vbTab & Format(IIf(QuantComprado - QuantsolicitadoN13 < 0, (QuantComprado - QuantsolicitadoN13) * -1, 0), "###,##0.0000") & vbTab & Format(Valor_Cofins_Serv, "###,##0.00000000") & vbTab & Format(valor, "###,##0.00000000") & vbTab & Format(Valor_Cofins_Prod, "###,##0.00000000") & vbTab & DataValidacao & vbTab & RespValidacao & vbTab & PartNumber & vbTab & TBNivel14!Obs & IIf(QuantComprado - QuantsolicitadoN13 < 0, 1, 0)
                End If
            Else
                Kg_un = Format(TBNivel14!PesoMetro, "###,##0.0000000000")
                Un_Kg = TBNivel14!Un_Kg
                Dim_mm = Format(TBNivel14!Dimensoes, "###,##0.0000000000")
                Kg_pc = Format(TBNivel14!Peso, "###,##0.0000000000")
                QtTexto = Format(TBNivel14!quantidade, "0.000")
                Peso_total = Format(TBNivel14!PesoTotal, "###,##0.0000000000")
                If Formulario = "Engenharia/Processos" Then
                    arrNodes(Contador1).Text = TBNivel14!Desenho & vbTab & PosicaoTexto & vbTab & Codproduto & vbTab & CodRef & vbTab & TBNivel14!Descricao & vbTab & TBNivel14!Unidade & vbTab & TBNivel14!Versao_desenho & vbTab & Kg_un & vbTab & Un_Kg & vbTab & Dim_mm & vbTab & Kg_pc & vbTab & QtTexto & vbTab & Peso_total & vbTab & DataValidacao & vbTab & RespValidacao & vbTab & PartNumber & vbTab & TBNivel14!Obs
                ElseIf Formulario = "Engenharia/Estrutura" Then
                    arrNodes(Contador1).Text = TBNivel14!Desenho & vbTab & PosicaoTexto & vbTab & Codproduto & vbTab & CodRef & vbTab & TBNivel14!Descricao & vbTab & TBNivel14!Unidade & vbTab & TBNivel14!Versao_desenho & vbTab & Kg_un & vbTab & Un_Kg & vbTab & Dim_mm & vbTab & Kg_pc & vbTab & IIf(IsNull(TBAbrir!Largura), 0, Format(TBAbrir!Largura, "###,##0.00")) & vbTab & IIf(IsNull(TBAbrir!Comprimento), "", Format(TBAbrir!Comprimento, "###,##0.00")) & vbTab & QtTexto & vbTab & Peso_total & vbTab & Format(valor, "###,##0.00000000") & vbTab & DataValidacao & vbTab & RespValidacao & vbTab & TBNivel14!CODIGO & vbTab & PartNumber & vbTab & TBNivel14!Obs
                Else
                    arrNodes(Contador1).Text = TBNivel14!Desenho & vbTab & PosicaoTexto & vbTab & Codproduto & vbTab & CodRef & vbTab & TBNivel14!Descricao & vbTab & TBNivel14!Unidade & vbTab & TBNivel14!Versao_desenho & vbTab & Kg_un & vbTab & Un_Kg & vbTab & Dim_mm & vbTab & Kg_pc & vbTab & QtTexto & vbTab & Peso_total & vbTab & Format(valor, "###,##0.00000000") & vbTab & DataValidacao & vbTab & RespValidacao & vbTab & TBNivel14!CODIGO & vbTab & PartNumber & vbTab & TBNivel14!Obs
                End If
            End If
            If CarregarNivelAbaixo = True Then ProcNivel15Estrutura NomeFormulario, TBNivel14!Versao_desenho, MostrarValores, Carteira_pedidos, CarregarNivelAbaixo, Simulacao_vendas
        End If
        TBNivel14.MoveNext
    Loop
End If

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcNivel15Estrutura(NomeFormulario As Form, versao As String, MostrarValores As Boolean, Carteira_pedidos As Boolean, CarregarNivelAbaixo As Boolean, Simulacao_vendas As Boolean)
On Error GoTo tratar_erro

If versao = "" Then TextoFiltro = " " Else TextoFiltro = " and Versao = '" & versao & "'"
Set TBNivel15 = CreateObject("adodb.recordset")
TBNivel15.Open "Select * from projconjunto where codproduto = " & Codproduto & TextoFiltro & " order by Posicao, Codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBNivel15.EOF = False Then
    Do While TBNivel15.EOF = False
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select P.codproduto, P.largura, P.comprimento, P.Subtipoitem, P.Producao, P.PCusto, PCDV.DtValidacao, PCDV.RespValidacao from projproduto P LEFT JOIN Projconjunto_desc_versao PCDV ON PCDV.Codproduto = P.Codproduto " & IIf(versao <> "", "and PCDV.Versao = '" & versao & "'", "") & " where P.desenho = '" & TBNivel15!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Codproduto = TBAbrir!Codproduto
            If TBNivel15!Unidade = "KG" And TBNivel15!PesoTotal > 0 Or TBNivel15!Unidade = "MT" And TBNivel15!Dimensoes > 0 Or TBNivel15!Unidade = "MM" And TBNivel15!Dimensoes > 0 Then
                Select Case TBNivel15!Unidade
                    Case "KG":  Peso = TBNivel15!PesoTotal
                    Case "MT":  Peso = (TBNivel15!Dimensoes * TBNivel15!quantidade) / 1000
                    Case "MM":  Peso = TBNivel15!Dimensoes * TBNivel15!quantidade
                End Select
            Else
                Peso = TBNivel15!quantidade
            End If
            
            QuantsolicitadoN14 = Peso * QuantsolicitadoN13
            FunCarregaValorEstrutura TBNivel15!CODIGO, TBNivel15!Desenho, MostrarValores, Simulacao_vendas, QuantsolicitadoN14, False
            
            If Carteira_pedidos = True Then
                Processos = "N√O"
                Set TBProcessos = CreateObject("adodb.recordset")
                TBProcessos.Open "Select PR.IDProcesso FROM Processos PR INNER JOIN Projproduto P ON PR.Codproduto = P.Codproduto WHERE P.Desenho = '" & TBNivel15!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBProcessos.EOF = False Then
                    Processos = "SIM"
                End If
                OrdemCarteira = ""
                Set TBOrdem = CreateObject("adodb.recordset")
                TBOrdem.Open "Select PR.Ordem FROM Producao PR INNER JOIN Producao_pedidos PP ON PR.Ordem = PP.Ordem where PP.IDCarteira = " & Ordem & " and PR.Desenho = '" & TBNivel15!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBOrdem.EOF = False Then
                    OrdemCarteira = TBOrdem!Ordem
                End If
            End If
            
            PosicaoTexto = FunTamanhoTextoZeroEsq(TBNivel15!Posicao, 3)
            
            CodRef = ""
            Set TBItem = CreateObject("adodb.recordset")
            TBItem.Open "Select n_referencia from item_aplicacoes where codproduto = " & TBAbrir!Codproduto, Conexao, adOpenKeyset, adLockOptimistic
            If TBItem.EOF = False Then
                CodRef = TBItem!N_referencia
            End If
            TBItem.Close
            
            If TBAbrir!SubTipoItem <> 0 Then
                DataValidacao = IIf(IsNull(TBAbrir!DtValidacao), "", TBAbrir!DtValidacao)
                RespValidacao = IIf(IsNull(TBAbrir!RespValidacao), "", TBAbrir!RespValidacao)
            Else
                DataValidacao = ""
                RespValidacao = ""
            End If
            
            PartNumber = ""
            If IsNull(TBNivel15!ID_partnumber_fabricante) = False Then
                Set TBProduto = CreateObject("adodb.recordset")
                TBProduto.Open "Select Part_number from Projproduto_fabricante where ID = " & TBNivel15!ID_partnumber_fabricante, Conexao, adOpenKeyset, adLockOptimistic
                If TBProduto.EOF = False Then PartNumber = TBProduto!Part_number
                TBProduto.Close
            End If
            
            Contador1 = Contador1 + 1
            arrNodes(Contador1).Level = 14
            If Carteira_pedidos = True Or Simulacao_vendas = True Then
                QtTexto = Format(QuantsolicitadoN14, "###,##0.0000")
                If Carteira_pedidos = True Then
                    arrNodes(Contador1).Text = TBNivel15!Desenho & vbTab & PosicaoTexto & vbTab & Codproduto & vbTab & CodRef & vbTab & TBNivel15!Descricao & vbTab & TBNivel15!Unidade & vbTab & TBNivel15!Versao_desenho & vbTab & "" & vbTab & "" & vbTab & QtTexto & vbTab & "" & vbTab & "" & vbTab & Processos & vbTab & OrdemCarteira & vbTab & DataValidacao & vbTab & RespValidacao & vbTab & PartNumber & vbTab & TBNivel15!Obs
                Else
                    arrNodes(Contador1).Text = TBNivel15!Desenho & vbTab & PosicaoTexto & vbTab & Codproduto & vbTab & CodRef & vbTab & TBNivel15!Descricao & vbTab & TBNivel15!Unidade & vbTab & TBNivel15!Versao_desenho & vbTab & QtTexto & vbTab & Format(QuantComprado, "###,##0.0000") & vbTab & Format(IIf(QuantComprado - QuantsolicitadoN14 < 0, (QuantComprado - QuantsolicitadoN14) * -1, 0), "###,##0.0000") & vbTab & Format(Valor_Cofins_Serv, "###,##0.00000000") & vbTab & Format(valor, "###,##0.00000000") & vbTab & Format(Valor_Cofins_Prod, "###,##0.00000000") & vbTab & DataValidacao & vbTab & RespValidacao & vbTab & PartNumber & vbTab & TBNivel15!Obs & IIf(QuantComprado - QuantsolicitadoN14 < 0, 1, 0)
                End If
            Else
                Kg_un = Format(TBNivel15!PesoMetro, "###,##0.0000000000")
                Un_Kg = TBNivel15!Un_Kg
                Dim_mm = Format(TBNivel15!Dimensoes, "###,##0.0000000000")
                Kg_pc = Format(TBNivel15!Peso, "###,##0.0000000000")
                QtTexto = Format(TBNivel15!quantidade, "0.000")
                Peso_total = Format(TBNivel15!PesoTotal, "###,##0.0000000000")
                If Formulario = "Engenharia/Processos" Then
                    arrNodes(Contador1).Text = TBNivel15!Desenho & vbTab & PosicaoTexto & vbTab & Codproduto & vbTab & CodRef & vbTab & TBNivel15!Descricao & vbTab & TBNivel15!Unidade & vbTab & TBNivel15!Versao_desenho & vbTab & Kg_un & vbTab & Un_Kg & vbTab & Dim_mm & vbTab & Kg_pc & vbTab & QtTexto & vbTab & Peso_total & vbTab & DataValidacao & vbTab & RespValidacao & vbTab & PartNumber & vbTab & TBNivel15!Obs
                ElseIf Formulario = "Engenharia/Estrutura" Then
                    arrNodes(Contador1).Text = TBNivel15!Desenho & vbTab & PosicaoTexto & vbTab & Codproduto & vbTab & CodRef & vbTab & TBNivel15!Descricao & vbTab & TBNivel15!Unidade & vbTab & TBNivel15!Versao_desenho & vbTab & Kg_un & vbTab & Un_Kg & vbTab & Dim_mm & vbTab & Kg_pc & vbTab & IIf(IsNull(TBAbrir!Largura), 0, Format(TBAbrir!Largura, "###,##0.00")) & vbTab & IIf(IsNull(TBAbrir!Comprimento), "", Format(TBAbrir!Comprimento, "###,##0.00")) & vbTab & QtTexto & vbTab & Peso_total & vbTab & Format(valor, "###,##0.00000000") & vbTab & DataValidacao & vbTab & RespValidacao & vbTab & TBNivel15!CODIGO & vbTab & PartNumber & vbTab & TBNivel15!Obs
                Else
                    arrNodes(Contador1).Text = TBNivel15!Desenho & vbTab & PosicaoTexto & vbTab & Codproduto & vbTab & CodRef & vbTab & TBNivel15!Descricao & vbTab & TBNivel15!Unidade & vbTab & TBNivel15!Versao_desenho & vbTab & Kg_un & vbTab & Un_Kg & vbTab & Dim_mm & vbTab & Kg_pc & vbTab & QtTexto & vbTab & Peso_total & vbTab & Format(valor, "###,##0.00000000") & vbTab & DataValidacao & vbTab & RespValidacao & vbTab & TBNivel15!CODIGO & vbTab & PartNumber & vbTab & TBNivel15!Obs
                End If
            End If
        End If
        TBNivel15.MoveNext
    Loop
End If

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Function FunVerifUNConversao(Un_est As String, Un_com As String) As Boolean
On Error GoTo tratar_erro

FunVerifUNConversao = False
If Un_est <> Un_com And (Un_est = "KG" Or Un_est = "MT" Or Un_est = "MM" Or Un_est = "BR" Or Un_est = "PC" Or Un_est = "P«" Or Un_est = "CH") And (Un_com = "KG" Or Un_com = "MT" Or Un_com = "MM" Or Un_com = "BR" Or Un_com = "PC" Or Un_com = "P«" Or Un_com = "CH") Then FunVerifUNConversao = True

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Function

Function FunConverteUN(Un_est As String, Un_com As String, quantidadeUN As Double, DesenhoUN As String) As Double
On Error GoTo tratar_erro
Dim quantidadeUN2 As Double

Set TBCFOP = CreateObject("adodb.recordset")
TBCFOP.Open "Select peso_metro, PBruto from projproduto where desenho = '" & DesenhoUN & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBCFOP.EOF = False Then
    If Un_est = "KG" Then
        If Un_com = "P«" Or Un_com = "PC" Or Un_com = "BR" Or Un_com = "CH" Then quantidadeUN2 = IIf(IsNull(TBCFOP!PBruto), 0, TBCFOP!PBruto) Else quantidadeUN2 = IIf(IsNull(TBCFOP!peso_metro), 0, TBCFOP!peso_metro)
        If Un_com = "MT" Then
            FunConverteUN = Format(quantidadeUN2 * quantidadeUN, "###,##0.0000000000")
        ElseIf Un_com = "MM" Then
                FunConverteUN = Format((quantidadeUN2 / 1000) * quantidadeUN, "###,##0.0000000000")
            ElseIf Un_com = "P«" Or Un_com = "PC" Or Un_com = "BR" Or Un_com = "CH" Then
                FunConverteUN = Format(quantidadeUN2 * quantidadeUN, "###,##0.0000000000")
        End If
    Else
        If Un_est = "P«" Or Un_est = "PC" Or Un_est = "BR" Or Un_est = "CH" Then quantidadeUN2 = IIf(IsNull(TBCFOP!PBruto), 0, TBCFOP!PBruto) Else quantidadeUN2 = IIf(IsNull(TBCFOP!peso_metro), 0, TBCFOP!peso_metro)
        If quantidadeUN2 > 0 Then
            If Un_est = "MT" Then
                FunConverteUN = Format(quantidadeUN / quantidadeUN2, "###,##0.0000000000")
            ElseIf Un_est = "MM" Then
                    FunConverteUN = Format((quantidadeUN * 0.001) / quantidadeUN2, "###,##0.0000000000")
                ElseIf Un_est = "P«" Or Un_est = "PC" Or Un_est = "BR" Or Un_est = "CH" Then
                    FunConverteUN = Format(quantidadeUN / quantidadeUN2, "###,##0.0000000000")
            End If
        End If
    End If
End If
TBCFOP.Close

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Public Sub ProcLibera_UN_Com(Un_est As ComboBox, Un_com As ComboBox)
On Error GoTo tratar_erro

If Un_est = "KG" Or Un_est = "MM" Or Un_est = "MT" Or Un_est = "PC" Or Un_est = "P«" Or Un_est = "BR" Or Un_est = "CH" Then
    With Un_com
        .Clear
        .Locked = False
        .TabStop = True
        If Un_est = "KG" Then
            TextoFiltro = "and (Unidade = 'KG' or Unidade = 'MM' or Unidade = 'MT' or Unidade = 'PC' or Unidade = 'P«' or Unidade = 'BR' or Unidade = 'CH' )"
        Else
            TextoFiltro = "and (Unidade = 'KG' or Unidade = '" & Un_est & "')"
        End If
        Set TBCarregarCombo = CreateObject("adodb.recordset")
        TBCarregarCombo.Open "Select Unidade, Codigo from Unidade_Medida where Unidade IS NOT NULL " & TextoFiltro & " group by Unidade, Codigo", Conexao, adOpenKeyset, adLockOptimistic
        If TBCarregarCombo.EOF = False Then
            Do While TBCarregarCombo.EOF = False
                .AddItem TBCarregarCombo!Unidade
                .ItemData(.NewIndex) = TBCarregarCombo!CODIGO
                TBCarregarCombo.MoveNext
            Loop
        End If
        Set TBCarregarCombo = CreateObject("adodb.recordset")
        TBCarregarCombo.Open "Select U.Unidade, U.Codigo from Unidade_Medida U INNER JOIN Tabela_conversao_unidade TCU ON TCU.Unidade_para = U.Unidade where TCU.Unidade_de = '" & Un_est & "' group by U.Unidade, U.Codigo", Conexao, adOpenKeyset, adLockOptimistic
        If TBCarregarCombo.EOF = False Then
            Do While TBCarregarCombo.EOF = False
                .AddItem TBCarregarCombo!Unidade
                .ItemData(.NewIndex) = TBCarregarCombo!CODIGO
                TBCarregarCombo.MoveNext
            Loop
        End If
        TBCarregarCombo.Close
    End With
Else
    ProcCarregaComboUnidade Un_com, False
End If

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Function FunVefificaModuloLocacao(Administrativo As Boolean, Manufatura As Boolean, Faturamento As Boolean) As Boolean
On Error GoTo tratar_erro

FunVefificaModuloLocacao = False
If TemInternet = True And ErroDriverMYSQL = False Then
    Set TBFIltro = CreateObject("adodb.recordset")
    TBFIltro.Open "Select * from Empresa", Conexao, adOpenKeyset, adLockOptimistic
    If TBFIltro.EOF = False Then
        Do While TBFIltro.EOF = False
            FunAbreBDSite
            If ConexaoMySql.State = 1 Then
                Set TBMySQL = New ADODB.Recordset
                TBMySQL.Open "Select Modulo From Clientes Where CNPJ = '" & TBFIltro!CNPJ & "' and Modulo = 'Full (ERP/MRP)'", ConexaoMySql, adOpenKeyset, adLockOptimistic, adCmdText
                If TBMySQL.EOF = False Then
                    FunVefificaModuloLocacao = True
                Else
                    If Administrativo = True And Faturamento = True Then
                        TextoFiltro = "(Modulo = 'Light I (Administrativo)' or Modulo = 'Light III (Faturamento)')"
                    ElseIf Administrativo = True Then
                            TextoFiltro = "Modulo = 'Light I (Administrativo)'"
                        ElseIf Faturamento = True Then
                                TextoFiltro = "Modulo = 'Light III (Faturamento)'"
                            Else
                                TextoFiltro = "Modulo = 'Light II (Manufatura)'"
                    End If
                    Set TBMySQL = New ADODB.Recordset
                    TBMySQL.Open "Select Modulo From Clientes Where CNPJ = '" & TBFIltro!CNPJ & "' and " & TextoFiltro, ConexaoMySql, adOpenKeyset, adLockOptimistic, adCmdText
                    If TBMySQL.EOF = False Then
                        FunVefificaModuloLocacao = True
                    End If
                End If
            Else
                GoTo VerifBDLocal
            End If
            TBFIltro.MoveNext
        Loop
    End If
    TBFIltro.Close
Else
VerifBDLocal:
    Set TBFIltro = CreateObject("adodb.recordset")
    TBFIltro.Open "Select * from Empresa where Modulo = 'Full (ERP/MRP)'", Conexao, adOpenKeyset, adLockOptimistic
    If TBFIltro.EOF = False Then
        FunVefificaModuloLocacao = True
    Else
        If Administrativo = True And Faturamento = True Then
            TextoFiltro = "(Modulo = 'Light I (Administrativo)' or Modulo = 'Light III (Faturamento)')"
        ElseIf Administrativo = True Then
                TextoFiltro = "Modulo = 'Light I (Administrativo)'"
            ElseIf Faturamento = True Then
                    TextoFiltro = "Modulo = 'Light III (Faturamento)'"
                Else
                    TextoFiltro = "Modulo = 'Light II (Manufatura)'"
        End If
        Set TBFIltro = CreateObject("adodb.recordset")
        TBFIltro.Open "Select * from Empresa where " & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
        If TBFIltro.EOF = False Then
            FunVefificaModuloLocacao = True
        End If
    End If
End If
If FunVefificaModuloLocacao = False Then Call USMsgBox("Prezado Cliente, " & vbCrLf & "O mÛdulo contratado " & Modulo_caprind & " n„o d· direito de acesso a este recurso.", vbExclamation, "CAPRIND v5.0", "ValidaÁ„o de contrato")

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Function FunVerifRegimeEmpresa(ID_empresa As Integer) As Integer
On Error GoTo tratar_erro

Set TBAliquota = CreateObject("adodb.recordset")
TBAliquota.Open "Select Simples, Presumido, Real, Simples1 from Empresa where Codigo = " & ID_empresa, Conexao, adOpenKeyset, adLockOptimistic
If TBAliquota.EOF = False Then
    If TBAliquota!Simples = True Then FunVerifRegimeEmpresa = 1
    If TBAliquota!Presumido = True Then FunVerifRegimeEmpresa = 2
    If TBAliquota!Real = True Then FunVerifRegimeEmpresa = 3
    If TBAliquota!Simples1 = True Then FunVerifRegimeEmpresa = 4
End If
TBAliquota.Close

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Sub ProcEmpenhoPIN1(Desenho As String, IDcarteira As Long, Versao_estrutura As String, QuantSolicitado1 As Double)
On Error GoTo tratar_erro

Set TBItem = CreateObject("adodb.recordset")
TBItem.Open "Select * from projproduto where desenho = '" & Desenho & "' and DtValidacaoConj IS NOT NULL", Conexao, adOpenKeyset, adLockOptimistic
If TBItem.EOF = False Then
    Set TBNivel1 = CreateObject("adodb.recordset")
    TBNivel1.Open "Select PC.*, P.PBruto, P.SubTipoItem, P.Desenho from projconjunto PC INNER JOIN projproduto P ON P.Desenho = PC.Desenho where PC.codproduto = " & TBItem!Codproduto & " and PC.Versao = '" & Versao_estrutura & "' and P.bloqueado = 'False' order by PC.Posicao, PC.codigo", Conexao, adOpenKeyset, adLockOptimistic
    If TBNivel1.EOF = False Then
        Do While TBNivel1.EOF = False
            Set TBMaterial = CreateObject("adodb.recordset")
            TBMaterial.Open "Select * from producaomaterial where ID_carteira = " & IDcarteira & " and codigo = '" & TBNivel1!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBMaterial.EOF = True Then TBMaterial.AddNew
            QuantSolicitado = QuantSolicitado1
            ProcEnviaDadosEmpenhoPI IDcarteira, TBNivel1!quantidade, TBNivel1!Unidade, TBNivel1!Desenho, TBNivel1!Descricao, IIf(IsNull(TBNivel1!PesoMetro), 0, TBNivel1!PesoMetro), TBNivel1!Peso, TBNivel1!PesoTotal, TBNivel1!Dimensoes, TBNivel1!Un_Kg, TBNivel1!versao, TBNivel1!SubTipoItem, IIf(IsNull(TBNivel1!PBruto), 0, TBNivel1!PBruto), IIf(IsNull(TBNivel1!Percentual_perda), 0, TBNivel1!Percentual_perda)
            TBMaterial.Update
            
            QuantsolicitadoN1 = QuantSolicitado
            ProcEmpenhoPIN2 TBNivel1!Desenho, IDcarteira, Versao_estrutura, TBMaterial!Requisitado
            
            TBNivel1.MoveNext
        Loop
    End If
    TBNivel1.Close
End If
QuantSolicitado = Quant

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEmpenhoPIN2(Desenho As String, IDcarteira As Long, Versao_estrutura As String, QuantSolicitado1 As Double)
On Error GoTo tratar_erro

Set TBItem = CreateObject("adodb.recordset")
TBItem.Open "Select * from projproduto where desenho = '" & Desenho & "' and DtValidacaoConj IS NOT NULL", Conexao, adOpenKeyset, adLockOptimistic
If TBItem.EOF = False Then
    Set TBNivel2 = CreateObject("adodb.recordset")
    TBNivel2.Open "Select PC.*, P.PBruto, P.SubTipoItem, P.Desenho from projconjunto PC INNER JOIN projproduto P ON P.Desenho = PC.Desenho where PC.codproduto = " & TBItem!Codproduto & " and PC.Versao = '" & Versao_estrutura & "' and P.bloqueado = 'False' order by PC.Posicao, PC.codigo", Conexao, adOpenKeyset, adLockOptimistic
    If TBNivel2.EOF = False Then
        Do While TBNivel2.EOF = False
            Set TBMaterial = CreateObject("adodb.recordset")
            TBMaterial.Open "Select * from producaomaterial where ID_carteira = " & IDcarteira & " and codigo = '" & TBNivel2!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBMaterial.EOF = True Then TBMaterial.AddNew
            QuantSolicitado = QuantSolicitado1
            ProcEnviaDadosEmpenhoPI IDcarteira, TBNivel2!quantidade, TBNivel2!Unidade, TBNivel2!Desenho, TBNivel2!Descricao, IIf(IsNull(TBNivel2!PesoMetro), 0, TBNivel2!PesoMetro), TBNivel2!Peso, TBNivel2!PesoTotal, TBNivel2!Dimensoes, TBNivel2!Un_Kg, TBNivel2!versao, TBNivel2!SubTipoItem, IIf(IsNull(TBNivel2!PBruto), 0, TBNivel2!PBruto), IIf(IsNull(TBNivel2!Percentual_perda), 0, TBNivel2!Percentual_perda)
            TBMaterial.Update
            
            QuantsolicitadoN2 = QuantSolicitado
            ProcEmpenhoPIN3 TBNivel2!Desenho, IDcarteira, Versao_estrutura, TBMaterial!Requisitado
            
            TBNivel2.MoveNext
        Loop
    End If
    TBNivel2.Close
End If
QuantSolicitado = QuantsolicitadoN1

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEmpenhoPIN3(Desenho As String, IDcarteira As Long, Versao_estrutura As String, QuantSolicitado1 As Double)
On Error GoTo tratar_erro

Set TBItem = CreateObject("adodb.recordset")
TBItem.Open "Select * from projproduto where desenho = '" & Desenho & "' and DtValidacaoConj IS NOT NULL", Conexao, adOpenKeyset, adLockOptimistic
If TBItem.EOF = False Then
    Set TBNivel3 = CreateObject("adodb.recordset")
    TBNivel3.Open "Select PC.*, P.PBruto, P.SubTipoItem, P.Desenho from projconjunto PC INNER JOIN projproduto P ON P.Desenho = PC.Desenho where PC.codproduto = " & TBItem!Codproduto & " and PC.Versao = '" & Versao_estrutura & "' and P.bloqueado = 'False' order by PC.Posicao, PC.codigo", Conexao, adOpenKeyset, adLockOptimistic
    If TBNivel3.EOF = False Then
        Do While TBNivel3.EOF = False
            Set TBMaterial = CreateObject("adodb.recordset")
            TBMaterial.Open "Select * from producaomaterial where ID_carteira = " & IDcarteira & " and codigo = '" & TBNivel3!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBMaterial.EOF = True Then TBMaterial.AddNew
            QuantSolicitado = QuantSolicitado1
            ProcEnviaDadosEmpenhoPI IDcarteira, TBNivel3!quantidade, TBNivel3!Unidade, TBNivel3!Desenho, TBNivel3!Descricao, IIf(IsNull(TBNivel3!PesoMetro), 0, TBNivel3!PesoMetro), TBNivel3!Peso, TBNivel3!PesoTotal, TBNivel3!Dimensoes, TBNivel3!Un_Kg, TBNivel3!versao, TBNivel3!SubTipoItem, IIf(IsNull(TBNivel3!PBruto), 0, TBNivel3!PBruto), IIf(IsNull(TBNivel3!Percentual_perda), 0, TBNivel3!Percentual_perda)
            TBMaterial.Update
            
            QuantsolicitadoN3 = QuantSolicitado
            ProcEmpenhoPIN4 TBNivel3!Desenho, IDcarteira, Versao_estrutura, TBMaterial!Requisitado
            
            TBNivel3.MoveNext
        Loop
    End If
    TBNivel3.Close
End If
QuantSolicitado = QuantsolicitadoN2

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEmpenhoPIN4(Desenho As String, IDcarteira As Long, Versao_estrutura As String, QuantSolicitado1 As Double)
On Error GoTo tratar_erro

Set TBItem = CreateObject("adodb.recordset")
TBItem.Open "Select * from projproduto where desenho = '" & Desenho & "' and DtValidacaoConj IS NOT NULL", Conexao, adOpenKeyset, adLockOptimistic
If TBItem.EOF = False Then
    Set TBNivel4 = CreateObject("adodb.recordset")
    TBNivel4.Open "Select PC.*, P.PBruto, P.SubTipoItem, P.Desenho from projconjunto PC INNER JOIN projproduto P ON P.Desenho = PC.Desenho where PC.codproduto = " & TBItem!Codproduto & " and PC.Versao = '" & Versao_estrutura & "' and P.bloqueado = 'False' order by PC.Posicao, PC.codigo", Conexao, adOpenKeyset, adLockOptimistic
    If TBNivel4.EOF = False Then
        Do While TBNivel4.EOF = False
            Set TBMaterial = CreateObject("adodb.recordset")
            TBMaterial.Open "Select * from producaomaterial where ID_carteira = " & IDcarteira & " and codigo = '" & TBNivel4!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBMaterial.EOF = True Then TBMaterial.AddNew
            QuantSolicitado = QuantSolicitado1
            ProcEnviaDadosEmpenhoPI IDcarteira, TBNivel4!quantidade, TBNivel4!Unidade, TBNivel4!Desenho, TBNivel4!Descricao, IIf(IsNull(TBNivel4!PesoMetro), 0, TBNivel4!PesoMetro), TBNivel4!Peso, TBNivel4!PesoTotal, TBNivel4!Dimensoes, TBNivel4!Un_Kg, TBNivel4!versao, TBNivel4!SubTipoItem, IIf(IsNull(TBNivel4!PBruto), 0, TBNivel4!PBruto), IIf(IsNull(TBNivel4!Percentual_perda), 0, TBNivel4!Percentual_perda)
            TBMaterial.Update
            
            QuantsolicitadoN4 = QuantSolicitado
            ProcEmpenhoPIN5 TBNivel4!Desenho, IDcarteira, Versao_estrutura, TBMaterial!Requisitado
            
            TBNivel4.MoveNext
        Loop
    End If
    TBNivel4.Close
End If
QuantSolicitado = QuantsolicitadoN3

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEmpenhoPIN5(Desenho As String, IDcarteira As Long, Versao_estrutura As String, QuantSolicitado1 As Double)
On Error GoTo tratar_erro

Set TBItem = CreateObject("adodb.recordset")
TBItem.Open "Select * from projproduto where desenho = '" & Desenho & "' and DtValidacaoConj IS NOT NULL", Conexao, adOpenKeyset, adLockOptimistic
If TBItem.EOF = False Then
    Set TBNivel5 = CreateObject("adodb.recordset")
    TBNivel5.Open "Select PC.*, P.PBruto, P.SubTipoItem, P.Desenho from projconjunto PC INNER JOIN projproduto P ON P.Desenho = PC.Desenho where PC.codproduto = " & TBItem!Codproduto & " and PC.Versao = '" & Versao_estrutura & "' and P.bloqueado = 'False' order by PC.Posicao, PC.codigo", Conexao, adOpenKeyset, adLockOptimistic
    If TBNivel5.EOF = False Then
        Do While TBNivel5.EOF = False
            Set TBMaterial = CreateObject("adodb.recordset")
            TBMaterial.Open "Select * from producaomaterial where ID_carteira = " & IDcarteira & " and codigo = '" & TBNivel5!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBMaterial.EOF = True Then TBMaterial.AddNew
            QuantSolicitado = QuantSolicitado1
            ProcEnviaDadosEmpenhoPI IDcarteira, TBNivel5!quantidade, TBNivel5!Unidade, TBNivel5!Desenho, TBNivel5!Descricao, IIf(IsNull(TBNivel5!PesoMetro), 0, TBNivel5!PesoMetro), TBNivel5!Peso, TBNivel5!PesoTotal, TBNivel5!Dimensoes, TBNivel5!Un_Kg, TBNivel5!versao, TBNivel5!SubTipoItem, IIf(IsNull(TBNivel5!PBruto), 0, TBNivel5!PBruto), IIf(IsNull(TBNivel5!Percentual_perda), 0, TBNivel5!Percentual_perda)
            TBMaterial.Update
            
            QuantsolicitadoN5 = QuantSolicitado
            ProcEmpenhoPIN6 TBNivel5!Desenho, IDcarteira, Versao_estrutura, TBMaterial!Requisitado
            
            TBNivel5.MoveNext
        Loop
    End If
    TBNivel5.Close
End If
QuantSolicitado = QuantsolicitadoN4

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEmpenhoPIN6(Desenho As String, IDcarteira As Long, Versao_estrutura As String, QuantSolicitado1 As Double)
On Error GoTo tratar_erro

Set TBItem = CreateObject("adodb.recordset")
TBItem.Open "Select * from projproduto where desenho = '" & Desenho & "' and DtValidacaoConj IS NOT NULL", Conexao, adOpenKeyset, adLockOptimistic
If TBItem.EOF = False Then
    Set TBNivel6 = CreateObject("adodb.recordset")
    TBNivel6.Open "Select PC.*, P.PBruto, P.SubTipoItem, P.Desenho from projconjunto PC INNER JOIN projproduto P ON P.Desenho = PC.Desenho where PC.codproduto = " & TBItem!Codproduto & " and PC.Versao = '" & Versao_estrutura & "' and P.bloqueado = 'False' order by PC.Posicao, PC.codigo", Conexao, adOpenKeyset, adLockOptimistic
    If TBNivel6.EOF = False Then
        Do While TBNivel6.EOF = False
            Set TBMaterial = CreateObject("adodb.recordset")
            TBMaterial.Open "Select * from producaomaterial where ID_carteira = " & IDcarteira & " and codigo = '" & TBNivel6!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBMaterial.EOF = True Then TBMaterial.AddNew
            QuantSolicitado = QuantSolicitado1
            ProcEnviaDadosEmpenhoPI IDcarteira, TBNivel6!quantidade, TBNivel6!Unidade, TBNivel6!Desenho, TBNivel6!Descricao, TBNivel6!PesoMetro, TBNivel6!Peso, TBNivel6!PesoTotal, TBNivel6!Dimensoes, TBNivel6!Un_Kg, TBNivel6!versao, TBNivel6!SubTipoItem, IIf(IsNull(TBNivel6!PBruto), 0, TBNivel6!PBruto), IIf(IsNull(TBNivel6!Percentual_perda), 0, TBNivel6!Percentual_perda)
            TBMaterial.Update
            
            QuantsolicitadoN6 = QuantSolicitado
            ProcEmpenhoPIN7 TBNivel6!Desenho, IDcarteira, Versao_estrutura, TBMaterial!Requisitado
            
            TBNivel6.MoveNext
        Loop
    End If
    TBNivel6.Close
End If
QuantSolicitado = QuantsolicitadoN5

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEmpenhoPIN7(Desenho As String, IDcarteira As Long, Versao_estrutura As String, QuantSolicitado1 As Double)
On Error GoTo tratar_erro

Set TBItem = CreateObject("adodb.recordset")
TBItem.Open "Select * from projproduto where desenho = '" & Desenho & "' and DtValidacaoConj IS NOT NULL", Conexao, adOpenKeyset, adLockOptimistic
If TBItem.EOF = False Then
    Set TBNivel7 = CreateObject("adodb.recordset")
    TBNivel7.Open "Select PC.*, P.PBruto, P.SubTipoItem, P.Desenho from projconjunto PC INNER JOIN projproduto P ON P.Desenho = PC.Desenho where PC.codproduto = " & TBItem!Codproduto & " and PC.Versao = '" & Versao_estrutura & "' and P.bloqueado = 'False' order by PC.Posicao, PC.codigo", Conexao, adOpenKeyset, adLockOptimistic
    If TBNivel7.EOF = False Then
        Do While TBNivel7.EOF = False
            Set TBMaterial = CreateObject("adodb.recordset")
            TBMaterial.Open "Select * from producaomaterial where ID_carteira = " & IDcarteira & " and codigo = '" & TBNivel7!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBMaterial.EOF = True Then TBMaterial.AddNew
            QuantSolicitado = QuantSolicitado1
            ProcEnviaDadosEmpenhoPI IDcarteira, TBNivel7!quantidade, TBNivel7!Unidade, TBNivel7!Desenho, TBNivel7!Descricao, TBNivel7!PesoMetro, TBNivel7!Peso, TBNivel7!PesoTotal, TBNivel7!Dimensoes, TBNivel7!Un_Kg, TBNivel7!versao, TBNivel7!SubTipoItem, IIf(IsNull(TBNivel7!PBruto), 0, TBNivel7!PBruto), IIf(IsNull(TBNivel7!Percentual_perda), 0, TBNivel7!Percentual_perda)
            TBMaterial.Update
            
            QuantsolicitadoN7 = QuantSolicitado
            ProcEmpenhoPIN8 TBNivel7!Desenho, IDcarteira, Versao_estrutura, TBMaterial!Requisitado
            
            TBNivel7.MoveNext
        Loop
    End If
    TBNivel7.Close
End If
QuantSolicitado = QuantsolicitadoN6

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEmpenhoPIN8(Desenho As String, IDcarteira As Long, Versao_estrutura As String, QuantSolicitado1 As Double)
On Error GoTo tratar_erro

Set TBItem = CreateObject("adodb.recordset")
TBItem.Open "Select * from projproduto where desenho = '" & Desenho & "' and DtValidacaoConj IS NOT NULL", Conexao, adOpenKeyset, adLockOptimistic
If TBItem.EOF = False Then
    Set TBNivel8 = CreateObject("adodb.recordset")
    TBNivel8.Open "Select PC.*, P.PBruto, P.SubTipoItem, P.Desenho from projconjunto PC INNER JOIN projproduto P ON P.Desenho = PC.Desenho where PC.codproduto = " & TBItem!Codproduto & " and PC.Versao = '" & Versao_estrutura & "' and P.bloqueado = 'False' order by PC.Posicao, PC.codigo", Conexao, adOpenKeyset, adLockOptimistic
    If TBNivel8.EOF = False Then
        Do While TBNivel8.EOF = False
            Set TBMaterial = CreateObject("adodb.recordset")
            TBMaterial.Open "Select * from producaomaterial where ID_carteira = " & IDcarteira & " and codigo = '" & TBNivel8!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBMaterial.EOF = True Then TBMaterial.AddNew
            QuantSolicitado = QuantSolicitado1
            ProcEnviaDadosEmpenhoPI IDcarteira, TBNivel8!quantidade, TBNivel8!Unidade, TBNivel8!Desenho, TBNivel8!Descricao, TBNivel8!PesoMetro, TBNivel8!Peso, TBNivel8!PesoTotal, TBNivel8!Dimensoes, TBNivel8!Un_Kg, TBNivel8!versao, TBNivel8!SubTipoItem, IIf(IsNull(TBNivel8!PBruto), 0, TBNivel8!PBruto), IIf(IsNull(TBNivel8!Percentual_perda), 0, TBNivel8!Percentual_perda)
            TBMaterial.Update
            
            QuantsolicitadoN8 = QuantSolicitado
            ProcEmpenhoPIN9 TBNivel8!Desenho, IDcarteira, Versao_estrutura, TBMaterial!Requisitado
            
            TBNivel8.MoveNext
        Loop
    End If
    TBNivel8.Close
End If
QuantSolicitado = QuantsolicitadoN7

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEmpenhoPIN9(Desenho As String, IDcarteira As Long, Versao_estrutura As String, QuantSolicitado1 As Double)
On Error GoTo tratar_erro

Set TBItem = CreateObject("adodb.recordset")
TBItem.Open "Select * from projproduto where desenho = '" & Desenho & "' and DtValidacaoConj IS NOT NULL", Conexao, adOpenKeyset, adLockOptimistic
If TBItem.EOF = False Then
    Set TBNivel9 = CreateObject("adodb.recordset")
    TBNivel9.Open "Select PC.*, P.PBruto, P.SubTipoItem, P.Desenho from projconjunto PC INNER JOIN projproduto P ON P.Desenho = PC.Desenho where PC.codproduto = " & TBItem!Codproduto & " and PC.Versao = '" & Versao_estrutura & "' and P.bloqueado = 'False' order by PC.Posicao, PC.codigo", Conexao, adOpenKeyset, adLockOptimistic
    If TBNivel9.EOF = False Then
        Do While TBNivel9.EOF = False
            Set TBMaterial = CreateObject("adodb.recordset")
            TBMaterial.Open "Select * from producaomaterial where ID_carteira = " & IDcarteira & " and codigo = '" & TBNivel9!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBMaterial.EOF = True Then TBMaterial.AddNew
            QuantSolicitado = QuantSolicitado1
            ProcEnviaDadosEmpenhoPI IDcarteira, TBNivel9!quantidade, TBNivel9!Unidade, TBNivel9!Desenho, TBNivel9!Descricao, TBNivel9!PesoMetro, TBNivel9!Peso, TBNivel9!PesoTotal, TBNivel9!Dimensoes, TBNivel9!Un_Kg, TBNivel9!versao, TBNivel9!SubTipoItem, IIf(IsNull(TBNivel9!PBruto), 0, TBNivel9!PBruto), IIf(IsNull(TBNivel9!Percentual_perda), 0, TBNivel9!Percentual_perda)
            TBMaterial.Update
            
            QuantsolicitadoN9 = QuantSolicitado
            ProcEmpenhoPIN10 TBNivel9!Desenho, IDcarteira, Versao_estrutura, TBMaterial!Requisitado
            
            TBNivel9.MoveNext
        Loop
    End If
    TBNivel9.Close
End If
QuantSolicitado = QuantsolicitadoN8

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEmpenhoPIN10(Desenho As String, IDcarteira As Long, Versao_estrutura As String, QuantSolicitado1 As Double)
On Error GoTo tratar_erro

Set TBItem = CreateObject("adodb.recordset")
TBItem.Open "Select * from projproduto where desenho = '" & Desenho & "' and DtValidacaoConj IS NOT NULL", Conexao, adOpenKeyset, adLockOptimistic
If TBItem.EOF = False Then
    Set TBNivel10 = CreateObject("adodb.recordset")
    TBNivel10.Open "Select PC.*, P.PBruto, P.SubTipoItem, P.Desenho from projconjunto PC INNER JOIN projproduto P ON P.Desenho = PC.Desenho where PC.codproduto = " & TBItem!Codproduto & " and PC.Versao = '" & Versao_estrutura & "' and P.bloqueado = 'False' order by PC.Posicao, PC.codigo", Conexao, adOpenKeyset, adLockOptimistic
    If TBNivel10.EOF = False Then
        Do While TBNivel10.EOF = False
            Set TBMaterial = CreateObject("adodb.recordset")
            TBMaterial.Open "Select * from producaomaterial where ID_carteira = " & IDcarteira & " and codigo = '" & TBNivel10!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBMaterial.EOF = True Then TBMaterial.AddNew
            QuantSolicitado = QuantSolicitado1
            ProcEnviaDadosEmpenhoPI IDcarteira, TBNivel10!quantidade, TBNivel10!Unidade, TBNivel10!Desenho, TBNivel10!Descricao, TBNivel10!PesoMetro, TBNivel10!Peso, TBNivel10!PesoTotal, TBNivel10!Dimensoes, TBNivel10!Un_Kg, TBNivel10!versao, TBNivel10!SubTipoItem, IIf(IsNull(TBNivel10!PBruto), 0, TBNivel10!PBruto), IIf(IsNull(TBNivel10!Percentual_perda), 0, TBNivel10!Percentual_perda)
            TBMaterial.Update
            
            QuantsolicitadoN10 = QuantSolicitado
            ProcEmpenhoPIN11 TBNivel10!Desenho, IDcarteira, Versao_estrutura, TBMaterial!Requisitado
            
            TBNivel10.MoveNext
        Loop
    End If
    TBNivel10.Close
End If
QuantSolicitado = QuantsolicitadoN9

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEmpenhoPIN11(Desenho As String, IDcarteira As Long, Versao_estrutura As String, QuantSolicitado1 As Double)
On Error GoTo tratar_erro

Set TBItem = CreateObject("adodb.recordset")
TBItem.Open "Select * from projproduto where desenho = '" & Desenho & "' and DtValidacaoConj IS NOT NULL", Conexao, adOpenKeyset, adLockOptimistic
If TBItem.EOF = False Then
    Set TBNivel11 = CreateObject("adodb.recordset")
    TBNivel11.Open "Select PC.*, P.PBruto, P.SubTipoItem, P.Desenho from projconjunto PC INNER JOIN projproduto P ON P.Desenho = PC.Desenho where PC.codproduto = " & TBItem!Codproduto & " and PC.Versao = '" & Versao_estrutura & "' and P.bloqueado = 'False' order by PC.Posicao, PC.codigo", Conexao, adOpenKeyset, adLockOptimistic
    If TBNivel11.EOF = False Then
        Do While TBNivel11.EOF = False
            Set TBMaterial = CreateObject("adodb.recordset")
            TBMaterial.Open "Select * from producaomaterial where ID_carteira = " & IDcarteira & " and codigo = '" & TBNivel11!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBMaterial.EOF = True Then TBMaterial.AddNew
            QuantSolicitado = QuantSolicitado1
            ProcEnviaDadosEmpenhoPI IDcarteira, TBNivel11!quantidade, TBNivel11!Unidade, TBNivel11!Desenho, TBNivel11!Descricao, TBNivel11!PesoMetro, TBNivel11!Peso, TBNivel11!PesoTotal, TBNivel11!Dimensoes, TBNivel11!Un_Kg, TBNivel11!versao, TBNivel11!SubTipoItem, IIf(IsNull(TBNivel11!PBruto), 0, TBNivel11!PBruto), IIf(IsNull(TBNivel11!Percentual_perda), 0, TBNivel11!Percentual_perda)
            TBMaterial.Update
            
            QuantsolicitadoN11 = QuantSolicitado
            ProcEmpenhoPIN12 TBNivel11!Desenho, IDcarteira, Versao_estrutura, TBMaterial!Requisitado
            
            TBNivel11.MoveNext
        Loop
    End If
    TBNivel11.Close
End If
QuantSolicitado = QuantsolicitadoN10

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEmpenhoPIN12(Desenho As String, IDcarteira As Long, Versao_estrutura As String, QuantSolicitado1 As Double)
On Error GoTo tratar_erro

Set TBItem = CreateObject("adodb.recordset")
TBItem.Open "Select * from projproduto where desenho = '" & Desenho & "' and DtValidacaoConj IS NOT NULL", Conexao, adOpenKeyset, adLockOptimistic
If TBItem.EOF = False Then
    Set TBNivel12 = CreateObject("adodb.recordset")
    TBNivel12.Open "Select PC.*, P.PBruto, P.SubTipoItem, P.Desenho from projconjunto PC INNER JOIN projproduto P ON P.Desenho = PC.Desenho where PC.codproduto = " & TBItem!Codproduto & " and PC.Versao = '" & Versao_estrutura & "' and P.bloqueado = 'False' order by PC.Posicao, PC.codigo", Conexao, adOpenKeyset, adLockOptimistic
    If TBNivel12.EOF = False Then
        Do While TBNivel12.EOF = False
            Set TBMaterial = CreateObject("adodb.recordset")
            TBMaterial.Open "Select * from producaomaterial where ID_carteira = " & IDcarteira & " and codigo = '" & TBNivel12!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBMaterial.EOF = True Then TBMaterial.AddNew
            QuantSolicitado = QuantSolicitado1
            ProcEnviaDadosEmpenhoPI IDcarteira, TBNivel12!quantidade, TBNivel12!Unidade, TBNivel12!Desenho, TBNivel12!Descricao, TBNivel12!PesoMetro, TBNivel12!Peso, TBNivel12!PesoTotal, TBNivel12!Dimensoes, TBNivel12!Un_Kg, TBNivel12!versao, TBNivel12!SubTipoItem, IIf(IsNull(TBNivel12!PBruto), 0, TBNivel12!PBruto), IIf(IsNull(TBNivel12!Percentual_perda), 0, TBNivel12!Percentual_perda)
            TBMaterial.Update
            
            QuantsolicitadoN12 = QuantSolicitado
            ProcEmpenhoPIN13 TBNivel12!Desenho, IDcarteira, Versao_estrutura, TBMaterial!Requisitado
            
            TBNivel12.MoveNext
        Loop
    End If
    TBNivel12.Close
End If
QuantSolicitado = QuantsolicitadoN11

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEmpenhoPIN13(Desenho As String, IDcarteira As Long, Versao_estrutura As String, QuantSolicitado1 As Double)
On Error GoTo tratar_erro

Set TBItem = CreateObject("adodb.recordset")
TBItem.Open "Select * from projproduto where desenho = '" & Desenho & "' and DtValidacaoConj IS NOT NULL", Conexao, adOpenKeyset, adLockOptimistic
If TBItem.EOF = False Then
    Set TBNivel13 = CreateObject("adodb.recordset")
    TBNivel13.Open "Select PC.*, P.PBruto, P.SubTipoItem, P.Desenho from projconjunto PC INNER JOIN projproduto P ON P.Desenho = PC.Desenho where PC.codproduto = " & TBItem!Codproduto & " and PC.Versao = '" & Versao_estrutura & "' and P.bloqueado = 'False' order by PC.Posicao, PC.codigo", Conexao, adOpenKeyset, adLockOptimistic
    If TBNivel13.EOF = False Then
        Do While TBNivel13.EOF = False
            Set TBMaterial = CreateObject("adodb.recordset")
            TBMaterial.Open "Select * from producaomaterial where ID_carteira = " & IDcarteira & " and codigo = '" & TBNivel13!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBMaterial.EOF = True Then TBMaterial.AddNew
            QuantSolicitado = QuantSolicitado1
            ProcEnviaDadosEmpenhoPI IDcarteira, TBNivel13!quantidade, TBNivel13!Unidade, TBNivel13!Desenho, TBNivel13!Descricao, TBNivel13!PesoMetro, TBNivel13!Peso, TBNivel13!PesoTotal, TBNivel13!Dimensoes, TBNivel13!Un_Kg, TBNivel13!versao, TBNivel13!SubTipoItem, IIf(IsNull(TBNivel13!PBruto), 0, TBNivel13!PBruto), IIf(IsNull(TBNivel13!Percentual_perda), 0, TBNivel13!Percentual_perda)
            TBMaterial.Update
            
            QuantsolicitadoN13 = QuantSolicitado
            ProcEmpenhoPIN14 TBNivel13!Desenho, IDcarteira, Versao_estrutura, TBMaterial!Requisitado
            
            TBNivel13.MoveNext
        Loop
    End If
    TBNivel13.Close
End If
QuantSolicitado = QuantsolicitadoN12

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEmpenhoPIN14(Desenho As String, IDcarteira As Long, Versao_estrutura As String, QuantSolicitado1 As Double)
On Error GoTo tratar_erro

Set TBItem = CreateObject("adodb.recordset")
TBItem.Open "Select * from projproduto where desenho = '" & Desenho & "' and DtValidacaoConj IS NOT NULL", Conexao, adOpenKeyset, adLockOptimistic
If TBItem.EOF = False Then
    Set TBNivel14 = CreateObject("adodb.recordset")
    TBNivel14.Open "Select PC.*, P.PBruto, P.SubTipoItem, P.Desenho from projconjunto PC INNER JOIN projproduto P ON P.Desenho = PC.Desenho where PC.codproduto = " & TBItem!Codproduto & " and PC.Versao = '" & Versao_estrutura & "' and P.bloqueado = 'False' order by PC.Posicao, PC.codigo", Conexao, adOpenKeyset, adLockOptimistic
    If TBNivel14.EOF = False Then
        Do While TBNivel14.EOF = False
            Set TBMaterial = CreateObject("adodb.recordset")
            TBMaterial.Open "Select * from producaomaterial where ID_carteira = " & IDcarteira & " and codigo = '" & TBNivel14!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBMaterial.EOF = True Then TBMaterial.AddNew
            QuantSolicitado = QuantSolicitado1
            ProcEnviaDadosEmpenhoPI IDcarteira, TBNivel14!quantidade, TBNivel14!Unidade, TBNivel14!Desenho, TBNivel14!Descricao, TBNivel14!PesoMetro, TBNivel14!Peso, TBNivel14!PesoTotal, TBNivel14!Dimensoes, TBNivel14!Un_Kg, TBNivel14!versao, TBNivel14!SubTipoItem, IIf(IsNull(TBNivel14!PBruto), 0, TBNivel14!PBruto), IIf(IsNull(TBNivel14!Percentual_perda), 0, TBNivel14!Percentual_perda)
            TBMaterial.Update
            
            QuantsolicitadoN14 = QuantSolicitado
            ProcEmpenhoPIN15 TBNivel14!Desenho, IDcarteira, Versao_estrutura, TBMaterial!Requisitado
            
            TBNivel14.MoveNext
        Loop
    End If
    TBNivel14.Close
End If
QuantSolicitado = QuantsolicitadoN13

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEmpenhoPIN15(Desenho As String, IDcarteira As Long, Versao_estrutura As String, QuantSolicitado1 As Double)
On Error GoTo tratar_erro

Set TBItem = CreateObject("adodb.recordset")
TBItem.Open "Select * from projproduto where desenho = '" & Desenho & "' and DtValidacaoConj IS NOT NULL", Conexao, adOpenKeyset, adLockOptimistic
If TBItem.EOF = False Then
    Set TBNivel15 = CreateObject("adodb.recordset")
    TBNivel15.Open "Select PC.*, P.PBruto, P.SubTipoItem, P.Desenho from projconjunto PC INNER JOIN projproduto P ON P.Desenho = PC.Desenho where PC.codproduto = " & TBItem!Codproduto & " and PC.Versao = '" & Versao_estrutura & "' and P.bloqueado = 'False' order by PC.Posicao, PC.codigo", Conexao, adOpenKeyset, adLockOptimistic
    If TBNivel15.EOF = False Then
        Do While TBNivel15.EOF = False
            Set TBMaterial = CreateObject("adodb.recordset")
            TBMaterial.Open "Select * from producaomaterial where ID_carteira = " & IDcarteira & " and codigo = '" & TBNivel15!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBMaterial.EOF = True Then TBMaterial.AddNew
            QuantSolicitado = QuantSolicitado1
            ProcEnviaDadosEmpenhoPI IDcarteira, TBNivel15!quantidade, TBNivel15!Unidade, TBNivel15!Desenho, TBNivel15!Descricao, TBNivel15!PesoMetro, TBNivel15!Peso, TBNivel15!PesoTotal, TBNivel15!Dimensoes, TBNivel15!Un_Kg, TBNivel15!versao, TBNivel15!SubTipoItem, IIf(IsNull(TBNivel15!PBruto), 0, TBNivel15!PBruto), IIf(IsNull(TBNivel15!Percentual_perda), 0, TBNivel15!Percentual_perda)
            TBMaterial.Update
            
            'QuantsolicitadoN15 = QuantSolicitado
            'ProcEmpenhoPIN16 TBNivel15!Desenho, IDCarteira, Versao_estrutura, TBMaterial!Requisitado
            
            TBNivel15.MoveNext
        Loop
    End If
    TBNivel15.Close
End If
QuantSolicitado = QuantsolicitadoN14

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviaDadosEmpenhoPI(IDcarteira As Long, quantidade As Double, Unidade As String, Desenho As String, Descricao As String, PesoMetro As Double, Peso As Double, PesoTotal As Double, Dimensoes As Double, Un_Kg As String, versao As String, SubTipoItem As Integer, PBruto As Double, Percentual_perda As Double)
On Error GoTo tratar_erro

TBMaterial!quantidade = TBMaterial!quantidade + (quantidade * QuantSolicitado)
TBMaterial!Unidade = Unidade
TBMaterial!CODIGO = Desenho
TBMaterial!Descricao = Descricao
TBMaterial!ID_carteira = IDcarteira
TBMaterial!PesoMetro = PesoMetro
TBMaterial!pesounidade = Peso
TBMaterial!PesoTotal = TBMaterial!PesoTotal + (PesoTotal * QuantSolicitado)
TBMaterial!Percentual_perda = Percentual_perda
TBMaterial!Dimensao = Dimensoes
If Un_Kg = "Mt≤" Then TBMaterial!DimensaoTotal = ((Dimensoes / 1000) / 1000) * TBMaterial!quantidade Else TBMaterial!DimensaoTotal = (Dimensoes / 1000) * TBMaterial!quantidade
TBMaterial!versao = versao
Peso = quantidade
If Un_Kg <> "N/a" And Un_Kg <> "" And (Unidade = "KG" Or Unidade = "MT" Or Unidade = "MM" Or Unidade = "M≥") Then
    Select Case Unidade
        Case "KG": Peso = PesoTotal
        Case "MT": Peso = (Dimensoes / 1000) * quantidade
        Case "MM": Peso = Dimensoes * quantidade
        Case "M≥": Peso = TBProcessos!PesoTotal
    End Select
End If
            
 If TBProcessos!Unidade = "M≥" Then
     TBMaterial!Requisitado = Peso * txtQuantidade
     TBMaterial!DimensaoTotal = TBProcessos!Dimensoes
     TBMaterial!Total_pc = TBProcessos!quantidade
 Else
    TBMaterial!Requisitado = Format(Peso * txtQuantidade, "###,##0.0000")
    If TBProcessos!Un_Kg = "Mt≤" Then TBMaterial!DimensaoTotal = ((TBProcessos!Dimensoes / 1000) / 1000) * TBMaterial!quantidade Else TBMaterial!DimensaoTotal = (TBProcessos!Dimensoes / 1000) * TBMaterial!quantidade
    If TBProcessos!Unidade = "KG" Or TBProcessos!SubTipoItem = 1 Or TBProcessos!SubTipoItem = 2 Or TBProcessos!SubTipoItem = 3 Then
        If TBProcessos!Unidade = "KG" And (TBProcessos!Un_Kg = "Mt≤" Or TBProcessos!Un_Kg = "Mt/L") Then
            If IsNull(TBProcessos!PBruto) = False And TBProcessos!PBruto > 0 And TBProcessos!PBruto <> "" Then TBMaterial!Total_pc = Format(TBMaterial!Requisitado / TBProcessos!PBruto, "###,##0.0000") Else TBMaterial!Total_pc = Null
        Else
            If TBProcessos!Unidade = "P«" Or TBProcessos!Unidade = "PC" Or TBProcessos!Unidade = "UN" Or TBProcessos!Unidade = "CJ" Then TBMaterial!Total_pc = TBMaterial!Requisitado Else TBMaterial!Total_pc = Null
        End If
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Public Function FunCalcularQtdeUnKg(LarguraPrin As Double, CompriPrin As Double, LarguraUnKg As Double, CompriUnKg As Double, UnKg As String)
On Error GoTo tratar_erro

If UnKg = "Mt≤" Then
    FunCalcularQtdeUnKg = (LarguraPrin * CompriPrin) / (LarguraUnKg * CompriUnKg)
Else
    FunCalcularQtdeUnKg = (CompriPrin * 1) / CompriUnKg
End If

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Function

Function FunBloqueiaUnConversao(ProdutoUN As String, Un_est As ComboBox, Un_com As ComboBox, ProdServUN As Boolean) As Boolean
On Error GoTo tratar_erro

'Se a unidade for diferente verifica se esta cadastrado o peso bruto e peso metro
If Un_est <> Un_com And (Un_est = "KG" Or Un_est = "MT" Or Un_est = "MM" Or Un_est = "BR" Or Un_est = "PC" Or Un_est = "P«") And (Un_com = "KG" Or Un_com = "MT" Or Un_com = "MM" Or Un_com = "BR" Or Un_com = "PC" Or Un_com = "P«") Then
    If Un_com = "P«" Or Un_com = "PC" Or Un_com = "BR" Or Un_est = "P«" Or Un_est = "PC" Or Un_est = "BR" Then
        MensagemUN = "peso bruto"
        PesquisaUN = "(PBruto IS NULL or PBruto <= 0)"
    Else
        MensagemUN = "Kg/Un"
        PesquisaUN = "(peso_metro IS NULL or peso_metro <= 0)"
    End If
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select Codproduto from projproduto where desenho = '" & ProdutoUN & "' and " & PesquisaUN, Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
        USMsgBox ("N„o È permitido salvar este " & IIf(ProdServUN = True, "produto", "serviÁo") & ", pois as unidades s„o diferentes e o " & MensagemUN & " n„o foi informado no cadastro do produto."), vbExclamation, "CAPRIND v5.0"
        FunBloqueiaUnConversao = True
    End If
    TBProduto.Close
End If

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Function

Function FunVerifCidadeEmpresa(IDempresa As Long) As String
On Error GoTo tratar_erro

FunVerifCidadeEmpresa = ""
Set TBCFOP = CreateObject("adodb.recordset")
TBCFOP.Open "Select Cidade from Empresa where Codigo = " & IDempresa, Conexao, adOpenKeyset, adLockOptimistic
If TBCFOP.EOF = False Then
    FunVerifCidadeEmpresa = TBCFOP!Cidade
End If
TBCFOP.Close

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Function FunVerifUFEmpresa(IDempresa As Long) As String
On Error GoTo tratar_erro

FunVerifUFEmpresa = ""
Set TBCFOP = CreateObject("adodb.recordset")
TBCFOP.Open "Select uf from Empresa where Codigo = " & IDempresa, Conexao, adOpenKeyset, adLockOptimistic
If TBCFOP.EOF = False Then

Select Case TBCFOP!UF
  Case "RO": nsUF = "11"
  Case "AC": nsUF = "12"
  Case "AM": nsUF = "13"
  Case "RR": nsUF = "14"
  Case "PA": nsUF = "15"
  Case "AP": nsUF = "16"
  Case "TO": nsUF = "17"
  Case "MA": nsUF = "21"
  Case "Pi": nsUF = "22"
  Case "CE": nsUF = "23"
  Case "RN": nsUF = "24"
  Case "PB": nsUF = "25"
  Case "PE": nsUF = "26"
  Case "AL": nsUF = "27"
  Case "SE": nsUF = "28"
  Case "BA": nsUF = "29"
  Case "MG": nsUF = "31"
  Case "ES": nsUF = "32"
  Case "RJ": nsUF = "33"
  Case "SP": nsUF = "35"
  Case "PR": nsUF = "41"
  Case "SC": nsUF = "42"
  Case "RS": nsUF = "43"
  Case "MS": nsUF = "50"
  Case "MT": nsUF = "51"
  Case "GO": nsUF = "52"
  Case "DF": nsUF = "53"
End Select

FunVerifUFEmpresa = nsUF
End If

TBCFOP.Close

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function
Public Function FunVerifNCPorDescricao(OS As Long) As Boolean
On Error GoTo tratar_erro

FunVerifNCPorDescricao = False
Set TBAliquota = CreateObject("adodb.recordset")
StrSql = "Select E.Codigo from (OrdemServico OS INNER JOIN Producao P ON P.Ordem = OS.Ordem) INNER JOIN Empresa E ON E.Codigo = P.ID_empresa where OS.Idproducao = " & OS & " and E.Apontar_NC_descricao = 1"

'Debug.print

TBAliquota.Open "Select E.Codigo from (OrdemServico OS INNER JOIN Producao P ON P.Ordem = OS.Ordem) INNER JOIN Empresa E ON E.Codigo = P.ID_empresa where OS.Idproducao = " & OS & " and E.Apontar_NC_descricao = 1", Conexao, adOpenKeyset, adLockReadOnly
If TBAliquota.EOF = False Then
    FunVerifNCPorDescricao = True
End If
TBAliquota.Close

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Function

Public Function FunVerifNCRejeitado(OS As Long) As Boolean
On Error GoTo tratar_erro

FunVerifNCRejeitado = False
Set TBAliquota = CreateObject("adodb.recordset")

TBAliquota.Open "Select E.Codigo from (OrdemServico OS INNER JOIN Producao P ON P.Ordem = OS.Ordem) INNER JOIN Empresa E ON E.Codigo = P.ID_empresa where OS.Idproducao = " & OS & " and E.Apontar_NC_descricao = 1", Conexao, adOpenKeyset, adLockReadOnly
If TBAliquota.EOF = False Then
    FunVerifNCRejeitado = True
End If
TBAliquota.Close

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Function

Public Function FunConversaoFinalUn(Unest As String, Uncom As String, Qtde As Double, Codinterno As String, UnComparaUnEst As Boolean) As Double
On Error GoTo tratar_erro

FunConversaoFinalUn = 0
If Unest <> Uncom Then
    If UnComparaUnEst = True Then
        If FunVerifUNConversao(Unest, Uncom) = True Then
            FunConversaoFinalUn = FunConverteUN(Unest, Uncom, Qtde, Codinterno)
        Else
            FunConversaoFinalUn = Qtde / FunVerificaTabelaConversaoUnidade(Unest, Uncom)
        End If
    Else
        If FunVerifUNConversao(Unest, Uncom) = True Then
            FunConversaoFinalUn = FunConverteUN(Uncom, Unest, Qtde, Codinterno)
        Else
            FunConversaoFinalUn = Qtde * FunVerificaTabelaConversaoUnidade(Unest, Uncom)
        End If
    End If
End If

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Function

Public Sub ProcConfigNFeTECNOSPEED(IDNota_Tecno As Integer, certificado_Tecno As String)
On Error GoTo tratar_erro

Set TBCFOP = CreateObject("adodb.recordset")
TBCFOP.Open "Select E.* from Empresa E INNER JOIN tbl_Dados_Nota_Fiscal NF ON E.Codigo = NF.ID_empresa where NF.ID = " & IDNota_Tecno, Conexao, adOpenKeyset, adLockReadOnly
If TBCFOP.EOF = False Then
    '[NFe]
    spdNFe.UF = TBCFOP!UF
    spdNFe.CNPJ = ReturnNumbersOnly(TBCFOP!CNPJ)
    spdNFe.Ambiente = "1" 'Informar 1 para produÁ„o e 2 para homologaÁ„o
    spdNFe.ArquivoServidoresHom = Localrel & "\NFe\arquivos\nfeServidoresHom.ini"
    spdNFe.ArquivoServidoresProd = Localrel & "\NFe\arquivos\nfeServidoresProd.ini"
    spdNFe.DiretorioEsquemas = Localrel & "\NFe\arquivos\Esquemas\"
    spdNFe.DiretorioTemplates = Localrel & "\NFe\arquivos\Templates\"
    spdNFe.DiretorioLog = Localrel & "\NFe\Log\"
    spdNFe.DiretorioLogErro = Localrel & "\NFe\LogErro\"
    spdNFe.DiretorioTemporario = Localrel & "\NFe\Temporario\"
    spdNFe.DiretorioXmlDestinatario = IIf(IsNull(TBCFOP!Caminho_XMLDanfe), Localrel & "\NFe\XmlDestinatario", TBCFOP!Caminho_XMLDanfe)
    spdNFe.TipoCertificado = ckFile
    spdNFe.NomeCertificado = certificado_Tecno
    spdNFe.VersaoManual = "5.0a"
    spdNFe.ConexaoSegura = True
    spdNFe.ValidarEsquemaAntesEnvio = True
    spdNFe.MaxSizeLoteEnvio = 500
    spdNFe.AnexarDanfePDF = True
    spdNFe.CaracteresRemoverAcentos = "·ÈÌÛ˙‡ËÏÚ˘‚ÍÓÙ˚‰ÎÔˆ¸„ıÒÁ¡…Õ”⁄¿»Ã“Ÿ¬ Œ‘€ƒÀœ÷‹√’—"
    
    Set TBCFOP_NFe = CreateObject("adodb.recordset")
    TBCFOP_NFe.Open "Select * from Empresa_Email where ID_empresa = " & TBCFOP!CODIGO & " AND Aplicacao = 'FA'", Conexao, adOpenKeyset, adLockReadOnly
    If TBCFOP_NFe.EOF = False Then
        '[MAIL
        spdNFe.EmailServidor = TBCFOP_NFe!Servidor_SMTP
        spdNFe.EmailRemetente = TBCFOP_NFe!Email
        spdNFe.EmailAssunto = "Segue em anexo DANFE e XML"
        spdNFe.EmailMensagem = "Segue em anexo nota fiscal eletrÙnica."
        spdNFe.EmailUsuario = TBCFOP_NFe!Usuario
        spdNFe.EmailSenha = TBCFOP_NFe!Senha
        spdNFe.EmailAutenticacao = IIf(TBCFOP_NFe!Seguranca = "S", True, False)
        spdNFe.EmailPorta = TBCFOP_NFe!Porta
        spdNFe.AnexarDanfePDF = True
    End If
    TBCFOP_NFe.Close
    
    '[DANFE
    spdNFe.LogotipoEmitente = TBCFOP!Logotipo
    spdNFe.InfCplMaxCol = "68"
    spdNFe.InfCplMaxRow = "7"
    spdNFe.FraseContingencia = "Danfe em contingÍncia - Impresso em decorrÍncia de problemas tÈcnicos."
    spdNFe.FraseHomologacao = "SEM VALOR FISCAL"
    spdNFe.ModeloRetrato = spdNFe.DiretorioTemplates + "Danfe\\retrato.rtm"
    spdNFe.ModeloPaisagem = spdNFe.DiretorioTemplates + "Danfe\\paisagem.rtm"
    spdNFe.QtdeCopias = "1"
        
    'GERAR DATASET ########################################################################################
    NumeroLote = 1
    spdNFe.VersaoManual = "5.0a"
    spdNFe.DiretorioTemplates = Localrel & "\NFe\arquivos\Templates"
    spdNFe.DiretorioEsquemas = Localrel & "\NFe\arquivos\Esquemas"
    spdNFeDataSet.VersaoEsquema = "pl_008h"
    spdNFeDataSet.DicionarioXML = spdNFe.DiretorioTemplates + "Conversor\NFeDataSets.xml"
    spdNFeDataSet.ValidaRegrasNegocio = False
    spdNFeDataSet.ValidaRegrasNegocioTecno = False
End If
TBCFOP.Close
         
Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Function funcSomaEmailNF(idNF_Email As Integer) As String
On Error GoTo tratar_erro
Dim TBSomaEmail  As ADODB.Recordset
Dim TBSomaEmail2  As ADODB.Recordset
Dim TBSomaEmail3 As ADODB.Recordset
Dim TBSomaEmail4 As ADODB.Recordset

Set TBSomaEmail = CreateObject("adodb.recordset")
TBSomaEmail.Open "Select * from tbl_Dados_Nota_Fiscal WHERE ID = " & idNF_Email, Conexao, adOpenKeyset, adLockOptimistic
If TBSomaEmail.EOF = False Then
    funcSomaEmailNF = ""
    'Verifica email e paÌs
    Set TBSomaEmail2 = CreateObject("adodb.recordset")
    If TBSomaEmail!txt_tipocliente = "E" Then
        'Empresa
        TBSomaEmail2.Open "Select Email, Pais, Codigo_pais from Empresa where Codigo=" & TBSomaEmail!Id_Int_Cliente, Conexao, adOpenKeyset, adLockOptimistic
        If TBSomaEmail2.EOF = False Then
            funcSomaEmailNF = IIf(IsNull(TBSomaEmail2!Email), "", TBSomaEmail2!Email)
        End If
    ElseIf TBSomaEmail!txt_tipocliente = "JP" Or TBSomaEmail!txt_tipocliente = "JR" Or TBSomaEmail!txt_tipocliente = "FP" Or TBSomaEmail!txt_tipocliente = "FR" Then
        'Cliente
        TBSomaEmail2.Open "Select Email, Pais, Codigo_pais from Clientes where IDcliente=" & TBSomaEmail!Id_Int_Cliente & " and NomeRazao= '" & TBSomaEmail!txt_Razao_Nome & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBSomaEmail2.EOF = False Then
            funcSomaEmailNF = IIf(IsNull(TBSomaEmail2!Email), "", TBSomaEmail2!Email)
            If funcSomaEmailNF <> "" Then TextoFiltro = " and Email <> '" & funcSomaEmailNF & "'" Else TextoFiltro = ""
            
            Set TBSomaEmail3 = CreateObject("adodb.recordset")
            TBSomaEmail3.Open "Select Email from Clientes_Contatos where IDcliente=" & TBSomaEmail!Id_Int_Cliente & TextoFiltro & " and Enviar_NFe= 'True' and EMail is not null", Conexao, adOpenKeyset, adLockOptimistic
            If TBSomaEmail3.EOF = False Then
                Do While TBSomaEmail3.EOF = False
                    If IsNull(TBSomaEmail3!Email) = False And TBSomaEmail3!Email <> "" Then
                        If funcSomaEmailNF <> "" Then funcSomaEmailNF = funcSomaEmailNF & ";" & TBSomaEmail3!Email Else funcSomaEmailNF = TBSomaEmail3!Email
                    End If
                    TBSomaEmail3.MoveNext
                Loop
            End If
            TBSomaEmail3.Close
        End If
    Else
        'Fornecedor
        TBSomaEmail2.Open "Select Email, Pais, Codigo_pais from Compras_fornecedores where IDcliente=" & TBSomaEmail!Id_Int_Cliente & " and Nome_Razao= '" & TBSomaEmail!txt_Razao_Nome & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBSomaEmail2.EOF = False Then
            funcSomaEmailNF = IIf(IsNull(TBSomaEmail2!Email), "", TBSomaEmail2!Email)
            If funcSomaEmailNF <> "" Then TextoFiltro = " and Email <> '" & funcSomaEmailNF & "'" Else TextoFiltro = ""
            
            Set TBSomaEmail3 = CreateObject("adodb.recordset")
            TBSomaEmail3.Open "Select Email from Contatos_fornecedor where IdFornecedor=" & TBSomaEmail!Id_Int_Cliente & TextoFiltro & " and Enviar_NFe= 'True' and Email is not null", Conexao, adOpenKeyset, adLockOptimistic
            If TBSomaEmail3.EOF = False Then
                Do While TBSomaEmail3.EOF = False
                    If IsNull(TBSomaEmail3!Email) = False And TBSomaEmail3!Email <> "" Then
                        If funcSomaEmailNF <> "" Then funcSomaEmailNF = funcSomaEmailNF & ";" & TBSomaEmail3!Email Else funcSomaEmailNF = TBSomaEmail3!Email
                    End If
                    TBSomaEmail3.MoveNext
                Loop
            End If
            TBSomaEmail3.Close
        End If
    End If
    TBSomaEmail2.Close
    
    'Verifica se tem transportadora na NF para consultar o e-mail
    Email1 = ""
    Set TBSomaEmail4 = CreateObject("adodb.recordset")
    TBSomaEmail4.Open "Select CNPJ from Empresa where Codigo=" & TBSomaEmail!ID_empresa, Conexao, adOpenKeyset, adLockOptimistic
    If TBSomaEmail4.EOF = False Then
        Set TBSomaEmail3ltro = CreateObject("adodb.recordset")
        TBSomaEmail3ltro.Open "Select IdIntTransp, txt_Razao from tbl_Dados_Transp where ID_nota = " & TBSomaEmail!ID & " and TXT_CNPJ <> '" & TBSomaEmail4!CNPJ & "' and TXT_CNPJ <> '" & TBSomaEmail!txt_CNPJ_CPF & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBSomaEmail3ltro.EOF = False Then
            'Cliente
            Set TBSomaEmail2 = CreateObject("adodb.recordset")
            TBSomaEmail2.Open "Select Email from Clientes where IDcliente=" & TBSomaEmail3ltro!IdIntTransp & " and NomeRazao= '" & TBSomaEmail3ltro!txt_Razao & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBSomaEmail2.EOF = False Then
                Email1 = IIf(IsNull(TBSomaEmail2!Email), "", TBSomaEmail2!Email)
                If Email1 <> "" Then
                    TextoFiltro = " and Email <> '" & Email1 & "'"
                    If funcSomaEmailNF <> "" Then funcSomaEmailNF = funcSomaEmailNF & ";" & Email1 Else funcSomaEmailNF = Email1
                Else
                    TextoFiltro = ""
                End If
                
                Set TBSomaEmail3 = CreateObject("adodb.recordset")
                TBSomaEmail3.Open "Select Email from Clientes_Contatos where IDcliente=" & TBSomaEmail3ltro!IdIntTransp & TextoFiltro & " and Enviar_NFe= 'True' and EMail is not null", Conexao, adOpenKeyset, adLockOptimistic
                If TBSomaEmail3.EOF = False Then
                    Do While TBSomaEmail3.EOF = False
                        If IsNull(TBSomaEmail3!Email) = False And TBSomaEmail3!Email <> "" Then
                            If funcSomaEmailNF <> "" Then funcSomaEmailNF = funcSomaEmailNF & ";" & TBSomaEmail3!Email Else funcSomaEmailNF = TBSomaEmail3!Email
                        End If
                        TBSomaEmail3.MoveNext
                    Loop
                End If
                TBSomaEmail3.Close
            Else
                'Fornecedor
                Set TBSomaEmail2 = CreateObject("adodb.recordset")
                TBSomaEmail2.Open "Select Email from Compras_fornecedores where IDcliente=" & TBSomaEmail3ltro!IdIntTransp & " and Nome_Razao= '" & TBSomaEmail3ltro!txt_Razao & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBSomaEmail2.EOF = False Then
                    Email1 = IIf(IsNull(TBSomaEmail2!Email), "", TBSomaEmail2!Email)
                    If Email1 <> "" Then
                        TextoFiltro = " and Email <> '" & Email1 & "'"
                        If funcSomaEmailNF <> "" Then funcSomaEmailNF = funcSomaEmailNF & ";" & Email1 Else funcSomaEmailNF = Email1
                    Else
                        TextoFiltro = ""
                    End If
                    
                    Set TBSomaEmail3 = CreateObject("adodb.recordset")
                    TBSomaEmail3.Open "Select Email from Contatos_fornecedor where IdFornecedor=" & TBSomaEmail3ltro!IdIntTransp & TextoFiltro & " and Enviar_NFe= 'True' and Email is not null", Conexao, adOpenKeyset, adLockOptimistic
                    If TBSomaEmail3.EOF = False Then
                        Do While TBSomaEmail3.EOF = False
                            If IsNull(TBSomaEmail3!Email) = False And TBSomaEmail3!Email <> "" Then
                                If funcSomaEmailNF <> "" Then funcSomaEmailNF = funcSomaEmailNF & ";" & TBSomaEmail3!Email Else funcSomaEmailNF = TBSomaEmail3!Email
                            End If
                            TBSomaEmail3.MoveNext
                        Loop
                    End If
                    TBSomaEmail3.Close
                End If
            End If
        End If
    End If
End If
TBSomaEmail.Close

Exit Function
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Function

Public Sub ProcEnviarEmailTecnoSpeed(emailEnviar As String, tipoEmail As Integer, mensagemEmail As Boolean, numeroNF_email As String, chaveAcesso_email As String, serieNF_email As String)
On Error GoTo tratar_erro

'1 - Envio normal
'2 - Cancelamento
'3 - Carta de correÁ„o

If spdNFe.EmailRemetente <> "" Then
    If emailEnviar <> "" And chaveAcesso_email <> "" Then
        spdNFe.EmailDestinatario = emailEnviar
        If tipoEmail = 1 Then
            spdNFe.EmailMensagem = "Segue em anexo DANFE e XML referente a nota fiscal: " & numeroNF_email & " Serie: " & serieNF_email & " Chave de acesso: " & chaveAcesso_email
            spdNFe.EnviarNotaDestinatario chaveAcesso_email, "", ""
        ElseIf tipoEmail = 2 Then
            ArquivoXML = spdNFe.DiretorioXmlDestinatario + chaveAcesso_email + "-caneve.XML"
            ArquivoDanfe = spdNFe.DiretorioXmlDestinatario + "PDF\" + chaveAcesso_email + "-canc.pdf"
            spdNFe.EmailMensagem = "A nota fiscal: " & numeroNF_email & " Serie: " & serieNF_email & " Chave de acesso: " & chaveAcesso_email & " foi cancelada."
            spdNFe.EnviarNotaDestinatarioAnexos ArquivoDanfe, ArquivoXML, ""
        Else
            spdNFe.EmailMensagem = "Foi registrada a seguinte correÁ„o para a sua nota fiscal: " & numeroNF_email & " Serie: " & serieNF_email & " Chave de acesso: " & chaveAcesso_email
            spdNFe.EnviarCCeDestinatario chaveAcesso_email
        End If
        If mensagemEmail = True Then USMsgBox "E-mail enviado com sucesso.", vbInformation, "CAPRIND v5.0"
    End If
Else
    If mensagemEmail = True Then USMsgBox "N„o foi possÌvel enviar o e-mail, pois o mesmo n„o foi configurado.", vbExclamation, "CAPRIND v5.0"
End If

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Public Sub procExcluirDevolucaoNF(ID_nota As Long, ID_nota_relacionada As Long)
On Error GoTo tratar_erro
Dim TBDevRel As ADODB.Recordset

Set TBDevRel = CreateObject("adodb.recordset")
TBDevRel.Open "Select * from Faturamento_Relacionamento where (ID_nota <> " & ID_nota & " and ID_nota_relacionada = " & ID_nota_relacionada & " Or ID_nota = " & ID_nota_relacionada & " And ID_nota_relacionada <> " & ID_nota & ") and qtde < 0", Conexao, adOpenKeyset, adLockOptimistic
Do While TBDevRel.EOF = False
    
    If TBDevRel!ID_nota = ID_nota_relacionada Then
        id_produto_entrada = TBDevRel!id_produto_relacionada
    Else
        id_produto_entrada = TBDevRel!ID_Produto
    End If
    
    Conexao.Execute "UPDATE tbl_Detalhes_Nota SET Saldo = saldo + " & Format(TBDevRel!Qtde, 2) & " WHERE Int_codigo = " & id_produto_entrada
    TBDevRel.Delete
    TBDevRel.MoveNext
Loop
TBDevRel.Close

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Sub ProcExcluirContas(ID_nota As Long, Saida As Boolean, TipoCliente As String)
On Error GoTo tratar_erro

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select CFOP.* from tbl_Detalhes_Nota NFP INNER JOIN tbl_NaturezaOperacao CFOP ON CFOP.IDCountCfop = NFP.ID_CFOP where NFP.ID_nota = " & ID_nota & " and CFOP.Devolucao = 'True'", Conexao, adOpenKeyset, adLockReadOnly
If TBAbrir.EOF = False Then
    'Fornecedor
    If Saida = True And Len(TipoCliente) = 1 Then GoTo ExcluirPagar
    'Cliente
    If Saida = False And Len(TipoCliente) = 2 Then GoTo ExcluirReceber
Else
    If Saida = True Then GoTo ExcluirReceber Else GoTo ExcluirPagar
End If
TBAbrir.Close

ExcluirReceber:
    Set TBContas = CreateObject("adodb.recordset")
    TBContas.Open "Select * from tbl_contas_Receber where id_nota = " & ID_nota & " and Bloqueado = 'False' and Status <> 'DUPLICATA DESCONTADA RECOMPRADA'", Conexao, adOpenKeyset, adLockOptimistic
    If TBContas.EOF = False Then
        Do While TBContas.EOF = False
            Conexao.Execute "DELETE from CC_realizado where Operacao = 'CrÈdito' and ID_financeiro = " & TBContas!IDintconta
            
            If (IsNull(TBContas!Proposta) = True Or TBContas!Proposta = "") And TBContas!Logsit = "N" Then
                'Contas contabeis
                Conexao.Execute "DELETE FROM Familia_financeiro WHERE IDConta = " & TBContas!IDintconta & " and Tipoconta = 'R' and Deposito_transf = 'False'"
                'Fluxo de Caixa
                Conexao.Execute "DELETE from tbl_Fluxo_de_caixa where IDFluxo = " & IIf(IsNull(TBContas!IDFluxo), 0, TBContas!IDFluxo)
                'N˙mero dos boletos
                Conexao.Execute "DELETE from tbl_Detalhes_Recebimento_Nboletos where IDContaReceber = " & TBContas!IDintconta
                'Conta
                Conexao.Execute "DELETE from tbl_contas_Receber where IDintconta = " & TBContas!IDintconta
            ElseIf IsNull(TBContas!Proposta) = False And TBContas!Proposta <> "" Then
                    TBContas!ID_nota = 0
                    TBContas!NFiscal = ""
                    TBContas.Update
            End If
            TBContas.MoveNext
        Loop
    End If
    TBContas.Close
    GoTo Prosseguir

ExcluirPagar:
    Set TBContas = CreateObject("adodb.recordset")
    TBContas.Open "Select * from tbl_ContasPagar where id_nota = " & ID_nota & " and Bloqueado = 'False' and Despesas_NF = 'False'", Conexao, adOpenKeyset, adLockOptimistic
    If TBContas.EOF = False Then
        Do While TBContas.EOF = False
            If (IsNull(TBContas!Txt_pedido) = True Or TBContas!Txt_pedido = "") And TBContas!Logsit = "N" Then
                'Contas contabeis
                Conexao.Execute "DELETE FROM Familia_financeiro WHERE IDConta = " & TBContas!IDintconta & " and Tipoconta = 'P' and Deposito_transf = 'False'"
                'Fluxo de Caixa
                Conexao.Execute "DELETE from tbl_Fluxo_de_caixa where IDFluxo = " & IIf(IsNull(TBContas!IDFluxo), 0, TBContas!IDFluxo)
                'Conta
                Conexao.Execute "DELETE from tbl_ContasPagar where IDintconta = " & TBContas!IDintconta
            ElseIf IsNull(TBContas!Txt_pedido) = False And TBContas!Txt_pedido <> "" Then
                    TBContas!ID_nota = 0
                    TBContas!txt_ndocumento = ""
                    TBContas.Update
            End If
            TBContas.MoveNext
        Loop
    End If
    TBContas.Close

Prosseguir:
    Conexao.Execute "Update CC set CC.ID_Financeiro = 0 from CC_realizado CC INNER JOIN tbl_Detalhes_Recebimento TBL on CC.ID_duplicata = TBL.ID where TBL.ID_nota = " & ID_nota

Exit Sub
tratar_erro:
    USMsgBox ("DescriÁ„o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub
