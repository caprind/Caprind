VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.Controls.v19.3.0.ocx"
Begin VB.Form frmEmpresa_Contrato 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Empresa - Contrato_Contrato"
   ClientHeight    =   10035
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15360
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10035
   ScaleMode       =   0  'User
   ScaleWidth      =   15480.95
   WindowState     =   2  'Maximized
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   26
      ScreenHeight    =   1080
      ScreenWidth     =   1920
      ScreenHeightDT  =   1080
      ScreenWidthDT   =   1920
      AutoResizeOnLoad=   0   'False
      ApplicationName =   "Active Resize Control Professional"
      FormHeightDT    =   10500
      FormWidthDT     =   15480
      FormScaleHeightDT=   10035
      FormScaleWidthDT=   15481
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin RichTextLib.RichTextBox txtContrato 
      Height          =   7995
      Left            =   390
      TabIndex        =   0
      Top             =   1590
      Width           =   14625
      _ExtentX        =   25797
      _ExtentY        =   14102
      _Version        =   393217
      BackColor       =   16777215
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmEmpresa_Contrato.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin DrawSuite2022.USButton btnImprimir 
      Height          =   615
      Left            =   12960
      TabIndex        =   1
      Top             =   570
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   1085
      DibPicture      =   "frmEmpresa_Contrato.frx":0080
      Caption         =   "Imprimir"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColorDown =   5249536
      BorderColorOver =   8076800
      PicAlign        =   8
      ShowFocusRect   =   0   'False
   End
   Begin DrawSuite2022.USButton btnSair 
      Height          =   615
      Left            =   14040
      TabIndex        =   2
      Top             =   570
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   1085
      DibPicture      =   "frmEmpresa_Contrato.frx":1D15
      Caption         =   "Sair"
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
      BorderColorDown =   2039646
      BorderColorOver =   3026574
      GradientColor1  =   5263559
      GradientColor2  =   5263559
      GradientColor3  =   5263559
      GradientColor4  =   5263559
      GradientColorOver1=   3026574
      GradientColorOver2=   3026574
      GradientColorOver3=   3026574
      GradientColorOver4=   3026574
      GradientColorDown1=   2039646
      GradientColorDown2=   2039646
      GradientColorDown3=   2039646
      GradientColorDown4=   2039646
      PicAlign        =   8
      ShowFocusRect   =   0   'False
      Theme           =   4
   End
   Begin DrawSuite2022.USGroupBox USGroupBox1 
      Height          =   9915
      Left            =   150
      TabIndex        =   3
      Top             =   90
      Width           =   15105
      _ExtentX        =   26644
      _ExtentY        =   17489
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin DrawSuite2022.USAlphaImage USAlphaImage1 
         Height          =   720
         Left            =   630
         TabIndex        =   5
         Top             =   480
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1270
         Image           =   "frmEmpresa_Contrato.frx":334FE
         Props           =   5
      End
      Begin XtremeSuiteControls.Label lblContrato 
         Height          =   285
         Left            =   1500
         TabIndex        =   4
         Top             =   720
         Width           =   7800
         _Version        =   1245187
         _ExtentX        =   13758
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "CONTRATO DE LOCAÇÃO DE LICENÇAS SUPORTE E MANUTENÇÃO Nº"
         ForeColor       =   4210752
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoSize        =   -1  'True
      End
   End
End
Attribute VB_Name = "frmEmpresa_Contrato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnImprimir_Click()
On Error GoTo tratar_erro

NomeRel = "Contrato.rpt"
FormulaRel = "{Clientes.Cnpj} = '" & CNPJCliente & "'"  '41.030.779/0001-75'
ProcImprimirContrato (FormulaRel)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimirContrato(FormulaRel As String)
On Error GoTo tratar_erro

Set Report = crAPP.OpenReport(Localrel & "\" & NomeRel)
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
        USMsgBox ("Não foi encontrado o relatório " & NomeRel & " na pasta " & Localrel), vbExclamation, "CAPRIND v5.0"
        GoTo 2
    End If
    If Err.Number = "-2147483638" Then
        USMsgBox ("Não foi possível visualizar o relatório, favor reiniciar o sistema."), vbExclamation, "CAPRIND v5.0"
        GoTo 2
    End If
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub btnSair_Click()
On Error GoTo tratar_erro

Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro
Dim Contrato As String


If TemInternet = True And ErroDriverMYSQL = False Then
Set TBAfericao = CreateObject("adodb.recordset")
TBAfericao.Open "Select * from Empresa where CNPJ = '" & frmOpcoesGeral.txtcnpj.Text & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBAfericao.EOF = False Then
FunAbreBDSite
If ConexaoMySql.State = 1 Then
Set TBMySQL = New ADODB.Recordset

StrSql = "select * from Clientes where CNPJ = '" & TBAfericao!CNPJ & "'"

TBMySQL.Open StrSql, ConexaoMySql, adOpenKeyset, adLockOptimistic, adCmdText
If TBMySQL.EOF = False Then
CNPJCliente = TBMySQL!CNPJ
lblContrato.Caption = "CONTRATO DE LOCAÇÃO DE LICENÇAS SUPORTE E MANUTENÇÃO Nº 00" & TBMySQL!ID & "/" & Year(TBMySQL!Data)
            
Contrato = _
"1. PARTES CONTRATANTES" & vbCrLf & _
"São partes neste CONTRATO DE LOCAÇÃO DE LICENÇAS, SUPORTE E MANUTENÇÃO, na qualidade de PRESTADOR DE SERVIÇOS, a FNL TECNOLOGIA -" & _
" CAPRIND SISTEMAS  inscrita no CNPJ sob o Nº 34.270.461/0001-04  e na qualidade de CLIENTE, a " & TBMySQL!NomeRazao & ", inscrita no CNPJ sob o Nº " & TBMySQL!CNPJ & ""

Contrato = Contrato & vbCrLf & vbCrLf & _
"2. OBJETO" & vbCrLf & _
" Prestação de serviços em locação de licenças, manutenção e suporte técnico aos produtos(s) assinalados no presente contrato."

Contrato = Contrato & vbCrLf & vbCrLf & _
"3. RELAÇÃO DE SERVIÇOS CONTRATADOS : LOCAÇÃO DE LICENÇAS, MANUTENÇÃO  E SUPORTE TÉCNICO" & vbCrLf & _
" Suporte técnico através de:" & vbCrLf & _
"   1 - Chat online (Sistema e Portal Caprind)" & vbCrLf & _
"   2 - TeamViewer (Conexão remota)" & vbCrLf & _
"   3 - Email (suporte@caprind.com.br)" & vbCrLf & _
"   4 - Telefone no horário comercial" & vbCrLf & _
"   5 - Abertura de chamado (SAS - Sistema)" & vbCrLf & _
"   6 - Update (Correção de erros)" & vbCrLf & _
"   7 - Upgrade (Atualização com melhorias e novos recursos)" & vbCrLf & _
" Manutenção do banco de dados (caso se faça necessário pela atualização)" & vbCrLf & _
" Locação de licença(s)"

Contrato = Contrato & vbCrLf & vbCrLf & _
"4. PRODUTO(S) E SERVIÇOS COBERTO(S) PELO CONTRATO" & vbCrLf & _
"    " & TBMySQL!Licencas & " licença(s) CAPRIND FULL GESTÃO INDUSTRIAL" & vbCrLf & _
"    " & TBMySQL!Licencas_gerprod & " licença(s) GERPROD COLETOR DE DADOS E APONTAMENTO DE PRODUÇÃO"

Contrato = Contrato & vbCrLf & vbCrLf & _
"5. RECURSOS E MÓDULOS DO SISTEMA" & vbCrLf & _
" O sistema CAPRIND é um sistema de gestão industrial, e como tal destina-se a organizar as informações inerentes a gestão da empresa de forma integrada." & vbCrLf & _
" Conta com módulos de gerenciamento da empresa que permitem ao usuário manipular as informações de forma online atendendo desde a formação do  preço de venda até a expedição final dos produtos." & vbCrLf & _
vbCrLf & " RELAÇÃO DE MÓDULOS DISPONÍVEIS NO SISTEMA:" & vbCrLf & _
"  ADMINISTRATIVO" & vbCrLf & _
"   RH"

Contrato = Contrato & vbCrLf & _
"    Cadastro de funcionários" & vbCrLf & _
"    Relatórios"

Contrato = Contrato & vbCrLf & vbCrLf & _
"   COMPRAS" & vbCrLf & _
"    Famílias" & vbCrLf & _
"    Produtos e serviços" & vbCrLf & _
"    Fornecedores" & vbCrLf & _
"    Programação" & vbCrLf & _
"    Cotação" & vbCrLf & _
"    Pedido" & vbCrLf & _
"    Necessidade" & vbCrLf & _
"    Não conformidade" & vbCrLf & _
"    Atualização de valores" & vbCrLf & _
"    Relatórios"

Contrato = Contrato & vbCrLf & vbCrLf & _
"   VENDAS" & vbCrLf & _
"    Famílias" & vbCrLf & _
"    Produtos e serviços" & vbCrLf & _
"    Clientes" & vbCrLf & _
"    Vendedores" & vbCrLf & _
"    Telemarketing" & vbCrLf & _
"    Empenho" & vbCrLf & _
"    Programação" & vbCrLf & _
"    Proposta Comercial" & vbCrLf & _
"    Pedido interno" & vbCrLf & _
"    Follow up" & vbCrLf & _
"    Situação da produção" & vbCrLf & _
"    Informações de faturamento" & vbCrLf & _
"    Atualização de valores" & vbCrLf & _
"    Relatórios"

Contrato = Contrato & vbCrLf & vbCrLf & _
"   FINANCEIRO" & vbCrLf & _
"    Plano de contas" & vbCrLf & _
"    Instituições" & vbCrLf & _
"    Contas a pagar" & vbCrLf & _
"    Contas pagas" & vbCrLf & _
"    Contas a receber" & vbCrLf & _
"    Contas recebidas" & vbCrLf & _
"    Desconto de duplicatas" & vbCrLf & _
"    Fluxo de caixa" & vbCrLf & _
"    Relatórios"

Contrato = Contrato & vbCrLf & vbCrLf & _
"   FATURAMENTO" & vbCrLf & _
"    Fiscal" & vbCrLf & _
"    Nota fiscal" & vbCrLf & _
"    Carta de correção" & vbCrLf & _
"    Minuta de despacho" & vbCrLf & _
"    Relatórios"

Contrato = Contrato & vbCrLf & vbCrLf & _
"   CUSTOS" & vbCrLf & _
"    Centro de custo" & vbCrLf & _
"    Relatórios"

Contrato = Contrato & vbCrLf & vbCrLf & _
"   ENGENHARIA" & vbCrLf & _
"    Carteira de pedidos" & vbCrLf & _
"    Famílias" & vbCrLf & _
"    Produtos e serviços" & vbCrLf & _
"    Conjuntos" & vbCrLf & _
"    Estrutura" & vbCrLf & _
"    Controle de projetos" & vbCrLf & _
"    Processos" & vbCrLf & _
"    Normas"

Contrato = Contrato & vbCrLf & vbCrLf & _
"   PPCP" & vbCrLf & _
"    Postos de trabalho" & vbCrLf & _
"    Códigos de trabalho" & vbCrLf & _
"    Carga posto de trabalho" & vbCrLf & _
"    Gerenciamento de ordem" & vbCrLf & _
"    Monitor de trabalho" & vbCrLf & _
"    Situação da produção" & vbCrLf & _
"    Programas CNC" & vbCrLf & _
"    Necessidade" & vbCrLf & _
"    Não conformidades" & vbCrLf & _
"    Programação da produção" & vbCrLf & _
"    Plano de apontamento" & vbCrLf & _
"    Relatórios"

Contrato = Contrato & vbCrLf & vbCrLf & _
"   QUALIDADE" & vbCrLf & _
"    Famílias" & vbCrLf & _
"    Instrumentos" & vbCrLf & _
"    Almoxarifado" & vbCrLf & _
"    Plano de inspeção" & vbCrLf & _
"    Controle de medição" & vbCrLf & _
"    Inspeção de recebimento" & vbCrLf & _
"    Ensaios" & vbCrLf & _
"    Controle de certificados" & vbCrLf & _
"    Controle de documentos e dados" & vbCrLf & _
"    Não conformidades" & vbCrLf & _
"    Solicitação de ação" & vbCrLf & _
"    Solicitação de desvio" & vbCrLf & _
"    RNC" & vbCrLf & _
"    PPAP" & vbCrLf & _
"    Histórico de revisão dos relatórios" & vbCrLf & _
"    Relatórios"

Contrato = Contrato & vbCrLf & vbCrLf & _
"   ESTOQUE" & vbCrLf & _
"    Almoxarifado" & vbCrLf & _
"    Local de armazenamento" & vbCrLf & _
"    Requisição de materiais" & vbCrLf & _
"    Recebimento" & vbCrLf & _
"    Inventário" & vbCrLf & _
"    Movimentação" & vbCrLf & _
"    Necessidade" & vbCrLf & _
"    Ordem de faturamento" & vbCrLf & _
"    Nota fiscal"

Contrato = Contrato & vbCrLf & vbCrLf & _
"   MANUTENÇÃO" & vbCrLf & _
"    Equipamentos" & vbCrLf & _
"    Relatórios"

Contrato = Contrato & vbCrLf & vbCrLf & _
"   OUTROS" & vbCrLf & _
"    Solicitação" & vbCrLf & _
"    Follow up de compras" & vbCrLf & _
"    Validação dos procedimentos" & vbCrLf & _
"    Análise crítica"

Contrato = Contrato & vbCrLf & vbCrLf & _
"   SUPORTE" & vbCrLf & _
"    Chamado" & vbCrLf & _
"    Chat (Online)" & vbCrLf & _
"    Atualização"

Contrato = Contrato & vbCrLf & vbCrLf & _
"6. OBRIGAÇÕES DA CAPRIND SISTEMAS" & vbCrLf & _
" 6.1. Prestar suporte técnico ao CLIENTE, conforme relação descrita no item três deste contrato, solucionando dúvidas quanto à instalação, configuração e utilização do PRODUTO, realizando quando necessário à manutenção corretiva e preventiva do banco de dados do cliente no horário de nove às dezessete horas, de segunda a sexta feira, exceto nos feriados." & vbCrLf & _
" 6.2. Manter a atualização do sistema disponível ao cliente através do PORTAL CAPRIND em forma de download dos arquivos." & vbCrLf & _
" 6.3. S.A.S (Solicitação de atendimento ao sistema), trata-se das formas de atendimento ao usuários disponibilizadas ao usuário do sistema." & vbCrLf & _
"  6.3.1. A CAPRIND SISTEMAS atenderá a todas as solicitações do CLIENTE desde que enviadas por uma das formas de atendimento válidas (S.A.S) e dentro dos limites contratados, sendo que as formas de atendimento válidas são:" & vbCrLf & _
"   6.3.1.1. Abertura de chamado através do sistema CAPRIND." & vbCrLf & _
"    6.3.1.1.1. Somente deverá ser utilizado por usuários comprovadamente treinados para reportar a CAPRIND SISTEMAS erros e falhas do sistema, ou para solicitação de desenvolvimento de recursos novos e relatórios." & vbCrLf & _
"    6.3.1.1.2. Tempo de resposta: De 01 á 24hs" & vbCrLf & _
"   6.3.1.2. Atendimento via chat online através do sistema CAPRIND ou através do Portal CAPRIND." & vbCrLf & _
"    6.3.1.2.1. Somente deverá ser utilizado por usuários comprovadamente treinados para esclarecimento de dúvidas." & vbCrLf & _
"    6.3.1.2.2. Tempo de reposta: De 10 minutos á 1h" & vbCrLf & _
"   6.3.1.3 .Envio de e-mail ao departamento de suporte ao sistema CAPRIND." & vbCrLf & _
"    6.3.1.3.1. Somente deverá ser utilizado pelo administrador do sistema ou por pessoas autorizadas pela CAPRIND SISTEMAS." & vbCrLf & _
"    6.3.1.3.2. Tempo de resposta: De um a cinco dias uteis" & vbCrLf & _
"  6.3.2. Em todos os casos acima, será retornada ao mesmo a solução, a previsão de solução ou a justificativa da impossibilidade de solução do problema;" & vbCrLf & _
"   6.3.2.1. Se necessário será deslocado um técnico da CAPRIND SISTEMAS as instalações do cliente para solução do problema detectado." & vbCrLf & _
" 6.4. Fornecer ao licenciado condições de utilização plena do sistema através de:" & vbCrLf & _
"  6.4.1 .Consultoria de implantação e treinamentos:" & vbCrLf & _
"   6.4.1.1. Fornecido através de aquisição de banco de horas de consultoria técnica prestada nas instalações do cliente." & vbCrLf & _
"    6.4.1.1.1. A consultoria técnica pode ser adquirida através da compra do banco de horas previamente orçadas, e agendadas no cliente conforme agenda do técnico e a necessidade do licenciado, e pode ser utilizado para treinamentos, consultoria técnica, esclarecimento de duvidas, e coleta de informações técnicas para desenvolvimento de novos recursos." & vbCrLf & _
"  6.4.2. Treinamentos:" & vbCrLf & _
"   6.4.2.1. Material de auxílio fornecido através de Tutoriais em vídeo postados no Portal CAPRIND."

Contrato = Contrato & vbCrLf & vbCrLf & _
"7. OBRIGAÇÕES DO CLIENTE" & vbCrLf & _
" 7.1. Efetuar os pagamentos dos valores ajustados no contrato, de acordo com datas e condições estabelecidas na cláusula nove;" & vbCrLf & _
" 7.2. Informar à CAPRIND SISTEMAS qualquer alteração cadastral (endereço, telefone, fax, e-mail, etc.);" & vbCrLf & _
" 7.3. Manter este contrato em local apropriado, e ceder a CAPRIND SISTEMAS todas as vezes que solicitado, devendo a CAPRIND SISTEMAS devolver imediatamente ao cliente após analise." & vbCrLf & _
" 7.4. Manter seus funcionários treinados e atualizados quanto às funcionalidades e novidades do sistema." & vbCrLf & _
" 7.5. Indicar um funcionário para atuar como gestor do sistema e informar a CAPRIND SISTEMAS quando esse funcionário for substituido."

Contrato = Contrato & vbCrLf & vbCrLf & _
"8. DISPOSIÇÕES GERAIS DO ATENDIMENTO AO CLIENTE" & vbCrLf & _
" 8.1. O serviço deverá ser solicitado através das formas de atendimento válidas para este contrato (Cláusula 6.1)." & vbCrLf & _
" 8.2. Quando for necessário o deslocamento de um profissional da CAPRIND SISTEMAS e/ou de um representante até as instalações do CLIENTE, serão observadas as seguintes condições:" & vbCrLf & _
"  8.2.1.1. O serviço será prestado sempre no horário comercial, nos dias úteis com horário previamente agendado;" & vbCrLf & _
"  8.2.1.2. O atendimento será registrado em Relatório de Visita assinado pelas partes;" & vbCrLf & _
" 8.3. O Atendimento de Suporte (manutenção) está restrito ao funcionamento do(s) PRODUTO(S). Este Contrato não contempla ocorrências referentes a problemas não inerentes ao(s) produto(s)." & vbCrLf & _
"      Nos casos onde a equipe técnica da CAPRIND SISTEMAS identificar que uma ocorrência fuja ao escopo do(s) PRODUTO(S), o CLIENTE será comunicado do fato e, caso haja interesse de ambas as partes, a solução  será prestada na forma de consultoria segundo valores vigentes à época;" & vbCrLf & _
" 8.4. Desenvolvimentos e/ou customizações específicas para o CLIENTE, não fazem parte deste contrato." & vbCrLf & _
"      Entretanto, poderão ser orçadas pela CAPRIND SISTEMAS e, mediante aprovação prévia do orçamento pelo CLIENTE, serem implementadas;"

Contrato = Contrato & vbCrLf & vbCrLf & _
"9. VALORES E CONDIÇÕES DE PAGAMENTO" & vbCrLf & _
" 9.1. Pelos serviços, objeto deste contrato, o CLIENTE pagará um valor mensal de  R$ " & Format(TBMySQL!Total_Contrato, "###,##0.00") & " em forma de boleto bancário referente aos valores informados." & vbCrLf & _
" 9.2. O pagamento deverá ser efetuado através de cobrança bancária (Boleto);" & vbCrLf & _
" 9.3. O débito será efetuado no mês seguinte ao de prestação do serviço."

Contrato = Contrato & vbCrLf & vbCrLf & _
"10. PENALIDADES" & vbCrLf & _
" 10.1.  A falta de pagamento dos valores convencionados, nas datas de seus respectivos vencimentos, acarretará multa, meramente penal, de 10% (dez por cento) do valor total do débito, acrescido de juros moratórias de 0,2% ao dia de correção monetária referente ao período de atraso;" & vbCrLf & _
" 10.2. No caso de inadimplência superior a 15 (QUINZE) dias facultará à CAPRIND SISTEMAS o direito de suspender os serviços descritos neste contrato."

Select Case Month(TBMySQL!Data)
    Case 1: Mes = "Janeiro"
    Case 2: Mes = "Fevereiro"
    Case 3: Mes = "Março"
    Case 4: Mes = "Abril"
    Case 5: Mes = "Maio"
    Case 6: Mes = "Junho"
    Case 7: Mes = "Julho"
    Case 8: Mes = "Agosto"
    Case 9: Mes = "Setembro"
    Case 10: Mes = "Outubro"
    Case 11: Mes = "Novembro"
    Case 12: Mes = "Dezembro"
End Select

Contrato = Contrato & vbCrLf & vbCrLf & _
"11. REAJUSTE" & vbCrLf & _
" 11.1. Anualmente exatamente no mês de " & Mes & ", o valor mencionado na cláusula 9.1 deste contrato será reajustado de acordo com a variação do IGPM/FGV ou INPC. Em sua falta, será adotado outro índice legalmente admitido."

Contrato = Contrato & vbCrLf & vbCrLf & _
"12. VIGÊNCIA, RESCISÃO E MULTA." & vbCrLf & _
" 12.1.Vigência:" & vbCrLf & _
"  12.1.1.Inicia na data da sua assinatura, e vigorará pelo prazo certo e determinado de 12 (doze) meses, sendo renovado automaticamente por períodos iguais e sucessivos, caso as partes não se manifestem em contrário, por escrito, com antecedência mínima de 30 dias da data do término do prazo contratual ou de cada renovação;" & vbCrLf & _
" 12.2.Rescisão:" & vbCrLf & _
"  12.2.1.Este contrato poderá ser rescindido, a qualquer momento, pela simples vontade de qualquer dos contratantes, manifestada com 30 (trinta) dias de antecedência, através de documento" & vbCrLf & _
"         escrito e comprovadamente entregue à outra parte." & vbCrLf & _
"  12.2.2.No caso da rescisão do contrato, as licenças do sistema serão desinstaladas, e o banco de dados com todas as informações nele contidas permanecerão em poder do cliente a disposição para utilização." & vbCrLf & _
" 12.3.Multa:" & vbCrLf & _
"  12.3.1. NÃO EXISTE QUALQUER TIPO DE COBRANÇA PARA RESCISÃO DO CONTRATO EXCETO ESTAR COM A(s) MENSALIDADE(s) EM DIA."

Contrato = Contrato & vbCrLf & vbCrLf & _
"13. RESPONSABILIDADE DE FUNCIONAMENTO" & vbCrLf & _
" 13.1. A CAPRIND SISTEMAS somente será responsável por qualquer dano direto ou indireto, lucro cessante, interrupção de negócios, perda de informações, decorrentes do mau funcionamento do sistema, quando for identificado por um técnico da própria CAPRIND que o defeito é do próprio sistema, e que não se trata de nenhum dos motivos relacionados abaixo:" & vbCrLf & _
"  13.1.1.Perda de conexão com o banco de dados SQL Server." & vbCrLf & _
"  13.1.2.Perda de conexão com a rede interna." & vbCrLf & _
"  13.1.3.Defeito do equipamento onde está instalado o sistema" & vbCrLf & _
"  13.1.4.Defeito do servidor de dados (SQL Server)." & vbCrLf & _
"  13.1.5.Perda de sinal de internet." & vbCrLf & _
"  13.1.6.Sistema fora da atualização válida." & vbCrLf & _
"  13.1.7.Banco de dados fora da atualização válida." & vbCrLf & _
"  13.1.8.Defeito causado por uma atualização do Windows." & vbCrLf & _
"  13.1.9.Equipamento não adequado para utilização do sistema." & vbCrLf & _
"  13.1.10.Falta de treinamento adequado do usuário." & vbCrLf & _
"  13.1.11.Defeito causado por uma instalação de aplicativo, posterior a instalação do sistema." & vbCrLf & _
"  13.1.12.Defeito causado por uma atualização do sistema com a implementação de novos recursos." & vbCrLf & _
"  13.1.13.Outras situações identificadas pelo técnico da Caprind."

Contrato = Contrato & vbCrLf & vbCrLf & _
"14. SIGILO DE INFORMAÇÕES" & vbCrLf & _
" 14.1. A CAPRIND SISTEMAS assume inteira responsabilidade pelo sigilo das informações contidas no banco de dados, e se compromete a não divulgar, ceder, copiar ou usar de qualquer outra forma para tornar público ou compartilhar essas informações."

Contrato = Contrato & vbCrLf & vbCrLf & _
"15. OBSERVAÇÕES" & vbCrLf & _
" 15.1.O numero de licenças locadas poderá ser alterado a qualquer tempo dentro do próprio sistema em módulo próprio pelo administrador." & vbCrLf & _
"      Observações: A mudança no numero de licenças alterará o valor mensal de locação." & vbCrLf & _
" 15.2.Todas as funcionalidades e recursos do sistema foram apresentados ao cliente, não deixando nenhuma dúvida quanto a sua utilidade e recursos disponíveis."

Contrato = Contrato & vbCrLf & vbCrLf & _
"16. DO FORO" & vbCrLf & _
" 16.1.As partes elegem o foro de Indaiatuba/SP, para dirimir quaisquer dúvidas decorrentes deste instrumento, renunciando a qualquer outro por mais privilegiado que possa parecer."

Contrato = Contrato & vbCrLf & vbCrLf & _
"E por estarem assim justas e contratadas, firmam o presente instrumento em 02 (duas) vias de igual forma e teor, para um só fim, obrigando-se por si e/ou seus sucessores a fielmente cumpri-lo em todas as suas disposições."

Contrato = Contrato & vbCrLf & vbCrLf & _
"Atenção: Esse contrato somente terá validade após o reconhecimento de firma em cartório."
End If
TBMySQL.Close
End If

End If
TBAfericao.Close
End If

txtContrato.Text = Contrato
frmEmpresa_Contrato.Refresh
   
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

