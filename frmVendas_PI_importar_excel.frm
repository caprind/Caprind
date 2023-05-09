VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmVendas_PI_importar_excel 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Vendas - Pedido interno - Importar pedido de compra do excel"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8880
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   8880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5220
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   4
      Top             =   0
      Width           =   8775
      _ExtentX        =   15478
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
      ButtonCaption1  =   "Importar"
      ButtonEnabled1  =   0   'False
      ButtonIconSize1 =   32
      ButtonToolTipText1=   "Importar (F3)"
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
      ButtonWidth1    =   50
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
      ButtonLeft2     =   54
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
      ButtonLeft3     =   58
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
      ButtonLeft4     =   96
      ButtonTop4      =   2
      ButtonWidth4    =   26
      ButtonHeight4   =   21
      ButtonUseMaskColor4=   0   'False
      ButtonEnabled5  =   0   'False
      ButtonIconSize5 =   32
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
      ButtonState5    =   5
      ButtonLeft5     =   124
      ButtonTop5      =   2
      ButtonWidth5    =   24
      ButtonHeight5   =   24
      ButtonUseMaskColor5=   0   'False
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   5880
         Top             =   180
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmVendas_PI_importar_excel.frx":0000
         Count           =   1
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   2085
      Left            =   55
      TabIndex        =   5
      Top             =   990
      Width           =   8780
      Begin VB.ComboBox Cmb_familia_material 
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
         Height          =   330
         Left            =   180
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "Família para cadastrar o material consignado."
         Top             =   1620
         Width           =   8415
      End
      Begin VB.ComboBox Cmb_familia_produto 
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
         Height          =   330
         Left            =   180
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "Família para cadastrar o produto final."
         Top             =   990
         Width           =   8415
      End
      Begin VB.TextBox Txt_caminho_arquivo 
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
         MaxLength       =   255
         TabIndex        =   0
         TabStop         =   0   'False
         ToolTipText     =   "Anexo."
         Top             =   390
         Width           =   8055
      End
      Begin VB.CommandButton Cmd_anexo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   8250
         Picture         =   "frmVendas_PI_importar_excel.frx":23C3
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Localizar arquivo."
         Top             =   390
         Width           =   315
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Família para cadastrar o material consignado"
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
         Left            =   2790
         TabIndex        =   9
         Top             =   1410
         Width           =   3195
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Família para cadastrar o produto final"
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
         Left            =   3045
         TabIndex        =   8
         Top             =   780
         Width           =   2685
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Caminho do arquivo excel"
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
         Left            =   3285
         TabIndex        =   6
         Top             =   180
         Width           =   1845
      End
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   60
      TabIndex        =   7
      Top             =   3090
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   450
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
Attribute VB_Name = "frmVendas_PI_importar_excel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private myXML As DOMDocument50
Private x As IXMLDOMNode
Dim xl As New Excel.Application
Dim xlw As Excel.Workbook

Private Sub Cmd_anexo_Click()
On Error GoTo tratar_erro

ProcCarregaCaminhoNomeArquivo CommonDialog1, "*.xls", "(*.xls) | *.xls| (*.xlsx) | *.xlsx"
Txt_caminho_arquivo = caminho

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyF3: procImportar
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

ProcCarregaToolBar1 Me, 8775, 5, True
ProcLimpaVariaveisPrincipais
ProcCarregaComboFamilia Cmb_familia_produto, "Vendas = 'True'", False
ProcCarregaComboFamilia Cmb_familia_material, "Vendas = 'True'", True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procImportar()
On Error GoTo tratar_erro

Acao = "importar"
If Txt_caminho_arquivo = "" Then
    NomeCampo = "o caminho do arquivo"
    ProcVerificaAcao
    Txt_caminho_arquivo.SetFocus
    Exit Sub
End If
If Cmb_familia_produto = "" Then
    NomeCampo = "a família para cadastrar o produto final"
    ProcVerificaAcao
    Cmb_familia_produto.SetFocus
    Exit Sub
End If

Nome_anexo1 = Replace(Left(Nome_anexo, Len(Nome_anexo) - 4), "-", "/")
Permitido2 = True
'Abrir o arquivo do Excel
Set xlw = xl.Workbooks.Open(Txt_caminho_arquivo)
'Definir qual a planilha de trabalho
NomePlanilha = "Plan1"
AbrePlanilha:
    xlw.Sheets(NomePlanilha).Select
    
    With xlw.Application
        If .Cells(1, 1) <> "INÍCIO" Then
            USMsgBox ("Este arquivo não é válido."), vbExclamation, "CAPRIND v5.0"
            Permitido2 = False
            GoTo Encerrar
        End If
        
        frmVendas_PI.RegimeEmpresa_PI = FunVerifRegimeEmpresa(frmVendas_PI.Cmb_empresa.ItemData(frmVendas_PI.Cmb_empresa.ListIndex))
        If frmVendas_PI.RegimeEmpresa_PI = 1 Then
            'Verifica se existe mais de uma tabela do simples cadastrada
            contador = 0
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select Tabela FROM Impostos_TabelaDAS where ID_empresa = " & frmVendas_PI.Cmb_empresa.ItemData(frmVendas_PI.Cmb_empresa.ListIndex) & " and Ativado = 1 group by Tabela", Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = True Then
                Do While TBAbrir.EOF = False
                    frmVendas_PI.TabelaSN_PI = TBAbrir!Tabela
                    contador = contador + 1
                    TBAbrir.MoveNext
                Loop
                If contador > 1 Then
                    USMsgBox ("Favor informar a tabela do simples nacional utilizada para esse pedido."), vbInformation, "CAPRIND v5.0"
                    frmVendas_proposta_tabelaSN.Show 1
                End If
            Else
                USMsgBox ("Não é permitido importar o pedido, pois não existe nenhuma tabela do simples nacional ativa."), vbExclamation, "CAPRIND v5.0"
                TBAbrir.Close
                xlw.Close False
                Set xlw = Nothing
                Set xl = Nothing
                Exit Sub
            End If
            TBAbrir.Close
        End If
        
        Linha = 1
        Do While .Cells(Linha, 1) <> "FIM"
            Linha = Linha + 1
        Loop
        PBLista.Min = 0
        PBLista.Max = Linha
        
        'CLIENTE
        Set TBClientes = CreateObject("adodb.recordset")
        TBClientes.Open "Select * from Clientes where CPF_CNPJ = '" & .Cells(2, 2).Value & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBClientes.EOF = True Then
            IDCliente = 1
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select IDCliente from clientes order by idcliente desc", Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                IDCliente = TBAbrir!IDCliente + 1
            End If
            TBAbrir.Close
            
            TBClientes.AddNew
            TBClientes!IDCliente = IDCliente
            TBClientes!Data = Date
            TBClientes!Responsavel = pubUsuario
            TBClientes!DtValidacao = Date
            TBClientes!RespValidacao = pubUsuario
            TBClientes!status = "Liberado"
            TBClientes!Prospecto = False
            TBClientes!idTipoEmpresa = 1
            TBClientes!Categoria = "A"
            TBClientes!Tipo_transp = ""
            TBClientes!txt_transportadora = ""
            TBClientes!idTransp = 0
            TBClientes!cxpostal = ""
            
            Evento = "Novo"
        Else
            'Verifica se já existe pedido interno gerado para este pedido de compra
            Set TBCotacao = CreateObject("adodb.recordset")
            TBCotacao.Open "Select VP.Cotacao from vendas_proposta VP INNER JOIN vendas_carteira VC ON VC.Cotacao = VP.Cotacao where VP.IDCliente = " & TBClientes!IDCliente & " and (VC.PCCliente = '" & Nome_anexo1 & "' or VC.PCCliente = '" & Replace(Nome_anexo1, "/", "-") & "')", Conexao, adOpenKeyset, adLockOptimistic
            If TBCotacao.EOF = False Then
                USMsgBox ("Favor informar outro pedido de compra, pois o mesmo já foi importado."), vbExclamation, "CAPRIND v5.0"
                TBCotacao.Close
                GoTo Encerrar
            End If
            TBCotacao.Close
            Evento = "Alterar"
        End If
        
        PBLista.Value = 1
        
        TBClientes!NomeRazao = .Cells(2, 1).Value
        TBClientes!Endereco = .Cells(2, 6).Value
        TBClientes!Bairro = .Cells(2, 10).Value
        TBClientes!Cidade = .Cells(2, 11).Value
        If .Cells(2, 20).Value = "Presumido" Then TBClientes!Presumido = True Else TBClientes!Presumido = False
        If .Cells(2, 20).Value = "Simples" Then TBClientes!Simples = True Else TBClientes!Simples = False
        If .Cells(2, 20).Value = "Real" Then TBClientes!Real = True Else TBClientes!Real = False
        TBClientes!Tipo = "JP"
        TBClientes!Pais = .Cells(2, 14).Value
        TBClientes!Codigo_pais = .Cells(2, 15).Value
        TBClientes!Tipo_endereco = .Cells(2, 5).Value
        TBClientes!Tipo_bairro = .Cells(2, 9).Value
        TBClientes!complemento = .Cells(2, 8).Value
        TBClientes!Tel01 = .Cells(2, 16).Value
        TBClientes!Fax = .Cells(2, 17).Value
        TBClientes!RG_IE = .Cells(2, 3).Value
        TBClientes!RG_IM = .Cells(2, 4).Value
        TBClientes!CPF_CNPJ = .Cells(2, 2).Value
        TBClientes!Email = .Cells(2, 18).Value
        TBClientes!Site = .Cells(2, 19).Value
        TBClientes!UF = .Cells(2, 13).Value
        TBClientes!CEP = .Cells(2, 12).Value
        TBClientes!Numero = .Cells(2, 7).Value
        TBClientes.Update
        '==================================
        Modulo = "Vendas/Clientes"
        ID_documento = TBClientes!IDCliente
        Documento = "Cliente: " & TBClientes!NomeRazao
        Documento1 = ""
        ProcGravaEvento
        '==================================
        
        PBLista.Value = 2
        
        'CRIA O PEDIDO INTERNO
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select Ncotacao from vendas_proposta where Year(Data) = '" & Year(Date) & "' order by Ordenarproposta desc", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
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
        
        Set TBCotacao = CreateObject("adodb.recordset")
        TBCotacao.Open "Select * from vendas_proposta", Conexao, adOpenKeyset, adLockOptimistic
        TBCotacao.AddNew
        TBCotacao!Regime = frmVendas_PI.RegimeEmpresa_PI
        TBCotacao!Ncotacao = NumeroCotacao
        TBCotacao!Tipo = "PE"
        TBCotacao!ID_empresa = frmVendas_PI.Cmb_empresa.ItemData(frmVendas_PI.Cmb_empresa.ListIndex)
        TBCotacao!status = "VENDIDA"
        TBCotacao!Revisao = 0
        TBCotacao!IDCliente = TBClientes!IDCliente
        TBCotacao!Cliente = TBClientes!NomeRazao
        TBCotacao!Remetente = .Cells(3, 9).Value
        TBCotacao!Fax = TBClientes!Fax
        
        If .Cells(3, 10).Value <> "" Then
            TBCotacao!Email = .Cells(3, 10).Value
        ElseIf .Cells(3, 10).Value = "" Then
                Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open "Select Email from Clientes_Contatos where IDCliente = " & TBClientes!IDCliente & " and NomeContato = '" & .Cells(3, 9).Value & "' and (Departamento = 'COMPRA' or Departamento IS NULL or Departamento = N'')", Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    TBCotacao!Email = TBAbrir!Email
                Else
                    TBCotacao!Email = TBClientes!Email
                End If
                TBAbrir.Close
        End If
        
        TBCotacao!Tipo_endereco = TBClientes!Tipo_endereco
        TBCotacao!Endereco = TBClientes!Endereco
        TBCotacao!Numero = TBClientes!Numero
        TBCotacao!complemento = TBClientes!complemento
        TBCotacao!Tipo_bairro = TBClientes!Tipo_bairro
        TBCotacao!Bairro = TBClientes!Bairro
        TBCotacao!Cidade = TBClientes!Cidade
        TBCotacao!telefone = TBClientes!Tel01
        TBCotacao!Departamento = "COMPRAS"
        TBCotacao!UF = TBClientes!UF
        TBCotacao!Tipo_cliente = "JP"
        TBCotacao!Datavendas = Date
        TBCotacao!Data = Date
        TBCotacao!Responsavel = pubUsuario
        TBCotacao.Update
        Conexao.Execute "Update Vendas_proposta set ordenarproposta = " & TBCotacao!Cotacao & " where cotacao = " & TBCotacao!Cotacao
        '==================================
        Modulo = "Vendas/Pedido interno"
        Evento = "Novo"
        ID_documento = TBCotacao!Cotacao
        Documento = "Nº pedido: " & TBCotacao!Ncotacao & " - Rev.: " & 0
        Documento1 = ""
        ProcGravaEvento
        '==================================
        
        PBLista.Value = 3
        
        'DADOS COMERCIAIS
        Set TBGravar = CreateObject("adodb.recordset")
        TBGravar.Open "Select * from vendas_comercial", Conexao, adOpenKeyset, adLockOptimistic
        TBGravar.AddNew
        TBGravar!condicoes = .Cells(3, 1).Value
        TBGravar!Cotacao = TBCotacao!Cotacao
        TBGravar!Observacoes = .Cells(3, 3).Value
        TBGravar!Local_entrega = .Cells(3, 5).Value
        TBGravar!Escopo_fornecimento = .Cells(3, 6).Value
        TBGravar!Moeda = .Cells(3, 7).Value
        TBGravar!Valor_moeda = .Cells(3, 8).Value
        TBGravar.Update
        TBGravar.Close
        
        'PRODUTOS
        contador = 4
        Linha = 4
Inicio:
        Col = 1
        'CADASTRA A UNIDADE DE MEDIDA, PRODUTO E MATERIA PRIMA
        'UNIDADE DE MEDIDA
        Set TBGravar = CreateObject("adodb.recordset")
        TBGravar.Open "Select * from Unidade_Medida where Unidade = '" & .Cells(Linha, 7).Value & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBGravar.EOF = True Then
            ProcNovaUNMedida .Cells(Linha, 7).Value
        End If
        
        Set TBGravar = CreateObject("adodb.recordset")
        TBGravar.Open "Select * from Unidade_Medida where Unidade = '" & .Cells(Linha, 23).Value & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBGravar.EOF = True Then
            ProcNovaUNMedida .Cells(Linha, 23).Value
        End If
        TBGravar.Close
        
        'PRODUTO
        Set TBProduto = CreateObject("adodb.recordset")
        TBProduto.Open "Select P.Codproduto from Projproduto P INNER JOIN item_aplicacoes IA ON IA.Codproduto = P.Codproduto where IA.n_referencia = '" & .Cells(Linha, 1).Value & "' and IA.ID_cliente_forn = " & TBClientes!IDCliente & " and IA.Tipo = 'C'", Conexao, adOpenKeyset, adLockOptimistic
        If TBProduto.EOF = True Then
            ProcNovoProduto IIf(.Cells(Linha, 17).Value = True, Cmb_familia_material, Cmb_familia_produto), .Cells(Linha, 7).Value, .Cells(Linha, 23).Value, .Cells(Linha, 2).Value, .Cells(Linha, 1).Value, IIf(.Cells(Linha, 17).Value = True, False, True), IIf(.Cells(Linha, 17).Value = True, True, False), IIf(.Cells(Linha, 17).Value = True, False, True), True, IIf(.Cells(Linha, 17).Value = True, 0, 1), TBClientes!IDCliente, TBClientes!NomeRazao, .Cells(Linha, 79).Value, .Cells(Linha, 84).Value
        End If
        
        'CADASTRA ESTRUTURA
        If .Cells(Linha, 76) <> "" Then
            Set TBProduto = CreateObject("adodb.recordset")
            TBProduto.Open "Select P.Codproduto from Projproduto P INNER JOIN item_aplicacoes IA ON IA.Codproduto = P.Codproduto where IA.n_referencia = '" & .Cells(Linha, 1).Value & "' and IA.ID_cliente_forn = " & TBClientes!IDCliente & " and IA.Tipo = 'C'", Conexao, adOpenKeyset, adLockOptimistic
            If TBProduto.EOF = False Then
                Set TBItem = CreateObject("adodb.recordset")
                TBItem.Open "Select P.Desenho from Projproduto P INNER JOIN item_aplicacoes IA ON IA.Codproduto = P.Codproduto where IA.n_referencia = '" & .Cells(Linha, 76).Value & "' and IA.ID_cliente_forn = " & TBClientes!IDCliente & " and IA.Tipo = 'C'", Conexao, adOpenKeyset, adLockOptimistic
                If TBItem.EOF = False Then
                    Set TBGravar = CreateObject("adodb.recordset")
                    TBGravar.Open "Select * from Projconjunto where Codproduto = " & TBProduto!Codproduto & " and Desenho = '" & TBItem!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBGravar.EOF = True Then
                        'Verifica última versão da estrutura
                        Set TBFI = CreateObject("adodb.recordset")
                        TBFI.Open "Select Versao from Projconjunto where Codproduto = " & TBProduto!Codproduto & " order by Versao desc", Conexao, adOpenKeyset, adLockOptimistic
                        If TBFI.EOF = True Then
                            versao = "A"
                        Else
                            versao = FunGerarVersaoEstrutura(TBFI!versao)
                        End If
                        TBGravar.AddNew
                        TBGravar!versao = versao
                        TBGravar!Codproduto = TBProduto!Codproduto
                        TBGravar!Desenho = TBItem!Desenho
                        TBGravar!Descricao = .Cells(Linha, 77).Value
                        TBGravar!PesoMetro = .Cells(Linha, 78).Value
                        TBGravar!PesoTotal = .Cells(Linha, 79).Value
                        TBGravar!quantidade = .Cells(Linha, 80).Value
                        TBGravar!Peso = .Cells(Linha, 81).Value
                        TBGravar!Unidade = .Cells(Linha, 82).Value
                        TBGravar!Dimensoes = .Cells(Linha, 83).Value
                        TBGravar!Un_Kg = .Cells(Linha, 84).Value
                        TBGravar!Posicao = .Cells(Linha, 85).Value
                        TBGravar.Update
                    End If
                    TBGravar.Close
                End If
                TBItem.Close
            End If
            TBProduto.Close
        End If
        
        'CADASTRA PRODUTOS E SERVIÇOS NO PEDIDO
        Set TBLISTA = CreateObject("adodb.recordset")
        TBLISTA.Open "Select * from vendas_carteira", Conexao, adOpenKeyset, adLockOptimistic
        TBLISTA.AddNew
        TBLISTA!Liberacao = "VENDIDA"
        TBLISTA!Datavendas = Date
        TBLISTA!Tem_ordem = False
        TBLISTA!Observacoes = .Cells(Linha, 12).Value
        
        Set TBProduto = CreateObject("adodb.recordset")
        TBProduto.Open "Select P.Codproduto, P.Desenho, P.Revdesenho, P.Classe from Projproduto P INNER JOIN item_aplicacoes IA ON IA.Codproduto = P.Codproduto where IA.n_referencia = '" & .Cells(Linha, 1).Value & "' and IA.ID_cliente_forn = " & TBClientes!IDCliente & " and IA.Tipo = 'C'", Conexao, adOpenKeyset, adLockOptimistic
        If TBProduto.EOF = False Then
            TBLISTA!Desenho = TBProduto!Desenho
            TBLISTA!N_referencia = .Cells(Linha, 1).Value
            TBLISTA!Rev_codinterno = TBProduto!RevDesenho
            TBLISTA!Familia = TBProduto!Classe
        End If
        TBLISTA!quantidade = .Cells(Linha, 3).Value
        TBLISTA!Qtde_produzir = TBLISTA!quantidade / FunVerificaTabelaConversaoUnidade(.Cells(Linha, 7).Value, .Cells(Linha, 23).Value)
        TBLISTA!Desconto = .Cells(Linha, 14).Value
        TBLISTA!ValorDesconto = .Cells(Linha, 15).Value
        TBLISTA!preco_unitario_desconto = .Cells(Linha, 16).Value
        TBLISTA!Descricao = .Cells(Linha, 22).Value
        TBLISTA!Antecipacao_fat = False
        TBLISTA!Faturamento_parcial = False
        TBLISTA!Comprimento = Null
        TBLISTA!Largura = Null
        TBLISTA!Espessura = Null
        TBLISTA!Dureza = Null
        TBLISTA!descricao_tecnica = .Cells(Linha, 2).Value
        TBLISTA!preco_unitario = .Cells(Linha, 4).Value
        TBLISTA!Unidade = .Cells(Linha, 7).Value
        TBLISTA!Unidade_com = .Cells(Linha, 23).Value
        TBLISTA!txt_CST = .Cells(Linha, 68).Value
        TBLISTA!Cotacao = TBCotacao!Cotacao
        TBLISTA!preco_lote = .Cells(Linha, 6).Value
        TBLISTA!Tipo = .Cells(Linha, 18).Value
        If .Cells(Linha, 17).Value = True Then
            TBLISTA!retorno = True
            TBLISTA!Data_retorno = .Cells(Linha, 13).Value
        Else
            TBLISTA!retorno = False
        End If
        TBLISTA!Comissao = 0
        TBLISTA!ValorComissao = 0
        TBLISTA!dbl_valoripi = .Cells(Linha, 9).Value
        TBLISTA!IntICMS = .Cells(Linha, 10).Value
        TBLISTA!int_IPI = .Cells(Linha, 5).Value
        TBLISTA!dbl_Valor_ICMS = .Cells(Linha, 8).Value
        TBLISTA!Embalagem = .Cells(3, 2).Value
        
        'CADASTRA A CFOP
        If .Cells(Linha, 24).Value <> "" Then
            TextoFiltro = "ID_CFOP = '" & Replace(.Cells(Linha, 24).Value, "1.", "5.") & "' and Txt_ICMS = '" & .Cells(Linha, 26).Value & "' and txt_IPI = '" & .Cells(Linha, 27).Value & "' and txt_Somar = '" & .Cells(Linha, 28).Value & "' and Vendas = " & FunConverterPInt(.Cells(Linha, 29).Value) & " and Retem = " & FunConverterPInt(.Cells(Linha, 30).Value) & " and Suframa = " & _
                FunConverterPInt(.Cells(Linha, 31).Value) & " and MaoObra = " & FunConverterPInt(.Cells(Linha, 32).Value) & " and Demonstracao = " & FunConverterPInt(.Cells(Linha, 33).Value) & " and Soma_retorno_totalnf = " & FunConverterPInt(.Cells(Linha, 34).Value) & " and TemPIS = " & FunConverterPInt(.Cells(Linha, 35).Value) & " and TemCOFINS = " & FunConverterPInt(.Cells(Linha, 36).Value) & " and MPA = " & FunConverterPInt(.Cells(Linha, 39).Value) & " and TemReducaoBC = " & FunConverterPInt(.Cells(Linha, 40).Value) & " and Remessa = " & FunConverterPInt(.Cells(Linha, 41).Value) & " and Retorno = " & FunConverterPInt(.Cells(Linha, 42).Value) & " and Somar_IPI_BC_ICMSST = " & FunConverterPInt(.Cells(Linha, 43).Value) & " and Devolucao = " & FunConverterPInt(.Cells(Linha, 44).Value)
            Set TBGravar = CreateObject("adodb.recordset")
            TBGravar.Open "Select * from tbl_NaturezaOperacao where " & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
            If TBGravar.EOF = True Then
                TBGravar.AddNew
                TBGravar!ID_CFOP = Replace(.Cells(Linha, 24).Value, "1.", "5.")
                TBGravar!Proprio = True
                TBGravar!Terceiros = False
                TBGravar!Txt_ICMS = .Cells(Linha, 26).Value
                TBGravar!txt_IPI = .Cells(Linha, 27).Value
                TBGravar!txt_Somar = .Cells(Linha, 28).Value
                TBGravar!Vendas = .Cells(Linha, 29).Value
                TBGravar!Retem = .Cells(Linha, 30).Value
                TBGravar!Suframa = .Cells(Linha, 31).Value
                TBGravar!MaoObra = .Cells(Linha, 32).Value
                TBGravar!Demonstracao = .Cells(Linha, 33).Value
                TBGravar!Soma_retorno_totalnf = .Cells(Linha, 34).Value
                TBGravar!TemPIS = .Cells(Linha, 35).Value
                TBGravar!TemCOFINS = .Cells(Linha, 36).Value
                TBGravar!De = .Cells(Linha, 37).Value
                TBGravar!FE = .Cells(Linha, 38).Value
                TBGravar!MPA = .Cells(Linha, 39).Value
                TBGravar!TemReducaoBC = .Cells(Linha, 40).Value
                TBGravar!Remessa = .Cells(Linha, 41).Value
                TBGravar!retorno = .Cells(Linha, 42).Value
                TBGravar!Somar_IPI_BC_ICMSST = .Cells(Linha, 43).Value
                TBGravar!Devolucao = .Cells(Linha, 44).Value
                TBGravar!Data = Date
                TBGravar!Responsavel = pubUsuario
                TBGravar!DtValidacao = Date
                TBGravar!RespValidacao = pubUsuario
            End If
            Select Case Right(TBGravar!ID_CFOP, 3)
                Case "101": TBGravar!Txt_descricao = "VENDA DE PRODUÇÃO DO ESTABELECIMENTO"
                Case "124": TBGravar!Txt_descricao = "INDUSTRIALIZAÇÃO EFETUADA PARA OUTRA EMPRESA"
                Case "902": TBGravar!Txt_descricao = "RET. DE MERC. UTIL. NA INDUSTR. POR ENCOMENDA"
                Case "903": TBGravar!Txt_descricao = "RET. MERC. RECEB. P/ IND. E NÃO APL. NO REF. PROC."
            End Select
            TBGravar.Update
            
            TBLISTA!ID_CFOP = TBGravar!IDCountCfop
            TBGravar.Close
        Else
            TBLISTA!ID_CFOP = 0
        End If
        
        'CADASTRA A NCM
        If .Cells(Linha, 45).Value <> "" Then
            valor = .Cells(Linha, 48).Value
            Valor1 = .Cells(Linha, 49).Value
            Valor2 = .Cells(Linha, 50).Value
            Valor3 = .Cells(Linha, 51).Value
            ValorIPI = .Cells(Linha, 52).Value
            Valor_Cofins_Prod = .Cells(Linha, 54).Value
            Valor_Cofins_Serv = .Cells(Linha, 55).Value
            Valor_PIS_Prod = .Cells(Linha, 56).Value
            Valor_PIS_Serv = .Cells(Linha, 57).Value
            Valor_IRPJ_Prod = .Cells(Linha, 59).Value
            Valor_IRPJ_Serv = .Cells(Linha, 60).Value
            Valor_Retencao_PIS = .Cells(Linha, 61).Value
            Valor_Retencao_Cofins = .Cells(Linha, 62).Value
            Valor_CSLL_Prod = .Cells(Linha, 63).Value
            Valor_CSLL_Serv = .Cells(Linha, 64).Value
            NovoValor = Replace(valor, ",", ".")
            NovoValor1 = Replace(Valor1, ",", ".")
            NovoValor2 = Replace(Valor2, ",", ".")
            NovoValor3 = Replace(Valor3, ",", ".")
            NovoValor4 = Replace(ValorIPI, ",", ".")
            NovoValor5 = Replace(Valor_Cofins_Prod, ",", ".")
            NovoValor6 = Replace(Valor_Cofins_Serv, ",", ".")
            NovoValor7 = Replace(Valor_PIS_Prod, ",", ".")
            NovoValor8 = Replace(Valor_PIS_Serv, ",", ".")
            NovoValor9 = Replace(Valor_IRPJ_Prod, ",", ".")
            NovoValor10 = Replace(Valor_IRPJ_Serv, ",", ".")
            NovoValor11 = Replace(Valor_Retencao_PIS, ",", ".")
            NovoValor12 = Replace(Valor_Retencao_Cofins, ",", ".")
            NovoValor13 = Replace(Valor_CSLL_Prod, ",", ".")
            NovoValor14 = Replace(Valor_CSLL_Serv, ",", ".")
            TextoFiltro = "IDIntClasse = '" & .Cells(Linha, 45).Value & "' and Txt_grupo = '" & .Cells(Linha, 46).Value & "' and dbl_ICMS_de = " & NovoValor & " and dbl_ICMS_ss = " & NovoValor1 & " and dbl_ICMS_nn = " & NovoValor2 & " and dbl_ICMS_co = " & NovoValor3 & " and dbl_IPI = " & NovoValor4 & " and BaseReduzida = " & FunConverterPInt(.Cells(Linha, 53).Value) & " and CTDE = " & NovoValor5 & " and CTNN = " & NovoValor6 & " and CTCO = " & NovoValor7 & " and CTSS = " & NovoValor8 & " and Retem_PIS_Cofins = " & FunConverterPInt(.Cells(Linha, 58).Value) & " and PIS = " & NovoValor9 & " and Cofins = " & NovoValor10 & " and PIS_destaca = " & NovoValor11 & " and Cofins_destaca = " & NovoValor12 & " and dbl_ICMS_ex = " & NovoValor13 & " and CTEX = " & NovoValor14
            
            Set TBGravar = CreateObject("adodb.recordset")
            TBGravar.Open "Select * from tbl_ClassificacaoFiscal where " & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
            If TBGravar.EOF = True Then
                TBGravar.AddNew
                TBGravar!IDIntClasse = .Cells(Linha, 45).Value
                TBGravar!Txt_grupo = .Cells(Linha, 46).Value
                TBGravar!txt_Class = .Cells(Linha, 47).Value
                TBGravar!dbl_ICMS_de = valor
                TBGravar!dbl_ICMS_ss = Valor1
                TBGravar!dbl_ICMS_nn = Valor2
                TBGravar!dbl_ICMS_co = Valor3
                TBGravar!dbl_IPI = ValorIPI
                TBGravar!basereduzida = .Cells(Linha, 53).Value
                TBGravar!CTDE = Valor_Cofins_Prod
                TBGravar!CTNN = Valor_Cofins_Serv
                TBGravar!CTCO = Valor_PIS_Prod
                TBGravar!CTSS = Valor_PIS_Serv
                TBGravar!Retem_PIS_Cofins = .Cells(Linha, 58).Value
                TBGravar!PIS = Valor_IRPJ_Prod
                TBGravar!Cofins = Valor_IRPJ_Serv
                TBGravar!PIS_destaca = Valor_Retencao_PIS
                TBGravar!Cofins_destaca = Valor_Retencao_Cofins
                TBGravar!dbl_ICMS_ex = Valor_CSLL_Prod
                TBGravar!CTEX = Valor_CSLL_Serv
                TBGravar!Desoneracao = .Cells(Linha, 65).Value
                TBGravar!Aliq_nacional = .Cells(Linha, 66).Value
                TBGravar!Aliq_importacao = .Cells(Linha, 67).Value
                TBGravar!Data = Date
                TBGravar!Responsavel = pubUsuario
                TBGravar!DtValidacao = Date
                TBGravar!RespValidacao = pubUsuario
                TBGravar.Update
            End If
            TBLISTA!ID_CF = TBGravar!Idclass
            TBGravar.Close
            
            'Salvar ID da NCM no produto
            Conexao.Execute "Update projproduto Set ID_CF = " & TBLISTA!ID_CF & " where Desenho = '" & TBLISTA!Desenho & "'"
        Else
            TBLISTA!ID_CF = Null
        End If
        
        TBLISTA!BC_ICMS = 0
        TBLISTA!BC_ICMS_ST = 0
        TBLISTA!Valor_ICMS_ST = 0
        If IsNull(TBLISTA!ID_CF) = False Then
            ProcValorImposto TBCotacao!Cotacao, TBLISTA!ID_CF, TBCotacao!IDCliente, TBCotacao!Cliente, TBCotacao!UF, TBCotacao!ID_empresa, False, IIf(IsNull(TBLISTA!ID_CFOP), 0, TBLISTA!ID_CFOP), frmVendas_PI.RegimeEmpresa_PI
            ProcControleImposto IIf(IsNull(TBLISTA!ID_CFOP), 0, TBLISTA!ID_CFOP), TBCotacao!IDCliente
            ProcCalculaBC TBCotacao!ID_empresa, Replace(.Cells(Linha, 24).Value, "1.", "5."), 0, TBLISTA!preco_lote, TBLISTA!dbl_valoripi, SomarIPI, SomarIPIST, TemReducaoBC, False, IIf(IsNull(TBLISTA!txt_CST), "", TBLISTA!txt_CST), "P", 0, ""
            If TemICMS = "SIM" And TBLISTA!dbl_Valor_ICMS <> 0 Then TBLISTA!BC_ICMS = BC
    
            If IsNull(TBLISTA!txt_CST) = False And TBLISTA!txt_CST <> "" And TBLISTA!retorno = False Then
                ProcSubstituicaoTributaria TBCotacao!UF, TBLISTA!txt_CST, TBLISTA!ID_CF, TBCotacao!IDCliente, TBCotacao!Cliente, TBLISTA!preco_unitario_desconto, TBLISTA!quantidade, BC, BCST, 0, 0, 0, False, False, 0
                TBLISTA!Valor_ICMS_ST = ICMSCST
                If ICMSCST <> 0 Then TBLISTA!BC_ICMS_ST = BCICMSCST
            End If
        End If
    
        'Impostos
        Valor_total = TBLISTA!preco_unitario_desconto * TBLISTA!quantidade
        Valor_IPI = TBLISTA!dbl_valoripi
    
        If IsNull(TBLISTA!ID_CF) = False Then
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select * from tbl_classificacaofiscal where Idclass = " & TBLISTA!ID_CF, Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False Then
                'Verifica se a CF tem retenção de PIS/Cofins, destaca PIS/Cofins e grava no produto
                If TBFI!Retem_PIS_Cofins = True Then
                    TBLISTA!Valor_Retencao_PIS = Format((Valor_total * IIf(IsNull(TBFI!PIS), 0, TBFI!PIS)) / 100, "###,##0.00")
                    TBLISTA!Valor_Retencao_Cofins = Format((Valor_total * IIf(IsNull(TBFI!Cofins), 0, TBFI!Cofins)) / 100, "###,##0.00")
                End If
    
                If Regime <> 1 Then
                    PIS_Prod = IIf(IsNull(TBFI!PIS_destaca), 0, TBFI!PIS_destaca)
                    Cofins_Prod = IIf(IsNull(TBFI!Cofins_destaca), 0, TBFI!Cofins_destaca)
                    If PIS_Prod <> 0 Then
                        TBLISTA!PIS_Prod = PIS_Prod
                        TBLISTA!Total_PIS_prod = Format((Valor_total * PIS_Prod) / 100, "###,##0.00")
                    End If
                    If Cofins_Prod <> 0 Then
                        TBLISTA!Cofins_Prod = Cofins_Prod
                        TBLISTA!Total_Cofins_prod = Format((Valor_total * Cofins_Prod) / 100, "###,##0.00")
                    End If
                End If
            End If
            TBFI.Close
        End If
    
        TBLISTA!PCCliente = Nome_anexo1
        Dataini = .Cells(contador, 13).Value
        TBLISTA!PrazoFinal = Dataini
        TBLISTA!prazofinaldias = Date - Dataini
    
        'Empresa
        ProcControleImposto IIf(IsNull(TBLISTA!ID_CFOP), 0, TBLISTA!ID_CFOP), TBCotacao!IDCliente
        ProcVerifImpostosEmpresa TBCotacao!ID_empresa, TBLISTA!retorno, "", False, 0, False, frmVendas_PI.TabelaSN_PI, 0
    
        TBLISTA!PIS_Prod = PIS_Prod
        If PIS_Prod <> 0 Then TBLISTA!Total_PIS_prod = Format((Valor_total * PIS_Prod) / 100, "###,##0.00") Else TBLISTA!Total_PIS_prod = 0
        TBLISTA!Cofins_Prod = Cofins_Prod
        If Cofins_Prod <> 0 Then TBLISTA!Total_Cofins_prod = Format((Valor_total * Cofins_Prod) / 100, "###,##0.00") Else TBLISTA!Total_Cofins_prod = 0
        TBLISTA!CSLL_Prod = CSLL_Prod
        If CSLL_Prod <> 0 Then TBLISTA!Total_CSLL_prod = Format((Valor_total * CSLL_Prod) / 100, "###,##0.00") Else TBLISTA!Total_CSLL_prod = 0
        TBLISTA!IRPJ_Prod = IRPJ_Prod
        If IRPJ_Prod <> 0 Then TBLISTA!Total_IRPJ_prod = Format((Valor_total * IRPJ_Prod) / 100, "###,##0.00") Else TBLISTA!Total_IRPJ_prod = 0
        TBLISTA!DAS = DAS
        If DAS <> 0 Then TBLISTA!Total_DAS = Format((Valor_total * DAS) / 100, "###,##0.00") Else TBLISTA!Total_DAS = 0
    
        TBLISTA.Update
    
        Set TBFI = CreateObject("adodb.recordset")
        TBFI.Open "Select codproduto from Projproduto where desenho = '" & TBLISTA!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBFI.EOF = False Then ProcAgregarProdutoCli TBFI!Codproduto, TBCotacao!IDCliente, TBCotacao!Tipo_cliente, TBLISTA!Unidade, TBLISTA!Unidade_com, TBLISTA!preco_unitario
        TBFI.Close
        
        TBLISTA.Close
        
        Linha = Linha + 1
        contador = contador + 1
        PBLista.Value = contador
        
        'Captura o valor das celulas e grava nas variaves
        If .Cells(Linha, 1) = "FIM" Then
            IDpedido = TBCotacao!Cotacao
            USMsgBox ("Importação efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
            '==================================
            Modulo = "Vendas/Pedido interno"
            Evento = "Novo (importação)"
            ID_documento = TBCotacao!Cotacao
            Documento = "Nº pedido: " & TBCotacao!Ncotacao & " - Rev.: " & TBCotacao!Revisao
            Documento1 = ""
            ProcGravaEvento
            '==================================
            TBClientes.Close
            TBCotacao.Close
Encerrar:
            xlw.Close False
            Set xlw = Nothing
            Set xl = Nothing
            If Permitido2 = True Then GoTo CarregaPedido
        End If
        
        GoTo Inicio
    End With
    
CarregaPedido:
        With frmVendas_PI
            .StrSql_PI_Localizar = "Select * from vendas_proposta where Cotacao = " & IDpedido & " and (Tipo = 'PE' or Tipo = 'PRPE')"
            .ProcCarregaLista (1)
            
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select * from vendas_proposta where cotacao = " & IDpedido, Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                .Novo_PI = False
                .ProcLimpar
                .ProcPuxaDados
                .ProcLimparTotais
                .ProcPuxaTotais
            End If
            TBAbrir.Close
        End With

Exit Sub
tratar_erro:
    Teste = Err.Number
    If Err.Number = 9 Then
        NomePlanilha = "Planilha1"
        GoTo AbrePlanilha
    End If
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovaUNMedida(Un As String)
On Error GoTo tratar_erro

TBGravar.AddNew
TBGravar!Data = Date
TBGravar!Responsavel = pubUsuario
TBGravar!Unidade = Un
TBGravar.Update

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovoProduto(Familia As String, Un As String, Un_com As String, Descricao As String, Cod_ref As String, Vendas As Boolean, Compras As Boolean, PCP As Boolean, Qualidade As Boolean, SubTipoItem As Integer, IDCliente As Long, NomeRazao As String, Peso As String, Un_Kg As String)
On Error GoTo tratar_erro

Set TBFamilia = CreateObject("adodb.recordset")
TBFamilia.Open "Select Letra from projfamilia where familia = '" & Familia & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBFamilia.EOF = False Then Letra = TBFamilia!Letra
TBFamilia.Close

CompLetra = Len(Letra)
Set TBComponente = CreateObject("adodb.recordset")
TBComponente.Open "Select * from projproduto where classe = '" & Familia & "' and codmanual = 'False' and Right(Desenho, " & CompLetra & ") = '" & Letra & "' and (subtipoitem = 0 or subtipoitem = 1 or subtipoitem = 4 or subtipoitem = 5) order by codproduto desc", Conexao, adOpenKeyset, adLockOptimistic
If TBComponente.EOF = False Then
    Numero = Left(TBComponente!Desenho, Len(TBComponente!Desenho) - (CompLetra + 1))
    Numero = Numero + 1
    Select Case Len(Numero)
        Case 5: Desenho = Numero & "-" & Letra
        Case 4: Desenho = "0" & Numero & "-" & Letra
        Case 3: Desenho = "00" & Numero & "-" & Letra
        Case 2: Desenho = "000" & Numero & "-" & Letra
        Case 1: Desenho = "0000" & Numero & "-" & Letra
    End Select
Else
    Desenho = "00001" & "-" & Letra
End If

VerifCodigo:
    Set TBItem = CreateObject("adodb.recordset")
    TBItem.Open "Select Codproduto from projproduto where desenho = '" & Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBItem.EOF = False Then
        CompLetra = Len(Letra)
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
    TBItem.Close
    
    TBComponente.AddNew
    TBComponente!Data = Date
    TBComponente!Responsavel = pubUsuario
    TBComponente!CodManual = False
    TBComponente!Desenho = Desenho
    TBComponente!RevDesenho = 0
    TBComponente!Unidade = Un
    TBComponente!Unidade_com = Un_com
    TBComponente!Descricao = Descricao
    TBComponente!descricaotecnica = Descricao
    TBComponente!Classe = Familia
            
    'Conta contábil e CC
    Set TBFamilia = CreateObject("adodb.recordset")
    TBFamilia.Open "select * from projfamilia where Familia = '" & Familia & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBFamilia.EOF = False Then
        If IsNull(TBFamilia!ID_PC) = False And TBFamilia!ID_PC <> "" Then TBComponente!ID_PC = TBFamilia!ID_PC
        If IsNull(TBFamilia!ID_PC1) = False And TBFamilia!ID_PC1 <> "" Then TBComponente!ID_PC1 = TBFamilia!ID_PC1
        If IsNull(TBFamilia!ID_CC) = False And TBFamilia!ID_CC <> "" Then TBComponente!ID_CC = TBFamilia!ID_CC
    End If
    TBFamilia.Close
        
    TBComponente!Vendas = Vendas
    TBComponente!Compras = Compras
    TBComponente!Producao = PCP
    TBComponente!Qualidade = Qualidade
    TBComponente!Tipo = "P"
    TBComponente!SubTipoItem = SubTipoItem
    If SubTipoItem = 0 Then
        TBComponente!peso_metro = IIf(Peso = "", Null, Peso)
        TBComponente!Un_Kg = IIf(Un_Kg = "", Null, Un_Kg)
    Else
        TBComponente!DtValidacaoConj = Date
        TBComponente!RespValidacaoConj = pubUsuario
    End If
    TBComponente!Leadtime = 0
    TBComponente.Update
                
    Set TBItem = CreateObject("adodb.recordset")
    TBItem.Open "Select * from item_aplicacoes where codproduto = " & TBComponente!Codproduto, Conexao, adOpenKeyset, adLockOptimistic
    If TBItem.EOF = True Then
        TBItem.AddNew
        TBItem!N_referencia = Cod_ref
        TBItem!Aplicacao = NomeRazao
        TBItem!Codproduto = TBComponente!Codproduto
        TBItem!Descricao = Descricao
        TBItem!ID_cliente_forn = IDCliente
        TBItem!Tipo = "C"
        TBItem.Update
    End If
    TBItem.Close
    
    TBComponente.Close

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

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: procImportar
    'Case 3: ProcAjuda
    Case 4: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Function FunConverterPInt(Texto As String) As Integer
On Error GoTo tratar_erro

FunConverterPInt = IIf(Texto = "True", 1, 0)

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function
