VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmproj_produto_PC 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Engenharia - Produtos e serviços - Conta contábil"
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12180
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   12180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USTreeView USTreeView1 
      Height          =   6645
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Width           =   12105
      _ExtentX        =   21352
      _ExtentY        =   11721
      BorderColor     =   12500670
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Theme           =   1
   End
   Begin DrawSuite2022.USButton Cmd_carregar 
      Height          =   405
      Left            =   7470
      TabIndex        =   1
      Top             =   6810
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   714
      BorderColor     =   8421504
      BorderColorDisabled=   0
      BorderColorDown =   15048022
      BorderColorOver =   15381630
      Caption         =   "Carregar (F3)"
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GradientColor2  =   14737632
      GradientColor3  =   12632256
      GradientColor4  =   12632256
      State           =   3
      Theme           =   1
   End
   Begin DrawSuite2022.USButton Cmd_sair 
      Height          =   405
      Left            =   10710
      TabIndex        =   3
      Top             =   6810
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   714
      BorderColor     =   8421504
      BorderColorDisabled=   0
      BorderColorDown =   15048022
      BorderColorOver =   15381630
      Caption         =   "Sair (Esc)"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GradientColor2  =   14737632
      GradientColor3  =   12632256
      GradientColor4  =   12632256
      Theme           =   1
   End
   Begin DrawSuite2022.USButton Cmd_carregar_vazio 
      Height          =   405
      Left            =   8970
      TabIndex        =   2
      Top             =   6810
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   714
      BorderColor     =   8421504
      BorderColorDisabled=   0
      BorderColorDown =   15048022
      BorderColorOver =   15381630
      Caption         =   "Carregar vazio (F4)"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GradientColor2  =   14737632
      GradientColor3  =   12632256
      GradientColor4  =   12632256
      Theme           =   1
   End
End
Attribute VB_Name = "frmproj_produto_PC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmd_carregar_Click()
On Error GoTo tratar_erro

If IDlista = 0 Then Exit Sub
Set TBFamilia = CreateObject("adodb.recordset")
TBFamilia.Open "Select * FROM tbl_familia where int_codfamilia = " & IDlista, Conexao, adOpenKeyset, adLockOptimistic
If TBFamilia.EOF = False Then
    If Plano_contas_produtos = True Then
        With frmproj_produto
            If Aplic = 1 Then
                .Txt_ID_PC = IDlista
                .Txt_codigo_PC = IIf(IsNull(TBFamilia!CODIGO), "", TBFamilia!CODIGO)
                .Txt_descricao_PC = IIf(IsNull(TBFamilia!Txt_descricao), "", TBFamilia!Txt_descricao)
            Else
                .Txt_ID_PC1 = IDlista
                .Txt_codigo_PC1 = IIf(IsNull(TBFamilia!CODIGO), "", TBFamilia!CODIGO)
                .Txt_descricao_PC1 = IIf(IsNull(TBFamilia!Txt_descricao), "", TBFamilia!Txt_descricao)
            End If
        End With
    ElseIf Plano_contas_familias = True Then
            With frmproj_familia
                If Aplic = 1 Then
                    .Txt_ID_PC = IDlista
                    .Txt_codigo_PC = IIf(IsNull(TBFamilia!CODIGO), "", TBFamilia!CODIGO)
                    .Txt_descricao_PC = IIf(IsNull(TBFamilia!Txt_descricao), "", TBFamilia!Txt_descricao)
                Else
                    .Txt_ID_PC1 = IDlista
                    .Txt_codigo_PC1 = IIf(IsNull(TBFamilia!CODIGO), "", TBFamilia!CODIGO)
                    .Txt_descricao_PC1 = IIf(IsNull(TBFamilia!Txt_descricao), "", TBFamilia!Txt_descricao)
                End If
            End With
        ElseIf Plano_centro_de_custo = True Then
                With Frm_centro_de_custo
                    Select Case Sit_REG
                        Case 1:
                            .Txt_ID_PC = IDlista
                            .Txt_codigo_PC = IIf(IsNull(TBFamilia!CODIGO), "", TBFamilia!CODIGO)
                            .Txt_descricao_PC = IIf(IsNull(TBFamilia!Txt_descricao), "", TBFamilia!Txt_descricao)
                        Case 2:
                            .Txt_ID_PC_depreciacao = IDlista
                            .Txt_codigo_PC_depreciacao = IIf(IsNull(TBFamilia!CODIGO), "", TBFamilia!CODIGO)
                            .Txt_descricao_PC_depreciacao = IIf(IsNull(TBFamilia!Txt_descricao), "", TBFamilia!Txt_descricao)
                        Case 3:
                            .Txt_ID_PC_rateio = IDlista
                            .Txt_codigo_PC_rateio = IIf(IsNull(TBFamilia!CODIGO), "", TBFamilia!CODIGO)
                            .Txt_descricao_PC_rateio = IIf(IsNull(TBFamilia!Txt_descricao), "", TBFamilia!Txt_descricao)
                    End Select
                End With
            ElseIf Plano_instituicao = True Then
                    With frm_Instituicoes
                        Select Case Sit_REG
                            Case 1:
                                .Txt_ID_PC_instituicao = IDlista
                                .Txt_codigo_PC_instituicao = IIf(IsNull(TBFamilia!CODIGO), "", TBFamilia!CODIGO)
                                .Txt_descricao_PC_instituicao = IIf(IsNull(TBFamilia!Txt_descricao), "", TBFamilia!Txt_descricao)
                            Case 2:
                                .Txt_ID_PC_instituicao_rec = IDlista
                                .Txt_codigo_PC_instituicao_rec = IIf(IsNull(TBFamilia!CODIGO), "", TBFamilia!CODIGO)
                                .Txt_descricao_PC_instituicao_rec = IIf(IsNull(TBFamilia!Txt_descricao), "", TBFamilia!Txt_descricao)
                            Case 3:
                                .Txt_ID_PC = IDlista
                                .Txt_codigo_PC = IIf(IsNull(TBFamilia!CODIGO), "", TBFamilia!CODIGO)
                                .Txt_descricao_PC = IIf(IsNull(TBFamilia!Txt_descricao), "", TBFamilia!Txt_descricao)
                        End Select
                    End With
                ElseIf Plano_opcoesgerais = True Then
                        With frmOpcoesGeral_PC
                            .Txt_ID_PC = IDlista
                            .Txt_codigo_PC = IIf(IsNull(TBFamilia!CODIGO), "", TBFamilia!CODIGO)
                            .Txt_descricao_PC = IIf(IsNull(TBFamilia!Txt_descricao), "", TBFamilia!Txt_descricao)
                        End With
                    ElseIf Plano_Faturamento = True Then
                            If Sit_REG = 1 Then
                                With frmFaturamento_Prod_Serv_FamiliasDuplicata
                                    .Txt_ID_PC = IDlista
                                    .Txt_codigo_PC = IIf(IsNull(TBFamilia!CODIGO), "", TBFamilia!CODIGO)
                                    .Txt_descricao_PC = IIf(IsNull(TBFamilia!Txt_descricao), "", TBFamilia!Txt_descricao)
                                End With
                            Else
                                With frmFaturamento_Prod_Serv_DI
                                    .Txt_ID_PC = IDlista
                                    .Txt_codigo_PC = IIf(IsNull(TBFamilia!CODIGO), "", TBFamilia!CODIGO)
                                    .Txt_descricao_PC = IIf(IsNull(TBFamilia!Txt_descricao), "", TBFamilia!Txt_descricao)
                                End With
                            End If
                        ElseIf Plano_PCP = True Then
                                With frmprod_outras_despesas
                                    .Txt_ID_PC = IDlista
                                    .Txt_codigo_PC = IIf(IsNull(TBFamilia!CODIGO), "", TBFamilia!CODIGO)
                                    .Txt_descricao_PC = IIf(IsNull(TBFamilia!Txt_descricao), "", TBFamilia!Txt_descricao)
                                End With
                            Else
                                With frmFamilia_financeiro
                                    .Txt_ID_PC = IDlista
                                    .Txt_codigo_PC = IIf(IsNull(TBFamilia!CODIGO), "", TBFamilia!CODIGO)
                                    .Txt_descricao_PC = IIf(IsNull(TBFamilia!Txt_descricao), "", TBFamilia!Txt_descricao)
                                End With
    End If
End If
TBFamilia.Close
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_carregar_vazio_Click()
On Error GoTo tratar_erro

If Plano_contas_produtos = True Then
    With frmproj_produto
        If Aplic = 1 Then
            .Txt_ID_PC = 0
            .Txt_codigo_PC = ""
            .Txt_descricao_PC = ""
        Else
            .Txt_ID_PC1 = 0
            .Txt_codigo_PC1 = ""
            .Txt_descricao_PC1 = ""
        End If
    End With
ElseIf Plano_contas_familias = True Then
        With frmproj_familia
            If Aplic = 1 Then
                .Txt_ID_PC = 0
                .Txt_codigo_PC = ""
                .Txt_descricao_PC = ""
            Else
                .Txt_ID_PC1 = 0
                .Txt_codigo_PC1 = ""
                .Txt_descricao_PC1 = ""
            End If
        End With
    ElseIf Plano_centro_de_custo = True Then
            With Frm_centro_de_custo
                Select Case Sit_REG
                    Case 1:
                        .Txt_ID_PC = 0
                        .Txt_codigo_PC = ""
                        .Txt_descricao_PC = ""
                    Case 2:
                        .Txt_ID_PC_depreciacao = 0
                        .Txt_codigo_PC_depreciacao = ""
                        .Txt_descricao_PC_depreciacao = ""
                    Case 3:
                        .Txt_ID_PC_rateio = 0
                        .Txt_codigo_PC_rateio = ""
                        .Txt_descricao_PC_rateio = ""
                End Select
            End With
        ElseIf Plano_instituicao = True Then
                With frm_Instituicoes
                    Select Case Sit_REG
                        Case 1:
                            .Txt_ID_PC_instituicao = 0
                            .Txt_codigo_PC_instituicao = ""
                            .Txt_descricao_PC_instituicao = ""
                        Case 2:
                            .Txt_ID_PC_instituicao_rec = 0
                            .Txt_codigo_PC_instituicao_rec = ""
                            .Txt_descricao_PC_instituicao_rec = ""
                        Case 3:
                            .Txt_ID_PC = 0
                            .Txt_codigo_PC = ""
                            .Txt_descricao_PC = ""
                    End Select
                End With
            ElseIf Plano_opcoesgerais = True Then
                    With frmOpcoesGeral_PC
                        .Txt_ID_PC = 0
                        .Txt_codigo_PC = ""
                        .Txt_descricao_PC = ""
                    End With
                ElseIf Plano_Faturamento = True Then
                        If Sit_REG = 1 Then
                            With frmFaturamento_Prod_Serv_FamiliasDuplicata
                                .Txt_ID_PC = 0
                                .Txt_codigo_PC = ""
                                .Txt_descricao_PC = ""
                            End With
                        Else
                            With frmFaturamento_Prod_Serv_DI
                                .Txt_ID_PC = 0
                                .Txt_codigo_PC = ""
                                .Txt_descricao_PC = ""
                            End With
                        End If
                    ElseIf Plano_PCP = True Then
                            With frmprod_outras_despesas
                                .Txt_ID_PC = 0
                                .Txt_codigo_PC = ""
                                .Txt_descricao_PC = ""
                            End With
                        Else
                            With frmFamilia_financeiro
                                .Txt_ID_PC = 0
                                .Txt_codigo_PC = ""
                                .Txt_descricao_PC = ""
                            End With
End If
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_sair_Click()
On Error GoTo tratar_erro

Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro
  
Select Case KeyCode
    Case vbKeyF3: If Cmd_carregar.Enabled = True Then Cmd_carregar_Click
    Case vbKeyF4: If Cmd_carregar_vazio.Visible = True Then Cmd_carregar_vazio_Click
    'Case vbKeyF1: ProcAjuda
    Case vbKeyEscape: Cmd_sair_Click
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

If Plano_contas_produtos = True Or Plano_contas_familias = True Then
    Cmd_carregar.Left = 7470
    With Cmd_carregar_vazio
        .Left = 8970
        .Visible = True
    End With
Else
    Cmd_carregar.Left = 8970
    Cmd_carregar_vazio.Visible = False
End If

Permitido = False 'Pagar
Permitido1 = False 'Receber
IDlista = 0
If Plano_contas_produtos = True Then
    If Compras_Produtos = True Then Caption = "Compras - Produtos e serviços - Conta contábil"
    If Vendas_Produtos = True Then Caption = "Vendas - Produtos e serviços - Conta contábil"
    If Aplic = 1 Then Permitido = True Else Permitido1 = True
ElseIf Plano_contas_familias = True Then
        If Compras_Familia = False And Vendas_Familia = False And Qualidade_Familia = False Then Caption = "Engenharia - Famílias - Conta contábil"
        If Compras_Familia = True Then Caption = "Compras - Famílias - Conta contábil"
        If Vendas_Familia = True Then Caption = "Vendas - Famílias - Conta contábil"
        If Qualidade_Familia = True Then Caption = "Qualidade - Famílias - Conta contábil"
        If Aplic = 1 Then Permitido = True Else Permitido1 = True
    ElseIf Plano_centro_de_custo = True Then
            Caption = "Custos - Centro de custo - Conta contábil"
            Permitido = True
            Permitido1 = True
        ElseIf Plano_instituicao = True Then
                Caption = "Financeiro - Instituições - Tarifas - Conta contábil"
                Select Case Sit_REG
                    Case 1: Permitido = True
                    Case 2: Permitido1 = True
                    Case 3: If frm_Instituicoes.Cmb_operacao = "Débito" Then Permitido = True Else Permitido1 = True
                End Select
            ElseIf Financeiro_Contas_Pagar = True Then
                    Caption = "Financeiro - Contas a pagar - Conta contábil"
                    Permitido = True
                ElseIf Financeiro_Contas_Receber = True Then
                        Caption = "Financeiro - Contas a receber - Conta contábil"
                        Permitido1 = True
                    ElseIf Financeiro_Contas_Pagas = True Then
                            Caption = "Financeiro - Contas pagas - Conta contábil"
                            Permitido = True
                        ElseIf Financeiro_Contas_Recebidas = True Then
                                Caption = "Financeiro - Contas recebidas - Conta contábil"
                                Permitido1 = True
                            ElseIf Plano_opcoesgerais = True Then
                                    Caption = "Configurações do sistema - Opções gerais - Conta contábil"
                                    Permitido = True
                                ElseIf Plano_Faturamento = True Then
                                        If Sit_REG = 1 Then
                                            Caption = "Faturamento - Nota fiscal - Conta contábil"
                                            With frmFaturamento_Prod_Serv
                                                'Verifica se é nota de devolução
                                                Set TBAbrir = CreateObject("adodb.recordset")
                                                TBAbrir.Open "Select CFOP.* from tbl_Detalhes_Nota NFP INNER JOIN tbl_NaturezaOperacao CFOP ON CFOP.IDCountCfop = NFP.ID_CFOP where NFP.ID_nota = " & .txtID & " and CFOP.Devolucao = 'True'", Conexao, adOpenKeyset, adLockOptimistic
                                                If TBAbrir.EOF = False Then
                                                    'Fornecedor
                                                    If .opt_Saida.Value = True And Len(.txttipocliente) = 1 Then Permitido = True
                                                    'Cliente
                                                    If .opt_Entrada.Value = True And Len(.txttipocliente) = 2 Then Permitido1 = True
                                                Else
                                                    If Faturamento_NF_Saida = True Then Permitido1 = True Else Permitido = True
                                                End If
                                                TBAbrir.Close
                                            End With
                                        Else
                                            Caption = "Faturamento - Nota fiscal - Importação - Despesas - Conta contábil"
                                            Permitido = True
                                        End If
                                    ElseIf Plano_PCP = True Then
                                            Caption = "PCP - Gerenciamento de ordem - Cadastrar custo de outras despesas - Localizar despesas"
                                            Permitido = True
End If
ProcCarregaPC

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaPC()
On Error GoTo tratar_erro

With USTreeView1
    .Clear
    
    'Adicionando as chaves principais
    If Permitido = True Then
        If Financeiro_Contas_Pagas = True Then Texto = "Contas pagas" Else Texto = "Contas a pagar"
        Set Pagar = .Nodes.AddNode(Texto, "B", , True, , , , 0, vbRed)
        TextoFiltro = "Destino = 'P'"
    End If
    If Permitido1 = True Then
        Set Receber = .Nodes.AddNode("Contas a receber", "A", , True, , , , 0, vbBlue)
        If Financeiro_Contas_Recebidas = True Then Texto = "Contas recebidas" Else Texto = "Contas a receber"
        TextoFiltro = "Destino = 'R'"
    End If
    If Permitido = True And Permitido1 = True Then TextoFiltro = "(Destino = 'R' or Destino = 'P')"
    
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from tbl_familia where Codigo is not null and " & TextoFiltro & " order by Codigo", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Do While TBAbrir.EOF = False
            
            Descricao = TBAbrir!CODIGO & " - " & TBAbrir!Txt_descricao
            IDlista = TBAbrir!int_codfamilia
            Nivel = IIf(IsNull(TBAbrir!Nivel), 0, TBAbrir!Nivel)
           
            If Nivel = 8 Then
                If TBAbrir!Destino = "R" Then
                    .Nodes.AddNode Descricao, IDlista, , , , , , , , NivelR7
                Else
                    .Nodes.AddNode Descricao, IDlista, , , , , , , , NivelP7
                End If
            ElseIf Nivel = 7 Then
                    If TBAbrir!Destino = "R" Then
                        Set NivelR7 = .Nodes.AddNode(Descricao, IDlista, , , , , , , , NivelR6)
                    Else
                        Set NivelP7 = .Nodes.AddNode(Descricao, IDlista, , , , , , , , NivelP6)
                    End If
                ElseIf Nivel = 6 Then
                        If TBAbrir!Destino = "R" Then
                            Set NivelR6 = .Nodes.AddNode(Descricao, IDlista, , , , , , , , NivelR5)
                        Else
                            Set NivelP6 = .Nodes.AddNode(Descricao, IDlista, , , , , , , , NivelP5)
                        End If
                    ElseIf Nivel = 5 Then
                            If TBAbrir!Destino = "R" Then
                                Set NivelR5 = .Nodes.AddNode(Descricao, IDlista, , , , , , , , NivelR4)
                            Else
                                Set NivelP5 = .Nodes.AddNode(Descricao, IDlista, , , , , , , , NivelP4)
                            End If
                        ElseIf Nivel = 4 Then
                                If TBAbrir!Destino = "R" Then
                                    Set NivelR4 = .Nodes.AddNode(Descricao, IDlista, , , , , , , , NivelR3)
                                Else
                                    Set NivelP4 = .Nodes.AddNode(Descricao, IDlista, , , , , , , , NivelP3)
                                End If
                            ElseIf Nivel = 3 Then
                                    If TBAbrir!Destino = "R" Then
                                        Set NivelR3 = .Nodes.AddNode(Descricao, IDlista, , , , , , , , NivelR2)
                                    Else
                                        Set NivelP3 = .Nodes.AddNode(Descricao, IDlista, , , , , , , , NivelP2)
                                    End If
                                ElseIf Nivel = 2 Then
                                        If TBAbrir!Destino = "R" Then
                                            Set NivelR2 = .Nodes.AddNode(Descricao, IDlista, , , , , , , , NivelR1)
                                        Else
                                            Set NivelP2 = .Nodes.AddNode(Descricao, IDlista, , , , , , , , NivelP1)
                                        End If
                                    Else
                                        If TBAbrir!Destino = "R" Then
                                            Set NivelR1 = .Nodes.AddNode(Descricao, IDlista, , , , , , , , Receber)
                                        Else
                                            Set NivelP1 = .Nodes.AddNode(Descricao, IDlista, , , , , , , , Pagar)
                                        End If
            End If
            
            TBAbrir.MoveNext
        Loop
    End If
    TBAbrir.Close
    .ExpandAllNodes False
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USTreeView1_NodeSelected(Node As DrawSuite2022.TreeNode)
On Error GoTo tratar_erro

If IsNumeric(Node.key) = True Then IDlista = Node.key Else IDlista = 0

Set TBFI = CreateObject("adodb.recordset")
TBFI.Open "Select * from tbl_familia where int_codfamilia = " & IDlista, Conexao, adOpenKeyset, adLockOptimistic
If TBFI.EOF = False Then
    
    Select Case TBFI!Nivel
        Case 1: TextoFiltro = "Left(Codigo,1) = '" & Left(TBFI!CODIGO, 1) & "'"
        Case 2: TextoFiltro = "Left(Codigo,4) = '" & Left(TBFI!CODIGO, 4) & "'"
        Case 3: TextoFiltro = "Left(Codigo,7) = '" & Left(TBFI!CODIGO, 7) & "'"
        Case 4: TextoFiltro = "Left(Codigo,10) = '" & Left(TBFI!CODIGO, 10) & "'"
        Case 5: TextoFiltro = "Left(Codigo,13) = '" & Left(TBFI!CODIGO, 13) & "'"
        Case 6: TextoFiltro = "Left(Codigo,16) = '" & Left(TBFI!CODIGO, 16) & "'"
        Case 7: TextoFiltro = "Left(Codigo,19) = '" & Left(TBFI!CODIGO, 19) & "'"
        Case 8: TextoFiltro = "Left(Codigo,22) = '" & Left(TBFI!CODIGO, 22) & "'"
    End Select
    
    Set TBFIltro = CreateObject("adodb.recordset")
    TBFIltro.Open "Select * from tbl_familia where " & TextoFiltro & " order by Nivel", Conexao, adOpenKeyset, adLockOptimistic
    If TBFIltro.EOF = False Then
        TBFIltro.MoveLast
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from tbl_familia where int_codfamilia = " & IDlista & " and Nivel = " & TBFIltro!Nivel, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Cmd_carregar.Enabled = True
        Else
            Cmd_carregar.Enabled = False
        End If
        TBAbrir.Close
    End If
    TBFIltro.Close
End If
TBFI.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
