VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Begin VB.Form FrmCalendario 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   2505
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   2370
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   2505
   ScaleWidth      =   2370
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.MonthView calendario 
      Height          =   2520
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2385
      _ExtentX        =   4207
      _ExtentY        =   4445
      _Version        =   393216
      ForeColor       =   0
      BackColor       =   16777215
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MonthBackColor  =   16777215
      StartOfWeek     =   136380417
      TitleBackColor  =   8421504
      TitleForeColor  =   16777215
      TrailingForeColor=   255
      CurrentDate     =   39059
   End
End
Attribute VB_Name = "FrmCalendario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub calendario_DateClick(ByVal DateClicked As Date)
On Error GoTo tratar_erro

If Compras_Requisicao = True Then frmCompras_Requisicao.txtprazo.Text = Format(DateClicked, "dd/mm/yyyy")
If Compras_Cotacao = True Then
    With frmcompras_reqcot
        If Sit_Data = 1 Then .Txt_data_validade = Format(DateClicked, "dd/mm/yyyy") Else .txtPrazoentregaforn.Text = Format(DateClicked, "dd/mm/yyyy")
    End With
End If

If Compras_Pedido = True Then frmCompras_Pedido.txtprazo_item.Text = Format(DateClicked, "dd/mm/yyyy")
If Compras_Pedido = True And Sit_Data = 2 Then frmCompras_Pedido.txtPrazo_serv.Text = Format(DateClicked, "dd/mm/yyyy")
If Faturamento = True And Sit_Data = 1 Then frmFaturamento_Prod_Serv.txt_EntSai.Text = Format(DateClicked, "dd/mm/yyyy")
If Faturamento = True And Sit_Data = 2 Then frmFaturamento_Prod_Serv.txt_EntSai.Text = Format(DateClicked, "dd/mm/yyyy")
If Faturamento = True And Sit_Data = 3 Then frmFaturamento_Prod_Serv.txt_DtEmissao.Text = Format(DateClicked, "dd/mm/yyyy")

If Vendas_Carteira = True And Sit_Data = 1 Then frmVendas_carteira.txtDatavendas.Text = Format(DateClicked, "dd/mm/yyyy")
If Vendas_Carteira = True And Sit_Data = 2 Then frmVendas_carteira.txtdatafaturado.Text = Format(DateClicked, "dd/mm/yyyy")
If Vendas_Carteira = True And Sit_Data = 3 Then frmVendas_carteira.txtprazofinal.Text = Format(DateClicked, "dd/mm/yyyy")
If Vendas_PI = True Or Vendas_Proposta = True Then
    With IIf(Vendas_PI = True, frmVendas_PI, frmVendas_proposta)
        If Sit_Data = 1 Then
            .mskprazo.Text = Format(DateClicked, "dd/mm/yyyy")
        ElseIf Sit_Data = 2 Then
                .mskprazoservico.Text = Format(DateClicked, "dd/mm/yyyy")
            Else
                .Txt_data_retorno = Format(DateClicked, "dd/mm/yyyy")
        End If
    End With
End If
If Manutencao = True Then
    With frmManutencao
        If Sit_REG = 1 Then .txtData_Solicitacao = Format(DateClicked, "dd/mm/yyyy") Else .Txt_data_conclusao1 = Format(DateClicked, "dd/mm/yyyy")
    End With
End If
If Usuarios = True Then frmUsuarios.txtexpiracao.Text = Format(DateClicked, "dd/mm/yyyy")
If Funcionario = True And Sit_Data = 1 Then frmRH_Funcionarios.txtNascimento.Text = Format(DateClicked, "dd/mm/yyyy")
If Funcionario = True And Sit_Data = 2 Then frmRH_Funcionarios.txtData_adm.Text = Format(DateClicked, "dd/mm/yyyy")
If Funcionario = True And Sit_Data = 3 Then frmRH_Funcionarios.txtdata_desligado.Text = Format(DateClicked, "dd/mm/yyyy")
If Funcionario = True And Sit_Data = 4 Then frmRH_Funcionarios.txtDtCurso.Text = Format(DateClicked, "dd/mm/yyyy")
If RNC = True Then
    With frmQualidade_RNC
        Select Case Sit_Data
            Case 1: .txtfim = Format(DateClicked, "dd/mm/yyyy")
            Case 2: .txtdata7 = Format(DateClicked, "dd/mm/yyyy")
            Case 3: .txtData5 = Format(DateClicked, "dd/mm/yyyy")
            Case 4: .txtData6 = Format(DateClicked, "dd/mm/yyyy")
        End Select
    End With
End If
If Qualidade_PPAP_PSW = True Then
    With frmQualidadePPAP
        Select Case Sit_Data
            Case 1: .txtData.Text = Format(DateClicked, "dd/mm/yy")
            Case 2: .txtData2.Text = Format(DateClicked, "dd/mm/yy")
            Case 3: .txtData3.Text = Format(DateClicked, "dd/mm/yy")
        End Select
    End With
End If
If Qualidade_PPAP_Plano = True Then frmQualidadePPAP_PlanoControle.txtDataEnsaio = Format(DateClicked, "dd/mm/yyyy")
If Qualidade_PPAP_FMEA = True Then frmQualidadePPAP_FMEA.txtDatachave = Format(DateClicked, "dd/mm/yyyy")
If SolicitacaoAcao = True Then
    With frmCQ_SA
        Select Case Sit_Data
            Case 1: .txtfim.Text = Format(DateClicked, "dd/mm/yyyy")
            Case 2: .txtData_revisar.Text = Format(DateClicked, "dd/mm/yyyy")
            Case 3: .txtData1.Text = Format(DateClicked, "dd/mm/yyyy")
            Case 4: .txtData2.Text = Format(DateClicked, "dd/mm/yyyy")
            Case 5: .txtData3.Text = Format(DateClicked, "dd/mm/yyyy")
            Case 6: .txtData4.Text = Format(DateClicked, "dd/mm/yyyy")
            Case 7: .txtData5.Text = Format(DateClicked, "dd/mm/yyyy")
            Case 8: .txtData6.Text = Format(DateClicked, "dd/mm/yyyy")
            Case 9: .txtdata7.Text = Format(DateClicked, "dd/mm/yyyy")
            Case 10: .txtData_riscos.Text = Format(DateClicked, "dd/mm/yyyy")
        End Select
    End With
End If
If Troca_Duplicata = True Then frm_trocaduplicata.Txt_data_operacao = Format(DateClicked, "dd/mm/yyyy")
If Financeiro_Contas_Recebidas = True Then frmContas_recebidas_dados_desconto.Txt_data_recebimento = Format(DateClicked, "dd/mm/yyyy")
If Engenharia_Normas = True Then frmNorma.Txt_data_rev = Format(DateClicked, "dd/mm/yyyy")
If Qualidade_sistema = True Then frmCQ_sistema.Txt_data_rev = Format(DateClicked, "dd/mm/yyyy")
If Engenharia = True Then frmproj_produto.txtPPAP_Datarev = Format(DateClicked, "dd/mm/yyyy")
If Compras_Fornecedores = True Then frmCompras_fornecedores.txtdata_venc = Format(DateClicked, "dd/mm/yyyy")
If Vendas_Programacao = True Then frmVendas_programacao.Txt_data_negociada = Format(DateClicked, "dd/mm/yyyy")

If Vendas_Analise = True Then
    With frmVendas_analise_prazos
        Select Case Sit_Data
            Case 1: .txtEngenharia = Format(DateClicked, "dd/mm/yyyy")
            Case 2: .txtProcesso = Format(DateClicked, "dd/mm/yyyy")
            Case 3: .txtPCP = Format(DateClicked, "dd/mm/yyyy")
            Case 4: .txtQualidade = Format(DateClicked, "dd/mm/yyyy")
            Case 5: .txtCompras = Format(DateClicked, "dd/mm/yyyy")
        End Select
    End With
End If

If Outros_solicitacaoPCP = True Then frmOutros_Solicitacao_PCP.txtprazo = Format(DateClicked, "dd/mm/yyyy")
If Estoque_recebimento = True Then
    If Sit_Data = 1 Then frmEstoque_Recebimento.txtDataemissao = Format(DateClicked, "dd/mm/yyyy")
    If Sit_Data = 2 Then frmEstoque_Recebimento.Txt_data_recebimento = Format(DateClicked, "dd/mm/yyyy")
End If

Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

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

Private Sub Form_Load()
On Error GoTo tratar_erro

calendario.Value = Date
FrmCalendario.Caption = "Calendário de " & Format(Date, "yyyy")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
