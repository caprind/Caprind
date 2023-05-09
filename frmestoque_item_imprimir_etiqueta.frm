VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmestoque_item_imprimir_etiqueta 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Etiqueta de identificação"
   ClientHeight    =   4185
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   3780
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
   Icon            =   "frmestoque_item_imprimir_etiqueta.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   3780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   14
      Top             =   3780
      Width           =   3780
      _ExtentX        =   6668
      _ExtentY        =   714
   End
   Begin DrawSuite2022.USButton btnImprimir 
      Height          =   855
      Left            =   330
      TabIndex        =   13
      ToolTipText     =   "Visualizar relatório para impressão"
      Top             =   2700
      Width           =   3165
      _ExtentX        =   5583
      _ExtentY        =   1508
      DibPicture      =   "frmestoque_item_imprimir_etiqueta.frx":1042
      BorderColor     =   5263559
      BorderColorDisabled=   13160660
      BorderColorDown =   4013465
      BorderColorOver =   4408288
      Caption         =   "Imprimir etiqueta"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
      ForeColorDown   =   16777215
      ForeColorOver   =   16777215
      GradientColor1  =   5263559
      GradientColor2  =   5263559
      GradientColor3  =   5263559
      GradientColor4  =   5263559
      GradientColorDisabled1=   13160660
      GradientColorDisabled2=   13160660
      GradientColorDisabled3=   13160660
      GradientColorDisabled4=   13160660
      GradientColorDown1=   4013465
      GradientColorDown2=   4013465
      GradientColorDown3=   4013465
      GradientColorDown4=   4013465
      GradientColorOver1=   4408288
      GradientColorOver2=   4408288
      GradientColorOver3=   4408288
      GradientColorOver4=   4408288
      PicAlign        =   7
      PicSize         =   3
      PicSizeH        =   32
      PicSizeW        =   32
      ShowFocusRect   =   0   'False
      ShowFocusRectDown=   0   'False
      Theme           =   4
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   465
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   3780
      _ExtentX        =   6668
      _ExtentY        =   820
      DibPicture      =   "frmestoque_item_imprimir_etiqueta.frx":54F63
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
      Icon            =   "frmestoque_item_imprimir_etiqueta.frx":A8E84
      ShowMaximize    =   0   'False
      ShowMinimize    =   0   'False
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1905
      Left            =   330
      TabIndex        =   6
      Top             =   660
      Width           =   3180
      Begin VB.ComboBox Cmb_posicao 
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
         Height          =   330
         ItemData        =   "frmestoque_item_imprimir_etiqueta.frx":A9ED6
         Left            =   2220
         List            =   "frmestoque_item_imprimir_etiqueta.frx":A9F16
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         ToolTipText     =   "Posição inicial."
         Top             =   910
         Width           =   765
      End
      Begin VB.TextBox Txt_qtde_emb 
         Alignment       =   2  'Center
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
         Height          =   315
         Left            =   180
         TabIndex        =   2
         ToolTipText     =   "Quantidade por embalagem."
         Top             =   910
         Width           =   1095
      End
      Begin VB.TextBox Txt_qtde_RE 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Left            =   1590
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade da RE."
         Top             =   360
         Width           =   1395
      End
      Begin VB.TextBox Txt_RE 
         Alignment       =   2  'Center
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
         Height          =   315
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         ToolTipText     =   "Número da RE."
         Top             =   360
         Width           =   1395
      End
      Begin VB.TextBox Txt_n_copias 
         Alignment       =   2  'Center
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
         Height          =   315
         Left            =   1290
         TabIndex        =   3
         ToolTipText     =   "Número de cópias."
         Top             =   910
         Width           =   915
      End
      Begin VB.CheckBox chk_imprimir_qtde_total_RE 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Imprimir qtde. total da RE"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   480
         TabIndex        =   5
         Top             =   1560
         Width           =   2175
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Pos."
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   2445
         TabIndex        =   11
         Top             =   720
         Width           =   315
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Qtde. emb."
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
         Left            =   270
         MousePointer    =   1  'Arrow
         TabIndex        =   10
         Top             =   720
         Width           =   915
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Qtde. RE"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1957
         MousePointer    =   1  'Arrow
         TabIndex        =   9
         Top             =   180
         Width           =   660
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "RE"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   780
         MousePointer    =   1  'Arrow
         TabIndex        =   8
         Top             =   180
         Width           =   195
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Nº cópias"
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
         Left            =   1357
         MousePointer    =   1  'Arrow
         TabIndex        =   7
         Top             =   720
         Width           =   780
      End
   End
End
Attribute VB_Name = "frmestoque_item_imprimir_etiqueta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnImprimir_Click()
On Error GoTo tratar_erro

ProcImprimir

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro
  
Select Case KeyCode
    Case vbKeyF5: ProcImprimir
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

Desenho = ""

If Faturamento = True Then
    Label2(0).Caption = "OF"
    txt_RE.ToolTipText = "Número da ordem de faturamento."
    Label3.Caption = "Qtde. OF"
    Txt_qtde_RE.ToolTipText = "Quantidade da ordem de faturamento."
    chk_imprimir_qtde_total_RE.Caption = "Imprimir qtde. total da OF"
    With frmEstoque_Ordem_Faturamento
        txt_RE = .txtID
        If .SSFaturamento.Tab = 2 Then
            Txt_qtde_RE = .ListaProdutos.SelectedItem.ListSubItems(7)
            Desenho = .ListaProdutos.SelectedItem.ListSubItems(1)
        Else
            With Txt_qtde_emb
                .Locked = True
                .TabStop = False
            End With
            With chk_imprimir_qtde_total_RE
                .Value = 1
                .Enabled = False
            End With
        End If
    End With
    If Permitido = False Then Cmb_posicao.Enabled = False
Else
    If Permitido = False Then
        If Inspecao_recebimento = True Then
            With frmCompras_recebimento.ListProdReceb.SelectedItem
                txt_RE = .ListSubItems(5)
                Txt_qtde_RE = .ListSubItems(4)
                Desenho = .ListSubItems(1)
            End With
        ElseIf Estoque_recebimento = True Then
                With frmEstoque_Recebimento
                    txt_RE = .Lista_movimentacao.SelectedItem.ListSubItems(1)
                    Txt_qtde_RE = .Lista_movimentacao.SelectedItem.ListSubItems(4)
                    Desenho = .txtCodigo
                End With
            Else
                With frmestoque_item
                    txt_RE = RE '.txtlocalização
                    Txt_qtde_RE = .Txt_qtde_estoqueRE
                    Desenho = .Lista.SelectedItem.ListSubItems(3)
                End With
        End If
        Cmb_posicao.Enabled = False
    Else
        If Inspecao_recebimento = True Then
            With frmCompras_recebimento.ListProdReceb
                TextoFiltro = "Compras_recebimento EC where EC.Id = " & .SelectedItem
                CamposFiltro = "EC.IDestoque, EC.Enc"
                Desenho = .SelectedItem.ListSubItems(1)
            End With
        ElseIf Estoque_recebimento = True Then
                With frmEstoque_Recebimento
                    TextoFiltro = "estoque_controle EC INNER JOIN estoque_movimentacao EM ON EC.IDEstoque = EM.idestoque where EM.IDoperacao = " & .Lista_movimentacao.SelectedItem
                    CamposFiltro = "EC.IDestoque, EM.Entrada"
                    Desenho = .txtCodigo
                End With
            Else
                With frmestoque_item.Lista
                    TextoFiltro = "estoque_controle EC where EC.IdEstoque = " & .SelectedItem
                    CamposFiltro = "EC.IdEstoque, EC.estoque_real"
                    Desenho = .SelectedItem.ListSubItems(3)
                End With
        End If
        Set TBEstoque = CreateObject("adodb.recordset")
        TBEstoque.Open "Select " & CamposFiltro & " as Quantsolicitado from " & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
        If TBEstoque.EOF = False Then
            txt_RE = TBEstoque!IDEstoque
            Txt_qtde_RE = IIf(IsNull(TBEstoque!QuantSolicitado), 0, Format(TBEstoque!QuantSolicitado, "###,##0.0000"))
        End If
        TBEstoque.Close
    End If
End If
If Desenho <> "" Then Txt_qtde_emb = Format(CarregaQtdeEmbProd(Desenho), "###,##0.0000")
   
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro

Acao = "visualizar impressão"
valor = IIf(Txt_qtde_RE = "", 0, Txt_qtde_RE)
If Faturamento = False Then
    If valor <= 0 Then
        USMsgBox ("Não é permitido gerar etiqueta para este RE, pois não existe quantidade disponível em estoque."), vbExclamation, "CAPRIND v5.0"
        Exit Sub
    End If
End If
If Txt_qtde_emb.Locked = False Then
    Valor1 = IIf(Txt_qtde_emb = "", 0, Txt_qtde_emb)
    If Valor1 <= 0 Then
        NomeCampo = "a quantidade por embalagem"
        ProcVerificaAcao
        Txt_qtde_emb.SetFocus
        Exit Sub
    End If
    If Valor1 > valor Then
        USMsgBox ("A quantidade por embalagem não pode ser maior que a quantidade " & IIf(Faturamento = False, "do RE", "da OF") & ", favor alterar."), vbExclamation, "CAPRIND v5.0"
        Txt_qtde_emb.SetFocus
        Exit Sub
    End If
End If
Cont = IIf(Txt_n_copias = "", 0, Txt_n_copias)
If Cont <= 0 Then
    NomeCampo = "o número de cópias"
    ProcVerificaAcao
    Txt_n_copias.SetFocus
    Exit Sub
End If

If Cmb_posicao.Enabled = False Then
    Conexao.Execute "DELETE from Compras_Recebimento_Relatorios where responsavel = '" & pubUsuario & "' and modulo = '" & Formulario & "'"
    Do While Cont > 0
        If chk_imprimir_qtde_total_RE.Value = 1 Then
            If Txt_qtde_emb.Locked = True Then
                Set TBFI = CreateObject("adodb.recordset")
                TBFI.Open "Select Int_codigo, int_Cod_Produto, int_Qtd from tbl_Detalhes_Nota where ID_nota = " & txt_RE, Conexao, adOpenKeyset, adLockOptimistic
                If TBFI.EOF = False Then
                    Do While TBFI.EOF = False
                        valor = TBFI!int_Qtd
                        Valor1 = Format(CarregaQtdeEmbProd(TBFI!int_Cod_Produto), "###,##0.0000")
                        Do While valor > 0
                            ProcGravarEtiqueta IIf(valor > Valor1, Valor1, valor)
                            valor = valor - Valor1
                        Loop
                        TBFI.MoveNext
                    Loop
                End If
                TBFI.Close
            Else
                valor = Txt_qtde_RE
                Do While valor > 0
                    ProcGravarEtiqueta IIf(valor > Valor1, Valor1, valor)
                    valor = valor - Valor1
                Loop
            End If
        Else
            ProcGravarEtiqueta Valor1
        End If
        Cont = Cont - 1
    Loop
    If Inspecao_recebimento = True Then
        NomeRel = "CQ_inspecao recebimento_identificacao.rpt"
        ProcImprimirRel "{Compras_Recebimento_Relatorios.Responsavel}= '" & pubUsuario & "' and {Compras_Recebimento_Relatorios.Modulo} = '" & Formulario & "' and {Compras_recebimento.IDestoque} = " & txt_RE, ""
    ElseIf Faturamento = True Then
            NomeRel = "Estoque_ordemfaturamento_identificacao.rpt"
            ProcImprimirRel "{Compras_Recebimento_Relatorios.Responsavel}= '" & pubUsuario & "' and {Compras_Recebimento_Relatorios.Modulo} = '" & Formulario & "'", ""
        Else
            NomeRel = "Estoque_identificacao.rpt"
            ProcImprimirRel "{Compras_Recebimento_Relatorios.Responsavel}= '" & pubUsuario & "' and {Compras_Recebimento_Relatorios.Modulo} = '" & Formulario & "'", ""
    End If
Else
    If Cmb_posicao = "" Then
        NomeCampo = "a posição inicial"
        ProcVerificaAcao
        Cmb_posicao.SetFocus
        Exit Sub
    End If
    
    Conexao.Execute "DELETE from etiqueta where Modulo = '" & Formulario & "' and Responsavel = '" & pubUsuario & "'"
    
    Posicao = Cmb_posicao.Text
    Cont = Cmb_posicao.Text
    Set TBFI = CreateObject("adodb.recordset")
    TBFI.Open "Select posicao from etiqueta where posicao <> 0 and posicao <> " & Cont & " and Modulo = '" & Formulario & "' and Responsavel = '" & pubUsuario & "' order by posicao", Conexao, adOpenKeyset, adLockOptimistic
    If TBFI.EOF = False Then
        TBFI.MoveLast
        If TBFI!Posicao < Cont Then
            i = TBFI!Posicao + 1
            If i >= 20 Then i = 1
        Else
            i = 1
        End If
    Else
        i = 1
    End If
    Do While i <> Cont
        Set TBGravar = CreateObject("adodb.recordset")
        TBGravar.Open "Select * from etiqueta where posicao = " & i & " and Modulo = '" & Formulario & "' and Responsavel = '" & pubUsuario & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBGravar.EOF = True Then TBGravar.AddNew
        TBGravar!Posicao = i
        TBGravar!Modulo = Formulario
        TBGravar!Responsavel = pubUsuario
        TBGravar.Update
        TBGravar.Close
        i = i + 1
    Loop
    TBFI.Close
    
    Posicao = Cmb_posicao
    Cont = Txt_n_copias
    Do While Cont > 0
        If chk_imprimir_qtde_total_RE.Value = 1 Then
            If Txt_qtde_emb.Locked = True Then
                Set TBFI = CreateObject("adodb.recordset")
                TBFI.Open "Select Int_codigo, int_Cod_Produto, int_Qtd from tbl_Detalhes_Nota where ID_nota = " & txt_RE, Conexao, adOpenKeyset, adLockOptimistic
                If TBFI.EOF = False Then
                    Do While TBFI.EOF = False
                        valor = TBFI!int_Qtd
                        Valor1 = Format(CarregaQtdeEmbProd(TBFI!int_Cod_Produto), "###,##0.0000")
                        Do While valor > 0
                            ProcGravarEtiquetaPers IIf(valor > Valor1, Valor1, valor)
                            valor = valor - Valor1
                        Loop
                        TBFI.MoveNext
                    Loop
                End If
                TBFI.Close
            Else
                valor = Txt_qtde_RE
                Do While valor > 0
                    ProcGravarEtiquetaPers IIf(valor > Valor1, Valor1, valor)
                    valor = valor - Valor1
                Loop
            End If
        Else
            ProcGravarEtiquetaPers Valor1
        End If
        Cont = Cont - 1
    Loop
    
    contador = 9999999
    Do While contador > 0
        contador = contador - 1
    Loop
    
    ProcImprimirRel "{etiqueta.Responsavel}= '" & pubUsuario & "' and {etiqueta.Modulo} = '" & Formulario & "'", ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcGravarEtiqueta(Qtde As Double)
On Error GoTo tratar_erro

Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from Compras_Recebimento_Relatorios", Conexao, adOpenKeyset, adLockOptimistic
TBGravar.AddNew
If Inspecao_recebimento = True Then
    TBGravar!ID_recebimento = frmCompras_recebimento.ListProdReceb.SelectedItem
ElseIf Faturamento = True Then
        If Txt_qtde_emb.Locked = True Then TBGravar!ID_recebimento = TBFI!Int_codigo Else TBGravar!ID_recebimento = frmFaturamento_Prod_Serv.ListaProdutos.SelectedItem
    Else
        TBGravar!ID_recebimento = txt_RE
End If
TBGravar!Modulo = Formulario
TBGravar!Responsavel = pubUsuario
TBGravar!Qtde = Qtde
TBGravar.Update
TBGravar.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcGravarEtiquetaPers(Qtde As Double)
On Error GoTo tratar_erro

Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from etiqueta", Conexao, adOpenKeyset, adLockOptimistic
TBGravar.AddNew
If Inspecao_recebimento = True Then
    NomeRel = "CQ_inspecao recebimento_identificacao_personalizada.rpt"
    TBGravar!ID_nome = frmCompras_recebimento.ListProdReceb.SelectedItem
ElseIf Estoque_recebimento = True Then
        NomeRel = "Estoque_identificacao_personalizada.rpt"
        Set TBEstoque = CreateObject("adodb.recordset")
        TBEstoque.Open "Select EC.IDestoque, EM.Documento, EM.Operacao from estoque_controle EC INNER JOIN estoque_movimentacao EM ON EC.IDEstoque = EM.idestoque where EM.IDoperacao = " & frmEstoque_Recebimento.Lista_movimentacao.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
        If TBEstoque.EOF = False Then
            TBGravar!ID_nome = TBEstoque!IDEstoque
            TBGravar!Documento = TBEstoque!Documento
            TBGravar!Endereco = TBEstoque!Operacao
            TBGravar!ID_lista = frmEstoque_Recebimento.Listprod.SelectedItem
        End If
        TBEstoque.Close
    ElseIf Faturamento = True Then
            NomeRel = "Estoque_ordemfaturamento_identificacao_personalizada.rpt"
            If Txt_qtde_emb.Locked = True Then TBGravar!ID_nome = TBFI!Int_codigo Else TBGravar!ID_nome = frmFaturamento_Prod_Serv.ListaProdutos.SelectedItem
        Else
            NomeRel = "Estoque_identificacao_personalizada.rpt"
            TBGravar!ID_nome = frmestoque_item.Lista.SelectedItem
            TBGravar!Documento = "NÃO"
End If
TBGravar!Nome = Qtde
TBGravar!Posicao = Posicao
TBGravar!Modulo = Formulario
TBGravar!Responsavel = pubUsuario
TBGravar.Update
TBGravar.Close

Posicao = Posicao + 1
If Posicao = 20 Then Posicao = 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

If Cmb_posicao.Enabled = False Then Conexao.Execute "DELETE from Compras_Recebimento_Relatorios where responsavel = '" & pubUsuario & "' and modulo = '" & Formulario & "'" Else Conexao.Execute "DELETE from etiqueta where Modulo = '" & Formulario & "' and Responsavel = '" & pubUsuario & "'"
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_n_copias_Change()
On Error GoTo tratar_erro

If Txt_n_copias <> "" Then
    VerifNumero = Txt_n_copias
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_n_copias = ""
        Txt_n_copias.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_qtde_emb_Change()
On Error GoTo tratar_erro

If Txt_qtde_emb <> "" Then
    VerifNumero = Txt_qtde_emb
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_qtde_emb = ""
        Txt_qtde_emb.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_qtde_emb_LostFocus()
On Error GoTo tratar_erro

 Txt_qtde_emb = Format(Txt_qtde_emb, "###,##0.0000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Function CarregaQtdeEmbProd(Codinterno As String) As Double
On Error GoTo tratar_erro

If Txt_qtde_emb.Locked = True Then CarregaQtdeEmbProd = valor Else CarregaQtdeEmbProd = Txt_qtde_RE
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select Qtde_embalagem from projproduto where Desenho = '" & Codinterno & "' and Qtde_embalagem is not null and Qtde_embalagem <> 0 and Qtde_embalagem <> N''", Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    CarregaQtdeEmbProd = TBProduto!Qtde_embalagem
End If
TBProduto.Close

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function
