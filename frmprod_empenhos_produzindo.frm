VERSION 5.00
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmprod_empenhos_produzindo 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "PCP | Lista de empenhos produzindo"
   ClientHeight    =   9585
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   14595
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmprod_empenhos_produzindo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmprod_empenhos_produzindo.frx":000C
   ScaleHeight     =   9585
   ScaleWidth      =   14595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   405
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   14595
      _ExtentX        =   25744
      _ExtentY        =   714
      DibPicture      =   "frmprod_empenhos_produzindo.frx":0316
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
      Icon            =   "frmprod_empenhos_produzindo.frx":65FA
   End
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   3
      Top             =   9180
      Width           =   14595
      _ExtentX        =   25744
      _ExtentY        =   714
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   150
      TabIndex        =   0
      Top             =   8700
      Width           =   14325
      _ExtentX        =   25268
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
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   0
      TabIndex        =   1
      Top             =   420
      Width           =   14595
      _ExtentX        =   25744
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
      ButtonCaption1  =   "Excluir"
      ButtonEnabled1  =   0   'False
      ButtonIconSize1 =   32
      ButtonToolTipText1=   "Excluir (F4)"
      ButtonKey1      =   "1"
      ButtonAlignment1=   2
      BeginProperty ButtonFont1 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft1     =   2
      ButtonTop1      =   2
      ButtonWidth1    =   45
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
      ButtonLeft2     =   49
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft3     =   53
      ButtonTop3      =   2
      ButtonWidth3    =   41
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft4     =   96
      ButtonTop4      =   2
      ButtonWidth4    =   30
      ButtonHeight4   =   21
      ButtonUseMaskColor4=   0   'False
      ButtonEnabled5  =   0   'False
      ButtonIconSize5 =   32
      ButtonKey5      =   "5"
      ButtonAlignment5=   2
      BeginProperty ButtonFont5 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState5    =   5
      ButtonLeft5     =   128
      ButtonTop5      =   2
      ButtonWidth5    =   24
      ButtonHeight5   =   24
      ButtonUseMaskColor5=   0   'False
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   5730
         Top             =   120
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmprod_empenhos_produzindo.frx":6616
         Count           =   1
      End
   End
   Begin MSComctlLib.ListView ListaEmpenhoProduzindo 
      Height          =   7215
      Left            =   150
      TabIndex        =   2
      Top             =   1470
      Width           =   14325
      _ExtentX        =   25268
      _ExtentY        =   12726
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   14
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "N"
         Text            =   "ID_carteira"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Object.Tag             =   "D"
         Text            =   "Data"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "Responsável"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Ped. int./SPR"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Object.Tag             =   "N"
         Text            =   "Rev."
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Object.Tag             =   "T"
         Text            =   "Cliente"
         Object.Width           =   2886
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Object.Tag             =   "T"
         Text            =   "Vendedor interno"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Object.Tag             =   "T"
         Text            =   "Vendedor externo"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   9
         Object.Tag             =   "D"
         Text            =   "Pr. final"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Object.Tag             =   "N"
         Text            =   "Qt. vend."
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   11
         Object.Tag             =   "N"
         Text            =   "Qt. emp."
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   12
         Object.Tag             =   "N"
         Text            =   "Qt. entr."
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   13
         Object.Tag             =   "N"
         Text            =   "Saldo"
         Object.Width           =   1587
      EndProperty
   End
End
Attribute VB_Name = "frmprod_empenhos_produzindo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Sql_Estoque_Movimentacao_empenho As String 'OK

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyF4: ProcExcluirEmpenho
    'Case vbKeyF1: ProcAjuda
    Case vbKeyEscape: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 14595, 4, True
Direitos
ProcLimpaVariaveisPrincipais
ProcCarregaListaEmpprod

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaListaEmpprod()
On Error GoTo tratar_erro

ListaEmpenhoProduzindo.ListItems.Clear
CamposFiltro = "PP.ID, PP.IDcarteira, PP.Data1 as Dataemp, PP.Responsavel as Respemp, PP.Qtde_Empenho, PP.Qtde_entrada, VP.Ncotacao, VP.Revisao, OPCP.Requisicaotexto, VP.Cliente, VP.vend_int, VP.vend_ext, VC.Codigo, VC.Qtde_produzir, VC.Prazofinal"
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select " & CamposFiltro & " from (((Producao_Pedidos PP INNER JOIN Producao P ON P.Ordem = PP.Ordem) INNER JOIN vendas_carteira VC ON PP.IDCarteira =  VC.Codigo) LEFT JOIN vendas_proposta VP ON VP.cotacao = VC.Cotacao) LEFT JOIN Outros_SolicitacaoPCP OPCP ON OPCP.ID = VC.ID_Solicitacao where P.Desenho = '" & frmprod.TxtEstCodigo & "' and PP.Qtde_empenho - ISNULL(PP.Qtde_entrada, 0) > 0", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBLISTA.EOF = False
        With ListaEmpenhoProduzindo.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = TBLISTA!IDcarteira
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Dataemp), "", Format(TBLISTA!Dataemp, "dd/mm/yy"))
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Respemp), "", TBLISTA!Respemp)
            If IsNull(TBLISTA!Ncotacao) = False Then
                .Item(.Count).SubItems(4) = TBLISTA!Ncotacao
                .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!Revisao), "", TBLISTA!Revisao)
                .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA!Cliente), "", TBLISTA!Cliente)
                .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA!vend_int), "", TBLISTA!vend_int)
                .Item(.Count).SubItems(8) = IIf(IsNull(TBLISTA!Vend_ext), "", TBLISTA!Vend_ext)
            Else
                .Item(.Count).SubItems(4) = TBLISTA!Requisicaotexto
            End If
            .Item(.Count).SubItems(9) = IIf(IsNull(TBLISTA!PrazoFinal), "", Format(TBLISTA!PrazoFinal, "dd/mm/yy"))
            .Item(.Count).SubItems(10) = IIf(IsNull(TBLISTA!Qtde_produzir), "", TBLISTA!Qtde_produzir)
            valor = IIf(IsNull(TBLISTA!Qtde_empenho), 0, TBLISTA!Qtde_empenho)
            .Item(.Count).SubItems(11) = valor
            Valor1 = IIf(IsNull(TBLISTA!Qtde_entrada), 0, TBLISTA!Qtde_entrada)
            .Item(.Count).SubItems(12) = Valor1
            .Item(.Count).SubItems(13) = valor - Valor1
        End With
        TBLISTA.MoveNext
        Contador = Contador + 1
        PBLista.Value = Contador
    Loop
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaEmpenhoProduzindo_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With ListaEmpenhoProduzindo
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If .ListItems.Item(InitFor).ListSubItems(12) <> 0 Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView ListaEmpenhoProduzindo, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaEmpenhoProduzindo_DblClick()
On Error GoTo tratar_erro

With ListaEmpenhoProduzindo
    If .ListItems.Count = 0 Then Exit Sub
    ProcVerifQtdeFaturadaProdServ .SelectedItem.ListSubItems(1), frmprod.TxtEstCodigo, False
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaEmpenhoProduzindo_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With ListaEmpenhoProduzindo
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If .ListItems.Item(InitFor).ListSubItems(12) <> 0 Then
                USMsgBox ("Não é permitido excluir este empenho, pois já foi dado entrada no estoque."), vbExclamation, "CAPRIND v5.0"
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
        End If
    Next InitFor
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcExcluirEmpenho
    'Case 3: ProcAjuda
    Case 4: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcExcluirEmpenho()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With ListaEmpenhoProduzindo
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir esse(s) empenho(s)?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            
            Conexao.Execute "DELETE from producao_pedidos where ID = " & .ListItems(InitFor)
            '==================================
            Modulo = "PCP/Lista de empenhos produzindo"
            Evento = "Excluir empenho produção"
            ID_documento = .ListItems(InitFor)
            Documento = "Cod. carteira: " & .ListItems(InitFor).SubItems(1) & " - Pedido: " & .ListItems(InitFor).SubItems(4) & " - Rev.: " & .ListItems(InitFor).SubItems(5)
            Documento = ""
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) empenho(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Empenho(s) excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcCarregaListaEmpprod
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
