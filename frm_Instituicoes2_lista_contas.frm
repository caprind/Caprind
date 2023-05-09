VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form frm_Instituicoes2_lista_contas 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Financeiro - Instituições - Lista de contas"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12420
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   12420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Txt_valor_movimentacao 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   2355
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Valor da movimentação."
      Top             =   1080
      Width           =   1620
   End
   Begin DrawSuite2022.USProgressBar PbLista 
      Height          =   255
      Left            =   60
      TabIndex        =   3
      Top             =   6900
      Width           =   12315
      _ExtentX        =   21722
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
      SearchText      =   ""
      Value           =   0
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   4
      Top             =   0
      Width           =   12315
      _ExtentX        =   21722
      _ExtentY        =   1720
      ButtonCount     =   3
      GradientColor2  =   14737632
      GradientColorOverRight1=   16315633
      GradientColorOverRight2=   15195350
      GripperColor    =   15195350
      IsStrech        =   -1  'True
      RightColor1     =   0
      RightColor2     =   0
      ShowEndPanel    =   0   'False
      Theme           =   1
      ButtonCaption1  =   "Ajuda"
      ButtonEnabled1  =   0   'False
      ButtonIconSize1 =   32
      ButtonToolTipText1=   "Ajuda (F1)"
      ButtonKey1      =   "8"
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
      ButtonWidth1    =   41
      ButtonHeight1   =   21
      ButtonUseMaskColor1=   0   'False
      ButtonCaption2  =   "Sair"
      ButtonEnabled2  =   0   'False
      ButtonIconSize2 =   32
      ButtonToolTipText2=   "Sair (Esc)"
      ButtonKey2      =   "9"
      ButtonAlignment2=   2
      BeginProperty ButtonFont2 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft2     =   45
      ButtonTop2      =   2
      ButtonWidth2    =   30
      ButtonHeight2   =   21
      ButtonUseMaskColor2=   0   'False
      ButtonEnabled3  =   0   'False
      ButtonIconSize3 =   32
      ButtonKey3      =   "10"
      ButtonAlignment3=   2
      BeginProperty ButtonFont3 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState3    =   5
      ButtonLeft3     =   77
      ButtonTop3      =   2
      ButtonWidth3    =   24
      ButtonHeight3   =   24
      ButtonUseMaskColor3=   0   'False
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   6750
         Top             =   60
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frm_Instituicoes2_lista_contas.frx":0000
         Count           =   1
      End
   End
   Begin MSComctlLib.ListView Lista_pagas 
      Height          =   5415
      Left            =   60
      TabIndex        =   1
      Top             =   1470
      Width           =   12315
      _ExtentX        =   21722
      _ExtentY        =   9551
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
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
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Object.Tag             =   "D"
         Text            =   "Dt. emissão"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Object.Tag             =   "D"
         Text            =   "Dt. vencto."
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Object.Tag             =   "N"
         Text            =   "Valor"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Nº documento"
         Object.Width           =   1941
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Object.Tag             =   "T"
         Text            =   "Parcela"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Object.Tag             =   "T"
         Text            =   "Fornecedor"
         Object.Width           =   6948
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Object.Tag             =   "N"
         Text            =   "Vlr. pago"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   8
         Object.Tag             =   "D"
         Text            =   "Dt. pagto."
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Object.Tag             =   "T"
         Text            =   "Responsável"
         Object.Width           =   2646
      EndProperty
   End
   Begin MSComctlLib.ListView Lista_recebidas 
      Height          =   5415
      Left            =   60
      TabIndex        =   2
      Top             =   1470
      Width           =   12315
      _ExtentX        =   21722
      _ExtentY        =   9551
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
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
      NumItems        =   11
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Object.Tag             =   "D"
         Text            =   "Dt. emissão"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Object.Tag             =   "D"
         Text            =   "Dt. venc."
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Object.Tag             =   "N"
         Text            =   "Valor"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Nº documento"
         Object.Width           =   1941
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Tag             =   "N"
         Text            =   "Nota fiscal"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Object.Tag             =   "T"
         Text            =   "Parcela"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Object.Tag             =   "T"
         Text            =   "Cliente"
         Object.Width           =   5362
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Object.Tag             =   "N"
         Text            =   "Vlr. receb."
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   9
         Object.Tag             =   "D"
         Text            =   "Dt. receb."
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Object.Tag             =   "T"
         Text            =   "Responsável"
         Object.Width           =   2646
      EndProperty
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Valor da movimentação :"
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
      Height          =   210
      Left            =   180
      TabIndex        =   5
      Top             =   1080
      Width           =   2100
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frm_Instituicoes2_lista_contas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro
  
Select Case KeyCode
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

ProcCarregaToolBar1 Me, 12315, 3, True

With frm_Instituicoes
    Txt_valor_movimentacao = .Lst_extrato.SelectedItem.ListSubItems(3)
    Set TBFluxo = CreateObject("adodb.recordset")
    TBFluxo.Open "Select * from Tbl_Fluxo_de_Caixa where IDFluxo = " & .Lst_extrato.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
    If TBFluxo.EOF = False Then
        If TBFluxo!Operacao = "Débito" Then
            Caption = "Financeiro - Instituições - Lista de contas pagas"
            Lista_pagas.Visible = True
            Lista_recebidas.Visible = False
            ProcCarregaListaContasPagas
        Else
            Lista_pagas.Visible = False
            Lista_recebidas.Visible = True
            Caption = "Financeiro - Instituições - Lista de contas recebidas"
            ProcCarregaListaContasRecebidas
        End If
    End If
    TBFluxo.Close
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaListaContasPagas()
On Error GoTo tratar_erro

Lista_pagas.ListItems.Clear
If Left(TBFluxo!Descricao, 6) = "Malote" Or Left(TBFluxo!Descricao, 6) = "Cheque" Or Left(TBFluxo!Descricao, 3) = "Doc" Or Left(TBFluxo!Descricao, 3) = "Ted" Then
    TextoFiltro = "ID_empresa = " & TBFluxo!ID_empresa & " and NDoctoBaixa = '" & TBFluxo!Cheque & "' and Banco = '" & TBFluxo!Instituicao & "' and Data_movimentacao = '" & TBFluxo!Data & "'"
ElseIf IsNull(TBFluxo!ID_varias) = False And TBFluxo!ID_varias <> "" And TBFluxo!ID_varias <> "0" Then
        TextoFiltro = "ID_empresa = " & TBFluxo!ID_empresa & " and ID_varias = " & TBFluxo!ID_varias
    Else
        TextoFiltro = "IDFluxo = " & TBFluxo!IDFluxo
End If
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from tbl_ContasPagar where " & TextoFiltro & " order by databaixa desc, IdIntConta", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    contador = 0
    Do While TBLISTA.EOF = False
        With Lista_pagas.ListItems
            .Add , , TBLISTA!IDintconta
            .Item(.Count).SubItems(1) = Format(TBLISTA!Dt_emissao, "dd/mm/yy")
            .Item(.Count).SubItems(2) = Format(TBLISTA!dt_Pagamento, "dd/mm/yy")
            
            If TBLISTA!Parcial = True And TBLISTA!status <> "TÍTULO PAGO PARCIAL LIQUIDADO" Then
                valor = IIf(IsNull(TBLISTA!pagoparcial), 0, TBLISTA!pagoparcial) + IIf(IsNull(TBLISTA!ValorPendente), 0, TBLISTA!ValorPendente)
            Else
                valor = IIf(IsNull(TBLISTA!dbl_valorpagto), 0, TBLISTA!dbl_valorpagto)
            End If
            .Item(.Count).SubItems(3) = Format(valor, "###,##0.00")
            
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!txt_ndocumento), "", TBLISTA!txt_ndocumento)
            .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!txt_Parcela), "", TBLISTA!txt_Parcela)
            .Item(.Count).SubItems(6) = Trim(TBLISTA!Txt_fornecedor)
            .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA!ValorPago), "", Format(TBLISTA!ValorPago, "###,##0.00"))
            .Item(.Count).SubItems(8) = IIf(IsNull(TBLISTA!DataBaixa), "", Format(TBLISTA!DataBaixa, "dd/mm/yy"))
            .Item(.Count).SubItems(9) = IIf(IsNull(TBLISTA!resppag), "", TBLISTA!resppag)
        End With
        TBLISTA.MoveNext
        contador = contador + 1
        PBLista.Value = contador
    Loop
Else
    Set TBLISTA = CreateObject("adodb.recordset")
    TBLISTA.Open "Select * from tbl_instituicoes_transf where IDFluxo = " & TBFluxo!IDFluxo & " order by data_transf desc, id_transf", Conexao, adOpenKeyset, adLockOptimistic
    If TBLISTA.EOF = False Then
        PBLista.Min = 0
        PBLista.Max = TBLISTA.RecordCount
        PBLista.Value = 1
        contador = 0
        Do While TBLISTA.EOF = False
            With Lista_pagas.ListItems
                .Add , , TBLISTA!id_transf
                .Item(.Count).SubItems(1) = Format(TBLISTA!data_transf, "dd/mm/yy")
                .Item(.Count).SubItems(2) = Format(TBLISTA!data_transf, "dd/mm/yy")
                .Item(.Count).SubItems(3) = Format(TBLISTA!valor_transf, "###,##0.00")
                .Item(.Count).SubItems(6) = TBLISTA!banco_recebedor
                .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA!valor_transf), "", Format(TBLISTA!valor_transf, "###,##0.00"))
                .Item(.Count).SubItems(8) = IIf(IsNull(TBLISTA!data_transf), "", Format(TBLISTA!data_transf, "dd/mm/yy"))
                .Item(.Count).SubItems(9) = IIf(IsNull(TBLISTA!Responsavel), "", TBLISTA!Responsavel)
            End With
            TBLISTA.MoveNext
            contador = contador + 1
            PBLista.Value = contador
        Loop
    End If
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaListaContasRecebidas()
On Error GoTo tratar_erro

Lista_recebidas.ListItems.Clear
If Left(TBFluxo!Descricao, 6) = "Malote" Or Left(TBFluxo!Descricao, 6) = "Cheque" Or Left(TBFluxo!Descricao, 3) = "Doc" Or Left(TBFluxo!Descricao, 3) = "Ted" Then
    TextoFiltro = "ID_empresa = " & TBFluxo!ID_empresa & " and NDoctoBaixa = '" & TBFluxo!Cheque & "' and Banco = '" & TBFluxo!Instituicao & "' and Data_movimentacao = '" & TBFluxo!Data & "'"
ElseIf IsNull(TBFluxo!ID_varias) = False And TBFluxo!ID_varias <> "" And TBFluxo!ID_varias <> "0" Then
        TextoFiltro = "ID_empresa = " & TBFluxo!ID_empresa & " and ID_varias = " & TBFluxo!ID_varias
    Else
        TextoFiltro = "IDFluxo = " & TBFluxo!IDFluxo
End If
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from tbl_contas_receber where " & TextoFiltro & " order by data_pagamento desc, IdIntConta", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    contador = 0
    Do While TBLISTA.EOF = False
        With Lista_recebidas.ListItems
            .Add.Text = TBLISTA!IDintconta
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!emissao), "", Format(TBLISTA!emissao, "dd/mm/yy"))
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Vencimento), "", Format(TBLISTA!Vencimento, "dd/mm/yy"))
            
            If TBLISTA!Parcial = True And TBLISTA!status <> "TÍTULO RECEBIDO PARCIAL LIQUIDADO" Then
                valor = IIf(IsNull(TBLISTA!RecebidoParcial), 0, TBLISTA!RecebidoParcial) + IIf(IsNull(TBLISTA!ValorPendente), 0, TBLISTA!ValorPendente)
            Else
                valor = IIf(IsNull(TBLISTA!valor), 0, TBLISTA!valor)
            End If
            .Item(.Count).SubItems(3) = Format(valor, "###,##0.00")
            
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!txt_ndocumento), "", TBLISTA!txt_ndocumento)
            .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!NFiscal), "", TBLISTA!NFiscal)
            .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA!Parcela), "", TBLISTA!Parcela)
            .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA!Nome_Razao), "", Trim(TBLISTA!Nome_Razao))
            .Item(.Count).SubItems(8) = IIf(IsNull(TBLISTA!valortitulorecebido), "", Format(TBLISTA!valortitulorecebido, "###,##0.00"))
            .Item(.Count).SubItems(9) = IIf(IsNull(TBLISTA!Data_pagamento), "", Format(TBLISTA!Data_pagamento, "dd/mm/yy"))
            .Item(.Count).SubItems(10) = IIf(IsNull(TBLISTA!resprec), "", TBLISTA!resprec)
        End With
        TBLISTA.MoveNext
        contador = contador + 1
        PBLista.Value = contador
    Loop
Else
    Set TBLISTA = CreateObject("adodb.recordset")
    TBLISTA.Open "Select * from tbl_instituicoes_transf where IDFluxo_rec = " & TBFluxo!IDFluxo & " order by data_transf desc, id_transf", Conexao, adOpenKeyset, adLockOptimistic
    If TBLISTA.EOF = False Then
        PBLista.Min = 0
        PBLista.Max = TBLISTA.RecordCount
        PBLista.Value = 1
        contador = 0
        Do While TBLISTA.EOF = False
            With Lista_recebidas.ListItems
                .Add , , TBLISTA!id_transf
                .Item(.Count).SubItems(1) = Format(TBLISTA!data_transf, "dd/mm/yy")
                .Item(.Count).SubItems(2) = Format(TBLISTA!data_transf, "dd/mm/yy")
                .Item(.Count).SubItems(3) = Format(TBLISTA!valor_transf, "###,##0.00")
                .Item(.Count).SubItems(7) = TBLISTA!banco_remetente
                .Item(.Count).SubItems(8) = IIf(IsNull(TBLISTA!valor_transf), "", Format(TBLISTA!valor_transf, "###,##0.00"))
                .Item(.Count).SubItems(9) = IIf(IsNull(TBLISTA!data_transf), "", Format(TBLISTA!data_transf, "dd/mm/yy"))
                .Item(.Count).SubItems(10) = IIf(IsNull(TBLISTA!Responsavel), "", TBLISTA!Responsavel)
            End With
            TBLISTA.MoveNext
            contador = contador + 1
            PBLista.Value = contador
        Loop
    End If
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    'Case 1: ProcAjuda
    Case 2: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

