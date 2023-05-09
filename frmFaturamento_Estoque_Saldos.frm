VERSION 5.00
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFaturamento_Estoque_Saldos 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Faturamento | Estoque | Movimentação | Saldos"
   ClientHeight    =   6885
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10725
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6885
   ScaleWidth      =   10725
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   2
      Top             =   6480
      Width           =   10725
      _ExtentX        =   18918
      _ExtentY        =   714
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10725
      _ExtentX        =   18918
      _ExtentY        =   714
      DibPicture      =   "frmFaturamento_Estoque_Saldos.frx":0000
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
      Icon            =   "frmFaturamento_Estoque_Saldos.frx":1CAD
   End
   Begin MSComctlLib.ListView Lista_Movimentacao 
      Height          =   4995
      Left            =   180
      TabIndex        =   1
      Top             =   450
      Width           =   10395
      _ExtentX        =   18336
      _ExtentY        =   8811
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   0
      BackColor       =   16777215
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Object.Tag             =   "N"
         Text            =   "Código"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Unidade"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "Descrição"
         Object.Width           =   5821
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Object.Tag             =   "N"
         Text            =   "Saldo em estoque"
         Object.Width           =   2647
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Baixar"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Baixado"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Saldo"
         Object.Width           =   1764
      EndProperty
   End
   Begin DrawSuite2022.USButton BtnMovimentar 
      Height          =   855
      Left            =   180
      TabIndex        =   3
      Top             =   5490
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   1508
      DibPicture      =   "frmFaturamento_Estoque_Saldos.frx":1FC7
      Caption         =   "Movimentar estoque"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   1154291
      BorderColorDisabled=   13160660
      BorderColorDown =   16576
      BorderColorOver =   8438015
      GradientColor1  =   1154291
      GradientColor2  =   1154291
      GradientColor3  =   1154291
      GradientColor4  =   1154291
      GradientColorDisabled1=   14215660
      GradientColorDisabled2=   14215660
      GradientColorDisabled3=   14215660
      GradientColorDisabled4=   14215660
      GradientColorOver1=   8438015
      GradientColorOver2=   8438015
      GradientColorOver3=   8438015
      GradientColorOver4=   8438015
      GradientColorDown1=   16576
      GradientColorDown2=   16576
      GradientColorDown3=   16576
      GradientColorDown4=   16576
      PicAlign        =   8
      PicSize         =   3
      PicSizeH        =   32
      PicSizeW        =   32
      ShowFocusRect   =   0   'False
      Theme           =   5
   End
   Begin DrawSuite2022.USButton BtnExcluirMovimentacao 
      Height          =   855
      Left            =   2790
      TabIndex        =   4
      Top             =   5490
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   1508
      DibPicture      =   "frmFaturamento_Estoque_Saldos.frx":3C74
      Caption         =   "Excluir movimentação"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   5263559
      BorderColorDisabled=   13160660
      BorderColorDown =   4013465
      BorderColorOver =   4408288
      GradientColor1  =   5263559
      GradientColor2  =   5263559
      GradientColor3  =   5263559
      GradientColor4  =   5263559
      GradientColorDisabled1=   13160660
      GradientColorDisabled2=   13160660
      GradientColorDisabled3=   13160660
      GradientColorDisabled4=   13160660
      GradientColorOver1=   4408288
      GradientColorOver2=   4408288
      GradientColorOver3=   4408288
      GradientColorOver4=   4408288
      GradientColorDown1=   4013465
      GradientColorDown2=   4013465
      GradientColorDown3=   4013465
      GradientColorDown4=   4013465
      PicAlign        =   8
      PicSize         =   3
      PicSizeH        =   32
      PicSizeW        =   32
      ShowFocusRect   =   0   'False
      Theme           =   4
   End
   Begin DrawSuite2022.USButton BtnConsultaMovimentacao 
      Height          =   855
      Left            =   5400
      TabIndex        =   5
      Top             =   5490
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   1508
      DibPicture      =   "frmFaturamento_Estoque_Saldos.frx":D4C0
      Caption         =   "Visualizar movimentação"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   4960354
      BorderColorDisabled=   13160660
      BorderColorDown =   4210752
      BorderColorOver =   49152
      GradientColor1  =   4960354
      GradientColor2  =   4960354
      GradientColor3  =   4960354
      GradientColor4  =   4960354
      GradientColorDisabled1=   14215660
      GradientColorDisabled2=   14215660
      GradientColorDisabled3=   14215660
      GradientColorDisabled4=   14215660
      GradientColorOver1=   49152
      GradientColorOver2=   49152
      GradientColorOver3=   49152
      GradientColorOver4=   49152
      GradientColorDown1=   32768
      GradientColorDown2=   32768
      GradientColorDown3=   32768
      GradientColorDown4=   32768
      PicAlign        =   8
      PicSize         =   3
      PicSizeH        =   32
      PicSizeW        =   32
      ShowFocusRect   =   0   'False
      Theme           =   3
   End
   Begin DrawSuite2022.USButton BtnConsultaSaldosNF 
      Height          =   855
      Left            =   8010
      TabIndex        =   6
      Top             =   5490
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   1508
      DibPicture      =   "frmFaturamento_Estoque_Saldos.frx":10B10
      Caption         =   "Conferir saldos do itens"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   8388608
      BorderColorDisabled=   13160660
      BorderColorDown =   16576
      BorderColorOver =   16711680
      GradientColor1  =   8388608
      GradientColor2  =   8388608
      GradientColor3  =   12582912
      GradientColor4  =   8388608
      GradientColorDisabled1=   14215660
      GradientColorDisabled2=   14215660
      GradientColorDisabled3=   14215660
      GradientColorDisabled4=   14215660
      GradientColorOver1=   12582912
      GradientColorOver2=   12582912
      GradientColorOver3=   8388608
      GradientColorOver4=   8388608
      GradientColorDown1=   16576
      GradientColorDown2=   16576
      GradientColorDown3=   16576
      GradientColorDown4=   16576
      PicAlign        =   8
      PicSize         =   3
      PicSizeH        =   32
      PicSizeW        =   32
      ShowFocusRect   =   0   'False
      Theme           =   5
   End
End
Attribute VB_Name = "frmFaturamento_Estoque_Saldos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BtnExcluirMovimentacao_Click()
On Error GoTo tratar_erro

With frmFaturamento_Prod_Serv
ID_nota = .txtId
If ID_nota <> 0 Then
ApagarMovimentacaoNFe
USMsgBox "Movimentação excluida com sucesso!", vbInformation, "CAPRIND v5.0"
End If
End With
ProcCarregaLista

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub BtnMovimentar_Click()
On Error GoTo tratar_erro

With frmFaturamento_Prod_Serv
If USMsgBox("Deseja movimentar estoque com essa NFe?", vbYesNo, "CAPRIND v5.0") = vbYes Then
  If .txtId <> "" Then
  ID_empresa = .txtIDEmpresa.Text
  ID_nota = .txtId.Text
    If ID_nota <> 0 And ID_empresa <> 0 Then
    '======================================
    ' Se for nota de saida baixa estoque
    '======================================
      If .opt_Entrada.Value = False Then
       BaixarEstoqueNF
       If Sair = True Then
       USMsgBox "Baixa executada com sucesso!", vbInformation, "CAPRIND v5.0"
       frmFaturamento_Prod_Serv.ProcExibePagina1 (PaginaAtual)
       Else
       USMsgBox "Movimentação no estoque não executada, pois não existe(m) mais saldo(s) no(s) item(ns) da nota!", vbInformation, "CAPRIND v5.0"
       End If
       
      Else
       EntrarEstoqueNF
       USMsgBox "Entrada estoque executada com sucesso!", vbInformation, "CAPRIND v5.0"
      End If
    End If
  End If
End If
End With
ProcCarregaLista

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaLista

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaLista()
On Error GoTo tratar_erro

Lista_Movimentacao.ListItems.Clear

Set TBLISTA = CreateObject("adodb.recordset")
StrSql = "Select TBL.Int_codigo, TBL.int_Cod_Produto,TBL.txt_Descricao, TBL.int_Qtd, TBL.txt_Unid, TBL.Unidade_com, TBL.qtde_estoque, TBL.ID_nota, TBL.N_Referencia, TBL.Int_NotaFiscal, P.ID_PC, P.Codproduto from (tbl_detalhes_nota TBL INNER JOIN ProjProduto P ON TBL.Codproduto = P.Codproduto) INNER JOIN tbl_NaturezaOperacao CFOP ON CFOP.IDCountCfop = TBL.ID_CFOP where TBL.ID_Nota = " & ID_nota & " and P.Estoque = 'True' and (CFOP.Retorno IS NULL or CFOP.Retorno = 'False')"
'Debug.print StrSql

TBLISTA.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
Contador = 1
If TBLISTA.EOF = False Then
Do While TBLISTA.EOF = False
    With Lista_Movimentacao.ListItems
        .Add , , TBLISTA!Int_codigo
        .Item(.Count).SubItems(1) = TBLISTA!int_Cod_Produto
        .Item(.Count).SubItems(2) = TBLISTA!txt_Unid
        .Item(.Count).SubItems(3) = TBLISTA!Txt_descricao
        VerificaSaldoItem (TBLISTA!int_Cod_Produto)
        .Item(.Count).SubItems(4) = Format(Saldo_Atual, "###,##0.00")
        .Item(.Count).SubItems(5) = Format(TBLISTA!int_Qtd, "###,##0.00")
       VerificaSaldoBaixarItem (TBLISTA!Int_codigo)
        .Item(.Count).SubItems(6) = Format(QTBaixadoItemNota, "###,##0.00")
        .Item(.Count).SubItems(7) = Format(TBLISTA!int_Qtd - QTBaixadoItemNota, "###,##0.00")
        
        TBLISTA.MoveNext
    End With
   Contador = Contador + 1
Loop
Else
'TBLISTA.Close
BtnExcluirMovimentacao.Enabled = False
BtnMovimentar.Enabled = False

USMsgBox "Não existem itens que movimentam estoque automaticamente nessa nota.", vbCritical, "CAPRIND v5.0"
'Unload Me
End If

TBLISTA.Close
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub BtnConsultaMovimentacao_Click()
On Error GoTo tratar_erro

  frmFaturamento_Estoque_Movimentacao.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub BtnConsultaSaldosNF_Click()
On Error GoTo tratar_erro

nNF = frmFaturamento_Prod_Serv.txtNFiscal.Text
  frmFaturamento_Estoque.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

