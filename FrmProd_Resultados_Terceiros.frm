VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form FrmProd_Resultados_Terceiros 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "PCP - Gerenciamento de ordem - Resultados da ordem detalhado - Terceiros"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11550
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmProd_Resultados_Terceiros.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   11550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Centralziar na Tela
   Begin MSComctlLib.ListView Lista 
      Height          =   5410
      Left            =   50
      TabIndex        =   0
      Top             =   20
      Width           =   11480
      _ExtentX        =   20241
      _ExtentY        =   9551
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      MousePointer    =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "FrmProd_Resultados_Terceiros.frx":000C
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Cód. interno"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Descrição"
         Object.Width           =   5186
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Un."
         Object.Width           =   970
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Qtde."
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Valor unit."
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Valor total"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Pedido"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Nota fiscal"
         Object.Width           =   2469
      EndProperty
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   50
      TabIndex        =   1
      Top             =   5430
      Width           =   11475
      _ExtentX        =   20241
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
End
Attribute VB_Name = "FrmProd_Resultados_Terceiros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Lista.ListItems.Clear
If OF = 0 Then Exit Sub
CamposFiltro = "NFP.Int_codigo, CPL.IDlista, CPL.Desenho, CPL.Descricao, CPL.Un, CP.Pedido, NFPP.Quantidade as Qtde, NFP.dbl_ValorUnitario as Valor, ROUND(NFP.dbl_ValorUnitario * NFPP.Quantidade, 2) as Valor1, NFP.Int_NotaFiscal as NotaFiscal"
Set TBMaterial = CreateObject("adodb.recordset")
TBMaterial.Open "Select " & CamposFiltro & " from ((Compras_pedido_lista CPL inner join Compras_pedido CP on CPL.IDpedido = CP.IDpedido) INNER JOIN tbl_Detalhes_Nota_pedidos NFPP ON NFPP.ID_carteira = CPL.IDlista and NFPP.Codinterno = CPL.Desenho) INNER JOIN tbl_Detalhes_Nota NFP ON NFP.Int_codigo = NFPP.ID_prod_NF where CPL.Ordem = " & OF & " and CPL.OS is not null and CPL.OS <> 0 and CPL.Remessa = 'False' and (CPL.Status_Item = 'RECEBIDO' or CPL.Status_Item = 'PARCIAL') order by CPL.Desenho", Conexao, adOpenKeyset, adLockOptimistic
If TBMaterial.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBMaterial.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBMaterial.EOF = False
        With Lista.ListItems
            Teste = TBMaterial!Int_codigo
            .Add , , TBMaterial!IDlista
            .Item(.Count).SubItems(1) = IIf(IsNull(TBMaterial!Desenho), "", TBMaterial!Desenho)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBMaterial!Descricao), "", TBMaterial!Descricao)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBMaterial!Un), "", TBMaterial!Un)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBMaterial!Qtde), "", Format(TBMaterial!Qtde, "###,##0.0000"))
            .Item(.Count).SubItems(5) = IIf(IsNull(TBMaterial!valor), "", Format(TBMaterial!valor, "###,##0.0000000000"))
            .Item(.Count).SubItems(6) = IIf(IsNull(TBMaterial!Valor1), "", Format(TBMaterial!Valor1, "###,##0.00"))
            .Item(.Count).SubItems(7) = IIf(IsNull(TBMaterial!Pedido), "", TBMaterial!Pedido)
            .Item(.Count).SubItems(8) = IIf(IsNull(TBMaterial!NotaFiscal), "", TBMaterial!NotaFiscal)
        End With
        TBMaterial.MoveNext
        Contador = Contador + 1
        PBLista.Value = Contador
    Loop
    
End If
TBMaterial.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView Lista, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
