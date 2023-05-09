VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form frmVendas_PI_Empenhos 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'Nenhum
   Caption         =   "CAPRIND v5.0 | Lista de RE(s) empenhadas"
   ClientHeight    =   3810
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5700
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
   ScaleHeight     =   3810
   ScaleWidth      =   5700
   StartUpPosition =   2  'Centralziar na Tela
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   435
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5700
      _ExtentX        =   10054
      _ExtentY        =   767
      DibPicture      =   "frmVendas_PI_Empenhos.frx":0000
      CaptionDelimiter=   "|"
      CaptionOnCenter =   -1  'True
      EnableMaximizeButton=   0   'False
      EnableMinimizeButton=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Icon            =   "frmVendas_PI_Empenhos.frx":1854
      ShowMaximize    =   0   'False
      ShowMinimize    =   0   'False
   End
   Begin MSComctlLib.ListView ListaEmpenhoEstoquePed 
      Height          =   2655
      Left            =   330
      TabIndex        =   1
      Top             =   720
      Width           =   4875
      _ExtentX        =   8599
      _ExtentY        =   4683
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
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Object.Tag             =   "N"
         Text            =   "RE"
         Object.Width           =   883
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Object.Tag             =   "N"
         Text            =   "Vendido"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Object.Tag             =   "N"
         Text            =   "Empenhado"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Object.Tag             =   "N"
         Text            =   "Baixado"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Object.Tag             =   "N"
         Text            =   "Saldo"
         Object.Width           =   1587
      EndProperty
   End
End
Attribute VB_Name = "frmVendas_PI_Empenhos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ProcCarregaListaEmpEstPed()
On Error GoTo tratar_erro

Valor3 = 0
ListaEmpenhoEstoquePed.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
'TBLISTA.Open "Select ECEV.ID, ECEV.ID_carteira, ECEV.Data1 as Dataemp, ECEV.Responsavel as Respemp, ECEV.Qtde_empenhada, ECEV.Qtde_saida, VP.*, VC.Codigo, VC.Qtde_produzir - VC.Qtdefaturada as Qtd, VC.Prazofinal from (Estoque_Controle_Empenho_Vendas ECEV INNER JOIN vendas_carteira VC ON ECEV.ID_Carteira =  VC.Codigo) INNER JOIN vendas_proposta VP ON VP.cotacao = VC.Cotacao where ECEV.ID_Estoque = " & ListaEstoque.SelectedItem & " and ECEV.Qtde_empenhada - ECEV.Qtde_saida > 0", Conexao, adOpenKeyset, adLockOptimistic
TBLISTA.Open "Select ECEV.ID_Estoque, ECEV.ID, ECEV.ID_carteira, ECEV.Data as Dataemp, ECEV.Responsavel as Respemp, ECEV.Qtde_empenhada, ECEV.Qtde_saida, VP.*, VC.Codigo, VC.Qtde_produzir as Qtd, VC.Prazofinal from (Estoque_Controle_Empenho_Vendas ECEV INNER JOIN vendas_carteira VC ON ECEV.ID_Carteira =  VC.Codigo) INNER JOIN vendas_proposta VP ON VP.cotacao = VC.Cotacao where ECEV.ID_CARTEIRA = " & frmVendas_PI.Listprod.SelectedItem & " and ECEV.Qtde_empenhada - ECEV.Qtde_saida > 0", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    TBLISTA.MoveLast
    'PBLista3.Min = 0
    'PBLista3.Max = TBLISTA.RecordCount
    'PBLista3.Value = 1
    contador = 0
    TBLISTA.MoveFirst
    Do While TBLISTA.EOF = False
        With ListaEmpenhoEstoquePed.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = TBLISTA!ID_estoque
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Qtd), "", TBLISTA!Qtd)
            valor = IIf(IsNull(TBLISTA!Qtde_empenhada), 0, TBLISTA!Qtde_empenhada)
            .Item(.Count).SubItems(3) = valor
            Valor1 = IIf(IsNull(TBLISTA!Qtde_saida), 0, TBLISTA!Qtde_saida)
            .Item(.Count).SubItems(4) = Valor1
            .Item(.Count).SubItems(5) = valor - Valor1
        End With
        Valor3 = Valor3 + (valor - Valor1)
        TBLISTA.MoveNext
        contador = contador + 1
        'PBLista3.Value = contador
    Loop
End If
TBLISTA.Close
'valor = ListaEstoque.SelectedItem.ListSubItems(6)
'Txt_qtde_total_estoque = valor
QuantEmpenho = 0
'Set TBAbrir = CreateObject("adodb.recordset")
'TBAbrir.Open "Select SUM(Quantidade - Qtde_saida) as QuantEmpenho from Producao_NF_Consignada where IDEstoque = " & ListaEstoque.SelectedItem & " and Quantidade - Qtde_saida > 0", Conexao, adOpenKeyset, adLockOptimistic
'If TBAbrir.EOF = False Then
'    QuantEmpenho = IIf(IsNull(TBAbrir!QuantEmpenho), 0, TBAbrir!QuantEmpenho)
'End If
'Txt_qtde_total_emp_estoque = Valor3 + QuantEmpenho
'Txt_qtde_total_disp_estoque = valor - (Valor3 + QuantEmpenho)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaListaEmpEstPed
Id_Item = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
