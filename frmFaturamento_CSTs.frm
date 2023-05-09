VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form frmFaturamento_CST 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Faturamento | CSTs"
   ClientHeight    =   3135
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4710
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4710
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Lista de CST Icms, Ipi, Pis, Cofins"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2205
      Left            =   150
      TabIndex        =   1
      Top             =   660
      Width           =   4335
      Begin MSComctlLib.ListView ListaCST 
         Height          =   1785
         Left            =   120
         TabIndex        =   2
         Top             =   270
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   3149
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "N"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "ICMS"
            Object.Width           =   1412
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "IPI"
            Object.Width           =   1412
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Object.Tag             =   "N"
            Text            =   "PIS"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "Cofins"
            Object.Width           =   1411
         EndProperty
      End
   End
   Begin DrawSuite2022.USForm frmFaturamento_CST 
      Align           =   1  'Align Top
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4710
      _ExtentX        =   8308
      _ExtentY        =   714
      DibPicture      =   "frmFaturamento_CSTs.frx":0000
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
      Icon            =   "frmFaturamento_CSTs.frx":A123
      ShowMaximize    =   0   'False
      ShowMinimize    =   0   'False
   End
End
Attribute VB_Name = "frmFaturamento_CST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCadastrarCST_Click()

frm_Natureza_OP_CST.Show 1

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro
        
 Select Case KeyCode
     Case vbKeyF2:
     Case vbKeyF3: ProcGravarCST
     Case vbKeyF1: 'ProcAjuda
     Case vbKeyEscape: Unload Me
 End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcGravarCST()
On Error GoTo tratar_erro

AlterarProduto = True
Unload Me
frmFaturamento_Prod_Serv.ProcSalvarProduto

'Conexao.Execute "Update tbl_Detalhes_Nota set txt_Cst = '" & vICMSCST & "', cst_ipi = '" & IPICST & "', CST_PIS = '" & PISCST & "', cst_COFINS = '" & COFINSCST & "' WHERE INT_CODIGO = " & frmFaturamento_Prod_Serv.txtidproduto



USMsgBox "Dados CST gravados com sucesso!"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ListaCST.ListItems.Clear

Do While TBAbrir.EOF = False
    With ListaCST.ListItems
        .Add , , IIf(IsNull(TBAbrir!ID), "", TBAbrir!ID)
        .Item(.Count).SubItems(1) = IIf(IsNull(TBAbrir!CST_ICMS), "", TBAbrir!CST_ICMS)
        .Item(.Count).SubItems(2) = IIf(IsNull(TBAbrir!CST_IPI), "", TBAbrir!CST_IPI)
        .Item(.Count).SubItems(3) = IIf(IsNull(TBAbrir!CST_PIS), "", TBAbrir!CST_PIS)
        .Item(.Count).SubItems(4) = IIf(IsNull(TBAbrir!CST_Cofins), "", TBAbrir!CST_Cofins)
    End With
    TBAbrir.MoveNext
Loop

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaCST_DblClick()
On Error GoTo tratar_erro

With frmFaturamento_Prod_Serv

If ListaCST.ListItems.Count <> 0 Then
.txtCST_ICMS.Text = ListaCST.SelectedItem.ListSubItems.Item(1).Text
.txtCST_IPI.Text = ListaCST.SelectedItem.ListSubItems.Item(2).Text
.txtCST_PIS.Text = ListaCST.SelectedItem.ListSubItems.Item(3).Text
.txtCST_Cofins.Text = ListaCST.SelectedItem.ListSubItems.Item(4).Text
End If

vICMSCST = ListaCST.SelectedItem.ListSubItems.Item(1).Text
IPICST = ListaCST.SelectedItem.ListSubItems.Item(2).Text
PISCST = ListaCST.SelectedItem.ListSubItems.Item(3).Text
COFINSCST = ListaCST.SelectedItem.ListSubItems.Item(4).Text

End With
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaCST_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With frmFaturamento_Prod_Serv

If ListaCST.ListItems.Count <> 0 Then
.txtCST_ICMS.Text = ListaCST.SelectedItem.ListSubItems.Item(1).Text
.txtCST_IPI.Text = ListaCST.SelectedItem.ListSubItems.Item(2).Text
.txtCST_PIS.Text = ListaCST.SelectedItem.ListSubItems.Item(3).Text
.txtCST_Cofins.Text = ListaCST.SelectedItem.ListSubItems.Item(4).Text

vICMSCST = ListaCST.SelectedItem.ListSubItems.Item(1).Text
IPICST = ListaCST.SelectedItem.ListSubItems.Item(2).Text
PISCST = ListaCST.SelectedItem.ListSubItems.Item(3).Text
COFINSCST = ListaCST.SelectedItem.ListSubItems.Item(4).Text

End If

End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
