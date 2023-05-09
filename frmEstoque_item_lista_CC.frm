VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form frmEstoque_item_lista_CC 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Estoque - Movimentação - Centro de custo"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8655
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   8655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ListView Lista 
      Height          =   4115
      Left            =   55
      TabIndex        =   0
      Top             =   30
      Width           =   8565
      _ExtentX        =   15108
      _ExtentY        =   7250
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
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "T"
         Text            =   "Código"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Descrição"
         Object.Width           =   9657
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Object.Tag             =   "N"
         Text            =   "Valor"
         Object.Width           =   2117
      EndProperty
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   55
      TabIndex        =   1
      Top             =   4155
      Width           =   8565
      _ExtentX        =   15108
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
Attribute VB_Name = "frmEstoque_item_lista_CC"
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

ProcLimpaVariaveisPrincipais
ProcCarregaLista

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaLista()
On Error GoTo tratar_erro

Permitido = False
Lista.ListItems.Clear
If Estoque_recebimento = True Then
    'Verifica se é recebimento com pedido ou programação
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select Estoque_controle_recebimento.*, Estoque_movimentacao.IDLista_recebimento from Estoque_movimentacao INNER JOIN Estoque_controle_recebimento ON Estoque_movimentacao.IDEstoque_recebimento = Estoque_controle_recebimento.Id where Estoque_movimentacao.IDoperacao = " & frmestoque_item.Lista_movimentacao.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Permitido = True
        If TBAbrir!Programacao = True Then
            Programacao = True
            TextoFiltro = ""
        Else
            Programacao = False
            TextoFiltro = "CC_realizado.ID_Lista = " & TBAbrir!idlista_recebimento
        End If
    End If
    TBAbrir.Close
Else
    Permitido = True
    Programacao = False
    TextoFiltro = "CC_realizado.ID_estoque = " & frmestoque_item.Lista_movimentacao.SelectedItem
End If

If Permitido = True Then
    If Programacao = True Then
        Exit Sub
    Else
        Set TBLISTA = CreateObject("adodb.recordset")
        TBLISTA.Open "Select CC_realizado.*, Usuarios_setor.Codigo, Usuarios_setor.Setor from CC_realizado INNER JOIN Usuarios_setor ON CC_realizado.ID_CC = Usuarios_setor.ID where " & TextoFiltro & " and (CC_realizado.ID_origem is null or CC_realizado.ID_origem = 0) order by Usuarios_setor.Codigo", Conexao, adOpenKeyset, adLockOptimistic
    End If
    If TBLISTA.EOF = False Then
        TBLISTA.MoveLast
        PBLista.Min = 0
        PBLista.Max = TBLISTA.RecordCount
        PBLista.Value = 1
        contador = 0
        TBLISTA.MoveFirst
        Do While TBLISTA.EOF = False
            With Lista.ListItems
                .Add , , TBLISTA!ID
                .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!CODIGO), "", TBLISTA!CODIGO)
                .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Setor), "", TBLISTA!Setor)
                
                Select Case TBLISTA!Operacao
                    Case "Crédito": ValorTexto = "+" & IIf(IsNull(TBLISTA!valor), "", Format(TBLISTA!valor, "###,##0.00"))
                    Case "Débito": ValorTexto = "-" & IIf(IsNull(TBLISTA!valor), "", Format(TBLISTA!valor, "###,##0.00"))
                End Select
                .Item(.Count).SubItems(3) = ValorTexto
            End With
            TBLISTA.MoveNext
            contador = contador + 1
            PBLista.Value = contador
        Loop
    End If
    TBLISTA.Close
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
