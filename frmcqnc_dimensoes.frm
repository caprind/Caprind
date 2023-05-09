VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form frmcqnc_dimensoes 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Qualidade - Não conformidade - Dimensões"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7095
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
   ScaleHeight     =   5430
   ScaleWidth      =   7095
   StartUpPosition =   2  'Centralziar na Tela
   Begin MSComctlLib.ListView Lista 
      Height          =   5160
      Left            =   55
      TabIndex        =   0
      Top             =   0
      Width           =   6990
      _ExtentX        =   12330
      _ExtentY        =   9102
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
      BorderStyle     =   1
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
         Text            =   "Id"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Índice"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Tipo da dimensão"
         Object.Width           =   3087
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Dimensão"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Tol. sup."
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Tol. inf."
         Object.Width           =   2293
      EndProperty
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   55
      TabIndex        =   1
      Top             =   5160
      Width           =   6990
      _ExtentX        =   12330
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
Attribute VB_Name = "frmcqnc_dimensoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro
  
Select Case KeyCode
    Case vbKeyEscape: Unload Me
    Case vbKeyReturn: Lista_DblClick
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

If Sit_REG = 0 Then
    With frmcqnc
        TextoFiltro = "PR.Ordem = " & .txtOF & " and P.Fase = " & .txtFase
    End With
Else
    With frmcqnc.ListaFases
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                TextoFiltro = "PR.Ordem = " & .ListItems.Item(InitFor).ListSubItems(2) & " and P.Fase = " & .ListItems.Item(InitFor).ListSubItems(5)
                GoTo Prosseguir
            End If
        Next InitFor
    End With
End If

Prosseguir:
    Lista.ListItems.Clear
    Set TBLISTA = CreateObject("adodb.recordset")
    TBLISTA.Open "Select PD.* from Plano P INNER JOIN planodimensao PD on P.idplano = PD.idplano INNER JOIN Producao PR ON PR.Desenho = P.Desenho where " & TextoFiltro & " and P.DtValidacao IS NOT NULL order by PD.indice", Conexao, adOpenKeyset, adLockOptimistic
    If TBLISTA.EOF = False Then
        PBLista.Min = 0
        PBLista.Max = TBLISTA.RecordCount
        PBLista.Value = 1
        Contador = 0
        Do While TBLISTA.EOF = False
            With Lista.ListItems
                .Add , , TBLISTA!idDimensao
                .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!indice), "", TBLISTA!indice)
                .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Tipo), "", TBLISTA!Tipo)
                .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!dimdesejada), "", Format(TBLISTA!dimdesejada, "###,##0.0000"))
                .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!TolSup), "", Format(TBLISTA!TolSup, "###,##0.0000"))
                .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!TolInf), "", Format(TBLISTA!TolInf, "###,##0.0000"))
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

Private Sub Lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView Lista, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_DblClick()
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 Then Exit Sub
ParecerTexto = Lista.SelectedItem.ListSubItems(1) & " - " & Lista.SelectedItem.ListSubItems(2) & " - " & Lista.SelectedItem.ListSubItems(3) & " - Tol. sup.: " & Lista.SelectedItem.ListSubItems(4) & " - Tol. inf.: " & Lista.SelectedItem.ListSubItems(5)
If Sit_REG = 0 Then frmcqnc.txtParecerF = ParecerTexto Else frmcqnc_disposicao.txtParecerF = ParecerTexto
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
