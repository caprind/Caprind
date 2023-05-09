VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form frmPlanomedicao_ListaRNC 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Controle de medição - Lista RNC"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3015
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
   ScaleHeight     =   4080
   ScaleWidth      =   3015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Centralziar na Tela
   Begin MSComctlLib.ListView Lista 
      Height          =   3750
      Left            =   0
      TabIndex        =   0
      Top             =   30
      Width           =   2985
      _ExtentX        =   5265
      _ExtentY        =   6615
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Dt. emissão"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "RNC"
         Object.Width           =   2496
      EndProperty
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   3780
      Width           =   2985
      _ExtentX        =   5265
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
Attribute VB_Name = "frmPlanomedicao_ListaRNC"
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

Permitido = True
With frmPlanomedicao
    Lista.ListItems.Clear
    Set TBLISTA = CreateObject("adodb.recordset")
    TBLISTA.Open "Select ID, Data, id_texto, Fim from CQ_RNC where Desenho = '" & .txtDesenho & "' order by ID desc", Conexao, adOpenKeyset, adLockOptimistic
    If TBLISTA.EOF = False Then
        TBLISTA.MoveLast
        PBLista.Min = 0
        PBLista.Max = TBLISTA.RecordCount
        PBLista.Value = 1
        Contador = 0
        TBLISTA.MoveFirst
        Do While TBLISTA.EOF = False
            If IsNull(TBLISTA!fim) = True Or TBLISTA!fim = "" Then
            With Lista.ListItems
                .Add , , TBLISTA!ID
                .Item(.Count).SubItems(1) = Format(TBLISTA!data, "dd/mm/yy/")
                .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!id_texto), "", TBLISTA!id_texto)
            End With
            End If
            TBLISTA.MoveNext
            Contador = Contador + 1
            PBLista.Value = Contador
        Loop
    Else
        Permitido = False
    End If
    TBLISTA.Close
End With

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
RNC_Inspecao_Recebimento = False
RNC_Controle_Medicao = False
RNC_Nao_Conformidade = False
RNC_Solicitacao_Desvio = False
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from CQ_RNC where ID = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    With frmQualidade_RNC
        Unload Me
        .ProcLimpaCampos
        .ProcPuxaDados
        .Show
    End With
End If
TBAbrir.Close
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
