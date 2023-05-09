VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form frmprod_ListaOrdem 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Lista de ordem"
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   1605
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
   ScaleHeight     =   4020
   ScaleWidth      =   1605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Centralziar na Tela
   Begin VB.ComboBox Cmb_prazo_final 
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
      ItemData        =   "frmprod_ListaOrdem.frx":0000
      Left            =   0
      List            =   "frmprod_ListaOrdem.frx":0028
      MouseIcon       =   "frmprod_ListaOrdem.frx":0091
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Prazo final."
      Top             =   270
      Width           =   1590
   End
   Begin VB.CommandButton Cmd_carregar_lista 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Carregar lista (F2)"
      Height          =   345
      Left            =   0
      MouseIcon       =   "frmprod_ListaOrdem.frx":039B
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   660
      Width           =   1590
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   2685
      Left            =   0
      TabIndex        =   2
      Top             =   1050
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   4736
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
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
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Ordem"
         Object.Width           =   2117
      EndProperty
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   3730
      Width           =   1590
      _ExtentX        =   2805
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparente
      Caption         =   "Prazo final"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   420
      TabIndex        =   3
      Top             =   60
      Width           =   750
   End
End
Attribute VB_Name = "frmprod_ListaOrdem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmd_carregar_lista_Click()
On Error GoTo tratar_erro

Select Case Cmb_prazo_final
    Case "Janeiro": M = 1
    Case "Fevereiro": M = 2
    Case "Março": M = 3
    Case "Abril": M = 4
    Case "Maio": M = 5
    Case "Junho": M = 6
    Case "Julho": M = 7
    Case "Agosto": M = 8
    Case "Setembro": M = 9
    Case "Outubro": M = 10
    Case "Novembro": M = 11
    Case "Dezembro": M = 12
End Select

OF = 0
Lista.ListItems.Clear
Set TBOrdem = CreateObject("adodb.recordset")
TBOrdem.Open "Select Ordem from Ordemservico Where Month(Prazofinal) = '" & M & "' group by Ordem", Conexao, adOpenKeyset, adLockReadOnly
If TBOrdem.EOF = False Then
    TBOrdem.MoveLast
    PBLista.Min = 0
    PBLista.Max = TBOrdem.RecordCount
    PBLista.Value = 1
    Contador = 0
    TBOrdem.MoveFirst
    Do While TBOrdem.EOF = False
        Set TBproducao = CreateObject("adodb.recordset")
        TBproducao.Open "Select * from Producao where Ordem = " & TBOrdem!Ordem, Conexao, adOpenKeyset, adLockOptimistic
        If TBproducao.EOF = True Then
            With Lista.ListItems
                .Add , , TBOrdem!Ordem
            End With
        End If
        TBproducao.Close
        TBOrdem.MoveNext
        Contador = Contador + 1
        PBLista.Value = Contador
    Loop
End If
TBOrdem.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyEscape: Unload Me
    Case vbKeyF2: Cmd_carregar_lista_Click
End Select
            
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

Cmb_prazo_final = "Janeiro"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
