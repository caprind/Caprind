VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form FrmProd_Resultados_Apontamentos 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'Nenhum
   Caption         =   "PCP | Gerenciamento de ordem - Resultados da ordem detalhado - Lista de apontamentos: PREPARANDO MÁQUINA"
   ClientHeight    =   7020
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   12090
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmProd_Resultados_Apontamentos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   12090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Centralziar na Tela
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   3
      Top             =   6615
      Width           =   12090
      _ExtentX        =   21325
      _ExtentY        =   714
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   435
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   12090
      _ExtentX        =   21325
      _ExtentY        =   767
      DibPicture      =   "FrmProd_Resultados_Apontamentos.frx":000C
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
      Icon            =   "FrmProd_Resultados_Apontamentos.frx":718C
      ShowMaximize    =   0   'False
      ShowMinimize    =   0   'False
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   5415
      Left            =   270
      TabIndex        =   0
      Top             =   645
      Width           =   11475
      _ExtentX        =   20241
      _ExtentY        =   9551
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
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
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "IDProducao"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "OS"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Início"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Final"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Tempo total"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Operador"
         Object.Width           =   4507
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Qtde. aprov."
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "Qtde. N/C"
         Object.Width           =   1676
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   8
         Text            =   "Pronta"
         Object.Width           =   1376
      EndProperty
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   270
      TabIndex        =   1
      Top             =   6060
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
Attribute VB_Name = "FrmProd_Resultados_Apontamentos"
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
If Sit_REG = 2 Then
    Caption = "PCP - Gerenciamento de ordem - Resultados da ordem detalhado - Lista de apontamentos: MÁQUINA EM PRODUÇÃO"
    TextoFiltro = "2"
Else
    TextoFiltro = "1"
End If
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from " & NomeTabelaAp & " where Ordem = " & OF & " and CodigoDesc = " & TextoFiltro & " order by Data, Tempoinicio, IDFase", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    contador = 0
    Do While TBLISTA.EOF = False
        With Lista.ListItems
            .Add , , TBLISTA!IDProducao
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!IDFase), "", TBLISTA!IDFase)
            .Item(.Count).SubItems(2) = Format(TBLISTA!TempoInicio, "dd/mm/yy hh:mm:ss")
            .Item(.Count).SubItems(3) = Format(TBLISTA!TempoFinal, "dd/mm/yy hh:mm:ss")
            If TBLISTA!Dias <> 0 Then
                TempoTotalDias = IIf(IsNull(TBLISTA!TempoTotal), 0, TBLISTA!TempoTotal) + TBLISTA!Dias
                ElapsedTime (TempoTotalDias)
                .Item(.Count).SubItems(3) = HoraTotal
            Else
                .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!TempoTotal), "", TBLISTA!TempoTotal)
            End If
            .Item(.Count).SubItems(5) = TBLISTA!Usuario
            .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA!quantidade), 0, TBLISTA!quantidade)
            .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA!Reprovada), 0, TBLISTA!Reprovada)
            .Item(.Count).SubItems(8) = TBLISTA!Pronto
        End With
        TBLISTA.MoveNext
        contador = contador + 1
        PBLista.Value = contador
    Loop
    
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
