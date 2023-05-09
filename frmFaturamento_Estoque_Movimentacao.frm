VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form frmFaturamento_Estoque_Movimentacao 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'Nenhum
   Caption         =   "Faturamento | Estoque - Movimentação"
   ClientHeight    =   5880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5760
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
   ScaleHeight     =   5880
   ScaleWidth      =   5760
   StartUpPosition =   2  'Centralziar na Tela
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   2
      Top             =   5475
      Width           =   5760
      _ExtentX        =   10160
      _ExtentY        =   714
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5760
      _ExtentX        =   10160
      _ExtentY        =   714
      DibPicture      =   "frmFaturamento_Estoque_Movimentacao.frx":0000
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
      Icon            =   "frmFaturamento_Estoque_Movimentacao.frx":62E4
      ShowMaximize    =   0   'False
      ShowMinimize    =   0   'False
   End
   Begin MSComctlLib.ListView Lista_Movimentacao 
      Height          =   4185
      Left            =   330
      TabIndex        =   1
      Top             =   750
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   7382
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
      NumItems        =   5
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
         Object.Tag             =   "T"
         Text            =   "Data"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Object.Tag             =   "N"
         Text            =   "RE"
         Object.Width           =   1412
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Movimentou"
         Object.Width           =   2117
      EndProperty
   End
End
Attribute VB_Name = "frmFaturamento_Estoque_Movimentacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
On Error GoTo tratar_erro

Lista_Movimentacao.ListItems.Clear

Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Estoque_Movimentacao where ID_NOTA = '" & ID_nota & "' and operacao ='SAIDA_NOTA'", Conexao, adOpenKeyset, adLockOptimistic
contador = 1
Do While TBLISTA.EOF = False
    With Lista_Movimentacao.ListItems
        .Add , , contador
        .Item(.Count).SubItems(1) = TBLISTA!Desenho
        .Item(.Count).SubItems(2) = TBLISTA!Data
        .Item(.Count).SubItems(3) = TBLISTA!IDEstoque
        .Item(.Count).SubItems(4) = Format(TBLISTA!Saida, "###,##0.00")
        TBLISTA.MoveNext
    End With
   contador = contador + 1
Loop

TBLISTA.Close
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
