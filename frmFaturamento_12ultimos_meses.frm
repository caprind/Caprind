VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form frmFaturamento_12ultimos_meses 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'Nenhum
   Caption         =   "Faturamento - Relatórios - Doze últimos meses"
   ClientHeight    =   4995
   ClientLeft      =   0
   ClientTop       =   45
   ClientWidth     =   9480
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFaturamento_12ultimos_meses.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   9480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Centralziar na Tela
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   7
      Top             =   4590
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   714
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   465
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   820
      DibPicture      =   "frmFaturamento_12ultimos_meses.frx":000C
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
      Icon            =   "frmFaturamento_12ultimos_meses.frx":A12F
      ShowMaximize    =   0   'False
      ShowMinimize    =   0   'False
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   270
      TabIndex        =   2
      Top             =   540
      Width           =   8940
      Begin VB.ComboBox Cmb_empresa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "frmFaturamento_12ultimos_meses.frx":A14B
         Left            =   990
         List            =   "frmFaturamento_12ultimos_meses.frx":A14D
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         ToolTipText     =   "Empresa."
         Top             =   180
         Width           =   3975
      End
      Begin VB.TextBox txttotal 
         Alignment       =   1  'Alinhar à Direita
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'Nenhum
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6960
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Valor total."
         Top             =   180
         Width           =   1860
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparente
         Caption         =   "Empresa :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   210
         TabIndex        =   6
         Top             =   180
         Width           =   720
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparente
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   6360
         TabIndex        =   3
         Top             =   180
         Width           =   570
         WordWrap        =   -1  'True
      End
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   3105
      Left            =   270
      TabIndex        =   0
      Top             =   1200
      Width           =   8940
      _ExtentX        =   15769
      _ExtentY        =   5477
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
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
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Text            =   "Mês"
         Object.Width           =   970
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Object.Tag             =   "N"
         Text            =   "Ano"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Object.Tag             =   "N"
         Text            =   "Valor"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Object.Tag             =   "N"
         Text            =   "Frete"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Object.Tag             =   "N"
         Text            =   "Seguro"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Object.Tag             =   "N"
         Text            =   "Acessórias"
         Object.Width           =   1852
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Object.Tag             =   "N"
         Text            =   "Desconto"
         Object.Width           =   1941
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Object.Tag             =   "N"
         Text            =   "Devolução"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Object.Tag             =   "N"
         Text            =   "Total fat."
         Object.Width           =   2117
      EndProperty
   End
End
Attribute VB_Name = "frmFaturamento_12ultimos_meses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ProcSair()
On Error GoTo tratar_erro

Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_empresa_Click()
On Error GoTo tratar_erro

ProcCarregaLista

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaLista()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Valor1 = 0
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "select * from Faturamento_12ultimos_meses_mensal where Codigo = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " ORDER BY codigo, Ano, Mes", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    contador = 0
    Do While TBAbrir.EOF = False
        With Lista.ListItems
            .Add , , FunVerificaNumeroMes(TBAbrir!Mes)
            .Item(.Count).SubItems(1) = TBAbrir!Ano
            .Item(.Count).SubItems(2) = FunFormataCasasDecimais(2, TBAbrir!VlrTotal)
            .Item(.Count).SubItems(3) = FunFormataCasasDecimais(2, TBAbrir!VlrTotalFrete)
            .Item(.Count).SubItems(4) = FunFormataCasasDecimais(2, TBAbrir!VlrTotalSeguro)
            .Item(.Count).SubItems(5) = FunFormataCasasDecimais(2, TBAbrir!VlrTotalAcessorias)
            .Item(.Count).SubItems(6) = FunFormataCasasDecimais(2, TBAbrir!VlrTotalDesc + TBAbrir!VlrTotalDescSUFRAMA)
            
            valor = 0
            Set TBItem = CreateObject("adodb.recordset")
            TBItem.Open "Select VlrTotalFat from Faturamento_12ultimos_meses_devolucao_mensal where Codigo = " & TBAbrir!CODIGO & " and Mes = " & TBAbrir!Mes & " and Ano = " & TBAbrir!Ano, Conexao, adOpenKeyset, adLockOptimistic
            If TBItem.EOF = False Then
                valor = TBItem!VlrTotalFat
            End If
            TBItem.Close
            .Item(.Count).SubItems(7) = FunFormataCasasDecimais(2, valor)
            .Item(.Count).SubItems(8) = FunFormataCasasDecimais(2, TBAbrir!VlrTotalFat - valor)
            
            Valor1 = Valor1 + (TBAbrir!VlrTotalFat - valor)
        End With
        TBAbrir.MoveNext
        contador = contador + 1
    Loop
    txttotal = FunFormataCasasDecimais(2, Valor1)
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

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

ProcCarregaComboEmpresa Cmb_empresa, False
ProcCarregaLista

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

