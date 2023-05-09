VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form frmVendas_Lista_Vendedores 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'Nenhum
   Caption         =   "CAPRIND v5.0 | Vendas | Proposta comercial | Vendedores"
   ClientHeight    =   5115
   ClientLeft      =   1635
   ClientTop       =   975
   ClientWidth     =   7095
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Centralziar na Tela
   Begin DrawSuite2022.USButton BtnCarregar 
      Height          =   345
      Left            =   5430
      TabIndex        =   4
      Top             =   4650
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   609
      BorderColor     =   5263559
      BorderColorDisabled=   13160660
      BorderColorDown =   4013465
      BorderColorOver =   4408288
      Caption         =   "Carregar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
      ForeColorDown   =   16777215
      ForeColorOver   =   16777215
      GradientColor1  =   5263559
      GradientColor2  =   5263559
      GradientColor3  =   5263559
      GradientColor4  =   5263559
      GradientColorDisabled1=   13160660
      GradientColorDisabled2=   13160660
      GradientColorDisabled3=   13160660
      GradientColorDisabled4=   13160660
      GradientColorDown1=   4013465
      GradientColorDown2=   4013465
      GradientColorDown3=   4013465
      GradientColorDown4=   4013465
      GradientColorOver1=   4408288
      GradientColorOver2=   4408288
      GradientColorOver3=   4408288
      GradientColorOver4=   4408288
      Theme           =   4
   End
   Begin DrawSuite2022.USLabel USLabel1 
      Height          =   720
      Left            =   1140
      Top             =   570
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   1270
      Caption         =   "Escolha um vendedor na lista abaixo e clique no botão carregar, ou execute um duplo clique na lista para carregar o vendedor."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   128
      NoHTMLCaption   =   $"frmVendas_Lista_Vendedores.frx":0000
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   661
      DibPicture      =   "frmVendas_Lista_Vendedores.frx":0083
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
      Icon            =   "frmVendas_Lista_Vendedores.frx":85DB
      ShowMaximize    =   0   'False
      ShowMinimize    =   0   'False
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   2880
      Left            =   210
      TabIndex        =   0
      Top             =   1350
      Width           =   6660
      _ExtentX        =   11748
      _ExtentY        =   5080
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Text            =   "Nº"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "T"
         Text            =   "Vendedor"
         Object.Width           =   7885
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Object.Tag             =   "N"
         Text            =   "Comissão"
         Object.Width           =   2293
      EndProperty
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   210
      TabIndex        =   2
      Top             =   4260
      Width           =   6660
      _ExtentX        =   11748
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
   Begin DrawSuite2022.USAlphaImage USAlphaImage1 
      Height          =   585
      Left            =   300
      Top             =   600
      Width           =   585
      _ExtentX        =   1032
      _ExtentY        =   1032
      Image           =   "frmVendas_Lista_Vendedores.frx":88F5
      Props           =   29
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4680
      TabIndex        =   1
      Top             =   1320
      Width           =   45
   End
End
Attribute VB_Name = "frmVendas_Lista_Vendedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub ProcAtualizalista()
On Error GoTo tratar_erro

TextoFiltro = ""
With IIf(Vendas_Proposta = True, frmVendas_proposta, frmVendas_PI)
    If .txtIDCliente <> "" And .txtIDCliente <> "0" Then TextoFiltro = " and (VVC.IDCliente = " & .txtIDCliente & " and VV.Bloquear_venda_cliente = 'True' or VV.Bloquear_venda_cliente = 'False')"
End With
CamposFiltro = "VV.n_vendedor, VV.vendedor, VV.comissao"
Lista.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select " & CamposFiltro & " from vendas_vendedores VV LEFT JOIN Vendas_Vendedores_Clientes VVC ON VVC.IDVendedor = VV.ID where VV.dtvalidacao IS NOT NULL" & TextoFiltro & " group by " & CamposFiltro & " order by VV.vendedor", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBLISTA.EOF = False
        With Lista.ListItems
            .Add , , TBLISTA!n_vendedor
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!vendedor), "", TBLISTA!vendedor)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Comissao), "", TBLISTA!Comissao)
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

Private Sub BtnCarregar_Click()
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 Then Exit Sub
If VE = True And VI = False Then
    With IIf(Vendas_PI = True, frmVendas_PI, frmVendas_proposta)
        .txtVE.Text = Lista.SelectedItem.Text
        .txtVend_Ext = Lista.SelectedItem.ListSubItems.Item(1).Text
        Set TBVendas = CreateObject("adodb.recordset")
        TBVendas.Open "Select * from vendas_vendedores where n_vendedor = " & .txtVE.Text, Conexao, adOpenKeyset, adLockOptimistic
        If TBVendas.EOF = False Then
            .txtregiao.Text = IIf(IsNull(TBVendas!regiao), "", TBVendas!regiao)
        End If
        TBVendas.Close
    End With
End If
If VI = True And VE = False Then
    With IIf(Vendas_PI = True, frmVendas_PI, frmVendas_proposta)
        .txtVI.Text = Lista.SelectedItem.Text
        .txtvend_Int = Lista.SelectedItem.ListSubItems.Item(1).Text
    End With
End If
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

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

If Vendas_PI = True Then Caption = "Administrativo - Vendas - Pedido interno - Lista de vendedores"
ProcAtualizalista

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
If VE = True And VI = False Then
    With IIf(Vendas_PI = True, frmVendas_PI, frmVendas_proposta)
        .txtVE.Text = Lista.SelectedItem.Text
        .txtVend_Ext = Lista.SelectedItem.ListSubItems.Item(1).Text
        Set TBVendas = CreateObject("adodb.recordset")
        TBVendas.Open "Select * from vendas_vendedores where n_vendedor = " & .txtVE.Text, Conexao, adOpenKeyset, adLockOptimistic
        If TBVendas.EOF = False Then
            .txtregiao.Text = IIf(IsNull(TBVendas!regiao), "", TBVendas!regiao)
        End If
        TBVendas.Close
    End With
End If
If VI = True And VE = False Then
    With IIf(Vendas_PI = True, frmVendas_PI, frmVendas_proposta)
        .txtVI.Text = Lista.SelectedItem.Text
        .txtvend_Int = Lista.SelectedItem.ListSubItems.Item(1).Text
    End With
End If
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
