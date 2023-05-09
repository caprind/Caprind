VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form frmQualidadePPAP_FMEA_LocalizarProduto 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Qualidade - PPAP - FMEA -  Localizar produto"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8925
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   8925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Centralziar na Tela
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   885
      Left            =   55
      TabIndex        =   0
      Top             =   10
      Width           =   8805
      Begin VB.CommandButton cmdAjuda 
         BackColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   7350
         MouseIcon       =   "frmQualidadePPAP_FMEA_LocalizarProduto.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "frmQualidadePPAP_FMEA_LocalizarProduto.frx":0152
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Ajuda (F1)"
         Top             =   180
         Width           =   630
      End
      Begin VB.CommandButton imgSair 
         BackColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   7980
         MouseIcon       =   "frmQualidadePPAP_FMEA_LocalizarProduto.frx":05F4
         MousePointer    =   99  'Custom
         Picture         =   "frmQualidadePPAP_FMEA_LocalizarProduto.frx":0746
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Sair (Esc)"
         Top             =   180
         Width           =   630
      End
      Begin VB.CommandButton cmdFiltrar 
         BackColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   180
         MouseIcon       =   "frmQualidadePPAP_FMEA_LocalizarProduto.frx":0F19
         MousePointer    =   99  'Custom
         Picture         =   "frmQualidadePPAP_FMEA_LocalizarProduto.frx":106B
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Filtrar (F2)"
         Top             =   180
         Width           =   630
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3585
      Left            =   60
      TabIndex        =   4
      Top             =   2430
      Width           =   8805
      _ExtentX        =   15531
      _ExtentY        =   6324
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
      MouseIcon       =   "frmQualidadePPAP_FMEA_LocalizarProduto.frx":17EA
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "T"
         Text            =   "Cód. interno"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "T"
         Text            =   "Descrição"
         Object.Width           =   11695
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Un."
         Object.Width           =   1058
      EndProperty
   End
   Begin MSComctlLib.ProgressBar PBLista 
      Height          =   255
      Left            =   60
      TabIndex        =   11
      Top             =   6030
      Width           =   8805
      _ExtentX        =   15531
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   1515
      Left            =   55
      TabIndex        =   5
      Top             =   900
      Width           =   8805
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0E0&
         Height          =   510
         Left            =   4620
         TabIndex        =   12
         Top             =   210
         Width           =   3975
         Begin VB.OptionButton Optmeio 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Meio frase"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1470
            MouseIcon       =   "frmQualidadePPAP_FMEA_LocalizarProduto.frx":1B04
            MousePointer    =   99  'Custom
            TabIndex        =   15
            Top             =   180
            Width           =   1275
         End
         Begin VB.OptionButton Optinicio 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Início frase"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   180
            MouseIcon       =   "frmQualidadePPAP_FMEA_LocalizarProduto.frx":1C56
            MousePointer    =   99  'Custom
            TabIndex        =   14
            Top             =   180
            Width           =   1275
         End
         Begin VB.OptionButton Optfim 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Fim frase"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   2760
            MouseIcon       =   "frmQualidadePPAP_FMEA_LocalizarProduto.frx":1DA8
            MousePointer    =   99  'Custom
            TabIndex        =   13
            Top             =   180
            Width           =   1155
         End
      End
      Begin VB.TextBox txtTexto 
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
         Height          =   315
         Left            =   180
         MaxLength       =   255
         TabIndex        =   7
         ToolTipText     =   "Texto para pesquisa."
         Top             =   1050
         Width           =   8415
      End
      Begin VB.ComboBox cmbfiltrarpor 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "frmQualidadePPAP_FMEA_LocalizarProduto.frx":1EFA
         Left            =   180
         List            =   "frmQualidadePPAP_FMEA_LocalizarProduto.frx":1F25
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         ToolTipText     =   "Opções para filtro."
         Top             =   390
         Width           =   4365
      End
      Begin VB.ComboBox cmbfamilia 
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
         Left            =   180
         MousePointer    =   99  'Custom
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   8
         ToolTipText     =   "Texto para pesquisa."
         Top             =   1050
         Width           =   8415
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparente
         Caption         =   "Filtrar por"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1942
         TabIndex        =   10
         Top             =   180
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparente
         Caption         =   "Texto para pesquisa"
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
         Left            =   3645
         TabIndex        =   9
         Top             =   840
         Width           =   1470
      End
   End
End
Attribute VB_Name = "frmQualidadePPAP_FMEA_LocalizarProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbFamilia_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
If cmbfamilia <> "" Then txtTexto = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbfiltrarpor_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
If cmbfiltrarpor = "Família" Or cmbfiltrarpor = "Cliente" Or cmbfiltrarpor = "Fornecedor" Then
    txtTexto.Visible = False
    cmbfamilia.Visible = True
    cmbfamilia.Clear
    cmbfamilia.AddItem ""
    If cmbfiltrarpor = "Família" Then
        ProcCarregaComboFamilia cmbfamilia, "familia <> 'Null'", True
    ElseIf cmbfiltrarpor = "Cliente" Then
            Set TBClientes = CreateObject("adodb.recordset")
            TBClientes.Open "Select IDCliente, NomeRazao from Clientes where NomeRazao <> 'Null' order by NomeRazao", Conexao, adOpenKeyset, adLockOptimistic
            If TBClientes.EOF = False Then
                Do While TBClientes.EOF = False
                    cmbfamilia.AddItem Trim(TBClientes!NomeRazao)
                    cmbfamilia.ItemData(cmbfamilia.NewIndex) = TBClientes!IDCliente
                    TBClientes.MoveNext
                Loop
            End If
            TBClientes.Close
            If frmQualidadePPAP_FMEA.txtCliente <> "" Then cmbfamilia = frmQualidadePPAP_FMEA.txtCliente
        Else
            Set TBFornecedor = CreateObject("adodb.recordset")
            TBFornecedor.Open "Select IDCliente, Nome_Razao from Compras_fornecedores where Nome_Razao <> 'Null' order by Nome_Razao", Conexao, adOpenKeyset, adLockOptimistic
            If TBFornecedor.EOF = False Then
                Do While TBFornecedor.EOF = False
                    cmbfamilia.AddItem Trim(TBFornecedor!Nome_Razao)
                    cmbfamilia.ItemData(cmbfamilia.NewIndex) = TBFornecedor!IDCliente
                    TBFornecedor.MoveNext
                Loop
            End If
            TBFornecedor.Close
            If frmQualidadePPAP_FMEA.txtfornecedor <> "" Then cmbfamilia = frmQualidadePPAP_FMEA.txtfornecedor
    End If
Else
    txtTexto.Visible = True
    cmbfamilia.Visible = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdFiltrar_Click()
On Error GoTo tratar_erro

CamposFiltro = "P.Codproduto, P.Desenho, P.Descricao, P.Unidade"
INNERJOINTEXTO = "Select " & CamposFiltro & " from ((Projproduto P LEFT JOIN item_aplicacoes IA ON IA.codproduto = P.codproduto) LEFT JOIN Projproduto_clientes PC ON PC.codproduto = P.codproduto) LEFT JOIN Projproduto_fabricante PFAB ON PFAB.Codproduto = P.codproduto"
If Cmb_ordenar = "Código interno" Then Ordenar = "P.desenho" Else Ordenar = "P.Descricao"
TextoFiltroPadrao = "P.Bloqueado = 'False' group by " & CamposFiltro & " order by " & Ordenar

If txtTexto.Visible = True And txtTexto <> "" Or cmbfamilia.Visible = True And cmbfamilia <> "" Then
    If cmbfiltrarpor = "Cliente" Then
        StrSqlLocProdPadrao = INNERJOINTEXTO & " where PC.IDCliente = " & cmbfamilia.ItemData(cmbfamilia.ListIndex) & " and " & TextoFiltroPadrao
    ElseIf cmbfiltrarpor = "Família" Then
            StrSqlLocProdPadrao = INNERJOINTEXTO & " where P.classe = '" & cmbfamilia & "' and " & TextoFiltroPadrao
        ElseIf cmbfiltrarpor = "Comprimento" Or cmbfiltrarpor = "Largura" Or cmbfiltrarpor = "Espessura" Then
                Select Case cmbfiltrarpor
                    Case "Comprimento": TextoFiltro = "P.Comprimento"
                    Case "Largura": TextoFiltro = "P.Largura"
                    Case "Espessura": TextoFiltro = "P.Espessura"
                End Select
                valor = txtTexto
                NovoValor = Replace(valor, ",", ".")
                StrSqlLocProdPadrao = INNERJOINTEXTO & " where " & TextoFiltro & " = " & NovoValor & " and " & TextoFiltroPadrao
            Else
                Select Case cmbfiltrarpor
                    Case "Código interno": TextoFiltro = "P.desenho"
                    Case "Código de referência": TextoFiltro = "IA.N_referencia"
                    Case "Número do desenho": TextoFiltro = "IA.desenho"
                    Case "Descrição": TextoFiltro = "P.descricao"
                    Case "Descrição comercial": TextoFiltro = "P.Descricaotecnica"
                    Case "Dureza": TextoFiltro = "P.Dureza"
                    Case "Part number": TextoFiltro = "PFAB.Part_number"
                End Select
                'StrSqlLocProdPadrao = INNERJOINTEXTO & " where " & TextoFiltro & " like " & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, txtTexto) & " and " & TextoFiltroPadrao
    End If
Else
    StrSqlLocProdPadrao = INNERJOINTEXTO & " where " & TextoFiltroPadrao
End If
ProcCarregaLista

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyEscape: Unload Me
    Case vbKeyReturn: ListView1_DblClick
    Case vbKeyF2: cmdFiltrar_Click
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

cmbfamilia.Clear
Set TBFamilia = CreateObject("adodb.recordset")
TBFamilia.Open "Select * from projfamilia order by familia", Conexao, adOpenKeyset, adLockOptimistic
If TBFamilia.EOF = False Then
    Do While TBFamilia.EOF = False
        cmbfamilia.AddItem TBFamilia!Familia
        TBFamilia.MoveNext
    Loop
End If
TBFamilia.Close
cmbfiltrarpor = "Código interno"
Optinicio.Value = True
txtTexto.Visible = True
cmbfamilia.Visible = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub imgSair_Click()
On Error GoTo tratar_erro

Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView ListView1, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListView1_DblClick()
On Error GoTo tratar_erro

If ListView1.ListItems.Count = 0 Then Exit Sub
With frmQualidadePPAP_FMEA
    Set TBItem = CreateObject("adodb.recordset")
    TBItem.Open "Select * from projproduto where desenho = '" & ListView1.SelectedItem & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBItem.EOF = False Then
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from item_aplicacoes where codproduto = " & TBItem!Codproduto & " order by n_referencia", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Cont = 0
            Do While TBAbrir.EOF = False
                .cmbreferencia.AddItem IIf(IsNull(TBAbrir!N_referencia), "", TBAbrir!N_referencia)
                TBAbrir.MoveNext
            Loop
        End If
        TBAbrir.Close
        .txtIDProduto.Text = IIf(IsNull(TBItem!Codproduto), "", TBItem!Codproduto)
        .txtCodInterno.Text = IIf(IsNull(TBItem!Desenho), "", TBItem!Desenho)
        .txtRevProduto.Text = IIf(IsNull(TBItem!RevDesenho), "", TBItem!RevDesenho)
        .txtDescricao.Text = IIf(IsNull(TBItem!Descricao), "", TBItem!Descricao)
    Else
        .txtCodInterno.Text = ListView1.SelectedItem
    End If
    TBItem.Close
End With
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaLista()
On Error GoTo tratar_erro

Desenho = ""
ListView1.ListItems.Clear
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open StrSqlLocProdPadrao, Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    TBProduto.MoveLast
    PBLista.Min = 0
    PBLista.Max = TBProduto.RecordCount
    PBLista.Value = 1
    Contador = 0
    TBProduto.MoveFirst
    Do While TBProduto.EOF = False
        If Desenho <> TBProduto!Desenho Then
            With ListView1.ListItems
                .Add , , TBProduto!Desenho
                .Item(.Count).SubItems(1) = IIf(IsNull(TBProduto!Descricao), "", TBProduto!Descricao)
                .Item(.Count).SubItems(2) = IIf(IsNull(TBProduto!Unidade), "", TBProduto!Unidade)
            End With
        End If
        Desenho = TBProduto!Desenho
        TBProduto.MoveNext
        Contador = Contador + 1
        PBLista.Value = Contador
    Loop
End If
TBProduto.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optfim_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optinicio_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optmeio_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtTexto_Change()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
If txtTexto <> "" Then cmbfamilia.ListIndex = -1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

