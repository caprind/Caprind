VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form frmQualidadePPAP_PlanoControle_Copiar 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Qualidade - PPAP - Plano de controle - Copiar"
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   1515
      Left            =   55
      TabIndex        =   5
      Top             =   900
      Width           =   8805
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
         TabIndex        =   12
         ToolTipText     =   "Familia."
         Top             =   1050
         Width           =   8415
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
         TabIndex        =   11
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
         ItemData        =   "frmQualidadePPAP_PlanoControle_Copiar.frx":0000
         Left            =   180
         List            =   "frmQualidadePPAP_PlanoControle_Copiar.frx":0025
         MousePointer    =   99  'Custom
         Style           =   2  'Dropdown List
         TabIndex        =   10
         ToolTipText     =   "Opções para filtro."
         Top             =   390
         Width           =   4365
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0E0&
         Height          =   510
         Left            =   4620
         TabIndex        =   6
         Top             =   210
         Width           =   3975
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
            MouseIcon       =   "frmQualidadePPAP_PlanoControle_Copiar.frx":00BD
            MousePointer    =   99  'Custom
            TabIndex        =   9
            Top             =   180
            Width           =   1155
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
            MouseIcon       =   "frmQualidadePPAP_PlanoControle_Copiar.frx":020F
            MousePointer    =   99  'Custom
            TabIndex        =   8
            Top             =   180
            Width           =   1275
         End
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
            MouseIcon       =   "frmQualidadePPAP_PlanoControle_Copiar.frx":0361
            MousePointer    =   99  'Custom
            TabIndex        =   7
            Top             =   180
            Width           =   1275
         End
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
         TabIndex        =   14
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
         TabIndex        =   13
         Top             =   840
         Width           =   1470
      End
   End
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
         MouseIcon       =   "frmQualidadePPAP_PlanoControle_Copiar.frx":04B3
         MousePointer    =   99  'Custom
         Picture         =   "frmQualidadePPAP_PlanoControle_Copiar.frx":0605
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
         MouseIcon       =   "frmQualidadePPAP_PlanoControle_Copiar.frx":0AA7
         MousePointer    =   99  'Custom
         Picture         =   "frmQualidadePPAP_PlanoControle_Copiar.frx":0BF9
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
         MouseIcon       =   "frmQualidadePPAP_PlanoControle_Copiar.frx":13CC
         MousePointer    =   99  'Custom
         Picture         =   "frmQualidadePPAP_PlanoControle_Copiar.frx":151E
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
      MouseIcon       =   "frmQualidadePPAP_PlanoControle_Copiar.frx":1C9D
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
      Left            =   55
      TabIndex        =   15
      Top             =   6030
      Width           =   8805
      _ExtentX        =   15531
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
End
Attribute VB_Name = "frmQualidadePPAP_PlanoControle_Copiar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public StrSql_PPAP_Localizar_Prod_Serv As String 'OK

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
If cmbfiltrarpor = "Família" Or cmbfiltrarpor = "Cliente" Then
    txtTexto.Visible = False
    cmbfamilia.Visible = True
    cmbfamilia.Clear
    cmbfamilia.AddItem ""
    If cmbfiltrarpor = "Família" Then
        ProcCarregaComboFamilia cmbfamilia, "familia <> 'Null'", True
    Else
        Set TBClientes = CreateObject("adodb.recordset")
        TBClientes.Open "Select IDCliente, NomeRazao from Clientes where NomeRazao <> 'Null' order by NomeRazao", Conexao, adOpenKeyset, adLockOptimistic
        If TBClientes.EOF = False Then
            Do While TBClientes.EOF = False
                cmbfamilia.AddItem Trim(TBClientes!NomeRazao)
                cmbfamilia.ItemData(cmbfamilia.NewIndex) = TBClientes!IDCliente
                TBClientes.MoveNext
            Loop
        End If
        cmbfamilia = frmQualidadePPAP.txtCliente
        TBClientes.Close
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

ListView1.ListItems.Clear
If txtTexto.Visible = True And txtTexto <> "" Or cmbfamilia.Visible = True And cmbfamilia <> "" Then
    If cmbfiltrarpor = "Código interno" Then
        If Optinicio.Value = True Then StrSql_PPAP_Localizar_Prod_Serv = "Select * from Projproduto where desenho like '" & txtTexto.Text & "%' order by desenho"
        If Optmeio.Value = True Then StrSql_PPAP_Localizar_Prod_Serv = "Select * from Projproduto where desenho like '%" & txtTexto.Text & "%' order by desenho"
        If Optfim.Value = True Then StrSql_PPAP_Localizar_Prod_Serv = "Select * from Projproduto where desenho like '%" & txtTexto.Text & "' order by desenho"
    End If
    If cmbfiltrarpor = "Código de referência" Then
        Set TBItem = CreateObject("adodb.recordset")
        If Optinicio.Value = True Then TBItem.Open "Select codproduto from item_aplicacoes where n_referencia like '" & txtTexto.Text & "%' order by codproduto", Conexao, adOpenKeyset, adLockOptimistic
        If Optmeio.Value = True Then TBItem.Open "Select codproduto from item_aplicacoes where n_referencia like '%" & txtTexto.Text & "%' order by codproduto", Conexao, adOpenKeyset, adLockOptimistic
        If Optfim.Value = True Then TBItem.Open "Select codproduto from item_aplicacoes where n_referencia like '%" & txtTexto.Text & "' order by codproduto", Conexao, adOpenKeyset, adLockOptimistic
        If TBItem.EOF = False Then
            Codproduto = 0
            Do While TBItem.EOF = False
                If Codproduto <> TBItem!Codproduto Then
                    StrSql_PPAP_Localizar_Prod_Serv = "Select * from Projproduto where codproduto = " & TBItem!Codproduto & " order by desenho"
                    ProcCarregaLista
                End If
                Codproduto = TBItem!Codproduto
                TBItem.MoveNext
            Loop
        End If
        TBItem.Close
        Exit Sub
    End If
    If cmbfiltrarpor = "Descrição" Then
        If Optinicio.Value = True Then StrSql_PPAP_Localizar_Prod_Serv = "Select * from Projproduto where descricao like '" & txtTexto.Text & "%' order by desenho"
        If Optmeio.Value = True Then StrSql_PPAP_Localizar_Prod_Serv = "Select * from Projproduto where descricao like '%" & txtTexto.Text & "%' order by desenho"
        If Optfim.Value = True Then StrSql_PPAP_Localizar_Prod_Serv = "Select * from Projproduto where descricao like '%" & txtTexto.Text & "' order by desenho"
    End If
    If cmbfiltrarpor = "Descrição comercial" Then
        If Optinicio.Value = True Then StrSql_PPAP_Localizar_Prod_Serv = "Select * from Projproduto where Descricaotecnica like '" & txtTexto.Text & "%' order by desenho"
        If Optmeio.Value = True Then StrSql_PPAP_Localizar_Prod_Serv = "Select * from Projproduto where Descricaotecnica like '%" & txtTexto.Text & "%' order by desenho"
        If Optfim.Value = True Then StrSql_PPAP_Localizar_Prod_Serv = "Select * from Projproduto where Descricaotecnica like '%" & txtTexto.Text & "' order by desenho"
    End If
    If cmbfiltrarpor = "Família" Then StrSql_PPAP_Localizar_Prod_Serv = "Select * from Projproduto where classe = '" & cmbfamilia & "' order by desenho"
    If cmbfiltrarpor = "Cliente" Then StrSql_PPAP_Localizar_Prod_Serv = "Select Projproduto.* FROM Projproduto_clientes INNER JOIN Projproduto ON Projproduto.codproduto = Projproduto_clientes.codproduto where Projproduto_clientes.IDCliente = " & cmbfamilia.ItemData(cmbfamilia.ListIndex) & " order by Projproduto.desenho"
    If cmbfiltrarpor = "Comprimento" Then
        If Optinicio.Value = True Then StrSql_PPAP_Localizar_Prod_Serv = "Select * from Projproduto where Comprimento like '" & txtTexto.Text & "%' order by desenho"
        If Optmeio.Value = True Then StrSql_PPAP_Localizar_Prod_Serv = "Select * from Projproduto where Comprimento like '%" & txtTexto.Text & "%' order by desenho"
        If Optfim.Value = True Then StrSql_PPAP_Localizar_Prod_Serv = "Select * from Projproduto where Comprimento like '%" & txtTexto.Text & "' order by desenho"
    End If
    If cmbfiltrarpor = "Largura" Then
        If Optinicio.Value = True Then StrSql_PPAP_Localizar_Prod_Serv = "Select * from Projproduto where Largura like '" & txtTexto.Text & "%' order by desenho"
        If Optmeio.Value = True Then StrSql_PPAP_Localizar_Prod_Serv = "Select * from Projproduto where Largura like '%" & txtTexto.Text & "%' order by desenho"
        If Optfim.Value = True Then StrSql_PPAP_Localizar_Prod_Serv = "Select * from Projproduto where Largura like '%" & txtTexto.Text & "' order by desenho"
    End If
    If cmbfiltrarpor = "Espessura" Then
        If Optinicio.Value = True Then StrSql_PPAP_Localizar_Prod_Serv = "Select * from Projproduto where Espessura like '" & txtTexto.Text & "%' order by desenho"
        If Optmeio.Value = True Then StrSql_PPAP_Localizar_Prod_Serv = "Select * from Projproduto where Espessura like '%" & txtTexto.Text & "%' order by desenho"
        If Optfim.Value = True Then StrSql_PPAP_Localizar_Prod_Serv = "Select * from Projproduto where Espessura like '%" & txtTexto.Text & "' order by desenho"
    End If
    If cmbfiltrarpor = "Dureza" Then
        If Optinicio.Value = True Then StrSql_PPAP_Localizar_Prod_Serv = "Select * from Projproduto where Dureza like '" & txtTexto.Text & "%' order by desenho"
        If Optmeio.Value = True Then StrSql_PPAP_Localizar_Prod_Serv = "Select * from Projproduto where Dureza like '%" & txtTexto.Text & "%' order by desenho"
        If Optfim.Value = True Then StrSql_PPAP_Localizar_Prod_Serv = "Select * from Projproduto where Dureza like '%" & txtTexto.Text & "' order by desenho"
    End If
    If cmbfiltrarpor = "Número do desenho" Then
        Set TBFIltro = CreateObject("adodb.recordset")
        If Optinicio.Value = True Then TBFIltro.Open "Select codproduto from item_aplicacoes where desenho like '" & txtTexto.Text & "%' order by Codproduto", Conexao, adOpenKeyset, adLockOptimistic
        If Optmeio.Value = True Then TBFIltro.Open "Select codproduto from item_aplicacoes where desenho like '%" & txtTexto.Text & "%' order by Codproduto", Conexao, adOpenKeyset, adLockOptimistic
        If Optfim.Value = True Then TBFIltro.Open "Select codproduto from item_aplicacoes where desenho like '%" & txtTexto.Text & "' order by Codproduto", Conexao, adOpenKeyset, adLockOptimistic
        If TBFIltro.EOF = False Then
            Codproduto = 0
            Do While TBFIltro.EOF = False
                If Codproduto <> TBFIltro!Codproduto Then
                    StrSql_PPAP_Localizar_Prod_Serv = "Select * from Projproduto where codproduto = " & TBFIltro!Codproduto & " order by desenho"
                    ProcCarregaLista
                End If
                Codproduto = TBFIltro!Codproduto
                TBFIltro.MoveNext
            Loop
        End If
        TBFIltro.Close
        Exit Sub
    End If
Else
    StrSql_PPAP_Localizar_Prod_Serv = "Select * from Projproduto order by desenho"
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
With frmQualidadePPAP_PlanoControle
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from QualidadePPAP_PlanoControle where ID = " & .TxtID, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Set TBItem = CreateObject("adodb.recordset")
        TBItem.Open "Select * from projproduto where desenho = '" & ListView1.SelectedItem & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBItem.EOF = False Then
            Set TBGravar = CreateObject("adodb.recordset")
            TBGravar.Open "Select * from QualidadePPAP_PlanoControle order by plano", Conexao, adOpenKeyset, adLockOptimistic
            TBGravar.AddNew
            TBGravar!Plano = TBAbrir!Plano
            TBGravar!Rev = 0
            TBGravar!DtEmissao = Date
            TBGravar!Responsavel = pubUsuario
            ProcEnviaDadosCopiar
            TBGravar.Update
            
            Set TBCiclo = CreateObject("adodb.recordset")
            TBCiclo.Open "Select plano.idplano, Planodimensao.IDdimensao from plano inner join Planodimensao on plano.IdPlano = Planodimensao.IdPlano where plano.Desenho = '" & TBItem!Desenho & "' and Planodimensao.PCP = 'True' order by plano.IdPlano, Planodimensao.indice", Conexao, adOpenKeyset, adLockOptimistic
            If TBCiclo.EOF = False Then
                
                Set TBFI = CreateObject("adodb.recordset")
                TBFI.Open "Select * from QualidadePPAP_PlanoControle_Dimensoes where IDPlanoControle = " & TBAbrir!ID, Conexao, adOpenKeyset, adLockOptimistic
                
                Do While TBCiclo.EOF = False
                    Set TBExecucao = CreateObject("adodb.recordset")
                    TBExecucao.Open "Select * from QualidadePPAP_PlanoControle_Dimensoes", Conexao, adOpenKeyset, adLockOptimistic
                    TBExecucao.AddNew
                    TBExecucao!IdPlanoControle = TBGravar!ID
                    TBExecucao!IDPlano = TBCiclo!IDPlano
                    TBExecucao!idDimensao = TBCiclo!idDimensao
                    TBExecucao!data = Date
                    TBExecucao!Responsavel = pubUsuario
                    
                    If TBFI.EOF = False Then
                        TBExecucao!Processo = TBFI!Processo
                        TBExecucao!Descricao = TBFI!Descricao
                        TBExecucao!Carac = TBFI!Carac
                        TBExecucao!TecnicaAvaliacao = TBFI!TecnicaAvaliacao
                        TBExecucao!AmostraFreq = TBFI!AmostraFreq
                        TBExecucao!MetodoControle = TBFI!MetodoControle
                        TBExecucao!PlanoReacao = TBFI!PlanoReacao
                        TBExecucao!DataEnsaio = TBFI!DataEnsaio
                        TBExecucao!QtdeEnsaiado = TBFI!QtdeEnsaiado
                        TBExecucao!Resultados = TBFI!Resultados
                        If TBFI!ok = True Then TBExecucao!ok = True Else TBExecucao!ok = False
                    End If
                    
                    TBExecucao.Update
                    TBExecucao.Close
                    TBCiclo.MoveNext
                    If TBFI.EOF = False Then TBFI.MoveNext
                Loop
                TBFI.Close
            End If
            TBCiclo.Close
                        
            TBGravar.Close
        End If
        TBItem.Close
    End If
    USMsgBox ("Plano de controle copiado com sucesso."), vbInformation, "CAPRIND v5.0"
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from QualidadePPAP_PlanoControle where ID = " & .TxtID, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        .ProcLimpaCampos
        .ProcPuxaDados
    End If
    TBAbrir.Close
    .Lista.ListItems.Clear
    .ProcCarregaLista
    '==================================
    Modulo = "Qualidade/PPAP/Plano de controle"
    Evento = "Novo"
    ID_documento = .TxtID
    Documento = "Plano de controle: " & txtPlano & " - Cód. interno:  " & .txtCodInterno
    Documento1 = ""
    ProcGravaEvento
    '==================================
End With
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviaDadosCopiar()
On Error GoTo tratar_erro

TBGravar!Tipo = TBAbrir!Tipo
TBGravar!contato = TBAbrir!contato
TBGravar!IDProduto = IIf(IsNull(TBItem!Codproduto), "", TBItem!Codproduto)
TBGravar!Codinterno = IIf(IsNull(TBItem!Desenho), "", TBItem!Desenho)
Set TBFI = CreateObject("adodb.recordset")
TBFI.Open "Select * from item_aplicacoes where codproduto = " & TBItem!Codproduto & " order by n_referencia", Conexao, adOpenKeyset, adLockOptimistic
If TBFI.EOF = False Then
    TBGravar!N_referencia = TBFI!N_referencia
End If
TBFI.Close
TBGravar!Aprovacao_engenharia = TBAbrir!Aprovacao_engenharia
TBGravar!Equipe = TBAbrir!Equipe
TBGravar!Organizacao_Planta = TBAbrir!Organizacao_Planta
TBGravar!CodOrganizacao = TBAbrir!CodOrganizacao
TBGravar!Organizacao_Aprovacao = TBAbrir!Organizacao_Aprovacao
TBGravar!Aprovacao_Qualidade = TBAbrir!Aprovacao_Qualidade
TBGravar!Outra_Aprovacao = TBAbrir!Outra_Aprovacao
TBGravar!Outra_Aprovacao2 = TBAbrir!Outra_Aprovacao2

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaLista()
On Error GoTo tratar_erro

Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open StrSql_PPAP_Localizar_Prod_Serv, Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    TBProduto.MoveLast
    PBLista.Min = 0
    PBLista.Max = TBProduto.RecordCount
    PBLista.Value = 1
    Contador = 0
    TBProduto.MoveFirst
    Do While TBProduto.EOF = False
        With ListView1.ListItems
            .Add , , TBProduto!Desenho
            .Item(.Count).SubItems(1) = IIf(IsNull(TBProduto!Descricao), "", TBProduto!Descricao)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBProduto!Unidade), "", TBProduto!Unidade)
        End With
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
