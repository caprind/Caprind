VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form Frm_Baixas_Lista_Saques 
   BackColor       =   &H00F9F9F9&
   BorderStyle     =   0  'None
   Caption         =   "Administrativo - Financeiro - Contas a pagar - Pagamento de contas - Lista de saques"
   ClientHeight    =   8070
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   7995
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Frm_Baixas_Lista_Saques.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8070
   ScaleWidth      =   7995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   7995
      _ExtentX        =   14102
      _ExtentY        =   609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Icon            =   "Frm_Baixas_Lista_Saques.frx":030A
   End
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   4530
      Top             =   150
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "Frm_Baixas_Lista_Saques.frx":0624
      Count           =   1
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Lista de conta(s) paga(s) com o saque selecionado"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   55
      TabIndex        =   13
      Top             =   4950
      Width           =   7875
      Begin MSComctlLib.ListView Lista1 
         Height          =   2355
         Left            =   180
         TabIndex        =   6
         Top             =   270
         Width           =   7515
         _ExtentX        =   13256
         _ExtentY        =   4154
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
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "N"
            Text            =   "Id"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Object.Tag             =   "D"
            Text            =   "Data"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Responsável"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Tag             =   "T"
            Text            =   "Fornecedor"
            Object.Width           =   6214
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Object.Tag             =   "N"
            Text            =   "Vlr. pago"
            Object.Width           =   1587
         EndProperty
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   55
      TabIndex        =   8
      Top             =   4110
      Width           =   7875
      Begin VB.TextBox Txt_valor_utilizado 
         Alignment       =   1  'Right Justify
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
         Left            =   6360
         TabIndex        =   5
         ToolTipText     =   "Valor pago."
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox Txt_data 
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
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Data da movimentação."
         Top             =   360
         Width           =   1365
      End
      Begin VB.TextBox Txt_valor 
         Alignment       =   1  'Right Justify
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
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Valor."
         Top             =   360
         Width           =   1305
      End
      Begin VB.TextBox Txt_responsavel 
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Responsável pelo cadastro."
         Top             =   360
         Width           =   3465
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vlr. pago*"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   6652
         TabIndex        =   12
         Top             =   150
         Width           =   750
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Data"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   690
         TabIndex        =   11
         Top             =   170
         Width           =   345
      End
      Begin VB.Label Label36 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5512
         TabIndex        =   10
         Top             =   165
         Width           =   360
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Responsável"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   2835
         TabIndex        =   9
         Top             =   165
         Width           =   915
      End
   End
   Begin VB.TextBox Txt_setfocus 
      Alignment       =   1  'Right Justify
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
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   2130
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Lista de saque(s) disponível(is)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2745
      Left            =   55
      TabIndex        =   7
      Top             =   1350
      Width           =   7875
      Begin MSComctlLib.ListView Lista 
         Height          =   2355
         Left            =   180
         TabIndex        =   1
         Top             =   240
         Width           =   7515
         _ExtentX        =   13256
         _ExtentY        =   4154
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
            Text            =   "Id"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Object.Tag             =   "D"
            Text            =   "Data"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Responsável"
            Object.Width           =   8333
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Object.Tag             =   "N"
            Text            =   "Valor"
            Object.Width           =   2117
         EndProperty
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   14
      Top             =   360
      Width           =   7875
      _ExtentX        =   13891
      _ExtentY        =   1720
      ButtonCount     =   5
      GradientColor2  =   14737632
      GradientColorOverRight1=   16315633
      GradientColorOverRight2=   15195350
      GripperColor    =   15195350
      IsStrech        =   -1  'True
      RightColor1     =   0
      RightColor2     =   0
      ShowEndPanel    =   0   'False
      Theme           =   1
      ButtonCaption1  =   "Salvar"
      ButtonEnabled1  =   0   'False
      ButtonIconSize1 =   32
      ButtonToolTipText1=   "Salvar (F3)"
      ButtonKey1      =   "1"
      ButtonAlignment1=   2
      BeginProperty ButtonFont1 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft1     =   2
      ButtonTop1      =   2
      ButtonWidth1    =   44
      ButtonHeight1   =   21
      ButtonUseMaskColor1=   0   'False
      ButtonEnabled2  =   0   'False
      ButtonIconSize2 =   32
      ButtonAlignment2=   2
      ButtonType2     =   1
      ButtonStyle2    =   -1
      BeginProperty ButtonFont2 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState2    =   -1
      ButtonLeft2     =   48
      ButtonTop2      =   4
      ButtonWidth2    =   2
      ButtonHeight2   =   54
      ButtonCaption3  =   "Ajuda"
      ButtonEnabled3  =   0   'False
      ButtonIconSize3 =   32
      ButtonToolTipText3=   "Ajuda (F1)"
      ButtonKey3      =   "3"
      ButtonAlignment3=   2
      BeginProperty ButtonFont3 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft3     =   52
      ButtonTop3      =   2
      ButtonWidth3    =   41
      ButtonHeight3   =   21
      ButtonUseMaskColor3=   0   'False
      ButtonCaption4  =   "Sair"
      ButtonEnabled4  =   0   'False
      ButtonIconSize4 =   32
      ButtonToolTipText4=   "Sair (Esc)"
      ButtonKey4      =   "4"
      ButtonAlignment4=   2
      BeginProperty ButtonFont4 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft4     =   95
      ButtonTop4      =   2
      ButtonWidth4    =   30
      ButtonHeight4   =   21
      ButtonUseMaskColor4=   0   'False
      ButtonEnabled5  =   0   'False
      ButtonIconSize5 =   32
      ButtonKey5      =   "5"
      ButtonAlignment5=   2
      BeginProperty ButtonFont5 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState5    =   5
      ButtonLeft5     =   127
      ButtonTop5      =   2
      ButtonWidth5    =   24
      ButtonHeight5   =   24
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   60
      TabIndex        =   15
      Top             =   7740
      Width           =   7875
      _ExtentX        =   13891
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
Attribute VB_Name = "Frm_Baixas_Lista_Saques"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Novo_Pagar_Conta_Saque As Boolean 'OK

Private Sub ProcSair()
On Error GoTo tratar_erro

Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvar()
On Error GoTo tratar_erro

With frm_Baixas
    If Txt_data = "" Then Exit Sub
    
    Acao = "salvar"
    Qtd = IIf(Txt_valor_utilizado = "", 0, Txt_valor_utilizado)
    If Qtd <= 0 Then
        NomeCampo = "o valor utilizado"
        ProcVerificaAcao
        Txt_valor_utilizado.SetFocus
        Permitido = False
        Exit Sub
    End If
    
    'Verifica saldo da antecipação
    Qtde = 0
    For InitFor = 1 To .Lista.ListItems.Count
        If .Lista.ListItems.Item(InitFor).Checked = True Then Qtde = Qtde + .Lista.ListItems.Item(InitFor).ListSubItems(3)
    Next InitFor
    qt = Qtd - Qtde
    
    valor = Lista.SelectedItem.ListSubItems(3)
    If qt > valor Then
        USMsgBox ("Não é permitido utilizar este saque, pois o valor utilizado é maior que o disponível."), vbExclamation, "CAPRIND v5.0"
        Txt_valor_utilizado.SetFocus
        Permitido = False
        Exit Sub
    End If
    
    'Verifica se o valor utilizado é maior que o valor disponivel no saque
    If .chbparcial.Value = 1 Then
        VP = Txt_valor_utilizado
        VD = .txt_VlrDocto.Text
        If VP = VD Then
            USMsgBox ("Não é permitido pagar parcial, pois o valor pago é o mesmo que o valor total da conta."), vbExclamation, "CAPRIND v5.0"
            Exit Sub
        End If
        If VP > VD Then
            USMsgBox ("Não é permitido pagar parcial, pois o valor pago é maior que o valor total da conta."), vbExclamation, "CAPRIND v5.0"
            Exit Sub
        End If
    End If
    
    Valor2 = 0
    For InitFor = 1 To .Lista.ListItems.Count
        If .Lista.ListItems.Item(InitFor).Checked = True Then
            Qtde = .Lista.ListItems.Item(InitFor).ListSubItems(3)
            qt = 0
            If Qtde = Qtd Then  'Valor antecipado igual a 0
                qt = Qtd
            ElseIf Qtde > Qtd Then 'Valor pago maior que o valor antecipado
                    qt = Qtd
                Else
                    qt = Qtde
            End If
            Qtd = Qtd - qt
        End If
    Next InitFor
    
    If Qtd > 0 Then
        For InitFor = 1 To frmContas_Pagar.lst_contas.ListItems.Count
            If frmContas_Pagar.lst_contas.ListItems.Item(InitFor).Checked = True Then
                Set TBGravar = CreateObject("adodb.recordset")
                TBGravar.Open "Select * from tbl_ContasPagar_Saque", Conexao, adOpenKeyset, adLockOptimistic
                TBGravar.AddNew
                TBGravar!IDSaque = Lista.SelectedItem
                TBGravar!Responsavel = pubUsuario
                
                TBGravar!Data = .txt_DtPagto.Value
                .txt_ValorPago = Txt_valor_utilizado
                If .chbparcial.Value = 0 Then TBGravar!IDintconta = frmContas_Pagar.lst_contas.ListItems.Item(InitFor)
                
                If Contador2 > 1 Then
                    Set TBAbrir = CreateObject("adodb.recordset")
                    TBAbrir.Open "Select * from tbl_contaspagar where idintconta = " & frmContas_Pagar.lst_contas.ListItems.Item(InitFor), Conexao, adOpenKeyset, adLockOptimistic
                    If TBAbrir.EOF = False Then
                         TBGravar!Valor_utilizado = TBAbrir!dbl_valorpagto
                    End If
                    TBAbrir.Close
                Else
                    TBGravar!Valor_utilizado = Qtd
                End If
                
                TBGravar.Update
            End If
        Next InitFor
        IDlista = Lista.SelectedItem 'ID do saque
        
        USMsgBox ("Saque utilizado com sucesso."), vbInformation, "CAPRIND v5.0"
        '==================================
        Modulo = "Financeiro/Contas a pagar"
        Evento = "Utilizar saque"
        ID_documento = Lista.SelectedItem
        Documento = "Data: " & Txt_data & " - Responsável: " & Txt_responsavel
        Documento1 = ""
        ProcGravaEvento
        '==================================
    Else
        .txt_ValorPago = Txt_valor_utilizado
    End If
End With
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyF3: ProcSalvar
    Case vbKeyEscape: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 7875, 5, True
ProcCarregaListaDisponiveis

With frmContas_Pagar
    Contador2 = 0
    For InitFor = 1 To .lst_contas.ListItems.Count
        If .lst_contas.ListItems.Item(InitFor).Checked = True Then Contador2 = Contador2 + 1
    Next InitFor
End With
With Txt_valor_utilizado
    If Contador2 > 1 Then
        .Locked = True
        .TabStop = False
    Else
        .Locked = False
        .TabStop = True
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaListaDisponiveis()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from tbl_instituicoes_transf where banco_remetente = '" & frm_Baixas.cmb_Banco & "' and Tipo = 'S' and Saldo > 0 order by data_transf desc", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    TBLISTA.MoveLast
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    contador = 0
    TBLISTA.MoveFirst
    Do While TBLISTA.EOF = False
        With Lista.ListItems
            .Add , , TBLISTA!id_transf
            .Item(.Count).SubItems(1) = Format(TBLISTA!data_transf, "dd/mm/yy")
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Responsavel), "", TBLISTA!Responsavel)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Saldo), "", Format(TBLISTA!Saldo, "###,##0.00"))
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

Private Sub ProcCarregaListaUtilizados()
On Error GoTo tratar_erro

Lista1.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from tbl_ContasPagar_Saque where IDSaque = " & Lista.SelectedItem & " order by data desc, ID desc", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    TBLISTA.MoveLast
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    contador = 0
    TBLISTA.MoveFirst
    Do While TBLISTA.EOF = False
        With Lista1.ListItems
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select * from tbl_contaspagar where IDintconta = " & TBLISTA!IDintconta & " and logsit = 'S'", Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    Do While TBAbrir.EOF = False
                        .Add , , TBAbrir!IDintconta
                        .Item(.Count).SubItems(1) = IIf(IsNull(TBAbrir!DataBaixa), "", Format(TBAbrir!DataBaixa, "dd/mm/yy"))
                        .Item(.Count).SubItems(2) = IIf(IsNull(TBAbrir!resppag), "", TBAbrir!resppag)
                        .Item(.Count).SubItems(3) = IIf(IsNull(TBAbrir!Txt_fornecedor), "", TBAbrir!Txt_fornecedor)
                        .Item(.Count).SubItems(4) = IIf(IsNull(TBAbrir!ValorPago), "", Format(TBAbrir!ValorPago, "###,##0.00"))
                        TBAbrir.MoveNext
                    Loop
                End If
                TBAbrir.Close
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

Private Sub Lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView Lista, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 Then Exit Sub
Txt_data = Lista.SelectedItem.ListSubItems(1)
Txt_responsavel = Lista.SelectedItem.ListSubItems(2)
Txt_valor = Lista.SelectedItem.ListSubItems(3)
If frm_Baixas.chbparcial.Value = 0 Then Txt_valor_utilizado = frm_Baixas.txt_VlrDocto
Frame8.Enabled = True
CodigoLista = Lista.SelectedItem.index
ProcCarregaListaUtilizados

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView Lista1, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_valor_utilizado_Change()
On Error GoTo tratar_erro

If Txt_valor_utilizado <> "" Then
    VerifNumero = Txt_valor_utilizado
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_valor_utilizado = ""
        Txt_valor_utilizado.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_valor_utilizado_LostFocus()
On Error GoTo tratar_erro

Txt_valor_utilizado = Format(Txt_valor_utilizado, "###,##0.00")
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcSalvar
    'Case 3: ProcAjuda
    Case 4: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
