VERSION 5.00
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOpcoesGeral_TotalFat 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Opções gerais | Faturamento por mês"
   ClientHeight    =   6135
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   4650
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   4650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   435
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   4650
      _ExtentX        =   8202
      _ExtentY        =   767
      DibPicture      =   "frmOpcoesGeral_TotalFat.frx":0000
      CaptionDelimiter=   "|"
      CaptionOnCenter =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Icon            =   "frmOpcoesGeral_TotalFat.frx":A123
   End
   Begin VB.TextBox Txt_ID 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
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
      Left            =   1410
      MaxLength       =   60
      TabIndex        =   1
      Text            =   "0"
      Top             =   3840
      Visible         =   0   'False
      Width           =   950
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   3375
      Left            =   360
      TabIndex        =   0
      Top             =   2580
      Width           =   3885
      _ExtentX        =   6853
      _ExtentY        =   5953
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
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
         Object.Width           =   512
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Object.Tag             =   "T"
         Text            =   "Mês"
         Object.Width           =   1759
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Object.Tag             =   "N"
         Text            =   "Ano"
         Object.Width           =   1759
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Object.Tag             =   "N"
         Text            =   "Valor"
         Object.Width           =   2117
      EndProperty
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   0
      TabIndex        =   3
      Top             =   450
      Width           =   4635
      _ExtentX        =   8176
      _ExtentY        =   1720
      ButtonCount     =   7
      GradientColor2  =   14737632
      GradientColorOverRight1=   16315633
      GradientColorOverRight2=   15195350
      GripperColor    =   15195350
      IsStrech        =   -1  'True
      RightColor1     =   0
      RightColor2     =   0
      ShowEndPanel    =   0   'False
      Theme           =   1
      ButtonCaption1  =   "Novo"
      ButtonEnabled1  =   0   'False
      ButtonIconSize1 =   32
      ButtonToolTipText1=   "Novo (Insert)"
      ButtonKey1      =   "1"
      ButtonAlignment1=   2
      BeginProperty ButtonFont1 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft1     =   2
      ButtonTop1      =   2
      ButtonWidth1    =   33
      ButtonHeight1   =   21
      ButtonUseMaskColor1=   0   'False
      ButtonCaption2  =   "Salvar"
      ButtonEnabled2  =   0   'False
      ButtonIconSize2 =   32
      ButtonToolTipText2=   "Salvar (F3)"
      ButtonKey2      =   "2"
      ButtonAlignment2=   2
      BeginProperty ButtonFont2 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft2     =   37
      ButtonTop2      =   2
      ButtonWidth2    =   38
      ButtonHeight2   =   21
      ButtonUseMaskColor2=   0   'False
      ButtonCaption3  =   "Excluir"
      ButtonEnabled3  =   0   'False
      ButtonIconSize3 =   32
      ButtonToolTipText3=   "Excluir (F4)"
      ButtonKey3      =   "3"
      ButtonAlignment3=   2
      BeginProperty ButtonFont3 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft3     =   77
      ButtonTop3      =   2
      ButtonWidth3    =   39
      ButtonHeight3   =   21
      ButtonUseMaskColor3=   0   'False
      ButtonEnabled4  =   0   'False
      ButtonIconSize4 =   32
      ButtonAlignment4=   2
      ButtonType4     =   1
      ButtonStyle4    =   -1
      BeginProperty ButtonFont4 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState4    =   -1
      ButtonLeft4     =   118
      ButtonTop4      =   4
      ButtonWidth4    =   2
      ButtonHeight4   =   54
      ButtonCaption5  =   "Ajuda"
      ButtonEnabled5  =   0   'False
      ButtonIconSize5 =   32
      ButtonToolTipText5=   "Ajuda (F1)"
      ButtonKey5      =   "5"
      ButtonAlignment5=   2
      BeginProperty ButtonFont5 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft5     =   122
      ButtonTop5      =   2
      ButtonWidth5    =   36
      ButtonHeight5   =   21
      ButtonUseMaskColor5=   0   'False
      ButtonCaption6  =   "Sair"
      ButtonEnabled6  =   0   'False
      ButtonIconSize6 =   32
      ButtonToolTipText6=   "Sair (Esc)"
      ButtonKey6      =   "6"
      ButtonAlignment6=   2
      BeginProperty ButtonFont6 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft6     =   160
      ButtonTop6      =   2
      ButtonWidth6    =   26
      ButtonHeight6   =   21
      ButtonUseMaskColor6=   0   'False
      ButtonEnabled7  =   0   'False
      ButtonIconSize7 =   32
      ButtonKey7      =   "7"
      ButtonAlignment7=   2
      BeginProperty ButtonFont7 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState7    =   5
      ButtonLeft7     =   188
      ButtonTop7      =   2
      ButtonWidth7    =   24
      ButtonHeight7   =   24
      ButtonUseMaskColor7=   0   'False
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   600
         Top             =   240
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmOpcoesGeral_TotalFat.frx":A43D
         Count           =   1
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   360
      TabIndex        =   2
      Top             =   1680
      Width           =   3885
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
         Left            =   2130
         MaxLength       =   50
         TabIndex        =   6
         ToolTipText     =   "Valor."
         Top             =   390
         Width           =   1275
      End
      Begin VB.ComboBox Cmb_ano 
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
         ItemData        =   "frmOpcoesGeral_TotalFat.frx":D826
         Left            =   1305
         List            =   "frmOpcoesGeral_TotalFat.frx":D828
         Style           =   2  'Dropdown List
         TabIndex        =   5
         ToolTipText     =   "Ano de."
         Top             =   390
         Width           =   825
      End
      Begin VB.ComboBox Cmb_mes 
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
         ItemData        =   "frmOpcoesGeral_TotalFat.frx":D82A
         Left            =   480
         List            =   "frmOpcoesGeral_TotalFat.frx":D852
         Style           =   2  'Dropdown List
         TabIndex        =   4
         ToolTipText     =   "Mês de."
         Top             =   390
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Valor"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   68
         Left            =   2580
         TabIndex        =   9
         Top             =   180
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Ano"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   67
         Left            =   1575
         TabIndex        =   8
         Top             =   180
         Width           =   285
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Mês"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   66
         Left            =   750
         TabIndex        =   7
         Top             =   180
         Width           =   285
      End
   End
End
Attribute VB_Name = "frmOpcoesGeral_TotalFat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Novo_Fat_Mes As Boolean 'OK

Private Sub ProcExcluir()
On Error GoTo tratar_erro

Permitido = False
With Lista
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) valor(es)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            
            Permitido = True
            Conexao.Execute "DELETE from Impostos_FaturamentoMes where ID = " & .ListItems(InitFor)
            '==================================
            Modulo = "Configuração do sistema/Opções gerais"
            Evento = "Excluir valor faturado no mês"
            ID_documento = .ListItems(InitFor)
            Documento = "Empresa: " & frmOpcoesGeral.txtRazao
            Documento1 = "Mês: " & .ListItems(InitFor).SubItems(1) & " - Ano: " & .ListItems(InitFor).SubItems(2) & " - Valor: " & .ListItems(InitFor).SubItems(3)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) valor(es) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Valor(es) excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCampos
    Frame1.Enabled = False
    ProcCarregaLista
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvar()
On Error GoTo tratar_erro

If Frame1.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
Acao = "salvar"
If Cmb_mes = "" Then
    NomeCampo = "o mês"
    ProcVerificaAcao
    Cmb_mes.SetFocus
    Exit Sub
End If
If Cmb_ano = "" Then
    NomeCampo = "o ano"
    ProcVerificaAcao
    Cmb_ano.SetFocus
    Exit Sub
End If
valor = IIf(Txt_valor = "", 0, Txt_valor)
If valor <= 0 Then
    NomeCampo = "o valor"
    ProcVerificaAcao
    Txt_valor.SetFocus
    Exit Sub
End If
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from Impostos_FaturamentoMes where Mes = " & FunVerificaMes(Cmb_mes) & " and Ano = " & Cmb_ano & " and ID <> " & Txt_ID & " And ID_empresa = " & frmOpcoesGeral.txtIDEmpresa, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = False Then
    USMsgBox ("Já existe um valor cadastrado para este mês/ano."), vbExclamation, "CAPRIND v5.0"
    TBAbrir.Close
    Exit Sub
Else
TBGravar.AddNew
TBGravar!ID_empresa = frmOpcoesGeral.txtIDEmpresa
TBGravar!Mes = FunVerificaMes(Cmb_mes)
TBGravar!Ano = Cmb_ano
TBGravar!valor = Txt_valor
TBGravar.Update
Txt_ID = TBGravar!ID
TBGravar.Close
End If

ProcCarregaLista
If Novo_Fat_Mes = True Then
    USMsgBox ("Novo valor cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo valor faturado no mês"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar valor faturado no mês"
    If CodigoLista <> 0 And Lista.ListItems.Count <> 0 Then
        Lista.SelectedItem = Lista.ListItems(CodigoLista)
        Lista.SetFocus
    End If
End If
'==================================
Modulo = "Configuração do sistema/Opções gerais"
ID_documento = Txt_ID
Documento = "Empresa: " & frmOpcoesGeral.txtRazao
Documento1 = "Mês: " & Cmb_mes & " - Ano: " & Cmb_ano & " - Valor: " & Format(Txt_valor, "###,##0.00")
ProcGravaEvento
'==================================
Novo_Fat_Mes = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovo()
On Error GoTo tratar_erro

If Contador = 12 Then
    USMsgBox ("Não é permitido cadastrar mais que 12 meses."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
ProcLimpaCampos
Novo_Fat_Mes = True
Frame1.Enabled = True
Cmb_mes.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 4585, 7, True

ProcLimpaVariaveisPrincipais
ProcCarregaComboAno Cmb_ano, Year(Date) - 1, 1
ProcCarregaLista

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

If Novo_Fat_Mes = True Then
    If USMsgBox("O valor ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvar
        If Novo_Fat_Mes = True Then Exit Sub Else Unload Me
    End If
End If
Novo_Fat_Mes = False
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaLista()
On Error GoTo tratar_erro

Contador = 0
valor = 0
Lista.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Impostos_FaturamentoMes where ID_empresa = " & frmOpcoesGeral.txtIDEmpresa & " order by Ano, Mes", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    Do While TBLISTA.EOF = False
        With Lista.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = FunVerificaNumeroMes(IIf(IsNull(TBLISTA!Mes), "", TBLISTA!Mes))
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Ano), "", TBLISTA!Ano)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!valor), "", Format(TBLISTA!valor, "###,##0.00"))
            valor = valor + TBLISTA!valor
            TBLISTA.MoveNext
            Contador = Contador + 1
        End With
    Loop
End If
TBLISTA.Close
With frmOpcoesGeral
    NovoValor = Replace(valor, ",", ".")
    Conexao.Execute "Update Impostos Set Vlr_total_faturado = " & NovoValor & " where ID_empresa = " & .txtIDEmpresa
    .Txt_valor_total_faturado = Format(valor, "###,##0.00")
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCampos()
On Error GoTo tratar_erro

Txt_ID = 0
Cmb_mes.ListIndex = -1
Cmb_ano.ListIndex = -1
Txt_valor = ""
CodigoLista = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Lista
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then .ListItems.Item(InitFor).Checked = False Else .ListItems.Item(InitFor).Checked = True
        Next InitFor
    End With
Else
    ProcOrdenaListView Lista, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Impostos_FaturamentoMes where ID = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    ProcLimpaCampos
    ProcPuxaDados
    CodigoLista = Lista.SelectedItem.index
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcPuxaDados()
On Error GoTo tratar_erro

Txt_ID = TBLISTA!ID
Cmb_mes = FunVerificaNumeroMes(IIf(IsNull(TBLISTA!Mes), "", TBLISTA!Mes))
Cmb_ano = IIf(IsNull(TBLISTA!Ano), "", TBLISTA!Ano)
Txt_valor = IIf(IsNull(TBLISTA!valor), "", Format(TBLISTA!valor, "###,##0.00"))
Frame1.Enabled = True
Novo_Fat_Mes = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyInsert: If USToolBar1.ButtonState(1) = 0 Then ProcNovo
    Case vbKeyF3: If USToolBar1.ButtonState(1) = 0 Then ProcSalvar
    Case vbKeyF4: If USToolBar1.ButtonState(2) = 0 Then ProcExcluir
    'Case vbKeyF1: ProcAjuda
    Case vbKeyEscape: ProcSair
End Select
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_Valor_Change()
On Error GoTo tratar_erro
    
If Txt_valor <> "" Then
    VerifNumero = Txt_valor
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_valor = ""
        Txt_valor.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_Valor_LostFocus()
On Error GoTo tratar_erro

Txt_valor = Format(Txt_valor, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcNovo
    Case 2: ProcSalvar
    Case 3: ProcExcluir
    'Case 5: ProcAjuda
    Case 6: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
