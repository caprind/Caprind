VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form Frm_centro_de_custo_copiar_prev 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Centro de custo - Copiar previsão"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3615
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Frm_centro_de_custo_copiar_prev.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   3615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ano"
      Height          =   660
      Left            =   0
      TabIndex        =   3
      Top             =   4530
      Width           =   3555
      Begin VB.ComboBox cmbAno 
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
         Height          =   330
         ItemData        =   "Frm_centro_de_custo_copiar_prev.frx":000C
         Left            =   1290
         List            =   "Frm_centro_de_custo_copiar_prev.frx":000E
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   180
         Width           =   855
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   3555
      _ExtentX        =   6271
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
      ButtonCaption1  =   "Copiar"
      ButtonEnabled1  =   0   'False
      ButtonIconSize1 =   32
      ButtonToolTipText1=   "Copiar (F7)"
      ButtonKey1      =   "3"
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
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   1410
      Top             =   1380
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "Frm_centro_de_custo_copiar_prev.frx":0010
      Count           =   1
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   3525
      Left            =   0
      TabIndex        =   0
      Top             =   975
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   6218
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "T"
         Text            =   "Mês"
         Object.Width           =   5054
      EndProperty
   End
End
Attribute VB_Name = "Frm_centro_de_custo_copiar_prev"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ProcCopiar()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & ", voce não está autorizado a alterar neste formulário."), vbInformation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With Lista
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            Permitido = True
            For InitFor1 = 1 To Frm_centro_de_custo.Lista2.ListItems.Count
                If Frm_centro_de_custo.Lista2.ListItems.Item(InitFor1).Checked = True Then
                    Set TBAbrir = CreateObject("adodb.recordset")
                    TBAbrir.Open "Select * from Usuarios_Setor_Previsao where ID = " & Frm_centro_de_custo.Lista2.ListItems(InitFor1), Conexao, adOpenKeyset, adLockOptimistic
                    If TBAbrir.EOF = False Then
                        Qtde = FunVerificaMes(IIf(.ListItems.Item(InitFor).ListSubItems(1) = "Maio", .ListItems.Item(InitFor).ListSubItems(1), Left(.ListItems.Item(InitFor).ListSubItems(1), 3)))
                        Set TBGravar = CreateObject("adodb.recordset")
                        'TBGravar.Open "Select * from Usuarios_Setor_Previsao where ID_CC = " & TBAbrir!ID_CC & " and Mes = " & Qtde & " and Ano = " & cmbAno & " and ID_PC = " & TBAbrir!ID_PC & " and Revisao = " & TBAbrir!Revisao, Conexao, adOpenKeyset, adLockOptimistic
                        TBGravar.Open "Select * from Usuarios_Setor_Previsao", Conexao, adOpenKeyset, adLockOptimistic
                        TBGravar.AddNew
                        TBGravar!ID_CC = TBAbrir!ID_CC
                        TBGravar!Data = Date
                        TBGravar!Responsavel = pubUsuario
                        TBGravar!Mes = Qtde
                        TBGravar!Ano = cmbAno
                        TBGravar!ID_PC = TBAbrir!ID_PC
                        TBGravar!valor = TBAbrir!valor
                        TBGravar!Revisao = TBAbrir!Revisao
                        TBGravar.Update
                        '==================================
                        Modulo = "Custos/Centro de custo/Copiar"
                        Evento = "Nova previsão orçamentária"
                        ID_documento = TBGravar!ID
                        Documento = "Código: " & Frm_centro_de_custo.txt_Codigo & " - Descrição: " & Frm_centro_de_custo.Txt_descricao
                        Documento1 = "Mês: " & Left(.ListItems.Item(InitFor).ListSubItems(1), 3) & " - Ano: " & cmbAno & " - Revisão : " & TBGravar!Revisao & " - Código contábil: " & Frm_centro_de_custo.Lista2.ListItems.Item(InitFor1).ListSubItems(6) & " - Conta contábil: " & Frm_centro_de_custo.Lista2.ListItems.Item(InitFor1).ListSubItems(7) & " - Valor: " & Format(TBGravar!valor, "###,##0.00")
                        ProcGravaEvento
                        '==================================
                    End If
                    TBAbrir.Close
                End If
            Next InitFor1
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) mes(es) antes de copiar a(s) previsão(ões) orçamentária(s)."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Previsão(ões) orçamentária(s) copiada(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    Unload Me
    With Frm_centro_de_custo
        M = FunVerificaMes(.TabFiltro.SelectedItem.key)
        If .Cmb_revisao_filtrar = "Todas" Then TextoFiltro = "" Else TextoFiltro = "and Usuarios_setor_previsao.Revisao = " & .Cmb_revisao_filtrar
        If .OptDomes.Value = True Then TextoFiltro1 = "Usuarios_setor_previsao.Mes = '" & M & "'" Else TextoFiltro1 = "Usuarios_setor_previsao.Mes <= '" & M & "'"
        Familiatext = "Select Usuarios_setor_previsao.*, tbl_familia.Codigo, tbl_familia.txt_descricao from Usuarios_setor_previsao INNER JOIN tbl_familia ON Usuarios_setor_previsao.ID_PC = tbl_familia.int_codfamilia where Usuarios_setor_previsao.ID_CC = " & .txtID & " and " & TextoFiltro1 & " and Usuarios_setor_previsao.Ano = '" & .cmbAno & "' " & TextoFiltro & " order by Usuarios_setor_previsao.mes desc"
        .ProcCarregaListaPrev Familiatext
    End With
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyF7: ProcCopiar
    'Case vbKeyF1: cmdAjuda_Click
    Case vbKeyEscape: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 3555, 5, True

With cmbAno
    .Clear
    AnoAtual = Year(Date)
    Do While AnoAtual <= (Year(Date) + 1)
        .AddItem AnoAtual
        AnoAtual = AnoAtual + 1
    Loop
    .Text = Year(Date)
End With

With Lista.ListItems
    .Clear
    contador = 1
    Do While contador <= 12
        .Add , , ""
        Select Case contador
            Case 1: Texto = "Janeiro"
            Case 2: Texto = "Fevereiro"
            Case 3: Texto = "Março"
            Case 4: Texto = "Abril"
            Case 5: Texto = "Maio"
            Case 6: Texto = "Junho"
            Case 7: Texto = "Julho"
            Case 8: Texto = "Agosto"
            Case 9: Texto = "Setembro"
            Case 10: Texto = "Outubro"
            Case 11: Texto = "Novembro"
            Case 12: Texto = "Dezembro"
        End Select
        .Item(.Count).SubItems(1) = Texto
        contador = contador + 1
    Loop
End With

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
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcCopiar
    'Case 3: ProcAjuda
    Case 4: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
