VERSION 5.00
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProcessos_gupoop 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Engenharia - Processos - Fases - Grupo/op."
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6285
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   6285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   5280
      Top             =   120
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frmProcessos_gupoop.frx":0000
      Count           =   1
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
      Left            =   1260
      MaxLength       =   60
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Text            =   "0"
      Top             =   3840
      Visible         =   0   'False
      Width           =   950
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   2630
      Left            =   60
      TabIndex        =   5
      Top             =   2875
      Width           =   6165
      _ExtentX        =   10874
      _ExtentY        =   4630
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Grupo/op."
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Descrição"
         Object.Width           =   6131
      EndProperty
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   55
      TabIndex        =   10
      Top             =   2010
      Width           =   6165
      Begin VB.CommandButton cmdFiltrar 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   5670
         Picture         =   "frmProcessos_gupoop.frx":410D
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Filtrar (F2)"
         Top             =   390
         Width           =   315
      End
      Begin VB.ComboBox cmbfiltrarpor 
         BackColor       =   &H00FFFFFF&
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
         Height          =   330
         ItemData        =   "frmProcessos_gupoop.frx":4528
         Left            =   180
         List            =   "frmProcessos_gupoop.frx":4532
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Opções para filtro."
         Top             =   390
         Width           =   2685
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
         Left            =   2880
         TabIndex        =   1
         ToolTipText     =   "Texto para pesquisa."
         Top             =   390
         Width           =   2775
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Texto para pesquisa"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3532
         TabIndex        =   12
         Top             =   180
         Width           =   1470
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   1102
         TabIndex        =   11
         Top             =   180
         Width           =   840
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   55
      TabIndex        =   13
      Top             =   0
      Width           =   6165
      _ExtentX        =   10874
      _ExtentY        =   1720
      ButtonCount     =   8
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
      ButtonCaption4  =   "Atualizar"
      ButtonEnabled4  =   0   'False
      ButtonIconSize4 =   32
      ButtonToolTipText4=   "Utilizado pelo administrador do sistema."
      ButtonKey4      =   "4"
      ButtonAlignment4=   2
      BeginProperty ButtonFont4 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft4     =   118
      ButtonTop4      =   2
      ButtonWidth4    =   50
      ButtonHeight4   =   21
      ButtonUseMaskColor4=   0   'False
      ButtonEnabled5  =   0   'False
      ButtonIconSize5 =   32
      ButtonAlignment5=   2
      ButtonType5     =   1
      ButtonStyle5    =   -1
      BeginProperty ButtonFont5 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState5    =   -1
      ButtonLeft5     =   170
      ButtonTop5      =   4
      ButtonWidth5    =   2
      ButtonHeight5   =   54
      ButtonCaption6  =   "Ajuda"
      ButtonEnabled6  =   0   'False
      ButtonIconSize6 =   32
      ButtonToolTipText6=   "Ajuda (F1)"
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
      ButtonLeft6     =   174
      ButtonTop6      =   2
      ButtonWidth6    =   36
      ButtonHeight6   =   21
      ButtonUseMaskColor6=   0   'False
      ButtonCaption7  =   "Sair"
      ButtonEnabled7  =   0   'False
      ButtonIconSize7 =   32
      ButtonToolTipText7=   "Sair (Esc)"
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
      ButtonLeft7     =   212
      ButtonTop7      =   2
      ButtonWidth7    =   26
      ButtonHeight7   =   21
      ButtonUseMaskColor7=   0   'False
      ButtonEnabled8  =   0   'False
      ButtonIconSize8 =   32
      ButtonKey8      =   "8"
      ButtonAlignment8=   2
      BeginProperty ButtonFont8 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState8    =   5
      ButtonLeft8     =   240
      ButtonTop8      =   2
      ButtonWidth8    =   24
      ButtonHeight8   =   24
      ButtonUseMaskColor8=   0   'False
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   60
      TabIndex        =   14
      Top             =   5520
      Width           =   6165
      _ExtentX        =   10874
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
   Begin VB.Frame Frame2 
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
      Height          =   1035
      Left            =   55
      TabIndex        =   7
      Top             =   960
      Width           =   6165
      Begin VB.TextBox Txt_descricao 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   4
         ToolTipText     =   "Descrição."
         Top             =   570
         Width           =   4905
      End
      Begin VB.TextBox Txt_grupo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   3
         ToolTipText     =   "Grupo/op."
         Top             =   180
         Width           =   4905
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Descricao :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   225
         TabIndex        =   9
         Top             =   570
         Width           =   795
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo/op. :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   8
         Top             =   180
         Width           =   840
      End
   End
End
Attribute VB_Name = "frmProcessos_gupoop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Novo_Grupo As Boolean 'OK
Dim StrSql_Grupo As String 'OK

Private Sub ProcExcluir()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With Lista
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) grupo/op.(s)?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            
            Conexao.Execute "DELETE from Grupo_op where Id = " & .ListItems(InitFor)
            
            '==================================
            Modulo = "Engenharia/Processos/Fases/Grupo/op."
            Evento = "Excluir"
            ID_documento = .ListItems(InitFor)
            Documento = "Grupo/op.: " & .ListItems(InitFor).SubItems(1) & " - Descrição: " & .ListItems(InitFor).SubItems(2)
            Documento1 = ""
            ProcGravaEvento
            '==================================
            
        End If
    Next InitFor
End With

If Permitido = False Then
    USMsgBox ("Informe o(s) grupo/op.(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Grupo/op.(s) excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCampos
    Frame2.Enabled = False
    StrSql_Grupo = "Select * from Grupo_op order by Grupo_op"
    ProcCarregaLista
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdFiltrar_Click()
On Error GoTo tratar_erro

If txtTexto <> "" Then
    Select Case cmbfiltrarpor
        Case "Grupo/op.": StrSql_Grupo = "Select * from Grupo_op where Grupo_op like '" & txtTexto & "%' order by Grupo_op"
        Case "Descrição": StrSql_Grupo = "Select * from Grupo_op where Descricao like '" & txtTexto & "%' order by Grupo_op"
    End Select
Else
    StrSql_Grupo = "Select * from Grupo_op order by Grupo_op"
End If
ProcCarregaLista

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvar()
On Error GoTo tratar_erro

If Frame2.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
If Txt_grupo.Text = "" Then
    USMsgBox ("Informe o grupo/op. antes de salvar."), vbInformation, "CAPRIND v5.0"
    Txt_grupo.SetFocus
    Exit Sub
End If
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from Grupo_op where ID = " & Txt_ID, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = False Then
    Conexao.Execute "Update Fases Set Grupo_op = '" & Txt_grupo & "' where Grupo_op = '" & TBGravar!Grupo_op & "'"
    Conexao.Execute "Update Ordemservico Set Grupo_op = '" & Txt_grupo & "' where Grupo_op = '" & TBGravar!Grupo_op & "'"
    Conexao.Execute "Update Plano Set Grupo_op = '" & Txt_grupo & "' where Grupo_op = '" & TBGravar!Grupo_op & "'"
Else
    TBGravar.AddNew
End If
TBGravar!Grupo_op = Txt_grupo
TBGravar!Descricao = Txt_descricao
TBGravar.Update
Txt_ID = TBGravar!ID
TBGravar.Close
StrSql_Grupo = "Select * from Grupo_op order by Grupo_op"
ProcCarregaLista
If Novo_Grupo = True Then
    USMsgBox ("Novo grupo/op. cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar"
    If CodigoLista <> 0 And Lista.ListItems.Count <> 0 Then
        Lista.SelectedItem = Lista.ListItems(CodigoLista)
        Lista.SetFocus
    End If
End If
'==================================
Modulo = "Engenharia/Processos/Fases/Grupo/op."
ID_documento = Txt_ID
Documento = "Grupo/op.: " & Txt_grupo & " - Descrição: " & Txt_descricao
Documento1 = ""
ProcGravaEvento
'==================================
Novo_Grupo = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovo()
On Error GoTo tratar_erro

ProcLimpaCampos
Novo_Grupo = True
Frame2.Enabled = True
Txt_grupo.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 6165, 8, True
ProcLimpaVariaveisPrincipais
StrSql_Grupo = "Select * from Grupo_op order by Grupo_op"
ProcCarregaLista
cmbfiltrarpor = "Grupo/op."

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procAtualiza()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If InputBox("Informe a senha para liberar.") = "280362P" Then
    If USMsgBox("Deseja realmente cadastrar os grupos/op. das fases dos processos?", vbYesNo, "CAPRIND v5.0") = vbYes Then
        Set TBFases = CreateObject("adodb.recordset")
        TBFases.Open "Select * from fases order by grupo_op", Conexao, adOpenKeyset, adLockOptimistic
        If TBFases.EOF = False Then
            TBFases.MoveLast
            PBLista.Min = 0
            PBLista.Max = TBFases.RecordCount
            PBLista.Value = 1
            Contador = 0
            TBFases.MoveFirst
            Do While TBFases.EOF = False
                If IsNull(TBFases!Grupo_op) = False And TBFases!Grupo_op <> "" And TBFases!Grupo_op <> "0" Then
                    Set TBGravar = CreateObject("adodb.recordset")
                    TBGravar.Open "Select * from Grupo_op where Grupo_op = '" & TBFases!Grupo_op & "'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBGravar.EOF = True Then TBGravar.AddNew
                    TBGravar!Grupo_op = TBFases!Grupo_op
                    TBGravar.Update
                    TBGravar.Close
                End If
                TBFases.MoveNext
                Contador = Contador + 1
                PBLista.Value = Contador
            Loop
        End If
        TBFases.Close
        USMsgBox ("Cadastros efetuados com sucesso."), vbInformation, "CAPRIND v5.0"
        '==================================
        Modulo = "Engenharia/Processos/Fases/Grupo/op."
        Evento = "Cadastrar"
        ID_documento = 0
        Documento = ""
        Documento1 = ""
        ProcGravaEvento
        '==================================
        StrSql_Grupo = "Select * from Grupo_op order by Grupo_op"
        ProcCarregaLista
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

If Novo_Grupo = True Then
    If USMsgBox("O grupo/op. ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvar
        If Novo_Grupo = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
Novo_Grupo = False
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaLista()
On Error GoTo tratar_erro

Lista.ListItems.Clear
If StrSql_Grupo = "" Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open StrSql_Grupo, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    TBLISTA.MoveLast
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    TBLISTA.MoveFirst
    Do While TBLISTA.EOF = False
        With Lista.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Grupo_op), "", TBLISTA!Grupo_op)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Descricao), "", TBLISTA!Descricao)
            TBLISTA.MoveNext
            Contador = Contador + 1
            PBLista.Value = Contador
        End With
    Loop
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCampos()
On Error GoTo tratar_erro

Txt_ID = 0
Txt_grupo = ""
Txt_descricao = ""
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
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                
                Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open "Select * from Fases where Grupo_op = '" & .ListItems(InitFor).SubItems(1) & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    TBAbrir.Close
                    GoTo Proximo
                End If
                Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open "Select * from Ordemservico where Grupo_op = '" & .ListItems(InitFor).SubItems(1) & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    TBAbrir.Close
                    GoTo Proximo
                End If
                Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open "Select * from Plano where Grupo_op = '" & .ListItems(InitFor).SubItems(1) & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    TBAbrir.Close
                    GoTo Proximo
                End If
                TBAbrir.Close
                
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
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

Private Sub Lista_DblClick()
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 Then Exit Sub
frmProcessos.txtgrupo_op = Lista.SelectedItem.ListSubItems(1)
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With Lista
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select * from Fases where Grupo_op = '" & .ListItems(InitFor).SubItems(1) & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                USMsgBox ("Não é permitido excluir este grupo/op., pois a mesma está sendo utilizada no módulo Processos/Gerenciamento."), vbExclamation, "CAPRIND v5.0"
                TBAbrir.Close
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select * from Ordemservico where Grupo_op = '" & .ListItems(InitFor).SubItems(1) & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                USMsgBox ("Não é permitido excluir este grupo/op., pois a mesma está sendo utilizada no módulo PCP/Gerenciamento de ordem."), vbExclamation, "CAPRIND v5.0"
                TBAbrir.Close
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select * from Plano where Grupo_op = '" & .ListItems(InitFor).SubItems(1) & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                USMsgBox ("Não é permitido excluir este grupo/op., pois a mesma está sendo utilizada no módulo Qualidade/Plano de inspeção."), vbExclamation, "CAPRIND v5.0"
                TBAbrir.Close
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            TBAbrir.Close
        End If
    Next InitFor
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Grupo_op where ID = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    ProcLimpaCampos
    ProcPuxaDados
    Frame2.Enabled = True
    CodigoLista = Lista.SelectedItem.index
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcPuxaDados()
On Error GoTo tratar_erro

Txt_ID = TBLISTA!ID
Txt_grupo = IIf(IsNull(TBLISTA!Grupo_op), "", TBLISTA!Grupo_op)
Txt_descricao = IIf(IsNull(TBLISTA!Descricao), "", TBLISTA!Descricao)
Novo_Grupo = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyInsert: ProcNovo
    Case vbKeyF3: ProcSalvar
    Case vbKeyF4: ProcExcluir
    Case vbKeyReturn: Lista_DblClick
    'Case vbKeyF1: cmdAjuda_Click
    Case vbKeyEscape: ProcSair
End Select
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_grupo_LostFocus()
On Error GoTo tratar_erro

If IsNumeric(Txt_grupo) = False Or Txt_grupo = "" Then Exit Sub
Select Case Len(Txt_grupo)
    Case 1: Txt_grupo = "00000" & Txt_grupo
    Case 2: Txt_grupo = "0000" & Txt_grupo
    Case 3: Txt_grupo = "000" & Txt_grupo
    Case 4: Txt_grupo = "00" & Txt_grupo
    Case 5: Txt_grupo = "0" & Txt_grupo
End Select
    
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
    Case 4: procAtualiza
    'Case 6: ProcAjuda
    Case 7: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
