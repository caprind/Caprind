VERSION 5.00
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProd_configurar_listas 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Produtividade | Configurar lista"
   ClientHeight    =   8265
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   7440
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmprod_configurar_listas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8265
   ScaleWidth      =   7440
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USButton btnSalvar 
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   6900
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   1508
      DibPicture      =   "frmprod_configurar_listas.frx":000C
      Caption         =   "Gravar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   5263559
      BorderColorDown =   2039646
      BorderColorOver =   3026574
      GradientColor1  =   5263559
      GradientColor2  =   5263559
      GradientColor3  =   5263559
      GradientColor4  =   5263559
      GradientColorOver1=   3026574
      GradientColorOver2=   3026574
      GradientColorOver3=   3026574
      GradientColorOver4=   3026574
      GradientColorDown1=   2039646
      GradientColorDown2=   2039646
      GradientColorDown3=   2039646
      GradientColorDown4=   2039646
      PicAlign        =   8
      PicSize         =   2
      PicSizeH        =   24
      PicSizeW        =   24
      ShowFocusRect   =   0   'False
      Theme           =   4
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   465
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   7440
      _ExtentX        =   13123
      _ExtentY        =   820
      CaptionDelimiter=   "|"
      CaptionOnCenter =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowMaximizeButton=   0   'False
      ShowMinimizeButton=   0   'False
   End
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   2
      Top             =   7860
      Width           =   7440
      _ExtentX        =   13123
      _ExtentY        =   714
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   6225
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   7125
      _ExtentX        =   12568
      _ExtentY        =   10980
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
         Object.Tag             =   "T"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "T"
         Text            =   "Descrição"
         Object.Width           =   11351
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Coluna"
         Object.Width           =   0
      EndProperty
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   60
      TabIndex        =   1
      Top             =   8730
      Width           =   7125
      _ExtentX        =   12568
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
Attribute VB_Name = "frmProd_configurar_listas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ProcCarregarLista()
On Error GoTo tratar_erro

Lista.ListItems.Clear

Contador2 = 30
Contador = 0
ContadorReg = 1
PBLista.Min = 0
PBLista.Max = Contador2
PBLista.Value = 1

Do While Contador <> Contador2
    Contador = Contador + 1
    PBLista.Value = Contador
    
    NColuna = IIf(Contador < 10, "0" & Contador, Contador)
    TextoColunas = ""
    

            Select Case Contador
                Case 1: TextoColunas = "ID"
                Case 2: TextoColunas = "Data"
                Case 3: TextoColunas = "Setor"
                Case 4: TextoColunas = "Ordem"
                Case 5: TextoColunas = "Cód. interno"
                Case 6: TextoColunas = "Cód. de ref."
                Case 7: TextoColunas = "Descrição"
                Case 8: TextoColunas = "OS"
                Case 9: TextoColunas = "Fase"
                Case 10: TextoColunas = "Grupo/op."
                Case 11: TextoColunas = "Posto de trabalho"
                Case 12: TextoColunas = "Descrição"
                Case 13: TextoColunas = "Operador."
                Case 14: TextoColunas = "Turno."
                Case 15: TextoColunas = "Prep. prevista"
                Case 16: TextoColunas = "Prep. utilizada"
                Case 17: TextoColunas = "Exec. prev."
                Case 18: TextoColunas = "Exec. utilizada"
                Case 19: TextoColunas = "Hrs prevista"
                Case 20: TextoColunas = "Qtde. prevista"
                Case 21: TextoColunas = "hrs. utilizadas"
                Case 22: TextoColunas = "QT. Apontamento"
                Case 23: TextoColunas = "QT. OK"
                Case 24: TextoColunas = "QT. NC"
                Case 25: TextoColunas = "QT. Acumulada OS"
                Case 26: TextoColunas = "Efic. Preparação"
                Case 27: TextoColunas = "Efic. Execução"
                Case 28: TextoColunas = "Efic. Média"
            End Select
    
    If TextoColunas <> "" Then
        With Lista.ListItems
            .Add , , Contador
            .Item(.Count).SubItems(1) = TextoColunas
            .Item(.Count).SubItems(2) = NColuna
            
            TextoFiltro = "Coluna" & Contador & " = 'True'"
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select ID from Usuarios_monitor_trabalho where ID_usuario = " & pubIDUsuario & " and Modulo = '" & Modulo & "' and " & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                .Item(ContadorReg).Checked = True
            End If
            TBAbrir.Close
        End With
        ContadorReg = ContadorReg + 1
      '  Contador = Contador + 1
    End If
Loop

1:

Exit Sub
tratar_erro:
    If Err.Number = 365 Then GoTo 1
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvar()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & ", voce não está autorizado a alterar neste formulário."), vbInformation, "CAPRIND v5.0"
    Exit Sub
End If
Acao = "salvar"

With Lista
    For InitFor = 1 To .ListItems.Count
        Set TBGravar = CreateObject("adodb.recordset")
        StrSql = "Select * from Usuarios_monitor_trabalho where ID_usuario = " & Int(pubIDUsuario) & " and Modulo = 'PCP/Relatórios/Produtividade'"
        'Debug.print StrSql
        
        TBGravar.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
        If TBGravar.EOF = True Then TBGravar.AddNew
        TBGravar!ID_usuario = pubIDUsuario
        TBGravar!Modulo = "PCP/Relatórios/Produtividade"
        If .ListItems.Item(InitFor).Checked = True Then
            Permitido = True
        Else
            Permitido = False
        End If
        
        
        Select Case .ListItems.Item(InitFor).ListSubItems(2)
            Case "01": TBGravar!Coluna1 = Permitido
            Case "02": TBGravar!Coluna2 = Permitido
            Case "03": TBGravar!Coluna3 = Permitido
            Case "04": TBGravar!Coluna4 = Permitido
            Case "05": TBGravar!Coluna5 = Permitido
            Case "06": TBGravar!Coluna6 = Permitido
            Case "07": TBGravar!Coluna7 = Permitido
            Case "08": TBGravar!Coluna8 = Permitido
            Case "09": TBGravar!Coluna9 = Permitido
            Case "10": TBGravar!Coluna10 = Permitido
            Case "11": TBGravar!Coluna11 = Permitido
            Case "12": TBGravar!Coluna12 = Permitido
            Case "13": TBGravar!Coluna13 = Permitido
            Case "14": TBGravar!Coluna14 = Permitido
            Case "15": TBGravar!Coluna15 = Permitido
            Case "16": TBGravar!Coluna16 = Permitido
            Case "17": TBGravar!Coluna17 = Permitido
            Case "18": TBGravar!Coluna18 = Permitido
            Case "19": TBGravar!Coluna19 = Permitido
            Case "20": TBGravar!Coluna20 = Permitido
            Case "21": TBGravar!Coluna21 = Permitido
            Case "22": TBGravar!Coluna22 = Permitido
            Case "23": TBGravar!Coluna23 = Permitido
            Case "24": TBGravar!Coluna24 = Permitido
            Case "25": TBGravar!Coluna25 = Permitido
            Case "26": TBGravar!Coluna26 = Permitido
            Case "27": TBGravar!Coluna27 = Permitido
            Case "28": TBGravar!Coluna28 = Permitido
            Case "29": TBGravar!Coluna29 = Permitido
            Case "30": TBGravar!Coluna30 = Permitido
        End Select
        TBGravar.Update
        TBGravar.Close
    Next InitFor
End With

Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from Usuarios_monitor_trabalho where ID_usuario = " & pubIDUsuario & " and Modulo = 'PCP/Relatórios/Produtividade'", Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = False Then
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar"
    '==================================
    Modulo = "PCP/Relatórios/Produtividade"
    ID_documento = TBGravar!ID
    Documento = "Módulo: " & Modulo
    Documento1 = ""
    ProcGravaEvento
    frmProd_Relatorios_Produtividade.ProcVerifColunas
    '==================================
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub btnSalvar_Click()
On Error GoTo tratar_erro
  
ProcSalvar

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro
  
Select Case KeyCode
    Case vbKeyF3: ProcSalvar
    Case vbKeyEscape: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcLimpaVariaveisPrincipais
Modulo = "PCP/Relatórios/Produtividade"
ProcCarregarLista

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

