VERSION 5.00
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUsuarios_configurar_listas 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configurações do sistema - Usuários - Configurar listas"
   ClientHeight    =   8385
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7215
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
   ScaleHeight     =   8385
   ScaleWidth      =   7215
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   5520
      Top             =   120
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frmUsuarios_configurar_listas.frx":0000
      Count           =   1
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
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
      Left            =   55
      TabIndex        =   3
      Top             =   990
      Width           =   7125
      Begin VB.ComboBox Cmb_modulo 
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
         ItemData        =   "frmUsuarios_configurar_listas.frx":1E02
         Left            =   180
         List            =   "frmUsuarios_configurar_listas.frx":1E1B
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Módulo."
         Top             =   390
         Width           =   4485
      End
      Begin VB.ComboBox Cmb_marcar_desmarcar 
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
         ItemData        =   "frmUsuarios_configurar_listas.frx":1F45
         Left            =   4800
         List            =   "frmUsuarios_configurar_listas.frx":1F4F
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "Marcar/desmarcar todos."
         Top             =   390
         Width           =   2145
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Módulo"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2205
         TabIndex        =   4
         Top             =   180
         Width           =   510
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   5
      Top             =   0
      Width           =   7125
      _ExtentX        =   12568
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
   Begin MSComctlLib.ListView Lista 
      Height          =   6225
      Left            =   60
      TabIndex        =   2
      Top             =   1890
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
      Left            =   55
      TabIndex        =   6
      Top             =   8130
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
Attribute VB_Name = "frmUsuarios_configurar_listas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cmb_marcar_desmarcar_Click()
On Error GoTo tratar_erro

With Lista
    For InitFor = 1 To .ListItems.Count
        If Cmb_marcar_desmarcar = "Marcar todos" Then .ListItems.Item(InitFor).Checked = True Else .ListItems.Item(InitFor).Checked = False
    Next InitFor
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_modulo_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Select Case Cmb_modulo
    Case "PCP/Gerenciamento de ordem - Carteira de produção": Contador2 = 26
    Case "PCP/Programação da produção/Localizar ordem": Contador2 = 13
    Case "PCP/Monitor de trabalho": Contador2 = 20
    Case "PCP/Relatórios/Produtividade": Contador2 = 23
    Case "Estoque/Ordem de faturamento - Carteira de fat. - PI": Contador2 = 30
    Case "Estoque/Ordem de faturamento - Carteira de fat. - PC": Contador2 = 30
    Case "Outros/Validação de procedimentos": Contador2 = 30
End Select

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
    
    Select Case Cmb_modulo
        Case "PCP/Gerenciamento de ordem - Carteira de produção":
            Select Case Contador
                Case 3: TextoColunas = "Empresa"
                Case 4: TextoColunas = "Cód. int"
                Case 5: TextoColunas = "Rev."
                Case 6: TextoColunas = "Cód. ref."
                Case 7: TextoColunas = "Descrição"
                Case 8: TextoColunas = "Qt. vend"
                Case 9: TextoColunas = "Qt. fatur."
                Case 10: TextoColunas = "Qt. fatura"
                Case 11: TextoColunas = "Emp. est."
                Case 12: TextoColunas = "Emp. prod"
                Case 13: TextoColunas = "Necessidade"
                Case 14: TextoColunas = "Pr. final"
                Case 15: TextoColunas = "Proc."
                Case 16: TextoColunas = "Plano"
                Case 17: TextoColunas = "Versão estrut."
                Case 18: TextoColunas = "Versão proc."
                Case 19: TextoColunas = "Dt. venda"
                Case 20: TextoColunas = "Ped. int."
                Case 21: TextoColunas = "Rev."
                Case 22: TextoColunas = "Cliente"
                Case 23: TextoColunas = "Ped. cliente"
                Case 24: TextoColunas = "Qt. expda."
                Case 25: TextoColunas = "Qt. expedir"
                Case 26: TextoColunas = "Observação"
            End Select
        Case "PCP/Programação da produção/Localizar ordem":
            Select Case Contador
                Case 1: TextoColunas = "Ordem"
                Case 2: TextoColunas = "Tipo"
                Case 3: TextoColunas = "Qtde."
                Case 4: TextoColunas = "Qtde. faturada"
                Case 5: TextoColunas = "Saldo"
                Case 6: TextoColunas = "Dt. emissão"
                Case 7: TextoColunas = "Prazo final"
                Case 8: TextoColunas = "Cód. interno"
                Case 9: TextoColunas = "Cód. de ref."
                Case 10: TextoColunas = "Descrição"
                Case 11: TextoColunas = "Cliente"
                Case 12: TextoColunas = "Status"
                Case 13: TextoColunas = "Observações"
            End Select
        Case "PCP/Monitor de trabalho":
            Select Case Contador
                Case 1: TextoColunas = "Situação"
                Case 2: TextoColunas = "Posto de trabalho"
                Case 3: TextoColunas = "Descrição"
                Case 4: TextoColunas = "Data"
                Case 5: TextoColunas = "Evento"
                Case 6: TextoColunas = "Início"
                Case 7: TextoColunas = "Cód. interno"
                Case 8: TextoColunas = "Cód. de ref."
                Case 9: TextoColunas = "Descrição"
                Case 10: TextoColunas = "Prazo final"
                Case 11: TextoColunas = "Ordem"
                Case 12: TextoColunas = "OS"
                Case 13: TextoColunas = "Fase"
                Case 14: TextoColunas = "Grupo/op."
                Case 15: TextoColunas = "Tempo total prev."
                Case 16: TextoColunas = "Operador"
                Case 17: TextoColunas = "Cliente"
                Case 18: TextoColunas = "Qtde. OK"
                Case 19: TextoColunas = "Qtde. NC"
                Case 20: TextoColunas = "Eficiência"
            End Select
        Case "PCP/Relatórios/Produtividade":
            Select Case Contador
                Case 1: TextoColunas = "ID"
                Case 2: TextoColunas = "Data"
                Case 3: TextoColunas = "Texto"
                Case 4: TextoColunas = "Ordem"
                Case 5: TextoColunas = "Cód. interno"
                Case 6: TextoColunas = "Cód. de ref."
                Case 7: TextoColunas = "Descrição"
                Case 8: TextoColunas = "OS"
                Case 9: TextoColunas = "Fase"
                Case 10: TextoColunas = "Grupo/op."
                Case 11: TextoColunas = "Posto de trabalho"
                Case 12: TextoColunas = "Descrição"
                Case 13: TextoColunas = "Prep. prev."
                Case 14: TextoColunas = "Exec. prev."
                Case 15: TextoColunas = "Hs. previstas"
                Case 16: TextoColunas = "Qtde. prev."
                Case 17: TextoColunas = "Hs. utilizadas"
                Case 18: TextoColunas = "Qtde. apont."
                Case 19: TextoColunas = "Qtde. OK"
                Case 20: TextoColunas = "Qtde. NC"
                Case 21: TextoColunas = "Qtde. acumul. OS"
                Case 22: TextoColunas = "Turno"
                Case 23: TextoColunas = "Eficiência"
            End Select
        Case "Estoque/Ordem de faturamento - Carteira de fat. - PI":
            Select Case Contador
                Case 3: TextoColunas = "Empresa"
                Case 4: TextoColunas = "Cód. int."
                Case 5: TextoColunas = "Rev."
                Case 6: TextoColunas = "Cód. ref."
                Case 7: TextoColunas = "Descrição"
                Case 8: TextoColunas = "Un. com."
                Case 9: TextoColunas = "Ped. cliente"
                Case 10: TextoColunas = "N. item"
                Case 11: TextoColunas = "Pr. final"
                Case 12: TextoColunas = "Vlr. total"
                Case 13: TextoColunas = "Ped. int."
                Case 14: TextoColunas = "Rev."
                Case 15: TextoColunas = "Programa"
                Case 16: TextoColunas = "Rev."
                Case 17: TextoColunas = "ID"
                Case 18: TextoColunas = "Cliente"
                Case 19: TextoColunas = "UF"
                Case 20: TextoColunas = "Antec."
                Case 21: TextoColunas = "Parc."
                Case 22: TextoColunas = "Moeda"
                Case 23: TextoColunas = "Transportadora"
                Case 24: TextoColunas = "Qtde. vend."
                Case 25: TextoColunas = "Qtde. faturar"
                Case 26: TextoColunas = "Qtde. faturada"
                Case 27: TextoColunas = "Saldo"
                Case 28: TextoColunas = "Emp. est"
                Case 29: TextoColunas = "Emp. prod."
                Case 30: TextoColunas = "Observações"
            End Select
        Case "Estoque/Ordem de faturamento - Carteira de fat. - PC":
            Select Case Contador
                Case 3: TextoColunas = "Empresa"
                Case 4: TextoColunas = "Cód. int."
                Case 6: TextoColunas = "Cód. ref."
                Case 7: TextoColunas = "Descrição"
                Case 8: TextoColunas = "Un. com."
                Case 11: TextoColunas = "Pr. final"
                Case 12: TextoColunas = "Vlr. total"
                Case 13: TextoColunas = "Ped. compra"
                Case 17: TextoColunas = "ID"
                Case 18: TextoColunas = "Fornecedor"
                Case 19: TextoColunas = "UF"
                Case 22: TextoColunas = "Moeda"
                Case 23: TextoColunas = "Transportadora"
                Case 24: TextoColunas = "Qtde. comp."
                Case 26: TextoColunas = "Qtde. faturada"
                Case 27: TextoColunas = "Saldo"
                Case 30: TextoColunas = "Observações"
            End Select
        Case "Outros/Validação de procedimentos":
            Select Case Contador
                Case 2: TextoColunas = "Pedido int."
                Case 3: TextoColunas = "Rev."
                Case 4: TextoColunas = "Cod. interno"
                Case 5: TextoColunas = "Rev."
                Case 6: TextoColunas = "Descrição"
                Case 7: TextoColunas = "Qtde."
                Case 8: TextoColunas = "Dimensões"
                Case 9: TextoColunas = "Observações do produto"
                Case 10: TextoColunas = "Observações da venda"
                Case 11: TextoColunas = "Inspeção"
                Case 12: TextoColunas = "Embalagem"
                Case 13: TextoColunas = "Gravação"
                Case 14: TextoColunas = "Novo projeto"
                Case 15: TextoColunas = "Prioridade"
                Case 16: TextoColunas = "Dt. venda"
                Case 17: TextoColunas = "Prazo final"
                Case 18: TextoColunas = "Validação do pedido interno"
                Case 19: TextoColunas = "Validação do produto"
                Case 20: TextoColunas = "Validação da estrutura"
                Case 21: TextoColunas = "Validação do processo"
                Case 22: TextoColunas = "Validação do plano de inspeção"
                Case 23: TextoColunas = "Validação da compra"
                Case 24: TextoColunas = "Validação da ordem"
                Case 25: TextoColunas = "Data da inspeção final"
                Case 26: TextoColunas = "Data da entrada no estoque"
                Case 27: TextoColunas = "Validação da nota fiscal"
                Case 28: TextoColunas = "Data da expedição"
                Case 29: TextoColunas = "Versão estrutura"
                Case 30: TextoColunas = "Versão processo"
            End Select
    End Select
    
    If TextoColunas <> "" Then
        With Lista.ListItems
            .Add , , Contador
            .Item(.Count).SubItems(1) = TextoColunas
            .Item(.Count).SubItems(2) = NColuna
            
            TextoFiltro = "Coluna" & Contador & " = 'True'"
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select ID from Usuarios_monitor_trabalho where ID_usuario = " & frmUsuarios.txtId & " and Modulo = '" & Cmb_modulo & "' and " & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                .Item(ContadorReg).Checked = True
            End If
            TBAbrir.Close
        End With
        ContadorReg = ContadorReg + 1
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
If Cmb_modulo = "" Then
    NomeCampo = "o módulo"
    ProcVerificaAcao
    Cmb_modulo.SetFocus
    Exit Sub
End If
With Lista
    For InitFor = 1 To .ListItems.Count
        Set TBGravar = CreateObject("adodb.recordset")
        TBGravar.Open "Select * from Usuarios_monitor_trabalho where ID_usuario = " & frmUsuarios.txtId & " and Modulo = '" & Cmb_modulo & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBGravar.EOF = True Then TBGravar.AddNew
        TBGravar!ID_usuario = frmUsuarios.txtId
        TBGravar!Modulo = Cmb_modulo
        If .ListItems.Item(InitFor).Checked = True Then Permitido = True Else Permitido = False
        
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
TBGravar.Open "Select * from Usuarios_monitor_trabalho where ID_usuario = " & frmUsuarios.txtId & " and Modulo = '" & Cmb_modulo & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = False Then
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar"
    '==================================
    Modulo = "Configuração do sistema/Usuários/Configurar listas"
    ID_documento = TBGravar!ID
    Documento = "Módulo: " & Cmb_modulo
    Documento1 = ""
    ProcGravaEvento
    '==================================
End If

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

ProcCarregaToolBar1 Me, 7125, 5, True
ProcLimpaVariaveisPrincipais

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
