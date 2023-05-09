VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmRH_funcionarios_localizar 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "RH - Cadastro de funcionários - Localizar"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8925
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   8925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   6840
      Top             =   180
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frmRH_funcionarios_localizar.frx":0000
      Count           =   1
   End
   Begin VB.ComboBox Cmb_empresa 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   330
      ItemData        =   "frmRH_funcionarios_localizar.frx":21F0
      Left            =   240
      List            =   "frmRH_funcionarios_localizar.frx":21F2
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Empresa."
      Top             =   1260
      Width           =   8445
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   1515
      Left            =   55
      TabIndex        =   10
      Top             =   1590
      Width           =   8805
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0E0&
         Height          =   510
         Left            =   3810
         TabIndex        =   15
         Top             =   210
         Width           =   4785
         Begin VB.OptionButton Optfim 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Fim frase"
            Height          =   255
            Left            =   2760
            TabIndex        =   8
            Top             =   180
            Width           =   1155
         End
         Begin VB.OptionButton Optinicio 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Início frase"
            Height          =   255
            Left            =   180
            TabIndex        =   6
            Top             =   180
            Value           =   -1  'True
            Width           =   1275
         End
         Begin VB.OptionButton Optmeio 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Meio frase"
            Height          =   255
            Left            =   1470
            TabIndex        =   7
            Top             =   180
            Width           =   1275
         End
         Begin VB.OptionButton optIgual 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Igual"
            Height          =   255
            Left            =   3930
            TabIndex        =   9
            Top             =   180
            Width           =   705
         End
      End
      Begin VB.ComboBox cmbfiltrarpor 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         ItemData        =   "frmRH_funcionarios_localizar.frx":21F4
         Left            =   180
         List            =   "frmRH_funcionarios_localizar.frx":2219
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "Opções para filtro."
         Top             =   390
         Width           =   3585
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
         TabIndex        =   2
         ToolTipText     =   "Texto para pesquisa."
         Top             =   1050
         Width           =   8415
      End
      Begin VB.ComboBox cmbAniversariantes 
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
         ItemData        =   "frmRH_funcionarios_localizar.frx":227C
         Left            =   180
         List            =   "frmRH_funcionarios_localizar.frx":22A4
         Style           =   2  'Dropdown List
         TabIndex        =   5
         ToolTipText     =   "Texto para pesquisa."
         Top             =   1050
         Visible         =   0   'False
         Width           =   8415
      End
      Begin VB.ComboBox cmbTexto 
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
         ItemData        =   "frmRH_funcionarios_localizar.frx":230D
         Left            =   180
         List            =   "frmRH_funcionarios_localizar.frx":230F
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "Texto para pesquisa."
         Top             =   1050
         Visible         =   0   'False
         Width           =   8415
      End
      Begin MSMask.MaskEdBox txtCpf 
         Height          =   315
         Left            =   180
         TabIndex        =   4
         ToolTipText     =   "Texto para pesquisa."
         Top             =   1050
         Visible         =   0   'False
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         AllowPrompt     =   -1  'True
         AutoTab         =   -1  'True
         MaxLength       =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "###.###.###-##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         TabIndex        =   12
         Top             =   840
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
         Left            =   1552
         TabIndex        =   11
         Top             =   180
         Width           =   840
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   14
      Top             =   0
      Width           =   8805
      _ExtentX        =   15531
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
      ButtonCaption1  =   "Filtrar"
      ButtonEnabled1  =   0   'False
      ButtonIconSize1 =   32
      ButtonToolTipText1=   "Filtrar (F2)"
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
      ButtonWidth1    =   42
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
      ButtonLeft2     =   46
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
      ButtonLeft3     =   50
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
      ButtonLeft4     =   93
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
      ButtonLeft5     =   125
      ButtonTop5      =   2
      ButtonWidth5    =   24
      ButtonHeight5   =   24
   End
   Begin VB.Label Label44 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Empresa"
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
      Left            =   4095
      TabIndex        =   13
      Top             =   1050
      Width           =   735
   End
End
Attribute VB_Name = "frmRH_funcionarios_localizar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbfiltrarpor_Click()
On Error GoTo tratar_erro

cmbTexto.Clear
cmbAniversariantes.Clear
If cmbfiltrarpor = "Divisão" Then
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select divisao from funcionarios_divisao order by divisao", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Do While TBAbrir.EOF = False
            If IsNull(TBAbrir!divisao) = False Then
                cmbTexto.AddItem TBAbrir!divisao
            End If
            TBAbrir.MoveNext
        Loop
    End If
    TBAbrir.Close
ElseIf cmbfiltrarpor = "Setor" Then
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select setor from Usuarios_Setor order by setor", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Do While TBAbrir.EOF = False
                cmbTexto.AddItem Trim(TBAbrir!Setor)
                TBAbrir.MoveNext
            Loop
        End If
        TBAbrir.Close
    ElseIf cmbfiltrarpor = "Sexo" Then
            cmbTexto.AddItem "Feminino"
            cmbTexto.AddItem "Masculino"
        ElseIf cmbfiltrarpor = "Tipo" Then
                cmbTexto.AddItem "Horista"
                cmbTexto.AddItem "Mensalista"
            ElseIf cmbfiltrarpor = "Aniversariantes" Then
                    cmbAniversariantes.AddItem "Janeiro"
                    cmbAniversariantes.AddItem "Fevereiro"
                    cmbAniversariantes.AddItem "Março"
                    cmbAniversariantes.AddItem "Abril"
                    cmbAniversariantes.AddItem "Maio"
                    cmbAniversariantes.AddItem "Junho"
                    cmbAniversariantes.AddItem "Julho"
                    cmbAniversariantes.AddItem "Agosto"
                    cmbAniversariantes.AddItem "Setembro"
                    cmbAniversariantes.AddItem "Outubro"
                    cmbAniversariantes.AddItem "Novembro"
                    cmbAniversariantes.AddItem "Dezembro"
                    txtTexto.Visible = False
                    cmbTexto.Visible = False
                    txtCpf.Visible = False
                    cmbAniversariantes.Visible = True
                    Exit Sub
                ElseIf cmbfiltrarpor = "CPF" Then
                        txtTexto.Visible = False
                        cmbTexto.Visible = False
                        txtCpf.Visible = True
                        cmbAniversariantes.Visible = False
                        Exit Sub
                    ElseIf cmbfiltrarpor = "Situação" Then
                            cmbTexto.AddItem "Normal"
                            cmbTexto.AddItem "Afastado"
                            cmbTexto.AddItem "Demitido"
                            cmbTexto.AddItem "Temporário"
                        Else
                            txtCpf.Visible = False
                            txtTexto.Visible = True
                            cmbTexto.Visible = False
                            cmbAniversariantes.Visible = False
                            Exit Sub
End If
txtTexto.Visible = False
cmbTexto.Visible = True
txtCpf.Visible = False
cmbAniversariantes.Visible = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbTexto_Click()
On Error GoTo tratar_erro

If cmbTexto <> "" Then
    txtTexto = ""
    txtCpf = "___.___.___-__"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

With frmRH_Funcionarios
    Empresarel = Cmb_empresa
    .Aniversario = False
    If txtTexto.Visible = True And txtTexto <> "" Or cmbTexto.Visible = True And cmbTexto <> "" Or txtCpf.Visible = True And txtCpf <> "___.___.___-__" Or cmbAniversariantes.Visible = True And cmbAniversariantes <> "" Then
        If cmbfiltrarpor = "Aniversariantes" Then
            .Aniversario = True
            Select Case cmbAniversariantes
                Case "Janeiro": M = 1
                Case "Fevereiro": M = 2
                Case "Março": M = 3
                Case "Abril": M = 4
                Case "Maio": M = 5
                Case "Junho": M = 6
                Case "Julho": M = 7
                Case "Agosto": M = 8
                Case "Setembro": M = 9
                Case "Outubro": M = 10
                Case "Novembro": M = 11
                Case "Dezembro": M = 12
            End Select
            .StrSql_Localizar_Funcionarios = "Select * from Funcionarios where month((data_nascimento)) = '" & M & "' and Situacao <> 'Demitido' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " order by nome"
            .FormulaRel_Funcionarios = "Month({Funcionarios.data_nascimento}) = " & M & " and {Funcionarios.Situacao} <> 'Demitido' and {Funcionarios.ID_empresa} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
        ElseIf cmbfiltrarpor = "Sexo" Or cmbfiltrarpor = "Divisão" Or cmbfiltrarpor = "Tipo" Or cmbfiltrarpor = "Setor" Or cmbfiltrarpor = "Situação" Then
                Select Case cmbfiltrarpor
                    Case "Sexo": TextoFiltro = "Sexo"
                    Case "Divisão": TextoFiltro = "divisao"
                    Case "Tipo": TextoFiltro = "tipo"
                    Case "Setor": TextoFiltro = "setor"
                    Case "Situação": TextoFiltro = "Situacao"
                End Select
                .StrSql_Localizar_Funcionarios = "Select * from Funcionarios where " & TextoFiltro & " = '" & cmbTexto.Text & "' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " order by nome"
                .FormulaRel_Funcionarios = "{Funcionarios." & TextoFiltro & "} like '" & cmbTexto.Text & "' and {Funcionarios.ID_empresa} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
            ElseIf cmbfiltrarpor = "CPF" Then
                    .StrSql_Localizar_Funcionarios = "Select * from Funcionarios where CPF = '" & txtCpf.Text & "' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " order by nome"
                    .FormulaRel_Funcionarios = "{Funcionarios.CPF} like '" & txtCpf.Text & "' and {Funcionarios.ID_empresa} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
                Else
                    Select Case cmbfiltrarpor
                        Case "Código": TextoFiltro = "Codigo"
                        Case "Nome": TextoFiltro = "Nome"
                        Case "RG": TextoFiltro = "RG"
                        Case "Centro de custo": TextoFiltro = "Centro"
                    End Select
                    If Optinicio.Value = True Then
                        .StrSql_Localizar_Funcionarios = "Select * from Funcionarios where " & TextoFiltro & " like '" & txtTexto.Text & "%' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " order by nome"
                        .FormulaRel_Funcionarios = "{Funcionarios." & TextoFiltro & "} like '" & txtTexto.Text & "*' and {Funcionarios.ID_empresa} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
                    End If
                    If Optmeio.Value = True Then
                        .StrSql_Localizar_Funcionarios = "Select * from Funcionarios where " & TextoFiltro & " like '%" & txtTexto.Text & "%' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " order by nome"
                        .FormulaRel_Funcionarios = "{Funcionarios." & TextoFiltro & "} like '*" & txtTexto.Text & "*' and {Funcionarios.ID_empresa} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
                    End If
                    If Optfim.Value = True Then
                        .StrSql_Localizar_Funcionarios = "Select * from Funcionarios where " & TextoFiltro & " like '%" & txtTexto.Text & "' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " order by nome"
                        .FormulaRel_Funcionarios = "{Funcionarios." & TextoFiltro & "} like '*" & txtTexto.Text & "' and {Funcionarios.ID_empresa} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
                    End If
                    If optIgual.Value = True Then
                        .StrSql_Localizar_Funcionarios = "Select * from Funcionarios where " & TextoFiltro & " = '" & txtTexto.Text & "' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " order by nome"
                        .FormulaRel_Funcionarios = "{Funcionarios." & TextoFiltro & "} = '" & txtTexto.Text & "' and {Funcionarios.ID_empresa} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
                    End If
        End If
    Else
        .StrSql_Localizar_Funcionarios = "Select * from Funcionarios where ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " order by nome"
        .FormulaRel_Funcionarios = "{Funcionarios.ID_empresa} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
    End If
    .ProcAtualizalista (1)
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
    Case vbKeyEscape: Unload Me
    Case vbKeyF2: ProcFiltrar
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 9015, 5, True
ProcCarregaComboEmpresa Cmb_empresa, False
cmbfiltrarpor = "Nome"

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

Private Sub txtCpf_Change()
On Error GoTo tratar_erro

If txtCpf <> "___.___.___-__" Then
    cmbTexto.Clear
    txtTexto = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtTexto_Change()
On Error GoTo tratar_erro

If txtTexto <> "" Then
    cmbTexto.Clear
    txtCpf = "___.___.___-__"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcFiltrar
    'Case 3: ProcAjuda
    Case 4: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
