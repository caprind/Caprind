VERSION 5.00
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Begin VB.Form frmOpcoesGeral_PC 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configurações do sistema - Opções gerais - Dados para criar contas a pagar"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11430
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   11430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   8880
      Top             =   240
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frmOpcoesGeral_PC.frx":0000
      Count           =   1
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   1455
      Left            =   55
      TabIndex        =   8
      Top             =   990
      Width           =   11325
      Begin VB.CommandButton cmdFornecedor 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   10800
         Picture         =   "frmOpcoesGeral_PC.frx":2AAF
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Localizar fornecedor."
         Top             =   390
         Width           =   315
      End
      Begin VB.TextBox txtIDforn 
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
         TabIndex        =   0
         ToolTipText     =   "Código."
         Top             =   390
         Width           =   765
      End
      Begin VB.TextBox txtFornecedor 
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
         Left            =   960
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Descrição."
         Top             =   390
         Width           =   9825
      End
      Begin VB.TextBox Txt_ID_PC 
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
         MouseIcon       =   "frmOpcoesGeral_PC.frx":2BB1
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Text            =   "0"
         ToolTipText     =   "ID PC."
         Top             =   990
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.TextBox txtDia_venc 
         Alignment       =   2  'Center
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
         Left            =   10350
         MaxLength       =   255
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Código."
         Top             =   990
         Width           =   765
      End
      Begin VB.TextBox Txt_codigo_PC 
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
         MaxLength       =   255
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Código."
         Top             =   990
         Width           =   1875
      End
      Begin VB.CommandButton Cmd_localizar_PC 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   9930
         Picture         =   "frmOpcoesGeral_PC.frx":2EBB
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Localizar conta contábil."
         Top             =   990
         Width           =   315
      End
      Begin VB.TextBox Txt_descricao_PC 
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
         Left            =   2070
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Descrição."
         Top             =   990
         Width           =   7845
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Fornecedor"
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
         Index           =   1
         Left            =   5460
         TabIndex        =   13
         Top             =   180
         Width           =   825
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Dia vcto."
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
         Index           =   0
         Left            =   10410
         TabIndex        =   12
         Top             =   780
         Width           =   645
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
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
         Index           =   5
         Left            =   855
         TabIndex        =   11
         Top             =   780
         Width           =   510
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição"
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
         Index           =   4
         Left            =   5370
         TabIndex        =   10
         Top             =   780
         Width           =   720
         WordWrap        =   -1  'True
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   7
      Top             =   0
      Width           =   11325
      _ExtentX        =   19976
      _ExtentY        =   1720
      ButtonCount     =   6
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
      ButtonCaption2  =   "Excluir"
      ButtonEnabled2  =   0   'False
      ButtonIconSize2 =   32
      ButtonToolTipText2=   "Excluir (F4)"
      ButtonKey2      =   "2"
      ButtonAlignment2=   2
      BeginProperty ButtonFont2 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft2     =   48
      ButtonTop2      =   2
      ButtonWidth2    =   45
      ButtonHeight2   =   21
      ButtonUseMaskColor2=   0   'False
      ButtonEnabled3  =   0   'False
      ButtonIconSize3 =   32
      ButtonAlignment3=   2
      ButtonType3     =   1
      ButtonStyle3    =   -1
      BeginProperty ButtonFont3 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState3    =   -1
      ButtonLeft3     =   95
      ButtonTop3      =   4
      ButtonWidth3    =   2
      ButtonHeight3   =   54
      ButtonCaption4  =   "Ajuda"
      ButtonEnabled4  =   0   'False
      ButtonIconSize4 =   32
      ButtonToolTipText4=   "Ajuda (F1)"
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
      ButtonLeft4     =   99
      ButtonTop4      =   2
      ButtonWidth4    =   41
      ButtonHeight4   =   21
      ButtonUseMaskColor4=   0   'False
      ButtonCaption5  =   "Sair"
      ButtonEnabled5  =   0   'False
      ButtonIconSize5 =   32
      ButtonToolTipText5=   "Sair (Esc)"
      ButtonKey5      =   "5"
      ButtonAlignment5=   2
      BeginProperty ButtonFont5 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft5     =   142
      ButtonTop5      =   2
      ButtonWidth5    =   30
      ButtonHeight5   =   21
      ButtonUseMaskColor5=   0   'False
      ButtonEnabled6  =   0   'False
      ButtonIconSize6 =   32
      ButtonKey6      =   "6"
      ButtonAlignment6=   2
      BeginProperty ButtonFont6 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState6    =   5
      ButtonLeft6     =   174
      ButtonTop6      =   2
      ButtonWidth6    =   24
      ButtonHeight6   =   24
   End
End
Attribute VB_Name = "frmOpcoesGeral_PC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Novo_geral_pc As Boolean 'OK

Private Sub Cmd_localizar_PC_Click()
On Error GoTo tratar_erro
  
Plano_contas_produtos = False
Plano_contas_familias = False
Plano_centro_de_custo = False
Plano_instituicao = False
Plano_opcoesgerais = True
Plano_Faturamento = False
Plano_PCP = False
Financeiro_Contas_Pagar = False
Financeiro_Contas_Pagas = False
Financeiro_Contas_Receber = False
Financeiro_Contas_Recebidas = False
frmproj_produto_PC.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdFornecedor_Click()
On Error GoTo tratar_erro

ProcConfVariaveisLocForn False, False, False, False, False, False, False, False, False, False, False, False, False, True, False, False, False, False, False, False
FrmCompras_localizafornecedor.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro
  
Select Case KeyCode
    Case vbKeyF3: ProcSalvar
    Case vbKeyF4: ProcExcluir
    'Case vbKeyF1: imgAjuda_Click
    Case vbKeyEscape: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 11325, 6, True
Formulario = "Configuração do sistema/Opções gerais"
Direitos
ProcLimpaVariaveisPrincipais
ProcCarregaDados

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvar()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

Acao = "salvar"
If txtIDforn = "" Then
    NomeCampo = "o fornecedor"
    ProcVerificaAcao
    cmdFornecedor_Click
    Exit Sub
End If
If Txt_codigo_PC = "" Then
    NomeCampo = "a conta contábil"
    ProcVerificaAcao
    Cmd_localizar_PC_Click
    Exit Sub
End If
If txtDia_venc.Text = "" Then
    NomeCampo = "o dia de vencimento"
    ProcVerificaAcao
    txtDia_venc.SetFocus
    Exit Sub
Else
    If txtDia_venc > 31 Then
        USMsgBox ("O dia de vencimento não pode ser maior que 31."), vbExclamation, "CAPRIND v5.0"
        txtDia_venc.SetFocus
        Exit Sub
    End If
End If

With frmOpcoesGeral
    If .PC_PIS = True Then
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from Impostos where id = " & .txtID_imposto & " and (ID_PC_Cofins = " & Txt_ID_PC & " or ID_PC_CSLL = " & Txt_ID_PC & " or ID_PC_ISSQN = " & Txt_ID_PC & " or ID_PC_IRRF = " & Txt_ID_PC & " or ID_PC_INSS = " & Txt_ID_PC & ") ", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            USMsgBox "Não é permitido utilizar esta conta contábil, pois a mesma está sendo utilizada para outro imposto.", vbExclamation, "CAPRIND v5.0"
            Exit Sub
        End If
        TBAbrir.Close
    ElseIf .PC_Cofins = True Then
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select * from Impostos where id = " & .txtID_imposto & " and (ID_PC_PIS = " & Txt_ID_PC & " or ID_PC_CSLL = " & Txt_ID_PC & " or ID_PC_ISSQN = " & Txt_ID_PC & " or ID_PC_IRRF = " & Txt_ID_PC & " or ID_PC_INSS = " & Txt_ID_PC & ") ", Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                USMsgBox "Não é permitido utilizar esta conta contábil, pois a mesma está sendo utilizada para outro imposto.", vbExclamation, "CAPRIND v5.0"
                Exit Sub
            End If
            TBAbrir.Close
        ElseIf .PC_CSLL = True Then
                Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open "Select * from Impostos where id = " & .txtID_imposto & " and (ID_PC_Cofins = " & Txt_ID_PC & " or ID_PC_PIS = " & Txt_ID_PC & " or ID_PC_ISSQN = " & Txt_ID_PC & " or ID_PC_IRRF = " & Txt_ID_PC & " or ID_PC_INSS = " & Txt_ID_PC & ") ", Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    USMsgBox "Não é permitido utilizar esta conta contábil, pois a mesma está sendo utilizada para outro imposto.", vbExclamation, "CAPRIND v5.0"
                    Exit Sub
                End If
                TBAbrir.Close
            ElseIf .PC_ISSQN = True Then
                    Set TBAbrir = CreateObject("adodb.recordset")
                    TBAbrir.Open "Select * from Impostos where id = " & .txtID_imposto & " and (ID_PC_Cofins = " & Txt_ID_PC & " or ID_PC_CSLL = " & Txt_ID_PC & " or ID_PC_PIS = " & Txt_ID_PC & " or ID_PC_IRRF = " & Txt_ID_PC & " or ID_PC_INSS = " & Txt_ID_PC & ") ", Conexao, adOpenKeyset, adLockOptimistic
                    If TBAbrir.EOF = False Then
                        USMsgBox "Não é permitido utilizar esta conta contábil, pois a mesma está sendo utilizada para outro imposto.", vbExclamation, "CAPRIND v5.0"
                        Exit Sub
                    End If
                    TBAbrir.Close
                ElseIf .PC_IRRF = True Then
                        Set TBAbrir = CreateObject("adodb.recordset")
                        TBAbrir.Open "Select * from Impostos where id = " & .txtID_imposto & " and (ID_PC_Cofins = " & Txt_ID_PC & " or ID_PC_CSLL = " & Txt_ID_PC & " or ID_PC_ISSQN = " & Txt_ID_PC & " or ID_PC_PIS = " & Txt_ID_PC & " or ID_PC_INSS = " & Txt_ID_PC & ") ", Conexao, adOpenKeyset, adLockOptimistic
                        If TBAbrir.EOF = False Then
                            USMsgBox "Não é permitido utilizar esta conta contábil, pois a mesma está sendo utilizada para outro imposto.", vbExclamation, "CAPRIND v5.0"
                            Exit Sub
                        End If
                        TBAbrir.Close
                    ElseIf .PC_INSS = True Then
                        Set TBAbrir = CreateObject("adodb.recordset")
                        TBAbrir.Open "Select * from Impostos where id = " & .txtID_imposto & " and (ID_PC_Cofins = " & Txt_ID_PC & " or ID_PC_CSLL = " & Txt_ID_PC & " or ID_PC_ISSQN = " & Txt_ID_PC & " or ID_PC_IRRF = " & Txt_ID_PC & " or ID_PC_PIS = " & Txt_ID_PC & ") ", Conexao, adOpenKeyset, adLockOptimistic
                        If TBAbrir.EOF = False Then
                            USMsgBox "Não é permitido utilizar esta conta contábil, pois a mesma está sendo utilizada para outro imposto.", vbExclamation, "CAPRIND v5.0"
                            Exit Sub
                        End If
                        TBAbrir.Close
    End If
    
    Set TBGravar = CreateObject("adodb.recordset")
    TBGravar.Open "Select * from Impostos where id = " & .txtID_imposto & " and Regime = " & .Regime, Conexao, adOpenKeyset, adLockOptimistic
    If TBGravar.EOF = False Then
        
        If .PC_PIS = True Then
            TBGravar!ID_PC_PIS = Txt_ID_PC
            TBGravar!Dia_PIS = txtDia_venc
            TBGravar!IDForn_PIS = txtIDforn
        ElseIf .PC_Cofins = True Then
                TBGravar!ID_PC_Cofins = Txt_ID_PC
                TBGravar!Dia_Cofins = txtDia_venc
                TBGravar!IDForn_Cofins = txtIDforn
            ElseIf .PC_CSLL = True Then
                    TBGravar!ID_PC_CSLL = Txt_ID_PC
                    TBGravar!Dia_CSLL = txtDia_venc
                    TBGravar!IDForn_CSLL = txtIDforn
                ElseIf .PC_ISSQN = True Then
                        TBGravar!ID_PC_ISSQN = Txt_ID_PC
                        TBGravar!Dia_ISSQN = txtDia_venc
                        TBGravar!IDForn_ISSQN = txtIDforn
                    ElseIf .PC_IRRF = True Then
                            TBGravar!ID_PC_IRRF = Txt_ID_PC
                            TBGravar!Dia_IRRF = txtDia_venc
                            TBGravar!IDForn_IRRF = txtIDforn
                        ElseIf .PC_INSS = True Then
                            TBGravar!ID_PC_INSS = Txt_ID_PC
                            TBGravar!Dia_INSS = txtDia_venc
                            TBGravar!IDForn_INSS = txtIDforn
        End If
        TBGravar.Update
    End If
    TBGravar.Close
    
    '==================================
    Modulo = "Configuração do sistema/Opções gerais"
    Evento = "Salvar"
    ID_documento = .txtID_imposto
    Documento = "Imposto: " & .txtID_imposto
    Documento1 = "Código do plano: " & Txt_codigo_PC & " - Descrição do plano: " & Txt_descricao_PC
    ProcGravaEvento
    '==================================
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluir()
On Error GoTo tratar_erro
  
If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

If USMsgBox("Deseja realmente excluir este registro?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    With frmOpcoesGeral
        Set TBGravar = CreateObject("adodb.recordset")
        TBGravar.Open "Select * from Impostos where id = " & .txtID_imposto & " and Regime = " & .Regime, Conexao, adOpenKeyset, adLockOptimistic
        If TBGravar.EOF = False Then
            If .PC_PIS = True Then
                TBGravar!ID_PC_PIS = Null
                TBGravar!Dia_PIS = Null
                TBGravar!IDForn_PIS = Null
            ElseIf .PC_Cofins = True Then
                    TBGravar!ID_PC_Cofins = Null
                    TBGravar!Dia_Cofins = Null
                    TBGravar!IDForn_Cofins = Null
                ElseIf .PC_CSLL = True Then
                        TBGravar!ID_PC_CSLL = Null
                        TBGravar!Dia_CSLL = Null
                        TBGravar!IDForn_CSLL = Null
                    ElseIf .PC_ISSQN = True Then
                            TBGravar!ID_PC_ISSQN = Null
                            TBGravar!Dia_ISSQN = Null
                            TBGravar!IDForn_ISSQN = Null
                        ElseIf .PC_IRRF = True Then
                                TBGravar!ID_PC_IRRF = Null
                                TBGravar!Dia_IRRF = Null
                                TBGravar!IDForn_IRRF = Null
                            ElseIf .PC_INSS = True Then
                                TBGravar!ID_PC_INSS = Null
                                TBGravar!Dia_INSS = Null
                                TBGravar!IDForn_INSS = Null
            End If
            TBGravar.Update
        End If
        TBGravar.Close
        USMsgBox ("Registros excluídos com sucesso."), vbInformation, "CAPRIND v5.0"
        
        '==================================
        Modulo = "Configuração do sistema/Opções gerais"
        Evento = "Excluir"
        ID_documento = .txtID_imposto
        Documento = "Imposto: " & .txtID_imposto
        Documento1 = ""
        ProcGravaEvento
        '==================================
        ProcLimpaCampos
    End With
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtDia_venc_Change()
On Error GoTo tratar_erro

If txtDia_venc.Text <> "" Then
    VerifNumero = txtDia_venc.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtDia_venc.Text = ""
        txtDia_venc.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtIDforn_Change()
On Error GoTo tratar_erro

txtFornecedor = ""
If txtIDforn <> "" Then
    VerifNumero = txtIDforn
    ProcVerificaNumero
    If VerifNumero = False Then
        txtIDforn = ""
        txtIDforn.SetFocus
        Exit Sub
    End If
    Set TBFornecedor = CreateObject("adodb.recordset")
    TBFornecedor.Open "Select Nome_Razao from compras_fornecedores where idcliente = " & txtIDforn, Conexao, adOpenKeyset, adLockOptimistic
    If TBFornecedor.EOF = False Then txtFornecedor = TBFornecedor!Nome_Razao
    TBFornecedor.Close
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
    Case 2: ProcExcluir
    'Case 4: ProcAjuda
    Case 5: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCampos()
On Error GoTo tratar_erro

txtIDforn = ""
txtFornecedor = ""
Txt_ID_PC = ""
Txt_codigo_PC = ""
Txt_descricao_PC = ""
txtDia_venc = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaDados()
On Error GoTo tratar_erro

With frmOpcoesGeral
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from Impostos where id = " & .txtID_imposto & " and Regime = " & .Regime, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        If .PC_PIS = True Then
            Txt_ID_PC = IIf(IsNull(TBAbrir!ID_PC_PIS), "", TBAbrir!ID_PC_PIS)
            txtDia_venc = IIf(IsNull(TBAbrir!Dia_PIS), "", TBAbrir!Dia_PIS)
            txtIDforn = IIf(IsNull(TBAbrir!IDForn_PIS), "", TBAbrir!IDForn_PIS)
        ElseIf .PC_Cofins = True Then
                Txt_ID_PC = IIf(IsNull(TBAbrir!ID_PC_Cofins), "", TBAbrir!ID_PC_Cofins)
                txtDia_venc = IIf(IsNull(TBAbrir!Dia_Cofins), "", TBAbrir!Dia_Cofins)
                txtIDforn = IIf(IsNull(TBAbrir!IDForn_Cofins), "", TBAbrir!IDForn_Cofins)
            ElseIf .PC_CSLL = True Then
                    Txt_ID_PC = IIf(IsNull(TBAbrir!ID_PC_CSLL), "", TBAbrir!ID_PC_CSLL)
                    txtDia_venc = IIf(IsNull(TBAbrir!Dia_CSLL), "", TBAbrir!Dia_CSLL)
                    txtIDforn = IIf(IsNull(TBAbrir!IDForn_CSLL), "", TBAbrir!IDForn_CSLL)
                ElseIf .PC_ISSQN = True Then
                        Txt_ID_PC = IIf(IsNull(TBAbrir!ID_PC_ISSQN), "", TBAbrir!ID_PC_ISSQN)
                        txtDia_venc = IIf(IsNull(TBAbrir!Dia_ISSQN), "", TBAbrir!Dia_ISSQN)
                        txtIDforn = IIf(IsNull(TBAbrir!IDForn_ISSQN), "", TBAbrir!IDForn_ISSQN)
                    ElseIf .PC_IRRF = True Then
                            Txt_ID_PC = IIf(IsNull(TBAbrir!ID_PC_IRRF), "", TBAbrir!ID_PC_IRRF)
                            txtDia_venc = IIf(IsNull(TBAbrir!Dia_IRRF), "", TBAbrir!Dia_IRRF)
                            txtIDforn = IIf(IsNull(TBAbrir!IDForn_IRRF), "", TBAbrir!IDForn_IRRF)
                        ElseIf .PC_INSS = True Then
                            Txt_ID_PC = IIf(IsNull(TBAbrir!ID_PC_INSS), "", TBAbrir!ID_PC_INSS)
                            txtDia_venc = IIf(IsNull(TBAbrir!Dia_INSS), "", TBAbrir!Dia_INSS)
                            txtIDforn = IIf(IsNull(TBAbrir!IDForn_INSS), "", TBAbrir!IDForn_INSS)
        End If
    End If
    TBAbrir.Close
End With

If Txt_ID_PC <> "" Then
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from tbl_familia where int_codfamilia = " & Txt_ID_PC, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Txt_codigo_PC = IIf(IsNull(TBAbrir!CODIGO), "", TBAbrir!CODIGO)
        Txt_descricao_PC = IIf(IsNull(TBAbrir!Txt_descricao), "", TBAbrir!Txt_descricao)
    End If
    TBAbrir.Close
End If

If txtIDforn <> "" Then
    Set TBFornecedor = CreateObject("adodb.recordset")
    TBFornecedor.Open "Select * from compras_fornecedores where idcliente = " & txtIDforn, Conexao, adOpenKeyset, adLockOptimistic
    If TBFornecedor.EOF = False Then txtFornecedor = IIf(IsNull(TBFornecedor!Nome_Razao), "", TBFornecedor!Nome_Razao)
    TBFornecedor.Close
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
