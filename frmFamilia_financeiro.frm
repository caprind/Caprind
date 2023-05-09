VERSION 5.00
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Begin VB.Form frmFamilia_financeiro 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   Caption         =   "Conta contábil"
   ClientHeight    =   10035
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   15360
   ControlBox      =   0   'False
   Icon            =   "frmFamilia_financeiro.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   10035
   ScaleWidth      =   15360
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   99
      ScreenHeight    =   1080
      ScreenWidth     =   2560
      ScreenHeightDT  =   1080
      ScreenWidthDT   =   1920
      AutoResizeOnLoad=   0   'False
      ApplicationName =   "Active Resize Control Professional"
      FormHeightDT    =   10500
      FormWidthDT     =   15480
      FormScaleHeightDT=   10035
      FormScaleWidthDT=   15360
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   825
      Left            =   55
      TabIndex        =   7
      Top             =   990
      Width           =   15225
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
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Descrição."
         Top             =   390
         Width           =   10515
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
         TabIndex        =   14
         Text            =   "0"
         ToolTipText     =   "ID PC."
         Top             =   390
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.CommandButton Cmd_localizar_PC 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   12600
         Picture         =   "frmFamilia_financeiro.frx":1042
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Localizar plano de contas."
         Top             =   390
         Width           =   315
      End
      Begin VB.TextBox txtValor 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Left            =   13350
         MaxLength       =   20
         TabIndex        =   4
         ToolTipText     =   "Valor."
         Top             =   390
         Width           =   1680
      End
      Begin VB.CommandButton Cmd_PC 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   12930
         Picture         =   "frmFamilia_financeiro.frx":1144
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Abrir formulário para cadastro de plano de contas."
         Top             =   390
         Width           =   315
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
         TabIndex        =   0
         TabStop         =   0   'False
         ToolTipText     =   "Código."
         Top             =   390
         Width           =   1875
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição*"
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
         Left            =   6937
         TabIndex        =   16
         Top             =   180
         Width           =   780
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Código*"
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
         Left            =   825
         TabIndex        =   15
         Top             =   180
         Width           =   585
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Valor*"
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
         Left            =   13920
         TabIndex        =   9
         Top             =   180
         Width           =   540
         WordWrap        =   -1  'True
      End
   End
   Begin VB.TextBox txtID 
      Height          =   285
      Left            =   6510
      TabIndex        =   8
      Text            =   "0"
      Top             =   5130
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.TextBox txtIDConta 
      Height          =   285
      Left            =   5820
      TabIndex        =   10
      Text            =   "0"
      Top             =   5130
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00E0E0E0&
      Height          =   585
      Left            =   60
      TabIndex        =   11
      Top             =   9420
      Width           =   15225
      Begin VB.TextBox txtSaldo 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Left            =   13440
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Saldo da conta."
         Top             =   180
         Width           =   1620
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo da conta :"
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
         Height          =   435
         Left            =   12000
         TabIndex        =   12
         Top             =   180
         Width           =   2190
         WordWrap        =   -1  'True
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   13
      Top             =   0
      Width           =   15225
      _ExtentX        =   26855
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft1     =   2
      ButtonTop1      =   2
      ButtonWidth1    =   36
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft2     =   40
      ButtonTop2      =   2
      ButtonWidth2    =   44
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft3     =   86
      ButtonTop3      =   2
      ButtonWidth3    =   45
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
      ButtonLeft4     =   133
      ButtonTop4      =   4
      ButtonWidth4    =   2
      ButtonHeight4   =   54
      ButtonCaption5  =   "Ajuda"
      ButtonEnabled5  =   0   'False
      ButtonIconSize5 =   32
      ButtonToolTipText5=   "Ajuda (F1)"
      ButtonKey5      =   "6"
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
      ButtonLeft5     =   137
      ButtonTop5      =   2
      ButtonWidth5    =   41
      ButtonHeight5   =   21
      ButtonUseMaskColor5=   0   'False
      ButtonCaption6  =   "Sair"
      ButtonEnabled6  =   0   'False
      ButtonIconSize6 =   32
      ButtonToolTipText6=   "Sair (Esc)"
      ButtonKey6      =   "7"
      ButtonAlignment6=   2
      BeginProperty ButtonFont6 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft6     =   180
      ButtonTop6      =   2
      ButtonWidth6    =   30
      ButtonHeight6   =   21
      ButtonUseMaskColor6=   0   'False
      ButtonEnabled7  =   0   'False
      ButtonIconSize7 =   32
      ButtonKey7      =   "8"
      ButtonAlignment7=   2
      BeginProperty ButtonFont7 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState7    =   5
      ButtonLeft7     =   212
      ButtonTop7      =   2
      ButtonWidth7    =   24
      ButtonHeight7   =   24
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   3840
         Top             =   180
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmFamilia_financeiro.frx":1226
         Count           =   1
      End
   End
   Begin DrawSuite2022.USTreeView USTreeView1 
      Height          =   7575
      Left            =   60
      TabIndex        =   5
      Top             =   1830
      Width           =   15225
      _ExtentX        =   26855
      _ExtentY        =   13361
      BorderColor     =   12500670
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Theme           =   1
   End
End
Attribute VB_Name = "frmFamilia_financeiro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Novo_familia_financeiro As Boolean 'OK

Private Sub Cmd_localizar_PC_Click()
On Error GoTo tratar_erro

Plano_contas_produtos = False
Plano_contas_familias = False
Plano_centro_de_custo = False
Plano_instituicao = False
Plano_opcoesgerais = False
Plano_Faturamento = False

'Não pode colocar essas variaveis neste modulo
'Financeiro_Contas_Pagar = False
'Financeiro_Contas_Pagas = False
'Financeiro_Contas_Receber = False
'Financeiro_Contas_Recebidas = False

Plano_PCP = False
frmproj_produto_PC.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_PC_Click()
On Error GoTo tratar_erro

frmFinanceiro_familia.Show

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
    'Case vbKeyF1: ProcAjuda
    Case vbKeyEscape: ProcSair
End Select

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
Acao = "salvar"
If Txt_ID_PC = 0 Then
    NomeCampo = "a conta contábil"
    ProcVerificaAcao
    Cmd_localizar_PC.SetFocus
    Exit Sub
End If

'Verifica se é devolução e altera o valor para negativo
valor = IIf(txtValor = "", 0, txtValor)
Permitido = False
If Financeiro_Contas_Pagar = True Or Financeiro_Contas_Pagas = True Then NomeTabela = "tbl_ContasPagar" Else NomeTabela = "tbl_contas_receber"
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from " & NomeTabela & " where IDIntconta = " & txtIDConta, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    IDnota2 = IIf(IsNull(TBAbrir!ID_nota), 0, TBAbrir!ID_nota)
    If TBAbrir!Devolucao = True And valor >= 0 Then
        txtValor = Format(txtValor, "-###,##0.00")
    ElseIf TBAbrir!Devolucao = False And valor <= 0 Then
            NomeCampo = "o valor"
            ProcVerificaAcao
            txtValor.SetFocus
            Exit Sub
    End If
End If

'Verifica se já existe o plano cadastrado
If Financeiro_Contas_Pagar = True Or Financeiro_Contas_Pagas = True Then Tipo = "TipoConta = 'P'" Else Tipo = "TipoConta = 'R'"
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from familia_financeiro where IDconta = " & txtIDConta & " and ID_PC = " & Txt_ID_PC & " and id <> " & txtId & " and " & Tipo, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    USMsgBox ("Esta conta contábil está sendo utilizada, favor alterar."), vbExclamation, "CAPRIND v5.0"
    TBAbrir.Close
    Exit Sub
End If
TBAbrir.Close

qt = Format(txtValor, "###,##0.00")
If qt < 0 Then qt = qt * -1
qtdeliberada = txtSaldo
If qtdeliberada < 0 Then qtdeliberada = qtdeliberada * -1

Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from familia_financeiro where id = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = False Then
    Qtd = Format(qtdeliberada + TBGravar!valor, "###,##0.00")
    If qt <> TBGravar!valor And Qtd < qt Then
        USMsgBox ("Não é permitido salvar pois o saldo da conta é menor que o valor da conta contábil."), vbExclamation, "CAPRIND v5.0"
        Exit Sub
    End If
Else
    If qtdeliberada < qt Then
        USMsgBox ("Não é permitido salvar pois o saldo da conta é menor que o valor da conta contábil."), vbExclamation, "CAPRIND v5.0"
        Exit Sub
    End If
    TBGravar.AddNew
    If Financeiro_Contas_Pagar = True Or Financeiro_Contas_Receber = True Then TBGravar!Pago_recebido = False Else TBGravar!Pago_recebido = True
End If
TBGravar!ID_PC = Txt_ID_PC
TBGravar!IDConta = txtIDConta
TBGravar!IDnota = IIf(IDnota2 = 0, Null, IDnota)
TBGravar!valor = txtValor
If Financeiro_Contas_Pagar = True Or Financeiro_Contas_Pagas = True Then TBGravar!TipoConta = "P" Else TBGravar!TipoConta = "R"
TBGravar.Update
txtId = TBGravar!ID
TBGravar.Close
If Novo_familia_financeiro = True Then
    USMsgBox ("Nova conta contábil cadastrada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar"
End If
ProcVerifPC
'==================================
Modulo = Formulario
ID_documento = txtId
Documento = "Documento: " & FamiliaAntiga
Documento1 = "Código do plano: " & Txt_codigo_PC & " - Descrição do plano: " & Txt_descricao_PC
ProcGravaEvento
'==================================
Novo_familia_financeiro = False

If Financeiro_Contas_Pagar = True Then
    frmContas_Pagar.ProcCarregaListaPC
ElseIf Financeiro_Contas_Pagas = True Then
        frmContas_Pagas.ProcCarregaListaPC
    ElseIf Financeiro_Contas_Receber = True Then
            frmContas_Receber.ProcCarregaListaPC
        Else
            frmContas_recebidas.ProcCarregaListaPC
End If
        
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15225, 7, True
If Financeiro_Contas_Pagar = True Then
    Caption = "Financeiro - Contas a pagar - Conta contábil"
    Formulario = "Financeiro/Contas a pagar/Conta contábil"
    With frmContas_Pagar
        txtIDConta = .txtidintconta
        FamiliaAntiga = .txtNDocumento
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select dbl_valorpagto from tbl_contaspagar where idintconta = " & txtIDConta, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Qtde = IIf(IsNull(TBAbrir!dbl_valorpagto), 0, TBAbrir!dbl_valorpagto)
        End If
        TBAbrir.Close
    End With
ElseIf Financeiro_Contas_Receber = True Then
        Caption = "Financeiro - Contas a receber - Conta contábil"
        Formulario = "Financeiro/Contas a receber/Conta contábil"
        With frmContas_Receber
            txtIDConta = .txtidintconta
            FamiliaAntiga = .txtDocumento
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select Valor from tbl_contas_receber where idintconta = " & txtIDConta, Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                Qtde = IIf(IsNull(TBAbrir!valor), 0, TBAbrir!valor)
            End If
            TBAbrir.Close
        End With
    ElseIf Financeiro_Contas_Pagas = True Then
            Caption = "Financeiro - Contas pagas - Conta contábil"
            Formulario = "Financeiro/Contas pagas/Conta contábil"
            With frmContas_Pagas
                txtIDConta = .txtidintconta
                FamiliaAntiga = .txtNDocumento
                Qtde = .txt_ValorPago
            End With
        Else
            Caption = "Financeiro - Contas recebidas - Conta contábil"
            Formulario = "Financeiro/Contas recebidas/Conta contábil"
            With frmContas_recebidas
                txtIDConta = .txtidintconta
                FamiliaAntiga = .txtDocumento
                Qtde = .txtvalortitrecebido
            End With
End If

Direitos
ProcLimpaVariaveisPrincipais
ProcVerifPC

ProcRemoveObjetosResize Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCampos()
On Error GoTo tratar_erro

txtId.Text = 0
Txt_ID_PC = 0
Txt_codigo_PC = ""
Txt_descricao_PC = ""
txtValor.Text = txtSaldo
CodigoLista = 0
Frame2.Enabled = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcVerifPC()
On Error GoTo tratar_erro

Conexao.Execute "DELETE from Plano_de_contas_totalizacao where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'"

If Financeiro_Contas_Pagar = True Then
    If frmContas_Pagar.Chk_antecipacao.Value = 1 Then TextoFiltro = "TipoConta = 'P'" Else TextoFiltro = "TipoConta = 'P' and Pago_recebido = 'False'"
ElseIf Financeiro_Contas_Receber = True Then
        If frmContas_Receber.Chk_antecipacao.Value = 1 Then TextoFiltro = "TipoConta = 'R'" Else TextoFiltro = "TipoConta = 'R' and Pago_recebido = 'False'"
    ElseIf Financeiro_Contas_Pagas = True Then
            TextoFiltro = "TipoConta = 'P' and Pago_recebido = 'True'"
        Else
            TextoFiltro = "TipoConta = 'R' and Pago_recebido = 'True'"
End If
Valor1 = 0
Set TBFamilia = CreateObject("adodb.recordset")
TBFamilia.Open "Select ID_PC, Sum(valor) as Valor from Familia_financeiro where idconta = " & txtIDConta & " and Deposito_transf = 'False' and " & TextoFiltro & " and ID_PC is not null and ID_PC <> 0 Group by ID_PC", Conexao, adOpenKeyset, adLockOptimistic
If TBFamilia.EOF = False Then
    Do While TBFamilia.EOF = False
        Valor1 = IIf(IsNull(TBFamilia!valor), 0, TBFamilia!valor)
        
        'Verifica o código e o nível do PC
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from tbl_familia where int_codfamilia = " & TBFamilia!ID_PC & " and Codigo is not null and nivel is not null", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Select Case TBAbrir!Nivel
                Case 8:
                    Set TBNivel8 = CreateObject("adodb.recordset")
                    TBNivel8.Open "Select * from tbl_familia where int_codfamilia = " & TBFamilia!ID_PC & " and Codigo is not null and nivel is not null", Conexao, adOpenKeyset, adLockOptimistic
                    If TBNivel8.EOF = False Then
                        
                        Set TBGravar = CreateObject("adodb.recordset")
                        TBGravar.Open "Select * from Plano_de_contas_totalizacao where Codigo = '" & TBNivel8!CODIGO & "' and Nivel = " & TBNivel8!Nivel & " and Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
                        ProcEnviaDadosPC TBNivel8!int_codfamilia, TBNivel8!CODIGO, TBNivel8!Txt_descricao, TBNivel8!Nivel
                        
                        ProcNivelPC7
                    End If
                    TBNivel8.Close
                Case 7:
                    Set TBNivel7 = CreateObject("adodb.recordset")
                    TBNivel7.Open "Select * from tbl_familia where int_codfamilia = " & TBFamilia!ID_PC & " and Codigo is not null and nivel is not null", Conexao, adOpenKeyset, adLockOptimistic
                    If TBNivel7.EOF = False Then
                        
                        Set TBGravar = CreateObject("adodb.recordset")
                        TBGravar.Open "Select * from Plano_de_contas_totalizacao where Codigo = '" & TBNivel7!CODIGO & "' and Nivel = " & TBNivel7!Nivel & " and Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
                        ProcEnviaDadosPC TBNivel7!int_codfamilia, TBNivel7!CODIGO, TBNivel7!Txt_descricao, TBNivel7!Nivel
                        
                        ProcNivelPC6
                    End If
                    TBNivel7.Close
                Case 6:
                    Set TBNivel6 = CreateObject("adodb.recordset")
                    TBNivel6.Open "Select * from tbl_familia where int_codfamilia = " & TBFamilia!ID_PC & " and Codigo is not null and nivel is not null", Conexao, adOpenKeyset, adLockOptimistic
                    If TBNivel6.EOF = False Then
                        
                        Set TBGravar = CreateObject("adodb.recordset")
                        TBGravar.Open "Select * from Plano_de_contas_totalizacao where Codigo = '" & TBNivel6!CODIGO & "' and Nivel = " & TBNivel6!Nivel & " and Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
                        ProcEnviaDadosPC TBNivel6!int_codfamilia, TBNivel6!CODIGO, TBNivel6!Txt_descricao, TBNivel6!Nivel
                        
                        ProcNivelPC5
                    End If
                    TBNivel6.Close
                Case 5:
                    Set TBNivel5 = CreateObject("adodb.recordset")
                    TBNivel5.Open "Select * from tbl_familia where int_codfamilia = " & TBFamilia!ID_PC & " and Codigo is not null and nivel is not null", Conexao, adOpenKeyset, adLockOptimistic
                    If TBNivel5.EOF = False Then
                        
                        Set TBGravar = CreateObject("adodb.recordset")
                        TBGravar.Open "Select * from Plano_de_contas_totalizacao where Codigo = '" & TBNivel5!CODIGO & "' and Nivel = " & TBNivel5!Nivel & " and Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
                        ProcEnviaDadosPC TBNivel5!int_codfamilia, TBNivel5!CODIGO, TBNivel5!Txt_descricao, TBNivel5!Nivel
                        
                        ProcNivelPC4
                    End If
                    TBNivel5.Close
                Case 4:
                    Set TBNivel4 = CreateObject("adodb.recordset")
                    TBNivel4.Open "Select * from tbl_familia where int_codfamilia = " & TBFamilia!ID_PC & " and Codigo is not null and nivel is not null", Conexao, adOpenKeyset, adLockOptimistic
                    If TBNivel4.EOF = False Then
                        
                        Set TBGravar = CreateObject("adodb.recordset")
                        TBGravar.Open "Select * from Plano_de_contas_totalizacao where Codigo = '" & TBNivel4!CODIGO & "' and Nivel = " & TBNivel4!Nivel & " and Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
                        ProcEnviaDadosPC TBNivel4!int_codfamilia, TBNivel4!CODIGO, TBNivel4!Txt_descricao, TBNivel4!Nivel
                        
                        ProcNivelPC3
                    End If
                    TBNivel4.Close
                Case 3:
                    Set TBNivel3 = CreateObject("adodb.recordset")
                    TBNivel3.Open "Select * from tbl_familia where int_codfamilia = " & TBFamilia!ID_PC & " and Codigo is not null and nivel is not null", Conexao, adOpenKeyset, adLockOptimistic
                    If TBNivel3.EOF = False Then
                        
                        Set TBGravar = CreateObject("adodb.recordset")
                        TBGravar.Open "Select * from Plano_de_contas_totalizacao where Codigo = '" & TBNivel3!CODIGO & "' and Nivel = " & TBNivel3!Nivel & " and Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
                        ProcEnviaDadosPC TBNivel3!int_codfamilia, TBNivel3!CODIGO, TBNivel3!Txt_descricao, TBNivel3!Nivel
                        
                        ProcNivelPC2
                    End If
                    TBNivel3.Close
                Case 2:
                    Set TBNivel2 = CreateObject("adodb.recordset")
                    TBNivel2.Open "Select * from tbl_familia where int_codfamilia = " & TBFamilia!ID_PC & " and Codigo is not null and nivel is not null", Conexao, adOpenKeyset, adLockOptimistic
                    If TBNivel2.EOF = False Then
                        
                        Set TBGravar = CreateObject("adodb.recordset")
                        TBGravar.Open "Select * from Plano_de_contas_totalizacao where Codigo = '" & TBNivel2!CODIGO & "' and Nivel = " & TBNivel2!Nivel & " and Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
                        ProcEnviaDadosPC TBNivel2!int_codfamilia, TBNivel2!CODIGO, TBNivel2!Txt_descricao, TBNivel2!Nivel
                        
                        ProcNivelPC1
                    End If
                    TBNivel2.Close
                Case 1:
                    Set TBNivel1 = CreateObject("adodb.recordset")
                    TBNivel1.Open "Select * from tbl_familia where int_codfamilia = " & TBFamilia!ID_PC & " and Codigo is not null and nivel is not null", Conexao, adOpenKeyset, adLockOptimistic
                    If TBNivel1.EOF = False Then
                        
                        Set TBGravar = CreateObject("adodb.recordset")
                        TBGravar.Open "Select * from Plano_de_contas_totalizacao where Codigo = '" & TBNivel1!CODIGO & "' and Nivel = " & TBNivel1!Nivel & " and Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
                        ProcEnviaDadosPC TBNivel1!int_codfamilia, TBNivel1!CODIGO, TBNivel1!Txt_descricao, TBNivel1!Nivel
                        
                    End If
                    TBNivel1.Close
            End Select
        End If
        TBAbrir.Close
        TBFamilia.MoveNext
    Loop
End If
TBFamilia.Close
ProcCarregaVisualizacao
txtSaldo = Format(Qtde - valor, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaVisualizacao()
On Error GoTo tratar_erro

valor = 0
With USTreeView1
    .Clear
    If Financeiro_Contas_Pagar = True Or Financeiro_Contas_Pagas = True Then
    
        'Adicionando as chaves principais
        If Financeiro_Contas_Pagar = True Then Tipo = "Contas a pagar" Else Tipo = "Contas pagas"
        Set Pagar = .Nodes.AddNode(Tipo, "B", , True, , , , 0, vbRed)
        
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from Plano_de_contas_totalizacao where Modulo = '" & Formulario & "' and Responsavel = '" & pubUsuario & "' order by Codigo", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Do While TBAbrir.EOF = False
                
                Descricao = TBAbrir!CODIGO & " - " & TBAbrir!Descricao & " - Valor: " & Format(TBAbrir!valor, "###,##0.00")
                IDlista = TBAbrir!ID_PC
                Nivel = TBAbrir!Nivel
               
                If Nivel = 8 Then
                    .Nodes.AddNode Descricao, IDlista, , , , , , , , NivelP7
                ElseIf Nivel = 7 Then
                        Set NivelP7 = .Nodes.AddNode(Descricao, IDlista, , , , , , , , NivelP6)
                    ElseIf Nivel = 6 Then
                            Set NivelP6 = .Nodes.AddNode(Descricao, IDlista, , , , , , , , NivelP5)
                        ElseIf Nivel = 5 Then
                                Set NivelP5 = .Nodes.AddNode(Descricao, IDlista, , , , , , , , NivelP4)
                            ElseIf Nivel = 4 Then
                                    Set NivelP4 = .Nodes.AddNode(Descricao, IDlista, , , , , , , , NivelP3)
                                ElseIf Nivel = 3 Then
                                        Set NivelP3 = .Nodes.AddNode(Descricao, IDlista, , , , , , , , NivelP2)
                                    ElseIf Nivel = 2 Then
                                            Set NivelP2 = .Nodes.AddNode(Descricao, IDlista, , , , , , , , NivelP1)
                                        Else
                                            Set NivelP1 = .Nodes.AddNode(Descricao, IDlista, , , , , , , , Pagar)
                                            valor = valor + TBAbrir!valor
                End If
                TBAbrir.MoveNext
            Loop
        End If
        TBAbrir.Close
    Else
        'Adicionando as chaves principais
        If Financeiro_Contas_Receber = True Then Tipo = "Contas a receber" Else Tipo = "Contas recebidas"
        Set Receber = .Nodes.AddNode(Tipo, "A", , True, , , , 0, vbBlue)
        
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from Plano_de_contas_totalizacao where Modulo = '" & Formulario & "' and Responsavel = '" & pubUsuario & "' order by Codigo", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Do While TBAbrir.EOF = False
                
                Descricao = TBAbrir!CODIGO & " - " & TBAbrir!Descricao & " - Valor: " & Format(TBAbrir!valor, "###,##0.00")
                IDlista = TBAbrir!ID_PC
                Nivel = TBAbrir!Nivel
               
                If Nivel = 8 Then
                    .Nodes.AddNode Descricao, IDlista, , , , , , , , NivelR7
                ElseIf Nivel = 7 Then
                        Set NivelR7 = .Nodes.AddNode(Descricao, IDlista, , , , , , , , NivelR6)
                    ElseIf Nivel = 6 Then
                            Set NivelR6 = .Nodes.AddNode(Descricao, IDlista, , , , , , , , NivelR5)
                        ElseIf Nivel = 5 Then
                                Set NivelR5 = .Nodes.AddNode(Descricao, IDlista, , , , , , , , NivelR4)
                            ElseIf Nivel = 4 Then
                                    Set NivelR4 = .Nodes.AddNode(Descricao, IDlista, , , , , , , , NivelR3)
                                ElseIf Nivel = 3 Then
                                        Set NivelR3 = .Nodes.AddNode(Descricao, IDlista, , , , , , , , NivelR2)
                                    ElseIf Nivel = 2 Then
                                            Set NivelR2 = .Nodes.AddNode(Descricao, IDlista, , , , , , , , NivelR1)
                                        Else
                                            Set NivelR1 = .Nodes.AddNode(Descricao, IDlista, , , , , , , , Receber)
                                            valor = valor + TBAbrir!valor
                End If
                
                TBAbrir.MoveNext
            Loop
        End If
        TBAbrir.Close
    End If
    .ExpandAllNodes True
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

If Financeiro_Contas_Pagar = True Then
    Formulario = "Financeiro/Contas a pagar/Conta contábil"
    With frmContas_Pagar
        txtIDConta = .txtidintconta
        FamiliaAntiga = .txtNDocumento
        Qtde = .txtValorTotal
    End With
ElseIf Financeiro_Contas_Receber = True Then
        Formulario = "Financeiro/Contas a receber/Conta contábil"
        With frmContas_Receber
            txtIDConta = .txtidintconta
            FamiliaAntiga = .txtDocumento
            Qtde = .txtValor
        End With
    ElseIf Financeiro_Contas_Pagas = True Then
            Formulario = "Financeiro/Contas pagas/Conta contábil"
            With frmContas_Pagas
                txtIDConta = .txtidintconta
                FamiliaAntiga = .txtNDocumento
                Qtde = .txt_ValorPago
            End With
        Else
            Formulario = "Financeiro/Contas recebidas/Conta contábil"
            With frmContas_recebidas
                txtIDConta = .txtidintconta
                FamiliaAntiga = .txtDocumento
                Qtde = .txtvalortitrecebido
            End With
End If
Direitos
ProcLimpaVariaveisPrincipais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluir()
On Error GoTo tratar_erro

If Txt_ID_PC = 0 Then
    USMsgBox ("Informe a conta contábil antes de excluir."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If USMsgBox("Deseja realmente excluir esta conta contábil?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    Conexao.Execute "DELETE FROM familia_financeiro WHERE Id = " & txtId
    USMsgBox ("Conta contábil excluído com sucesso."), vbInformation, "CAPRIND v5.0"
    '====================================
    Modulo = Formulario
    Evento = "Excluir"
    ID_documento = txtId
    Documento = "Documento: " & FamiliaAntiga
    Documento1 = "Código do plano: " & Txt_codigo_PC & " - Descrição do plano: " & Txt_descricao_PC
    ProcGravaEvento
    '===================================
    ProcLimpaCampos
    ProcVerifPC
    Novo_familia_financeiro = False
    Frame2.Enabled = False
    
    If Financeiro_Contas_Pagar = True Then
        frmContas_Pagar.ProcCarregaListaPC
    ElseIf Financeiro_Contas_Pagas = True Then
            frmContas_Pagas.ProcCarregaListaPC
        ElseIf Financeiro_Contas_Receber = True Then
                frmContas_Receber.ProcCarregaListaPC
            Else
                frmContas_recebidas.ProcCarregaListaPC
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovo()
On Error GoTo tratar_erro

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If txtSaldo = "0,00" Then
    USMsgBox ("Não é permitido criar nova conta contábil pois o saldo da conta é zero."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
ProcLimpaCampos
Novo_familia_financeiro = True
Frame2.Enabled = True
Cmd_localizar_PC_Click

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

If Novo_familia_financeiro = True Then
    If USMsgBox("A conta contábil ainda não foi salva, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvar
        If Novo_familia_financeiro = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
Novo_familia_financeiro = False
Unload Me

Conexao.Execute "DELETE from Plano_de_contas_totalizacao where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcPuxadaDados()
On Error GoTo tratar_erro

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from familia_financeiro where Idconta = " & txtIDConta & " and ID_PC = " & IDlista, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    txtId = TBAbrir!ID
    Txt_ID_PC = IIf(IsNull(TBAbrir!ID_PC), 0, TBAbrir!ID_PC)
    Set TBFamilia = CreateObject("adodb.recordset")
    TBFamilia.Open "Select * FROM tbl_familia where int_codfamilia = " & Txt_ID_PC, Conexao, adOpenKeyset, adLockOptimistic
    If TBFamilia.EOF = False Then
        Txt_codigo_PC = IIf(IsNull(TBFamilia!CODIGO), "", TBFamilia!CODIGO)
        Txt_descricao_PC = IIf(IsNull(TBFamilia!Txt_descricao), "", TBFamilia!Txt_descricao)
    End If
    txtValor.Text = Format(TBAbrir!valor, "###,##0.00")
    Novo_familia_financeiro = False
    Frame2.Enabled = True
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtValor_Change()
On Error GoTo tratar_erro

If txtValor <> "" Then
    VerifNumero = txtValor
    ProcVerificaNumero
    If VerifNumero = False Then
        txtValor = ""
        txtValor.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtValor_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus txtValor

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtValor_LostFocus()
On Error GoTo tratar_erro

txtValor = Format(txtValor, "###,##0.00")
    
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

Private Sub USTreeView1_NodeSelected(Node As DrawSuite2022.TreeNode)
On Error GoTo tratar_erro

If IsNumeric(Node.key) = True Then IDlista = Node.key Else IDlista = 0

With USToolBar1
    Set TBFI = CreateObject("adodb.recordset")
    TBFI.Open "Select * from tbl_familia where int_codfamilia = " & IDlista, Conexao, adOpenKeyset, adLockOptimistic
    If TBFI.EOF = False Then
        
        Select Case TBFI!Nivel
            Case 1: TextoFiltro = "Left(Codigo,1) = '" & Left(TBFI!CODIGO, 1) & "'"
            Case 2: TextoFiltro = "Left(Codigo,4) = '" & Left(TBFI!CODIGO, 4) & "'"
            Case 3: TextoFiltro = "Left(Codigo,7) = '" & Left(TBFI!CODIGO, 7) & "'"
            Case 4: TextoFiltro = "Left(Codigo,10) = '" & Left(TBFI!CODIGO, 10) & "'"
            Case 5: TextoFiltro = "Left(Codigo,13) = '" & Left(TBFI!CODIGO, 13) & "'"
            Case 6: TextoFiltro = "Left(Codigo,16) = '" & Left(TBFI!CODIGO, 16) & "'"
            Case 7: TextoFiltro = "Left(Codigo,19) = '" & Left(TBFI!CODIGO, 19) & "'"
            Case 8: TextoFiltro = "Left(Codigo,22) = '" & Left(TBFI!CODIGO, 22) & "'"
        End Select
        
        Set TBFIltro = CreateObject("adodb.recordset")
        TBFIltro.Open "Select * from tbl_familia where " & TextoFiltro & " order by Nivel", Conexao, adOpenKeyset, adLockOptimistic
        If TBFIltro.EOF = False Then
            TBFIltro.MoveLast
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select * from tbl_familia where int_codfamilia = " & IDlista & " and Nivel = " & TBFIltro!Nivel, Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                .ButtonState(3) = 0
            Else
                .ButtonState(3) = 5
            End If
            TBAbrir.Close
        End If
        TBFIltro.Close
    End If
    TBFI.Close
    .Refresh
End With
ProcLimpaCampos
ProcPuxadaDados

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
