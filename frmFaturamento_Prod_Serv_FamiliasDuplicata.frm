VERSION 5.00
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Begin VB.Form frmFaturamento_Prod_Serv_FamiliasDuplicata 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Administrativo - Faturamento - Nota fiscal - Conta contábil"
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12195
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   12195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame7 
      BackColor       =   &H00E0E0E0&
      Height          =   585
      Left            =   60
      TabIndex        =   12
      Top             =   7350
      Width           =   12105
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
         Left            =   10290
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Saldo."
         Top             =   180
         Width           =   1620
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo :"
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
         Left            =   9630
         TabIndex        =   13
         Top             =   180
         Width           =   570
         WordWrap        =   -1  'True
      End
   End
   Begin DrawSuite2022.USTreeView USTreeView1 
      Height          =   5505
      Left            =   60
      TabIndex        =   4
      Top             =   1830
      Width           =   12105
      _ExtentX        =   21352
      _ExtentY        =   9710
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
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   6
      Top             =   0
      Width           =   12105
      _ExtentX        =   21352
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
         Left            =   5790
         Top             =   -30
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmFaturamento_Prod_Serv_FamiliasDuplicata.frx":0000
         Count           =   1
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   825
      Left            =   55
      TabIndex        =   7
      Top             =   990
      Width           =   12105
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
         Left            =   10620
         MaxLength       =   20
         TabIndex        =   3
         ToolTipText     =   "Valor."
         Top             =   390
         Width           =   1290
      End
      Begin VB.CommandButton Cmd_localizar_PC 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   10230
         Picture         =   "frmFaturamento_Prod_Serv_FamiliasDuplicata.frx":33EC
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Localizar plano de contas."
         Top             =   390
         Width           =   315
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
         TabIndex        =   8
         Text            =   "0"
         ToolTipText     =   "ID PC."
         Top             =   390
         Visible         =   0   'False
         Width           =   765
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
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Descrição."
         Top             =   390
         Width           =   8145
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
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Valor"
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
         Left            =   10620
         TabIndex        =   11
         Top             =   180
         Width           =   1290
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
         Left            =   862
         TabIndex        =   10
         Top             =   180
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
         Left            =   5782
         TabIndex        =   9
         Top             =   180
         Width           =   720
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "frmFaturamento_Prod_Serv_FamiliasDuplicata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Novo_familia_faturamento As Boolean 'OK

Private Sub Cmd_localizar_PC_Click()
On Error GoTo tratar_erro

Plano_contas_produtos = False
Plano_contas_familias = False
Plano_centro_de_custo = False
Plano_instituicao = False
Plano_opcoesgerais = False
Plano_Faturamento = True
Financeiro_Contas_Pagar = False
Financeiro_Contas_Pagas = False
Financeiro_Contas_Receber = False
Financeiro_Contas_Recebidas = False
Plano_PCP = False
Sit_REG = 1
frmproj_produto_PC.Show 1

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
    Case vbKeyEscape:
        Conexao.Execute "DELETE from Plano_de_contas_totalizacao where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'"
        Unload Me
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
If IsNumeric(txtValor) = False Then
    NomeCampo = "o valor"
    ProcVerificaAcao
    txtValor.SetFocus
    Exit Sub
End If

With frmFaturamento_Prod_Serv
    'Verifica saldo
    qtdeliberada = txtSaldo
    qt = Format(txtValor, "###,##0.00")
    Set TBGravar = CreateObject("adodb.recordset")
    TBGravar.Open "Select Sum(Valor) as Valor from familia_financeiro where IDnota = " & .txtId, Conexao, adOpenKeyset, adLockOptimistic
    If qtdeliberada < 0 Then qtdeliberada = qtdeliberada * -1
    valor = IIf(IsNull(TBGravar!valor), 0, TBGravar!valor)
    If valor > 0 Then
        Qtd = Format(qtdeliberada + valor, "###,##0.00")
        If qt <> valor And Qtd < qt Then
            USMsgBox ("Não é permitido salvar pois o saldo da conta é menor que o valor da conta contábil."), vbExclamation, "CAPRIND v5.0"
            Exit Sub
        End If
    Else
        Qtd = txtSaldo
        If Qtd < qt Then
            USMsgBox ("Não é permitido salvar pois o saldo da conta é menor que o valor da conta contábil."), vbExclamation, "CAPRIND v5.0"
            Exit Sub
        End If
    End If
    
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from tbl_Detalhes_Recebimento where id_nota = " & .txtId, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Do While TBAbrir.EOF = False
            'Verifica se é nota de devolução
            Permitido = False
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select CFOP.* from tbl_Detalhes_Nota NFP INNER JOIN tbl_NaturezaOperacao CFOP ON CFOP.IDCountCfop = NFP.ID_CFOP where NFP.ID_nota = " & .txtId & " and CFOP.Devolucao = 'True'", Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False Then
                Permitido = True
                
                'Fornecedor
                If .opt_Saida.Value = True And Len(.txttipocliente) = 1 Then Permitido1 = True
                'Cliente
                If .opt_Entrada.Value = True And Len(.txttipocliente) = 2 Then Permitido1 = False
                
                ValorTotal = "-" & Qtde
                Valor3 = "-" & TBAbrir!dbl_Valor
            Else
                ValorTotal = Qtde
                Valor3 = TBAbrir!dbl_Valor
            End If
            TBFI.Close
            NovoValor = Replace(Valor3, ",", ".") 'Verifica a porcentagem representada pelo valor da duplicata e ID da conta
            
            If Permitido = True Then
                If Permitido1 = True Then GoTo SalvarCP Else GoTo SalvarCR
            End If
            
            If .opt_Saida.Value = True Then
SalvarCR:
                Set TBReceber = CreateObject("adodb.recordset")
                TBReceber.Open "Select * from tbl_contas_receber where id_nota = " & .txtId & " and Vencimento = '" & TBAbrir!dt_Vencimento & "' and Valor = " & NovoValor, Conexao, adOpenKeyset, adLockOptimistic
                Valor2 = TBReceber!valor
                TipoConta = "R"
            Else
SalvarCP:
                Set TBReceber = CreateObject("adodb.recordset")
                TBReceber.Open "Select * from tbl_ContasPagar where id_nota = " & .txtId & " and dt_Pagamento = '" & TBAbrir!dt_Vencimento & "' and dbl_valorpagto = " & NovoValor, Conexao, adOpenKeyset, adLockOptimistic
                Valor2 = TBReceber!dbl_valorpagto
                TipoConta = "P"
            End If
            
            Valor1 = Format((Valor2 * 100) / ValorTotal, "###,##0.0000000000")
            
            Set TBCiclo = CreateObject("adodb.recordset")
            If Novo_familia_faturamento = True Then TextoFiltro = "ID_PC = " & Txt_ID_PC Else TextoFiltro = "ID_PC = " & IDpedido
            TBCiclo.Open "select * from familia_financeiro where " & TextoFiltro & " and IDConta = " & TBReceber!IDintconta & " and TipoConta = '" & TipoConta & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBCiclo.EOF = True Then TBCiclo.AddNew
            TBCiclo!ID_PC = Txt_ID_PC
            TBCiclo!IDConta = TBReceber!IDintconta
            TBCiclo!IDnota = .txtId
            
            If Permitido = True Then qtdeliberada = "-" & txtValor Else qtdeliberada = txtValor
            valor = Format((qtdeliberada * Valor1) / 100, "###,##0.00")
            TBCiclo!valor = Format(valor, "###,##0.00")
            
            TBCiclo!TipoConta = TipoConta
            TBCiclo.Update
            TBCiclo.Close
            
            TBAbrir.MoveNext
        Loop
    End If
    
    If Novo_familia_faturamento = True Then
        USMsgBox ("Nova conta contábil cadastrada com sucesso."), vbInformation, "CAPRIND v5.0"
        Evento = "Nova conta conábil"
    Else
        USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
        Evento = "Alterar conta conábil"
    End If
    ProcVerifPC
    '==================================
    Modulo = Formulario
    ID_documento = .txtId
    .ProcVerificaTipoNF False
    If .txtNFiscal = "" Then NomeCampo = "N° ordem: " & .txtId Else NomeCampo = "N° nota: " & .txtNFiscal
    Documento = NomeCampo & " - Tipo: " & TipoNF & " - Série: " & .txtSerie
    Documento1 = "Código do plano: " & Txt_codigo_PC & " - Descrição do plano: " & Txt_descricao_PC
    ProcGravaEvento
    '==================================
    Novo_familia_faturamento = False
End With
        
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 12105, 7, True
If Formulario = "Faturamento/Nota fiscal/Própria" Then
    Caption = "Administrativo - Faturamento - Nota fiscal - Própria - Conta contábil"
ElseIf Formulario = "Faturamento/Nota fiscal/Terceiros" Then
        Caption = "Administrativo - Faturamento - Nota fiscal - Terceiros - Conta contábil"
    ElseIf Formulario = "Estoque/Ordem de faturamento" Then
            Caption = "Estoque - Ordem de faturamento - Conta contábil"
        Else
            Caption = "Estoque - Nota fiscal - Conta contábil"
End If
ProcLimpaVariaveisPrincipais

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select Sum(dbl_Valor) as Qtde from tbl_Detalhes_Recebimento where id_nota = " & frmFaturamento_Prod_Serv.txtId, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    Qtde = IIf(IsNull(TBAbrir!Qtde), 0, TBAbrir!Qtde)
End If
TBAbrir.Close

ProcVerifPC

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCampos()
On Error GoTo tratar_erro

'txtID.Text = 0
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

Valor1 = 0
With frmFaturamento_Prod_Serv
    Set TBFamilia = CreateObject("adodb.recordset")
    TBFamilia.Open "Select ID_PC, Sum(valor) as Valor from Familia_financeiro where IDNota = '" & .txtId & "' and ID_PC is not null and ID_PC <> 0 Group by ID_PC", Conexao, adOpenKeyset, adLockOptimistic
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
End With
ProcCarregaVisualizacao
txtSaldo = Format(Qtde - IIf(valor < 0, (valor * -1), valor), "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaVisualizacao()
On Error GoTo tratar_erro

valor = 0
Permitido = False
With frmFaturamento_Prod_Serv
    'Verifica se é nota de devolução
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select CFOP.* from tbl_Detalhes_Nota NFP INNER JOIN tbl_NaturezaOperacao CFOP ON CFOP.IDCountCfop = NFP.ID_CFOP where NFP.ID_nota = " & .txtId & " and CFOP.Devolucao = 'True'", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Permitido = True
        
        'Fornecedor
        If .opt_Saida.Value = True And Len(.txttipocliente) = 1 Then Permitido1 = True
        'Cliente
        If .opt_Entrada.Value = True And Len(.txttipocliente) = 2 Then Permitido1 = False
    End If
    TBAbrir.Close
End With

With USTreeView1
    .Clear
    
    If Permitido = True Then
        If Permitido1 = True Then GoTo VisualizarCP Else GoTo VisualizarCR
    End If
    
    If frmFaturamento_Prod_Serv.opt_Entrada.Value = True Then
VisualizarCP:
        'Adicionando as chaves principais
        Set Pagar = .Nodes.AddNode("Contas a pagar", "B", , True, , , , 0, vbRed)
        
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
VisualizarCR:
        'Adicionando as chaves principais
        Set Receber = .Nodes.AddNode("Contas a receber", "A", , True, , , , 0, vbBlue)
        
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

Private Sub ProcExcluir()
On Error GoTo tratar_erro

If Txt_ID_PC = 0 Then
    USMsgBox ("Informe a conta contábil antes de excluir."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If USMsgBox("Deseja realmente excluir esta conta contábil?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    With frmFaturamento_Prod_Serv
        Conexao.Execute "DELETE FROM familia_financeiro WHERE IDNota = " & .txtId & " and ID_PC = " & Txt_ID_PC & " and Deposito_transf = 'False'"
        USMsgBox ("Conta contábil excluído com sucesso."), vbInformation, "CAPRIND v5.0"
        '==================================
        Modulo = Formulario
        ID_documento = .txtId
        .ProcVerificaTipoNF False
        If .txtNFiscal = "" Then NomeCampo = "N° ordem: " & .txtId Else NomeCampo = "N° nota: " & .txtNFiscal
        Documento = NomeCampo & " - Tipo: " & TipoNF & " - Série: " & .txtSerie
        Documento1 = "Código do plano: " & Txt_codigo_PC & " - Descrição do plano: " & Txt_descricao_PC
        ProcGravaEvento
        '==================================
        ProcLimpaCampos
        ProcVerifPC
        Novo_familia_faturamento = False
        Frame2.Enabled = False
    End With
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
Novo_familia_faturamento = True
Frame2.Enabled = True
Cmd_localizar_PC_Click

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

If Novo_familia_faturamento = True Then
    If USMsgBox("A conta contábil ainda não foi salva, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvar
        If Novo_familia_faturamento = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
Novo_familia_faturamento = False
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
TBAbrir.Open "Select Sum(valor) as Valor from familia_financeiro where IdNota = " & frmFaturamento_Prod_Serv.txtId & " and ID_PC = " & IDlista, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    Txt_ID_PC = IDlista
    Set TBFamilia = CreateObject("adodb.recordset")
    TBFamilia.Open "Select * FROM tbl_familia where int_codfamilia = " & Txt_ID_PC, Conexao, adOpenKeyset, adLockOptimistic
    If TBFamilia.EOF = False Then
        Txt_codigo_PC = IIf(IsNull(TBFamilia!CODIGO), "", TBFamilia!CODIGO)
        Txt_descricao_PC = IIf(IsNull(TBFamilia!Txt_descricao), "", TBFamilia!Txt_descricao)
    End If
    txtValor.Text = IIf(IsNull(TBAbrir!valor), 0, Format(TBAbrir!valor, "###,##0.00"))
    Novo_familia_faturamento = False
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

If IsNumeric(Node.key) = True Then
    IDlista = Node.key
    IDpedido = Node.key
Else
    IDlista = 0
    IDpedido = 0
End If

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
