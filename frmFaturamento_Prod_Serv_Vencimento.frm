VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Begin VB.Form frmFaturamento_Prod_Serv_Vencimento 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Faturamento | Duplicatas"
   ClientHeight    =   4770
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFaturamento_Prod_Serv_Vencimento.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4770
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   16
      Top             =   4365
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   714
   End
   Begin VB.TextBox Txt_ID_duplicata 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   -600
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   1110
      Width           =   555
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   450
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   794
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
      ShowMaximizeButton=   0   'False
      ShowMinimizeButton=   0   'False
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Dados da duplicata"
      Height          =   3525
      Left            =   210
      TabIndex        =   1
      Top             =   570
      Width           =   4215
      Begin VB.ComboBox cmbtipo_conta 
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
         ItemData        =   "frmFaturamento_Prod_Serv_Vencimento.frx":000C
         Left            =   2670
         List            =   "frmFaturamento_Prod_Serv_Vencimento.frx":000E
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   13
         ToolTipText     =   "Tipo do documento."
         Top             =   1875
         Width           =   1185
      End
      Begin VB.ComboBox cmb_tipopagto 
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
         ItemData        =   "frmFaturamento_Prod_Serv_Vencimento.frx":0010
         Left            =   1650
         List            =   "frmFaturamento_Prod_Serv_Vencimento.frx":004A
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   12
         ToolTipText     =   "Forma de pagamento/recebimento."
         Top             =   2265
         Width           =   2205
      End
      Begin VB.TextBox txt_Duplicata 
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
         Height          =   315
         Left            =   2685
         MaxLength       =   50
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Numero da parcela"
         Top             =   300
         Width           =   1170
      End
      Begin DrawSuite2022.USButton btnAlterar 
         Height          =   585
         Left            =   1650
         TabIndex        =   9
         Top             =   2700
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1032
         DibPicture      =   "frmFaturamento_Prod_Serv_Vencimento.frx":0158
         Caption         =   "Gravar alterações"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   5263559
         BorderColorDisabled=   13160660
         BorderColorDown =   4013465
         BorderColorOver =   4408288
         GradientColor1  =   5263559
         GradientColor2  =   5263559
         GradientColor3  =   5263559
         GradientColor4  =   5263559
         GradientColorDisabled1=   13160660
         GradientColorDisabled2=   13160660
         GradientColorDisabled3=   13160660
         GradientColorDisabled4=   13160660
         GradientColorOver1=   4408288
         GradientColorOver2=   4408288
         GradientColorOver3=   4408288
         GradientColorOver4=   4408288
         GradientColorDown1=   4013465
         GradientColorDown2=   4013465
         GradientColorDown3=   4013465
         GradientColorDown4=   4013465
         PicAlign        =   7
         ShowFocusRect   =   0   'False
         Theme           =   4
      End
      Begin VB.TextBox txtVencimentoAntigo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   2685
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Vencimento atual da duplicata"
         Top             =   1080
         Width           =   1170
      End
      Begin VB.TextBox txtValorDuplicata 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   2685
         MaxLength       =   50
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Valor da duplicata"
         Top             =   690
         Width           =   1170
      End
      Begin MSComCtl2.DTPicker txt_Vencimento 
         Height          =   330
         Left            =   2670
         TabIndex        =   0
         ToolTipText     =   "Novo vencimento da duplicata"
         Top             =   1470
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   16777215
         CalendarForeColor=   0
         CalendarTitleBackColor=   8421504
         CalendarTitleForeColor=   16777215
         CalendarTrailingForeColor=   255
         Format          =   171376643
         CurrentDate     =   39057
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Documento:"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   19
         Left            =   1770
         TabIndex        =   15
         Top             =   1980
         Width           =   870
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Recebimento em:"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   137
         Left            =   375
         MousePointer    =   1  'Arrow
         TabIndex        =   14
         Top             =   2340
         Width           =   1260
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Novo Vencimento:"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   0
         Left            =   1320
         MousePointer    =   1  'Arrow
         TabIndex        =   7
         Top             =   1560
         Width           =   1305
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Valor da parcela:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   134
         Left            =   1440
         MousePointer    =   1  'Arrow
         TabIndex        =   6
         Top             =   780
         Width           =   1215
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Vencimento atual:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   135
         Left            =   1365
         MousePointer    =   1  'Arrow
         TabIndex        =   5
         Top             =   1170
         Width           =   1290
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Parcela:"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   136
         Left            =   2055
         MousePointer    =   1  'Arrow
         TabIndex        =   4
         Top             =   390
         Width           =   585
      End
   End
End
Attribute VB_Name = "frmFaturamento_Prod_Serv_Vencimento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ProcCarregaComboForma()
On Error GoTo tratar_erro

Familiatext = ""
cmb_tipopagto.Clear
If frmFaturamento_Prod_Serv.opt_Saida.Value = True Then TextoFiltro = "Tipo = 'R'" Else TextoFiltro = "Tipo = 'P'"

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from tbl_FormaPagto where " & TextoFiltro & " order by Descricao", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    Do While TBAbrir.EOF = False
        If Familiatext <> TBAbrir!Descricao Then cmb_tipopagto.AddItem TBAbrir!Descricao
        Familiatext = TBAbrir!Descricao
        TBAbrir.MoveNext
    Loop
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub btnAlterar_Click()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If frmFaturamento_Prod_Serv.lst_Duplicata.ListItems.Count = 0 Then Exit Sub
frmFaturamento_Prod_Serv.FunVerifAlteracaoNF frmFaturamento_Prod_Serv.txtid, True, frmFaturamento_Prod_Serv.optServico
If Permitido1 = False Then
    'USMsgBox ("Não é permitido salvar, " & MsgTexto & "."), vbExclamation, "CAPRIND v5.0"
    'Exit Sub
End If
Acao = "salvar"
If Txt_ID_duplicata = 0 Then
    NomeCampo = "a duplicata na lista"
    ProcVerificaAcao
    Exit Sub
End If
If txtValorDuplicata = "" Then
    NomeCampo = "o valor"
    ProcVerificaAcao
    txtValorDuplicata.SetFocus
    Exit Sub
End If
If txtValorDuplicata <= 0 Then
    USMsgBox ("Não é permitido salvar, pois o valor da duplicata não pode ser menor ou igual a zero."), vbExclamation, "CAPRIND v5.0"
    txtValorDuplicata.SetFocus
    Exit Sub
End If
With txt_Duplicata
'    '.PromptInclude = False
'    If Len(.Text) < 6 Then
'        USMsgBox ("O número da parcela digitada não é válido, digite o número correto."), vbExclamation, "CAPRIND v5.0"
'        .SetFocus
'        Exit Sub
'    End If
'    .PromptInclude = True
End With
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from tbl_detalhes_recebimento WHERE ID = " & Txt_ID_duplicata, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = False Then
    If TBGravar!Nosso_Numero <> "" Then
        USMsgBox ("Não é permitido alterar os dados desta duplicata, pois a mesma já tem boleto emitido."), vbExclamation, "CAPRIND v5.0"
        TBGravar.Close
        Exit Sub
    End If
    TBGravar!txt_Portador_Banco = frmFaturamento_Prod_Serv.cbo_PortBanco
    TBGravar!txt_Agencia = frmFaturamento_Prod_Serv.txt_Agencia
    TBGravar!txt_Conta = frmFaturamento_Prod_Serv.txt_Conta
    TBGravar!dt_Vencimento = txt_Vencimento.Value
    TBGravar!dbl_Valor = txtValorDuplicata
    TBGravar!txt_Parcela = txt_Duplicata
    TBGravar!Valor_Extenso = FunValorExtenso(TBGravar!dbl_Valor)
    TBGravar!txt_tipopagto = cmb_tipopagto
    TBGravar!Tipo_doc = cmbtipo_conta.Text
    TBGravar.Update
    frmFaturamento_Prod_Serv.ProcCarregaListaDuplicatas IIf(frmFaturamento_Prod_Serv.txtid = "", 0, frmFaturamento_Prod_Serv.txtid)
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    '==================================
    Modulo = Formulario
    Evento = "Alterar duplicata"
    ID_documento = Txt_ID_duplicata
    frmFaturamento_Prod_Serv.ProcVerificaTipoNF False
    If txtNFiscal = "" Then NomeCampo = "N° ordem: " & txtid Else NomeCampo = "N° nota: " & txtNFiscal
    Documento = NomeCampo & " - Tipo: " & TipoNF & " - Série: " & frmFaturamento_Prod_Serv.txtSerie
    Documento1 = "Data vencimento: " & Format(txt_Vencimento.Value, "dd/mm/yy") & " - Valor: " & Format(txtValorDuplicata, "###,##0.00") & " - Parcela: " & txt_Duplicata
    ProcGravaEvento
    '==================================
    
'    If frmFaturamento_Prod_Serv.txtNFiscal <> "" And frmFaturamento_Prod_Serv.txtDtValidacao <> "" Then
'        If USMsgBox("Deseja enviar a(s) duplicata(s) para o financeiro agora?", vbyesno, "CAPRIND v5.0") = vbYes Then frmFaturamento_Prod_Serv.ProcEnviarDupFinanceiro IIf(txtid = "", 0, txtid), True
'    End If
    
End If
TBGravar.Close
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub


Private Sub Form_Load()
On Error GoTo tratar_erro
                
If frmFaturamento_Prod_Serv.opt_Saida.Value = True Then
TextoFiltro = "Tipo = 'R'"
ProcCarregaComboTipoDocto cmbtipo_conta, "Tipo = 'R'"

Else
TextoFiltro = "Tipo = 'P'"
ProcCarregaComboTipoDocto cmbtipo_conta, "Tipo = 'P'"
End If
                
ProcCarregaComboForma

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub
