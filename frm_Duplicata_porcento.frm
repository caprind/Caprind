VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frm_Duplicata_porcento 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Nota fiscal - Parcelamento em porcentagem"
   ClientHeight    =   2715
   ClientLeft      =   2865
   ClientTop       =   2385
   ClientWidth     =   5850
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frm_Duplicata_porcento.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   5850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Chk_alterar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Utilizar saldo da nota para cálculo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   3375
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   55
      TabIndex        =   16
      Top             =   0
      Width           =   5715
      _ExtentX        =   10081
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft1     =   2
      ButtonTop1      =   2
      ButtonWidth1    =   38
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
      ButtonLeft2     =   42
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft3     =   46
      ButtonTop3      =   2
      ButtonWidth3    =   36
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft4     =   84
      ButtonTop4      =   2
      ButtonWidth4    =   26
      ButtonHeight4   =   21
      ButtonUseMaskColor4=   0   'False
      ButtonEnabled5  =   0   'False
      ButtonIconSize5 =   32
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
      ButtonState5    =   5
      ButtonLeft5     =   112
      ButtonTop5      =   2
      ButtonWidth5    =   24
      ButtonHeight5   =   24
      ButtonUseMaskColor5=   0   'False
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   2910
         Top             =   150
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frm_Duplicata_porcento.frx":1042
         Count           =   1
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   1395
      Left            =   55
      MousePointer    =   1  'Arrow
      TabIndex        =   10
      Top             =   1290
      Width           =   5715
      Begin VB.OptionButton optParcela 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Valor parcela(R$)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   300
         TabIndex        =   6
         Top             =   780
         Width           =   1545
      End
      Begin VB.OptionButton optPorcentagem 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Vlr. parcela(%)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4080
         TabIndex        =   4
         Top             =   180
         Value           =   -1  'True
         Width           =   1395
      End
      Begin VB.TextBox txtvalordin 
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
         Height          =   315
         Left            =   180
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   7
         Text            =   "0,00"
         ToolTipText     =   "Valor da parcela em reais."
         Top             =   990
         Width           =   1785
      End
      Begin VB.TextBox txtvalorporc 
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
         Height          =   315
         Left            =   4020
         MaxLength       =   50
         TabIndex        =   5
         ToolTipText     =   "Valor da parcela em porcentagem."
         Top             =   390
         Width           =   1515
      End
      Begin VB.TextBox txtsaldoporc 
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
         Height          =   315
         Left            =   3780
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   "0"
         ToolTipText     =   "Saldo da nota em porcentagem."
         Top             =   990
         Width           =   1755
      End
      Begin VB.TextBox txtsaldodin 
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
         Height          =   315
         Left            =   1980
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Saldo da nota em reais."
         Top             =   990
         Width           =   1785
      End
      Begin VB.TextBox txtvalorduplicata 
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
         Height          =   315
         Left            =   180
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Valor da nota."
         Top             =   390
         Width           =   1635
      End
      Begin MSComCtl2.DTPicker txtvenc 
         Height          =   315
         Left            =   2610
         TabIndex        =   3
         ToolTipText     =   "Data de vencimento da conta."
         Top             =   390
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
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
         Format          =   488570881
         CurrentDate     =   39057
      End
      Begin MSMask.MaskEdBox txtparcela 
         Height          =   315
         Left            =   1830
         TabIndex        =   2
         ToolTipText     =   "Número da parcela."
         Top             =   390
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         MaxLength       =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "###/###"
         PromptChar      =   "_"
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Parcela"
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
         Left            =   1960
         MousePointer    =   1  'Arrow
         TabIndex        =   15
         Top             =   180
         Width           =   525
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo nota(%)"
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
         Left            =   4132
         MousePointer    =   1  'Arrow
         TabIndex        =   14
         Top             =   780
         Width           =   1050
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo nota(R$)"
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
         Left            =   2332
         MousePointer    =   1  'Arrow
         TabIndex        =   13
         Top             =   780
         Width           =   1080
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Vencimento"
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
         Left            =   2945
         MousePointer    =   1  'Arrow
         TabIndex        =   12
         Top             =   180
         Width           =   825
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Valor nota"
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
         Left            =   615
         MousePointer    =   1  'Arrow
         TabIndex        =   11
         Top             =   180
         Width           =   735
      End
   End
End
Attribute VB_Name = "frm_Duplicata_porcento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Lista As Double 'OK

Private Sub Chk_alterar_Click()
On Error GoTo tratar_erro

ProcVerifValorParcela

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvar()
On Error GoTo tratar_erro

With frmFaturamento_Prod_Serv
    Acao = "salvar"
    txtparcela.PromptInclude = False
    If Len(txtparcela) < 6 Then
        USMsgBox ("O número da parcela digitada não é válido, digite o número correto."), vbExclamation, "CAPRIND v5.0"
        txtparcela.SetFocus
        Exit Sub
    End If
    txtparcela.PromptInclude = True
    If optPorcentagem.Value = True Then
        If txtvalorporc = "" Then
            NomeCampo = "o valor da parcela(%)"
            ProcVerificaAcao
            txtvalorporc.SetFocus
            Exit Sub
        End If
    End If
    If optParcela.Value = True Then
        If txtvalordin = "" Then
            NomeCampo = "o valor da parcela"
            ProcVerificaAcao
            txtvalordin.SetFocus
            Exit Sub
        End If
    End If
    Parcela = txtparcela.Text
    Lista = 0
    Do Until Lista = .lst_Duplicata.ListItems.Count
        If Parcela = .lst_Duplicata.ListItems((Lista + 1)).ListSubItems(2) Then
            USMsgBox ("Não é permitido adicionar mais de uma parcela com o mesmo nome."), vbExclamation, "CAPRIND v5.0"
            txtparcela.Text = "___/___"
            txtparcela.SetFocus
            Lista = 0
            Exit Sub
        End If
        Lista = Lista + 1
    Loop
    'Verifica saldo em porcentagem
    SaldoPorc = txtsaldoporc.Text
    If SaldoPorc <= 0 Then
        USMsgBox ("Não é permitido salvar, pois o saldo da nota é igual a zero."), vbExclamation, "CAPRIND v5.0"
        Exit Sub
    End If
    'Calcula saldo da nota
    SaldoPedido = txtsaldodin.Text
    Valorparcela = txtvalordin.Text
        If Valorparcela > SaldoPedido Then
            If optPorcentagem.Value = True Then
            USMsgBox ("Verifique a porcentagem digitada."), vbInformation, "CAPRIND v5.0"
            txtvalorporc.SetFocus
        Else
            USMsgBox ("Verifique o valor digitado."), vbInformation, "CAPRIND v5.0"
            txtvalordin.SetFocus
        End If
        Exit Sub
    End If
    txtsaldodin.Text = Format(SaldoPedido - Valorparcela, "###,##0.00")

    Set TBVendas = CreateObject("adodb.recordset")
    TBVendas.Open "Select * from tbl_detalhes_recebimento", Conexao, adOpenKeyset, adLockOptimistic
    TBVendas.AddNew
    TBVendas!int_NotaFiscal = IIf(.txtNFiscal = "", Null, .txtNFiscal)
    TBVendas!ID_nota = .txtID.Text
    TBVendas!txt_Parcela = txtparcela
    TBVendas!dt_Vencimento = txtvenc
    TBVendas!txt_Portador_Banco = .cbo_PortBanco
    TBVendas!txt_Agencia = .txt_Agencia
    TBVendas!txt_Conta = .txt_Conta
    TBVendas!dbl_Valor = txtvalordin
    TBVendas!Valor_Extenso = FunValorExtenso(TBVendas!dbl_Valor)
    TBVendas!txt_tipopagto = .cmb_tipopagto
    ID = TBVendas!ID
    TBVendas.Update
    TBVendas.Close
    .ProcCarregaListaDuplicatas IIf(.txtID = "", 0, .txtID)
End With
ProcLimpaCampos
txtparcela.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyF3: ProcSalvar
    Case vbKeyEscape: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 5715, 5, True

If Formulario = "Faturamento/Nota fiscal/Própria" Then
    Caption = "Nota fiscal - Própria - Parc. em porcentagem"
ElseIf Formulario = "Faturamento/Nota fiscal/Terceiros" Then
        Caption = "Nota fiscal - Terceiros - Parc. em porcentagem"
    ElseIf Formulario = "Estoque/Ordem de faturamento" Then
            Caption = "Ordem de faturamento - Parc. em porcentagem"
        Else
            Caption = "Nota fiscal - Parc. em porcentagem"
End If
contador = 0
ValoresParcelas = 0
With frmFaturamento_Prod_Serv
    Set TBTotaisnota = CreateObject("adodb.recordset")
    TBTotaisnota.Open "Select * from tbl_Totais_Nota where id_nota = " & .txtID, Conexao, adOpenKeyset, adLockOptimistic
    If TBTotaisnota.EOF = False Then
        Valor_Retencao_PIS = IIf(IsNull(TBTotaisnota!Total_retencao_PIS), 0, TBTotaisnota!Total_retencao_PIS)
        Valor_Retencao_Cofins = IIf(IsNull(TBTotaisnota!Total_retencao_Cofins), 0, TBTotaisnota!Total_retencao_Cofins)
    End If
    TBTotaisnota.Close
    
    'Verifica produtos de remessa
    qt = 0
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select Sum(dbl_ValorTotal) as qt from tbl_Detalhes_Nota where id_nota = " & .txtID.Text & " and Remessa = 'True'", Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
        qt = IIf(IsNull(TBProduto!qt), 0, TBProduto!qt)
    End If
    
    'Verifica produtos de retorno que somam o valor no total da nota
    VlrTotalRetorno = 0
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select Sum(dbl_ValorTotal) as VlrTotalRetorno from tbl_Detalhes_Nota where id_nota = " & .txtID.Text & " and Retorno = 'True' and Soma_retorno_totalnf = 'True'", Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
        VlrTotalRetorno = IIf(IsNull(TBProduto!VlrTotalRetorno), 0, TBProduto!VlrTotalRetorno)
    End If
    TBProduto.Close
    
    If .lst_Duplicata.ListItems.Count > 0 Then
        Do Until contador = .lst_Duplicata.ListItems.Count
            ValoresParcelas = Format(ValoresParcelas + .lst_Duplicata.ListItems((contador + 1)).ListSubItems(1), "###,##0.00")
            contador = contador + 1
        Loop
    End If
    .ProcVerificaTipoNF False
    If TipoNF = "M1" Then
        valor = IIf(.txt_ValorNota = "", 0, .txt_ValorNota) - Valor_Retencao_PIS - Valor_Retencao_Cofins - VlrTotalRetorno - qt
    ElseIf TipoNF = "SA" Then
            valor = IIf(.txttotalservicos = "", 0, .txttotalservicos)
        Else
            valor = (IIf(.txt_ValorNota = "", 0, .txt_ValorNota) - Valor_Retencao_PIS - Valor_Retencao_Cofins - VlrTotalRetorno - qt) + IIf(.txttotalservicos = "", 0, .txttotalservicos)
    End If
    .ProcVerificaEmpresaCliente
    If Permitido = True Then .ProcVerificaImposto Else ValorTotal = valor
    
    txtValorDuplicata = Format(valor, "###,##0.00")
    txtsaldodin = Format(valor - ValoresParcelas, "###,##0.00")
End With
txtvenc.Value = Date

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

Private Sub optParcela_Click()
On Error GoTo tratar_erro

txtvalorporc = ""
txtvalordin = ""
txtvalordin.Locked = False
txtvalorporc.Locked = True
txtvalordin.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optPorcentagem_Click()
On Error GoTo tratar_erro

txtvalorporc = ""
txtvalordin = ""
txtvalordin.Locked = True
txtvalorporc.Locked = False
txtvalorporc.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtsaldodin_Change()
On Error GoTo tratar_erro

Totalduplicata = txtValorDuplicata.Text
Valorparcela = txtsaldodin.Text
If Totalduplicata <> 0 Then
    Porcentagem = (Valorparcela * 100) / Totalduplicata
    txtsaldoporc = Porcentagem
Else
    txtsaldoporc = 0
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtvalordin_Change()
On Error GoTo tratar_erro

If txtvalordin.Text <> "" Then
    VerifNumero = txtvalordin.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtvalordin.Text = ""
        txtvalordin.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtvalordin_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus txtvalordin

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtvalorporc_Change()
On Error GoTo tratar_erro

If txtvalorporc.Text <> "" Then
    If txtvalorporc.Text <> "" Then
        VerifNumero = txtvalorporc.Text
        ProcVerificaNumero
        If VerifNumero = False Then
            txtvalorporc.Text = ""
            txtvalorporc.SetFocus
            Exit Sub
        End If
    End If
End If
ProcVerifValorParcela

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcVerifValorParcela()
On Error GoTo tratar_erro

If Chk_alterar.Value = 0 Then Totalduplicata = txtValorDuplicata Else Totalduplicata = txtsaldodin
Porcentagem = IIf(txtvalorporc = "", 0, txtvalorporc)
Valorparcela = (Totalduplicata * Porcentagem) / 100
txtvalordin.Text = Format(Valorparcela, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCampos()
On Error GoTo tratar_erro

txtparcela.Text = "___/___"
txtvenc.Value = Date
txtvalorporc.Text = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtvalorporc_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus txtvalorporc

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
