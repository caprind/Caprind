VERSION 5.00
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form frmprod_outras_despesas 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "PCP | Gerenciamento de ordem - Cadastrar custo de outras despesas"
   ClientHeight    =   8640
   ClientLeft      =   60
   ClientTop       =   15
   ClientWidth     =   10485
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmprod_outras_despesas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   10485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   17
      Top             =   8235
      Width           =   10485
      _ExtentX        =   18494
      _ExtentY        =   714
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   405
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   10485
      _ExtentX        =   18494
      _ExtentY        =   714
      DibPicture      =   "frmprod_outras_despesas.frx":1042
      CaptionDelimiter=   "|"
      CaptionOnCenter =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Icon            =   "frmprod_outras_despesas.frx":81C2
      ShowMaximizeButton=   0   'False
      ShowMinimizeButton=   0   'False
   End
   Begin DrawSuite2022.USTreeView USTreeView1 
      Height          =   5145
      Left            =   210
      TabIndex        =   4
      Top             =   2340
      Width           =   10125
      _ExtentX        =   17859
      _ExtentY        =   9075
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
      Height          =   825
      Left            =   210
      TabIndex        =   7
      Top             =   1500
      Width           =   10125
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
         Width           =   5715
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
         TabIndex        =   13
         Text            =   "0"
         ToolTipText     =   "ID PC."
         Top             =   390
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.CommandButton Cmd_localizar_PC 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7800
         Picture         =   "frmprod_outras_despesas.frx":9214
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Localizar despesas."
         Top             =   390
         Width           =   315
      End
      Begin VB.TextBox Txt_valor 
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
         Left            =   8250
         MaxLength       =   20
         TabIndex        =   3
         ToolTipText     =   "Valor."
         Top             =   390
         Width           =   1680
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
         Left            =   4567
         TabIndex        =   15
         Top             =   180
         Width           =   720
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
         TabIndex        =   14
         Top             =   180
         Width           =   510
         WordWrap        =   -1  'True
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
         Left            =   8250
         TabIndex        =   9
         Top             =   180
         Width           =   1680
         WordWrap        =   -1  'True
      End
   End
   Begin VB.TextBox Txt_ID 
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
      Left            =   5970
      TabIndex        =   8
      Text            =   "0"
      Top             =   3450
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Frame Frame7 
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
      Height          =   585
      Left            =   210
      TabIndex        =   10
      Top             =   7500
      Width           =   10125
      Begin VB.TextBox Txt_valor_total 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   8310
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Valor total."
         Top             =   180
         Width           =   1620
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Valor total :"
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
         Left            =   7230
         TabIndex        =   11
         Top             =   180
         Width           =   990
         WordWrap        =   -1  'True
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   0
      TabIndex        =   12
      Top             =   420
      Width           =   10485
      _ExtentX        =   18494
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
         Img1            =   "frmprod_outras_despesas.frx":9316
         Count           =   1
      End
   End
   Begin MSComctlLib.ListView Lista_RM 
      Height          =   1560
      Left            =   210
      TabIndex        =   5
      Top             =   5910
      Visible         =   0   'False
      Width           =   10125
      _ExtentX        =   17859
      _ExtentY        =   2752
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   16777215
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "T"
         Text            =   "Nº requisição"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "T"
         Text            =   "Cód. interno"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Descrição"
         Object.Width           =   9763
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "Un."
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Object.Tag             =   "V"
         Text            =   "Vlr. total"
         Object.Width           =   2117
      EndProperty
   End
End
Attribute VB_Name = "frmprod_outras_despesas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Novo_prod_outras As Boolean 'OK

Private Sub Cmd_localizar_PC_Click()
On Error GoTo tratar_erro

Plano_contas_produtos = False
Plano_contas_familias = False
Plano_centro_de_custo = False
Plano_instituicao = False
Plano_opcoesgerais = False
Plano_Faturamento = False
Financeiro_Contas_Pagar = False
Financeiro_Contas_Pagas = False
Financeiro_Contas_Receber = False
Financeiro_Contas_Recebidas = False
Plano_PCP = True
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
If frmprod.txtDtValidacao_custo <> "" Then
    USMsgBox ("Não é permitido salvar esta despesa, pois o resultado da ordem já foi validado."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

Acao = "salvar"
If Txt_ID_PC = 0 Then
    NomeCampo = "a conta contábil"
    ProcVerificaAcao
    Cmd_localizar_PC.SetFocus
    Exit Sub
End If
valor = IIf(Txt_valor = "", 0, Txt_valor)
If valor <= 0 Then
    NomeCampo = "o valor"
    ProcVerificaAcao
    Txt_valor.SetFocus
    Exit Sub
End If

'Verifica se já existe o plano cadastrado
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select ID from Producao_outras_despesas where Ordem = " & frmprod.txtof & " and ID_PC = " & Txt_ID_PC & " and ID <> " & Txt_ID, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    USMsgBox ("Esta despesa já foi cadastrada, favor alterar."), vbExclamation, "CAPRIND v5.0"
    TBAbrir.Close
    Exit Sub
End If
TBAbrir.Close

Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from Producao_outras_despesas where ID = " & Txt_ID, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then
    TBGravar.AddNew
Else
    If TBGravar!valor <> valor Then
        'Verifica se tem RM vinculada a essa conta contábil e não deixa alterar o valor para menor que o valor da RM
        Set TBMateriaprima = CreateObject("adodb.recordset")
        TBMateriaprima.Open "Select SUM(EM.VlrTotal) AS Valor1 from (Requisicao_materiais RM INNER JOIN Requisicao_materiais_lista RML ON RML.IDrequisicao = RM.ID) INNER JOIN Estoque_movimentacao EM ON EM.Desenho = RML.Desenho and EM.Documento = RM.requisicao where RML.Ordem = " & frmprod.txtof & " and RML.ID_PC = " & IDlista, Conexao, adOpenKeyset, adLockOptimistic
        If TBMateriaprima.EOF = False Then
            If valor < IIf(IsNull(TBMateriaprima!Valor1), 0, TBMateriaprima!Valor1) Then
                USMsgBox ("O valor não pode ser menor que " & Format(TBMateriaprima!Valor1, "###,##0.00") & ", pois esta despesa está vinculada a requisição de materiais do estoque."), vbExclamation, "CAPRIND v5.0"
                TBMateriaprima.Close
                Exit Sub
            End If
        End If
        TBMateriaprima.Close
    End If
End If
TBGravar!Ordem = frmprod.txtof
TBGravar!ID_PC = Txt_ID_PC
TBGravar!valor = Txt_valor
TBGravar.Update
Txt_ID = TBGravar!ID
TBGravar.Close
If Novo_prod_outras = True Then
    USMsgBox ("Nova despesa cadastrada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Nova despesa"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar despesa"
End If
ProcVerifPC
'==================================
Modulo = Formulario
ID_documento = Txt_ID
Documento = "Ordem: " & frmprod.txtof
Documento1 = "Código: " & Txt_codigo_PC & " - Descrição: " & Txt_descricao_PC
ProcGravaEvento
'==================================
Novo_prod_outras = False
frmprod.ProcAbrirRe
        
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15225, 7, True
Formulario = "PCP/Gerenciamento de ordem"
ProcLimpaVariaveisPrincipais
ProcVerifPC

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCampos()
On Error GoTo tratar_erro

Txt_ID = 0
Txt_ID_PC = 0
Txt_codigo_PC = ""
Txt_descricao_PC = ""
Txt_valor = ""
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
Set TBFamilia = CreateObject("adodb.recordset")
TBFamilia.Open "Select ID_PC, Sum(valor) as Valor from Producao_outras_despesas where Ordem = " & frmprod.txtof & " and ID_PC is not null and ID_PC <> 0 Group by ID_PC", Conexao, adOpenKeyset, adLockOptimistic
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
Txt_valor_total = Format(valor, "###,##0.00")

NovoValor = Replace(valor, ",", ".")
Conexao.Execute "Update Producao Set CTOutras = " & NovoValor & " where Ordem = " & frmprod.txtof

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
    USMsgBox ("Informe a despesa antes de excluir."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If USMsgBox("Deseja realmente excluir esta despesa?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    If frmprod.txtDtValidacao_custo <> "" Then
        USMsgBox ("Não é permitido excluir esta despesa, pois o resultado da ordem já foi validado."), vbExclamation, "CAPRIND v5.0"
        Exit Sub
    End If
    Conexao.Execute "DELETE from Producao_outras_despesas where ID = " & Txt_ID
    USMsgBox ("Despesa excluída com sucesso."), vbInformation, "CAPRIND v5.0"
    '====================================
    Modulo = Formulario
    Evento = "Excluir despesa"
    ID_documento = Txt_ID
    Documento = "Ordem: " & frmprod.txtof
    Documento1 = "Código: " & Txt_codigo_PC & " - Descrição: " & Txt_descricao_PC
    ProcGravaEvento
    '===================================
    ProcLimpaCampos
    ProcVerifPC
    Novo_prod_outras = False
    Frame2.Enabled = False
    frmprod.ProcAbrirRe
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
ProcLimpaCampos
Novo_prod_outras = True
Frame2.Enabled = True
Cmd_localizar_PC_Click

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

If Novo_prod_outras = True Then
    If USMsgBox("A despesa ainda não foi salva, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvar
        If Novo_prod_outras = True Then Exit Sub Else Unload Me
    End If
End If
Novo_prod_outras = False
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
TBAbrir.Open "Select * from Producao_outras_despesas where Ordem = " & frmprod.txtof & " and ID_PC = " & IDlista, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    Txt_ID = TBAbrir!ID
    Txt_ID_PC = IIf(IsNull(TBAbrir!ID_PC), 0, TBAbrir!ID_PC)
    Set TBFamilia = CreateObject("adodb.recordset")
    TBFamilia.Open "Select * FROM tbl_familia where int_codfamilia = " & Txt_ID_PC, Conexao, adOpenKeyset, adLockOptimistic
    If TBFamilia.EOF = False Then
        Txt_codigo_PC = IIf(IsNull(TBFamilia!CODIGO), "", TBFamilia!CODIGO)
        Txt_descricao_PC = IIf(IsNull(TBFamilia!Txt_descricao), "", TBFamilia!Txt_descricao)
    End If
    Txt_valor = Format(TBAbrir!valor, "###,##0.00")
    Novo_prod_outras = False
    Frame2.Enabled = True
    
    valor = 0
    Lista_RM.ListItems.Clear
    Set TBMateriaprima = CreateObject("adodb.recordset")
    TBMateriaprima.Open "Select RM.requisicao, RML.Desenho, RML.Descricao, RML.UN, SUM(EM.VlrTotal) AS Valor from (Requisicao_materiais RM INNER JOIN Requisicao_materiais_lista RML ON RML.IDrequisicao = RM.ID) INNER JOIN Estoque_movimentacao EM ON EM.Desenho = RML.Desenho and EM.Documento = RM.requisicao where RML.Ordem = " & frmprod.txtof & " and RML.ID_PC = " & IDlista & " group by RM.requisicao, RML.Desenho, RML.Descricao, RML.UN", Conexao, adOpenKeyset, adLockOptimistic
    If TBMateriaprima.EOF = False Then
        USTreeView1.Height = 6015
        With USToolBar1
            .ButtonState(3) = 5
            .Refresh
        End With
        Lista_RM.Visible = True
        Do While TBMateriaprima.EOF = False
            With Lista_RM.ListItems
                .Add , , TBMateriaprima!requisicao
                .Item(.Count).SubItems(1) = IIf(IsNull(TBMateriaprima!Desenho), "", TBMateriaprima!Desenho)
                .Item(.Count).SubItems(2) = IIf(IsNull(TBMateriaprima!Descricao), "", TBMateriaprima!Descricao)
                .Item(.Count).SubItems(3) = IIf(IsNull(TBMateriaprima!Un), "", TBMateriaprima!Un)
                .Item(.Count).SubItems(4) = IIf(IsNull(TBMateriaprima!valor), "", Format(TBMateriaprima!valor, "###,##0.00"))
            End With
            TBMateriaprima.MoveNext
        Loop
    Else
        USTreeView1.Height = 10125
        With USToolBar1
            .ButtonState(3) = 1
            .Refresh
        End With
        Lista_RM.Visible = False
    End If
    TBMateriaprima.Close
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_Valor_Change()
On Error GoTo tratar_erro

If Txt_valor <> "" Then
    VerifNumero = Txt_valor
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_valor = ""
        Txt_valor.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_valor_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus Txt_valor

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_Valor_LostFocus()
On Error GoTo tratar_erro

Txt_valor = Format(Txt_valor, "###,##0.00")
    
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
