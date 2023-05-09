VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmCQ_sistema_retirar 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Qualidade - Controle de documentos e dados - Retirar"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9345
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
   ScaleHeight     =   4290
   ScaleWidth      =   9345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2475
      Left            =   55
      TabIndex        =   10
      Top             =   990
      Width           =   9270
      Begin VB.TextBox Txt_rev 
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
         Left            =   2550
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Revisão."
         Top             =   375
         Width           =   525
      End
      Begin VB.TextBox Txt_observacao 
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
         Height          =   675
         Left            =   150
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         ToolTipText     =   "Observações."
         Top             =   1644
         Width           =   8925
      End
      Begin VB.TextBox Txt_codigo 
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
         MaxLength       =   50
         TabIndex        =   0
         TabStop         =   0   'False
         ToolTipText     =   "Código."
         Top             =   375
         Width           =   2355
      End
      Begin VB.TextBox Txt_descricao 
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
         Left            =   3090
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Descrição."
         Top             =   375
         Width           =   5985
      End
      Begin VB.ComboBox Cmb_funcionario 
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
         Left            =   1540
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         ToolTipText     =   "Funcionário."
         Top             =   1007
         Width           =   7565
      End
      Begin MSComCtl2.DTPicker Cmb_data_retirada 
         Height          =   315
         Left            =   210
         TabIndex        =   3
         ToolTipText     =   "Data da retirada."
         Top             =   1007
         Width           =   1335
         _ExtentX        =   2355
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
         Format          =   487456769
         CurrentDate     =   39057
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rev."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2640
         TabIndex        =   17
         Top             =   180
         Width           =   345
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Funcionário*"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   4865
         TabIndex        =   16
         Top             =   810
         Width           =   915
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dt. retirada"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   457
         TabIndex        =   15
         Top             =   810
         Width           =   840
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   5737
         TabIndex        =   14
         Top             =   180
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   7
         Left            =   150
         TabIndex        =   13
         Top             =   3030
         Width           =   45
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   9
         Left            =   1110
         TabIndex        =   12
         Top             =   180
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Observações"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   4140
         TabIndex        =   11
         Top             =   1440
         Width           =   945
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   55
      TabIndex        =   9
      Top             =   0
      Width           =   9270
      _ExtentX        =   16351
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
      ButtonCaption1  =   "Retirar"
      ButtonEnabled1  =   0   'False
      ButtonIconSize1 =   32
      ButtonToolTipText1=   "Retirar (F3)"
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
      ButtonWidth1    =   41
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
      ButtonLeft2     =   45
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
      ButtonLeft3     =   49
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
      ButtonLeft4     =   87
      ButtonTop4      =   2
      ButtonWidth4    =   26
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
      ButtonLeft5     =   115
      ButtonTop5      =   2
      ButtonWidth5    =   24
      ButtonHeight5   =   24
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   7470
         Top             =   90
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmCQ_sistema_retirar.frx":0000
         Count           =   1
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   55
      TabIndex        =   18
      Top             =   3450
      Width           =   9270
      Begin VB.TextBox Txt_qtde_copias 
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   5355
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Text            =   "0"
         ToolTipText     =   "Quantidade de cópias."
         Top             =   375
         Width           =   1215
      End
      Begin VB.TextBox Txt_qtde_retirada 
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   6585
         TabIndex        =   7
         Text            =   "0"
         ToolTipText     =   "Quantidade retirada."
         Top             =   375
         Width           =   1275
      End
      Begin VB.TextBox Txt_qtde_atualizada 
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   7875
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   "0"
         ToolTipText     =   "Quantidade de cópias atualizada."
         Top             =   375
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qtde. cópias"
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
         Index           =   6
         Left            =   5445
         TabIndex        =   22
         Top             =   180
         Width           =   1035
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qtde. retirada*"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   8
         Left            =   6660
         TabIndex        =   21
         Top             =   180
         Width           =   1125
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qtde. atual."
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
         Index           =   10
         Left            =   7995
         TabIndex        =   20
         Top             =   180
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   12
         Left            =   150
         TabIndex        =   19
         Top             =   3030
         Width           =   45
      End
   End
End
Attribute VB_Name = "frmCQ_sistema_retirar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyF3: ProcRetirar
    'Case vbKeyF1: ProcAjuda
    Case vbKeyEscape: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 9270, 5, True
ProcLimpaVariaveisPrincipais
With frmCQ_sistema
    txt_Codigo = .txtCodigo
    txt_rev = .txt_rev
    Txt_descricao = .Txt_descricao
    Cmb_data_retirada = Date
    ProcCarregaComboFuncionario Cmb_funcionario, "Situacao <> 'Afastado' and Situacao <> 'Demitido'", False
    Txt_qtde_copias = .Txt_qtde_copias
    Txt_qtde_atualizada = .Txt_qtde_copias
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_qtde_retirada_Change()
On Error GoTo tratar_erro

If Txt_qtde_retirada <> "" Then
    VerifNumero = Txt_qtde_retirada
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_qtde_retirada = ""
        Txt_qtde_retirada.SetFocus
        Exit Sub
    End If
End If
Qtd_Real = IIf(Txt_qtde_copias = "", 0, Txt_qtde_copias)
Qtd_Retirada = IIf(Txt_qtde_retirada = "", 0, Txt_qtde_retirada)
Txt_qtde_atualizada = Qtd_Real - Qtd_Retirada

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_qtde_retirada_GotFocus()
On Error GoTo tratar_erro

If Txt_qtde_retirada = "0" Then Txt_qtde_retirada = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcRetirar()
On Error GoTo tratar_erro

Acao = "retirar"
If Cmb_funcionario = "" Then
    NomeCampo = "o funcionário"
    ProcVerificaAcao
    Cmb_funcionario.SetFocus
    Exit Sub
End If
QuantEmpenho = IIf(Txt_qtde_retirada = "", 0, Txt_qtde_retirada)
If QuantEmpenho <= 0 Then
    NomeCampo = "a quantidade retirada"
    ProcVerificaAcao
    Txt_qtde_retirada.SetFocus
    Exit Sub
End If
Qtd_Real = Txt_qtde_copias
If Qtd_Real < QuantEmpenho Then
    USMsgBox ("A quantidade retirada não pode ser maior que a quantidade de cópias."), vbExclamation, "CAPRIND v5.0"
    Txt_qtde_retirada.SetFocus
    Exit Sub
End If

With frmCQ_sistema
    Set TBEstoque = CreateObject("adodb.recordset")
    TBEstoque.Open "Select * from CQ_Sistema_Movimentacoes", Conexao, adOpenKeyset, adLockOptimistic
    TBEstoque.AddNew
    TBEstoque!ID_sistema = .txtID
    TBEstoque!Data = Cmb_data_retirada
    TBEstoque!Responsavel = pubUsuario
    TBEstoque!ID_funcionario = Cmb_funcionario.ItemData(Cmb_funcionario.ListIndex)
    TBEstoque!Obs = Txt_observacao
    TBEstoque!Qtde_saida = Txt_qtde_retirada
    TBEstoque.Update
    USMsgBox ("Controle de documentos e dados retirado com sucesso."), vbInformation, "CAPRIND v5.0"
    '==================================
    Modulo = "Qualidade/Controle de documentos e dados"
    Evento = "Retirar"
    ID_documento = TBEstoque!ID
    Documento = "Código: " & txt_Codigo & " - Rev.: " & txt_rev
    Documento1 = "Data: " & Format(Cmb_data_retirada, "dd/mm/yy") & " - Qtde.: " & Txt_qtde_retirada
    ProcGravaEvento
    '==================================
    .ProcAtualizaQtdeCopias .txtID, 0, Txt_qtde_retirada
    Set TBLISTA = CreateObject("adodb.recordset")
    TBLISTA.Open "Select * from CQ_sistema where id = " & .txtID, Conexao, adOpenKeyset, adLockOptimistic
    If TBLISTA.EOF = False Then .ProcPuxaDados
    .ProcCarregaListaMovimentacao (1)
    .ProcCarregaLista (IIf(ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5))))
End With
Unload Me
    
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

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcRetirar
    'Case 3: ProcAjuda
    Case 4: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
