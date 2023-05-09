VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmproj_conjunto_criar_versao 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Criar versão"
   ClientHeight    =   3105
   ClientLeft      =   0
   ClientTop       =   -30
   ClientWidth     =   2835
   ClipControls    =   0   'False
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
   Icon            =   "frmproj_conjunto_criar_versao.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   2835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   435
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   767
      DibPicture      =   "frmproj_conjunto_criar_versao.frx":000C
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
      ShowMaximize    =   0   'False
      ShowMinimize    =   0   'False
   End
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   4
      Top             =   2700
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   714
   End
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   2040
      Top             =   210
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frmproj_conjunto_criar_versao.frx":02E9
      Count           =   1
   End
   Begin VB.Frame Frame1 
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
      Height          =   915
      Left            =   360
      TabIndex        =   1
      Top             =   1590
      Width           =   1965
      Begin VB.ComboBox Cmb_versao 
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
         ItemData        =   "frmproj_conjunto_criar_versao.frx":2354
         Left            =   360
         List            =   "frmproj_conjunto_criar_versao.frx":2356
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Versão da fase."
         Top             =   390
         Width           =   1275
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Escolha a versão"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   390
         TabIndex        =   2
         Top             =   180
         Width           =   1365
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   0
      TabIndex        =   3
      Top             =   450
      Width           =   3495
      _ExtentX        =   6165
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
      ButtonCaption1  =   "Criar versão"
      ButtonEnabled1  =   0   'False
      ButtonIconSize1 =   32
      ButtonToolTipText1=   "Criar versão (F7)"
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
      ButtonWidth1    =   67
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
      ButtonLeft2     =   71
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
      ButtonLeft3     =   75
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
      ButtonLeft4     =   113
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
      ButtonLeft5     =   141
      ButtonTop5      =   2
      ButtonWidth5    =   24
      ButtonHeight5   =   24
      ButtonUseMaskColor5=   0   'False
   End
End
Attribute VB_Name = "frmproj_conjunto_criar_versao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ProcCriarVersao()
On Error GoTo tratar_erro

'If FunVerifValidacaoRegistro("copiar", frmProcessos.txtDtValidacao, "processo", "a fase", True) = False Then Exit Sub

Acao = "Criar versão"
If Cmb_versao = "" Then
    NomeCampo = "versão"
    ProcVerificaAcao
    Cmb_versao.SetFocus
    Exit Sub
End If

If Engenharia_Conjuntos = True Then
    With frmproj_conjunto
        ProcCopiarEstrutura .Txt_cod_produto, .Txt_cod_produto, .cmbVersao_pesquisar, Cmb_versao
        USMsgBox ("Versão " & Cmb_versao & " do conjunto criada com sucesso."), vbExclamation, "CAPRIND v5.0"
        '==================================
        Modulo = "Engenharia/Conjuntos"
        Evento = "Criar versão"
        ID_documento = .Txt_cod_produto
        Documento = "Cód. interno: " & .txtdesenhoproduto.Text & " - Versão: " & Cmb_versao
        Documento1 = ""
        ProcGravaEvento
        '==================================
        .ProcCarregaVersao Cmb_versao
    End With
End If

If Formulario = "Engenharia/Estrutura/Detalhada" Then
    With frmproj_produto_estrutura
        ProcCopiarEstrutura .IDProduto, .IDProduto, .VersaoEstrutura, Cmb_versao
        Set TBProduto = CreateObject("adodb.recordset")
        TBProduto.Open "Select Desenho from projproduto where codproduto = " & .IDProduto, Conexao, adOpenKeyset, adLockOptimistic
        If TBProduto.EOF = False Then
            USMsgBox ("Versão " & Cmb_versao & " da estrutura criada com sucesso."), vbExclamation, "CAPRIND v5.0"
            '==================================
            Modulo = "Engenharia/Estrutura"
            Evento = "Criar versão"
            ID_documento = .IDProduto
            Documento = "Cód. interno: " & TBProduto!Desenho & " - Versão: " & Cmb_versao
            Documento1 = ""
            ProcGravaEvento
            '==================================
        End If
    End With
End If

If Formulario = "Engenharia/Estrutura/Resumida" Then
    With frmproj_produto_estrutura_Resumida
        ProcCopiarEstrutura .IDProduto, .IDProduto, .VersaoEstrutura, Cmb_versao
        Set TBProduto = CreateObject("adodb.recordset")
        TBProduto.Open "Select Desenho from projproduto where codproduto = " & .IDProduto, Conexao, adOpenKeyset, adLockOptimistic
        If TBProduto.EOF = False Then
            USMsgBox ("Versão " & Cmb_versao & " da estrutura criada com sucesso."), vbExclamation, "CAPRIND v5.0"
            '==================================
            Modulo = "Engenharia/Estrutura"
            Evento = "Criar versão"
            ID_documento = .IDProduto
            Documento = "Cód. interno: " & TBProduto!Desenho & " - Versão: " & Cmb_versao
            Documento1 = ""
            ProcGravaEvento
            '==================================
        End If
    End With
End If
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 2805, 5, True

If Engenharia_Conjuntos = True Then
IDlista = frmproj_conjunto.Txt_cod_produto
End If

If Formulario = "Engenharia/Estrutura/Detalhada" Then
IDlista = frmproj_produto_estrutura.IDProduto
End If

If Formulario = "Engenharia/Estrutura/Resumida" Then
IDlista = frmproj_produto_estrutura_Resumida.IDProduto
End If

With Cmb_versao
    .Clear
    If FunVerifVersaoCriadaEstrutura("A", IDlista) = False Then .AddItem "A"
    If FunVerifVersaoCriadaEstrutura("B", IDlista) = False Then .AddItem "B"
    If FunVerifVersaoCriadaEstrutura("C", IDlista) = False Then .AddItem "C"
    If FunVerifVersaoCriadaEstrutura("D", IDlista) = False Then .AddItem "D"
    If FunVerifVersaoCriadaEstrutura("E", IDlista) = False Then .AddItem "E"
    If FunVerifVersaoCriadaEstrutura("F", IDlista) = False Then .AddItem "F"
    If FunVerifVersaoCriadaEstrutura("G", IDlista) = False Then .AddItem "G"
    If FunVerifVersaoCriadaEstrutura("H", IDlista) = False Then .AddItem "H"
    If FunVerifVersaoCriadaEstrutura("I", IDlista) = False Then .AddItem "I"
    If FunVerifVersaoCriadaEstrutura("J", IDlista) = False Then .AddItem "J"
    If FunVerifVersaoCriadaEstrutura("K", IDlista) = False Then .AddItem "K"
    If FunVerifVersaoCriadaEstrutura("L", IDlista) = False Then .AddItem "L"
    If FunVerifVersaoCriadaEstrutura("M", IDlista) = False Then .AddItem "M"
    If FunVerifVersaoCriadaEstrutura("N", IDlista) = False Then .AddItem "N"
    If FunVerifVersaoCriadaEstrutura("O", IDlista) = False Then .AddItem "O"
    If FunVerifVersaoCriadaEstrutura("P", IDlista) = False Then .AddItem "P"
    If FunVerifVersaoCriadaEstrutura("Q", IDlista) = False Then .AddItem "Q"
    If FunVerifVersaoCriadaEstrutura("R", IDlista) = False Then .AddItem "R"
    If FunVerifVersaoCriadaEstrutura("S", IDlista) = False Then .AddItem "S"
    If FunVerifVersaoCriadaEstrutura("T", IDlista) = False Then .AddItem "T"
    If FunVerifVersaoCriadaEstrutura("U", IDlista) = False Then .AddItem "U"
    If FunVerifVersaoCriadaEstrutura("V", IDlista) = False Then .AddItem "V"
    If FunVerifVersaoCriadaEstrutura("W", IDlista) = False Then .AddItem "W"
    If FunVerifVersaoCriadaEstrutura("X", IDlista) = False Then .AddItem "X"
    If FunVerifVersaoCriadaEstrutura("Y", IDlista) = False Then .AddItem "Y"
    If FunVerifVersaoCriadaEstrutura("Z", IDlista) = False Then .AddItem "Z"
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyF7: ProcCriarVersao
    Case vbKeyEscape: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcCriarVersao
    'Case 3: ProcAjuda
    Case 4: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
