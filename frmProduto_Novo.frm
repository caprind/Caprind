VERSION 5.00
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Begin VB.Form frmProduto_Novo 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Importação XML | Novo item"
   ClientHeight    =   4425
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8100
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProduto_Novo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   8100
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   855
      Left            =   150
      TabIndex        =   18
      Top             =   2010
      Width           =   7725
      Begin VB.CheckBox chkInspecao 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Inspeção recebimento?"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   390
         Left            =   5670
         TabIndex        =   21
         Top             =   300
         Value           =   1  'Checked
         Width           =   1995
      End
      Begin VB.ComboBox cmbClassificacao_produto 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000080&
         Height          =   360
         ItemData        =   "frmProduto_Novo.frx":000C
         Left            =   150
         List            =   "frmProduto_Novo.frx":000E
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   19
         ToolTipText     =   "Classificação do produto (tipo do item)."
         Top             =   360
         Width           =   5415
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Classificação (Bloco K) | Tipo do item"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   71
         Left            =   1552
         TabIndex        =   20
         Top             =   180
         Width           =   2610
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tipo"
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
      Height          =   585
      Left            =   150
      TabIndex        =   11
      Top             =   1440
      Width           =   7725
      Begin VB.CheckBox chkEstoque 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Movimenta estoque"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   390
         Left            =   5880
         TabIndex        =   17
         Top             =   150
         Value           =   1  'Checked
         Width           =   1725
      End
      Begin VB.OptionButton Opt1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Produto"
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
         Height          =   240
         Left            =   150
         TabIndex        =   16
         Top             =   240
         Width           =   885
      End
      Begin VB.OptionButton Opt2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Subconjunto"
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
         Height          =   240
         Left            =   1050
         TabIndex        =   15
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton opt3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Componente"
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
         Height          =   240
         Left            =   2310
         TabIndex        =   14
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton Opt4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Matéria-prima"
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
         Height          =   240
         Left            =   3540
         TabIndex        =   13
         Top             =   240
         Width           =   1305
      End
      Begin VB.OptionButton Opt5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Outros"
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
         Height          =   240
         Left            =   4890
         TabIndex        =   12
         Top             =   240
         Width           =   825
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Local de armazenamento"
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
      Height          =   735
      Left            =   2820
      TabIndex        =   9
      Top             =   660
      Width           =   5055
      Begin VB.ComboBox cmbLocal_armaz 
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
         Left            =   150
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   10
         ToolTipText     =   "Local de armazenamento."
         Top             =   270
         Width           =   4770
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Gerar cód. interno"
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
      Height          =   735
      Left            =   150
      TabIndex        =   5
      Top             =   660
      Width           =   2655
      Begin VB.OptionButton Optautomatico 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Automático"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   300
         TabIndex        =   7
         Top             =   360
         Value           =   -1  'True
         Width           =   1305
      End
      Begin VB.OptionButton Optmanual 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Manual"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   1590
         TabIndex        =   6
         Top             =   360
         Width           =   945
      End
   End
   Begin VB.Frame Frame4 
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
      Height          =   945
      Left            =   150
      TabIndex        =   2
      Top             =   2850
      Width           =   7725
      Begin VB.ComboBox cmbfamilia 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   210
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "Família."
         Top             =   450
         Width           =   5805
      End
      Begin DrawSuite2022.USButton btnOK 
         Height          =   705
         Left            =   6210
         TabIndex        =   8
         Top             =   150
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   1244
         DibPicture      =   "frmProduto_Novo.frx":0010
         Caption         =   "Salvar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
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
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Família"
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
         Left            =   2865
         TabIndex        =   4
         Top             =   240
         Width           =   600
      End
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   465
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8100
      _ExtentX        =   14288
      _ExtentY        =   820
      DibPicture      =   "frmProduto_Novo.frx":8A15
      CaptionDelimiter=   "|"
      CaptionOnCenter =   -1  'True
      EnableMaximizeButton=   0   'False
      EnableMinimizeButton=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowControlBox  =   0   'False
   End
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   4020
      Width           =   8100
      _ExtentX        =   14288
      _ExtentY        =   714
   End
End
Attribute VB_Name = "frmProduto_Novo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnok_Click()
On Error GoTo tratar_erro

    Familia = cmbfamilia.Text
    LocalArmazenamento = cmbLocal_armaz.Text
    CodManual = Optmanual.Value
    
    If Familia = "" Or LocalArmazenamento = "" Then
    USMsgBox "Escolha a familia e o local de armazenamento", vbInformation, "CAPRIND V5.0"
        Exit Sub
    End If
    
    If ID_Tipo = Null Then
    USMsgBox "Escolha a classificação do item no bloco K", vbInformation, "CAPRIND V5.0"
        Exit Sub
    End If
    
    Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkEstoque_Click()
On Error GoTo tratar_erro
   
    If chkEstoque.Value = 1 Then
        Estoque = True
    Else
       Estoque = False
    End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkInspecao_Click()
On Error GoTo tratar_erro
    
    If Chk_insp_recebimento.Value = 1 Then
        Inspecao_recebimento = True
    Else
       Inspecao_recebimento = False
    End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbClassificacao_produto_Change()
On Error GoTo tratar_erro
    
    If cmbClassificacao_produto <> "" Then
        ID_Tipo = cmbClassificacao_produto.ItemData(cmbClassificacao_produto.ListIndex)
    Else
        ID_Tipo = Null
    End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbClassificacao_produto_Click()
On Error GoTo tratar_erro
    
    If cmbClassificacao_produto <> "" Then
        ID_Tipo = cmbClassificacao_produto.ItemData(cmbClassificacao_produto.ListIndex)
    Else
        ID_Tipo = Null
    End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

cmbfamilia.Clear

ProcCarregaComboFamilia cmbfamilia, "familia <> 'Null' and Compras = 'True'", False
ProcCarregaComboLA cmbLocal_armaz, False, False

With cmbClassificacao_produto
    .Clear
    Set TBCarregarCombo = CreateObject("adodb.recordset")
    TBCarregarCombo.Open "Select * from Projproduto_Tipo order by codigo", Conexao, adOpenKeyset, adLockOptimistic
    If TBCarregarCombo.EOF = False Then
        Do While TBCarregarCombo.EOF = False
            .AddItem TBCarregarCombo!CODIGO & " - " & TBCarregarCombo!Descricao
            .ItemData(.NewIndex) = TBCarregarCombo!ID
            TBCarregarCombo.MoveNext
        Loop
    End If
    TBCarregarCombo.Close
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub Opt1_Click()
On Error GoTo tratar_erro

If Opt1.Value = True Then
SubTipoItem = 1
End If

If Opt2.Value = True Then
SubTipoItem = 2
End If

If opt3.Value = True Then
TipoItem = 3
End If

If Opt4.Value = True Then
SubTipoItem = 0
End If

If Opt5.Value = True Then
SubTipoItem = 4
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt2_Click()
On Error GoTo tratar_erro

If Opt1.Value = True Then
SubTipoItem = 1
End If

If Opt2.Value = True Then
SubTipoItem = 2
End If

If opt3.Value = True Then
TipoItem = 3
End If

If Opt4.Value = True Then
SubTipoItem = 0
End If

If Opt5.Value = True Then
SubTipoItem = 4
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub opt3_Click()
On Error GoTo tratar_erro

If Opt1.Value = True Then
SubTipoItem = 1
End If

If Opt2.Value = True Then
SubTipoItem = 2
End If

If opt3.Value = True Then
TipoItem = 3
End If

If Opt4.Value = True Then
SubTipoItem = 0
End If

If Opt5.Value = True Then
SubTipoItem = 4
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub OPT4_Click()
On Error GoTo tratar_erro

If Opt1.Value = True Then
SubTipoItem = 1
End If

If Opt2.Value = True Then
SubTipoItem = 2
End If

If opt3.Value = True Then
TipoItem = 3
End If

If Opt4.Value = True Then
SubTipoItem = 0
End If

If Opt5.Value = True Then
SubTipoItem = 4
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt5_Click()
On Error GoTo tratar_erro

If Opt1.Value = True Then
SubTipoItem = 1
End If

If Opt2.Value = True Then
SubTipoItem = 2
End If

If opt3.Value = True Then
TipoItem = 3
End If

If Opt4.Value = True Then
SubTipoItem = 0
End If

If Opt5.Value = True Then
SubTipoItem = 4
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
