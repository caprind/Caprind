VERSION 5.00
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Begin VB.Form frmProj_Produto_Copiar 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Engenharia - Produtos e serviços - Copiar"
   ClientHeight    =   2775
   ClientLeft      =   1680
   ClientTop       =   1365
   ClientWidth     =   9660
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   9660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
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
      Height          =   885
      Left            =   60
      TabIndex        =   23
      Top             =   990
      Width           =   1785
      Begin VB.OptionButton Optautomatico 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Automático"
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
         Height          =   210
         Left            =   210
         TabIndex        =   0
         Top             =   300
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   210
         TabIndex        =   1
         Top             =   540
         Width           =   945
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aplicação"
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
      Height          =   885
      Left            =   1860
      TabIndex        =   22
      Top             =   990
      Width           =   2235
      Begin VB.CheckBox Chk_vendas 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Vendas"
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
         Left            =   150
         TabIndex        =   2
         Top             =   300
         Width           =   825
      End
      Begin VB.CheckBox Chk_compras 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Compras"
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
         Left            =   1050
         TabIndex        =   4
         Top             =   300
         Width           =   975
      End
      Begin VB.CheckBox Chk_qualidade 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Qualidade"
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
         Left            =   1050
         TabIndex        =   5
         Top             =   540
         Width           =   1035
      End
      Begin VB.CheckBox Chk_PCP 
         BackColor       =   &H00E0E0E0&
         Caption         =   "PCP"
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
         Left            =   150
         TabIndex        =   3
         Top             =   540
         Width           =   825
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
      Height          =   885
      Left            =   4110
      TabIndex        =   21
      Top             =   990
      Width           =   5505
      Begin VB.OptionButton Opt5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Serviço"
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
         Left            =   2820
         TabIndex        =   11
         Top             =   540
         Width           =   825
      End
      Begin VB.OptionButton Opt4 
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
         Height          =   195
         Left            =   1545
         TabIndex        =   10
         Top             =   540
         Width           =   825
      End
      Begin VB.OptionButton Opt1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Produto final"
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
         Left            =   180
         TabIndex        =   6
         Top             =   300
         Width           =   1215
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
         Height          =   195
         Left            =   1545
         TabIndex        =   7
         Top             =   300
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
         Height          =   195
         Left            =   2820
         TabIndex        =   8
         Top             =   300
         Width           =   1215
      End
      Begin VB.OptionButton Opt0 
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
         Height          =   195
         Left            =   180
         TabIndex        =   9
         Top             =   540
         Width           =   1305
      End
   End
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   4110
      Top             =   180
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frmProj_Produto_Copiar.frx":0000
      Count           =   1
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   20
      Top             =   0
      Width           =   9555
      _ExtentX        =   16854
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
      ButtonCaption1  =   "Copiar"
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
      ButtonWidth1    =   39
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
      ButtonLeft2     =   43
      ButtonTop2      =   4
      ButtonWidth2    =   2
      ButtonHeight2   =   54
      ButtonCaption3  =   "Ajuda"
      ButtonEnabled3  =   0   'False
      ButtonIconSize3 =   32
      ButtonToolTipText3=   "Ajuda (F1)"
      ButtonKey3      =   "6"
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
      ButtonLeft3     =   47
      ButtonTop3      =   2
      ButtonWidth3    =   36
      ButtonHeight3   =   21
      ButtonUseMaskColor3=   0   'False
      ButtonCaption4  =   "Sair"
      ButtonEnabled4  =   0   'False
      ButtonIconSize4 =   32
      ButtonToolTipText4=   "Sair (Esc)"
      ButtonKey4      =   "7"
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
      ButtonLeft4     =   85
      ButtonTop4      =   2
      ButtonWidth4    =   26
      ButtonHeight4   =   21
      ButtonUseMaskColor4=   0   'False
      ButtonEnabled5  =   0   'False
      ButtonIconSize5 =   32
      ButtonKey5      =   "8"
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
      ButtonLeft5     =   113
      ButtonTop5      =   2
      ButtonWidth5    =   24
      ButtonHeight5   =   24
      ButtonUseMaskColor5=   0   'False
   End
   Begin VB.Frame Frame2 
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
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   55
      TabIndex        =   15
      Top             =   1890
      Width           =   9555
      Begin VB.ComboBox cmbfamilia 
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
         Left            =   180
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   12
         ToolTipText     =   "Família."
         Top             =   390
         Width           =   4995
      End
      Begin VB.TextBox txtcodref 
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
         Left            =   7290
         MaxLength       =   50
         TabIndex        =   14
         ToolTipText     =   "Código de referência."
         Top             =   390
         Width           =   2085
      End
      Begin VB.TextBox txtcodinterno 
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
         Left            =   5190
         MaxLength       =   50
         TabIndex        =   13
         ToolTipText     =   "Código interno."
         Top             =   390
         Width           =   2085
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Família"
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
         Left            =   2437
         TabIndex        =   19
         Top             =   180
         Width           =   480
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código de referência"
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
         Left            =   7582
         TabIndex        =   18
         Top             =   180
         Width           =   1500
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código interno"
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
         Left            =   5707
         TabIndex        =   17
         Top             =   180
         Width           =   1050
      End
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   4680
      TabIndex        =   16
      Top             =   1320
      Width           =   45
   End
End
Attribute VB_Name = "frmProj_Produto_Copiar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CodReferencia   As String 'OK

Private Sub ProcCarregaFamilia()
On Error GoTo tratar_erro

cmbfamilia.Clear
'Vendas + Compras + PCP + Qualidade
If Chk_vendas.Value = 1 And Chk_compras.Value = 1 And Chk_PCP.Value = 1 And Chk_qualidade.Value = 1 Then
    ProcCarregaComboFamilia cmbfamilia, "familia <> 'Null'", False
    'Vendas + Compras + PCP
ElseIf Chk_vendas.Value = 1 And Chk_compras.Value = 1 And Chk_PCP.Value = 1 And Chk_qualidade.Value = 0 Then
        ProcCarregaComboFamilia cmbfamilia, "familia <> 'Null' and (Vendas = 'True' or Compras = 'True' or Fabricacao = 'True')", False
        'Vendas + Compras + Qualidade
    ElseIf Chk_vendas.Value = 1 And Chk_compras.Value = 1 And Chk_PCP.Value = 0 And Chk_qualidade.Value = 1 Then
            ProcCarregaComboFamilia cmbfamilia, "familia <> 'Null' and (Vendas = 'True' or Compras = 'True' or Qualidade = 'True')", False
            'Vendas + Compras
        ElseIf Chk_vendas.Value = 1 And Chk_compras.Value = 1 And Chk_PCP.Value = 0 And Chk_qualidade.Value = 0 Then
                ProcCarregaComboFamilia cmbfamilia, "familia <> 'Null' and (Vendas = 'True' or Compras = 'True')", False
                'Vendas + PCP + Qualidade
            ElseIf Chk_vendas.Value = 1 And Chk_compras.Value = 0 And Chk_PCP.Value = 1 And Chk_qualidade.Value = 1 Then
                    ProcCarregaComboFamilia cmbfamilia, "familia <> 'Null' and (Vendas = 'True' or Fabricacao = 'True' or Qualidade = 'True')", False
                    'Vendas + Qualidade
                ElseIf Chk_vendas.Value = 1 And Chk_compras.Value = 0 And Chk_PCP.Value = 0 And Chk_qualidade.Value = 1 Then
                        ProcCarregaComboFamilia cmbfamilia, "familia <> 'Null' and (Vendas = 'True' or Qualidade = 'True')", False
                        'Vendas + PCP
                    ElseIf Chk_vendas.Value = 1 And Chk_compras.Value = 0 And Chk_PCP.Value = 1 And Chk_qualidade.Value = 0 Then
                            ProcCarregaComboFamilia cmbfamilia, "familia <> 'Null' and (Vendas = 'True' or Fabricacao = 'True')", False
                            'Compras + PCP + Qualidade
                        ElseIf Chk_vendas.Value = 0 And Chk_compras.Value = 1 And Chk_PCP.Value = 1 And Chk_qualidade.Value = 1 Then
                                ProcCarregaComboFamilia cmbfamilia, "familia <> 'Null' and (Compras = 'True' or Fabricacao = 'True' or Qualidade = 'True')", False
                                'Compras + Qualidade
                            ElseIf Chk_vendas.Value = 0 And Chk_compras.Value = 1 And Chk_PCP.Value = 0 And Chk_qualidade.Value = 1 Then
                                    ProcCarregaComboFamilia cmbfamilia, "familia <> 'Null' and (Compras = 'True' or Qualidade = 'True')", False
                                    'Compras + PCP
                                ElseIf Chk_vendas.Value = 0 And Chk_compras.Value = 1 And Chk_PCP.Value = 1 And Chk_qualidade.Value = 0 Then
                                        ProcCarregaComboFamilia cmbfamilia, "familia <> 'Null' and (Compras = 'True' or Fabricacao = 'True')", False
                                        'PCP + Qualidade
                                    ElseIf Chk_vendas.Value = 0 And Chk_compras.Value = 0 And Chk_PCP.Value = 1 And Chk_qualidade.Value = 1 Then
                                            ProcCarregaComboFamilia cmbfamilia, "familia <> 'Null' and (Fabricacao = 'True' or Qualidade = 'True')", False
                                            'Vendas
                                        ElseIf Chk_vendas.Value = 1 And Chk_compras.Value = 0 And Chk_PCP.Value = 0 And Chk_qualidade.Value = 0 Then
                                                ProcCarregaComboFamilia cmbfamilia, "familia <> 'Null' and Vendas = 'True'", False
                                                'Compras
                                            ElseIf Chk_vendas.Value = 0 And Chk_compras.Value = 1 And Chk_PCP.Value = 0 And Chk_qualidade.Value = 0 Then
                                                    ProcCarregaComboFamilia cmbfamilia, "familia <> 'Null' and Compras = 'True'", False
                                                    'PCP
                                                ElseIf Chk_vendas.Value = 0 And Chk_compras.Value = 0 And Chk_PCP.Value = 1 And Chk_qualidade.Value = 0 Then
                                                        ProcCarregaComboFamilia cmbfamilia, "familia <> 'Null' and Fabricacao = 'True'", False
                                                    ElseIf Chk_vendas.Value = 0 And Chk_compras.Value = 0 And Chk_PCP.Value = 0 And Chk_qualidade.Value = 1 Then
                                                        'Qualidade
                                                        ProcCarregaComboFamilia cmbfamilia, "familia <> 'Null' and Qualidade = 'True'", False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCopiar()
On Error GoTo tratar_erro

Acao = "copiar"
If Chk_vendas.Value = 0 And Chk_compras.Value = 0 And Chk_PCP.Value = 0 And Chk_qualidade.Value = 0 Then
    NomeCampo = "a aplicação"
    ProcVerificaAcao
    Exit Sub
End If
If Opt5.Value = False And Opt4.Value = False And opt3.Value = False And Opt2.Value = False And Opt1.Value = False And Opt0.Value = False Then
    NomeCampo = "o tipo"
    ProcVerificaAcao
    Exit Sub
End If
If cmbfamilia = "" Then
    NomeCampo = "a família"
    ProcVerificaAcao
    cmbfamilia.SetFocus
    Exit Sub
End If
If Optmanual.Value = True And txtCodinterno = "" Then
    NomeCampo = "o código interno"
    ProcVerificaAcao
    txtCodinterno.SetFocus
    Exit Sub
End If
'Verifica se já existe produto cadastrado com esse cód. interno
If Optmanual.Value = True Then
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select * from projproduto where desenho = '" & txtCodinterno & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
        USMsgBox ("Não é permitido cadastrar um novo registro, pois esse código interno já está cadastrado."), vbExclamation, "CAPRIND v5.0"
        txtCodinterno.SetFocus
        TBProduto.Close
        Exit Sub
    End If
    TBProduto.Close
End If

If Opt2.Value = True Or opt3.Value = True Then ProcNovoItem Else ProcNovoProduto
USMsgBox ("Registro copiado com sucesso."), vbInformation, "CAPRIND v5.0"
With frmproj_produto
    '==================================
    Modulo = Formulario
    Evento = "Novo"
    ID_documento = .txtcodproduto
    Documento = "Cód. interno: " & .txtdesenhoproduto
    Documento1 = ""
    ProcGravaEvento
    '==================================
End With
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovoProduto()
On Error GoTo tratar_erro

With frmproj_produto
    Set TBFamilia = CreateObject("adodb.recordset")
    TBFamilia.Open "Select * from projfamilia where familia = '" & cmbfamilia.Text & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBFamilia.EOF = False Then Letra = TBFamilia!Letra
    TBFamilia.Close
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from projproduto where codproduto = " & .txtcodproduto, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        If Optautomatico.Value = True Then
            CompLetra = Len(Letra)
            Set TBComponente = CreateObject("adodb.recordset")
            TBComponente.Open "Select * from projproduto where classe = '" & cmbfamilia.Text & "' and codmanual = 'False' and Right(Desenho, " & CompLetra & ") = '" & Letra & "' and (subtipoitem = 0 or subtipoitem = 1 or subtipoitem = 4 or subtipoitem = 5) order by codproduto desc", Conexao, adOpenKeyset, adLockOptimistic
            If TBComponente.EOF = False Then
                Numero = Left(TBComponente!Desenho, Len(TBComponente!Desenho) - (CompLetra + 1))
                Numero = Numero + 1
                With txtCodinterno
                    Select Case Len(Numero)
                        Case 5: .Text = Numero & "-" & Letra
                        Case 4: .Text = "0" & Numero & "-" & Letra
                        Case 3: .Text = "00" & Numero & "-" & Letra
                        Case 2: .Text = "000" & Numero & "-" & Letra
                        Case 1: .Text = "0000" & Numero & "-" & Letra
                    End Select
                End With
VerifCodigo:
                Set TBFI = CreateObject("adodb.recordset")
                TBFI.Open "Select * from projproduto where desenho = '" & txtCodinterno & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBFI.EOF = False Then
                    Numero = Left(TBFI!Desenho, Len(TBFI!Desenho) - (CompLetra + 1))
                    Numero = Numero + 1
                    With txtCodinterno
                        Select Case Len(Numero)
                            Case 5: .Text = Numero & "-" & Letra
                            Case 4: .Text = "0" & Numero & "-" & Letra
                            Case 3: .Text = "00" & Numero & "-" & Letra
                            Case 2: .Text = "000" & Numero & "-" & Letra
                            Case 1: .Text = "0000" & Numero & "-" & Letra
                        End Select
                    End With
                    GoTo VerifCodigo
                End If
            Else
                txtCodinterno = "00001" & "-" & Letra
            End If
            TBComponente.AddNew
            ProcEnviaDados
            TBComponente.Close
        Else
            Set TBComponente = CreateObject("adodb.recordset")
            TBComponente.Open "Select * from projproduto where classe = '" & cmbfamilia.Text & "' order by codproduto", Conexao, adOpenKeyset, adLockOptimistic
            TBComponente.AddNew
            ProcEnviaDados
            TBComponente.Close
        End If
    End If
    TBAbrir.Close
    
    .ProcLimpaCampos
    Set TBProduto = CreateObject("adodb.recordset")
    If Engenharia_Produtos = True Or Engenharia_Produtos = False And .Chk_vendas.Value = Chk_vendas.Value And .Chk_compras.Value = Chk_compras.Value And .Chk_PCP.Value = Chk_PCP.Value And .Chk_qualidade = Chk_qualidade.Value Then
        TBProduto.Open "Select * from projproduto where desenho = '" & txtCodinterno & "'", Conexao, adOpenKeyset, adLockOptimistic
    Else
        TBProduto.Open "Select * from projproduto where desenho = '" & .Lista.SelectedItem.ListSubItems(1) & "'", Conexao, adOpenKeyset, adLockOptimistic
    End If
    If TBProduto.EOF = False Then
        .Copiar_Produto = True
        .ProcCarregaDados
        .ProcCarregaDadosOutros
        .ProcCarregaDadosImpostos
        .ProcCarregaDadosValoresDesc
        .ProcCarregaDadosForn_clientes
        .ProcCarregaDadosPC
        .ProcCarregaDadosImagem
    End If
    TBProduto.Close
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovoItem()
On Error GoTo tratar_erro

With frmproj_produto
    Set TBFamilia = CreateObject("adodb.recordset")
    TBFamilia.Open "Select * from projfamilia where familia = '" & cmbfamilia.Text & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBFamilia.EOF = False Then Letra = TBFamilia!Letra
    TBFamilia.Close
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from projproduto where codproduto = " & .txtcodproduto, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        If Optautomatico.Value = True Then
            CompLetra = Len(Letra)
            Set TBComponente = CreateObject("adodb.recordset")
            TBComponente.Open "Select * from projproduto where classe = '" & cmbfamilia.Text & "' and Left(Desenho, " & CompLetra & ") = '" & Letra & "' and codmanual = 'False' and (subtipoitem = 2 or subtipoitem = 3) order by codproduto desc", Conexao, adOpenKeyset, adLockOptimistic
            If TBComponente.EOF = False Then
                Numero = Right(TBComponente!Desenho, Len(TBComponente!Desenho) - (CompLetra + 1))
                Numero = Numero + 1
                With txtCodinterno
                    Select Case Len(Numero)
                        Case 5: .Text = Letra & "-" & Numero
                        Case 4: .Text = Letra & "-0" & Numero
                        Case 3: .Text = Letra & "-00" & Numero
                        Case 2: .Text = Letra & "-000" & Numero
                        Case 1: .Text = Letra & "-0000" & Numero
                    End Select
                End With
VerifCodigo:
                Set TBFIltro = CreateObject("adodb.recordset")
                TBFIltro.Open "Select * from projproduto where desenho = '" & txtCodinterno & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBFIltro.EOF = False Then
                    Numero = Right(txtCodinterno, Len(txtCodinterno) - (CompLetra + 1))
                    Numero = Numero + 1
                    With txtCodinterno
                        Select Case Len(Numero)
                            Case 5: .Text = Letra & "-" & Numero
                            Case 4: .Text = Letra & "-0" & Numero
                            Case 3: .Text = Letra & "-00" & Numero
                            Case 2: .Text = Letra & "-000" & Numero
                            Case 1: .Text = Letra & "-0000" & Numero
                        End Select
                    End With
                    GoTo VerifCodigo
                End If
            Else
                txtCodinterno = Letra & "-00001"
            End If
            TBComponente.AddNew
            ProcEnviaDados
            TBComponente.Close
        Else
            Set TBComponente = CreateObject("adodb.recordset")
            TBComponente.Open "Select * from projproduto where classe = '" & cmbfamilia.Text & "' order by codproduto", Conexao, adOpenKeyset, adLockOptimistic
            TBComponente.AddNew
            ProcEnviaDados
            TBComponente.Close
        End If
    End If
    TBAbrir.Close
    
    .ProcLimpaCampos
    Set TBProduto = CreateObject("adodb.recordset")
    If Engenharia_Produtos = True Or Engenharia_Produtos = False And .Chk_vendas.Value = Chk_vendas.Value And .Chk_compras.Value = Chk_compras.Value And .Chk_PCP.Value = Chk_PCP.Value And .Chk_qualidade = Chk_qualidade.Value Then
        TBProduto.Open "Select * from projproduto where desenho = '" & txtCodinterno & "'", Conexao, adOpenKeyset, adLockOptimistic
    Else
        TBProduto.Open "Select * from projproduto where desenho = '" & .Lista.SelectedItem.ListSubItems(1) & "'", Conexao, adOpenKeyset, adLockOptimistic
    End If
    If TBProduto.EOF = False Then
        .Copiar_Produto = True
        .ProcCarregaDados
    End If
    TBProduto.Close
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviaDados()
On Error GoTo tratar_erro

With frmproj_produto
    If Optmanual.Value = True Then TBComponente!CodManual = True Else TBComponente!CodManual = False
    If Chk_vendas.Value = 1 Then TBComponente!Vendas = True Else TBComponente!Vendas = False
    If Chk_compras.Value = 1 Then TBComponente!Compras = True Else TBComponente!Compras = False
    If Chk_PCP.Value = 1 Then TBComponente!Producao = True Else TBComponente!Producao = False
    If Chk_qualidade.Value = 1 Then TBComponente!Qualidade = True Else TBComponente!Qualidade = False
    If Opt5.Value = True Then
        TBComponente!SubTipoItem = 5
        TBComponente!Tipo = "S"
    Else
        TBComponente!Tipo = "P"
        If Opt4.Value = True Then TBComponente!SubTipoItem = 4
        If opt3.Value = True Then TBComponente!SubTipoItem = 3
        If Opt2.Value = True Then TBComponente!SubTipoItem = 2
        If Opt1.Value = True Then TBComponente!SubTipoItem = 1
        If Opt0.Value = True Then TBComponente!SubTipoItem = 0
    End If
    TBComponente!Classe = cmbfamilia
    TBComponente!Desenho = txtCodinterno
    TBComponente!Data = Date
    TBComponente!Descricao = TBAbrir!Descricao
    TBComponente!descricaotecnica = TBAbrir!descricaotecnica
    TBComponente!Unidade = TBAbrir!Unidade
    TBComponente!Unidade_com = TBAbrir!Unidade_com
    TBComponente!RevDesenho = TBAbrir!RevDesenho
    TBComponente!Leadtime = TBAbrir!Leadtime
    TBComponente!Responsavel = pubUsuario
    
    'Copia dados outros
    TBComponente!peso_metro = TBAbrir!peso_metro
    If TBAbrir!Un_Kg <> "" Then TBComponente!Un_Kg = TBAbrir!Un_Kg
    If TBAbrir!PBruto <> "" Then TBComponente!PBruto = TBAbrir!PBruto
    If TBAbrir!PLiquido <> "" Then TBComponente!PLiquido = TBAbrir!PLiquido
    TBComponente!nacional = TBAbrir!nacional
    TBComponente!importacao = TBAbrir!importacao
    TBComponente!exportacao = TBAbrir!exportacao
    TBComponente!Estoque_minimo = TBAbrir!Estoque_minimo
    TBComponente!Estoque = TBAbrir!Estoque
    TBComponente!Insp_recebimento = TBAbrir!Insp_recebimento
    TBComponente!Submetido = TBAbrir!Submetido
    TBComponente!PPAP_Rev = TBAbrir!PPAP_Rev
    TBComponente!PPAP_Datarev = TBAbrir!PPAP_Datarev
    TBComponente!FINAME = TBAbrir!FINAME
    TBComponente!Cor = TBAbrir!Cor
    TBComponente!Comprimento = TBAbrir!Comprimento
    TBComponente!Largura = TBAbrir!Largura
    TBComponente!Espessura = TBAbrir!Espessura
    TBComponente!Dureza = TBAbrir!Dureza
    TBComponente!Skip_lote = TBAbrir!Skip_lote
    TBComponente!qtde_LoteMinimo = TBAbrir!qtde_LoteMinimo
    TBComponente!Observacoes = TBAbrir!Observacoes
    TBComponente!Cod_servico = TBAbrir!Cod_servico
    TBComponente!Cod_servico_NFSE = TBAbrir!Cod_servico_NFSE
    TBComponente!ID_Genero = TBAbrir!ID_Genero
    TBComponente!GTIN = TBAbrir!GTIN
    TBComponente!ID_Tipo = TBAbrir!ID_Tipo
    
    'Copia dados impostos
    TBComponente!ID_CF = TBAbrir!ID_CF
    TBComponente!ID_CFOP = TBAbrir!ID_CFOP
    TBComponente!ID_CFOP1 = TBAbrir!ID_CFOP1
    TBComponente!Servico_cliente = TBAbrir!Servico_cliente
    TBComponente!INSS = TBAbrir!INSS
    TBComponente!Credita_ICMS = TBAbrir!Credita_ICMS
    
    'Copia dados valores/descontos
    If TBAbrir!MLucro <> "" Then TBComponente!MLucro = TBAbrir!MLucro
    If TBAbrir!PCusto <> "" Then TBComponente!PCusto = TBAbrir!PCusto
    If TBAbrir!PConsumo <> "" Then TBComponente!PConsumo = TBAbrir!PConsumo
    If TBAbrir!PRevenda <> "" Then TBComponente!PRevenda = TBAbrir!PRevenda
    TBComponente!Valor_bloqueado = TBAbrir!Valor_bloqueado
    
    'Conta contábil e CC
    Set TBFamilia = CreateObject("adodb.recordset")
    TBFamilia.Open "Select ID_PC, ID_PC1, ID_CC from Projfamilia where Familia = '" & cmbfamilia & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBFamilia.EOF = False Then
        TBComponente!ID_PC = TBFamilia!ID_PC
        TBComponente!ID_PC1 = TBFamilia!ID_PC1
        TBComponente!ID_CC = TBFamilia!ID_CC
    End If
    TBFamilia.Close
        
    TBComponente.Update
    ProcCopiarEstrutura frmproj_produto.txtcodproduto, TBComponente!Codproduto, "", ""
    
    Codproduto = TBComponente!Codproduto
    If Chk_vendas.Value = 1 Then
        Set TBFI = CreateObject("adodb.recordset")
        TBFI.Open "Select * from Projproduto_clientes where codproduto = " & .txtcodproduto, Conexao, adOpenKeyset, adLockOptimistic
        If TBFI.EOF = False Then
            Do While TBFI.EOF = False
                Set TBFornecedor = CreateObject("adodb.recordset")
                TBFornecedor.Open "Select * from Projproduto_clientes where codproduto = " & Codproduto & " and idcliente = " & TBFI!IDCliente, Conexao, adOpenKeyset, adLockOptimistic
                If TBFornecedor.EOF = True Then TBFornecedor.AddNew
                TBFornecedor!Codproduto = Codproduto
                TBFornecedor!IDCliente = TBFI!IDCliente
                TBFornecedor!PConsumo = TBFI!PConsumo
                TBFornecedor!PRevenda = TBFI!PRevenda
                TBFornecedor.Update
                TBFornecedor.Close
                TBFI.MoveNext
            Loop
        End If
    End If
    
    If Chk_compras.Value = 1 Then
        Set TBFI = CreateObject("adodb.recordset")
        TBFI.Open "Select * from Projproduto_fornecedor where codproduto = " & .txtcodproduto, Conexao, adOpenKeyset, adLockOptimistic
        If TBFI.EOF = False Then
            Do While TBFI.EOF = False
                Set TBFornecedor = CreateObject("adodb.recordset")
                TBFornecedor.Open "Select * from Projproduto_fornecedor where codproduto = " & Codproduto & " and idfornecedor = " & TBFI!IDFornecedor, Conexao, adOpenKeyset, adLockOptimistic
                If TBFornecedor.EOF = True Then TBFornecedor.AddNew
                TBFornecedor!Codproduto = Codproduto
                TBFornecedor!IDFornecedor = TBFI!IDFornecedor
                TBFornecedor!PCusto = TBFI!PCusto
                TBFornecedor!Leadtime = TBFI!Leadtime
                TBFornecedor.Update
                TBFornecedor.Close
                TBFI.MoveNext
            Loop
        End If
    End If
        
    Set TBFI = CreateObject("adodb.recordset")
    TBFI.Open "Select * from Projproduto_fabricante where codproduto = " & .txtcodproduto, Conexao, adOpenKeyset, adLockOptimistic
    If TBFI.EOF = False Then
        Do While TBFI.EOF = False
            Set TBFornecedor = CreateObject("adodb.recordset")
            TBFornecedor.Open "Select * from Projproduto_fabricante where codproduto = " & Codproduto & " and idfabricante = " & TBFI!Idfabricante, Conexao, adOpenKeyset, adLockOptimistic
            If TBFornecedor.EOF = True Then TBFornecedor.AddNew
            TBFornecedor!Codproduto = Codproduto
            TBFornecedor!Idfabricante = TBFI!Idfabricante
            TBFornecedor!Part_number = TBFI!Part_number
            TBFornecedor.Update
            TBFornecedor.Close
            TBFI.MoveNext
        Loop
    End If
    TBFI.Close
         
    If txtCodref <> "" Then
        Set TBItem = CreateObject("adodb.recordset")
        TBItem.Open "Select * from item_aplicacoes where n_referencia = '" & txtCodref & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBItem.EOF = False Then
            Set TBProduto = CreateObject("adodb.recordset")
            TBProduto.Open "Select * from projproduto where codproduto = " & TBItem!Codproduto, Conexao, adOpenKeyset, adLockOptimistic
            If TBProduto.EOF = False Then
                If TBProduto!Desenho <> txtCodinterno Then
                    If USMsgBox("Este código de referência está sendo utilizado em outro(s) produto(s), deseja excluir para ser salvo no produto " & txtCodinterno & "?", vbYesNo) = vbYes Then
                        If USMsgBox("Deseja realmente excluir o código de referência " & txtreferencia & " no(s) outro(s) produto(s)?", vbYesNo) = vbYes Then
                            Conexao.Execute "DELETE from item_aplicacoes where  n_referencia = '" & txtCodref & "'"
                        End If
                    End If
                End If
            End If
            TBProduto.Close
        End If
        TBItem.Close
        
        Set TBProduto = CreateObject("adodb.recordset")
        TBProduto.Open "Select * from item_aplicacoes where codproduto = " & Codproduto, Conexao, adOpenKeyset, adLockOptimistic
        If TBProduto.EOF = True Then TBProduto.AddNew
        ProcEnviaDadosRef
        TBProduto.Update
        TBProduto.Close
    Else
        If USMsgBox("Deseja copiar código de referência do produto " & TBAbrir!Desenho & "?", vbYesNo) = vbYes Then
            Permitido = False
            Permitido1 = False
            Set TBCompras = CreateObject("adodb.recordset")
            TBCompras.Open "Select * from item_aplicacoes where codproduto = " & TBAbrir!Codproduto, Conexao, adOpenKeyset, adLockOptimistic
            Do While TBCompras.EOF = False
                Set TBItem = CreateObject("adodb.recordset")
                TBItem.Open "Select * from item_aplicacoes where n_referencia = '" & TBCompras!N_referencia & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBItem.EOF = False Then
                    Set TBProduto = CreateObject("adodb.recordset")
                    TBProduto.Open "Select * from projproduto where codproduto = " & TBAbrir!Codproduto, Conexao, adOpenKeyset, adLockOptimistic
                    If TBProduto.EOF = False Then
                        If TBProduto!Desenho <> txtCodinterno Then
                            If Permitido = False Then
                                If USMsgBox("Este código de referência está sendo utilizado em outro(s) produto(s), deseja excluir para ser salvo no produto " & txtCodinterno & "?", vbYesNo) = vbYes Then
                                    If USMsgBox("Deseja realmente excluir o código de referência no(s) outro(s) produto(s)?", vbYesNo) = vbYes Then
                                        Permitido1 = True
                                    End If
                                End If
                            End If
                            Permitido = True
                        End If
                    End If
                    TBProduto.Close
                End If
                TBItem.Close
                If Permitido1 = True Then Conexao.Execute "DELETE from item_aplicacoes where  n_referencia = '" & TBCompras!N_referencia & "'"
                
                Set TBProduto = CreateObject("adodb.recordset")
                TBProduto.Open "Select * from item_aplicacoes", Conexao, adOpenKeyset, adLockOptimistic
                TBProduto.AddNew
                TBProduto!N_referencia = TBCompras!N_referencia
                TBProduto!Codproduto = Codproduto
                TBProduto!Descricao = TBAbrir!Descricao
                Set TBFI = CreateObject("adodb.recordset")
                TBFI.Open "Select * from item_aplicacoes where codproduto = " & Codproduto, Conexao, adOpenKeyset, adLockOptimistic
                If TBFI.EOF = False Then
                    TBProduto!Rev = TBFI!Rev
                    TBProduto!Aplicacao = TBFI!Aplicacao
                    TBProduto!ID_cliente_forn = TBFI!ID_cliente_forn
                    TBProduto!Tipo = TBFI!Tipo
                    TBProduto!Desenho = TBFI!Desenho
                End If
                TBFI.Close
                TBProduto.Update
                TBProduto.Close
                
                TBCompras.MoveNext
            Loop
            TBCompras.Close
        End If
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviaDadosRef()
On Error GoTo tratar_erro

TBProduto!N_referencia = txtCodref.Text
TBProduto!Codproduto = Codproduto
TBProduto!Descricao = TBAbrir!Descricao
Set TBFI = CreateObject("adodb.recordset")
TBFI.Open "Select * from item_aplicacoes where codproduto = " & Codproduto, Conexao, adOpenKeyset, adLockOptimistic
If TBFI.EOF = False Then
    TBProduto!Rev = TBFI!Rev
    TBProduto!Aplicacao = TBFI!Aplicacao
    TBProduto!ID_cliente_forn = TBFI!ID_cliente_forn
    TBProduto!Tipo = TBFI!Tipo
    TBProduto!Desenho = TBFI!Desenho
End If
TBFI.Close
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_compras_Click()
On Error GoTo tratar_erro

ProcCarregaFamilia

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_PCP_Click()
On Error GoTo tratar_erro

ProcCarregaFamilia

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_qualidade_Click()
On Error GoTo tratar_erro

ProcCarregaFamilia

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_vendas_Click()
On Error GoTo tratar_erro

ProcCarregaFamilia

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyF3: ProcCopiar
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

ProcCarregaToolBar1 Me, 10485, 5, True
If Compras = True Then Caption = "Compras - Produtos e serviços - Copiar"
If Vendas = True Then Caption = "Vendas - Produtos e serviços - Copiar"
 
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

Private Sub Optautomatico_Click()
On Error GoTo tratar_erro

If Optautomatico.Value = True Then
    With txtCodinterno
        .Locked = True
        .TabStop = False
    End With
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optmanual_Click()
On Error GoTo tratar_erro

If Optmanual.Value = True Then
    With txtCodinterno
        .Locked = False
        .TabStop = True
        .Text = ""
        .SetFocus
    End With
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcCopiar
    'Case 3: ProcAjuda
    Case 4: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
