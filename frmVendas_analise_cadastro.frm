VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Begin VB.Form frmVendas_analise_cadastro 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Outros - Análise crítica - Cadastrar produto/processo"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9570
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
   ForeColor       =   &H8000000D&
   Icon            =   "frmVendas_analise_cadastro.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   9570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1050
         TabIndex        =   3
         Top             =   300
         Width           =   975
      End
      Begin VB.CheckBox Chk_qualidade 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Qualidade"
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   150
         TabIndex        =   4
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
      Width           =   5445
      Begin VB.OptionButton Opt5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Serviço"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2820
         TabIndex        =   11
         Top             =   540
         Width           =   825
      End
      Begin VB.OptionButton Opt4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Insumo"
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   180
         TabIndex        =   9
         Top             =   540
         Width           =   1305
      End
   End
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   7110
      Top             =   150
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frmVendas_analise_cadastro.frx":0442
      Count           =   1
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   55
      TabIndex        =   20
      Top             =   0
      Width           =   9495
      _ExtentX        =   16748
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
         Name            =   "Tahoma"
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
      ButtonUseMaskColor2=   0   'False
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
      ButtonEnabled5  =   0   'False
      ButtonIconSize5 =   32
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
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Height          =   2085
      Left            =   60
      TabIndex        =   24
      Top             =   1860
      Width           =   9495
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
         Left            =   180
         MaxLength       =   255
         TabIndex        =   18
         ToolTipText     =   "Descrição início."
         Top             =   1620
         Width           =   8135
      End
      Begin VB.TextBox txtdesenho 
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
         Left            =   180
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Código interno."
         Top             =   990
         Width           =   3225
      End
      Begin VB.TextBox txtreferencia 
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
         Left            =   3420
         MaxLength       =   50
         TabIndex        =   14
         ToolTipText     =   "Código de referência."
         Top             =   990
         Width           =   3255
      End
      Begin VB.TextBox txtleadtime 
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
         Left            =   8370
         MaxLength       =   255
         TabIndex        =   17
         ToolTipText     =   "Lead time."
         Top             =   990
         Width           =   945
      End
      Begin VB.ComboBox cmbun 
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
         ItemData        =   "frmVendas_analise_cadastro.frx":2243
         Left            =   6690
         List            =   "frmVendas_analise_cadastro.frx":2245
         Locked          =   -1  'True
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Unidade de estoque."
         Top             =   990
         Width           =   825
      End
      Begin VB.ComboBox Cmb_un_com 
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
         ItemData        =   "frmVendas_analise_cadastro.frx":2247
         Left            =   7530
         List            =   "frmVendas_analise_cadastro.frx":2249
         Locked          =   -1  'True
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Unidade comercial."
         Top             =   990
         Width           =   825
      End
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
         Locked          =   -1  'True
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Família."
         Top             =   390
         Width           =   9135
      End
      Begin MSMask.MaskEdBox Txt_cod_serv 
         Height          =   315
         Left            =   8325
         TabIndex        =   19
         ToolTipText     =   "Código do serviço conforme Tabela de Serviços da Lei Complementar 116/2003."
         Top             =   1620
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Enabled         =   0   'False
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##.##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3902
         TabIndex        =   32
         Top             =   1410
         Width           =   690
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código interno"
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
         Left            =   1177
         TabIndex        =   31
         Top             =   780
         Width           =   1230
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código de referência"
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
         Left            =   4177
         TabIndex        =   30
         Top             =   780
         Width           =   1740
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "L. time"
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
         Left            =   8520
         TabIndex        =   29
         Top             =   780
         Width           =   570
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Un. est."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   6810
         TabIndex        =   28
         Top             =   780
         Width           =   585
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Família"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5032
         TabIndex        =   27
         Top             =   180
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Un. com."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   7620
         TabIndex        =   26
         Top             =   780
         Width           =   645
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cód. serv."
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   8438
         TabIndex        =   25
         Top             =   1410
         Width           =   765
      End
   End
End
Attribute VB_Name = "frmVendas_analise_cadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CodReferencia   As String 'OK

Private Sub ProcGravar()
On Error GoTo tratar_erro

Acao = "salvar"
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
If cmbfamilia.Text = "" Then
    NomeCampo = "a família"
    ProcVerificaAcao
    cmbfamilia.SetFocus
    Exit Sub
End If
If Optmanual.Value = True Then
    If txtdesenho.Text = "" Then
        NomeCampo = "o código interno"
        ProcVerificaAcao
        txtdesenho.SetFocus
        Exit Sub
    End If
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select * from projproduto where desenho = '" & txtdesenho.Text & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
        USMsgBox ("Não é permitido cadastrar um novo registro, pois esse código interno já está cadastrado."), vbExclamation, "CAPRIND v5.0"
        txtdesenho.SetFocus
        TBProduto.Close
        Exit Sub
    End If
    TBProduto.Close
End If
If txtleadtime.Text = "" Then
    NomeCampo = "o lead time"
    ProcVerificaAcao
    txtleadtime.SetFocus
    Exit Sub
End If
If Txt_descricao = "" Then
    NomeCampo = "a descrição"
    ProcVerificaAcao
    Txt_descricao.SetFocus
    Exit Sub
End If
If Opt5.Value = True And Txt_cod_serv <> "__.__" Then
    Txt_cod_serv.PromptInclude = False
    If Len(Txt_cod_serv.Text) < 4 Then
        USMsgBox ("Verifique se faltam dados no campo código do serviço á serem preenchidos."), vbExclamation, "CAPRIND v5.0"
        Txt_cod_serv.SetFocus
        Exit Sub
    End If
    Txt_cod_serv.PromptInclude = True
End If
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select Codproduto from projproduto where descricao = '" & Txt_descricao & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    If USMsgBox("Já existe um registro com esta descrição, deseja prosseguir assim mesmo?", vbYesNo, "CAPRIND v5.0") = vbNo Then
        Txt_descricao.SetFocus
        TBProduto.Close
        Exit Sub
    End If
End If
TBProduto.Close
ProcNovoProduto
Permitido1 = True
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovoProduto()
On Error GoTo tratar_erro

With frmVendas_analise
    Set TBFamilia = CreateObject("adodb.recordset")
    TBFamilia.Open "Select * from projfamilia where familia = '" & cmbfamilia & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBFamilia.EOF = False Then Letra = TBFamilia!Letra
    TBFamilia.Close
    
    CompLetra = Len(Letra)
    If Optautomatico.Value = True Then
        If Permitido = True Then
            Set TBComponente = CreateObject("adodb.recordset")
            TBComponente.Open "Select * from projproduto where classe = '" & cmbfamilia & "' and codmanual = 'False' and Right(Desenho, " & CompLetra & ") = '" & Letra & "' and (subtipoitem = 0 or subtipoitem = 1 or subtipoitem = 4 or subtipoitem = 5) order by codproduto", Conexao, adOpenKeyset, adLockOptimistic
            If TBComponente.EOF = False Then
                TBComponente.MoveLast
                CompLetra = Len(Letra) + 1
                Numero = Left(TBComponente!Desenho, Len(TBComponente!Desenho) - CompLetra)
                Numero = Numero + 1
                If Len(Numero) = 5 Then Desenho = Numero & "-" & Letra
                If Len(Numero) = 4 Then Desenho = "0" & Numero & "-" & Letra
                If Len(Numero) = 3 Then Desenho = "00" & Numero & "-" & Letra
                If Len(Numero) = 2 Then Desenho = "000" & Numero & "-" & Letra
                If Len(Numero) = 1 Then Desenho = "0000" & Numero & "-" & Letra
            Else
                Desenho = "00001" & "-" & Letra
            End If
        Else
            Set TBComponente = CreateObject("adodb.recordset")
            TBComponente.Open "Select * from projproduto where classe = '" & cmbfamilia.Text & "' and codmanual = 'False' and Left(Desenho, " & CompLetra & ") = '" & Letra & "' and (subtipoitem = 2 or subtipoitem = 3) order by codproduto", Conexao, adOpenKeyset, adLockOptimistic
            If TBComponente.EOF = False Then
                TBComponente.MoveLast
                CompLetra = Len(Letra) + 1
                Numero = Right(TBComponente!Desenho, Len(TBComponente!Desenho) - CompLetra)
                Numero = Numero + 1
                If Len(Numero) = 5 Then Desenho = Letra & "-" & Numero
                If Len(Numero) = 4 Then Desenho = Letra & "-0" & Numero
                If Len(Numero) = 3 Then Desenho = Letra & "-00" & Numero
                If Len(Numero) = 2 Then Desenho = Letra & "-000" & Numero
                If Len(Numero) = 1 Then Desenho = Letra & "-0000" & Numero
            Else
                Desenho = Letra & "-00001"
            End If
        End If
        TBComponente.AddNew
        ProcEnviaDados
        TBComponente.Close
    Else
        Set TBComponente = CreateObject("adodb.recordset")
        TBComponente.Open "Select * from projproduto", Conexao, adOpenKeyset, adLockOptimistic
        TBComponente.AddNew
        Desenho = txtdesenho
        TBComponente!CodManual = True
        ProcEnviaDados
        TBComponente.Close
    End If
End With

If Permitido = True Then Conexao.Execute "Update Vendas_analise Set idproduto = " & Codproduto & ", codinterno = '" & Desenho & "', RevDesenho = 0, N_Referencia = '" & txtreferencia & "', Descricao = '" & Txt_descricao & "' where ID = " & frmVendas_analise.txtId

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcEnviaDados()
On Error GoTo tratar_erro

TBComponente!Data = Date
TBComponente!Responsavel = pubUsuario
TBComponente!Desenho = Desenho
TBComponente!RevDesenho = 0
TBComponente!Unidade = cmbun
TBComponente!Unidade_com = Cmb_un_com
TBComponente!Leadtime = txtleadtime
TBComponente!Descricao = Txt_descricao
TBComponente!descricaotecnica = Txt_descricao
TBComponente!Cod_servico = Txt_cod_serv
TBComponente!Classe = cmbfamilia

'Conta contábil e CC
Set TBFamilia = CreateObject("adodb.recordset")
TBFamilia.Open "Select * from Projfamilia where Familia = '" & cmbfamilia & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBFamilia.EOF = False Then
    TBComponente!ID_PC = IIf(IsNull(TBFamilia!ID_PC), 0, TBFamilia!ID_PC)
End If
TBFamilia.Close

ProcEnviaDadosAplicacaoFamilia
TBComponente.Update
Codproduto = TBComponente!Codproduto
CodReferencia = txtreferencia

If CodReferencia <> "" Then
    Set TBRecebidos = CreateObject("adodb.recordset")
    TBRecebidos.Open "Select * from item_aplicacoes where n_referencia = '" & CodReferencia & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBRecebidos.EOF = False Then
        Set TBProduto = CreateObject("adodb.recordset")
        TBProduto.Open "Select * from projproduto where codproduto = " & TBRecebidos!Codproduto, Conexao, adOpenKeyset, adLockOptimistic
        If TBProduto.EOF = False Then
            If TBProduto!Desenho <> Desenho Then
                If USMsgBox("Este código de referência está sendo utilizado no produto " & TBProduto!Desenho & " deseja alterar para o produto " & Desenho & "?", vbYesNo) = vbYes Then
                    Conexao.Execute "DELETE from item_aplicacoes where  n_referencia = '" & txtreferencia & "'"
                    TBProduto.Close
                    GoTo Referencia
                Else
                    txtreferencia = ""
                End If
            Else
                TBProduto.Close
                GoTo Referencia
            End If
        Else
            Conexao.Execute "DELETE from item_aplicacoes where codproduto = " & TBRecebidos!Codproduto
            GoTo continuar
        End If
        TBProduto.Close
    Else
Referencia:
        Set TBItem = CreateObject("adodb.recordset")
        TBItem.Open "Select * from item_aplicacoes where codproduto = " & Codproduto, Conexao, adOpenKeyset, adLockOptimistic
        If TBItem.EOF = False Then
            If txtreferencia.Text <> "" Then
                ProcEnviaDadosRef
                TBItem.Update
            End If
        Else
            If txtreferencia.Text <> "" Then
                TBItem.AddNew
                ProcEnviaDadosRef
                TBItem.Update
            End If
        End If
        TBItem.Close
    End If
    TBRecebidos.Close
End If
continuar:

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviaDadosRef()
On Error GoTo tratar_erro

TBItem!N_referencia = CodReferencia
TBItem!Codproduto = Codproduto
TBItem!Descricao = Txt_descricao

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviaDadosAplicacaoFamilia()
On Error GoTo tratar_erro

If Chk_vendas.Value = 1 Then TBComponente!Vendas = True Else TBComponente!Vendas = False
If Chk_compras.Value = 1 Then TBComponente!Compras = True Else TBComponente!Compras = False
If Chk_PCP.Value = 1 Then TBComponente!Producao = True Else TBComponente!Producao = False
If Chk_qualidade.Value = 1 Then TBComponente!Qualidade = True Else TBComponente!Qualidade = False
If Opt5.Value = True Then
    TBComponente!Tipo = "S"
Else
    TBComponente!Tipo = "P"
    TBComponente!SubTipoItem = 5
End If
If Opt4.Value = True Then TBComponente!SubTipoItem = 4
If opt3.Value = True Then TBComponente!SubTipoItem = 3
If Opt2.Value = True Then TBComponente!SubTipoItem = 2
If Opt1.Value = True Then TBComponente!SubTipoItem = 1
If Opt0.Value = True Then TBComponente!SubTipoItem = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt5_Click()
On Error GoTo tratar_erro

With Txt_cod_serv
    If Opt5.Value = True Then
        .Enabled = True
    Else
        .Enabled = False
        .Text = "__.__"
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro
  
Select Case KeyCode
    Case vbKeyF3: ProcGravar
    Case vbKeyEscape: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 9495, 5, True

With frmVendas_analise
    Caption = "Outros - Análise crítica - Cadastrar produto/processo (Cód. interno : " & TBClientes!Codinterno & " . Descrição : " & TBClientes!Descricao & ")"
    ProcCarregaComboFamilia cmbfamilia, "familia <> 'Null'", False
    ProcCarregaComboUnidade cmbun, False
    ProcCarregaComboUnidade Cmb_un_com, False
    
    cmbfamilia = TBClientes!Familia
    txtrevdesproduto = 0
    cmbun = TBClientes!Un
    Cmb_un_com = TBClientes!Unidade_com
    Txt_descricao = TBClientes!Descricao
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

Permitido1 = False
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optautomatico_Click()
On Error GoTo tratar_erro

If Optautomatico.Value = True Then
    With txtdesenho
        .Locked = True
        .TabStop = False
        .Text = ""
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
    With txtdesenho
        .Locked = False
        .TabStop = True
        .Text = TBClientes!Codinterno
    End With
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtleadtime_Change()
On Error GoTo tratar_erro

If txtleadtime.Text <> "" Then
    VerifNumero = txtleadtime.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtleadtime.Text = ""
        txtleadtime.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcGravar
    'Case 3: ProcAjuda
    Case 4: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
