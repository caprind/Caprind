VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Begin VB.Form frmprojproduto_novo 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Engenharia - Produtos e serviços - Novo"
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
   ForeColor       =   &H00000000&
   Icon            =   "frmprojProduto_novo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   9570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      TabIndex        =   34
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
         Caption         =   "Outros"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1545
         TabIndex        =   10
         Top             =   540
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
      Begin VB.OptionButton opt3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Componente"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2820
         TabIndex        =   8
         Top             =   300
         Width           =   1305
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
      Begin VB.OptionButton Opt1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Produto final"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   180
         TabIndex        =   6
         Top             =   300
         Width           =   1305
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
      TabIndex        =   33
      Top             =   990
      Width           =   2235
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
      Height          =   885
      Left            =   55
      TabIndex        =   24
      Top             =   990
      Width           =   1785
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
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Height          =   2085
      Left            =   55
      TabIndex        =   22
      Top             =   1860
      Width           =   9495
      Begin VB.TextBox txtRev 
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
         Left            =   6840
         MaxLength       =   255
         TabIndex        =   17
         ToolTipText     =   "Revisão."
         Top             =   990
         Width           =   795
      End
      Begin VB.TextBox txtDias_antecipacao 
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
         Left            =   8520
         MaxLength       =   255
         TabIndex        =   14
         ToolTipText     =   "Dias para antecipação da produção."
         Top             =   390
         Width           =   795
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
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   12
         ToolTipText     =   "Família."
         Top             =   390
         Width           =   7515
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
         ItemData        =   "frmprojProduto_novo.frx":0442
         Left            =   8490
         List            =   "frmprojProduto_novo.frx":0444
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   19
         ToolTipText     =   "Unidade comercial."
         Top             =   990
         Width           =   825
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
         ItemData        =   "frmprojProduto_novo.frx":0446
         Left            =   7650
         List            =   "frmprojProduto_novo.frx":0448
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   18
         ToolTipText     =   "Unidade de estoque."
         Top             =   990
         Width           =   825
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
         Left            =   7710
         MaxLength       =   255
         TabIndex        =   13
         ToolTipText     =   "Lead time."
         Top             =   390
         Width           =   795
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
         Left            =   3990
         MaxLength       =   50
         TabIndex        =   16
         ToolTipText     =   "Código de referência."
         Top             =   990
         Width           =   2835
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
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Código interno."
         Top             =   990
         Width           =   3795
      End
      Begin VB.TextBox txtdescinicio 
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
         TabIndex        =   20
         ToolTipText     =   "Descrição início."
         Top             =   1620
         Width           =   8135
      End
      Begin MSMask.MaskEdBox Txt_cod_serv 
         Height          =   315
         Left            =   8325
         TabIndex        =   21
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
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Revisão"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   6952
         TabIndex        =   36
         Top             =   780
         Width           =   570
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dias ante."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   8550
         TabIndex        =   35
         Top             =   180
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cód. serv."
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   8438
         TabIndex        =   32
         Top             =   1410
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Un. com."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   8580
         TabIndex        =   31
         Top             =   780
         Width           =   645
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Família"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3697
         TabIndex        =   30
         Top             =   180
         Width           =   480
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Un. est."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   7770
         TabIndex        =   28
         Top             =   780
         Width           =   585
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
         Left            =   7822
         TabIndex        =   27
         Top             =   180
         Width           =   570
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
         Left            =   4537
         TabIndex        =   26
         Top             =   780
         Width           =   1740
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
         Left            =   1462
         TabIndex        =   25
         Top             =   780
         Width           =   1230
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3902
         TabIndex        =   23
         Top             =   1410
         Width           =   690
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   55
      TabIndex        =   29
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
      ButtonLeft3     =   46
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
      ButtonLeft4     =   84
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
      ButtonLeft5     =   112
      ButtonTop5      =   2
      ButtonWidth5    =   24
      ButtonHeight5   =   24
      ButtonUseMaskColor5=   0   'False
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   7800
         Top             =   150
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmprojProduto_novo.frx":044A
         Count           =   1
      End
   End
End
Attribute VB_Name = "frmprojproduto_novo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CodReferencia   As String 'OK

Private Sub ProcNovoProduto()
On Error GoTo tratar_erro

'Verifica se no cadastro da empresa esta marcado para gerar codigo sequencial para produto final
Permitido = False
If Opt1.Value = True Then
    Set TBAliquota = CreateObject("adodb.recordset")
    TBAliquota.Open "Select Codigo_sequencial from Empresa where Codigo_sequencial = 'True'", Conexao, adOpenKeyset, adLockOptimistic
    If TBAliquota.EOF = False Then Permitido = True
    TBAliquota.Close
End If

With frmproj_produto
    Set TBFamilia = CreateObject("adodb.recordset")
    TBFamilia.Open "Select * from projfamilia where familia = '" & cmbfamilia.Text & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBFamilia.EOF = False Then Letra = TBFamilia!Letra
    TBFamilia.Close
    If Optautomatico.Value = True Then
        CompLetra = Len(Letra)
        valor = 6 + CompLetra
        Set TBComponente = CreateObject("adodb.recordset")
        If Permitido = False Then
            TBComponente.Open "Select * from projproduto where classe = '" & cmbfamilia.Text & "' and Right(Desenho, " & CompLetra & ") = '" & Letra & "' and Len(Desenho) = " & valor & " and (subtipoitem = 0 or subtipoitem = 1 or subtipoitem = 4 or subtipoitem = 5) order by codproduto desc", Conexao, adOpenKeyset, adLockOptimistic
        Else
            'O Codigo sequencial não verifica a familia para gerar o codigo interno
            TBComponente.Open "Select * from projproduto where codmanual = 'False' and subtipoitem = 1 order by codproduto desc", Conexao, adOpenKeyset, adLockOptimistic
        End If
        If TBComponente.EOF = False Then
            If Permitido = False Then
                Numero = Left(TBComponente!Desenho, Len(TBComponente!Desenho) - (CompLetra + 1))
            Else
                Numero = Left(TBComponente!Desenho, 5)
            End If
            Numero = Numero + 1
            Select Case Len(Numero)
                Case 5: Desenho = Numero & "-" & Letra
                Case 4: Desenho = "0" & Numero & "-" & Letra
                Case 3: Desenho = "00" & Numero & "-" & Letra
                Case 2: Desenho = "000" & Numero & "-" & Letra
                Case 1: Desenho = "0000" & Numero & "-" & Letra
            End Select
        Else
            Desenho = "00001" & "-" & Letra
        End If
        .txtdesenhoproduto.Text = Desenho
        TBComponente.AddNew
        ProcEnviaDados
        TBComponente.Close
        .txtdesenhoproduto.Locked = True
        .Optmanual = False
    Else
        Set TBComponente = CreateObject("adodb.recordset")
        TBComponente.Open "Select * from projproduto", Conexao, adOpenKeyset, adLockOptimistic
        TBComponente.AddNew
        .txtdesenhoproduto = txtdesenho
        Desenho = txtdesenho
        ProcEnviaDados
        TBComponente.Close
        .txtdesenhoproduto.Locked = False
        .Optmanual.Value = True
    End If
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
    If Optautomatico.Value = True Then
        CompLetra = Len(Letra)
        valor = 6 + CompLetra
        Set TBComponente = CreateObject("adodb.recordset")
        TBComponente.Open "Select * from projproduto where classe = '" & cmbfamilia.Text & "' and Left(Desenho, " & CompLetra & ") = '" & Letra & "' and Len(Desenho) = " & valor & " and (subtipoitem = 2 or subtipoitem = 3) order by codproduto desc", Conexao, adOpenKeyset, adLockOptimistic
        If TBComponente.EOF = False Then
            Numero = Right(TBComponente!Desenho, Len(TBComponente!Desenho) - (CompLetra + 1))
            Numero = Numero + 1
            Select Case Len(Numero)
                Case 5: Desenho = Letra & "-" & Numero
                Case 4: Desenho = Letra & "-0" & Numero
                Case 3: Desenho = Letra & "-00" & Numero
                Case 2: Desenho = Letra & "-000" & Numero
                Case 1: Desenho = Letra & "-0000" & Numero
            End Select
        Else
            Desenho = Letra & "-00001"
        End If
        .txtdesenhoproduto.Text = Desenho
        TBComponente.AddNew
        ProcEnviaDados
        TBComponente.Update
        TBComponente.Close
        .txtdesenhoproduto.Locked = True
        .Optmanual.Value = False
    Else
        Set TBComponente = CreateObject("adodb.recordset")
        TBComponente.Open "Select * from projproduto", Conexao, adOpenKeyset, adLockOptimistic
        TBComponente.AddNew
        Desenho = txtdesenho
        .txtdesenhoproduto.Text = Desenho
        ProcEnviaDados
        TBComponente.Update
        TBComponente.Close
        .txtdesenhoproduto.Locked = False
        .Optmanual.Value = True
    End If
End With

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
    TBComponente!SubTipoItem = 5
Else
    TBComponente!Tipo = "P"
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
With frmproj_produto
    If Chk_PCP.Value = 1 Then
        Opt4.Value = False
        Opt4.Enabled = False
        .Opt4.Enabled = False
    Else
        Opt4.Enabled = True
        .Opt4.Enabled = True
    End If
End With
        
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

Private Sub Opt0_Click()
On Error GoTo tratar_erro

ProcBloqueiaLiberaTipo
ProcBloqueiaLiberaAplicacao

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt1_Click()
On Error GoTo tratar_erro

ProcBloqueiaLiberaTipo
ProcBloqueiaLiberaAplicacao

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcBloqueiaLiberaTipo()
On Error GoTo tratar_erro

With frmproj_produto
    If Optautomatico.Value = True Then
        If Opt0.Value = True Or Opt1.Value = True Or Opt4.Value = True Or Opt5.Value = True Then
            .Opt0.Enabled = True
            .Opt1.Enabled = True
            .Opt2.Enabled = False
            .opt3.Enabled = False
            .Opt4.Enabled = True
            .Opt5.Enabled = True
        Else
            .Opt0.Enabled = False
            .Opt1.Enabled = False
            .Opt2.Enabled = True
            .opt3.Enabled = True
            .Opt4.Enabled = False
            .Opt5.Enabled = False
        End If
    Else
        .Opt0.Enabled = True
        .Opt1.Enabled = True
        .Opt2.Enabled = True
        .opt3.Enabled = True
        .Opt4.Enabled = True
        .Opt5.Enabled = True
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcBloqueiaLiberaAplicacao()
On Error GoTo tratar_erro

With frmproj_produto
    If Opt4.Value = True Then
        Chk_PCP.Value = 0
        Chk_PCP.Enabled = False
        .Chk_PCP.Enabled = False
    Else
        Chk_PCP.Enabled = True
        .Chk_PCP.Enabled = True
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt2_Click()
On Error GoTo tratar_erro

ProcBloqueiaLiberaTipo
ProcBloqueiaLiberaAplicacao

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub opt3_Click()
On Error GoTo tratar_erro

ProcBloqueiaLiberaTipo
ProcBloqueiaLiberaAplicacao

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub OPT4_Click()
On Error GoTo tratar_erro

ProcBloqueiaLiberaTipo
ProcBloqueiaLiberaAplicacao

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt5_Click()
On Error GoTo tratar_erro

ProcBloqueiaLiberaTipo
ProcBloqueiaLiberaAplicacao
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

Private Sub ProcCarregaFamilia()
On Error GoTo tratar_erro

TextoFiltroObrig = ""
If Chk_vendas.Enabled = False Then
    TextoFiltroObrig = "and Vendas = 'True'"
ElseIf Chk_compras.Enabled = False Then
        TextoFiltroObrig = "and Compras = 'True'"
End If

cmbfamilia.Clear
'Vendas + Compras + PCP + Qualidade
If Chk_vendas.Value = 1 And Chk_compras.Value = 1 And Chk_PCP.Value = 1 And Chk_qualidade.Value = 1 Then
    ProcCarregaComboFamilia cmbfamilia, "familia <> 'Null' " & TextoFiltroObrig, False
    'Vendas + Compras + PCP
ElseIf Chk_vendas.Value = 1 And Chk_compras.Value = 1 And Chk_PCP.Value = 1 And Chk_qualidade.Value = 0 Then
        ProcCarregaComboFamilia cmbfamilia, "familia <> 'Null' and (Vendas = 'True' or Compras = 'True' or Fabricacao = 'True') " & TextoFiltroObrig, False
        'Vendas + Compras + Qualidade
    ElseIf Chk_vendas.Value = 1 And Chk_compras.Value = 1 And Chk_PCP.Value = 0 And Chk_qualidade.Value = 1 Then
            ProcCarregaComboFamilia cmbfamilia, "familia <> 'Null' and (Vendas = 'True' or Compras = 'True' or Qualidade = 'True') " & TextoFiltroObrig, False
            'Vendas + Compras
        ElseIf Chk_vendas.Value = 1 And Chk_compras.Value = 1 And Chk_PCP.Value = 0 And Chk_qualidade.Value = 0 Then
                ProcCarregaComboFamilia cmbfamilia, "familia <> 'Null' and (Vendas = 'True' or Compras = 'True') " & TextoFiltroObrig, False
                'Vendas + PCP + Qualidade
            ElseIf Chk_vendas.Value = 1 And Chk_compras.Value = 0 And Chk_PCP.Value = 1 And Chk_qualidade.Value = 1 Then
                    ProcCarregaComboFamilia cmbfamilia, "familia <> 'Null' and (Vendas = 'True' or Fabricacao = 'True' or Qualidade = 'True') " & TextoFiltroObrig, False
                    'Vendas + Qualidade
                ElseIf Chk_vendas.Value = 1 And Chk_compras.Value = 0 And Chk_PCP.Value = 0 And Chk_qualidade.Value = 1 Then
                        ProcCarregaComboFamilia cmbfamilia, "familia <> 'Null' and (Vendas = 'True' or Qualidade = 'True') " & TextoFiltroObrig, False
                        'Vendas + PCP
                    ElseIf Chk_vendas.Value = 1 And Chk_compras.Value = 0 And Chk_PCP.Value = 1 And Chk_qualidade.Value = 0 Then
                            ProcCarregaComboFamilia cmbfamilia, "familia <> 'Null' and (Vendas = 'True' or Fabricacao = 'True') " & TextoFiltroObrig, False
                            'Compras + PCP + Qualidade
                        ElseIf Chk_vendas.Value = 0 And Chk_compras.Value = 1 And Chk_PCP.Value = 1 And Chk_qualidade.Value = 1 Then
                                ProcCarregaComboFamilia cmbfamilia, "familia <> 'Null' and (Compras = 'True' or Fabricacao = 'True' or Qualidade = 'True') " & TextoFiltroObrig, False
                                'Compras + Qualidade
                            ElseIf Chk_vendas.Value = 0 And Chk_compras.Value = 1 And Chk_PCP.Value = 0 And Chk_qualidade.Value = 1 Then
                                    ProcCarregaComboFamilia cmbfamilia, "familia <> 'Null' and (Compras = 'True' or Qualidade = 'True') " & TextoFiltroObrig, False
                                    'Compras + PCP
                                ElseIf Chk_vendas.Value = 0 And Chk_compras.Value = 1 And Chk_PCP.Value = 1 And Chk_qualidade.Value = 0 Then
                                        ProcCarregaComboFamilia cmbfamilia, "familia <> 'Null' and (Compras = 'True' or Fabricacao = 'True') " & TextoFiltroObrig, False
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
ProcAtualizaUNCom

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbFamilia_Click()
On Error GoTo tratar_erro

txtdescinicio = FunBuscaDescPadraoFamilia(cmbfamilia, txtdesenho, txtdescinicio)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvar()
On Error GoTo tratar_erro

Acao = "criar um novo registro"
'If Chk_vendas.Value = 0 And Chk_compras.Value = 0 And Chk_PCP.Value = 0 And Chk_qualidade.Value = 0 Then
'    NomeCampo = "a aplicação"
'    ProcVerificaAcao
'    Exit Sub
'End If
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
        USMsgBox ("Não é permitido cadastrar um novo registro, pois este código interno já está cadastrado."), vbExclamation, "CAPRIND v5.0"
        txtdesenho.SetFocus
        TBProduto.Close
        Exit Sub
    End If
    TBProduto.Close
End If
If cmbun.Text = "" Then
    NomeCampo = "a unidade de estoque"
    ProcVerificaAcao
    cmbun.SetFocus
    Exit Sub
End If
If Cmb_un_com.Text = "" Then
    NomeCampo = "a unidade comercial"
    ProcVerificaAcao
    Cmb_un_com.SetFocus
    Exit Sub
End If
valor = IIf(txtleadtime = "", 0, txtleadtime)
If txtleadtime = "" Or valor < 0 Then
    NomeCampo = "o lead time"
    ProcVerificaAcao
    txtleadtime.SetFocus
    Exit Sub
End If
If txtdescinicio.Text = "" Then
    NomeCampo = "a descrição"
    ProcVerificaAcao
    txtdescinicio.SetFocus
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
TBProduto.Open "Select * from projproduto where descricao = '" & txtdescinicio & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    If USMsgBox("Já existe um registro com esta descrição, deseja prosseguir assim mesmo?", vbYesNo, "CAPRIND v5.0") = vbNo Then
        txtdescinicio.SetFocus
        TBProduto.Close
        Exit Sub
    End If
End If
TBProduto.Close

With frmproj_produto
    .Optautomatico = Optautomatico
    .Optmanual = Optmanual
    .cmbfamilia = cmbfamilia.Text
    .cmbfamilia.Locked = True
    .cmbfamilia.TabStop = False
    .Chk_vendas = Chk_vendas
    .Chk_compras = Chk_compras
    .Chk_PCP = Chk_PCP
    .Chk_qualidade = Chk_qualidade
    .Opt5 = Opt5
    .Opt4 = Opt4
    .opt3 = opt3
    .Opt2 = Opt2
    .Opt1 = Opt1
    .Opt0 = Opt0
    .txtDescricaoProduto.Text = txtdescinicio
    .txtespecificacao.Text = txtdescinicio
    .cmbun = cmbun
    .Cmb_un_com = Cmb_un_com
    .txtleadtime = txtleadtime
    .txtDias_antecipacao = txtDias_antecipacao
    .txtrevdesproduto = IIf(txtRev = "", 0, txtRev)
    .txtreferencia = txtreferencia
    .Txt_cod_serv = Txt_cod_serv
    .txtreferencia.Enabled = True
    .Lista.ListItems.Clear
    .Novo_Produto = True
    If Opt2.Value = True Or opt3.Value = True Then ProcNovoItem Else ProcNovoProduto
    .Frame2.Enabled = True
    .Frame12.Enabled = False
    .ProcGravar
    
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select * From projproduto where Desenho = '" & .txtdesenhoproduto & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
        .ProcCarregaDados
    End If
End With
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbun_Click()
On Error GoTo tratar_erro
  
ProcAtualizaUNCom

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

ProcCarregaToolBar1 Me, 9495, 5, True
If Compras_Produtos = True Then
    Caption = "Compras - Produtos e serviços - Novo"
    With Chk_compras
        .Enabled = False
        .Value = 1
    End With
End If
If Vendas_Produtos = True Then
    Caption = "Vendas - Produtos e serviços - Novo"
    With Chk_vendas
        .Enabled = False
        .Value = 1
    End With
End If
ProcCarregaComboUnidade cmbun, False
ProcCarregaComboUnidade Cmb_un_com, False

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
    With txtdesenho
        .Locked = True
        .TabStop = False
        .Text = ""
    End With
    ProcBloqueiaLiberaTipo
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
    End With
    ProcBloqueiaLiberaTipo
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviaDadosRef()
On Error GoTo tratar_erro

TBItem!N_referencia = CodReferencia
TBItem!Codproduto = Codproduto
TBItem!Descricao = txtdescinicio

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcEnviaDados()
On Error GoTo tratar_erro

With frmproj_produto
    TBComponente!Data = Date
    TBComponente!Responsavel = pubUsuario
    If Optmanual.Value = True Then TBComponente!CodManual = True Else TBComponente!CodManual = False
    TBComponente!Desenho = .txtdesenhoproduto
    TBComponente!RevDesenho = 0
    TBComponente!Unidade = cmbun
    TBComponente!Unidade_com = Cmb_un_com
    TBComponente!Leadtime = txtleadtime
    TBComponente!Dias_antecipacao = IIf(txtDias_antecipacao = "", Null, txtDias_antecipacao)
    TBComponente!RevDesenho = IIf(txtRev = "", 0, txtRev)
    TBComponente!Descricao = txtdescinicio
    TBComponente!descricaotecnica = txtdescinicio
    TBComponente!Cod_servico = Txt_cod_serv
    TBComponente!Classe = cmbfamilia.Text
    TBComponente!peso_metro = 0
    TBComponente!Un_Kg = "N/a"
    
    'Conta contábil, CC e NCM
    Set TBFamilia = CreateObject("adodb.recordset")
    TBFamilia.Open "select ID_PC, ID_PC1, ID_CC, ID_CF from projfamilia where Familia = '" & cmbfamilia & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBFamilia.EOF = False Then
        If IsNull(TBFamilia!ID_PC) = False And TBFamilia!ID_PC <> "" Then TBComponente!ID_PC = TBFamilia!ID_PC
        If IsNull(TBFamilia!ID_PC1) = False And TBFamilia!ID_PC1 <> "" Then TBComponente!ID_PC1 = TBFamilia!ID_PC1
        If IsNull(TBFamilia!ID_CC) = False And TBFamilia!ID_CC <> "" Then TBComponente!ID_CC = TBFamilia!ID_CC
        If IsNull(TBFamilia!ID_CF) = False And TBFamilia!ID_CF <> "" Then TBComponente!ID_CF = TBFamilia!ID_CF
    End If
    TBFamilia.Close
    
    ProcEnviaDadosAplicacaoFamilia
    TBComponente.Update
    .txtcodproduto = TBComponente!Codproduto
    Codproduto = TBComponente!Codproduto
    CodReferencia = txtreferencia

    If CodReferencia <> "" Then
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from item_aplicacoes where n_referencia = '" & CodReferencia & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Set TBProduto = CreateObject("adodb.recordset")
            TBProduto.Open "Select * from projproduto where codproduto = " & TBAbrir!Codproduto, Conexao, adOpenKeyset, adLockOptimistic
            If TBProduto.EOF = False Then
                If TBProduto!Desenho <> Desenho Then
                    If USMsgBox("Este código de referência está sendo utilizado em outro(s) produto(s), deseja excluir para ser salvo no produto?", vbYesNo) = vbYes Then
                        If USMsgBox("Deseja realmente excluir o código de referência " & txtreferencia & " no(s) outro(s) produto(s)?", vbYesNo) = vbYes Then
                            Conexao.Execute "DELETE from item_aplicacoes where n_referencia = '" & txtreferencia & "'"
                        End If
                    End If
                End If
            Else
                Conexao.Execute "DELETE from item_aplicacoes where codproduto = " & TBAbrir!Codproduto
            End If
            TBProduto.Close
        End If
        TBAbrir.Close
        
        Set TBItem = CreateObject("adodb.recordset")
        TBItem.Open "Select * from item_aplicacoes where codproduto = " & Codproduto, Conexao, adOpenKeyset, adLockOptimistic
        If TBItem.EOF = True Then TBItem.AddNew
        ProcEnviaDadosRef
        TBItem.Update
        TBItem.Close
    End If
continuar:
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtDias_antecipacao_Change()
On Error GoTo tratar_erro

If txtDias_antecipacao.Text <> "" Then
    VerifNumero = txtDias_antecipacao.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtDias_antecipacao.Text = ""
        txtDias_antecipacao.SetFocus
        Exit Sub
    End If
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
    Case 1: ProcSalvar
    'Case 3: ProcAjuda
    Case 4: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAtualizaUNCom()
On Error GoTo tratar_erro

'With Cmb_un_com
'    If cmbun <> "" Then .Text = cmbun
'    If Chk_vendas.Value = 1 Or Chk_compras.Value = 1 Then
'        .Locked = False
'        .TabStop = True
'    Else
'        .Locked = True
'        .TabStop = False
'    End If
'End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
