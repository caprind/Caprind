VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Begin VB.Form frmVendas_cliente_localizar 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Administrativo | Vendas - Clientes - Localizar"
   ClientHeight    =   3030
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   9480
   ClipControls    =   0   'False
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   9480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   18
      Top             =   2625
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   714
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   465
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   820
      DibPicture      =   "frmVendas_cliente_localizar.frx":0000
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
      Icon            =   "frmVendas_cliente_localizar.frx":3650
   End
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
      Height          =   1575
      Left            =   1470
      TabIndex        =   8
      Top             =   690
      Width           =   7785
      Begin DrawSuite2022.USButton btnFiltrar 
         Height          =   1185
         Left            =   6030
         TabIndex        =   20
         Top             =   270
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   2090
         DibPicture      =   "frmVendas_cliente_localizar.frx":396A
         Caption         =   "Filtrar"
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
         PicAlign        =   8
         PicSize         =   3
         PicSizeH        =   32
         PicSizeW        =   32
         ShowFocusRect   =   0   'False
         Theme           =   4
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   3000
         TabIndex        =   12
         Top             =   210
         Width           =   2955
         Begin VB.OptionButton optIgual 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Igual"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2130
            TabIndex        =   16
            Top             =   180
            Width           =   705
         End
         Begin VB.OptionButton Optmeio 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Meio"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   900
            TabIndex        =   15
            Top             =   180
            Width           =   675
         End
         Begin VB.OptionButton Optinicio 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Início"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   180
            TabIndex        =   14
            Top             =   180
            Value           =   -1  'True
            Width           =   1275
         End
         Begin VB.OptionButton Optfim 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Fim"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1560
            TabIndex        =   13
            Top             =   180
            Width           =   585
         End
      End
      Begin VB.TextBox txtTexto 
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
         TabIndex        =   3
         ToolTipText     =   "Texto para pesquisa."
         Top             =   1110
         Width           =   5775
      End
      Begin VB.ComboBox cmbfiltrarpor 
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
         Height          =   330
         ItemData        =   "frmVendas_cliente_localizar.frx":6FBA
         Left            =   210
         List            =   "frmVendas_cliente_localizar.frx":6FD9
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "Opções para filtro."
         Top             =   390
         Width           =   2715
      End
      Begin MSMask.MaskEdBox txtcnpj 
         Height          =   315
         Left            =   180
         TabIndex        =   6
         ToolTipText     =   "Número do CNPJ."
         Top             =   1110
         Visible         =   0   'False
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         AllowPrompt     =   -1  'True
         AutoTab         =   -1  'True
         MaxLength       =   18
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##.###.###/####-##"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox cmbstatus 
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
         ItemData        =   "frmVendas_cliente_localizar.frx":7041
         Left            =   180
         List            =   "frmVendas_cliente_localizar.frx":704E
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         ToolTipText     =   "Status."
         Top             =   1110
         Visible         =   0   'False
         Width           =   5775
      End
      Begin MSMask.MaskEdBox txtCpf 
         Height          =   315
         Left            =   180
         TabIndex        =   7
         ToolTipText     =   "Número do CPF."
         Top             =   1110
         Visible         =   0   'False
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         AllowPrompt     =   -1  'True
         AutoTab         =   -1  'True
         MaxLength       =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "###.###.###-##"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox cmbFamilia 
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
         Left            =   180
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         ToolTipText     =   "Texto para pesquisa."
         Top             =   1110
         Visible         =   0   'False
         Width           =   5775
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Texto para pesquisa"
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
         TabIndex        =   11
         Top             =   900
         Width           =   1470
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filtrar por"
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
         Left            =   1140
         TabIndex        =   10
         Top             =   180
         Width           =   705
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
      Height          =   1575
      Left            =   150
      TabIndex        =   9
      Top             =   690
      Width           =   1305
      Begin VB.OptionButton optTodos 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Todos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   270
         TabIndex        =   21
         Top             =   660
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.CheckBox Chk_prospecto 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Prospecto"
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
         TabIndex        =   19
         Top             =   330
         Width           =   1035
      End
      Begin VB.OptionButton optFisica 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Física"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   270
         TabIndex        =   1
         Top             =   1140
         Width           =   855
      End
      Begin VB.OptionButton optJuridica 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Jurídica"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   270
         TabIndex        =   0
         Top             =   900
         Width           =   975
      End
   End
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   7110
      Top             =   150
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frmVendas_cliente_localizar.frx":7070
      Count           =   1
   End
End
Attribute VB_Name = "frmVendas_cliente_localizar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnFiltrar_Click()
On Error GoTo tratar_erro

ProcFiltrar

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbFamilia_Click()
On Error GoTo tratar_erro

If cmbfamilia.Text <> "" Then
    txtTexto.Text = ""
    cmbStatus.ListIndex = -1
    txtcnpj.Text = "__.___.___/____-__"
    txtCpf.Text = "___.___.___-__"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbfiltrarpor_Click()
On Error GoTo tratar_erro

If cmbfiltrarpor = "Razão social" Or cmbfiltrarpor = "Nome fantasia" Or cmbfiltrarpor = "Cidade" Or cmbfiltrarpor = "Código do cliente" Then
    txtTexto.Visible = True
    cmbfamilia.Visible = False
    cmbStatus.Visible = False
    txtcnpj.Visible = False
    txtCpf.Visible = False
End If
If cmbfiltrarpor = "Família" Or cmbfiltrarpor = "Grupo" Or cmbfiltrarpor = "Vendedor" Then
    txtTexto.Visible = False
    cmbfamilia.Visible = True
    cmbStatus.Visible = False
    txtcnpj.Visible = False
    txtCpf.Visible = False
    cmbfamilia.Clear
    If cmbfiltrarpor = "Família" Then
        ProcCarregaComboFamilia cmbfamilia, "familia <> 'Null' and vendas = 'True'", True
    ElseIf cmbfiltrarpor = "Vendedor" Then
        Set TBFamilia = CreateObject("adodb.recordset")
        TBFamilia.Open "Select * from Vendas_Vendedores order by Vendedor", Conexao, adOpenKeyset, adLockOptimistic
        If TBFamilia.EOF = False Then
            Do While TBFamilia.EOF = False
                cmbfamilia.AddItem TBFamilia!vendedor
                TBFamilia.MoveNext
            Loop
        End If
        TBFamilia.Close
    Else
        Set TBFamilia = CreateObject("adodb.recordset")
        TBFamilia.Open "Select * from Clientes_grupos where Texto <> 'Null' order by Texto", Conexao, adOpenKeyset, adLockOptimistic
        If TBFamilia.EOF = False Then
            Do While TBFamilia.EOF = False
                cmbfamilia.AddItem TBFamilia!Texto
                TBFamilia.MoveNext
            Loop
        End If
        TBFamilia.Close
    End If
End If
If cmbfiltrarpor = "Status" Then
    txtTexto.Visible = False
    cmbfamilia.Visible = False
    cmbStatus.Visible = True
    txtcnpj.Visible = False
    txtCpf.Visible = False
End If
If cmbfiltrarpor = "CNPJ/CPF" And optJuridica.Value = True Then
    txtTexto.Visible = False
    cmbfamilia.Visible = False
    cmbStatus.Visible = False
    txtcnpj.Visible = True
    txtCpf.Visible = False
End If
If cmbfiltrarpor = "CNPJ/CPF" And optFisica.Value = True Then
    txtTexto.Visible = False
    cmbfamilia.Visible = False
    cmbStatus.Visible = False
    txtcnpj.Visible = False
    txtCpf.Visible = True
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbstatus_Click()
On Error GoTo tratar_erro

If cmbStatus.Text <> "" Then
    txtTexto.Text = ""
    cmbfamilia.ListIndex = -1
    txtcnpj.Text = "__.___.___/____-__"
    txtCpf.Text = "___.___.___-__"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

With frmVendas_cliente

    If optFisica.Value = True Then
'        TipoPessoa = "(C.tipo = 'FP' or C.tipo = 'FR') AND "
'        TipoPessoaRel = "{clientes.tipo} = 'FP' or {clientes.tipo} = 'FR'"
        CpfCnpj = "C.cpf_cnpj = '" & txtCpf.Text & "'"
        CPFCNPJRel = "{clientes.cpf_cnpj} = '" & txtCpf.Text & "'"
    End If
    
   If optJuridica.Value = True Then
'        TipoPessoa = "(C.tipo = 'JP' or C.tipo = 'JR') "
'        TipoPessoaRel = "({clientes.tipo} = 'JP' or {clientes.tipo} = 'JR') "
        CpfCnpj = "C.cpf_cnpj = '" & txtcnpj.Text & "'"
        CPFCNPJRel = "{clientes.cpf_cnpj} = '" & txtcnpj.Text & "'"
    End If
    
   If optTodos.Value = True Then
'        TipoPessoa = "(C.tipo = 'JP'or C.tipo = 'JR')"
'        TipoPessoaRel = "({clientes.tipo} = 'JP' or {clientes.tipo} = 'JR')"
        CpfCnpj = "C.cpf_cnpj <> '""'"
        CPFCNPJRel = "{clientes.cpf_cnpj} <> '""'"
    End If
    
    
    If Chk_prospecto.Value = 1 Then
        Prospecto = "C.Prospecto = 'True'"
        ProspectoRel = "{clientes.Prospecto} = True"
    Else
        Prospecto = "C.Prospecto = 'False'"
        ProspectoRel = "{clientes.Prospecto} = False"
    End If
    
    CamposFiltro = "C.IDCliente, C.Data, C.Responsavel, C.NomeRazao, C.DtValidacao, C.ID"
    INNERJOINTEXTO = "Select " & CamposFiltro & " from (Clientes C LEFT JOIN compras_fornecedores_familia CFF ON C.IDCliente = CFF.IDCliente) LEFT JOIN Clientes_grupos CG ON CG.ID = C.IDGrupo "
    TextoFiltroPadrao = Prospecto & " group by " & CamposFiltro & " order by C.nomerazao"
    TextoFiltroPadraoRel = ProspectoRel & IIf(cmbfiltrarpor <> "Status", " and {Clientes.status} = 'Liberado'", "")
    
    If txtTexto <> "" Or cmbfamilia <> "" Or cmbStatus <> "" Or txtcnpj <> "__.___.___/____-__" Or txtCpf <> "___.___.___-__" Then
        If cmbfiltrarpor = "Status" Then
            .StrSql_Cliente = INNERJOINTEXTO & " where C.status = '" & cmbStatus.Text & "' and " & TextoFiltroPadrao
            .FormulaRel_Cliente = "{Clientes.status} = '" & cmbStatus.Text & "' and " & TextoFiltroPadraoRel
        ElseIf cmbfiltrarpor = "Família" Then
                .StrSql_Cliente = INNERJOINTEXTO & " where CFF.Familia = '" & cmbfamilia & "' and CFF.tipo = 'C' and " & TextoFiltroPadrao
                .FormulaRel_Cliente = "{compras_fornecedores_familia.Familia} = '" & cmbfamilia & "' and {compras_fornecedores_familia.tipo} = 'C' and " & TextoFiltroPadraoRel
            ElseIf cmbfiltrarpor = "Grupo" Then
                    .StrSql_Cliente = INNERJOINTEXTO & " where CG.Texto = '" & cmbfamilia & "' and " & TextoFiltroPadrao
                    .FormulaRel_Cliente = "{Clientes_grupos.Texto} = '" & cmbfamilia & "' and " & TextoFiltroPadraoRel
                ElseIf cmbfiltrarpor = "CNPJ/CPF" Then
                        .StrSql_Cliente = INNERJOINTEXTO & " where " & CpfCnpj & " and " & TextoFiltroPadrao
                        .FormulaRel_Cliente = CPFCNPJRel & " and " & TextoFiltroPadraoRel
                    ElseIf cmbfiltrarpor = "Vendedor" Then
                            .StrSql_Cliente = "Select CL.*, VVC.*,VV.Vendedor from Clientes CL Inner join Vendas_Vendedores_Clientes VVC on VVC.IDCliente = CL.Id Inner Join Vendas_Vendedores VV on VVC.IDVendedor = VV.Id where VV.Vendedor = '" & cmbfamilia.Text & "'"
                            
                            .FormulaRel_Cliente = "{Vendas_Clientes_Vendedores.Vendedor} = '" & cmbfamilia & "' and " & TextoFiltroPadraoRel
                        'Debug.print .StrSql_Cliente
                        
                    ElseIf cmbfiltrarpor = "Código do cliente" Then
                            .StrSql_Cliente = INNERJOINTEXTO & " where C.IDCliente = " & txtTexto & " and " & TextoFiltroPadrao
                        Else
                            Select Case cmbfiltrarpor
                                Case "Razão social": TextoFiltro = "C.nomerazao"
                                Case "Nome fantasia": TextoFiltro = "C.nomefantasia"
                                Case "Cidade": TextoFiltro = "C.cidade"
                            End Select
                            .StrSql_Cliente = INNERJOINTEXTO & " where " & TextoFiltro & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and " & TextoFiltroPadrao
                            .FormulaRel_Cliente = "{" & Replace(TextoFiltro, "C.", "Clientes.") & "}" & FunVerifTipoFiltroIMFRel(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and " & TextoFiltroPadraoRel
        End If
    Else
        .StrSql_Cliente = INNERJOINTEXTO & " where " & TextoFiltroPadrao
        .FormulaRel_Cliente = TextoFiltroPadraoRel
    End If
    'Debug.print .StrSql_Cliente
    
    .ProcCarregaLista (1)
End With
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

'ProcCarregaToolBar1 Me, 9285, 5, True
cmbfiltrarpor = "Razão social"
optTodos.Value = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optFisica_Click()
On Error GoTo tratar_erro

If optFisica.Value = True And cmbfiltrarpor = "CNPJ/CPF" Then
    txtTexto.Visible = False
    txtTexto = ""
    cmbfamilia.Visible = False
    cmbfamilia.ListIndex = -1
    cmbStatus.Visible = False
    cmbStatus.ListIndex = -1
    txtcnpj.Visible = False
    txtcnpj.Text = "__.___.___/____-__"
    txtCpf.Visible = True
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optJuridica_Click()
On Error GoTo tratar_erro

If optJuridica.Value = True And cmbfiltrarpor = "CNPJ/CPF" Then
    txtTexto.Visible = False
    txtTexto = ""
    cmbfamilia.Visible = False
    cmbfamilia.ListIndex = -1
    cmbStatus.Visible = False
    cmbStatus.ListIndex = -1
    txtcnpj.Visible = True
    txtCpf.Visible = False
    txtCpf.Text = "___.___.___-__"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtTexto_Change()
On Error GoTo tratar_erro

If txtTexto <> "" Then
    cmbfamilia.ListIndex = -1
    cmbStatus.ListIndex = -1
    txtcnpj.Text = "__.___.___/____-__"
    txtCpf.Text = "___.___.___-__"
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtcnpj_Change()
On Error GoTo tratar_erro
  
If txtcnpj.Text <> "__.___.___/____-__" Then
    txtTexto.Text = ""
    cmbfamilia.ListIndex = -1
    cmbStatus.ListIndex = -1
    txtCpf.Text = "___.___.___-__"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtCpf_Change()
On Error GoTo tratar_erro
  
If txtCpf.Text <> "___.___.___-__" Then
    txtTexto.Text = ""
    cmbfamilia.ListIndex = -1
    cmbStatus.ListIndex = -1
    txtcnpj.Text = "__.___.___/____-__"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyEscape: Unload Me
    Case vbKeyF2: ProcFiltrar
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

