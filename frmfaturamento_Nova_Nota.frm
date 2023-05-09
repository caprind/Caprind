VERSION 5.00
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Begin VB.Form frmfaturamento_Nova_Nota 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Faturamento | Nota fiscal - Própria - Emissão"
   ClientHeight    =   5685
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7245
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
   Icon            =   "frmfaturamento_Nova_Nota.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   7245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   13
      Top             =   5280
      Width           =   7245
      _ExtentX        =   12779
      _ExtentY        =   714
   End
   Begin DrawSuite2022.USButton btnNovaNota 
      Height          =   1185
      Left            =   270
      TabIndex        =   11
      ToolTipText     =   "Criar nota fiscal manualmente"
      Top             =   2670
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   2090
      DibPicture      =   "frmfaturamento_Nova_Nota.frx":1042
      Caption         =   "Emitir nota fiscal eletrônica"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
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
      PicSize         =   4
      PicSizeH        =   48
      PicSizeW        =   48
      ShowFocusRect   =   0   'False
      Theme           =   4
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   465
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   7245
      _ExtentX        =   12779
      _ExtentY        =   820
      DibPicture      =   "frmfaturamento_Nova_Nota.frx":81C2
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
      Icon            =   "frmfaturamento_Nova_Nota.frx":11C6F
   End
   Begin VB.Frame Frame3 
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
      ForeColor       =   &H80000007&
      Height          =   1455
      Left            =   270
      TabIndex        =   5
      Top             =   570
      Width           =   6645
      Begin VB.TextBox txtSerie 
         Alignment       =   2  'Center
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
         Height          =   330
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   990
         Width           =   795
      End
      Begin VB.ComboBox Cmb_tipo_TBSN 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "frmfaturamento_Nova_Nota.frx":12CC1
         Left            =   180
         List            =   "frmfaturamento_Nova_Nota.frx":12CD4
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         ToolTipText     =   "Tabela do simples nacional."
         Top             =   1650
         Width           =   6345
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Tipo da nota"
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
         Height          =   525
         Left            =   150
         TabIndex        =   8
         Top             =   210
         Width           =   2265
         Begin VB.OptionButton optServico 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Serviços"
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
            Left            =   1200
            TabIndex        =   3
            Top             =   270
            Width           =   915
         End
         Begin VB.OptionButton optProduto 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Produtos"
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
            Left            =   180
            TabIndex        =   2
            Top             =   270
            Width           =   945
         End
      End
      Begin VB.ComboBox Cmb_modelo 
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
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "frmfaturamento_Nova_Nota.frx":12E09
         Left            =   1020
         List            =   "frmfaturamento_Nova_Nota.frx":12E70
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "Modelo da nota fiscal."
         Top             =   1005
         Width           =   5505
      End
      Begin VB.ComboBox Cmb_empresa 
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
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "frmfaturamento_Nova_Nota.frx":133D7
         Left            =   2640
         List            =   "frmfaturamento_Nova_Nota.frx":133D9
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Empresa."
         Top             =   390
         Width           =   3885
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Série"
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
         Left            =   390
         TabIndex        =   15
         Top             =   810
         Width           =   360
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Tabela do simples nacional"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   2370
         TabIndex        =   9
         Top             =   1440
         Width           =   1890
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Modelo"
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
         Left            =   3472
         TabIndex        =   7
         Top             =   810
         Width           =   510
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Empresa"
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
         Left            =   4215
         TabIndex        =   6
         Top             =   180
         Width           =   615
      End
   End
   Begin DrawSuite2022.USButton btnImportarxml 
      Height          =   1005
      Left            =   270
      TabIndex        =   12
      ToolTipText     =   "Importar nota por XML"
      Top             =   4050
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   1773
      DibPicture      =   "frmfaturamento_Nova_Nota.frx":133DB
      Caption         =   "Emitir nota fiscal por importação XML"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   4960354
      BorderColorDisabled=   13160660
      BorderColorDown =   4210752
      BorderColorOver =   49152
      GradientColor1  =   4960354
      GradientColor2  =   4960354
      GradientColor3  =   4960354
      GradientColor4  =   4960354
      GradientColorDisabled1=   14215660
      GradientColorDisabled2=   14215660
      GradientColorDisabled3=   14215660
      GradientColorDisabled4=   14215660
      GradientColorOver1=   49152
      GradientColorOver2=   49152
      GradientColorOver3=   49152
      GradientColorOver4=   49152
      GradientColorDown1=   32768
      GradientColorDown2=   32768
      GradientColorDown3=   32768
      GradientColorDown4=   32768
      PicAlign        =   7
      PicSize         =   3
      PicSizeH        =   32
      PicSizeW        =   32
      ShowFocusRect   =   0   'False
      Theme           =   3
   End
   Begin VB.Label lblXML 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Aguarde, importação XML sendo executada..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   390
      TabIndex        =   16
      Top             =   2190
      Visible         =   0   'False
      Width           =   6495
   End
End
Attribute VB_Name = "frmfaturamento_Nova_Nota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnImportarxml_Click()
On Error GoTo tratar_erro
strCaminho = ""

If USMsgBox("Deseja realmente importar o XML?", vbYesNo, "CAPRIND v5.0") = vbYes Then
ProcImportarXML
Unload Me
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub btnNovaNota_Click()
On Error GoTo tratar_erro

ProcNovo
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_empresa_Click()
On Error GoTo tratar_erro
IDempresa = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)

ProcVerificaTPNFe
txtSerie.Text = NF_Serie
If Formulario <> "Estoque/Ordem de faturamento" Then
frmFaturamento_Prod_Serv.txtEmpresa.Text = Cmb_empresa
Else
frmEstoque_Ordem_Faturamento.txtEmpresa.Text = Cmb_empresa
End If

ProcVerifTabelaSN IDempresa

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcGerarNumero()
On Error GoTo tratar_erro

With frmFaturamento_Prod_Serv
    If optProduto = True Then TipoNF = "M1" Else TipoNF = "SA"
    Set TBAbrir = CreateObject("adodb.recordset")
    StrSql = "Select CAST(int_NotaFiscal AS int) AS NF, Serie FROM tbl_Dados_Nota_Fiscal where Serie = '" & NF_Serie & "'and Modelo = '" & Left(Cmb_modelo.Text, 2) & "' and tipoNF = '" & TipoNF & "' and Aplicacao = 'P' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and int_NotaFiscal IS NOT NULL order by dt_DataEmissao desc, NF desc"
    'Debug.print StrSql
    
    TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        QuantsolicitadoN1 = TBAbrir!NF + 1
        FamiliaAntiga = QuantsolicitadoN1
        Familiatext = FunTamanhoTextoZeroEsq(FamiliaAntiga, 9)
        SerieNF = IIf(IsNull(TBAbrir!Serie), 1, TBAbrir!Serie)
    Else
        Familiatext = "000000001"
        SerieNF = NF_Serie
    End If
    .txtNFiscal.Text = FunVerifExisteNumNF(TipoNF, Cmb_empresa.ItemData(Cmb_empresa.ListIndex), Familiatext, SerieNF, Left(Cmb_modelo.Text, 2))
    .txtSerie = SerieNF
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_modelo_Click()
On Error GoTo tratar_erro

NFCe = False

If Left(Cmb_modelo.Text, 2) = "65" Then
    NFCe = True
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro
  
Select Case KeyCode
    Case vbKeyInsert: ProcNovo
    Case vbKeyEscape: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro
btnImportarxml.Visible = True

If Formulario = "Faturamento/Nota fiscal/Própria" Then
    USForm1.Caption = "Faturamento | Nota fiscal | Própria | Emissão"
            frmfaturamento_Nova_Nota.Height = 3870
            btnNovaNota.Top = 2130
    
ElseIf Formulario = "Faturamento/Nota fiscal/Terceiros" Then
        USForm1.Caption = "Faturamento | Nota fiscal | Terceiros | Emissão"
        frmfaturamento_Nova_Nota.Height = 5700
        lblXML.Caption = "Escolha uma opção abaixo."
        lblXML.Visible = True
        btnNovaNota.Top = 2670
        
    ElseIf Formulario = "Estoque/Ordem de faturamento" Then
            USForm1.Caption = "Estoque | Ordem de faturamento | Emissão"
            frmfaturamento_Nova_Nota.Height = 4700
            btnNovaNota.Caption = "Emitir ordem de faturamento"
            btnNovaNota.ToolTipText = "Emitir nova ordem de faturamento"
        Else
            USForm1.Caption = "Estoque - Nota fiscal - Nova"
            frmfaturamento_Nova_Nota.Height = 5700
End If

optProduto.Value = True
ProcCarregaComboEmpresa Cmb_empresa, False
Cmb_modelo = "55 - Nota Fiscal Eletrônica"
OutraMoeda = False
txtSerie.Text = NF_Serie

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcVerifTabelaSN(ID_empresa As Integer)
On Error GoTo tratar_erro

'Verifica se existe mais de uma tabela do simples cadastrada
With Cmb_tipo_TBSN
    Label1.Visible = False
    .Visible = False
    Frame3.Height = 1455

 '   Height = 2940
    If FunVerifRegimeEmpresa(ID_empresa) = 1 Then
        .Clear
        Contador = 0
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select Tabela FROM Impostos_TabelaDAS where ID_empresa = " & ID_empresa & " and Ativado = 1 group by Tabela", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Label1.Visible = True
            .Visible = True
            Frame3.Height = 2115
         '   Height = 3570
            
            Do While TBAbrir.EOF = False
                Select Case TBAbrir!Tabela
                    Case 1: .AddItem "Tabela I - Partilha do Simples Nacional – Comércio"
                    Case 2: .AddItem "Tabela II - Partilha do Simples Nacional - Indústria"
                    Case 3: .AddItem "Tabela III - Partilha do Simples Nacional - Serviços e Locação de Bens Móveis"
                    Case 4: .AddItem "Tabela IV - Partilha do Simples Nacional - Serviços"
                    Case 5: .AddItem "Tabela V - Partilha do Simples Nacional - Partilha do Simples Nacional - Receitas decorrentes da prestação de serviços relacionados no § 5º-I do art. 18 da LC 123/2016"
                End Select
                
                TabelaSN = TBAbrir!Tabela
                Contador = Contador + 1
                TBAbrir.MoveNext
            Loop
            If Contador = 1 Then
                Select Case TabelaSN
                    Case 1: .Text = "Tabela I - Partilha do Simples Nacional – Comércio"
                    Case 2: .Text = "Tabela II - Partilha do Simples Nacional - Indústria"
                    Case 3: .Text = "Tabela III - Partilha do Simples Nacional - Serviços e Locação de Bens Móveis"
                    Case 4: .Text = "Tabela IV - Partilha do Simples Nacional - Serviços"
                    Case 5: .Text = "Tabela V - Partilha do Simples Nacional - Partilha do Simples Nacional - Receitas decorrentes da prestação de serviços relacionados no § 5º-I do art. 18 da LC 123/2016"
                End Select
                .Locked = True
                .TabStop = False
            Else
                .Locked = False
                .TabStop = True
            End If
        End If
        TBAbrir.Close
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovo()
On Error GoTo tratar_erro

If FunVerificaProsseguir = False Then Exit Sub

If Formulario <> "Estoque/Ordem de faturamento" Then
With frmFaturamento_Prod_Serv
    .NF_enviada = False
    .ProcLimpaCamposProd
    .ProcLimpaCamposTotaisNota
    .ProcLimpaCamposServicos
    .ProcLimpaCamposDuplicata
    .ProcLimpaCamposTransp
    .cmbFinalidade_emissao.Text = "1 - Normal"
    .Cmb_consumidor.Text = "1 - Sim"
    
    If Faturamento_NF_Saida = True And Formulario <> "Estoque/Ordem de faturamento" Then
    ProcGerarNumero
    End If
    
    .txt_DtEmissao.Text = Format(Date, "dd/mm/yyyy")
    .txtSerie.Locked = False
    .txtSerie.TabStop = True
    .Cmb_modelo = Cmb_modelo
    
    If optProduto.Value = True Then
    .optProduto.Value = optProduto
    Else
    .OptServico.Value = OptServico
    End If
    
    
    '.RegimeEmpresa = FunVerifRegimeEmpresa(Cmb_empresa.ItemData(Cmb_empresa.ListIndex))
    If Label1.Visible = True Then
        If Left(Cmb_tipo_TBSN, 1) = "T" Then
            Select Case Mid(Cmb_tipo_TBSN, 8, 3)
                Case "I -": TabelaSN = 1
                Case "II ": TabelaSN = 2
                Case "III": TabelaSN = 3
                Case "IV ": TabelaSN = 4
            End Select
        Else
            TabelaSN = 6
        End If
    Else
        TabelaSN = 0
    End If
        
    .Frame1(6).Enabled = True
    .Novo_Nota = True
    
    Unload Me
    .cmdcliente_Click
End With
Else
With frmEstoque_Ordem_Faturamento
    .NF_enviada = False
    .ProcLimpaCamposProd
'    .ProcLimpaCamposTotaisNota
    .ProcLimpaCamposServicos
    .ProcLimpaCamposDuplicata
    .ProcLimpaCamposTransp
    .cmbFinalidade_emissao.Text = "1 - Normal"
    .Cmb_consumidor.Text = "1 - Sim"
    
    'If Faturamento_NF_Saida = True And Formulario <> "Estoque/Ordem de faturamento" Then ProcGerarNumero
    .txt_DtEmissao.Text = Format(Date, "dd/mm/yyyy")
    .txtSerie.Locked = False
    .txtSerie.TabStop = True
    .Cmb_modelo = Cmb_modelo
    
    If optProduto.Value = True Then
    .optProduto.Value = optProduto
    Else
    .OptServico.Value = OptServico
    End If
    
    
    '.RegimeEmpresa = FunVerifRegimeEmpresa(Cmb_empresa.ItemData(Cmb_empresa.ListIndex))
    If Label1.Visible = True Then
        If Left(Cmb_tipo_TBSN, 1) = "T" Then
            Select Case Mid(Cmb_tipo_TBSN, 8, 3)
                Case "I -": TabelaSN = 1
                Case "II ": TabelaSN = 2
                Case "III": TabelaSN = 3
                Case "IV ": TabelaSN = 4
            End Select
        Else
            TabelaSN = 6
        End If
    Else
        TabelaSN = 0
    End If
        
    .Frame1(6).Enabled = True
    .Novo_Nota = True
    
    Unload Me
    .cmdcliente_Click
End With
End If


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
End Sub

Private Function FunVerificaProsseguir() As Boolean
On Error GoTo tratar_erro

FunVerificaProsseguir = True
If Formulario = "Estoque/Ordem de faturamento" Then NomeCampo = "ordem de faturamento" Else NomeCampo = "nota fiscal"
If optProduto.Value = False And OptServico.Value = False Then
    USMsgBox ("Informe o tipo da " & NomeCampo & "."), vbExclamation, "CAPRIND v5.0"
    FunVerificaProsseguir = False
    Exit Function
End If
If Cmb_modelo = "" Then
    USMsgBox ("Informe o modelo da " & NomeCampo & "."), vbExclamation, "CAPRIND v5.0"
    Cmb_modelo.SetFocus
    FunVerificaProsseguir = False
    Exit Function
End If

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function
