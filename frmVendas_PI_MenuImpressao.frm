VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmVendas_PI_MenuImpressao 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Vendas - Proposta comercial  Menu impressão"
   ClientHeight    =   3615
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   6540
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
   ScaleHeight     =   3615
   ScaleWidth      =   6540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   435
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   6540
      _ExtentX        =   11536
      _ExtentY        =   767
      DibPicture      =   "frmVendas_PI_MenuImpressao.frx":0000
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
      Icon            =   "frmVendas_PI_MenuImpressao.frx":1C95
      ShowMaximize    =   0   'False
      ShowMinimize    =   0   'False
   End
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   15
      Top             =   3210
      Width           =   6540
      _ExtentX        =   11536
      _ExtentY        =   714
   End
   Begin DrawSuite2022.USButton Cmd_avancar 
      Height          =   975
      Left            =   180
      TabIndex        =   13
      Top             =   2160
      Width           =   3105
      _ExtentX        =   5477
      _ExtentY        =   1720
      DibPicture      =   "frmVendas_PI_MenuImpressao.frx":1FAF
      BorderColor     =   4960354
      BorderColorDisabled=   13160660
      BorderColorDown =   4210752
      BorderColorOver =   49152
      Caption         =   "Opções >>>>"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
      ForeColorDown   =   16777215
      ForeColorOver   =   16777215
      GradientColor1  =   4960354
      GradientColor2  =   4960354
      GradientColor3  =   4960354
      GradientColor4  =   4960354
      GradientColorDisabled1=   14215660
      GradientColorDisabled2=   14215660
      GradientColorDisabled3=   14215660
      GradientColorDisabled4=   14215660
      GradientColorDown1=   32768
      GradientColorDown2=   32768
      GradientColorDown3=   32768
      GradientColorDown4=   32768
      GradientColorOver1=   49152
      GradientColorOver2=   49152
      GradientColorOver3=   49152
      GradientColorOver4=   49152
      PicAlign        =   7
      PicSize         =   3
      PicSizeH        =   32
      PicSizeW        =   32
      ShowFocusRect   =   0   'False
      ShowFocusRectDown=   0   'False
      Theme           =   3
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Opções para impressão"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1545
      Left            =   210
      TabIndex        =   12
      Top             =   540
      Width           =   6195
      Begin VB.CheckBox chkPersonalizado 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Imprimir proposta(s) personalizada"
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
         TabIndex        =   1
         Top             =   555
         Width           =   4155
      End
      Begin VB.CheckBox chkImprimir_alteracao 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Imprimir alterações"
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
         TabIndex        =   3
         Top             =   1050
         Width           =   4155
      End
      Begin VB.CheckBox Chk_visualizar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Visualizando impressão"
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
         Left            =   180
         TabIndex        =   4
         Top             =   1290
         Width           =   2025
      End
      Begin VB.CheckBox Chk_Imprimir 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Imprimir proposta(s)"
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
         TabIndex        =   0
         Top             =   300
         Width           =   4155
      End
      Begin VB.CheckBox Chk_Imprimir2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Imprimir proposta(s) resumida"
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
         TabIndex        =   2
         Top             =   795
         Width           =   4155
      End
   End
   Begin VB.Frame Frame1 
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
      ForeColor       =   &H00000000&
      Height          =   705
      Left            =   210
      TabIndex        =   10
      Top             =   3210
      Visible         =   0   'False
      Width           =   6195
      Begin VB.ComboBox Cmb_rev_de 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
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
         Left            =   2580
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   7
         ToolTipText     =   "Número da revisão."
         Top             =   240
         Width           =   555
      End
      Begin VB.ComboBox Cmb_rev_ate 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
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
         Left            =   5460
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   9
         ToolTipText     =   "Número da revisão."
         Top             =   240
         Width           =   555
      End
      Begin VB.ComboBox Cmb_Ate 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
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
         Left            =   3690
         Style           =   2  'Dropdown List
         TabIndex        =   8
         ToolTipText     =   "Número da proposta."
         Top             =   240
         Width           =   1755
      End
      Begin VB.ComboBox Cmb_De 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
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
         ItemData        =   "frmVendas_PI_MenuImpressao.frx":8293
         Left            =   810
         List            =   "frmVendas_PI_MenuImpressao.frx":8295
         Style           =   2  'Dropdown List
         TabIndex        =   6
         ToolTipText     =   "Número da proposta."
         Top             =   240
         Width           =   1755
      End
      Begin VB.OptionButton Opt_De 
         BackColor       =   &H00E0E0E0&
         Caption         =   "De :"
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
         TabIndex        =   5
         Top             =   330
         Width           =   615
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Até :"
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
         Index           =   0
         Left            =   3285
         TabIndex        =   11
         Top             =   330
         Width           =   360
      End
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   210
      TabIndex        =   14
      Top             =   3990
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor2      =   0
      SearchText      =   ""
      Value           =   0
   End
   Begin DrawSuite2022.USButton btnImprimir 
      Height          =   975
      Left            =   3300
      TabIndex        =   17
      Top             =   2160
      Width           =   3105
      _ExtentX        =   5477
      _ExtentY        =   1720
      DibPicture      =   "frmVendas_PI_MenuImpressao.frx":8297
      BorderColor     =   5263559
      BorderColorDisabled=   13160660
      BorderColorDown =   4013465
      BorderColorOver =   4408288
      Caption         =   "Relatório"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
      ForeColorDown   =   16777215
      ForeColorOver   =   16777215
      GradientColor1  =   5263559
      GradientColor2  =   5263559
      GradientColor3  =   5263559
      GradientColor4  =   5263559
      GradientColorDisabled1=   13160660
      GradientColorDisabled2=   13160660
      GradientColorDisabled3=   13160660
      GradientColorDisabled4=   13160660
      GradientColorDown1=   4013465
      GradientColorDown2=   4013465
      GradientColorDown3=   4013465
      GradientColorDown4=   4013465
      GradientColorOver1=   4408288
      GradientColorOver2=   4408288
      GradientColorOver3=   4408288
      GradientColorOver4=   4408288
      PicAlign        =   7
      PicSize         =   3
      PicSizeH        =   32
      PicSizeW        =   32
      ShowFocusRect   =   0   'False
      ShowFocusRectDown=   0   'False
      Theme           =   4
   End
End
Attribute VB_Name = "frmVendas_PI_MenuImpressao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FormulaRel_Vendas_PI As String 'OK
Dim TipoFiltro As String

Private Sub btnImprimir_Click()
On Error GoTo tratar_erro

 ProcRelatorio

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_Imprimir_Click()
On Error GoTo tratar_erro

ProcHabDesabVisualizandoImpressao

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_Imprimir2_Click()
On Error GoTo tratar_erro

ProcHabDesabVisualizandoImpressao

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcHabDesabVisualizandoImpressao()
On Error GoTo tratar_erro

If Chk_Imprimir.Value = 0 And Chk_Imprimir2.Value = 0 And chkImprimir_alteracao.Value = 0 And chkPersonalizado.Value = 0 Then
    Chk_visualizar.Value = 0
    Chk_visualizar.Enabled = False
Else
    Chk_visualizar.Enabled = True
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkImprimir_alteracao_Click()
On Error GoTo tratar_erro

ProcHabDesabVisualizandoImpressao
With Cmd_avancar
    If chkImprimir_alteracao.Value = 1 Then
        Avancar = False
        .Caption = "Avançar >>>>"
        .Enabled = False
        Height = 3000
        Frame1.Visible = False
        PBLista.Visible = False
        Opt_De.Value = False
    Else
        .Enabled = True
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkPersonalizado_Click()
On Error GoTo tratar_erro

ProcHabDesabVisualizandoImpressao

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_Ate_Click()
On Error GoTo tratar_erro

With Cmb_rev_ate
    .Clear
    .Enabled = True
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from Vendas_proposta where Ncotacao = '" & Cmb_Ate & "' and " & TipoFiltro & " order by Revisao", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Do While TBAbrir.EOF = False
            .AddItem TBAbrir!Revisao
            .ItemData(.NewIndex) = TBAbrir!ordenarproposta
            TBAbrir.MoveNext
        Loop
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_De_Click()
On Error GoTo tratar_erro

With Cmb_rev_de
    .Clear
    .Enabled = True
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from Vendas_proposta where Ncotacao = '" & Cmb_de & "' and " & TipoFiltro & " order by Revisao", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Do While TBAbrir.EOF = False
            .AddItem TBAbrir!Revisao
            .ItemData(.NewIndex) = TBAbrir!ordenarproposta
            TBAbrir.MoveNext
        Loop
    End If
    Cmb_Ate.Clear
    Cmb_rev_ate.Clear
    Cmb_Ate.Enabled = False
    Cmb_rev_ate.Enabled = False
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcRelatorio()
On Error GoTo tratar_erro

Acao = "visualizar impressão"
If Chk_Imprimir.Value = 0 And Chk_Imprimir2.Value = 0 And chkImprimir_alteracao.Value = 0 And chkPersonalizado.Value = 0 Then
    NomeCampo = "uma das opções"
    ProcVerificaAcao
    Exit Sub
End If
If Avancar = True Then
    NomeCampo = IIf(Vendas_Proposta = True, "o número da proposta", "o número do pedido")
    If Opt_De.Value = True And Cmb_de = "" Then
        ProcVerificaAcao
        Cmb_de.SetFocus
        Exit Sub
    End If
    If Opt_De.Value = True And Cmb_Ate = "" Then
        ProcVerificaAcao
        Cmb_Ate.SetFocus
        Exit Sub
    End If
    NomeCampo = "o número da revisão"
    If Opt_De.Value = True And Cmb_de <> "" And Cmb_rev_de = "" Then
        ProcVerificaAcao
        Cmb_rev_de.SetFocus
        Exit Sub
    End If
    If Opt_De.Value = True And Cmb_Ate <> "" And Cmb_rev_ate = "" Then
        ProcVerificaAcao
        Cmb_rev_ate.SetFocus
        Exit Sub
    End If
    
    If Opt_De.Value = True Then
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select cotacao, OrdenarProposta from Vendas_proposta where ordenarproposta >= " & Cmb_rev_de.ItemData(Cmb_rev_de.ListIndex) & " and ordenarproposta <= " & Cmb_rev_ate.ItemData(Cmb_rev_ate.ListIndex) & " and " & TipoFiltro & " Group by cotacao, OrdenarProposta, Revisao order by OrdenarProposta, Revisao", Conexao, adOpenKeyset, adLockReadOnly
        If TBAbrir.EOF = False Then
            PBLista.Min = 0
            PBLista.Max = TBAbrir.RecordCount
            PBLista.Value = 1
            Contador1 = 0
            Do While TBAbrir.EOF = False
                If Chk_Imprimir.Value = 1 Then If Vendas_Proposta = True Then ProcImprimirPropostaAssinatura False Else ProcImprimirPI False
                If chkPersonalizado.Value = 1 Then If Vendas_Proposta = True Then ProcImprimirPropostaAssinatura True Else ProcImprimirPI True
                If Chk_Imprimir2.Value = 1 Then If Vendas_Proposta = True Then ProcImprimirPropostaResumida Else ProcImprimirCheck
                'If chkImprimir_alteracao.Value = 1 Then ProcImprimirAlteracao
                TBAbrir.MoveNext
                Contador1 = Contador1 + 1
                PBLista.Value = Contador1
            Loop
        End If
        TBAbrir.Close
    End If
ElseIf chkImprimir_alteracao.Value = 1 Then
        ProcImprimirAlteracao
    Else
        If Vendas_PI = True Then
            IDlista = IIf(frmVendas_PI.txtID = "", 0, frmVendas_PI.txtID)
            NomeCampo = "o pedido"
        Else
            IDlista = IIf(frmVendas_proposta.txtID = "", 0, frmVendas_proposta.txtID)
            NomeCampo = "a proposta"
        End If
        If IDlista = 0 Then
            ProcVerificaAcao
            Unload Me
            frmVendas_PI_lista.Show 1
            Exit Sub
        End If
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from Vendas_proposta where Cotacao = " & IDlista, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            If Chk_Imprimir.Value = 1 Then If Vendas_Proposta = True Then ProcImprimirPropostaAssinatura False Else ProcImprimirPI False
            If chkPersonalizado.Value = 1 Then If Vendas_Proposta = True Then ProcImprimirPropostaAssinatura True Else ProcImprimirPI True
            If Chk_Imprimir2.Value = 1 Then If Vendas_Proposta = True Then ProcImprimirPropostaResumida Else ProcImprimirCheck
        End If
        TBAbrir.Close
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimirPI(Personalizado_PI As Boolean)
On Error GoTo tratar_erro

If Personalizado_PI = True Then
    NomeRel = "Vendas_pedidointerno_personalizado.rpt"
Else
    NomeRel = "Vendas_pedidointerno.rpt"
End If
If Avancar = False Then
    If Vendas_PI = True Then IDlista = frmVendas_PI.txtID Else IDlista = frmVendas_proposta.txtID
    FormulaRel_Vendas_PI = "{Vendas_proposta.Cotacao} = " & IDlista
Else
    IDlista = TBAbrir!Cotacao
    FormulaRel_Vendas_PI = "{Vendas_proposta.Cotacao} = " & TBAbrir!Cotacao
End If
FormulaRel_Vendas_PI = FormulaRel_Vendas_PI & " and ({Vendas_carteira.Liberacao} = 'VENDIDA' or {Vendas_carteira.Liberacao} = 'REVISADA' or {Vendas_carteira.Liberacao} = 'FATURAR' or {Vendas_carteira.Liberacao} = 'FATURAR PARCIAL' or {Vendas_carteira.Liberacao} = 'FATURADO' or {Vendas_carteira.Liberacao} = 'FATURADO PARCIAL')"

Set TBItem = CreateObject("adodb.recordset")
TBItem.Open "Select VC.Ordem as OrdemVendas, P.ordem as OrdemProducao FROM (vendas_carteira VC INNER JOIN Producao_pedidos PP ON PP.IDCarteira = VC.Codigo) INNER JOIN Producao P ON P.Ordem = PP.Ordem and P.Desenho = VC.Desenho where VC.cotacao = " & IDlista, Conexao, adOpenKeyset, adLockOptimistic
If TBItem.EOF = False Then
    Do While TBItem.EOF = False
        TBItem!OrdemVendas = TBItem!OrdemProducao
        TBItem.Update
        TBItem.MoveNext
    Loop
End If
TBItem.Close

If Chk_visualizar.Value = 1 Then ProcImprimirRel FormulaRel_Vendas_PI, "" Else ProcImprimir FormulaRel_Vendas_PI, ""

Set TBItem = CreateObject("adodb.recordset")
TBItem.Open "Select Ordem from vendas_carteira where cotacao = " & IDlista, Conexao, adOpenKeyset, adLockOptimistic
If TBItem.EOF = False Then
    Do While TBItem.EOF = False
        TBItem!Ordem = Null
        TBItem.Update
        TBItem.MoveNext
    Loop
End If
TBItem.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimirCheck()
On Error GoTo tratar_erro

NomeRel = "Vendas_pedidointerno_check list.rpt"
If Avancar = False Then FormulaRel_Vendas_PI = "{Vendas_proposta.Cotacao} = " & frmVendas_PI.txtID Else FormulaRel_Vendas_PI = "{Vendas_proposta.Cotacao} = " & TBAbrir!Cotacao
FormulaRel_Vendas_PI = FormulaRel_Vendas_PI & " and (LEFT({Vendas_carteira.Liberacao}, 7) = 'VENDIDA' or LEFT({Vendas_carteira.Liberacao}, 7) = 'FATURAR' or LEFT({Vendas_carteira.Liberacao}, 8) = 'FATURADO')"
If Chk_visualizar.Value = 1 Then ProcImprimirRel FormulaRel_Vendas_PI, "" Else ProcImprimir FormulaRel_Vendas_PI, ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimir(FormulaRel As String, FormulaRelSubReport As String)
On Error GoTo tratar_erro

ProcVerifRelPersonalizado
            
If PermitidoRel = False Then LocalrelNovo = Localrel Else LocalrelNovo = LocalRelPersonalizado
Set Report = crAPP.OpenReport(LocalrelNovo & "\" & NomeRel, crptToPrinter)
'Login SQL
contador = Report.Database.Tables.Count
Do While contador > 0
    Set DBTable = Report.Database.Tables(contador)
    ProcLogonBDSQL
    contador = contador - 1
Loop
ProcVerifSubReport FormulaRelSubReport

Report.FormulaSyntax = crCrystalSyntaxFormula 'Configura a sintaxe da formula
Report.RecordSelectionFormula = FormulaRel 'Formula de seleção do relatório
Report.PrintOut False 'Configura a seleção de impressora com false, enviando para impressora padrão
Set Report = Nothing 'Cancela a variavel report
Set crAPP = Nothing 'Cancela a variavel report

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_rev_de_Click()
On Error GoTo tratar_erro

With Cmb_Ate
    .Clear
    .Enabled = True
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select Ncotacao, ordenarproposta from Vendas_proposta where ordenarproposta >= " & Cmb_rev_de.ItemData(Cmb_rev_de.ListIndex) & " and " & TipoFiltro & " Group by Ncotacao, ordenarproposta order by ordenarproposta", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Do While TBAbrir.EOF = False
            .AddItem TBAbrir!Ncotacao
            TBAbrir.MoveNext
        Loop
    End If
    TBAbrir.Close
    Cmb_rev_ate.Clear
    Cmb_rev_ate.Enabled = False
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_avancar_Click()
On Error GoTo tratar_erro

With chkImprimir_alteracao
    If Avancar = False Then
        Avancar = True
        Cmd_avancar.Caption = "<<<< Recuar"
        Height = 4725
        Frame1.Visible = True
        PBLista.Visible = True
        Opt_De.Value = True
        
        .Value = 0
        .Enabled = False
    Else
        Avancar = False
        Cmd_avancar.Caption = "Avançar >>>>"
        Height = 3615
        Frame1.Visible = False
        PBLista.Visible = False
        Opt_De.Value = False
        
        .Enabled = True
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
    Case vbKeyF5: ProcRelatorio
    Case vbKeyEscape: Unload Me
End Select
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

Avancar = False
Height = 3615
If Vendas_Proposta = True Then
    TipoFiltro = "(Tipo = 'PR' or Tipo = 'PRPE')"
Else
    Caption = "Vendas - Pedido interno - Menu impressão"
    Chk_Imprimir.Caption = "Imprimir pedido(s)"
    chkPersonalizado.Caption = "Imprimir pedido(s) personalizado"
    Chk_Imprimir2.Caption = "Imprimir check list"
    TipoFiltro = "(Tipo = 'PE' or tipo = 'PRPE')"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt_De_Click()
On Error GoTo tratar_erro

If Opt_De.Value = True Then
    With Cmb_de
        Cmb_Ate.Clear
        .Clear
        .Enabled = True
        .SetFocus
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select Ncotacao, ordenarproposta from Vendas_proposta where " & TipoFiltro & " Group by Ncotacao, ordenarproposta order by Ordenarproposta desc", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Do While TBAbrir.EOF = False
                .AddItem TBAbrir!Ncotacao
                .ItemData(.NewIndex) = TBAbrir!ordenarproposta
                TBAbrir.MoveNext
            Loop
        End If
        TBAbrir.Close
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
    Case 1: ProcRelatorio
    'Case 3: ProcAjuda
    Case 4: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimirPropostaAssinatura(Personalizado_proposta As Boolean)
On Error GoTo tratar_erro

If Personalizado_proposta = True Then
    NomeRel = "Vendas_proposta_personalizado.rpt"
Else
    NomeRel = "Vendas_proposta.rpt"
End If

ID_empresa = frmVendas_proposta.Cmb_empresa.ItemData(frmVendas_proposta.Cmb_empresa.ListIndex)

If Avancar = False Then FormulaRel_Vendas_PI = "{Vendas_proposta.Cotacao} = " & frmVendas_proposta.txtID & " And {Empresa.codigo}= " & ID_empresa Else FormulaRel_Vendas_PI = "{Vendas_proposta.Cotacao} = " & TBAbrir!Cotacao & " And {Empresa.codigo}= " & ID_empresa
If Chk_visualizar.Value = 1 Then ProcImprimirRel FormulaRel_Vendas_PI, "" Else ProcImprimir FormulaRel_Vendas_PI, ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimirAlteracao()
On Error GoTo tratar_erro

NomeRel = "Vendas_Alteracoes.rpt"
'If Avancar = False Then
    If Vendas_PI = True Then FormulaRel_Vendas_PI = frmVendas_PI.StrSql_PI_LocalizarRel Else FormulaRel_Vendas_PI = frmVendas_proposta.StrSql_Proposta_LocalizarRel
'Else
    'If Vendas_PI = True Then TextoFiltroAltRel = " and {vendas_carteira_alteracoes.Tipo} = 'VPI'" Else TextoFiltroAltRel = " and {vendas_carteira_alteracoes.Tipo} = 'VPR'"
    'TextoFiltroAltRel = TextoFiltroAltRel & " and Not(IsNull({vendas_carteira_alteracoes.ID}))"
    'FormulaRel_Vendas_PI = "{Vendas_proposta.Cotacao} = " & TBAbrir!Cotacao & TextoFiltroAltRel
'End If
If Chk_visualizar.Value = 1 Then ProcImprimirRel FormulaRel_Vendas_PI, "" Else ProcImprimir FormulaRel_Vendas_PI, ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimirPropostaResumida()
On Error GoTo tratar_erro

NomeRel = "Vendas_proposta_resumido.rpt"
If Avancar = False Then FormulaRel_Vendas_PI = "{Vendas_proposta.Cotacao} = " & frmVendas_proposta.txtID Else FormulaRel_Vendas_PI = "{Vendas_proposta.Cotacao} = " & TBAbrir!Cotacao
If Chk_visualizar.Value = 1 Then ProcImprimirRel FormulaRel_Vendas_PI, "" Else ProcImprimir FormulaRel_Vendas_PI, ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
