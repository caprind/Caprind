VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmFaturamento_Relatorios_menuimpressao 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Faturamento | Menu relatórios"
   ClientHeight    =   3855
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   4620
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   4620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   9
      Top             =   3450
      Width           =   4620
      _ExtentX        =   8149
      _ExtentY        =   714
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   435
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   4620
      _ExtentX        =   8149
      _ExtentY        =   741
      DibPicture      =   "frmFaturamento_Relatorios_menuimpressao.frx":0000
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
      Icon            =   "frmFaturamento_Relatorios_menuimpressao.frx":1C95
      ShowMaximize    =   0   'False
      ShowMinimize    =   0   'False
   End
   Begin VB.Frame Frame2 
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
      ForeColor       =   &H00800000&
      Height          =   885
      Left            =   660
      TabIndex        =   6
      Top             =   3510
      Visible         =   0   'False
      Width           =   3225
      Begin VB.CommandButton Cmd_imprimir 
         BackColor       =   &H00C0C0C0&
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
         Left            =   2715
         Picture         =   "frmFaturamento_Relatorios_menuimpressao.frx":1FAF
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Visualizar impressão (F5)"
         Top             =   390
         Width           =   315
      End
      Begin VB.ComboBox Cmb_nome_relatorio 
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
         ItemData        =   "frmFaturamento_Relatorios_menuimpressao.frx":20A5
         Left            =   180
         List            =   "frmFaturamento_Relatorios_menuimpressao.frx":20A7
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "Opções de relatório."
         Top             =   390
         Width           =   2535
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome do relatório"
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
         Left            =   690
         TabIndex        =   7
         Top             =   180
         Width           =   1515
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
      Height          =   2745
      Left            =   660
      TabIndex        =   5
      Top             =   570
      Width           =   3225
      Begin DrawSuite2022.USButton Cmd_padrao 
         Height          =   780
         Left            =   180
         TabIndex        =   0
         Top             =   180
         Width           =   2865
         _ExtentX        =   5054
         _ExtentY        =   1376
         DibPicture      =   "frmFaturamento_Relatorios_menuimpressao.frx":20A9
         BorderColor     =   4960354
         BorderColorDisabled=   13160660
         BorderColorDown =   4210752
         BorderColorOver =   49152
         Caption         =   "Padrão"
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
         PicSize         =   2
         PicSizeH        =   24
         PicSizeW        =   24
         ShowFocusRect   =   0   'False
         ShowFocusRectDown=   0   'False
         Theme           =   3
      End
      Begin DrawSuite2022.USButton Cmd_personalizado 
         Height          =   780
         Left            =   180
         TabIndex        =   1
         Top             =   990
         Width           =   2865
         _ExtentX        =   5054
         _ExtentY        =   1376
         DibPicture      =   "frmFaturamento_Relatorios_menuimpressao.frx":BB56
         BorderColor     =   5263559
         BorderColorDisabled=   13160660
         BorderColorDown =   4013465
         BorderColorOver =   4408288
         Caption         =   "Personalizado"
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
         PicSize         =   2
         PicSizeH        =   24
         PicSizeW        =   24
         ShowFocusRect   =   0   'False
         ShowFocusRectDown=   0   'False
         Theme           =   4
      End
      Begin DrawSuite2022.USButton Cmd_grafico 
         Height          =   780
         Left            =   180
         TabIndex        =   2
         Top             =   1800
         Width           =   2865
         _ExtentX        =   5054
         _ExtentY        =   1376
         DibPicture      =   "frmFaturamento_Relatorios_menuimpressao.frx":15D03
         BorderColor     =   1154291
         BorderColorDisabled=   13160660
         BorderColorDown =   16576
         BorderColorOver =   8438015
         Caption         =   "Gráfico"
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
         GradientColor1  =   1154291
         GradientColor2  =   1154291
         GradientColor3  =   1154291
         GradientColor4  =   1154291
         GradientColorDisabled1=   14215660
         GradientColorDisabled2=   14215660
         GradientColorDisabled3=   14215660
         GradientColorDisabled4=   14215660
         GradientColorDown1=   16576
         GradientColorDown2=   16576
         GradientColorDown3=   16576
         GradientColorDown4=   16576
         GradientColorOver1=   8438015
         GradientColorOver2=   8438015
         GradientColorOver3=   8438015
         GradientColorOver4=   8438015
         PicAlign        =   7
         PicSize         =   2
         PicSizeH        =   24
         PicSizeW        =   24
         ShowFocusRect   =   0   'False
         ShowFocusRectDown=   0   'False
         Theme           =   5
      End
   End
End
Attribute VB_Name = "frmFaturamento_Relatorios_menuimpressao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cmd_grafico_Click()
On Error GoTo tratar_erro

Frame2.Visible = False
frmFaturamento_Relatorios_menuimpressao.Height = 5000
With frmFaturamento_Relatorios
    If .Opt_individual.Value = True Then
        If .optDetalhado.Value = True Then
            NomeRel = "Faturamento_relatorio_individual_detalhado grafico.rpt"
        Else
            USMsgBox ("Não existe relatório disponível para esta pesquisa."), vbExclamation, "CAPRIND v5.0"
            Exit Sub
        End If
    Else
        NomeRel = "Faturamento_relatorio_comparativo_resumido grafico.rpt"
    End If
End With
ProcImprimirRelGrafico "{Producao_Relatorios.Responsavel}= '" & pubUsuario & "' and {Producao_Relatorios.Modulo} = '" & Formulario & "' and {Producao_Relatorios_Total.Responsavel}= '" & pubUsuario & "' and {Producao_Relatorios_Total.Modulo} = '" & Formulario & "'", ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_padrao_Click()
On Error GoTo tratar_erro

Frame2.Visible = False
frmFaturamento_Relatorios_menuimpressao.Height = 3850
With frmFaturamento_Relatorios
    If .Opt_individual.Value = True Then
        If .optDetalhado.Value = True Then
            NomeRel = "Faturamento_relatorio_individual_detalhado.rpt"
        Else
            NomeRel = "Faturamento_relatorio_individual_resumido.rpt"
        End If
    Else
        NomeRel = "Faturamento_relatorio_comparativo_resumido.rpt"
    End If
    TabelaRel = 0
    TabelaRel1 = 0
    TabelaRel2 = 0
    CampoRel = 0
    CampoRel1 = 0
    CampoRel2 = 0
    If .chkVlrTotal.Value = 1 Then
        TabelaRel = 1
        CampoRel = 47
        OrdenarRel = 1
        TabelaRel1 = 1
        CampoRel1 = 18
        OrdenarRel1 = 0
    Else
        TabelaRel = 1
        CampoRel = 18
        OrdenarRel = 0
    End If
End With
ProcImprimirRelOrdenado "{Producao_Relatorios.Responsavel}= '" & pubUsuario & "' and {Producao_Relatorios.Modulo} = '" & Formulario & "' and {Producao_Relatorios_Total.Responsavel}= '" & pubUsuario & "' and {Producao_Relatorios_Total.Modulo} = '" & Formulario & "'", ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_imprimir_Click()
On Error GoTo tratar_erro

If Cmb_nome_relatorio = "" Then
    USMsgBox ("Informe o nome de relatório antes de visualizar impressão."), vbExclamation, "CAPRIND v5.0"
    Cmb_nome_relatorio.SetFocus
    Exit Sub
End If
Select Case Cmb_nome_relatorio
    Case "Período detalhado": NomeRel = "Faturamento_nota fiscal_periodo_detalhado.rpt"
    Case "Período resumido": NomeRel = "Faturamento_nota fiscal_periodo_resumido.rpt"
    Case "Totais resumido": NomeRel = "Faturamento_nota fiscal_totais_resumido.rpt"
End Select
ProcImprimirRel "{Producao_Relatorios.Responsavel}= '" & pubUsuario & "' and {Producao_Relatorios.Modulo} = '" & Formulario & "' and {Producao_Relatorios_Total.Responsavel}= '" & pubUsuario & "' and {Producao_Relatorios_Total.Modulo} = '" & Formulario & "'", ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_personalizado_Click()
On Error GoTo tratar_erro

Permitido1 = True
With frmFaturamento_Relatorios
    If .Opt_individual.Value = True Then
        If .optDetalhado.Value = True Then
            If .cmbfiltrarpor = "Cliente/fornecedor" Or .cmbfiltrarpor = "Nota fiscal" Then Permitido1 = False
        Else
            Permitido1 = False
        End If
    Else
        If .cmbfiltrarpor <> "Nota fiscal" And .cmbfiltrarpor <> "Nota fiscal x Cliente/fornecedor" Then Permitido1 = False
    End If
End With
If Permitido1 = False Then
    USMsgBox ("Não existe relatório disponível para esta pesquisa."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Frame2.Visible = True
frmFaturamento_Relatorios_menuimpressao.Height = 5055

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyEscape: Unload Me
    Case vbKeyF5: If Frame2.Visible = True Then Cmd_imprimir_Click
End Select
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

With Cmb_nome_relatorio
    .Clear
    If frmFaturamento_Relatorios.Opt_individual.Value = True Then
        If frmFaturamento_Relatorios.cmbfiltrarpor <> "Cliente/fornecedor" And frmFaturamento_Relatorios.cmbfiltrarpor <> "Nota fiscal" Then
            .AddItem "Período detalhado"
            .AddItem "Período resumido"
            .AddItem "Totais resumido"
        End If
    ElseIf frmFaturamento_Relatorios.cmbfiltrarpor = "Nota fiscal" Or frmFaturamento_Relatorios.cmbfiltrarpor = "Nota fiscal x Cliente/fornecedor" Then
            .AddItem "Período resumido"
            .AddItem "Totais resumido"
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
