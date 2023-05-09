VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmQualidade_Relatorios_historico 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Qualidade - Relatórios - Histórico"
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5145
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   5145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Cmb_empresa 
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
      ItemData        =   "frmQualidade_Relatorios_historico.frx":0000
      Left            =   1080
      List            =   "frmQualidade_Relatorios_historico.frx":0002
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Empresa."
      Top             =   1080
      Width           =   3825
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   2205
      Left            =   60
      TabIndex        =   5
      Top             =   1440
      Width           =   5055
      Begin VB.TextBox txtCodigoInterno 
         Enabled         =   0   'False
         Height          =   345
         Left            =   2400
         TabIndex        =   15
         Top             =   1620
         Width           =   2505
      End
      Begin VB.OptionButton opt5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Controle de processo"
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
         Height          =   255
         Left            =   180
         TabIndex        =   14
         Top             =   1680
         Width           =   2415
      End
      Begin VB.ComboBox Cmb_mes_de 
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
         ItemData        =   "frmQualidade_Relatorios_historico.frx":0004
         Left            =   3480
         List            =   "frmQualidade_Relatorios_historico.frx":002C
         Style           =   2  'Dropdown List
         TabIndex        =   11
         ToolTipText     =   "Mês de."
         Top             =   420
         Width           =   675
      End
      Begin VB.ComboBox Cmb_mes_ate 
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
         ItemData        =   "frmQualidade_Relatorios_historico.frx":006D
         Left            =   3480
         List            =   "frmQualidade_Relatorios_historico.frx":0095
         Style           =   2  'Dropdown List
         TabIndex        =   10
         ToolTipText     =   "Mês até."
         Top             =   810
         Width           =   675
      End
      Begin VB.ComboBox Cmb_ano_ate 
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
         ItemData        =   "frmQualidade_Relatorios_historico.frx":00D6
         Left            =   4170
         List            =   "frmQualidade_Relatorios_historico.frx":00D8
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Ano até."
         Top             =   810
         Width           =   795
      End
      Begin VB.ComboBox Cmb_ano_de 
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
         ItemData        =   "frmQualidade_Relatorios_historico.frx":00DA
         Left            =   4170
         List            =   "frmQualidade_Relatorios_historico.frx":00DC
         Style           =   2  'Dropdown List
         TabIndex        =   8
         ToolTipText     =   "Ano de."
         Top             =   420
         Width           =   795
      End
      Begin VB.OptionButton Opt4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Custo de retrabalho"
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
         Height          =   255
         Left            =   180
         TabIndex        =   4
         Top             =   960
         Width           =   2415
      End
      Begin VB.OptionButton Opt3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Custos de falhas externa"
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
         Height          =   255
         Left            =   180
         TabIndex        =   3
         Top             =   720
         Width           =   2415
      End
      Begin VB.OptionButton Opt1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Inpeção final"
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
         Height          =   255
         Left            =   180
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton Opt2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Inspeção de recebimento"
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
         Height          =   255
         Left            =   180
         TabIndex        =   2
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Codigo Interno"
         Height          =   225
         Left            =   3150
         TabIndex        =   16
         Top             =   1380
         Width           =   1215
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
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
         Height          =   285
         Left            =   3030
         TabIndex        =   13
         Top             =   810
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Height          =   285
         Left            =   3090
         TabIndex        =   12
         Top             =   420
         Width           =   300
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   6
      Top             =   0
      Width           =   5025
      _ExtentX        =   8864
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
      ButtonCaption1  =   "Relatório"
      ButtonEnabled1  =   0   'False
      ButtonIconSize1 =   32
      ButtonToolTipText1=   "Relatório (F5)"
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
      ButtonWidth1    =   51
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
      ButtonLeft2     =   55
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
      ButtonLeft3     =   59
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
      ButtonLeft4     =   97
      ButtonTop4      =   2
      ButtonWidth4    =   26
      ButtonHeight4   =   21
      ButtonUseMaskColor4=   0   'False
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
      ButtonLeft5     =   125
      ButtonTop5      =   2
      ButtonWidth5    =   24
      ButtonHeight5   =   24
      ButtonUseMaskColor5=   0   'False
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   2520
         Top             =   180
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmQualidade_Relatorios_historico.frx":00DE
         Count           =   1
      End
   End
   Begin VB.Label Label44 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Empresa :"
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
      Left            =   180
      TabIndex        =   7
      Top             =   1080
      Width           =   825
   End
End
Attribute VB_Name = "frmQualidade_Relatorios_historico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cmb_ano_de_Click()
On Error GoTo tratar_erro

Cmb_ano_ate = Cmb_ano_de

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_mes_ate_Click()
On Error GoTo tratar_erro

If Cmb_mes_de <> "" And Cmb_mes_ate <> "" Then
    qt = FunVerificaMes(Cmb_mes_de)
    Qtd = FunVerificaMes(Cmb_mes_ate)
    If Qtd < qt Then
        USMsgBox ("O mês final não pode ser menor que o mês inicial."), vbExclamation, "CAPRIND v5.0"
        Cmb_mes_ate = Cmb_mes_de
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_mes_de_Click()
On Error GoTo tratar_erro

If Cmb_mes_de <> "" And Cmb_mes_ate <> "" Then
    qt = FunVerificaMes(Cmb_mes_de)
    Qtd = FunVerificaMes(Cmb_mes_ate)
    If qt > Qtd Then
        USMsgBox ("O mês inicial não pode ser maior que o mês final."), vbExclamation, "CAPRIND v5.0"
        Cmb_mes_de = Cmb_mes_ate
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyF5: ProcImprimir
    'Case vbKeyF1: ProcAjuda
    Case vbKeyEscape: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 4845, 4, True
Formulario = "Qualidade/Relatórios/Histórico"
Direitos
ProcLimpaVariaveisPrincipais
ProcCarregaComboEmpresa Cmb_empresa, True
ProcCarregaComboAno Cmb_ano_ate, "2005", 1
ProcCarregaComboAno Cmb_ano_de, "2005", 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro

Acao = "visualizar impressão"
If Cmb_mes_de = "" Then
    NomeCampo = "o mês"
    ProcVerificaAcao
    Cmb_mes_de.SetFocus
    Exit Sub
End If
If Cmb_ano_de = "" Then
    NomeCampo = "o ano"
    ProcVerificaAcao
    Cmb_ano_de.SetFocus
    Exit Sub
End If
If Cmb_mes_de = "" Then
    NomeCampo = "o mês"
    ProcVerificaAcao
    Cmb_mes_de.SetFocus
    Exit Sub
End If
If opt1.Value = True Then
    NomeRel = "CQ_relatorio_inspecao_final.rpt"
    NomeView = "Qualidade_relatorio_inspecao_final"
ElseIf opt2.Value = True Then
        NomeRel = "CQ_relatorio_inspecao_recebimento.rpt"
        NomeView = "Qualidade_relatorio_inspecao_recebimento"
    ElseIf opt3.Value = True Then
            NomeRel = "CQ_relatorio_devolucao_clientes.rpt"
            NomeView = "Qualidade_relatorio_devolucao_clientes"
        ElseIf opt4.Value = True Then
            NomeRel = "CQ_relatorio_custo_retrabalho.rpt"
            NomeView = "Qualidade_relatorio_custo_retrabalho"
            ElseIf opt5.Value = True Then
            
            Set TBLISTA = CreateObject("adodb.recordset")
            TBLISTA.Open ("Select * from Medicao where Desenho = '" & txtCodigoInterno & "'"), Conexao, adOpenKeyset, adLockReadOnly
            If TBLISTA.EOF = True Then
                MsgBox ("Não foi encontrado nenhum controle de medição para o codigo informado."), vbInformation + vbOKOnly
                Exit Sub
            End If
                
            NomeRel = "CQ_ControleProcesso.rpt"
            ProcImprimirRel "{Medicao.desenho} = '" & txtCodigoInterno.Text & "' and Month({Medicao.Data}) >= " & Cmb_mes_de.ListIndex + 1 & " and Month({Medicao.Data}) <= " & Cmb_mes_ate.ListIndex + 1 & " and Year({Medicao.Data}) <= " & Cmb_ano_de & " and Year({Medicao.Data}) >= " & Cmb_ano_ate & "", """"
            Exit Sub
               
End If

ProcImprimirRel "{" & NomeView & ".ID_empresa} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and {" & NomeView & ".Mes} >= " & FunVerificaMes(Cmb_mes_de) & " and {" & NomeView & ".Ano} = " & Cmb_ano_de & " and {" & NomeView & ".Mes} <= " & FunVerificaMes(Cmb_mes_ate) & " and {" & NomeView & ".Ano} = " & Cmb_ano_ate, ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt5_Click()
On Error GoTo tratar_erro

txtCodigoInterno.Enabled = True


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub

End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcImprimir
    'Case 3: ProcAjuda
    Case 4: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

