VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmFaturamento_Prod_serv_filtrar_carteira 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Administrativo - Faturamento - Nota fiscal - Filtrar carteira de pedidos"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8490
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmFaturamento_Prod_serv_filtrar_carteira.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   8490
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
      ItemData        =   "frmFaturamento_Prod_serv_filtrar_carteira.frx":1042
      Left            =   1170
      List            =   "frmFaturamento_Prod_serv_filtrar_carteira.frx":1044
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Empresa."
      Top             =   1110
      Width           =   7125
   End
   Begin VB.CheckBox optperiodo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Prazo final"
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
      Left            =   1500
      TabIndex        =   9
      Top             =   3270
      Width           =   1185
   End
   Begin VB.CheckBox Chk_data_venda 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Dt. venda"
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
      Left            =   255
      TabIndex        =   8
      Top             =   3270
      Width           =   1155
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Height          =   1515
      Left            =   55
      TabIndex        =   14
      Top             =   1470
      Width           =   8415
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Height          =   510
         Left            =   3450
         TabIndex        =   19
         Top             =   210
         Width           =   4785
         Begin VB.OptionButton Optfim 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Fim frase"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2760
            TabIndex        =   6
            Top             =   180
            Width           =   1155
         End
         Begin VB.OptionButton Optinicio 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Início frase"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   180
            TabIndex        =   4
            Top             =   180
            Value           =   -1  'True
            Width           =   1275
         End
         Begin VB.OptionButton Optmeio 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Meio frase"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1470
            TabIndex        =   5
            Top             =   180
            Width           =   1275
         End
         Begin VB.OptionButton optIgual 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Igual"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3930
            TabIndex        =   7
            Top             =   180
            Width           =   705
         End
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
         ItemData        =   "frmFaturamento_Prod_serv_filtrar_carteira.frx":1046
         Left            =   180
         List            =   "frmFaturamento_Prod_serv_filtrar_carteira.frx":1062
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "Opções para filtro."
         Top             =   390
         Width           =   3195
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
         TabIndex        =   2
         ToolTipText     =   "Texto para pesquisa."
         Top             =   1050
         Width           =   8055
      End
      Begin VB.ComboBox cmbTexto 
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
         ItemData        =   "frmFaturamento_Prod_serv_filtrar_carteira.frx":10DE
         Left            =   180
         List            =   "frmFaturamento_Prod_serv_filtrar_carteira.frx":10E0
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "Familia."
         Top             =   1050
         Width           =   8055
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filtrar por"
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
         Left            =   1357
         TabIndex        =   16
         Top             =   180
         Width           =   840
      End
      Begin VB.Label Label3 
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
         Left            =   3472
         TabIndex        =   15
         Top             =   840
         Width           =   1470
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   55
      TabIndex        =   17
      Top             =   0
      Width           =   8415
      _ExtentX        =   14843
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
      ButtonCaption1  =   "Filtrar"
      ButtonEnabled1  =   0   'False
      ButtonIconSize1 =   32
      ButtonToolTipText1=   "Filtrar (F2)"
      ButtonKey1      =   "1"
      ButtonAlignment1=   2
      BeginProperty ButtonFont1 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft1     =   2
      ButtonTop1      =   2
      ButtonWidth1    =   42
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
      ButtonLeft2     =   46
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft3     =   50
      ButtonTop3      =   2
      ButtonWidth3    =   41
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft4     =   93
      ButtonTop4      =   2
      ButtonWidth4    =   30
      ButtonHeight4   =   21
      ButtonUseMaskColor4=   0   'False
      ButtonEnabled5  =   0   'False
      ButtonIconSize5 =   32
      ButtonKey5      =   "5"
      ButtonAlignment5=   2
      BeginProperty ButtonFont5 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   4200
         Top             =   210
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmFaturamento_Prod_serv_filtrar_carteira.frx":10E2
         Count           =   1
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   675
      Left            =   55
      TabIndex        =   12
      Top             =   3000
      Width           =   8415
      Begin MSComCtl2.DTPicker msk_fltFim 
         Height          =   315
         Left            =   6840
         TabIndex        =   11
         ToolTipText     =   "Data final para pesquisa."
         Top             =   210
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   16777215
         CalendarForeColor=   0
         CalendarTitleBackColor=   8421504
         CalendarTitleForeColor=   16777215
         CalendarTrailingForeColor=   255
         Format          =   488243201
         CurrentDate     =   39057
      End
      Begin MSComCtl2.DTPicker msk_fltInicio 
         Height          =   315
         Left            =   5100
         TabIndex        =   10
         ToolTipText     =   "Data início para pesquisa."
         Top             =   210
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   16777215
         CalendarForeColor=   0
         CalendarTitleBackColor=   8421504
         CalendarTitleForeColor=   16777215
         CalendarTrailingForeColor=   255
         Format          =   488243203
         CurrentDate     =   39057
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "à"
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
         Height          =   255
         Left            =   6600
         TabIndex        =   13
         Top             =   240
         Width           =   255
      End
   End
   Begin VB.Label Label1 
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
      Left            =   270
      TabIndex        =   18
      Top             =   1110
      Width           =   825
   End
End
Attribute VB_Name = "frmFaturamento_Prod_serv_filtrar_carteira"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Chk_com_ordem_Click()
On Error GoTo tratar_erro

If cmbfiltrarpor = "Cliente" Then ProcCarregaComboTexto

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_data_venda_Click()
On Error GoTo tratar_erro

optPeriodo.Value = 0
If Chk_data_venda.Value = 1 Then
    Frame2.Enabled = True
    msk_fltInicio.SetFocus
Else
    Frame2.Enabled = False
    msk_fltFim.Value = Date
    msk_fltInicio.Value = Date
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_expedido_Click()
On Error GoTo tratar_erro

If cmbfiltrarpor = "Cliente" Then ProcCarregaComboTexto

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_faturar_faturados_Click()
On Error GoTo tratar_erro

If cmbfiltrarpor = "Cliente" Then ProcCarregaComboTexto

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_gerar_MRP_Click()
On Error GoTo tratar_erro

If cmbfiltrarpor = "Cliente" Then ProcCarregaComboTexto

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_MRP_gerado_Click()
On Error GoTo tratar_erro

If cmbfiltrarpor = "Cliente" Then ProcCarregaComboTexto

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_sem_ordem_Click()
On Error GoTo tratar_erro

If cmbfiltrarpor = "Cliente" Then ProcCarregaComboTexto

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ChkAtrasado_Click()
On Error GoTo tratar_erro

If cmbfiltrarpor = "Cliente" Then ProcCarregaComboTexto
If ChkAtrasado.Value = 1 Then
    optPeriodo.Value = 0
    optPeriodo.Enabled = False
    Chk_data_venda.Value = 0
    Chk_data_venda.Enabled = False
    msk_fltFim.Enabled = False
    msk_fltInicio.Enabled = False
    msk_fltFim.Value = Date
    msk_fltInicio.Value = Date
Else
    optPeriodo.Enabled = True
    Chk_data_venda.Enabled = True
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
    Exit Sub
End Sub

Private Sub cmbfiltrarpor_Click()
On Error GoTo tratar_erro

If cmbfiltrarpor = "Cliente" Then
    txtTexto.Visible = False
    cmbTexto.Visible = True
    ProcCarregaComboTexto
Else
    txtTexto.Visible = True
    cmbTexto.Visible = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaComboTexto()
On Error GoTo tratar_erro

ProcVerifFiltros
With cmbTexto
    .Clear
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select VP.cliente from vendas_carteira VC INNER JOIN vendas_proposta VP ON VC.cotacao = VP.cotacao where " & StatusFiltro & " and " & FiltroMRP & " and Qtdeexpedida " & Expedido & " VC.Quantidade and VP.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and " & TemOrdem & " and " & AtrasadoFiltro & " group by VP.cliente", Conexao, adOpenKeyset, adLockReadOnly
    If TBAbrir.EOF = False Then
        Do While TBAbrir.EOF = False
            .AddItem TBAbrir!Cliente
            TBAbrir.MoveNext
        Loop
    End If
    TBAbrir.Close
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 8415, 5, True

ProcCarregaComboEmpresa Cmb_empresa, False
cmbfiltrarpor = "Código interno"
msk_fltInicio.Value = Date
msk_fltFim.Value = Date

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro
  
Select Case KeyCode
    Case vbKeyF2: ProcFiltrar
    'Case vbKeyF1: ProcAjuda
    Case vbKeyEscape: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
   
Private Sub ProcFiltrar()
On Error GoTo tratar_erro

If Chk_gerar_MRP.Value = 0 And Chk_MRP_gerado.Value = 0 Then
    Acao = "filtrar"
    NomeCampo = "uma das opções do MRP"
    ProcVerificaAcao
    Exit Sub
End If
With msk_fltFim
    If FunVerificaDataFinal(msk_fltInicio.Value, .Value) = False Then
        .Value = Date
        .SetFocus
        Exit Sub
    End If
End With

'Deleta registros
ProcExcluirDadosProducaoRelatoriosTotal

ProcVerifFiltros
If Chk_data_venda.Value = 1 Then DataTexto = "Datavendas" Else DataTexto = "prazofinal"
If Chk_data_venda.Value = 1 Or optPeriodo.Value = 1 Then
    DataFiltro = "VC." & DataTexto & " Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "'"
    DataFiltroRel = "{Vendas_carteira." & DataTexto & "} >= Date(" & Year(msk_fltInicio.Value) & "," & Month(msk_fltInicio.Value) & "," & Day(msk_fltInicio.Value) & ") and {Vendas_carteira." & DataTexto & "} <= Date(" & _
                                Year(msk_fltFim.Value) & "," & Month(msk_fltFim.Value) & "," & Day(msk_fltFim.Value) & ")"
Else
    DataFiltro = "VC.desenho is not null"
    DataFiltroRel = "{vendas_carteira.desenho} <> 'Null'"
End If
With frmprod
    If txtTexto <> "" Or cmbTexto <> "" Then
        If cmbfiltrarpor = "Grupo do cliente" Then
            If Optinicio.Value = True Then
                .StrSql_Ordem_MRP = "Select VC.*, VP.ID_empresa, VP.Ncotacao, VP.revisao, VP.cliente FROM ((Clientes_grupos CG INNER JOIN Clientes C ON CG.ID = C.IDgrupo) INNER JOIN vendas_proposta VP ON VP.IDCliente = C.IDCliente) INNER JOIN vendas_carteira VC ON VP.cotacao = VC.cotacao where CP.Texto like '" & txtTexto & "%' and " & StatusFiltro & " and " & FiltroMRP & " and VC.Qtdeexpedida " & Expedido & " VC.Quantidade and VP.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and " & TemOrdem & " and " & DataFiltro & " and " & AtrasadoFiltro & " order by " & DataTexto
                .FormulaRel_Ordem_Carteira = "{Clientes_grupos.Texto} like '" & txtTexto & "*' and " & StatusFiltroRel & " and " & FiltroMRPRel & " and {vendas_carteira.Qtdeexpedida} " & Expedido & " {vendas_carteira.Quantidade} and {vendas_proposta.ID_empresa} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and " & TemOrdemRel & " and " & DataFiltroRel & " and " & AtrasadoFiltroRel
            End If
            If Optmeio.Value = True Then
                .StrSql_Ordem_MRP = "Select VC.*, VP.ID_empresa, VP.Ncotacao, VP.revisao, VP.cliente FROM ((Clientes_grupos CG INNER JOIN Clientes C ON CG.ID = C.IDgrupo) INNER JOIN vendas_proposta VP ON VP.IDCliente = C.IDCliente) INNER JOIN vendas_carteira VC ON VP.cotacao = VC.cotacao where CP.Texto like '%" & txtTexto & "%' and " & StatusFiltro & " and " & FiltroMRP & " and VC.Qtdeexpedida " & Expedido & " VC.Quantidade and VP.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and " & TemOrdem & " and " & DataFiltro & " and " & AtrasadoFiltro & " order by " & DataTexto
                .FormulaRel_Ordem_Carteira = "{Clientes_grupos.Texto} like '*" & txtTexto & "*' and " & StatusFiltroRel & " and " & FiltroMRPRel & " and {vendas_carteira.Qtdeexpedida} " & Expedido & " {vendas_carteira.Quantidade} and {vendas_proposta.ID_empresa} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and " & TemOrdemRel & " and " & DataFiltroRel & " and " & AtrasadoFiltroRel
            End If
            If Optfim.Value = True Then
                .StrSql_Ordem_MRP = "Select VC.*, VP.ID_empresa, VP.Ncotacao, VP.revisao, VP.cliente FROM ((Clientes_grupos CG INNER JOIN Clientes C ON CG.ID = C.IDgrupo) INNER JOIN vendas_proposta VP ON VP.IDCliente = C.IDCliente) INNER JOIN vendas_carteira VC ON VP.cotacao = VC.cotacao where CP.Texto like '%" & txtTexto & "' and " & StatusFiltro & " and " & FiltroMRP & " and VC.Qtdeexpedida " & Expedido & " VC.Quantidade and VP.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and " & TemOrdem & " and " & DataFiltro & " and " & AtrasadoFiltro & " order by " & DataTexto
                .FormulaRel_Ordem_Carteira = "{Clientes_grupos.Texto} like '*" & txtTexto & "' and " & StatusFiltroRel & " and " & FiltroMRPRel & " and {vendas_carteira.Qtdeexpedida} " & Expedido & " {vendas_carteira.Quantidade} and {vendas_proposta.ID_empresa} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and " & TemOrdemRel & " and " & DataFiltroRel & " and " & AtrasadoFiltroRel
            End If
        ElseIf cmbfiltrarpor = "Cliente" Then
                .StrSql_Ordem_MRP = "Select VC.*, VP.ID_empresa, VP.Ncotacao, VP.revisao, VP.cliente from vendas_carteira VC INNER JOIN vendas_proposta VP ON VC.cotacao = VP.cotacao where VP.Cliente = '" & cmbTexto & "' and " & StatusFiltro & " and " & FiltroMRP & " and Qtdeexpedida " & Expedido & " VC.Quantidade and VP.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and " & TemOrdem & " and " & DataFiltro & " and " & AtrasadoFiltro & " order by " & DataTexto
                .FormulaRel_Ordem_Carteira = "{Vendas_proposta.Cliente} = '" & cmbTexto & "' and " & StatusFiltroRel & " and " & FiltroMRPRel & " and {vendas_carteira.Qtdeexpedida} " & Expedido & " {vendas_carteira.Quantidade} and {vendas_proposta.ID_empresa} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and " & TemOrdemRel & " and " & DataFiltroRel & " and " & AtrasadoFiltroRel
            Else
                Select Case cmbfiltrarpor
                    Case "Código de referência":
                        TextoFiltro = "VC.n_referencia"
                        TextoFiltroRel = "Vendas_carteira.n_referencia"
                    Case "Código interno":
                        TextoFiltro = "VC.Desenho"
                        TextoFiltroRel = "Vendas_carteira.Desenho"
                    Case "Descrição":
                        TextoFiltro = "VC.Descricao_tecnica"
                        TextoFiltroRel = "Vendas_carteira.Descricao_tecnica"
                    Case "Família":
                        TextoFiltro = "VC.Familia"
                        TextoFiltroRel = "Vendas_carteira.Familia"
                    Case "Pedido do cliente":
                        TextoFiltro = "VC.PCcliente"
                        TextoFiltroRel = "Vendas_carteira.PCcliente"
                    Case "Pedido interno":
                        TextoFiltro = "VP.Ncotacao"
                        TextoFiltroRel = "Vendas_proposta.Ncotacao"
                End Select
                If Optinicio.Value = True Then
                    .StrSql_Ordem_MRP = "Select VC.*, VP.ID_empresa, VP.Ncotacao, VP.revisao, VP.cliente from vendas_carteira VC INNER JOIN vendas_proposta VP ON VC.cotacao = VP.cotacao where " & TextoFiltro & " like '" & txtTexto & "%' and " & StatusFiltro & " and " & FiltroMRP & " and VC.Qtdeexpedida " & Expedido & " VC.Quantidade and VP.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and " & TemOrdem & " and " & DataFiltro & " and " & AtrasadoFiltro & " order by " & DataTexto
                    .FormulaRel_Ordem_Carteira = "{" & TextoFiltroRel & "} like '" & txtTexto & "*' and " & StatusFiltroRel & " and " & FiltroMRPRel & " and {vendas_carteira.Qtdeexpedida} " & Expedido & " {vendas_carteira.Quantidade} and {vendas_proposta.ID_empresa} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and " & TemOrdemRel & " and " & DataFiltroRel & " and " & AtrasadoFiltroRel
                End If
                If Optmeio.Value = True Then
                    .StrSql_Ordem_MRP = "Select VC.*, VP.ID_empresa, VP.Ncotacao, VP.revisao, VP.cliente from vendas_carteira VC INNER JOIN vendas_proposta VP ON VC.cotacao = VP.cotacao where " & TextoFiltro & " like '%" & txtTexto & "%' and " & StatusFiltro & " and " & FiltroMRP & " and VC.Qtdeexpedida " & Expedido & " VC.Quantidade and VP.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and " & TemOrdem & " and " & DataFiltro & " and " & AtrasadoFiltro & " order by " & DataTexto
                    .FormulaRel_Ordem_Carteira = "{" & TextoFiltroRel & "} like '*" & txtTexto & "*' and " & StatusFiltroRel & " and " & FiltroMRPRel & " and {vendas_carteira.Qtdeexpedida} " & Expedido & " {vendas_carteira.Quantidade} and {vendas_proposta.ID_empresa} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and " & TemOrdemRel & " and " & DataFiltroRel & " and " & AtrasadoFiltroRel
                End If
                If Optfim.Value = True Then
                    .StrSql_Ordem_MRP = "Select VC.*, VP.ID_empresa, VP.Ncotacao, VP.revisao, VP.cliente from vendas_carteira VC INNER JOIN vendas_proposta VP ON VC.cotacao = VP.cotacao where " & TextoFiltro & " like '%" & txtTexto & "' and " & StatusFiltro & " and " & FiltroMRP & " and VC.Qtdeexpedida " & Expedido & " VC.Quantidade and VP.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and " & TemOrdem & " and " & DataFiltro & " and " & AtrasadoFiltro & " order by " & DataTexto
                    .FormulaRel_Ordem_Carteira = "{" & TextoFiltroRel & "} like '*" & txtTexto & "' and " & StatusFiltroRel & " and " & FiltroMRPRel & " and {vendas_carteira.Qtdeexpedida} " & Expedido & " {vendas_carteira.Quantidade} and {vendas_proposta.ID_empresa} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and " & TemOrdemRel & " and " & DataFiltroRel & " and " & AtrasadoFiltroRel
                End If
                If optIgual.Value = True Then
                    .StrSql_Ordem_MRP = "Select VC.*, VP.ID_empresa, VP.Ncotacao, VP.revisao, VP.cliente from vendas_carteira VC INNER JOIN vendas_proposta VP ON VC.cotacao = VP.cotacao where " & TextoFiltro & " = '" & txtTexto & "' and " & StatusFiltro & " and " & FiltroMRP & " and VC.Qtdeexpedida " & Expedido & " VC.Quantidade and VP.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and " & TemOrdem & " and " & DataFiltro & " and " & AtrasadoFiltro & " order by " & DataTexto
                    .FormulaRel_Ordem_Carteira = "{" & TextoFiltroRel & "} = '" & txtTexto & "' and " & StatusFiltroRel & " and " & FiltroMRPRel & " and {vendas_carteira.Qtdeexpedida} " & Expedido & " {vendas_carteira.Quantidade} and {vendas_proposta.ID_empresa} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and " & TemOrdemRel & " and " & DataFiltroRel & " and " & AtrasadoFiltroRel
                End If
        End If
    Else
        .StrSql_Ordem_MRP = "Select VC.*, VP.ID_empresa, VP.Ncotacao, VP.revisao, VP.cliente from vendas_carteira VC INNER JOIN vendas_proposta VP ON VC.cotacao = VP.cotacao where " & StatusFiltro & " and " & FiltroMRP & " and VC.Qtdeexpedida " & Expedido & " VC.Quantidade and VP.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and " & TemOrdem & " and " & DataFiltro & " and " & AtrasadoFiltro & " order by " & DataTexto
        .FormulaRel_Ordem_Carteira = StatusFiltroRel & " and " & FiltroMRPRel & " and {vendas_carteira.Qtdeexpedida} " & Expedido & " {vendas_carteira.Quantidade} and {vendas_proposta.ID_empresa} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and " & TemOrdemRel & " and " & DataFiltroRel & " and " & AtrasadoFiltroRel
    End If
    Call .m_Tree.Nodes.Clear
    .Grid1.rows = 1
    .ProcAtualizalista_carteira
    .Grid1.Visible = False
    .listaitens.Visible = True
   ' .PBLista.Visible = True
    .Frame1(2).Visible = True
    .ProcEsconderMostrarBotoes
End With

Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from Producao_Relatorios_Total", Conexao, adOpenKeyset, adLockOptimistic
TBGravar.AddNew
TBGravar!Responsavel = pubUsuario
TBGravar!Modulo = Formulario
TBGravar!Data_inicial = msk_fltInicio
TBGravar!Data_final = msk_fltInicio
TBGravar.Update

Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcVerifFiltros()
On Error GoTo tratar_erro

If Chk_gerar_MRP.Value = 1 And Chk_MRP_gerado.Value = 1 Then
    FiltroMRP = "(VC.OE = 'False' or VC.OE = 'True')"
    FiltroMRPRel = "({Vendas_carteira.OE} = False or {Vendas_carteira.OE} = True)"
ElseIf Chk_gerar_MRP.Value = 1 And Chk_MRP_gerado.Value = 0 Then
        FiltroMRP = "VC.OE = 'False'"
        FiltroMRPRel = "{Vendas_carteira.OE} = False"
    Else
        FiltroMRP = "VC.OE = 'True'"
        FiltroMRPRel = "{Vendas_carteira.OE} = True"
End If
If Chk_com_ordem.Value = 1 And Chk_sem_ordem.Value = 1 Then
    TemOrdem = "VC.Desenho is not null"
    TemOrdemRel = "{Vendas_carteira.Desenho} <> 'Null'"
ElseIf Chk_com_ordem.Value = 1 Then
        TemOrdem = "VC.Tem_ordem = 'True'"
        TemOrdemRel = "{Vendas_carteira.Tem_ordem} = True"
    Else
        TemOrdem = "VC.Tem_ordem = 'False'"
        TemOrdemRel = "{Vendas_carteira.Tem_ordem} = False"
End If
If Chk_expedido.Value = 1 Then Expedido = ">=" Else Expedido = "<"
If ChkAtrasado.Value = 1 Then
    AtrasadoFiltro = "VC.prazofinal < '" & Format(Date, "Short Date") & "'"
    AtrasadoFiltroRel = "{vendas_carteira.prazofinal} < Date(" & Year(Date) & "," & Month(Date) & "," & Day(Date) & ")"
Else
    AtrasadoFiltro = "VC.desenho is not null"
    AtrasadoFiltroRel = "{vendas_carteira.desenho} <> 'Null'"
End If
StatusFiltro = "VC.liberacao <> 'ABERTA EM ANALISE' and VC.liberacao <> 'REVISADA' and VC.liberacao <> 'CANCELADO' and VC.liberacao <> 'PERDIDO P/ PRAZO' and VC.liberacao <> 'PERDIDO P/ PREÇO' and VC.liberacao <> 'PORTAL ELETRONICO'"
StatusFiltroRel = "{vendas_carteira.liberacao} <> 'ABERTA EM ANALISE' and {vendas_carteira.liberacao} <> 'REVISADA' and {vendas_carteira.liberacao} <> 'CANCELADO' and {vendas_carteira.liberacao} <> 'PERDIDO P/ PRAZO' and {vendas_carteira.liberacao} <> 'PERDIDO P/ PREÇO' and {vendas_carteira.liberacao} <> 'PORTAL ELETRONICO'"
If Chk_faturar_faturados.Value = 0 Then
    StatusFiltro = StatusFiltro & " and VC.liberacao <> 'FATURAR' and VC.liberacao <> 'FATURADO'"
    StatusFiltroRel = StatusFiltroRel & " and {vendas_carteira.liberacao} <> 'FATURAR' and {vendas_carteira.liberacao} <> 'FATURADO'"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optPeriodo_Click()
On Error GoTo tratar_erro

Chk_data_venda.Value = 0
If optPeriodo.Value = 1 Then
    Frame2.Enabled = True
    msk_fltInicio.SetFocus
Else
    Frame2.Enabled = False
    msk_fltFim.Value = Date
    msk_fltInicio.Value = Date
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcFiltrar
    'Case 3: ProcAjuda
    Case 4: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
