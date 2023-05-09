VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frm_Instituicoes2_localizar_cheque 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Financeiro - Instituições - Localizar - Cheques"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4500
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frm_Instituicoes2_localizar_cheque.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   3060
      Top             =   150
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frm_Instituicoes2_localizar_cheque.frx":1042
      Count           =   1
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   55
      TabIndex        =   9
      Top             =   0
      Width           =   4410
      _ExtentX        =   7779
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft1     =   2
      ButtonTop1      =   2
      ButtonWidth1    =   36
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
      ButtonLeft2     =   40
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft3     =   44
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
      ButtonLeft4     =   82
      ButtonTop4      =   2
      ButtonWidth4    =   26
      ButtonHeight4   =   21
      ButtonUseMaskColor4=   0   'False
      ButtonEnabled5  =   0   'False
      ButtonIconSize5 =   32
      ButtonKey5      =   "5"
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
      ButtonLeft5     =   110
      ButtonTop5      =   2
      ButtonWidth5    =   24
      ButtonHeight5   =   24
      ButtonUseMaskColor5=   0   'False
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   60
      TabIndex        =   10
      Top             =   2265
      Width           =   4410
      _ExtentX        =   7779
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
      SearchText      =   "Atualizando..."
      Value           =   0
   End
   Begin VB.ComboBox Cmb_cheque 
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
      ItemData        =   "frm_Instituicoes2_localizar_cheque.frx":3233
      Left            =   990
      List            =   "frm_Instituicoes2_localizar_cheque.frx":323D
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   3
      ToolTipText     =   "Cheque."
      Top             =   1770
      Visible         =   0   'False
      Width           =   3315
   End
   Begin VB.Frame frm_filtro 
      BackColor       =   &H00E0E0E0&
      Height          =   1260
      Left            =   55
      TabIndex        =   4
      Top             =   990
      Width           =   4410
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
         ItemData        =   "frm_Instituicoes2_localizar_cheque.frx":3252
         Left            =   180
         List            =   "frm_Instituicoes2_localizar_cheque.frx":325C
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Opções para filtro."
         Top             =   390
         Width           =   4065
      End
      Begin MSComCtl2.DTPicker msk_fltFim 
         Height          =   315
         Left            =   2850
         TabIndex        =   2
         ToolTipText     =   "Data final para pesquisa."
         Top             =   780
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
         MouseIcon       =   "frm_Instituicoes2_localizar_cheque.frx":3271
         CalendarBackColor=   16777215
         CalendarForeColor=   0
         CalendarTitleBackColor=   8421504
         CalendarTitleForeColor=   16777215
         CalendarTrailingForeColor=   255
         Format          =   487849985
         CurrentDate     =   39057
      End
      Begin MSComCtl2.DTPicker msk_fltInicio 
         Height          =   315
         Left            =   1140
         TabIndex        =   1
         ToolTipText     =   "Data início para pesquisa."
         Top             =   780
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
         MouseIcon       =   "frm_Instituicoes2_localizar_cheque.frx":358B
         CalendarBackColor=   16777215
         CalendarForeColor=   0
         CalendarTitleBackColor=   8421504
         CalendarTitleForeColor=   16777215
         CalendarTrailingForeColor=   255
         Format          =   487849987
         CurrentDate     =   39057
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Cheque :"
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
         Left            =   180
         TabIndex        =   8
         Top             =   780
         Visible         =   0   'False
         Width           =   675
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
         Left            =   1792
         TabIndex        =   7
         Top             =   180
         Width           =   840
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
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
         Left            =   2580
         TabIndex        =   6
         Top             =   870
         Width           =   255
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Período de :"
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
         Left            =   180
         TabIndex        =   5
         Top             =   780
         Width           =   885
      End
   End
End
Attribute VB_Name = "frm_Instituicoes2_localizar_cheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbfiltrarpor_Click()
On Error GoTo tratar_erro

Cmb_cheque.Clear
Texto = ""
With frm_Instituicoes
    If cmbfiltrarpor = "Cheque" Then
        Label2.Visible = True
        Cmb_cheque.Visible = True
        Label1.Visible = False
        msk_fltInicio.Visible = False
        Label3.Visible = False
        msk_fltFim.Visible = False
        PBLista.Visible = True
        frm_Instituicoes2_localizar_cheque.Height = 2985
        Set TBLISTA = CreateObject("adodb.recordset")
        If .Cheques_Emitidos = True Then
            TBLISTA.Open "Select NDoctoBaixa, Status from tbl_contaspagar where Banco = '" & .txtdescricao & "' and ID_empresa = " & .Cmb_empresa.ItemData(.Cmb_empresa.ListIndex) & " and (logsit = 'S' or logsit is null) and (FormaBaixa = 'CHEQUE' or FormaBaixa = 'CHEQUE PRÉ-DATADO') order by NDoctoBaixa", Conexao, adOpenKeyset, adLockOptimistic
        Else
            TBLISTA.Open "Select NDoctoBaixa, Status from tbl_contas_receber where (logsit = 'S' or logsit is null) and Banco = '" & .txtdescricao & "' and ID_empresa = " & .Cmb_empresa.ItemData(.Cmb_empresa.ListIndex) & " and (FormaBaixa = 'CHEQUE' or FormaBaixa = 'CHEQUE PRÉ-DATADO') order by NDoctoBaixa", Conexao, adOpenKeyset, adLockOptimistic
        End If
        If TBLISTA.EOF = False Then
            TBLISTA.MoveLast
            PBLista.Min = 0
            PBLista.Max = TBLISTA.RecordCount
            PBLista.Value = 1
            contador = 0
            TBLISTA.MoveFirst
            Do While TBLISTA.EOF = False
                If IsNull(TBLISTA!NDoctoBaixa) = False And TBLISTA!NDoctoBaixa <> "" Then
                    If Texto <> TBLISTA!NDoctoBaixa Then
                        Cmb_cheque.AddItem Trim(TBLISTA!NDoctoBaixa)
                        If TBLISTA!status = "CANCELADO" Then
                            Cmb_cheque.ItemData(Cmb_cheque.NewIndex) = 2
                        Else
                            Cmb_cheque.ItemData(Cmb_cheque.NewIndex) = 1
                        End If
                    End If
                    Texto = TBLISTA!NDoctoBaixa
                End If
                TBLISTA.MoveNext
                contador = contador + 1
                PBLista.Value = contador
            Loop
        End If
        TBLISTA.Close
    Else
        Label2.Visible = False
        Cmb_cheque.Visible = False
        Label1.Visible = True
        msk_fltInicio.Visible = True
        Label3.Visible = False
        msk_fltFim.Visible = True
        PBLista.Visible = False
        frm_Instituicoes2_localizar_cheque.Height = 2700
    End If
End With
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 4590, 5, True

cmbfiltrarpor = "Período"
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

With msk_fltFim
    If FunVerificaDataFinal(msk_fltInicio.Value, .Value) = False Then
        .Value = Date
        .SetFocus
        Exit Sub
    End If
End With
With frm_Instituicoes
    If cmbfiltrarpor = "Cheque" Then
        If Cmb_cheque = "" Then
            USMsgBox ("Informe o número do cheque antes de localizar."), vbExclamation, "CAPRIND v5.0"
            Cmb_cheque.SetFocus
            Exit Sub
        End If
        If .Cheques_Emitidos = True Then
            .StrSql_Instituicoes_Localizar_Cheque = "Select * from tbl_contaspagar where NDoctoBaixa = '" & Cmb_cheque & "' and (logsit = 'S' or logsit is null) and Banco = '" & .txtdescricao & "' and ID_empresa = " & .Cmb_empresa.ItemData(.Cmb_empresa.ListIndex) & " and (FormaBaixa = 'CHEQUE' or FormaBaixa = 'CHEQUE PRÉ-DATADO') order by NDoctoBaixa"
        Else
            .StrSql_Instituicoes_Localizar_Cheque_Recebidos = "Select * from tbl_contas_receber where NDoctoBaixa = '" & Cmb_cheque & "' and (logsit = 'S' or logsit is null) and Banco = '" & .txtdescricao & "' and ID_empresa = " & .Cmb_empresa.ItemData(.Cmb_empresa.ListIndex) & " and (FormaBaixa = 'CHEQUE' or FormaBaixa = 'CHEQUE PRÉ-DATADO') order by NDoctoBaixa"
        End If
    Else
        If .Cheques_Emitidos = True Then
            .StrSql_Instituicoes_Localizar_Cheque = "Select * from tbl_ContasPagar where Banco = '" & .txtdescricao & "' and Status <> 'CHEQUE CANCELADO' and NDoctoBaixa <> 'Null' and (DataBaixa) Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "' and (FormaBaixa = 'CHEQUE' or FormaBaixa = 'CHEQUE PRÉ-DATADO') and ID_empresa = " & .Cmb_empresa.ItemData(.Cmb_empresa.ListIndex) & " order by FormaBaixa, DataBaixa, NDoctoBaixa, IdIntConta"
            .StrSql_Instituicoes_Localizar_Cheque_Cancelados = "Select tbl_ContasPagar.IdIntConta, tbl_ContasPagar.FormaBaixa, tbl_ContasPagar.DataBaixa, tbl_ContasPagar.DataBaixa, tbl_ContasPagar.NDoctoBaixa, tbl_ContasPagar.txt_Fornecedor, tbl_ContasPagar.ValorPago, tbl_ContasPagar.status, tbl_ContasPagar.Obs, Cheques_Cancelados.Data_cancelamento, Cheques_Cancelados.responsavel, Cheques_Cancelados.motivo from tbl_ContasPagar INNER JOIN Cheques_Cancelados on tbl_ContasPagar.IdIntConta = Cheques_Cancelados.ID_conta where tbl_ContasPagar.Banco = '" & .txtdescricao & "' and tbl_ContasPagar.logsit is null and tbl_ContasPagar.Status = 'CHEQUE CANCELADO' and (tbl_ContasPagar.DataBaixa) Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "' and ID_empresa = " & .Cmb_empresa.ItemData(.Cmb_empresa.ListIndex) & " and (tbl_ContasPagar.FormaBaixa = 'CHEQUE PRÉ-DATADO' or tbl_ContasPagar.FormaBaixa = 'CHEQUE')"
        Else
            .StrSql_Instituicoes_Localizar_Cheque_Recebidos = "Select * from tbl_contas_receber where Banco = '" & .txtdescricao & "' and NDoctoBaixa <> 'Null' and (Data_pagamento) Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "' and (FormaBaixa = 'CHEQUE' or FormaBaixa = 'CHEQUE PRÉ-DATADO') and ID_empresa = " & .Cmb_empresa.ItemData(.Cmb_empresa.ListIndex) & " order by FormaBaixa, Data_pagamento, NDoctoBaixa, IdIntConta"
        End If
    End If
    .ProcCarregaListaCheque
End With
Unload Me

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
