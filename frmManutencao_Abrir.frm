VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmManutencao_abrir 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Manutenção - Equipamentos - Localizar"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8925
   ControlBox      =   0   'False
   ForeColor       =   &H00000080&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   8925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox optPeriodo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Por período :"
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
      Left            =   4170
      TabIndex        =   4
      Top             =   2730
      Width           =   1245
   End
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   6420
      Top             =   240
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frmManutencao_Abrir.frx":0000
      Count           =   1
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   1515
      Left            =   55
      TabIndex        =   10
      Top             =   990
      Width           =   8805
      Begin VB.ComboBox cmbtipo 
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
         ItemData        =   "frmManutencao_Abrir.frx":21ED
         Left            =   180
         List            =   "frmManutencao_Abrir.frx":21FA
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "Texto para pesquisa."
         Top             =   1050
         Visible         =   0   'False
         Width           =   8415
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00E0E0E0&
         Height          =   510
         Left            =   3810
         TabIndex        =   16
         Top             =   210
         WhatsThisHelpID =   210
         Width           =   4785
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
            TabIndex        =   8
            Top             =   180
            Width           =   705
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
            TabIndex        =   6
            Top             =   180
            Width           =   1275
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
            TabIndex        =   2
            Top             =   180
            Value           =   -1  'True
            Width           =   1275
         End
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
            TabIndex        =   7
            Top             =   180
            Width           =   1155
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
         TabIndex        =   1
         ToolTipText     =   "Texto para pesquisa."
         Top             =   1050
         Width           =   8415
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
         ItemData        =   "frmManutencao_Abrir.frx":2220
         Left            =   180
         List            =   "frmManutencao_Abrir.frx":222A
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Opções para filtro."
         Top             =   390
         Width           =   3555
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
         Left            =   1537
         TabIndex        =   12
         Top             =   180
         Width           =   840
      End
      Begin VB.Label Label1 
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
         Left            =   3645
         TabIndex        =   11
         Top             =   840
         Width           =   1470
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   15
      Top             =   0
      Width           =   8805
      _ExtentX        =   15531
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
      ButtonKey1      =   "2"
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
      ButtonKey3      =   "12"
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
      ButtonKey4      =   "13"
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
      ButtonKey5      =   "14"
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
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   675
      Left            =   55
      TabIndex        =   13
      Top             =   2490
      Width           =   8805
      Begin VB.OptionButton optManutencao 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Manutenção"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   270
         Value           =   -1  'True
         Width           =   1635
      End
      Begin VB.OptionButton OptSolicitacao 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Solicitação"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   1950
         TabIndex        =   17
         Top             =   270
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker Msk_final 
         Height          =   315
         Left            =   7320
         TabIndex        =   9
         ToolTipText     =   "Data final."
         Top             =   210
         Width           =   1305
         _ExtentX        =   2302
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
         Format          =   133496833
         CurrentDate     =   39057
      End
      Begin MSComCtl2.DTPicker Msk_inicio 
         Height          =   315
         Left            =   5430
         TabIndex        =   5
         ToolTipText     =   "Data inicio."
         Top             =   210
         Width           =   1305
         _ExtentX        =   2302
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
         Format          =   133496833
         CurrentDate     =   39057
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
         Height          =   195
         Left            =   6855
         TabIndex        =   14
         Top             =   240
         Width           =   360
      End
   End
End
Attribute VB_Name = "frmManutencao_abrir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbTipo_Click()
On Error GoTo tratar_erro

If cmbTipo <> "" Then txtTexto = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbfiltrarpor_Click()
On Error GoTo tratar_erro

If cmbfiltrarpor = "Tipo" Then
    txtTexto.Visible = False
    cmbTipo.Visible = True
Else
    txtTexto.Visible = True
    cmbTipo.Visible = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

With Msk_final
    If FunVerificaDataFinal(Msk_inicio.Value, .Value) = False Then
        .Value = Date
        .SetFocus
        Exit Sub
    End If
End With

With frmManutencao
    .Lista.ListItems.Clear
    
    INNERJOINTEXTO = "manutencao M"
If OptSolicitacao.Value = True Then
    DataFiltro = "M.Tipo = 'S'"
Else
    DataFiltro = "M.Tipo <> 'S'"

End If
    DataFiltroRel = ""
'=============================================
' Se for solicitacao por período
'==============================================
    If optPeriodo.Value = 1 Then
        If OptSolicitacao.Value = True Then
            INNERJOINTEXTO = "manutencao"
            DataFiltro = "(M.Data_Solicitacao) Between '" & Format(Msk_inicio.Value, "Short Date") & "' And '" & Format(Msk_final.Value, "Short Date") & " 23:59:59.997'"
            DataFiltroRel = " and {manutencao.Data_Solicitacao} >= Date(" & Year(Msk_inicio.Value) & "," & Month(Msk_inicio.Value) & "," & Day(Msk_inicio.Value) & ") and {manutencao.Data_Solicitacao}<= Date(" & _
                                    Year(Msk_final.Value) & "," & Month(Msk_final.Value) & "," & Day(Msk_final.Value) & ")"
        Else
            INNERJOINTEXTO = "manutencao M INNER JOIN manutencao_data MD on M.codigo = MD.idManutencao"
            DataFiltro = "(MD.data) Between '" & Format(Msk_inicio.Value, "Short Date") & "' And '" & Format(Msk_final.Value, "Short Date") & " 23:59:59.997'"
            DataFiltroRel = " and {manutencao_data.data} >= Date(" & Year(Msk_inicio.Value) & "," & Month(Msk_inicio.Value) & "," & Day(Msk_inicio.Value) & ") and {manutencao_data.data}<= Date(" & _
                                    Year(Msk_final.Value) & "," & Month(Msk_final.Value) & "," & Day(Msk_final.Value) & ")"
        End If
    End If
'================================================
' Filtrar por código do equipamento
'================================================
    .FormulaRelSubReport_Manutencao = "{Manutencao_data.idManutencao} = {?Pm-Manutencao.Código} and {manutencao_data.Status} = 'Aberta' " & DataFiltroRel
    
    If txtTexto <> "" Or cmbTipo <> "" Then
        If cmbfiltrarpor = "Código do equipamento" Then
            TextoFiltro = "M.idmaquina" & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto)
            TextoFiltroRel = "{manutencao.idmaquina}" & FunVerifTipoFiltroIMFRel(Optinicio, Optmeio, Optfim, optIgual, txtTexto)
        Else
            Select Case cmbTipo
                'Case "Solicitação": TipoTexto = "S"
                Case "Preventiva": TipoTexto = "P"
                Case "Corretiva": TipoTexto = "C"
                Case "Predetiva": TipoTexto = "PR"
            End Select
            TextoFiltro = "M.tipo = '" & TipoTexto & "'"
            TextoFiltroRel = "{manutencao.tipo} = '" & TipoTexto & "'"
        End If
        .Sql_Manutencao_Localizar = "Select M.* from " & INNERJOINTEXTO & " where " & TextoFiltro & " and " & DataFiltro & " order by M.CodMan, M.data desc, M.tipo"
        .FormulaRel_Manutencao = "{manutencao_data.Status} = 'Aberta' and {manutencao.idmaquina}" & FunVerifTipoFiltroIMFRel(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & DataFiltroRel
    Else
        .Sql_Manutencao_Localizar = "Select M.* from " & INNERJOINTEXTO & " where " & DataFiltro & " order by M.CodMan, M.data desc, M.tipo"
        .FormulaRel_Manutencao = "{manutencao_data.Status} = 'Aberta' " & DataFiltroRel
    End If
    .ProcCarregaLista
End With
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyEscape: ProcSair
    Case vbKeyF2: ProcFiltrar
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 8805, 5, True
cmbfiltrarpor = "Código do equipamento"
Msk_final.Value = Date
Msk_inicio.Value = Date

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

Private Sub optPeriodo_Click()
On Error GoTo tratar_erro

If optPeriodo.Value = 1 Then
    Msk_inicio.SetFocus
Else
    Msk_inicio.Value = Date
    Msk_final.Value = Date
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtTexto_Change()
On Error GoTo tratar_erro

If txtTexto <> "" Then cmbTipo.ListIndex = -1

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
