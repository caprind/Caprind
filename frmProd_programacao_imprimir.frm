VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmProd_programacao_imprimir 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'Nenhum
   Caption         =   "PCP | Progr. da produção | impressão"
   ClientHeight    =   2625
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   4635
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
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
   ScaleHeight     =   2625
   ScaleWidth      =   4635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Centralziar na Tela
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   9
      Top             =   2220
      Width           =   4635
      _ExtentX        =   8176
      _ExtentY        =   714
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   405
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   4635
      _ExtentX        =   8176
      _ExtentY        =   714
      DibPicture      =   "frmProd_programacao_imprimir.frx":0000
      CaptionDelimiter=   "|"
      CaptionOnCenter =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Icon            =   "frmProd_programacao_imprimir.frx":1C95
      ShowMaximize    =   0   'False
      ShowMinimize    =   0   'False
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Período"
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
      Height          =   855
      Left            =   90
      TabIndex        =   0
      Top             =   1335
      Width           =   4485
      Begin MSComCtl2.DTPicker txtDataFinal 
         Height          =   315
         Left            =   2880
         TabIndex        =   2
         ToolTipText     =   "Data final."
         Top             =   330
         Width           =   1215
         _ExtentX        =   2143
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
         Format          =   135069699
         CurrentDate     =   39057
      End
      Begin MSComCtl2.DTPicker txtDataInicio 
         Height          =   315
         Left            =   960
         TabIndex        =   3
         ToolTipText     =   "Data de início."
         Top             =   330
         Width           =   1410
         _ExtentX        =   2487
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
         Format          =   135069699
         CurrentDate     =   39057
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Alinhar à Direita
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparente
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
         Index           =   1
         Left            =   600
         TabIndex        =   4
         Top             =   330
         Width           =   300
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Alinhar à Direita
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparente
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
         Left            =   2460
         TabIndex        =   1
         Top             =   330
         Width           =   360
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   885
      Left            =   0
      TabIndex        =   7
      Top             =   420
      Width           =   5130
      _ExtentX        =   9049
      _ExtentY        =   1561
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
      ButtonHeight2   =   48
      ButtonCaption3  =   "Ajuda"
      ButtonEnabled3  =   0   'False
      ButtonIconSize3 =   32
      ButtonToolTipText3=   "Ajuda (F1)"
      ButtonKey3      =   "14"
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
      ButtonKey4      =   "15"
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
      ButtonKey5      =   "16"
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
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   3000
         Top             =   150
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmProd_programacao_imprimir.frx":1FAF
         Count           =   1
      End
   End
   Begin VB.OptionButton optData 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Data"
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
      Height          =   210
      Left            =   1335
      TabIndex        =   5
      Top             =   990
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparente
      Caption         =   "Filtrar por:"
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
      TabIndex        =   6
      Top             =   990
      Visible         =   0   'False
      Width           =   885
   End
End
Attribute VB_Name = "frmProd_programacao_imprimir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ProcImprimir()
On Error GoTo tratar_erro

With frmProd_programacao
    NomeRel = "PCP_programacao.rpt"
    If .FiltroImpressao = True Then
        If optData.Value = True Then
            If .FormulaRel_Ordem_Programacao = "" Then
                ProcImprimirRel .FormulaRel_Ordem_Programacao & " {pcp_programacao_ordem.inicio}>=Date(" & Year(txtDataInicio.Value) & "," & Month(txtDataInicio.Value) & "," & Day(txtDataInicio.Value) & ") and {pcp_programacao_ordem.inicio}<=Date(" & Year(txtDataFinal.Value) & "," & Month(txtDataFinal.Value) & "," & Day(txtDataFinal.Value) & ") or {pcp_programacao_ordem.final}>=Date(" & Year(txtDataInicio.Value) & "," & Month(txtDataInicio.Value) & "," & Day(txtDataInicio.Value) & ") and {pcp_programacao_ordem.final}<=Date(" & Year(txtDataFinal.Value) & "," & Month(txtDataFinal.Value) & "," & Day(txtDataFinal.Value) & ")", ""
            Else
                ProcImprimirRel .FormulaRel_Ordem_Programacao & " and {pcp_programacao_ordem.inicio}>=Date(" & Year(txtDataInicio.Value) & "," & Month(txtDataInicio.Value) & "," & Day(txtDataInicio.Value) & ") and {pcp_programacao_ordem.inicio}<=Date(" & Year(txtDataFinal.Value) & "," & Month(txtDataFinal.Value) & "," & Day(txtDataFinal.Value) & ") or {pcp_programacao_ordem.final}>=Date(" & Year(txtDataInicio.Value) & "," & Month(txtDataInicio.Value) & "," & Day(txtDataInicio.Value) & ") and {pcp_programacao_ordem.final}<=Date(" & Year(txtDataFinal.Value) & "," & Month(txtDataFinal.Value) & "," & Day(txtDataFinal.Value) & ")", ""
            End If
        End If
    Else
        If optData.Value = True Then
            ProcImprimirRel "{pcp_programacao.id}= " & .txtId & " and {pcp_programacao_ordem.inicio}>=Date(" & Year(txtDataInicio.Value) & "," & Month(txtDataInicio.Value) & "," & Day(txtDataInicio.Value) & ") and {pcp_programacao_ordem.inicio}<=Date(" & Year(txtDataFinal.Value) & "," & Month(txtDataFinal.Value) & "," & Day(txtDataFinal.Value) & ") or {pcp_programacao.id}= " & .txtId & " and {pcp_programacao_ordem.final}>=Date(" & Year(txtDataInicio.Value) & "," & Month(txtDataInicio.Value) & "," & Day(txtDataInicio.Value) & ") and {pcp_programacao_ordem.final}<=Date(" & Year(txtDataFinal.Value) & "," & Month(txtDataFinal.Value) & "," & Day(txtDataFinal.Value) & ")", ""
        Else
            ProcImprimirRel "{pcp_programacao.id}= " & .txtId, ""
        End If
    End If
End With

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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyF5: ProcImprimir
    Case vbKeyEscape: ProcSair
End Select
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

txtDataInicio = Date
txtDataFinal = Date

ProcCarregaToolBar1 Me, 15195, 4, True


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optData_Click()
On Error GoTo tratar_erro

Frame1.Enabled = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optTudo_Click()
On Error GoTo tratar_erro

txtDataInicio = Date
txtDataFinal = Date
Frame1.Enabled = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcImprimir
    Case 4: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub

End Sub

