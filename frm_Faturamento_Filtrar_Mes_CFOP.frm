VERSION 5.00
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Begin VB.Form frm_Faturamento_Filtrar_Mes_CFOP 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   ClientHeight    =   3135
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3210
   Icon            =   "frm_Faturamento_Filtrar_Mes_CFOP.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   3210
   StartUpPosition =   1  'CenterOwner
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   7
      Top             =   2730
      Width           =   3210
      _ExtentX        =   5662
      _ExtentY        =   714
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Opções de filtro"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2025
      Left            =   180
      TabIndex        =   1
      Top             =   540
      Width           =   2775
      Begin DrawSuite2022.USButton btn_Filtrar 
         Height          =   915
         Left            =   330
         TabIndex        =   6
         ToolTipText     =   "Visualizar relatório"
         Top             =   960
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   1614
         DibPicture      =   "frm_Faturamento_Filtrar_Mes_CFOP.frx":000C
         Caption         =   "Visualizar Relatório"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         PicAlign        =   7
         PicSize         =   5
         PicSizeH        =   32
         PicSizeW        =   32
         ShowFocusRect   =   0   'False
         Theme           =   4
      End
      Begin VB.ComboBox cmbMes 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frm_Faturamento_Filtrar_Mes_CFOP.frx":108D0
         Left            =   1440
         List            =   "frm_Faturamento_Filtrar_Mes_CFOP.frx":108F8
         TabIndex        =   3
         Top             =   390
         Width           =   1065
      End
      Begin VB.ComboBox cmbAno 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frm_Faturamento_Filtrar_Mes_CFOP.frx":10961
         Left            =   300
         List            =   "frm_Faturamento_Filtrar_Mes_CFOP.frx":10983
         TabIndex        =   2
         Top             =   390
         Width           =   1065
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Mês"
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
         Left            =   1710
         TabIndex        =   5
         Top             =   210
         Width           =   405
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Ano"
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
         Left            =   630
         TabIndex        =   4
         Top             =   210
         Width           =   405
      End
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3210
      _ExtentX        =   5662
      _ExtentY        =   714
      EnableMaximizeButton=   0   'False
      EnableMinimizeButton=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frm_Faturamento_Filtrar_Mes_CFOP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btn_Filtrar_Click()
On Error GoTo tratar_erro
    
Dim FormulaRel_CFOP As String
Dim Mes As Integer
Dim Ano As Integer

Select Case cmbMes.Text
    Case "Janeiro": Mes = 1
    Case "Fevereiro": Mes = 2
    Case "Março": Mes = 3
    Case "Abril": Mes = 4
    Case "Maio": Mes = 5
    Case "Junho": Mes = 6
    Case "Julho": Mes = 7
    Case "Agosto": Mes = 8
    Case "Setembro": Mes = 9
    Case "Outubro": Mes = 10
    Case "Novembro": Mes = 11
    Case "Dezembro": Mes = 12
End Select

Ano = cmbAno.Text

'FormulaRel_CFOP = "year({tbl_Dados_Nota_Fiscal.dt_DataEmissao})=" & Ano & " and month({tbl_Dados_Nota_Fiscal.dt_DataEmissao})= " & Mes & " and {tbl_Dados_Nota_Fiscal.int_status}=1 and {tbl_Dados_Nota_Fiscal.int_NotaFiscal} <> """""
FormulaRel_CFOP = "{Faturamento_Relatorio_Total_CFOP_Mes.Ano} =" & Ano & " AND {Faturamento_Relatorio_Total_CFOP_Mes.Mes} = " & Mes
NomeRel = "Notas emitidas mensal_CFOP.rpt"
'Debug.print FormulaRel_CFOP

ProcImprimirRel FormulaRel_CFOP, ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

Mes = Month(Date)

Select Case Mes
    Case 1: cmbMes.Text = "Janeiro"
    Case 2: cmbMes.Text = "Fevereiro"
    Case 3: cmbMes.Text = "Março"
    Case 4: cmbMes.Text = "Abril"
    Case 5: cmbMes.Text = "Maio"
    Case 6: cmbMes.Text = "Junho"
    Case 7: cmbMes.Text = "Julho"
    Case 8: cmbMes.Text = "Agosto"
    Case 9: cmbMes.Text = "Setembro"
    Case 10: cmbMes.Text = "Outubro"
    Case 11: cmbMes.Text = "Novembro"
    Case 12: cmbMes.Text = "Dezembro"
End Select


cmbAno.Text = Year(Date)
'cmbMes.Text = Mes

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub

End Sub
