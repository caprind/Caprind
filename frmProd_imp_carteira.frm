VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmProd_imp_carteira 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "PCP | Gerenciamento de ordem - Filtrar carteira de produção"
   ClientHeight    =   5325
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   9045
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
   Icon            =   "frmProd_imp_carteira.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   9045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   29
      Top             =   4920
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   714
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Opções para filtrar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   270
      TabIndex        =   24
      Top             =   750
      Width           =   8415
      Begin VB.CheckBox chkProcesso 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Com processo"
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
         TabIndex        =   28
         Top             =   270
         Width           =   1605
      End
      Begin VB.CheckBox chkMRP 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Gerar MRP"
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
         Left            =   7110
         TabIndex        =   27
         Top             =   270
         Value           =   1  'Checked
         Width           =   1065
      End
      Begin VB.CheckBox chkConjunto 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Com conjunto"
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
         Left            =   2100
         TabIndex        =   26
         Top             =   270
         Value           =   1  'Checked
         Width           =   1365
      End
      Begin VB.CheckBox chkPlanoinspecao 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Com plano de inspeção"
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
         Left            =   4050
         TabIndex        =   25
         Top             =   270
         Width           =   1995
      End
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   873
      DibPicture      =   "frmProd_imp_carteira.frx":000C
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
      ShowMaximize    =   0   'False
      ShowMinimize    =   0   'False
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   270
      TabIndex        =   18
      Top             =   1395
      Width           =   8415
      Begin VB.ComboBox Cmb_filtrar 
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
         ItemData        =   "frmProd_imp_carteira.frx":365C
         Left            =   4950
         List            =   "frmProd_imp_carteira.frx":3669
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "Filtrar."
         Top             =   380
         Width           =   3285
      End
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
         ItemData        =   "frmProd_imp_carteira.frx":3696
         Left            =   180
         List            =   "frmProd_imp_carteira.frx":3698
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Empresa."
         Top             =   380
         Width           =   4665
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filtrar"
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
         Left            =   6390
         TabIndex        =   20
         Top             =   180
         Width           =   420
      End
      Begin VB.Label Label1 
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2010
         TabIndex        =   19
         Top             =   180
         Width           =   615
      End
   End
   Begin VB.CheckBox optperiodo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Prazo final?"
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
      Left            =   1920
      TabIndex        =   10
      Top             =   4110
      Width           =   1185
   End
   Begin VB.CheckBox Chk_data_venda 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Data venda?"
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
      Left            =   555
      TabIndex        =   9
      Top             =   4110
      Width           =   1245
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1515
      Left            =   270
      TabIndex        =   15
      Top             =   2220
      Width           =   8415
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   3450
         TabIndex        =   21
         Top             =   210
         Width           =   4785
         Begin VB.OptionButton Optfim 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Fim frase"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
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
         Begin VB.OptionButton Optinicio 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Início frase"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   180
            TabIndex        =   5
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
               Weight          =   400
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
         Begin VB.OptionButton optIgual 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Igual"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
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
         ItemData        =   "frmProd_imp_carteira.frx":369A
         Left            =   180
         List            =   "frmProd_imp_carteira.frx":36B9
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "Opções para filtro."
         Top             =   390
         Width           =   3165
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
         ItemData        =   "frmProd_imp_carteira.frx":374E
         Left            =   180
         List            =   "frmProd_imp_carteira.frx":3750
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1335
         TabIndex        =   17
         Top             =   180
         Width           =   705
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
         TabIndex        =   16
         Top             =   840
         Width           =   1470
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Filtrar por?"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   270
      TabIndex        =   13
      Top             =   3750
      Width           =   5895
      Begin MSComCtl2.DTPicker msk_fltFim 
         Height          =   315
         Left            =   4350
         TabIndex        =   12
         ToolTipText     =   "Data final para pesquisa."
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
         Format          =   489422849
         CurrentDate     =   39057
      End
      Begin MSComCtl2.DTPicker msk_fltInicio 
         Height          =   315
         Left            =   2850
         TabIndex        =   11
         ToolTipText     =   "Data início para pesquisa."
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
         Format          =   489422851
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
         Left            =   4170
         TabIndex        =   14
         Top             =   390
         Width           =   135
      End
   End
   Begin DrawSuite2022.USButton cmdFiltrar 
      Height          =   825
      Left            =   6180
      TabIndex        =   23
      Top             =   3750
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   1455
      DibPicture      =   "frmProd_imp_carteira.frx":3752
      BorderColor     =   5263559
      BorderColorDisabled=   13160660
      BorderColorDown =   4013465
      BorderColorOver =   4408288
      Caption         =   "Filtrar carteira de produção"
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
      PicAlign        =   3
      PicSize         =   2
      PicSizeH        =   24
      PicSizeW        =   24
      Theme           =   4
   End
End
Attribute VB_Name = "frmProd_imp_carteira"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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

If cmbfiltrarpor = "Cliente" Or cmbfiltrarpor = "Família" Then
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
    If cmbfiltrarpor = "Cliente" Then CampoFiltro = "Cliente" Else CampoFiltro = "Familia"
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select " & CampoFiltro & " as NomeCampo from Carteira_producao where ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and " & FiltroMRP & " and " & CampoFiltro & " IS NOT NULL group by " & CampoFiltro, Conexao, adOpenKeyset, adLockReadOnly
    If TBAbrir.EOF = False Then
        Do While TBAbrir.EOF = False
            .AddItem TBAbrir!NomeCampo
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

Private Sub cmdFiltrar_Click()
On Error GoTo tratar_erro

  ProcFiltrar
  
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaComboEmpresa Cmb_empresa, False
Cmb_filtrar = "Com necessidade"

Permitido = True
ProcFiltroPadrao cmbfiltrarpor, Optmeio, Optfim, optIgual, Cmb_empresa.ItemData(Cmb_empresa.ListIndex), "Produtos/Serviços", "P", True
If Permitido = False Then cmbfiltrarpor = "Código interno"

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

'Deleta registros
ProcExcluirDadosProducaoRelatoriosTotal

ProcVerifFiltros
DataFiltro = ""
DataFiltroRel = ""
If Chk_data_venda.Value = 1 Then DataTexto = "Datavendas" Else DataTexto = "prazofinal"
If Chk_data_venda.Value = 1 Or optPeriodo.Value = 1 Then
    DataFiltro = " and " & DataTexto & " Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "'"
    DataFiltroRel = " and {Carteira_producao." & DataTexto & "} >= Date(" & Year(msk_fltInicio.Value) & "," & Month(msk_fltInicio.Value) & "," & Day(msk_fltInicio.Value) & ") and {Carteira_producao." & DataTexto & "} <= Date(" & _
                                Year(msk_fltFim.Value) & "," & Month(msk_fltFim.Value) & "," & Day(msk_fltFim.Value) & ")"
End If

TextoFiltroValid = TextoFiltroValidEst & TextoFiltroValidProc & TextoFiltroValidPlano
TextoFiltroValidRel = TextoFiltroValidEstRel & TextoFiltroValidProcRel & TextoFiltroValidPlanoRel
TextoFiltroPadrao = "ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and " & FiltroMRP & DataFiltro
TextoFiltroPadraoRel = "{Carteira_producao.ID_empresa} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and " & FiltroMRPRel & DataFiltroRel

With frmprod
    If txtTexto <> "" Or cmbTexto <> "" Then
        If cmbfiltrarpor = "Cliente" Then
            .StrSql_Ordem_MRP = "Select * from Carteira_producao where Cliente = '" & cmbTexto & "' and " & TextoFiltroPadrao
            .FormulaRel_Ordem_Carteira = "{Carteira_producao.Cliente} = '" & cmbTexto & "' and " & TextoFiltroPadraoRel
        ElseIf cmbfiltrarpor = "Família" Then
            .StrSql_Ordem_MRP = "Select * from Carteira_producao where Familia = '" & cmbTexto & "' and " & TextoFiltroPadrao
            .FormulaRel_Ordem_Carteira = "{Carteira_producao.Familia} = '" & cmbTexto & "' and " & TextoFiltroPadraoRel
        Else
            Select Case cmbfiltrarpor
                Case "Código de referência":  TextoFiltro = "n_referencia"
                Case "Código interno": TextoFiltro = "Desenho"
                Case "Descrição": TextoFiltro = "Descricao_tecnica"
                Case "Família": TextoFiltro = "Familia"
                Case "Pedido do cliente": TextoFiltro = "PCcliente"
                Case "Pedido interno": TextoFiltro = "Ncotacao"
                Case "Grupo do cliente": TextoFiltro = "Grupo_cliente"
                Case "Solicitação de produção": TextoFiltro = "Requisicaotexto"
            End Select
            .StrSql_Ordem_MRP = "Select * from Carteira_producao where " & TextoFiltro & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and " & TextoFiltroPadrao
            .FormulaRel_Ordem_Carteira = "{Carteira_producao." & TextoFiltro & "}" & FunVerifTipoFiltroIMFRel(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and " & TextoFiltroPadraoRel
        End If
    Else
        .StrSql_Ordem_MRP = "Select * from Carteira_producao where " & TextoFiltroPadrao
        .FormulaRel_Ordem_Carteira = TextoFiltroPadraoRel
    End If
    .ID_empresa_MRP = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
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

TemProcessoTexto = ""
TemProcessoRel = ""
TemMRPTexto = ""
TemMRPRel = ""

If chkProcesso.Value = 1 Then
    TemProcessoTexto = " and Tem_processo = 'SIM'"
    TemProcessoRel = " and {Carteira_producao.Tem_processo} = 'SIM'"
End If
If chkMRP.Value = 1 Then
    TemMRPTexto = " and MRP = 'NÃO'"
    TemMRPRel = " and {Carteira_producao.MRP} = 'NÃO'"
End If

Select Case Cmb_filtrar
    Case "Com necessidade":
        FiltroMRP = "Necessidade > 0"
        FiltroMRPRel = "{Carteira_producao.Necessidade} > 0"
    Case "Sem necessidade":
        FiltroMRP = "Necessidade <= 0"
        FiltroMRPRel = "{Carteira_producao.Necessidade} <= 0"
    Case "Todos"
        FiltroMRP = "Desenho IS NOT NULL"
        FiltroMRPRel = "Not(IsNull({Carteira_producao.Desenho}))"
End Select
FiltroMRP = FiltroMRP & TemProcessoTexto & TemMRPTexto
FiltroMRPRel = FiltroMRPRel & TemProcessoRel & TemMRPRel

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

