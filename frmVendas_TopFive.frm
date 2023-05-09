VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2014.ocx"
Object = "{50D37AD9-8D3C-43DD-ADD5-7C957C951843}#1.9#0"; "FlexCell.ocx"
Begin VB.Form frmVendas_TopFive 
   Caption         =   "Administrativo - Vendas - Relatorios - Desempenho"
   ClientHeight    =   10005
   ClientLeft      =   480
   ClientTop       =   405
   ClientWidth     =   15375
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10005
   ScaleWidth      =   15375
   WindowState     =   2  'Maximized
   Begin VB.Frame frameAnoMes 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Filtrar do ano e do mês"
      Height          =   1125
      Left            =   1860
      TabIndex        =   15
      Top             =   0
      Width           =   2625
      Begin VB.ComboBox cmbdoAno 
         Height          =   315
         ItemData        =   "frmVendas_TopFive.frx":0000
         Left            =   210
         List            =   "frmVendas_TopFive.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   17
         ToolTipText     =   "Escolha o ano para filtrar"
         Top             =   600
         Width           =   945
      End
      Begin VB.ComboBox cmbdoMes 
         Height          =   315
         ItemData        =   "frmVendas_TopFive.frx":0004
         Left            =   1380
         List            =   "frmVendas_TopFive.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   16
         ToolTipText     =   "Escolha o ano para filtrar"
         Top             =   600
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Do ano"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   0
         Left            =   420
         TabIndex        =   19
         Top             =   390
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mês"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   1
         Left            =   1590
         TabIndex        =   18
         Top             =   390
         Width           =   285
      End
   End
   Begin VB.Frame FramePeriodo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Filtrar do período"
      Height          =   1125
      Left            =   1860
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   2625
      Begin MSComCtl2.DTPicker msk_ate 
         Height          =   315
         Left            =   1380
         TabIndex        =   11
         ToolTipText     =   "Data final."
         Top             =   450
         Width           =   1185
         _ExtentX        =   2090
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
         Format          =   113180673
         CurrentDate     =   39057
      End
      Begin MSComCtl2.DTPicker msk_de 
         Height          =   315
         Left            =   90
         TabIndex        =   12
         ToolTipText     =   "Data inicio."
         Top             =   450
         Width           =   1185
         _ExtentX        =   2090
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
         Format          =   113180673
         CurrentDate     =   39057
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Até"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   2
         Left            =   1830
         TabIndex        =   14
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "De"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   4
         Left            =   510
         TabIndex        =   13
         Top             =   240
         Width           =   195
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   1125
      Left            =   4500
      TabIndex        =   4
      Top             =   0
      Width           =   10875
      Begin VB.ComboBox cmbOpcao 
         Height          =   315
         ItemData        =   "frmVendas_TopFive.frx":0008
         Left            =   2370
         List            =   "frmVendas_TopFive.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   6
         ToolTipText     =   "Escolha um valor para filtrar"
         Top             =   570
         Width           =   4755
      End
      Begin VB.ComboBox cmbFiltrarPor 
         Height          =   315
         ItemData        =   "frmVendas_TopFive.frx":000C
         Left            =   270
         List            =   "frmVendas_TopFive.frx":0019
         Style           =   2  'Dropdown List
         TabIndex        =   5
         ToolTipText     =   "Escolha uma opção pra filtrar"
         Top             =   570
         Width           =   2085
      End
      Begin DrawSuite2014.USButton btnSair 
         Height          =   765
         Left            =   9450
         TabIndex        =   7
         ToolTipText     =   "Fechar formulário"
         Top             =   240
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   1349
         DibPicture      =   "frmVendas_TopFive.frx":0039
         BorderColor     =   5263559
         BorderColorDisabled=   13160660
         BorderColorDown =   4013465
         BorderColorOver =   4408288
         Caption         =   "Sair"
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
      Begin DrawSuite2014.USButton btnFiltrar 
         Height          =   765
         Left            =   8160
         TabIndex        =   20
         ToolTipText     =   "Filtrar registros de vendas"
         Top             =   240
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   1349
         DibPicture      =   "frmVendas_TopFive.frx":2DE6
         BorderColor     =   4960354
         BorderColorDisabled=   13160660
         BorderColorDown =   4210752
         BorderColorOver =   49152
         Caption         =   "Filtrar"
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
         PicAlign        =   8
         PicSize         =   2
         PicSizeH        =   24
         PicSizeW        =   24
         ShowFocusRect   =   0   'False
         ShowFocusRectDown=   0   'False
         Theme           =   3
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Opção de filtro"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   3
         Left            =   4215
         TabIndex        =   9
         Top             =   360
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filtrar por"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   20
         Left            =   870
         TabIndex        =   8
         Top             =   360
         Width           =   705
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Opções para pesquisa"
      Height          =   1125
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1845
      Begin DrawSuite2014.USOptionButton optPeriodo 
         Height          =   315
         Left            =   270
         TabIndex        =   2
         Top             =   630
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
         Caption         =   "Período"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   128
         ShowFocusRect   =   0   'False
      End
      Begin DrawSuite2014.USOptionButton optMesAno 
         CausesValidation=   0   'False
         Height          =   315
         Left            =   270
         TabIndex        =   3
         Top             =   390
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         Caption         =   "Ano x Mês"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   128
         ShowFocusRect   =   0   'False
      End
   End
   Begin FlexCell.Grid GridSerie 
      Height          =   8655
      Left            =   0
      TabIndex        =   0
      Top             =   1140
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   15266
      AllowUserReorderColumn=   -1  'True
      AllowUserResizing=   0   'False
      Appearance      =   0
      BackColor2      =   14737632
      BackColorBkg    =   -2147483644
      BorderColor     =   12632256
      CellBorderColor =   8421504
      SelectionBorderColor=   4210752
      Cols            =   9
      DefaultFontSize =   6.75
      FixedRowColStyle=   2
      GridColor       =   12632256
      Rows            =   1
      ScrollBars      =   2
      ScrollBarStyle  =   0
      SelectionMode   =   1
      MultiSelect     =   0   'False
      EnterKeyMoveTo  =   1
      AllowUserPaste  =   3
   End
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   99
      ScreenHeight    =   768
      ScreenWidth     =   1360
      ScreenHeightDT  =   1080
      ScreenWidthDT   =   1920
      AutoResizeOnLoad=   0   'False
      ApplicationName =   "Active Resize Control Professional"
      FormHeightDT    =   10470
      FormWidthDT     =   15495
      FormScaleHeightDT=   10005
      FormScaleWidthDT=   15375
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
End
Attribute VB_Name = "frmVendas_TopFive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnSair_Click()
On Error GoTo tratar_erro

Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbdoAno_Click()
On Error GoTo tratar_erro

procCarregaComboMes

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbfiltrarpor_Click()
On Error GoTo tratar_erro
Dim Mes As String
'==========================================================
' Se for filtrar por ano x mês
'==========================================================
If optMesAno.Value = True Then
    Mes = cmbdoMes.Text
    
            Select Case Mes
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
    
    
    Select Case cmbFiltrarPor.Text
        Case "Cliente":  StrSql = "Select DISTINCT Cliente from Vendas_relatorios_historico_clientes Where Mes = '" & Mes & "' and Ano = '" & cmbdoAno.Text & "'"
        Case "Produto":  StrSql = "Select DISTINCT Desenho from Vendas_relatorios_historico_Produtos Where Mes = '" & Mes & "' and Ano = '" & cmbdoAno.Text & "'"
        Case "Vendedor": StrSql = "Select DISTINCT Vendedor from Vendas_relatorios_historico_Vendedor Where Mes = '" & Mes & "' and Ano = '" & cmbdoAno.Text & "'"
    End Select
Else
'=====================================================
' Se for filtrar por período
'=====================================================
    Select Case cmbFiltrarPor.Text
        Case "Cliente":  StrSql = "Select DISTINCT Cliente from Vendas_relatorios_historico_Detalhado Where DataVendas >= '" & msk_de & "' and DataVendas <= '" & msk_ate & "'"
        Case "Produto":  StrSql = "Select DISTINCT Desenho from Vendas_relatorios_historico_Detalhado Where DataVendas >= '" & msk_de & "' and DataVendas <= '" & msk_ate & "'"
        Case "Vendedor": StrSql = "Select DISTINCT Vend_ext as Vendedor from Vendas_relatorios_historico_Detalhado Where DataVendas >= '" & msk_de & "' and DataVendas <= '" & msk_ate & "'"
    End Select
End If
Debug.Print StrSql


procCarregaComboOpcao

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procCarregaComboOpcao()
On Error GoTo tratar_erro
cmbOpcao.Clear
Dim Texto As String

'StrSql = "Select DISTINCT " & OpcaoFiltro & " from Producao_NumeroSerie where Ordem = '" & frmprod.txtof.Text & "' AND Status <> '' "
'Debug.Print StrSql

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
  If TBAbrir.EOF = False Then
    cmbOpcao.AddItem "TOP FIVE"
    Do While TBAbrir.EOF = False
        If cmbFiltrarPor.Text = "Cliente" Then Texto = TBAbrir!Cliente
        If cmbFiltrarPor.Text = "Produto" Then Texto = TBAbrir!Desenho
        If cmbFiltrarPor.Text = "Vendedor" Then Texto = IIf(IsNull(TBAbrir!Vendedor), "", TBAbrir!Vendedor)
        
        With cmbOpcao
            .AddItem Texto
        End With
        TBAbrir.MoveNext
    Loop
  End If
  TBAbrir.Close
  
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaComboAno()
On Error GoTo tratar_erro
cmbdoAno.Clear

StrSql = "Select DISTINCT Ano from Vendas_relatorios_historico_detalhado order by Ano Desc"

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
  If TBAbrir.EOF = False Then
    Do While TBAbrir.EOF = False
        
        With cmbdoAno
            .AddItem TBAbrir!Ano
        End With
        TBAbrir.MoveNext
    Loop
  End If
  TBAbrir.Close
  
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procCarregaComboMes()
On Error GoTo tratar_erro
cmbdoMes.Clear
Dim Mes As String

If cmbdoAno.Text <> "" Then
StrSql = "Select DISTINCT Mes from Vendas_relatorios_historico_detalhado Where Ano = '" & cmbdoAno.Text & "'order by Mes"

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
  If TBAbrir.EOF = False Then
    Do While TBAbrir.EOF = False
        Select Case TBAbrir!Mes
            Case 1: Mes = "Janeiro"
            Case 2: Mes = "Fevereiro"
            Case 3: Mes = "Março"
            Case 4: Mes = "Abril"
            Case 5: Mes = "Maio"
            Case 6: Mes = "Junho"
            Case 7: Mes = "Julho"
            Case 8: Mes = "Agosto"
            Case 9: Mes = "Setembro"
            Case 10: Mes = "Outubro"
            Case 11: Mes = "Novembro"
            Case 12: Mes = "Dezembro"
        End Select
        
        With cmbdoMes
            .AddItem Mes
        End With
        TBAbrir.MoveNext
    Loop
  End If
  TBAbrir.Close
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaComboAno
procCarregaComboMes
optMesAno.Value = True
msk_de.Value = Date
msk_ate.Value = Date

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optMesAno_Click()
On Error GoTo tratar_erro

FramePeriodo.Visible = optPeriodo.Value
frameAnoMes.Visible = optMesAno.Value

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optPeriodo_Click()
On Error GoTo tratar_erro


FramePeriodo.Visible = optPeriodo.Value
frameAnoMes.Visible = optMesAno.Value

  
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
