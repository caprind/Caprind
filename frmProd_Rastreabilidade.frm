VERSION 5.00
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{50D37AD9-8D3C-43DD-ADD5-7C957C951843}#1.9#0"; "FlexCell.ocx"
Begin VB.Form frmProd_Rastreabilidade 
   BackColor       =   &H00E0E0E0&
   Caption         =   "PCP - Rastreabilidade numero de série"
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
   Icon            =   "frmProd_Rastreabilidade.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10005
   ScaleWidth      =   15375
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Opções para pesquisa"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1125
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   15225
      Begin DrawSuite2022.USButton btnSair 
         Height          =   645
         Left            =   13770
         TabIndex        =   7
         ToolTipText     =   "Fechar formulário"
         Top             =   270
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   1138
         DibPicture      =   "frmProd_Rastreabilidade.frx":000C
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
         PicSize         =   2
         PicSizeH        =   24
         PicSizeW        =   24
         ShowFocusRect   =   0   'False
         Theme           =   4
      End
      Begin VB.ComboBox cmbOpcao 
         Height          =   315
         ItemData        =   "frmProd_Rastreabilidade.frx":2DB9
         Left            =   2790
         List            =   "frmProd_Rastreabilidade.frx":2DBB
         Style           =   2  'Dropdown List
         TabIndex        =   5
         ToolTipText     =   "Escolha um valor para filtrar"
         Top             =   600
         Width           =   1905
      End
      Begin VB.ComboBox cmbFiltrarPor 
         Height          =   315
         ItemData        =   "frmProd_Rastreabilidade.frx":2DBD
         Left            =   270
         List            =   "frmProd_Rastreabilidade.frx":2DBF
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "Escolha uma opção pra filtrar"
         Top             =   600
         Width           =   2505
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Opção de filtro"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   3
         Left            =   3210
         TabIndex        =   6
         Top             =   390
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filtrar por"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   20
         Left            =   1170
         TabIndex        =   4
         Top             =   390
         Width           =   705
      End
   End
   Begin DrawSuite2022.USGroupBox USGroupBox1 
      Height          =   8595
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   15225
      _ExtentX        =   26855
      _ExtentY        =   15161
      Caption         =   "Lista de numeros de série utilizados na produção"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   128
      GradientColor1  =   14737632
      Begin FlexCell.Grid GridSerie 
         Height          =   8265
         Left            =   0
         TabIndex        =   1
         Top             =   330
         Width           =   15225
         _ExtentX        =   26855
         _ExtentY        =   14579
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
   End
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   26
      ScreenHeight    =   1080
      ScreenWidth     =   1920
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
Attribute VB_Name = "frmProd_Rastreabilidade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private OpcaoFiltro As String

Public Sub ProcAjustaGridSerie()
On Error GoTo tratar_erro

With GridSerie

    .AllowUserPaste = cellTextOnly
    .AllowUserResizing = False
    .ExtendLastCol = True
    .BoldFixedCell = False
    .DisplayDateTimeMask = True
    .DisplayFocusRect = False
    .SelectionMode = cellSelectionFree
    .Cols = 9
    
    .DrawMode = cellOwnerDraw
    
    .Appearance = Flat
    .ScrollBarStyle = Flat
    .FixedRowColStyle = Flat
    .Cell(0, 1).Text = "Item"

    .Cell(0, 2).Text = "OS"

    .Cell(0, 3).Text = "Fase"

    .Cell(0, 4).Text = "Posto trabalho"

    .Cell(0, 5).Text = "Numero de série"
    
    .Cell(0, 6).Text = "Status"
         
    .Cell(0, 7).Text = "Data"
         
    .Cell(0, 8).Text = "Responsável"
    
    .Column(0).Width = 8
    .Column(1).Width = 30
    .Column(2).Width = 60
    .Column(3).Width = 50
    .Column(4).Width = 250
    .Column(5).Width = 100
    .Column(6).Width = 180
    .Column(7).Width = 180
    .Column(8).Width = 180
        
          
    .Column(1).CellType = cellTextBox
    .Column(1).Alignment = cellCenterCenter
    .Column(1).Locked = True
    
    .Column(2).CellType = cellTextBox
    .Column(2).Alignment = cellCenterCenter
    
    .Column(3).CellType = cellTextBox
    .Column(3).Alignment = cellCenterCenter
    
    .Column(4).CellType = cellTextBox
    .Column(4).Alignment = cellCenterCenter
    
    .Column(5).CellType = cellTextBox
    .Column(5).Alignment = cellCenterCenter
    
    .Column(6).CellType = cellTextBox
    .Column(6).Alignment = cellCenterCenter
    
    .Column(7).CellType = cellTextBox
    .Column(7).Alignment = cellCenterCenter
    
    .Column(8).CellType = cellTextBox
    .Column(8).Alignment = cellCenterCenter
    
  
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaListaItens()
On Error GoTo tratar_erro
Dim L As Long

With GridSerie
    
 L = 1
.rows = 1
.Cols = 9

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
  If TBAbrir.EOF = False Then
    Contador2 = 1 'TBAbrir.RecordCount
    
        Do While TBAbrir.EOF = False
         .AddItem Contador2
         .Cell(Contador2, 2).Text = TBAbrir!OS
         .Cell(Contador2, 3).Text = TBAbrir!Fase
         .Cell(Contador2, 4).Text = TBAbrir!maquina
         .Cell(Contador2, 5).Text = TBAbrir!N_Serie
         .Cell(Contador2, 6).Text = TBAbrir!status
         .Cell(Contador2, 7).Text = IIf(IsNull(TBAbrir!Data), "", TBAbrir!Data)
         .Cell(Contador2, 8).Text = IIf(IsNull(TBAbrir!Responsavel), "", TBAbrir!Responsavel)
         Contador2 = Contador2 + 1
         TBAbrir.MoveNext
        Loop
  End If


End With
TBAbrir.Close

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Private Sub btnSair_Click()
On Error GoTo tratar_erro

Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbfiltrarpor_Change()
On Error GoTo tratar_erro

Select Case cmbfiltrarpor.Text
    Case "ORDEM DE SERVIÇO": OpcaoFiltro = "OS"
    Case "FASE": OpcaoFiltro = "Fase"
    Case "POSTO DE TRABALHO": OpcaoFiltro = "Maquina"
    Case "NUMERO DE SÉRIE": OpcaoFiltro = "N_serie"
    Case "STATUS": OpcaoFiltro = "Status"
End Select

StrSql = "Select * from Producao_NumeroSerie where " & OpcaoFiltro & " = '" & cmbOpcao & "' and Ordem '" & frmprod.txtof.Text & "' AND Status <> '' ORDER BY OS, N_serie"
'Debug.print StrSql


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

StrSql = "Select DISTINCT " & OpcaoFiltro & " from Producao_Rastreavel where Ordem = '" & frmprod.txtof.Text & "' AND Status <> '' "
'Debug.print StrSql

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
  If TBAbrir.EOF = False Then
    Do While TBAbrir.EOF = False
    
        If OpcaoFiltro = "OS" Then Texto = TBAbrir!OS
        If OpcaoFiltro = "Fase" Then Texto = TBAbrir!Fase
        If OpcaoFiltro = "Maquina" Then Texto = TBAbrir!maquina
        If OpcaoFiltro = "N_serie" Then Texto = TBAbrir!N_Serie
        If OpcaoFiltro = "Status" Then Texto = TBAbrir!status
        
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

Private Sub cmbfiltrarpor_Click()
On Error GoTo tratar_erro

Select Case cmbfiltrarpor.Text
    Case "ORDEM DE SERVIÇO": OpcaoFiltro = "OS"
    Case "FASE": OpcaoFiltro = "Fase"
    Case "POSTO DE TRABALHO": OpcaoFiltro = "Maquina"
    Case "NUMERO DE SÉRIE": OpcaoFiltro = "N_serie"
    Case "STATUS": OpcaoFiltro = "Status"
End Select

StrSql = "Select * from Producao_Rastreavel where " & OpcaoFiltro & " = '" & cmbOpcao & "' and Ordem '" & frmprod.txtof.Text & "' AND Status <> '' ORDER BY OS, N_serie"
'Debug.print StrSql


procCarregaComboOpcao

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbOpcao_Click()
On Error GoTo tratar_erro

If cmbfiltrarpor <> "" Then
Select Case cmbfiltrarpor.Text
    Case "ORDEM DE SERVIÇO": OpcaoFiltro = "OS"
    Case "FASE": OpcaoFiltro = "Fase"
    Case "POSTO DE TRABALHO": OpcaoFiltro = "Maquina"
    Case "NUMERO DE SÉRIE": OpcaoFiltro = "N_serie"
    Case "STATUS": OpcaoFiltro = "Status"
End Select

StrSql = "Select * from Producao_Rastreavel where " & OpcaoFiltro & " = '" & cmbOpcao & "' and Ordem = '" & frmprod.txtof.Text & "' AND Status <> '' ORDER BY OS, N_serie"
'Debug.print StrSql

ProcCarregaListaItens
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcRemoveObjetosResize Me
ProcAjustaGridSerie
procCarregarComboFiltrarPor

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procCarregarComboFiltrarPor()
On Error GoTo tratar_erro

    With cmbfiltrarpor
        .AddItem "FASE"
        .AddItem "NUMERO DE SÉRIE"
        .AddItem "ORDEM DE SERVIÇO"
        .AddItem "POSTO DE TRABALHO"
        .AddItem "STATUS"
    End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

