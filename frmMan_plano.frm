VERSION 5.00
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{50D37AD9-8D3C-43DD-ADD5-7C957C951843}#1.9#0"; "FlexCell.ocx"
Begin VB.Form frmMan_plano 
   Caption         =   "Plano de manutenção"
   ClientHeight    =   10035
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15660
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10035
   ScaleWidth      =   15660
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Dados plano de manutenção"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   0
      TabIndex        =   2
      Top             =   1020
      Width           =   15645
      Begin VB.ComboBox cmbSetor 
         Height          =   315
         ItemData        =   "frmMan_plano.frx":0000
         Left            =   10920
         List            =   "frmMan_plano.frx":0002
         TabIndex        =   14
         Top             =   480
         Width           =   3375
      End
      Begin VB.TextBox txt_Tipo 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   1470
         TabIndex        =   13
         Top             =   480
         Width           =   1815
      End
      Begin VB.ComboBox cmbPeridiocidade 
         Height          =   315
         ItemData        =   "frmMan_plano.frx":0004
         Left            =   14310
         List            =   "frmMan_plano.frx":001A
         TabIndex        =   10
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txt_Descricao 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   4650
         TabIndex        =   8
         Top             =   480
         Width           =   6255
      End
      Begin VB.TextBox txt_Codigo 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   3300
         TabIndex        =   6
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txt_CodPlano 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   210
         TabIndex        =   4
         Top             =   480
         Width           =   1245
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Tipo"
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
         Left            =   2205
         TabIndex        =   12
         Top             =   270
         Width           =   315
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Peridiocidade"
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
         Left            =   14385
         TabIndex        =   11
         Top             =   270
         Width           =   975
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Setor"
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
         Left            =   12315
         TabIndex        =   9
         Top             =   270
         Width           =   405
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Descrição"
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
         Left            =   7425
         TabIndex        =   7
         Top             =   270
         Width           =   705
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Código"
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
         Left            =   3705
         TabIndex        =   5
         Top             =   270
         Width           =   525
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Código plano"
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
         Left            =   390
         TabIndex        =   3
         Top             =   270
         Width           =   945
      End
   End
   Begin FlexCell.Grid GridPL 
      Height          =   7995
      Left            =   0
      TabIndex        =   1
      Top             =   2040
      Width           =   15645
      _ExtentX        =   27596
      _ExtentY        =   14102
      BackColor2      =   14737632
      BackColorBkg    =   16777215
      Cols            =   13
      DefaultFontSize =   8.25
      GridColor       =   12632256
      Rows            =   33
   End
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   10425
      Top             =   1350
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frmMan_plano.frx":0055
      Count           =   1
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   1005
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16470
      _ExtentX        =   29051
      _ExtentY        =   1773
      ButtonCount     =   9
      GradientColor1  =   16777215
      GradientColor2  =   14737632
      GradientColorDown1=   10802943
      GradientColorDown2=   7979263
      GradientColorDownRight1=   10802943
      GradientColorDownRight2=   7979263
      GradientColorOver1=   14417407
      GradientColorOver2=   12317439
      GradientColorOverRight1=   14417407
      GradientColorOverRight2=   12317439
      IsStrech        =   -1  'True
      RightColor1     =   14737632
      RightColor2     =   16777215
      ShowEndPanel    =   0   'False
      Theme           =   1
      ButtonCaption1  =   "Novo"
      ButtonEnabled1  =   0   'False
      ButtonIconSize1 =   32
      ButtonToolTipText1=   "Novo (Insert)"
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
      ButtonWidth1    =   36
      ButtonHeight1   =   21
      ButtonUseMaskColor1=   0   'False
      ButtonCaption2  =   "Filtrar"
      ButtonEnabled2  =   0   'False
      ButtonIconSize2 =   32
      ButtonToolTipText2=   "Filtrar (F2)"
      ButtonKey2      =   "2"
      ButtonAlignment2=   2
      BeginProperty ButtonFont2 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft2     =   40
      ButtonTop2      =   2
      ButtonWidth2    =   42
      ButtonHeight2   =   21
      ButtonUseMaskColor2=   0   'False
      ButtonCaption3  =   "Salvar"
      ButtonEnabled3  =   0   'False
      ButtonIconSize3 =   32
      ButtonToolTipText3=   "Salvar (F3)"
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
      ButtonLeft3     =   84
      ButtonTop3      =   2
      ButtonWidth3    =   44
      ButtonHeight3   =   21
      ButtonUseMaskColor3=   0   'False
      ButtonCaption4  =   "Excluir"
      ButtonEnabled4  =   0   'False
      ButtonIconSize4 =   32
      ButtonToolTipText4=   "Excluir (F4)"
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
      ButtonLeft4     =   130
      ButtonTop4      =   2
      ButtonWidth4    =   45
      ButtonHeight4   =   21
      ButtonUseMaskColor4=   0   'False
      ButtonCaption5  =   "Relatório"
      ButtonEnabled5  =   0   'False
      ButtonIconSize5 =   32
      ButtonToolTipText5=   "Relatório (F5)"
      ButtonKey5      =   "5"
      ButtonAlignment5=   2
      BeginProperty ButtonFont5 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft5     =   177
      ButtonTop5      =   2
      ButtonWidth5    =   60
      ButtonHeight5   =   21
      ButtonUseMaskColor5=   0   'False
      ButtonCaption6  =   "Filtrar todos"
      ButtonEnabled6  =   0   'False
      ButtonIconSize6 =   32
      ButtonToolTipText6=   "Filtrar todos os registros."
      ButtonKey6      =   "8"
      ButtonAlignment6=   2
      BeginProperty ButtonFont6 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft6     =   239
      ButtonTop6      =   2
      ButtonWidth6    =   77
      ButtonHeight6   =   21
      ButtonUseMaskColor6=   0   'False
      ButtonEnabled7  =   0   'False
      ButtonIconSize7 =   32
      ButtonAlignment7=   2
      ButtonType7     =   1
      ButtonStyle7    =   -1
      BeginProperty ButtonFont7 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState7    =   -1
      ButtonLeft7     =   318
      ButtonTop7      =   4
      ButtonWidth7    =   2
      ButtonHeight7   =   56
      ButtonCaption8  =   "Sair"
      ButtonEnabled8  =   0   'False
      ButtonIconSize8 =   32
      ButtonToolTipText8=   "Sair (Esc)"
      ButtonKey8      =   "13"
      ButtonAlignment8=   2
      BeginProperty ButtonFont8 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft8     =   322
      ButtonTop8      =   2
      ButtonWidth8    =   30
      ButtonHeight8   =   21
      ButtonUseMaskColor8=   0   'False
      ButtonEnabled9  =   0   'False
      ButtonIconSize9 =   32
      ButtonKey9      =   "15"
      ButtonAlignment9=   2
      BeginProperty ButtonFont9 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState9    =   5
      ButtonLeft9     =   354
      ButtonTop9      =   2
      ButtonWidth9    =   24
      ButtonHeight9   =   24
      ButtonUseMaskColor9=   0   'False
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
         FormHeightDT    =   10500
         FormWidthDT     =   15780
         FormScaleHeightDT=   10035
         FormScaleWidthDT=   15660
         ResizeFormBackground=   -1  'True
         ResizePictureBoxContents=   -1  'True
      End
   End
End
Attribute VB_Name = "frmMan_plano"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15500, 9, True
ProcAjustaGridPL
ProcExibePaginaGrid
ProcBuscaDadosSetor

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub GridPL_Click()
On Error GoTo tratar_erro

txt_CodPlano = GridPL.Cell(GridPL.ActiveCell.Row, 1).Text
txt_Codigo = GridPL.Cell(GridPL.ActiveCell.Row, 2).Text
txt_Descricao = GridPL.Cell(GridPL.ActiveCell.Row, 3).Text
cmbSetor.Text = GridPL.Cell(GridPL.ActiveCell.Row, 4).Text
cmbPeridiocidade = GridPL.Cell(GridPL.ActiveCell.Row, 5).Text

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txt_CodPlano_Change()
On Error GoTo tratar_erro

ProcBuscaDadosPlano

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcNovo
'    Case 2: ProcFiltrar
    Case 3: ProcEnviaDadosPlano
'    Case 4: ProcExcluir
'    Case 5: ProcImprimir
    Case 8: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovo()
On Error GoTo tratar_erro

If USMsgBox("Deseja realmente criar um novo plano de manutenção?", vbYesNo, "CAPRIND v5.0") = vbYes Then
frmManutencao_Plano_novo.Show 1
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcAjustaGridPL()
On Error GoTo tratar_erro

With GridPL
    .Cols = 9
    .rows = 31
    .AllowUserPaste = cellTextOnly
    .AllowUserResizing = False
    .ExtendLastCol = True
    .BoldFixedCell = False
    .DisplayDateTimeMask = True
    .DisplayFocusRect = False
    .SelectionMode = cellSelectionByRow

    .DrawMode = cellOwnerDraw
    
    .Appearance = Flat
    .ScrollBarStyle = Flat
    .FixedRowColStyle = Flat
    .Cell(0, 1).Text = "Plano"
    .Cell(0, 2).Text = "Equipamento"
    .Cell(0, 3).Text = "Descrição"
    .Cell(0, 4).Text = "Setor"
    .Cell(0, 5).Text = "Periodicidade"
    .Cell(0, 6).Text = "Realizada"
    .Cell(0, 7).Text = "Próximo"
    .Cell(0, 8).Text = "Status"
        
    .Column(1).CellType = cellTextBox
    .Column(1).Alignment = cellCenterCenter
    
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
 
    .Column(0).Width = 10
    .Column(1).Width = 90
    .Column(2).Width = 75
    .Column(3).Width = 280
    .Column(4).Width = 190
    .Column(5).Width = 100
    .Column(6).Width = 100
    .Column(7).Width = 120
    .Column(8).Width = 120
End With
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcBuscaDadosPlano()
On Error GoTo tratar_erro

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Manutencao_Plano where CodPlano = '" & txt_CodPlano.Text & "'order by ID", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then

If TBAbrir!Predial = True Then
txt_Tipo = "Predial"
End If

If TBAbrir!Produto = True Then
txt_Tipo = "Produto"
End If

If TBAbrir!Equipamento = True Then
txt_Tipo = "Equipamento"
End If

End If
TBAbrir.Close
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcBuscaDadosSetor()
On Error GoTo tratar_erro

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from usuarios_setor order by Setor", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then

Do While TBAbrir.EOF = False
cmbSetor.AddItem TBAbrir!Setor
TBAbrir.MoveNext
Loop

End If
TBAbrir.Close
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub ProcEnviaDadosPlano()
On Error GoTo tratar_erro

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Manutencao_Plano where CodPlano = '" & txt_CodPlano.Text & "'order by ID", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = True Then
TBAbrir.AddNew
End If

If txt_Tipo = "Predial" Then
TBAbrir!Predial = True
TBAbrir!Produto = False
TBAbrir!Equipamento = False
End If

If txt_Tipo = "Produto" Then
TBAbrir!Predial = False
TBAbrir!Produto = True
TBAbrir!Equipamento = False
End If

If txt_Tipo = "Equipamento" Then
TBAbrir!Predial = False
TBAbrir!Produto = False
TBAbrir!Equipamento = True
End If

TBAbrir!Codplano = txt_CodPlano
TBAbrir!CODIGO = txt_Codigo
TBAbrir!Descricao = txt_Descricao
TBAbrir!Setor = cmbSetor.Text
TBAbrir!Peridiocidade = cmbPeridiocidade


TBAbrir.Update

TBAbrir.Close
ProcExibePaginaGrid
USMsgBox "Dados salvos com sucesso!", vbInformation, "CAPRIND v5.0"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePaginaGrid()
On Error GoTo tratar_erro

GridPL.rows = 1
Contador = 1

Set TBLISTA = CreateObject("adodb.recordset")
        TBLISTA.Open "Select * from Manutencao_Plano order by ID", Conexao, adOpenKeyset, adLockOptimistic

    Do While TBLISTA.EOF = False
    
        GridPL.AddItem TBLISTA!Codplano & vbTab & _
            TBLISTA!CODIGO & vbTab & _
            TBLISTA!Descricao & vbTab & _
            TBLISTA!Setor & vbTab & _
            TBLISTA!Peridiocidade & vbTab & _
            TBLISTA!Realizada & vbTab & _
            TBLISTA!Realizada & vbTab & _
            TBLISTA!Realizada & vbTab & _
            TBLISTA!status
            Contador = Contador + 1
            TBLISTA.MoveNext
    Loop

GridPL.AutoRedraw = True
GridPL.Refresh


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
