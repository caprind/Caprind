VERSION 5.00
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{50D37AD9-8D3C-43DD-ADD5-7C957C951843}#1.9#0"; "FlexCell.ocx"
Begin VB.Form frmVendas_Comissoes_Metas 
   Caption         =   "Administrativo - Vendas - Comissões - Metas"
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Status recebimentos"
      Height          =   1125
      Left            =   30
      TabIndex        =   10
      Top             =   990
      Width           =   1905
      Begin DrawSuite2022.USOptionButton optrecebida 
         Height          =   315
         Left            =   180
         TabIndex        =   11
         ToolTipText     =   "Filtrar por período"
         Top             =   660
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
         Caption         =   "Recebidas"
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
      Begin DrawSuite2022.USOptionButton optreceber 
         CausesValidation=   0   'False
         Height          =   315
         Left            =   180
         TabIndex        =   12
         ToolTipText     =   "Filtrar por ano e mês"
         Top             =   390
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         Caption         =   "Á receber"
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Filtro"
      Height          =   1125
      Left            =   4710
      TabIndex        =   1
      Top             =   990
      Width           =   10665
      Begin VB.ComboBox cmbOpcao 
         Height          =   315
         ItemData        =   "frmVendas_Comissoes_Metas.frx":0000
         Left            =   270
         List            =   "frmVendas_Comissoes_Metas.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "Escolha um vendedor para filtrar"
         Top             =   570
         Width           =   4425
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedor"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   3
         Left            =   2137
         TabIndex        =   3
         Top             =   360
         Width           =   690
      End
   End
   Begin FlexCell.Grid GridLista 
      Height          =   7875
      Left            =   30
      TabIndex        =   0
      Top             =   2130
      Width           =   15345
      _ExtentX        =   27067
      _ExtentY        =   13891
      AllowUserReorderColumn=   -1  'True
      AllowUserResizing=   0   'False
      Appearance      =   0
      BackColor2      =   14737632
      BackColorBkg    =   16777215
      BackColorFixed  =   14737632
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
   Begin VB.Frame frameAnoMes 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Filtrar do ano e do mês"
      Height          =   1125
      Left            =   1950
      TabIndex        =   4
      Top             =   990
      Width           =   2745
      Begin VB.ComboBox cmbdoAno 
         Height          =   315
         ItemData        =   "frmVendas_Comissoes_Metas.frx":0004
         Left            =   270
         List            =   "frmVendas_Comissoes_Metas.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   6
         ToolTipText     =   "Escolha o ano para filtrar"
         Top             =   570
         Width           =   945
      End
      Begin VB.ComboBox cmbdoMes 
         Height          =   315
         ItemData        =   "frmVendas_Comissoes_Metas.frx":0008
         Left            =   1440
         List            =   "frmVendas_Comissoes_Metas.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   5
         ToolTipText     =   "Escolha o mês para filtrar"
         Top             =   570
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Do ano"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   0
         Left            =   487
         TabIndex        =   8
         Top             =   360
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mês"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   1
         Left            =   1770
         TabIndex        =   7
         Top             =   360
         Width           =   285
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   30
      TabIndex        =   9
      Top             =   0
      Width           =   15345
      _ExtentX        =   27067
      _ExtentY        =   1720
      ButtonCount     =   6
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
      ButtonCaption2  =   "Relatório"
      ButtonEnabled2  =   0   'False
      ButtonIconSize2 =   32
      ButtonToolTipText2=   "Relatório (F5)"
      ButtonKey2      =   "2"
      ButtonAlignment2=   2
      BeginProperty ButtonFont2 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft2     =   40
      ButtonTop2      =   2
      ButtonWidth2    =   51
      ButtonHeight2   =   21
      ButtonUseMaskColor2=   0   'False
      ButtonCaption3  =   "Exportar"
      ButtonEnabled3  =   0   'False
      ButtonIconSize3 =   32
      ButtonToolTipText3=   "Exportar relatório"
      ButtonKey3      =   "3"
      ButtonAlignment3=   2
      BeginProperty ButtonFont3 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft3     =   93
      ButtonTop3      =   2
      ButtonWidth3    =   47
      ButtonHeight3   =   21
      ButtonUseMaskColor3=   0   'False
      ButtonEnabled4  =   0   'False
      ButtonIconSize4 =   32
      ButtonAlignment4=   2
      ButtonType4     =   1
      ButtonStyle4    =   -1
      BeginProperty ButtonFont4 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState4    =   -1
      ButtonLeft4     =   142
      ButtonTop4      =   4
      ButtonWidth4    =   2
      ButtonHeight4   =   54
      ButtonCaption5  =   "Ajuda"
      ButtonEnabled5  =   0   'False
      ButtonIconSize5 =   32
      ButtonToolTipText5=   "Ajuda (F1)"
      ButtonKey5      =   "4"
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
      ButtonLeft5     =   146
      ButtonTop5      =   2
      ButtonWidth5    =   36
      ButtonHeight5   =   21
      ButtonUseMaskColor5=   0   'False
      ButtonCaption6  =   "Sair"
      ButtonEnabled6  =   0   'False
      ButtonIconSize6 =   32
      ButtonToolTipText6=   "Sair (Esc)"
      ButtonKey6      =   "5"
      ButtonAlignment6=   2
      BeginProperty ButtonFont6 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft6     =   184
      ButtonTop6      =   2
      ButtonWidth6    =   26
      ButtonHeight6   =   21
      ButtonUseMaskColor6=   0   'False
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   9990
         Top             =   120
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmVendas_Comissoes_Metas.frx":000C
         Count           =   1
      End
   End
End
Attribute VB_Name = "frmVendas_Comissoes_Metas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ProcFiltrarCompleto()
On Error GoTo tratar_erro
Dim Mes As String

'==========================================================
' Se for filtrar por ano x mês
'==========================================================

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
    
    
            If cmbOpcao.Text <> "" Then
                If optrecebida.Value = True Then
                    StrSql = "Select * from Vendas_relatorios_historico_vendedor_Mes Where Mes = '" & Mes & "' and Ano = '" & cmbdoAno.Text & "' and vendedor  = '" & cmbOpcao.Text & "' and logSit = 'S' ORDER BY ValorRecebido DESC"
                Else
                    StrSql = "Select * from Vendas_relatorios_historico_vendedor_Mes Where Mes = '" & Mes & "' and Ano = '" & cmbdoAno.Text & "' and vendedor  = '" & cmbOpcao.Text & "' and logSit = 'N' ORDER BY ValorRecebido DESC"
                End If
            Else
                If optrecebida.Value = True Then
                    StrSql = "Select * from Vendas_relatorios_historico_vendedor_Mes Where Mes = '" & Mes & "' and Ano = '" & cmbdoAno.Text & "' and logSit = 'S' ORDER BY ValorRecebido DESC"
                Else
                    StrSql = "Select * from Vendas_relatorios_historico_vendedor_Mes Where Mes = '" & Mes & "' and Ano = '" & cmbdoAno.Text & "' and logSit = 'N' ORDER BY ValorRecebido DESC"
                End If
            End If
            'Debug.print StrSql
            ProcCarregaListaVendedores
            


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub btnFiltrar_Click()
On Error GoTo tratar_erro


    ProcFiltrarCompleto

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub btnImprimir_Click()
On Error GoTo tratar_erro

ProcImprimir

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

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

If optrecebida.Value = True Then
    StrSql = "Select DISTINCT Mes from Vendas_relatorios_recebimentos_vendedor_Mes Where Ano = '" & cmbdoAno.Text & "'order by Mes"
Else
    StrSql = "Select DISTINCT MesVencer as Mes from Vendas_relatorios_recebimentos_vendedor_Mes Where AnoVencer = '" & cmbdoAno.Text & "'order by MesVencer"
End If
    
procCarregaComboMes

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub cmbfiltrarpor_Click()
On Error GoTo tratar_erro
Dim Mes As String

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
    
    
StrSql = "Select DISTINCT Vendedor from Vendas_relatorios_recebimentos_vendedor_Mes Where Mes = '" & Mes & "' and Ano = '" & cmbdoAno.Text & "'"

'Debug.print StrSql
ProcAjustaGridVendedor

    

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
'Debug.print StrSql

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
  If TBAbrir.EOF = False Then
    cmbOpcao.AddItem ""
    Do While TBAbrir.EOF = False
        
        With cmbOpcao
            .AddItem TBAbrir!vendedor
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
cmbdoMes.Clear

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
  If TBAbrir.EOF = False Then
    Do While TBAbrir.EOF = False
      If IsNull(TBAbrir!Ano) = False Then
        With cmbdoAno
            .AddItem IIf(IsNull(TBAbrir!Ano), "", TBAbrir!Ano)
        End With
       End If
        TBAbrir.MoveNext
    Loop
  End If
  TBAbrir.Close
  
ProcAjustaGridVendedor
 
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


Private Sub cmbdoMes_Change()
On Error GoTo tratar_erro
Dim Mes As String

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
    
If optrecebida.Value = True Then
StrSql = "Select DISTINCT Vendedor from Vendas_relatorios_recebimentos_vendedor_Mes Where Mes = '" & Mes & "' and Ano = '" & cmbdoAno.Text & "' and logsit = 'S'"
Else
StrSql = "Select DISTINCT Vendedor from Vendas_relatorios_recebimentos_vendedor_Mes Where MesVencer = '" & Mes & "' and AnoVencer = '" & cmbdoAno.Text & "' and logsit = 'N'"
End If

'Debug.print StrSql
ProcAjustaGridVendedor

    

procCarregaComboOpcao

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub

End Sub

Private Sub cmbdoMes_Click()
On Error GoTo tratar_erro
Dim Mes As String

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
    
If optrecebida.Value = True Then
StrSql = "Select DISTINCT Vendedor from Vendas_relatorios_recebimentos_vendedor_Mes Where Mes = '" & Mes & "' and Ano = '" & cmbdoAno.Text & "' and logsit = 'S'"
Else
StrSql = "Select DISTINCT Vendedor from Vendas_relatorios_recebimentos_vendedor_Mes Where MesVencer = '" & Mes & "' and AnoVencer = '" & cmbdoAno.Text & "' and logsit = 'N'"
End If

'Debug.print StrSql
ProcAjustaGridVendedor

procCarregaComboOpcao

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro
ProcCarregaToolBar1 Me, 16200, 6, True


'cmbdoMes.Text = Month(Date)
ProcAjustaGridVendedor


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro

 With GridLista.PageSetup
        .HeaderAlignment = cellCenter
        .HeaderFont.Name = "Tahoma"
        .HeaderFont.size = 12
        .PrintCellBorders = True
        .PrintTitleColumns = True
        .PrintFixedColumn = True
        .PrintFixedRow = True
        .PrintGridlines = True
        .ThinBorderLineWidth = 1
        
        
        .Orientation = cellLandscape
        .PaperSize = cellPaperA4
        .LeftMargin = 1
        .TopMargin = 2
        .RightMargin = 1
        .BottomMargin = 2
        .HeaderMargin = 1
        .FooterMargin = 1
        .Footer = "Pag &P de &N"
        .FooterAlignment = cellRight
        .FooterFont.Name = "Tahoma"
        .FooterFont.size = 8
        
        .Header = "Comissões de vendas do mês de " & cmbdoMes & " de " & cmbdoAno
        
    End With

GridLista.PrintPreview 100


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub optreceber_Click()
On Error GoTo tratar_erro

If optreceber.Value = True Then
    StrSql = "Select DISTINCT AnoVencer As Ano from Vendas_relatorios_recebimentos_vendedor_Mes order by AnoVencer Desc"
    ProcCarregaComboAno
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optrecebida_Click()
On Error GoTo tratar_erro

If optrecebida.Value = True Then
    StrSql = "Select DISTINCT Ano from Vendas_relatorios_recebimentos_vendedor_Mes order by Ano Desc"
    ProcCarregaComboAno
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcFiltrarCompleto
    Case 2: ProcImprimir
    Case 3: frmVendas_Comissoes_Metas_Exportar.Show 1
    'Case 4: ProcAjuda
    Case 6: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcAjustaGridVendedor()
On Error GoTo tratar_erro

With GridLista
    .rows = 1
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
    .Cell(0, 1).Text = "Vendedor"

    .Cell(0, 2).Text = "% Comissão"
    .Cell(0, 3).Text = "Meta"
    .Cell(0, 4).Text = "Valor vendido"
    .Cell(0, 5).Text = "% Meta"
    .Cell(0, 6).Text = IIf(optreceber.Value = False, "Valor recebido", "Valor receber")
    .Cell(0, 7).Text = "Valor Comissão"

    .Column(0).Width = 18
    .Column(1).Width = 250
    .Column(2).Width = 80
    .Column(3).Width = 100
    .Column(4).Width = 100
    .Column(5).Width = 100
    .Column(6).Width = 100
    .Column(7).Width = 80
    
    .Column(1).CellType = cellTextBox
    .Column(1).Alignment = cellLeftCenter
    .Column(1).Locked = True
    
    .Column(2).CellType = cellTextBox
    .Column(2).Alignment = cellCenterCenter
    
    .Column(3).CellType = cellTextBox
    .Column(3).Alignment = cellCenterCenter
    
    .Column(4).CellType = cellTextBox
    .Column(4).Alignment = cellRightCenter
      
    .Column(5).CellType = cellTextBox
    .Column(5).Alignment = cellCenterCenter
    
    .Column(6).CellType = cellTextBox
    .Column(6).Alignment = cellRightCenter

    .Column(7).CellType = cellTextBox
    .Column(7).Alignment = cellRightCenter
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaListaVendedores()
On Error GoTo tratar_erro
Dim L As Long
Total = 0

With GridLista
    
 L = 1
.rows = 1
.Cols = 9

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
  If TBAbrir.EOF = False Then
    Contador2 = 1 'TBAbrir.RecordCount
 
 TotalPagar = 0
 TotalVendido = 0
 TotalMeta = 0
 TotalRecebido = 0
 
        Do While TBAbrir.EOF = False
        '==================================================
        ' Busca comissao do vendedor de acorodo com os recebimentos
        '=======================================================
        StrSql = "Select max(Comissao) as Comissao from Vendas_Vendedores_Comissoes Where IDVendedor = '" & TBAbrir!IDvendedor & "' and De <= " & Replace(TBAbrir!valorVendido, ",", ".")
        'Debug.print StrSql
        Set TBVendas = CreateObject("adodb.recordset")
        TBVendas.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
         If TBVendas.EOF = False Then
         Comissao = Format(IIf(IsNull(TBVendas!Comissao), 0, TBVendas!Comissao), "###,##0.00")

         If IsNull(TBAbrir!ValorRecebido) = False And IsNull(TBVendas!Comissao) = False Then
            ValorPagar = (TBAbrir!ValorRecebido * TBVendas!Comissao) / 100
         Else
            ValorPagar = 0
         End If
         
         End If

       '===============================================================
        
         .AddItem UCase(TBAbrir!vendedor)
         .Cell(Contador2, 2).Text = Format(IIf(IsNull(TBVendas!Comissao), 0, TBVendas!Comissao), "###,##0.00") & " %"
         .Cell(Contador2, 3).Text = "R$" & Format(TBAbrir!Meta, "###,##0.00")
         .Cell(Contador2, 4).Text = "R$" & Format(TBAbrir!valorVendido, "###,##0.00")
         .Cell(Contador2, 5).Text = Format(IIf(IsNull(TBAbrir!TotalAtingido), 0, TBAbrir!TotalAtingido), "###,##0.00") & " %"
         .Cell(Contador2, 6).Text = "R$" & Format(IIf(IsNull(TBAbrir!ValorRecebido), 0, TBAbrir!ValorRecebido), "###,##0.00")
         .Cell(Contador2, 7).Text = "R$" & Format(ValorPagar, "###,##0.00")
         
         TotalMeta = TotalMeta + TBAbrir!Meta
         TotalVendido = TotalVendido + TBAbrir!valorVendido
         TotalPagar = TotalPagar + ValorPagar
         TotalRecebido = TotalRecebido + IIf(IsNull(TBAbrir!ValorRecebido), 0, TBAbrir!ValorRecebido)
         
         Contador2 = Contador2 + 1
         TBVendas.Close
         TBAbrir.MoveNext
        Loop

          pMeta = (TotalVendido / TotalMeta) * 100
         .AddItem ""
         .Cell(Contador2, 3).Text = "R$" & Format(TotalMeta, "###,##0.00")
         .Cell(Contador2, 4).Text = "R$" & Format(TotalVendido, "###,##0.00")
         .Cell(Contador2, 5).Text = Format(pMeta, "###,##0.00") & " %"
         .Cell(Contador2, 6).Text = "R$" & Format(TotalRecebido, "###,##0.00")
         .Cell(Contador2, 7).Text = "R$" & Format(TotalPagar, "###,##0.00")
End If
End With

TBAbrir.Close

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

