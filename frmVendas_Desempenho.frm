VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{50D37AD9-8D3C-43DD-ADD5-7C957C951843}#1.9#0"; "FlexCell.ocx"
Begin VB.Form frmVendas_Desempenho 
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
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tipo relatório"
      Height          =   1125
      Left            =   30
      TabIndex        =   19
      Top             =   990
      Width           =   1965
      Begin VB.ComboBox cmbTipo 
         Height          =   315
         ItemData        =   "frmVendas_Desempenho.frx":0000
         Left            =   240
         List            =   "frmVendas_Desempenho.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   20
         ToolTipText     =   "Escolha o tipo do relatório para filtrar"
         Top             =   540
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Filtro"
      Height          =   1125
      Left            =   6240
      TabIndex        =   4
      Top             =   990
      Width           =   9135
      Begin VB.ComboBox cmbOpcao 
         Height          =   315
         ItemData        =   "frmVendas_Desempenho.frx":0028
         Left            =   1320
         List            =   "frmVendas_Desempenho.frx":002A
         Style           =   2  'Dropdown List
         TabIndex        =   6
         ToolTipText     =   "Escolha um valor para filtrar"
         Top             =   570
         Width           =   4425
      End
      Begin VB.ComboBox cmbFiltrarPor 
         Height          =   315
         ItemData        =   "frmVendas_Desempenho.frx":002C
         Left            =   180
         List            =   "frmVendas_Desempenho.frx":0039
         Style           =   2  'Dropdown List
         TabIndex        =   5
         ToolTipText     =   "Escolha uma opção pra filtrar"
         Top             =   570
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Opção de filtro"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   3
         Left            =   3015
         TabIndex        =   8
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
         Left            =   390
         TabIndex        =   7
         Top             =   360
         Width           =   705
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Período por"
      Height          =   1125
      Left            =   2010
      TabIndex        =   1
      Top             =   990
      Width           =   1395
      Begin DrawSuite2022.USOptionButton optPeriodo 
         Height          =   315
         Left            =   150
         TabIndex        =   2
         ToolTipText     =   "Filtrar por período"
         Top             =   750
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
         Caption         =   "Data"
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
      Begin DrawSuite2022.USOptionButton optMesAno 
         CausesValidation=   0   'False
         Height          =   315
         Left            =   150
         TabIndex        =   3
         ToolTipText     =   "Filtrar por ano e mês"
         Top             =   480
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
      Begin DrawSuite2022.USOptionButton optAno 
         CausesValidation=   0   'False
         Height          =   315
         Left            =   150
         TabIndex        =   21
         ToolTipText     =   "Filtrar por ano"
         Top             =   210
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         Caption         =   "Ano"
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
         Value           =   -1  'True
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
      DefaultFontSize =   5.25
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
   Begin VB.Frame FramePeriodo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Filtrar do período"
      Height          =   1125
      Left            =   3420
      TabIndex        =   9
      Top             =   990
      Visible         =   0   'False
      Width           =   2805
      Begin MSComCtl2.DTPicker msk_ate 
         Height          =   315
         Left            =   1440
         TabIndex        =   10
         ToolTipText     =   "Data final."
         Top             =   570
         Width           =   1275
         _ExtentX        =   2249
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
         Format          =   127467521
         CurrentDate     =   39057
      End
      Begin MSComCtl2.DTPicker msk_de 
         Height          =   315
         Left            =   150
         TabIndex        =   11
         ToolTipText     =   "Data inicio."
         Top             =   570
         Width           =   1275
         _ExtentX        =   2249
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
         Format          =   127467521
         CurrentDate     =   39057
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Até"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   2
         Left            =   1890
         TabIndex        =   13
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "De"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   4
         Left            =   570
         TabIndex        =   12
         Top             =   360
         Width           =   195
      End
   End
   Begin VB.Frame FrameAno 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Filtrar do ano"
      Height          =   1125
      Left            =   3420
      TabIndex        =   22
      Top             =   990
      Width           =   2805
      Begin VB.ComboBox cmbAno 
         Height          =   315
         ItemData        =   "frmVendas_Desempenho.frx":0059
         Left            =   270
         List            =   "frmVendas_Desempenho.frx":005B
         Style           =   2  'Dropdown List
         TabIndex        =   23
         ToolTipText     =   "Escolha o ano para filtrar"
         Top             =   570
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Do ano"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   5
         Left            =   480
         TabIndex        =   24
         Top             =   360
         Width           =   510
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   30
      TabIndex        =   25
      Top             =   0
      Width           =   15315
      _ExtentX        =   27014
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
      ButtonKey3      =   "6"
      ButtonAlignment3=   2
      BeginProperty ButtonFont3 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft3     =   93
      ButtonTop3      =   2
      ButtonWidth3    =   56
      ButtonHeight3   =   24
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
      ButtonLeft4     =   151
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
      ButtonLeft5     =   155
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
      ButtonLeft6     =   193
      ButtonTop6      =   2
      ButtonWidth6    =   26
      ButtonHeight6   =   21
      ButtonUseMaskColor6=   0   'False
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   9990
         Top             =   120
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmVendas_Desempenho.frx":005D
         Count           =   1
      End
   End
   Begin VB.Frame frameAnoMes 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Filtrar do ano e do mês"
      Height          =   1125
      Left            =   3420
      TabIndex        =   14
      Top             =   990
      Visible         =   0   'False
      Width           =   2805
      Begin VB.ComboBox cmbdoAno 
         Height          =   315
         ItemData        =   "frmVendas_Desempenho.frx":3188
         Left            =   270
         List            =   "frmVendas_Desempenho.frx":318A
         Style           =   2  'Dropdown List
         TabIndex        =   16
         ToolTipText     =   "Escolha o ano para filtrar"
         Top             =   570
         Width           =   1005
      End
      Begin VB.ComboBox cmbdoMes 
         Height          =   315
         ItemData        =   "frmVendas_Desempenho.frx":318C
         Left            =   1440
         List            =   "frmVendas_Desempenho.frx":318E
         Style           =   2  'Dropdown List
         TabIndex        =   15
         ToolTipText     =   "Escolha o mês para filtrar"
         Top             =   570
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Do ano"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   18
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
         Left            =   1650
         TabIndex        =   17
         Top             =   360
         Width           =   285
      End
   End
End
Attribute VB_Name = "frmVendas_Desempenho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ProcFiltrarCompleto()
On Error GoTo tratar_erro
Dim Mes As String

'==========================================================
' Se for filtrar por ano
'==========================================================
If optAno.Value = True Then
    Select Case cmbfiltrarpor.Text
        Case "Cliente":
            If cmbOpcao.Text <> "" Then
                If cmbOpcao.Text = "TOP FIVE" Then
                    StrSql = "Select TOP (5) Cliente, Sum(Total) as Total from Vendas_relatorios_historico_clientes Where Ano = '" & cmbAno.Text & "' Group By Cliente,Ano ORDER BY TOTAL DESC"
                Else
                    StrSql = "Select Cliente, Sum(Total) as Total from Vendas_relatorios_historico_clientes Where Ano = '" & cmbAno.Text & "' and Cliente  = '" & cmbOpcao.Text & "' Group By Cliente,Ano ORDER BY TOTAL DESC"
                End If
            Else
                StrSql = "Select Cliente, Sum(Total) as Total from Vendas_relatorios_historico_clientes Where Ano = '" & cmbAno.Text & "' Group By Cliente,Ano ORDER BY TOTAL DESC"
            End If
            ProcCarregaListaClientesResumido
            
        Case "Produto":
            If cmbOpcao.Text <> "" Then
                If cmbOpcao.Text = "TOP FIVE" Then
                    StrSql = "Select TOP (5) Desenho, Descricao_Tecnica, sum(total) as TOTAL from Vendas_relatorios_historico_Produtos Where Ano = '" & cmbAno.Text & "' Group By Desenho,Descricao_tecnica ORDER BY TOTAL DESC"
                Else
                    StrSql = "Select Desenho, Descricao_Tecnica, sum(total) as TOTAL from Vendas_relatorios_historico_Produtos Where Ano = '" & cmbAno.Text & "' and Desenho  = '" & cmbOpcao.Text & "' Group By Desenho,Descricao_tecnica ORDER BY TOTAL DESC"
                End If
            Else
                    StrSql = "Select Desenho, Descricao_Tecnica, sum(total) as TOTAL from Vendas_relatorios_historico_Produtos Where Ano = '" & cmbAno.Text & "' Group By Desenho,Descricao_tecnica ORDER BY TOTAL DESC"
            End If
            'Debug.print StrSql
            
            ProcCarregaListaProdutosResumido
            
        Case "Vendedor":
            If cmbOpcao.Text <> "" Then
                If cmbOpcao.Text = "TOP FIVE" Then
                    StrSql = "Select TOP (5) Vendedor, Sum(Total) as Total from Vendas_relatorios_historico_Vendedor Where Ano = '" & cmbAno.Text & "' GROUP BY Vendedor,Ano ORDER BY TOTAL DESC"
                Else
                    StrSql = "Select Vendedor, Sum(Total) as Total from Vendas_relatorios_historico_Vendedor Where Ano = '" & cmbAno.Text & "' and vendedor  = '" & cmbOpcao.Text & "' GROUP BY vendedor,Ano ORDER BY TOTAL DESC"
                End If
            Else
                StrSql = "Select Vendedor, Sum(Total) as Total from Vendas_relatorios_historico_Vendedor Where Ano = '" & cmbAno.Text & "' GROUP BY vendedor, Ano ORDER BY SUM(TOTAL) DESC"
            End If
            ProcCarregaListaVendedoresResumido
            
    End Select
End If


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
    
    
    Select Case cmbfiltrarpor.Text
        Case "Cliente":
            If cmbOpcao.Text <> "" Then
                If cmbOpcao.Text = "TOP FIVE" Then
                    StrSql = "Select TOP (5) * from Vendas_relatorios_historico_clientes Where Mes = '" & Mes & "' and Ano = '" & cmbdoAno.Text & "' ORDER BY TOTAL DESC"
                Else
                    StrSql = "Select * from Vendas_relatorios_historico_clientes Where Mes = '" & Mes & "' and Ano = '" & cmbdoAno.Text & "' and Cliente  = '" & cmbOpcao.Text & "' ORDER BY TOTAL DESC"
                End If
            Else
                StrSql = "Select * from Vendas_relatorios_historico_clientes Where Mes = '" & Mes & "' and Ano = '" & cmbdoAno.Text & "' ORDER BY TOTAL DESC"
            End If
            ProcCarregaListaClientesResumido
            
        Case "Produto":
            If cmbOpcao.Text <> "" Then
                If cmbOpcao.Text = "TOP FIVE" Then
                    StrSql = "Select TOP (5) * from Vendas_relatorios_historico_Produtos Where Mes = '" & Mes & "' and Ano = '" & cmbdoAno.Text & "' ORDER BY TOTAL DESC"
                Else
                    StrSql = "Select * from Vendas_relatorios_historico_Produtos Where Mes = '" & Mes & "' and Ano = '" & cmbdoAno.Text & "' and Desenho  = '" & cmbOpcao.Text & "' ORDER BY TOTAL DESC"
                End If
            Else
                If Mes <> "" Then
                    StrSql = "Select * from Vendas_relatorios_historico_Produtos Where Mes = '" & Mes & "' and Ano = '" & cmbdoAno.Text & "' ORDER BY TOTAL DESC"
                Else
                    StrSql = "Select * from Vendas_relatorios_historico_Produtos Where Ano = '" & cmbdoAno.Text & "' ORDER BY TOTAL DESC"
                End If
            End If
            ProcCarregaListaProdutosResumido
            
        Case "Vendedor":
            If cmbOpcao.Text <> "" Then
                If cmbOpcao.Text = "TOP FIVE" Then
                    StrSql = "Select TOP (5) * from Vendas_relatorios_historico_Vendedor Where Mes = '" & Mes & "' and Ano = '" & cmbdoAno.Text & "' ORDER BY TOTAL DESC"
                Else
                    StrSql = "Select * from Vendas_relatorios_historico_Vendedor Where Mes = '" & Mes & "' and Ano = '" & cmbdoAno.Text & "' and vendedor  = '" & cmbOpcao.Text & "' ORDER BY TOTAL DESC"
                End If
            Else
                StrSql = "Select * from Vendas_relatorios_historico_Vendedor Where Mes = '" & Mes & "' and Ano = '" & cmbdoAno.Text & "' ORDER BY TOTAL DESC"
            End If
            ProcCarregaListaVendedoresResumido
            
    End Select
End If

If optPeriodo.Value = True Then
'=====================================================
' Se for filtrar por período
'=====================================================
    Select Case cmbfiltrarpor.Text
        Case "Cliente":
        If cmbOpcao.Text <> "" Then
        
            If cmbOpcao.Text = "TOP FIVE" Then
                StrSql = "Select TOP (5) Cliente, Sum(preco_lote) as Total from Vendas_relatorios_historico_Detalhado Where DataVendas >= '" & msk_de & "' and DataVendas <= '" & msk_ate & "' group By Cliente order BY Total desc"
            Else
                StrSql = "Select Cliente, Sum(preco_lote) as Total from Vendas_relatorios_historico_Detalhado Where DataVendas >= '" & msk_de & "' and DataVendas <= '" & msk_ate & "' and Cliente  = '" & cmbOpcao.Text & "' group By Cliente order BY Total desc"
            End If
        
        Else
            StrSql = "Select Cliente, Sum(preco_lote) as Total from Vendas_relatorios_historico_Detalhado Where DataVendas >= '" & msk_de & "' and DataVendas <= '" & msk_ate & "' group By Cliente order BY Total desc"
        End If
        'Debug.print StrSql
        
        ProcCarregaListaClientesResumido
        
        Case "Produto":
        If cmbOpcao.Text <> "" Then
            If cmbOpcao.Text = "TOP FIVE" Then
                StrSql = "Select TOP (5) Desenho,Descricao_tecnica, Sum(preco_lote) as Total from Vendas_relatorios_historico_Detalhado Where DataVendas >= '" & msk_de & "' and DataVendas <= '" & msk_ate & "' Group BY Desenho, Descricao_tecnica order BY Total desc"
            Else
                StrSql = "Select Desenho,Descricao_tecnica, Sum(preco_lote) as Total from Vendas_relatorios_historico_Detalhado Where DataVendas >= '" & msk_de & "' and DataVendas <= '" & msk_ate & "' and Desenho  = '" & cmbOpcao.Text & "' Group BY Desenho, Descricao_tecnica order BY Total desc"
            End If
            
        Else
            StrSql = "Select Desenho,Descricao_tecnica, Sum(preco_lote) as Total from Vendas_relatorios_historico_Detalhado Where DataVendas >= '" & msk_de & "' and DataVendas <= '" & msk_ate & "' group By Desenho, Descricao_tecnica order BY Total desc"
        End If
        ProcCarregaListaProdutosResumido
        
        Case "Vendedor":
        If cmbOpcao.Text <> "" Then
            If cmbOpcao.Text = "TOP FIVE" Then
                StrSql = "Select TOP (5) Vend_Ext as Vendedor, Sum(preco_lote) as Total from Vendas_relatorios_historico_Detalhado Where DataVendas >= '" & msk_de & "' and DataVendas <= '" & msk_ate & "' group By Vend_Ext order BY Total desc"
            Else
                StrSql = "Select Vend_Ext as Vendedor, Sum(preco_lote) as Total from Vendas_relatorios_historico_Detalhado Where DataVendas >= '" & msk_de & "' and DataVendas <= '" & msk_ate & "' and Vend_ext  = '" & cmbOpcao.Text & "' group By Vend_Ext order BY Total desc"
            End If
        Else
            StrSql = "Select Vend_Ext as Vendedor, Sum(preco_lote) as Total from Vendas_relatorios_historico_Detalhado Where DataVendas >= '" & msk_de & "' and DataVendas <= '" & msk_ate & "' group By Vend_Ext order BY Total desc"
        End If
        
        ProcCarregaListaVendedoresResumido
        
    End Select
End If

'Debug.print StrSql


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrarUltimasVendas()
On Error GoTo tratar_erro
Dim Mes As String

'==========================================================
' Se for filtrar por ano
'==========================================================
If optAno.Value = True Then
    Select Case cmbfiltrarpor.Text
        Case "Cliente":
            If cmbOpcao.Text <> "" Then
                If cmbOpcao.Text = "TOP FIVE" Then
                    StrSql = "select Top (5) cliente,Vend_ext as Vendedor,Endereco, Cidade, Bairro, Numero, max(datavendas) as ultimavenda from Vendas_relatorios_historico_detalhado Where Ano = '" & cmbAno.Text & "' group by Cliente, Vend_ext, Endereco, Cidade, Bairro, Numero order by max(datavendas)"
                Else
                    StrSql = "select cliente,Vend_ext as Vendedor,Endereco, Cidade, Bairro,Numero, max(datavendas) as ultimavenda from Vendas_relatorios_historico_detalhado Where Ano = '" & cmbAno.Text & "' and Cliente  = '" & cmbOpcao.Text & "' group by Cliente, Vend_ext, Endereco, Cidade, Bairro, Numero order by max(datavendas)"
                End If
            Else
                StrSql = "select cliente,Vend_ext as Vendedor,Endereco, Cidade,Bairro,Numero, max(datavendas) as ultimavenda from Vendas_relatorios_historico_detalhado Where Ano = '" & cmbAno.Text & "' group by Cliente, Vend_ext, Endereco, Cidade, Bairro, Numero order by max(datavendas)"
            End If
            
            ProcCarregaListaClientesMenos
            
        Case "Produto":
            If cmbOpcao.Text <> "" Then
                If cmbOpcao.Text = "TOP FIVE" Then
                    StrSql = "select Top (5) Desenho,Descricao_tecnica, max(datavendas) as ultimavenda from Vendas_relatorios_historico_detalhado Where Ano = '" & cmbAno.Text & "' group by Desenho,Descricao_tecnica order by max(datavendas)"
                Else
                    StrSql = "select Desenho,Descricao_tecnica, max(datavendas) as ultimavenda from Vendas_relatorios_historico_detalhado Where Ano = '" & cmbAno.Text & "' and Desenho  = '" & cmbOpcao.Text & "' group by Desenho,Descricao_tecnica order by max(datavendas)"
                End If
            Else
                StrSql = "select Desenho,Descricao_tecnica, max(datavendas) as ultimavenda from Vendas_relatorios_historico_detalhado Where Ano = '" & cmbAno.Text & "' group by Desenho,Descricao_tecnica order by max(datavendas)"
            End If
            ProcCarregaListaProdutosMenos
            
        Case "Vendedor":
            If cmbOpcao.Text <> "" Then
                If cmbOpcao.Text = "TOP FIVE" Then
                    StrSql = "Select TOP (5) Vend_ext as Vendedor, Cliente, Endereco, Cidade, Bairro, Numero, max(datavendas) as ultimavenda from Vendas_relatorios_historico_detalhado Where Ano = '" & cmbAno.Text & "' GROUP BY Vend_ext,Cliente, Endereco,Cidade,Bairro,Numero,Ano ORDER BY max(datavendas)"
                Else
                    StrSql = "Select Vend_ext as Vendedor,Cliente, Endereco, Cidade, Bairro, Numero, max(datavendas) as ultimavenda from Vendas_relatorios_historico_detalhado Where Ano = '" & cmbAno.Text & "' and vend_ext  = '" & cmbOpcao.Text & "' GROUP BY Vend_ext,Cliente, Endereco,Cidade,Bairro,Numero,Ano ORDER BY max(datavendas)"
                End If
            Else
                    StrSql = "Select Vend_ext as Vendedor,Cliente, Endereco, Cidade, Bairro, Numero, max(datavendas) as ultimavenda from Vendas_relatorios_historico_detalhado Where Ano = '" & cmbAno.Text & "' GROUP BY Vend_ext, Cliente,Endereco, Cidade, Bairro, Numero, Ano ORDER BY max(datavendas)"
            End If
            ProcCarregaListaVendedoresMenos
            
    End Select
End If


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
    
    
    Select Case cmbfiltrarpor.Text
        Case "Cliente":
            If cmbOpcao.Text <> "" Then
                If cmbOpcao.Text = "TOP FIVE" Then
                    StrSql = "select Top (5) Cliente,Vend_ext as Vendedor, max(datavendas) as ultimavenda from Vendas_relatorios_historico_detalhado group by Cliente,Vend_ext Having Month(max(datavendas)) = '" & Mes & "' and Year(max(datavendas)) = '" & cmbdoAno.Text & "' order by max(datavendas)"
                Else
                    StrSql = "select Cliente,Vend_ext as Vendedor, max(datavendas) as ultimavenda from Vendas_relatorios_historico_detalhado Where Cliente  = '" & cmbOpcao.Text & "' group by Cliente,Vend_ext Having Month(max(datavendas)) = '" & Mes & "' and Year(max(datavendas)) = '" & cmbdoAno.Text & "' order by max(datavendas)"
                End If
            Else
                StrSql = "select cliente,Vend_ext as Vendedor,Endereco, Cidade,Bairro,Numero, max(datavendas) as ultimavenda from Vendas_relatorios_historico_detalhado  group by Cliente,Vend_ext, Endereco, Cidade, Bairro, Numero Having Month(max(datavendas)) = '" & Mes & "' and Year(max(datavendas)) = '" & cmbdoAno.Text & "' order by max(datavendas)"
                'StrSql = "select cliente,Vend_ext as Vendedor,Endereco, Cidade,Bairro,Numero, max(datavendas) as ultimavenda from Vendas_relatorios_historico_detalhado Where Ano = '" & cmbAno.Text & "' group by Cliente, Vend_ext, Endereco, Cidade, Bairro, Numero order by max(datavendas)"
            End If
            ProcCarregaListaClientesMenos
            
        Case "Produto":
            If cmbOpcao.Text <> "" Then
                If cmbOpcao.Text = "TOP FIVE" Then
                    StrSql = "select Top (5) Desenho,Descricao_tecnica, max(datavendas) as ultimavenda from Vendas_relatorios_historico_detalhado group by Desenho,Descricao_tecnica Having Month(max(datavendas)) = '" & Mes & "' and Year(max(datavendas)) = '" & cmbdoAno.Text & "' order by max(datavendas)"
                Else
                    StrSql = "select Desenho,Descricao_tecnica, max(datavendas) as ultimavenda from Vendas_relatorios_historico_detalhado Where Desenho  = '" & cmbOpcao.Text & "' group by Desenho,Descricao_tecnica Having Month(max(datavendas)) = '" & Mes & "' and Year(max(datavendas)) = '" & cmbdoAno.Text & "' order by max(datavendas)"
                End If
            Else
                StrSql = "select Desenho,Descricao_tecnica, max(datavendas) as ultimavenda from Vendas_relatorios_historico_detalhado  group by Desenho,Descricao_tecnica Having Month(max(datavendas)) = '" & Mes & "' and Year(max(datavendas)) = '" & cmbdoAno.Text & "' order by max(datavendas)"
            End If
            ProcCarregaListaProdutosMenos

            
        Case "Vendedor":
            If cmbOpcao.Text <> "" Then
                If cmbOpcao.Text = "TOP FIVE" Then
                    StrSql = "select Top (5) vend_ext as Vendedor,Cliente,Endereco, Cidade, Bairro, Numero, max(datavendas) as ultimavenda from Vendas_relatorios_historico_detalhado group by Vend_ext,Cliente,Endereco, Cidade, Bairro, Numero Having Month(max(datavendas)) = '" & Mes & "' and Year(max(datavendas)) = '" & cmbdoAno.Text & "' order by max(datavendas)"
                Else
                    StrSql = "select vend_ext as Vendedor,Cliente,Endereco, Cidade, Bairro, Numero, max(datavendas) as ultimavenda from Vendas_relatorios_historico_detalhado Where Vend_ext  = '" & cmbOpcao.Text & "' group by Vend_ext,Cliente,Endereco, Cidade, Bairro, Numero Having Month(max(datavendas)) = '" & Mes & "' and Year(max(datavendas)) = '" & cmbdoAno.Text & "' order by max(datavendas)"
                End If
            Else
                StrSql = "select vend_ext as Vendedor,Cliente,Endereco, Cidade, Bairro, Numero, max(datavendas) as ultimavenda from Vendas_relatorios_historico_detalhado  group by Vend_ext,Cliente,Endereco, Cidade, Bairro, Numero Having Month(max(datavendas)) = '" & Mes & "' and Year(max(datavendas)) = '" & cmbdoAno.Text & "' order by max(datavendas)"
            End If
            ProcCarregaListaVendedoresMenos
            'Debug.print StrSql
            
    End Select
End If

If optPeriodo.Value = True Then
'=====================================================
' Se for filtrar por período
'=====================================================
    Select Case cmbfiltrarpor.Text
        Case "Cliente":
        If cmbOpcao.Text <> "" Then
        
            If cmbOpcao.Text = "TOP FIVE" Then
                StrSql = "Select TOP (5) Cliente,Vend_ext as Vendedor,Endereco,Cidade,Bairro,Numero, max(datavendas) as ultimavenda from Vendas_relatorios_historico_Detalhado group By Cliente,Vend_ext, Endereco, Cidade, Bairro, Numero Having max(datavendas) >= '" & msk_de.Value & "' and Max(datavendas) <= '" & msk_ate.Value & "' order by max(datavendas)"
            Else
                StrSql = "Select Cliente,Vend_ext as Vendedor,Endereco,Cidade,Bairro,Numero, max(datavendas) as ultimavenda from Vendas_relatorios_historico_Detalhado Where Cliente  = '" & cmbOpcao.Text & "' group By Cliente,Vend_ext, Endereco, Cidade, Bairro, Numero Having max(datavendas) >= '" & msk_de.Value & "' and Max(datavendas) <= '" & msk_ate.Value & "' order by max(datavendas)"
            End If
        
        Else
            StrSql = "Select Cliente,vend_ext as Vendedor, Endereco, Cidade, Bairro, Numero, max(datavendas) as ultimavenda from Vendas_relatorios_historico_Detalhado group By Cliente,Vend_ext, Endereco, Cidade,Bairro,Numero Having max(datavendas) >= '" & msk_de.Value & "' and Max(datavendas) <= '" & msk_ate.Value & "' order by max(datavendas)"
        End If
        'Debug.print StrSql
        
        ProcCarregaListaClientesMenos
        
        Case "Produto":
        If cmbOpcao.Text <> "" Then
            If cmbOpcao.Text = "TOP FIVE" Then
                StrSql = "Select TOP (5) Desenho,Descricao_tecnica, max(datavendas) as ultimavenda from Vendas_relatorios_historico_Detalhado Group BY Desenho, Descricao_tecnica Having max(datavendas) >= '" & msk_de.Value & "' and Max(datavendas) <= '" & msk_ate.Value & "' order by max(datavendas)"
            Else
                StrSql = "Select Desenho,Descricao_tecnica, max(datavendas) as ultimavenda from Vendas_relatorios_historico_Detalhado Where Desenho  = '" & cmbOpcao.Text & "' Group BY Desenho, Descricao_tecnica Having max(datavendas) >= '" & msk_de.Value & "' and Max(datavendas) <= '" & msk_ate.Value & "' order by max(datavendas)"
            End If
            
        Else
            StrSql = "Select Desenho,Descricao_tecnica, max(datavendas) as ultimavenda from Vendas_relatorios_historico_Detalhado group By Desenho, Descricao_tecnica Having max(datavendas) >= '" & msk_de.Value & "' and Max(datavendas) <= '" & msk_ate.Value & "' order by max(datavendas)"
        End If
        'Debug.print StrSql
        
        ProcCarregaListaProdutosMenos
        
        Case "Vendedor":
        If cmbOpcao.Text <> "" Then
            If cmbOpcao.Text = "TOP FIVE" Then
                StrSql = "Select TOP (5) Vend_Ext as Vendedor,Cliente,Endereco, Cidade, Bairro, Numero, max(datavendas) as ultimavenda from Vendas_relatorios_historico_Detalhado Where DataVendas >= '" & msk_de & "' and DataVendas <= '" & msk_ate & "' group By Vend_Ext, Endereco, Cidade, Bairro, Numero order by max(datavendas)"
            Else
                StrSql = "Select Cliente, Vend_Ext as Vendedor, Endereco, Cidade, Bairro, Numero, max(datavendas) as ultimavenda from Vendas_relatorios_historico_Detalhado Where DataVendas >= '" & msk_de & "' and DataVendas <= '" & msk_ate & "' and Vend_ext  = '" & cmbOpcao.Text & "' group By Vend_Ext,Cliente, Endereco, Cidade, Bairro, Numero order by max(datavendas)"
            End If
        Else
            StrSql = "Select Cliente, Vend_Ext as Vendedor, Endereco, Cidade, Bairro, Numero, max(datavendas) as ultimavenda from Vendas_relatorios_historico_Detalhado Where DataVendas >= '" & msk_de & "' and DataVendas <= '" & msk_ate & "' group By Vend_Ext,Cliente, Endereco, Cidade, Bairro, Numero order by max(datavendas)"
        End If
        
        ProcCarregaListaVendedoresMenos
        
    End Select
End If

'Debug.print StrSql


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

If cmbTipo.Text = "Completo" Then
    ProcFiltrarCompleto
Else
    ProcFiltrarUltimasVendas
End If

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
    
    
    Select Case cmbfiltrarpor.Text
        Case "Cliente":  StrSql = "Select DISTINCT Cliente from Vendas_relatorios_historico_clientes Where Mes = '" & Mes & "' and Ano = '" & cmbdoAno.Text & "'"
        Case "Produto":  StrSql = "Select DISTINCT Desenho from Vendas_relatorios_historico_Produtos Where Mes = '" & Mes & "' and Ano = '" & cmbdoAno.Text & "'"
        Case "Vendedor": StrSql = "Select DISTINCT Vendedor from Vendas_relatorios_historico_Vendedor Where Mes = '" & Mes & "' and Ano = '" & cmbdoAno.Text & "'"
    End Select
End If

If optPeriodo.Value = True Then
'=====================================================
' Se for filtrar por período
'=====================================================
    Select Case cmbfiltrarpor.Text
        Case "Cliente":  StrSql = "Select DISTINCT Cliente from Vendas_relatorios_historico_Detalhado Where DataVendas >= '" & msk_de & "' and DataVendas <= '" & msk_ate & "'"
        Case "Produto":  StrSql = "Select DISTINCT Desenho from Vendas_relatorios_historico_Detalhado Where DataVendas >= '" & msk_de & "' and DataVendas <= '" & msk_ate & "'"
        Case "Vendedor": StrSql = "Select DISTINCT Vend_ext as Vendedor from Vendas_relatorios_historico_Detalhado Where DataVendas >= '" & msk_de & "' and DataVendas <= '" & msk_ate & "'"
    End Select
End If

'=====================================================
' Se for filtrar por ano
'=====================================================
If optAno.Value = True Then
    Select Case cmbfiltrarpor.Text
        Case "Cliente":  StrSql = "Select DISTINCT Cliente from Vendas_relatorios_historico_clientes Where  Ano = '" & cmbAno.Text & "'"
        Case "Produto":  StrSql = "Select DISTINCT Desenho from Vendas_relatorios_historico_Produtos Where Ano = '" & cmbAno.Text & "'"
        Case "Vendedor": StrSql = "Select DISTINCT Vendedor from Vendas_relatorios_historico_Vendedor Where Ano = '" & cmbAno.Text & "'"
    End Select
End If

'Debug.print StrSql

If cmbTipo.Text = "Completo" Then
    If cmbfiltrarpor.Text = "Vendedor" Then
    ProcAjustaGridVendedor
    End If
    
    If cmbfiltrarpor.Text = "Produto" Then
    ProcAjustaGridProduto
    End If
    
    If cmbfiltrarpor.Text = "Cliente" Then
    ProcAjustaGridCliente
    End If
Else
    If cmbfiltrarpor.Text = "Vendedor" Then
     ProcAjustaGridVendedorMenos
    End If
    
    If cmbfiltrarpor.Text = "Produto" Then
    ProcAjustaGridProdutoMenos
    End If
    
    If cmbfiltrarpor.Text = "Cliente" Then
    ProcAjustaGridClienteMenos
    End If

End If
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
    cmbOpcao.AddItem "TOP FIVE"
    Do While TBAbrir.EOF = False
        If cmbfiltrarpor.Text = "Cliente" Then Texto = TBAbrir!Cliente
        If cmbfiltrarpor.Text = "Produto" Then Texto = TBAbrir!Desenho
        If cmbfiltrarpor.Text = "Vendedor" Then Texto = IIf(IsNull(TBAbrir!vendedor), "", TBAbrir!vendedor)
        
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
cmbAno.Clear

StrSql = "Select DISTINCT Ano from Vendas_relatorios_historico_detalhado order by Ano Desc"

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
  If TBAbrir.EOF = False Then
    Do While TBAbrir.EOF = False
        
        With cmbdoAno
            .AddItem TBAbrir!Ano
        End With
        
        With cmbAno
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

Private Sub cmbTipo_Click()
On Error GoTo tratar_erro

If cmbTipo.Text = "Completo" Then
    If cmbfiltrarpor.Text = "Vendedor" Then
    ProcAjustaGridVendedor
    End If
    
    If cmbfiltrarpor.Text = "Produto" Then
    ProcAjustaGridProduto
    End If
    
    If cmbfiltrarpor.Text = "Cliente" Then
    ProcAjustaGridCliente
    End If
Else
    If cmbfiltrarpor.Text = "Vendedor" Then
     ProcAjustaGridVendedorMenos
    End If
    
    If cmbfiltrarpor.Text = "Produto" Then
    ProcAjustaGridProdutoMenos
    End If
    
    If cmbfiltrarpor.Text = "Cliente" Then
    ProcAjustaGridClienteMenos
    End If

End If

'procCarregaComboOpcao

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro
ProcCarregaToolBar1 Me, 16200, 6, True

ProcCarregaComboAno
procCarregaComboMes
optAno.Value = True
msk_de.Value = Date
msk_ate.Value = Date

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optAno_Click()
On Error GoTo tratar_erro

FramePeriodo.Visible = optPeriodo.Value
frameAnoMes.Visible = optMesAno.Value
FrameAno.Visible = optAno.Value

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optMesAno_Click()
On Error GoTo tratar_erro

FramePeriodo.Visible = optPeriodo.Value
frameAnoMes.Visible = optMesAno.Value
FrameAno.Visible = optAno.Value

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optPeriodo_Click()
On Error GoTo tratar_erro

FramePeriodo.Visible = optPeriodo.Value
frameAnoMes.Visible = optMesAno.Value
FrameAno.Visible = optAno.Value

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
    .Cols = 4
    
    .DrawMode = cellOwnerDraw
    
    .Appearance = Flat
    .ScrollBarStyle = Flat
    .FixedRowColStyle = Flat
    .Cell(0, 1).Text = "Item"

    .Cell(0, 2).Text = "Vendedor"

    .Cell(0, 3).Text = "Valor total"

    .Column(0).Width = 8
    .Column(1).Width = 30
    .Column(2).Width = 570
    .Column(3).Width = 100
          
    .Column(1).CellType = cellTextBox
    .Column(1).Alignment = cellCenterCenter
    .Column(1).Locked = True
    
    .Column(2).CellType = cellTextBox
    .Column(2).Alignment = cellLeftCenter
    
    .Column(3).CellType = cellTextBox
    .Column(3).Alignment = cellRightCenter
    
  
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcAjustaGridVendedorMenos()
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

    .Cell(0, 1).Text = "Item"



    .Cell(0, 2).Text = "Vendedor"
    .Cell(0, 3).Text = "Cliente"
    .Cell(0, 4).Text = "Endereço"
    .Cell(0, 5).Text = "Cidade"
    .Cell(0, 6).Text = "Bairro"
    .Cell(0, 7).Text = "Nº"
    
    .Cell(0, 8).Text = "Ultima venda"

    .Column(0).Width = 8
    .Column(1).Width = 30
    .Column(2).Width = 150
    .Column(3).Width = 180
    .Column(4).Width = 200
    .Column(5).Width = 100
    .Column(6).Width = 180
    .Column(7).Width = 40
    .Column(8).Width = 100
    
    .Column(1).CellType = cellTextBox
    .Column(1).Alignment = cellCenterCenter
    .Column(1).Locked = True
    
    .Column(2).CellType = cellTextBox
    .Column(2).Alignment = cellCenterCenter
    
    .Column(3).CellType = cellTextBox
    .Column(3).Alignment = cellLeftCenter
    
    .Column(4).CellType = cellTextBox
    .Column(4).Alignment = cellLeftCenter
  
    .Column(5).CellType = cellTextBox
    .Column(5).Alignment = cellCenterCenter
  
    .Column(6).CellType = cellTextBox
    .Column(6).Alignment = cellLeftCenter
  
    .Column(7).CellType = cellTextBox
    .Column(7).Alignment = cellCenterCenter
  
    .Column(8).CellType = cellTextBox
    .Column(8).Alignment = cellRightCenter
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcAjustaGridClienteMenos()
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
    .Cell(0, 1).Text = "Item"

    .Cell(0, 2).Text = "Cliente"

    .Cell(0, 3).Text = "Vendedor"
    .Cell(0, 4).Text = "Endereço"
    .Cell(0, 5).Text = "Cidade"
    .Cell(0, 6).Text = "Bairro"
    .Cell(0, 7).Text = "Nº"
    
    .Cell(0, 8).Text = "Ultima venda"

    .Column(0).Width = 8
    .Column(1).Width = 30
    .Column(2).Width = 180
    .Column(3).Width = 150
    .Column(4).Width = 200
    .Column(5).Width = 100
    .Column(6).Width = 180
    .Column(7).Width = 40
    .Column(8).Width = 100
    
    .Column(1).CellType = cellTextBox
    .Column(1).Alignment = cellCenterCenter
    .Column(1).Locked = True
    
    .Column(2).CellType = cellTextBox
    .Column(2).Alignment = cellLeftCenter
    
    .Column(3).CellType = cellTextBox
    .Column(3).Alignment = cellCenterCenter
    
    .Column(4).CellType = cellTextBox
    .Column(4).Alignment = cellLeftCenter
  
    .Column(5).CellType = cellTextBox
    .Column(5).Alignment = cellCenterCenter
  
    .Column(6).CellType = cellTextBox
    .Column(6).Alignment = cellLeftCenter
  
    .Column(7).CellType = cellTextBox
    .Column(7).Alignment = cellCenterCenter
  
    .Column(8).CellType = cellTextBox
    .Column(8).Alignment = cellRightCenter
  
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcAjustaGridCliente()
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
    .Cols = 4
    
    .DrawMode = cellOwnerDraw
    
    .Appearance = Flat
    .ScrollBarStyle = Flat
    .FixedRowColStyle = Flat
    .Cell(0, 1).Text = "Item"

    .Cell(0, 2).Text = "Cliente"

    .Cell(0, 3).Text = "Valor total"

    .Column(0).Width = 8
    .Column(1).Width = 30
    .Column(2).Width = 570
    .Column(3).Width = 100
          
    .Column(1).CellType = cellTextBox
    .Column(1).Alignment = cellCenterCenter
    .Column(1).Locked = True
    
    .Column(2).CellType = cellTextBox
    .Column(2).Alignment = cellLeftCenter
    
    .Column(3).CellType = cellTextBox
    .Column(3).Alignment = cellRightCenter
    
  
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcAjustaGridProduto()
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
    .Cols = 5
    
    .DrawMode = cellOwnerDraw
    
    .Appearance = Flat
    .ScrollBarStyle = Flat
    .FixedRowColStyle = Flat
    .Cell(0, 1).Text = "Item"

    .Cell(0, 2).Text = "Código"

    .Cell(0, 3).Text = "Descrição"

    .Cell(0, 4).Text = "Valor total"

    .Column(0).Width = 8
    .Column(1).Width = 30
    .Column(2).Width = 60
    .Column(3).Width = 510
    .Column(4).Width = 100
   ' .Column(5).Width = 1
          
    .Column(1).CellType = cellTextBox
    .Column(1).Alignment = cellCenterCenter
    .Column(1).Locked = True
    
    .Column(2).CellType = cellTextBox
    .Column(2).Alignment = cellCenterCenter
    
    .Column(3).CellType = cellTextBox
    .Column(3).Alignment = cellLeftCenter
    
    .Column(4).CellType = cellTextBox
    .Column(4).Alignment = cellRightCenter
  '  .Column(5).CellType = cellTextBox
   ' .Column(5).Alignment = cellRightCenter
    
  
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcAjustaGridProdutoMenos()
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
    .Cols = 5
    
    .DrawMode = cellOwnerDraw
    
    .Appearance = Flat
    .ScrollBarStyle = Flat
    .FixedRowColStyle = Flat
    .Cell(0, 1).Text = "Item"

    .Cell(0, 2).Text = "Código"

    .Cell(0, 3).Text = "Descrição"

    .Cell(0, 4).Text = "Ultima venda"

    .Column(0).Width = 8
    .Column(1).Width = 30
    .Column(2).Width = 60
    .Column(3).Width = 510
    .Column(4).Width = 100
   ' .Column(5).Width = 1
          
    .Column(1).CellType = cellTextBox
    .Column(1).Alignment = cellCenterCenter
    .Column(1).Locked = True
    
    .Column(2).CellType = cellTextBox
    .Column(2).Alignment = cellCenterCenter
    
    .Column(3).CellType = cellTextBox
    .Column(3).Alignment = cellLeftCenter
    
    .Column(4).CellType = cellTextBox
    .Column(4).Alignment = cellRightCenter
  '  .Column(5).CellType = cellTextBox
   ' .Column(5).Alignment = cellRightCenter
    
  
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub



Private Sub ProcCarregaListaClientesResumido()
On Error GoTo tratar_erro
Dim L As Long
Total = 0

With GridLista
    
 L = 1
.rows = 1
.Cols = 4

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
  If TBAbrir.EOF = False Then
    Contador2 = 1 'TBAbrir.RecordCount
    
        Do While TBAbrir.EOF = False
         .AddItem Contador2
         .Cell(Contador2, 2).Text = UCase(TBAbrir!Cliente)
         .Cell(Contador2, 3).Text = "R$" & Format(TBAbrir!Total, "###,##0.00")
         Contador2 = Contador2 + 1
         Total = Total + TBAbrir!Total
         TBAbrir.MoveNext
        Loop
  End If

         .AddItem Contador2
         .Cell(Contador2, 3).Text = "R$" & Format(Total, "###,##0.00")
         
End With
TBAbrir.Close

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Private Sub ProcCarregaListaClientesMenos()
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
    
        Do While TBAbrir.EOF = False
         .AddItem Contador2
         .Cell(Contador2, 2).Text = UCase(TBAbrir!Cliente)
         .Cell(Contador2, 3).Text = UCase(TBAbrir!vendedor)
         .Cell(Contador2, 4).Text = UCase(TBAbrir!Endereco)
         .Cell(Contador2, 5).Text = UCase(TBAbrir!Cidade)
         .Cell(Contador2, 6).Text = UCase(TBAbrir!Bairro)
         .Cell(Contador2, 7).Text = TBAbrir!Numero
         .Cell(Contador2, 8).Text = TBAbrir!Ultimavenda
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

Private Sub ProcCarregaListaProdutosMenos()
On Error GoTo tratar_erro
Dim L As Long
Total = 0

With GridLista
    
 L = 1
.rows = 1
.Cols = 6

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
  If TBAbrir.EOF = False Then
    Contador2 = 1 'TBAbrir.RecordCount
    
        Do While TBAbrir.EOF = False
         .AddItem Contador2
         .Cell(Contador2, 2).Text = UCase(TBAbrir!Desenho)
         .Cell(Contador2, 3).Text = UCase(TBAbrir!descricao_tecnica)
         .Cell(Contador2, 4).Text = TBAbrir!Ultimavenda
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

Private Sub ProcCarregaListaProdutosResumido()
On Error GoTo tratar_erro
Dim L As Long
Total = 0

With GridLista
    
 L = 1
.rows = 1
.Cols = 6

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
  If TBAbrir.EOF = False Then
    Contador2 = 1 'TBAbrir.RecordCount
    
        Do While TBAbrir.EOF = False
         .AddItem Contador2
         .Cell(Contador2, 2).Text = UCase(TBAbrir!Desenho)
         .Cell(Contador2, 3).Text = UCase(TBAbrir!descricao_tecnica)
         .Cell(Contador2, 4).Text = "R$" & Format(TBAbrir!Total, "###,##0.00")
         Contador2 = Contador2 + 1
         Total = Total + TBAbrir!Total
         TBAbrir.MoveNext
        Loop
  End If
         .AddItem Contador2
         .Cell(Contador2, 4).Text = "R$" & Format(Total, "###,##0.00")

End With
TBAbrir.Close

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Private Sub ProcCarregaListaVendedoresResumido()
On Error GoTo tratar_erro
Dim L As Long
Total = 0

With GridLista
    
 L = 1
.rows = 1
.Cols = 4

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
  If TBAbrir.EOF = False Then
    Contador2 = 1 'TBAbrir.RecordCount
    
        Do While TBAbrir.EOF = False
         .AddItem Contador2
         .Cell(Contador2, 2).Text = UCase(TBAbrir!vendedor)
         .Cell(Contador2, 3).Text = "R$" & Format(TBAbrir!Total, "###,##0.00")
         Contador2 = Contador2 + 1
         Total = Total + TBAbrir!Total
         TBAbrir.MoveNext
        Loop


         .AddItem Contador2
         .Cell(Contador2, 3).Text = "R$" & Format(Total, "###,##0.00")

  End If
  
End With
TBAbrir.Close

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Private Sub ProcCarregaListaVendedoresMenos()
On Error GoTo tratar_erro
Dim L As Long
Total = 0

With GridLista
    
 L = 1
.rows = 1
.Cols = 9

Set TBAbrir = CreateObject("adodb.recordset")
'Debug.print StrSql

TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
  If TBAbrir.EOF = False Then
    Contador2 = 1 'TBAbrir.RecordCount
    
        Do While TBAbrir.EOF = False
         .AddItem Contador2
         .Cell(Contador2, 2).Text = UCase(IIf(IsNull(TBAbrir!vendedor) = False, TBAbrir!vendedor, ""))
         .Cell(Contador2, 3).Text = UCase(IIf(IsNull(TBAbrir!Cliente) = False, TBAbrir!Cliente, ""))
         .Cell(Contador2, 4).Text = UCase(IIf(IsNull(TBAbrir!Endereco) = False, TBAbrir!Endereco, ""))
         .Cell(Contador2, 5).Text = UCase(IIf(IsNull(TBAbrir!Cidade) = False, TBAbrir!Cidade, ""))
         .Cell(Contador2, 6).Text = UCase(IIf(IsNull(TBAbrir!Bairro) = False, TBAbrir!Bairro, ""))
         .Cell(Contador2, 7).Text = UCase(IIf(IsNull(TBAbrir!Numero) = False, TBAbrir!Numero, ""))
         
         .Cell(Contador2, 8).Text = IIf(IsNull(TBAbrir!Ultimavenda) = False, TBAbrir!Ultimavenda, "")
         
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
        
        If optAno.Value = True Then
            .Header = "Desempenho de vendas por " & cmbfiltrarpor.Text & " no ano de " & cmbAno
        End If
        
        
        If optMesAno.Value = True Then
            .Header = "Desempenho de vendas por " & cmbfiltrarpor.Text & " no mês de " & cmbdoMes.Text & " de " & cmbdoAno
        End If
        
        If optPeriodo.Value = True Then
            .Header = "Desempenho de vendas por " & cmbfiltrarpor.Text & " período de " & msk_de & " de " & msk_ate
        End If
        
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
    End With

GridLista.PrintPreview 100


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcFiltrar
    Case 2: ProcImprimir
    Case 3: frmVendas_Desempenho_Exportar.Show 1
    'Case 4: ProcAjuda
    Case 6: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
