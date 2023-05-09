VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2014.ocx"
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
      Left            =   0
      TabIndex        =   22
      Top             =   60
      Width           =   1965
      Begin VB.ComboBox cmbTipo 
         Height          =   315
         ItemData        =   "frmVendas_Relatorios_Comissoes_Metas.frx":0000
         Left            =   240
         List            =   "frmVendas_Relatorios_Comissoes_Metas.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   23
         ToolTipText     =   "Escolha o ano para filtrar"
         Top             =   540
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Filtro"
      Height          =   1125
      Left            =   6210
      TabIndex        =   4
      Top             =   60
      Width           =   9165
      Begin VB.ComboBox cmbOpcao 
         Height          =   315
         ItemData        =   "frmVendas_Relatorios_Comissoes_Metas.frx":0028
         Left            =   1290
         List            =   "frmVendas_Relatorios_Comissoes_Metas.frx":002A
         Style           =   2  'Dropdown List
         TabIndex        =   6
         ToolTipText     =   "Escolha um valor para filtrar"
         Top             =   570
         Width           =   4425
      End
      Begin VB.ComboBox cmbFiltrarPor 
         Height          =   315
         ItemData        =   "frmVendas_Relatorios_Comissoes_Metas.frx":002C
         Left            =   150
         List            =   "frmVendas_Relatorios_Comissoes_Metas.frx":0039
         Style           =   2  'Dropdown List
         TabIndex        =   5
         ToolTipText     =   "Escolha uma opção pra filtrar"
         Top             =   570
         Width           =   1125
      End
      Begin DrawSuite2014.USButton btnSair 
         Height          =   765
         Left            =   7980
         TabIndex        =   7
         ToolTipText     =   "Fechar formulário"
         Top             =   240
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   1349
         DibPicture      =   "frmVendas_Relatorios_Comissoes_Metas.frx":0059
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
         Left            =   5820
         TabIndex        =   20
         ToolTipText     =   "Filtrar registros de vendas"
         Top             =   240
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   1349
         DibPicture      =   "frmVendas_Relatorios_Comissoes_Metas.frx":2E06
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
      Begin DrawSuite2014.USButton btnImprimir 
         Height          =   765
         Left            =   6900
         TabIndex        =   21
         ToolTipText     =   "Gerar relatório"
         Top             =   240
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   1349
         DibPicture      =   "frmVendas_Relatorios_Comissoes_Metas.frx":6456
         BorderColor     =   1154291
         BorderColorDisabled=   13160660
         BorderColorDown =   16576
         BorderColorOver =   8438015
         Caption         =   "Imprimir"
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
         GradientColor1  =   1154291
         GradientColor2  =   1154291
         GradientColor3  =   1154291
         GradientColor4  =   1154291
         GradientColorDisabled1=   14215660
         GradientColorDisabled2=   14215660
         GradientColorDisabled3=   14215660
         GradientColorDisabled4=   14215660
         GradientColorDown1=   16576
         GradientColorDown2=   16576
         GradientColorDown3=   16576
         GradientColorDown4=   16576
         GradientColorOver1=   8438015
         GradientColorOver2=   8438015
         GradientColorOver3=   8438015
         GradientColorOver4=   8438015
         PicAlign        =   8
         PicSize         =   2
         PicSizeH        =   24
         PicSizeW        =   24
         ShowFocusRect   =   0   'False
         ShowFocusRectDown=   0   'False
         Theme           =   5
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Opção de filtro"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   3
         Left            =   2985
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
         Left            =   360
         TabIndex        =   8
         Top             =   360
         Width           =   705
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Período por"
      Height          =   1125
      Left            =   1980
      TabIndex        =   1
      Top             =   60
      Width           =   1395
      Begin DrawSuite2014.USOptionButton optPeriodo 
         Height          =   315
         Left            =   150
         TabIndex        =   2
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
      Begin DrawSuite2014.USOptionButton optMesAno 
         CausesValidation=   0   'False
         Height          =   315
         Left            =   150
         TabIndex        =   3
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
      Begin DrawSuite2014.USOptionButton optAno 
         CausesValidation=   0   'False
         Height          =   315
         Left            =   150
         TabIndex        =   24
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
      Height          =   8745
      Left            =   0
      TabIndex        =   0
      Top             =   1200
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   15425
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
   Begin VB.Frame frameAnoMes 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Filtrar do ano e do mês"
      Height          =   1125
      Left            =   3390
      TabIndex        =   15
      Top             =   60
      Visible         =   0   'False
      Width           =   2805
      Begin VB.ComboBox cmbdoAno 
         Height          =   315
         ItemData        =   "frmVendas_Relatorios_Comissoes_Metas.frx":80EB
         Left            =   270
         List            =   "frmVendas_Relatorios_Comissoes_Metas.frx":80ED
         Style           =   2  'Dropdown List
         TabIndex        =   17
         ToolTipText     =   "Escolha o ano para filtrar"
         Top             =   570
         Width           =   945
      End
      Begin VB.ComboBox cmbdoMes 
         Height          =   315
         ItemData        =   "frmVendas_Relatorios_Comissoes_Metas.frx":80EF
         Left            =   1440
         List            =   "frmVendas_Relatorios_Comissoes_Metas.frx":80F1
         Style           =   2  'Dropdown List
         TabIndex        =   16
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
         Index           =   0
         Left            =   480
         TabIndex        =   19
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
         TabIndex        =   18
         Top             =   360
         Width           =   285
      End
   End
   Begin VB.Frame FramePeriodo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Filtrar do período"
      Height          =   1125
      Left            =   3390
      TabIndex        =   10
      Top             =   60
      Visible         =   0   'False
      Width           =   2805
      Begin MSComCtl2.DTPicker msk_ate 
         Height          =   315
         Left            =   1440
         TabIndex        =   11
         ToolTipText     =   "Data final."
         Top             =   570
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
         Format          =   848429057
         CurrentDate     =   39057
      End
      Begin MSComCtl2.DTPicker msk_de 
         Height          =   315
         Left            =   150
         TabIndex        =   12
         ToolTipText     =   "Data inicio."
         Top             =   570
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
         Format          =   848429057
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
         TabIndex        =   14
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
         TabIndex        =   13
         Top             =   360
         Width           =   195
      End
   End
   Begin VB.Frame FrameAno 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Filtrar do ano"
      Height          =   1125
      Left            =   3390
      TabIndex        =   25
      Top             =   60
      Width           =   2805
      Begin VB.ComboBox cmbAno 
         Height          =   315
         ItemData        =   "frmVendas_Relatorios_Comissoes_Metas.frx":80F3
         Left            =   270
         List            =   "frmVendas_Relatorios_Comissoes_Metas.frx":80F5
         Style           =   2  'Dropdown List
         TabIndex        =   26
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
         TabIndex        =   27
         Top             =   360
         Width           =   510
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
            Debug.Print StrSql
            
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
        Debug.Print StrSql
        
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

Debug.Print StrSql


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
                    StrSql = "select Top (5) cliente,Vend_ext as Vendedor, max(datavendas) as ultimavenda from Vendas_relatorios_historico_detalhado Where Ano = '" & cmbAno.Text & "' group by Cliente, Vend_ext order by max(datavendas)"
                Else
                    StrSql = "select cliente,Vend_ext as Vendedor, max(datavendas) as ultimavenda from Vendas_relatorios_historico_detalhado Where Ano = '" & cmbAno.Text & "' and Cliente  = '" & cmbOpcao.Text & "' group by Cliente, Vend_ext order by max(datavendas)"
                End If
            Else
                StrSql = "select cliente,Vend_ext as Vendedor, max(datavendas) as ultimavenda from Vendas_relatorios_historico_detalhado Where Ano = '" & cmbAno.Text & "' group by Cliente, Vend_ext order by max(datavendas)"
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
                    StrSql = "Select TOP (5) Vend_ext as Vendedor, max(datavendas) as ultimavenda from Vendas_relatorios_historico_detalhado Where Ano = '" & cmbAno.Text & "' GROUP BY Vend_ext,Ano ORDER BY max(datavendas)"
                Else
                    StrSql = "Select Vend_ext as Vendedor, max(datavendas) as ultimavenda from Vendas_relatorios_historico_detalhado Where Ano = '" & cmbAno.Text & "' and vend_ext  = '" & cmbOpcao.Text & "' GROUP BY Vend_ext,Ano ORDER BY max(datavendas)"
                End If
            Else
                StrSql = "Select Vend_ext as Vendedor, max(datavendas) as ultimavenda from Vendas_relatorios_historico_detalhado Where Ano = '" & cmbAno.Text & "' GROUP BY Vend_ext, Ano ORDER BY max(datavendas)"
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
                StrSql = "select Cliente,Vend_ext as Vendedor, max(datavendas) as ultimavenda from Vendas_relatorios_historico_detalhado  group by Cliente,Vend_ext Having Month(max(datavendas)) = '" & Mes & "' and Year(max(datavendas)) = '" & cmbdoAno.Text & "' order by max(datavendas)"
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
                    StrSql = "select Top (5) vend_ext as Vendedor, max(datavendas) as ultimavenda from Vendas_relatorios_historico_detalhado group by Vend_ext Having Month(max(datavendas)) = '" & Mes & "' and Year(max(datavendas)) = '" & cmbdoAno.Text & "' order by max(datavendas)"
                Else
                    StrSql = "select vend_ext as Vendedor, max(datavendas) as ultimavenda from Vendas_relatorios_historico_detalhado Where Vend_ext  = '" & cmbOpcao.Text & "' group by Vend_ext Having Month(max(datavendas)) = '" & Mes & "' and Year(max(datavendas)) = '" & cmbdoAno.Text & "' order by max(datavendas)"
                End If
            Else
                StrSql = "select vend_ext as Vendedor, max(datavendas) as ultimavenda from Vendas_relatorios_historico_detalhado  group by Vend_ext Having Month(max(datavendas)) = '" & Mes & "' and Year(max(datavendas)) = '" & cmbdoAno.Text & "' order by max(datavendas)"
            End If
            ProcCarregaListaVendedoresMenos
            Debug.Print StrSql
            
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
                StrSql = "Select TOP (5) Cliente,Vend_ext as Vendedor, max(datavendas) as ultimavenda from Vendas_relatorios_historico_Detalhado group By Cliente,Vend_ext Having max(datavendas) >= '" & msk_de.Value & "' and Max(datavendas) <= '" & msk_ate.Value & "' order by max(datavendas)"
            Else
                StrSql = "Select Cliente,Vend_ext as Vendedor, max(datavendas) as ultimavenda from Vendas_relatorios_historico_Detalhado Where Cliente  = '" & cmbOpcao.Text & "' group By Cliente,Vend_ext Having max(datavendas) >= '" & msk_de.Value & "' and Max(datavendas) <= '" & msk_ate.Value & "' order by max(datavendas)"
            End If
        
        Else
            StrSql = "Select Cliente,vend_ext as Vendedor, max(datavendas) as ultimavenda from Vendas_relatorios_historico_Detalhado group By Cliente,Vend_ext Having max(datavendas) >= '" & msk_de.Value & "' and Max(datavendas) <= '" & msk_ate.Value & "' order by max(datavendas)"
        End If
        Debug.Print StrSql
        
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
        Debug.Print StrSql
        
        ProcCarregaListaProdutosMenos
        
        Case "Vendedor":
        If cmbOpcao.Text <> "" Then
            If cmbOpcao.Text = "TOP FIVE" Then
                StrSql = "Select TOP (5) Vend_Ext as Vendedor, max(datavendas) as ultimavenda from Vendas_relatorios_historico_Detalhado Where DataVendas >= '" & msk_de & "' and DataVendas <= '" & msk_ate & "' group By Vend_Ext order by max(datavendas)"
            Else
                StrSql = "Select Vend_Ext as Vendedor, max(datavendas) as ultimavenda from Vendas_relatorios_historico_Detalhado Where DataVendas >= '" & msk_de & "' and DataVendas <= '" & msk_ate & "' and Vend_ext  = '" & cmbOpcao.Text & "' group By Vend_Ext order by max(datavendas)"
            End If
        Else
            StrSql = "Select Vend_Ext as Vendedor, max(datavendas) as ultimavenda from Vendas_relatorios_historico_Detalhado Where DataVendas >= '" & msk_de & "' and DataVendas <= '" & msk_ate & "' group By Vend_Ext order by max(datavendas)"
        End If
        
        ProcCarregaListaVendedoresMenos
        
    End Select
End If

Debug.Print StrSql


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub btnFiltrar_Click()
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

Debug.Print StrSql

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
'Debug.Print StrSql

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
  If TBAbrir.EOF = False Then
    cmbOpcao.AddItem "TOP FIVE"
    Do While TBAbrir.EOF = False
        If cmbfiltrarpor.Text = "Cliente" Then Texto = TBAbrir!Cliente
        If cmbfiltrarpor.Text = "Produto" Then Texto = TBAbrir!Desenho
        If cmbfiltrarpor.Text = "Vendedor" Then Texto = IIf(IsNull(TBAbrir!Vendedor), "", TBAbrir!Vendedor)
        
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

procCarregaComboOpcao

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

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

    .Cell(0, 2).Text = "Valor total"

    .Cell(0, 3).Text = "Ultima venda"

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
    .Cols = 4
    
    .DrawMode = cellOwnerDraw
    
    .Appearance = Flat
    .ScrollBarStyle = Flat
    .FixedRowColStyle = Flat
    .Cell(0, 1).Text = "Item"

    .Cell(0, 2).Text = "Vendedor"

    .Cell(0, 3).Text = "Ultima venda"

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
    .Cols = 5
    
    .DrawMode = cellOwnerDraw
    
    .Appearance = Flat
    .ScrollBarStyle = Flat
    .FixedRowColStyle = Flat
    .Cell(0, 1).Text = "Item"

    .Cell(0, 2).Text = "Cliente"

    .Cell(0, 3).Text = "Vendedor"
    
    .Cell(0, 4).Text = "Ultima venda"

    .Column(0).Width = 8
    .Column(1).Width = 30
    .Column(2).Width = 320
    .Column(3).Width = 250
    .Column(4).Width = 100
    
    .Column(1).CellType = cellTextBox
    .Column(1).Alignment = cellCenterCenter
    .Column(1).Locked = True
    
    .Column(2).CellType = cellTextBox
    .Column(2).Alignment = cellLeftCenter
    
    .Column(3).CellType = cellTextBox
    .Column(3).Alignment = cellCenterCenter
    
    .Column(4).CellType = cellTextBox
    .Column(4).Alignment = cellRightCenter
  
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
.Cols = 5

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
  If TBAbrir.EOF = False Then
    Contador2 = 1 'TBAbrir.RecordCount
    
        Do While TBAbrir.EOF = False
         .AddItem Contador2
         .Cell(Contador2, 2).Text = UCase(TBAbrir!Cliente)
         .Cell(Contador2, 3).Text = UCase(TBAbrir!Vendedor)
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
         .Cell(Contador2, 2).Text = UCase(TBAbrir!Vendedor)
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

Private Sub ProcCarregaListaVendedoresMenos()
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
         .Cell(Contador2, 2).Text = UCase(TBAbrir!Vendedor)
         .Cell(Contador2, 3).Text = TBAbrir!Ultimavenda
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
        
        .Orientation = cellPortrait
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
