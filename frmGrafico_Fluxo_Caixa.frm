VERSION 5.00
Object = "{C215CB9A-0AE1-499F-A101-48B3C370D3DF}#19.3#0"; "Codejock.ChartControl.v19.3.0.ocx"
Begin VB.Form frmGrafico_Fluxo_Caixa 
   Caption         =   "Graficos de fluxo de caixa"
   ClientHeight    =   10035
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15360
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
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   669
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   Begin VB.PictureBox panelOptions 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9855
      Left            =   12360
      ScaleHeight     =   9855
      ScaleWidth      =   2895
      TabIndex        =   1
      Top             =   120
      Width           =   2895
      Begin VB.CommandButton Command3 
         Caption         =   "Atualizar"
         Height          =   315
         Left            =   2430
         TabIndex        =   14
         Top             =   510
         Width           =   345
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Atualizar"
         Height          =   495
         Left            =   300
         TabIndex        =   13
         Top             =   7050
         Width           =   1215
      End
      Begin VB.TextBox txtSaldoInicial 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   240
         TabIndex        =   11
         Text            =   "00,00"
         Top             =   510
         Width           =   2175
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Sair"
         Height          =   495
         Left            =   240
         TabIndex        =   10
         Top             =   9150
         Width           =   1215
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Salva como imagem"
         Height          =   435
         Left            =   240
         TabIndex        =   6
         Top             =   8640
         Width           =   1215
      End
      Begin VB.ComboBox cmbStyle 
         Height          =   315
         ItemData        =   "frmGrafico_Fluxo_Caixa.frx":0000
         Left            =   240
         List            =   "frmGrafico_Fluxo_Caixa.frx":0013
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1530
         Width           =   2535
      End
      Begin VB.ComboBox cmbPalette 
         Height          =   315
         ItemData        =   "frmGrafico_Fluxo_Caixa.frx":007A
         Left            =   240
         List            =   "frmGrafico_Fluxo_Caixa.frx":00DB
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2970
         Width           =   2535
      End
      Begin VB.ComboBox cmbAppearance 
         Height          =   315
         ItemData        =   "frmGrafico_Fluxo_Caixa.frx":0226
         Left            =   240
         List            =   "frmGrafico_Fluxo_Caixa.frx":0233
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   2250
         Width           =   2535
      End
      Begin VB.CommandButton cmdPrintPreview 
         Caption         =   "Visualizar impressão"
         Height          =   435
         Left            =   1560
         TabIndex        =   2
         Top             =   8640
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Saldo atual"
         Height          =   285
         Left            =   1110
         TabIndex        =   12
         Top             =   270
         Width           =   1275
      End
      Begin VB.Label lblStyle 
         Caption         =   "Estilo:"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1290
         Width           =   2175
      End
      Begin VB.Label lblPalette 
         Caption         =   "Cores:"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   2730
         Width           =   2295
      End
      Begin VB.Label lblAppearance 
         Caption         =   "Aparência:"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   2010
         Width           =   2295
      End
   End
   Begin XtremeChartControl.ChartControl ChartControl 
      Height          =   9855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12165
      _Version        =   1245187
      _ExtentX        =   21458
      _ExtentY        =   17383
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmGrafico_Fluxo_Caixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function ShellExecute Lib "shell32.dll" _
            Alias "ShellExecuteA" _
            (ByVal hWnd As Long, _
            ByVal lpOperation As String, _
            ByVal lpFile As String, _
            ByVal lpParameters As String, _
            ByVal lpDirectory As String, _
            ByVal nShowCmd As Long) As Long

Option Explicit

Dim Diagram As ChartDiagram2D
Dim Strip As ChartAxisStrip

Sub AddTitles()
On Error GoTo tratar_erro

ChartControl.Content.Titles.Add "Gráfico de fluxo de caixa"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub AddSeries()
On Error GoTo tratar_erro
    
    lblLabelPosition.Visible = False
    cmbPieLabelPosition.Visible = False
    cmbPyramidLabelPosition.Visible = False
    lbl3dHint.Visible = False
    
    If ChartControl.Content.Series.Count > 0 Then
        ChartControl.Content.Series.DeleteAll
    End If
    
    ChartControl.Content.Titles(0).Text = "Simple Chart Sample"
    ChartControl.Content.Legend.Visible = True
    ChartControl.Content.Legend.HorizontalAlignment = xtpChartLegendFar
    
    Dim Series As ChartSeries
    
    If ChartControl.Content.Series.Count > 0 Then
        Set Series = ChartControl.Content.Series(0)
    Else
        Set Series = ChartControl.Content.Series.Add("Series 1")
    End If
    
    ChartControl.Content.Legend.Visible = True

    Series.Points.Add 0, 0.5
    Series.Points.Add 1, 2
    Series.Points.Add 2, 1
    Series.Points.Add 3, 1.5
    Series.Points.Add 4, 3
    Series.Points.Add 5, 2.5
    Series.Points.Add 6, 1
    Series.Points.Add 7, 0.5
    Series.Points.Add 8, 1.5
    Series.Points.Add 9, 2.5
    Series.Points.Add 10, 0.5
    Series.Points.Add 11, 2
    Series.Points.Add 12, 1
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub AddPieSeries()
On Error GoTo tratar_erro
    
    lblLabelPosition.Visible = True
    cmbPieLabelPosition.Visible = True
    cmbPyramidLabelPosition.Visible = False
    lbl3dHint.Visible = False
    
    If ChartControl.Content.Series.Count > 0 Then
        ChartControl.Content.Series.DeleteAll
    End If
    
    Dim Series As ChartSeries
    
    ChartControl.Content.Titles(0).Text = "Top 10 States by Population"
    ChartControl.Content.Legend.Visible = True
    ChartControl.Content.Legend.HorizontalAlignment = xtpChartLegendNearOutside
    
    Set Series = ChartControl.Content.Series.Add("Top 10 States by Population")
    
    CreateSeriesPoint Series.Points, "California", 11.95
    CreateSeriesPoint Series.Points, "Texas", 7.81
    CreateSeriesPoint Series.Points, "New York", 6.31
    CreateSeriesPoint Series.Points, "Florida", 5.97
    CreateSeriesPoint Series.Points, "Illinois", 4.2
    CreateSeriesPoint Series.Points, "Pennsylvania", 4.06
    CreateSeriesPoint Series.Points, "Ohio", 3.75
    CreateSeriesPoint Series.Points, "Michigan", 3.29
    CreateSeriesPoint Series.Points, "Georgia", 3.12
    CreateSeriesPoint Series.Points, "North Carolina", 3
                
    Series.Points(9).Special = True
    Series.Points(1).Special = True
    Series.Points(7).Special = True
    

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub CreateSeriesPoint(ByVal pPointCollection As ChartSeriesPointCollection, vArg As String, nValue As Double)
On Error GoTo tratar_erro
    
    Dim pPoint As ChartSeriesPoint
    Set pPoint = pPointCollection.Add(vArg, nValue)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Sub AddBubbleSeries()
On Error GoTo tratar_erro
    
    lblLabelPosition.Visible = False
    cmbPieLabelPosition.Visible = False
    cmbPyramidLabelPosition.Visible = False
    lbl3dHint.Visible = False
    
    If ChartControl.Content.Series.Count > 0 Then
        ChartControl.Content.Series.DeleteAll
    End If
    
    ChartControl.Content.Titles.DeleteAll
    ChartControl.Content.Titles.Add "Top 10 States by Population"
    ChartControl.Content.Legend.Visible = True
    ChartControl.Content.Legend.HorizontalAlignment = xtpChartLegendFarOutside
   
    Dim Series As ChartSeries
    Set Series = ChartControl.Content.Series.Add("Top 10 States by Population")
    
    CreateBubblePoint Series.Points, "California", 8, 37, 11.95
    CreateBubblePoint Series.Points, "Texas", 7, 25, 7.81
    CreateBubblePoint Series.Points, "New York", 6, 20, 6.31
    CreateBubblePoint Series.Points, "Florida", 5, 18, 5.97
    CreateBubblePoint Series.Points, "Illinois", 9, 13, 4.2
    CreateBubblePoint Series.Points, "Pennsylvania", 10, 12, 4.06
    CreateBubblePoint Series.Points, "Ohio", 4, 11, 3.75
    CreateBubblePoint Series.Points, "Michigan", 3, 10, 3.29
    CreateBubblePoint Series.Points, "Georgia", 2, 9, 3.12
    CreateBubblePoint Series.Points, "North Carolina", 1, 9, 3
                            
    Series.ArgumentScaleType = xtpChartScaleQualitative
           
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub CreateBubblePoint(ByVal pPointCollection As ChartSeriesPointCollection, lpszLegendText As String, nYear As Integer, nValue As Integer, dWidth As Double)
On Error GoTo tratar_erro
    
    Dim pPoint As ChartSeriesPoint
    Set pPoint = pPointCollection.Add2(nYear, nValue, dWidth)
    pPoint.LegendText = lpszLegendText

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub AddCandelStyleSeries()
On Error GoTo tratar_erro
    
    lblLabelPosition.Visible = False
    cmbPieLabelPosition.Visible = False
    cmbPyramidLabelPosition.Visible = False
    lbl3dHint.Visible = False
    
    If ChartControl.Content.Series.Count > 0 Then
        ChartControl.Content.Series.DeleteAll
    End If
    
    ChartControl.Content.Titles.DeleteAll
    ChartControl.Content.Titles.Add "Historical Stock Prices"
    
    Dim Series As ChartSeries
    Set Series = ChartControl.Content.Series.Add("Stock Series 1")
       
    ChartControl.Content.Legend.Visible = False
       
    'Series.Points.Add4 DateTime.Date, 0.5, 3, 1.3, 2.5
    'Series.Points.Add4 (DateTime.Date + 1), 1, 7, 2.5, 4.5
    'Series.Points.Add4 (DateTime.Date + 2), 2, 6, 4.5, 5.5
    'Series.Points.Add4 (DateTime.Date + 3), 3, 8, 5.5, 4.5
    'Series.Points.Add4 (DateTime.Date + 4), 1, 8, 4.5, 7
    'Series.Points.Add4 (DateTime.Date + 5), 0, 9, 7, 2
    'Series.Points.Add4 (DateTime.Date + 6), 1.5, 8, 2, 6
    'Series.Points.Add4 (DateTime.Date + 7), 4, 6.5, 6, 4.5
    'Series.Points.Add4 (DateTime.Date + 8), 3, 5, 4.5, 4.5

    Series.Points.Add4 "2009-12-28", 30.89, 31.18, 31#, 31.17
    Series.Points.Add4 "2009-12-29", 31.23, 31.5, 31.35, 31.39
    Series.Points.Add4 "2009-12-30", 30.8, 31.29, 31.15, 30.96
    Series.Points.Add4 "2009-12-31", 30.48, 30.99, 30.98, 30.48
    Series.Points.Add4 "2010-01-04", 30.59, 31.1, 30.62, 30.95
    Series.Points.Add4 "2010-01-05", 30.64, 31.1, 30.85, 30.96
    Series.Points.Add4 "2010-01-06", 30.52, 31.08, 30.88, 30.77
    Series.Points.Add4 "2010-01-07", 30.19, 30.7, 30.63, 30.45
    Series.Points.Add4 "2010-01-08", 30.24, 30.88, 30.28, 30.66
    Series.Points.Add4 "2010-01-11", 30.12, 30.76, 30.71, 30.27
    Series.Points.Add4 "2010-01-12", 29.91, 30.4, 30.15, 30.07
    Series.Points.Add4 "2010-01-13", 30.01, 30.52, 30.26, 30.35
    Series.Points.Add4 "2010-01-14", 30.26, 31.1, 30.31, 30.96
    Series.Points.Add4 "2010-01-15", 30.71, 31.24, 31.08, 30.86
    Series.Points.Add4 "2010-01-19", 30.68, 31.24, 30.75, 31.1
    Series.Points.Add4 "2010-01-20", 30.31, 30.94, 30.81, 30.59
    Series.Points.Add4 "2010-01-21", 30#, 30.72, 30.61, 30.01
    Series.Points.Add4 "2010-01-22", 28.84, 30.2, 30#, 28.96
    Series.Points.Add4 "2010-01-25", 29.1, 29.66, 29.24, 29.32
    Series.Points.Add4 "2010-01-26", 29.09, 29.85, 29.2, 29.5
    Series.Points.Add4 "2010-01-27", 29.02, 29.82, 29.35, 29.67
    Series.Points.Add4 "2010-01-28", 28.89, 29.87, 29.84, 29.16
    Series.Points.Add4 "2010-01-29", 27.66, 29.92, 29.9, 28.18
    Series.Points.Add4 "2010-02-01", 27.92, 28.48, 28.39, 28.41
    Series.Points.Add4 "2010-02-02", 28.14, 28.5, 28.37, 28.46
    Series.Points.Add4 "2010-02-03", 28.12, 28.79, 28.26, 28.63
    Series.Points.Add4 "2010-02-04", 27.81, 28.5, 28.38, 27.84
    Series.Points.Add4 "2010-02-05", 27.57, 28.28, 28#, 28.02
    Series.Points.Add4 "2010-02-08", 27.57, 28.08, 28.01, 27.72
    Series.Points.Add4 "2010-02-09", 27.75, 28.34, 27.97, 28.01
    Series.Points.Add4 "2010-02-10", 27.84, 28.24, 28.03, 27.99
    Series.Points.Add4 "2010-02-11", 27.7, 28.4, 27.93, 28.12
    Series.Points.Add4 "2010-02-12", 27.58, 28.06, 27.81, 27.93
    Series.Points.Add4 "2010-02-16", 28.02, 28.37, 28.13, 28.35
    Series.Points.Add4 "2010-02-17", 28.36, 28.65, 28.53, 28.59
    Series.Points.Add4 "2010-02-18", 28.51, 29.03, 28.59, 28.97
    Series.Points.Add4 "2010-02-19", 28.69, 28.92, 28.79, 28.77
    Series.Points.Add4 "2010-02-22", 28.65, 28.94, 28.84, 28.73
    Series.Points.Add4 "2010-02-23", 28.09, 28.83, 28.68, 28.33
    Series.Points.Add4 "2010-02-24", 28.38, 28.79, 28.52, 28.63
    Series.Points.Add4 "2010-02-25", 28.02, 28.65, 28.27, 28.6
    Series.Points.Add4 "2010-02-26", 28.51, 28.85, 28.65, 28.67
    Series.Points.Add4 "2010-03-01", 28.53, 29.05, 28.77, 29.02
    Series.Points.Add4 "2010-03-02", 28.24, 29.3, 29.08, 28.46
    Series.Points.Add4 "2010-03-03", 28.35, 28.61, 28.51, 28.46
    'Series.Points.Add4 "2010-03-04", 28.27, 28.65, 28.46, 28.63
    'Series.Points.Add4 "2010-03-05", 28.42, 28.68, 28.66, 28.59
    'Series.Points.Add4 "2010-03-08", 28.5, 28.93, 28.52, 28.63
    'Series.Points.Add4 "2010-03-09", 28.55, 29.11, 28.56, 28.8
    'Series.Points.Add4 "2010-03-10", 28.8, 29.11, 28.86, 28.97
    'Series.Points.Add4 "2010-03-11", 28.85, 29.19, 28.89, 29.18
    'Series.Points.Add4 "2010-03-12", 29.04, 29.38, 29.32, 29.27
    'Series.Points.Add4 "2010-03-15", 29.01, 29.37, 29.18, 29.29
    'Series.Points.Add4 "2010-03-16", 29.2, 29.49, 29.42, 29.37
    'Series.Points.Add4 "2010-03-17", 29.4, 29.87, 29.5, 29.63
    'Series.Points.Add4 "2010-03-18", 29.5, 29.72, 29.63, 29.61
    'Series.Points.Add4 "2010-03-19", 29.35, 29.9, 29.76, 29.59
    'Series.Points.Add4 "2010-03-22", 29.39, 29.7, 29.5, 29.6
    'Series.Points.Add4 "2010-03-23", 29.41, 29.9, 29.59, 29.88
    'Series.Points.Add4 "2010-03-24", 29.6, 29.85, 29.72, 29.65
    
    Series.ArgumentScaleType = xtpChartScaleDateTime
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub AddPopulationSeries()
On Error GoTo tratar_erro
    
    If ChartControl.Content.Series.Count > 0 Then
        ChartControl.Content.Series.DeleteAll
    End If
    
    ChartControl.Content.Titles(0).Text = "Gráfico de fluxo de caixa"
    ChartControl.Content.Legend.Visible = True
       
    Dim Series As ChartSeries
    Set Series = ChartControl.Content.Series.Add("Á Pagar")
    
    
    Set TBAbrir = CreateObject("adodb.recordset")
    
    StrSql = "select Day(Data) as Dia, Month(Data) as Mes,Year(Data) as Ano, Sum(Valor) as Valor from Fluxo_de_caixa_projetado_resumido where Operacao = 'À Debitar' and data > '01/12/2020' and data < '01/01/2021' and ID_Empresa = '1' Group By Day(Data),Month(Data),Year(Data) order by  Year(Data)"
    
    TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
        
        Do While TBAbrir.EOF = False
            Series.Points.Add TBAbrir!Dia & "/" & TBAbrir!Mes, TBAbrir!valor
        TBAbrir.MoveNext
        Loop
    End If
    
    Set Series = ChartControl.Content.Series.Add("Á receber")
    Set TBAbrir = CreateObject("adodb.recordset")
    
    StrSql = "select Day(Data) as Dia, Month(Data) as Mes,Year(Data) as Ano, Sum(Valor) as Valor from Fluxo_de_caixa_projetado_resumido where Operacao = 'À Creditar' and data > '01/12/2020' and data < '01/01/2021' and ID_Empresa = '1' Group By Day(Data),Month(Data),Year(Data) order by  Year(Data)"
     TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
        
        Do While TBAbrir.EOF = False
            Series.Points.Add TBAbrir!Dia & "/" & TBAbrir!Mes, TBAbrir!valor
        TBAbrir.MoveNext
        Loop
    End If
    
    Series.PointLegendFormat = "{A}"
    Series.Points(3).Special = True
       

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub AddSeriesCombination()
On Error GoTo tratar_erro
    
Dim SaldoInicial As Double

    
    If ChartControl.Content.Series.Count > 0 Then
        ChartControl.Content.Series.DeleteAll
    End If
    
    ChartControl.Content.Legend.Visible = True
    ChartControl.Content.Legend.HorizontalAlignment = xtpChartLegendNearOutside
    
    Dim Series As ChartSeries
            
    Set Series = ChartControl.Content.Series.Add("Á Pagar")
    
    
    Set TBAbrir = CreateObject("adodb.recordset")
    
    StrSql = "select Day(Data) as Dia, Month(Data) as Mes,Year(Data) as Ano, Sum(Valor) as Valor from Fluxo_de_caixa_projetado_resumido where Operacao = 'À Debitar' and data > '01/12/2020' and data < '01/01/2021' and ID_Empresa = '1' Group By Day(Data),Month(Data),Year(Data) order by  Year(Data)"
    
    TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
        
        Do While TBAbrir.EOF = False
            Series.Points.Add TBAbrir!Dia & "/" & TBAbrir!Mes, TBAbrir!valor
        TBAbrir.MoveNext
        Loop
    End If
    
   Set Series = ChartControl.Content.Series.Add("Á receber")
    Set TBAbrir = CreateObject("adodb.recordset")
    
    StrSql = "select Day(Data) as Dia, Month(Data) as Mes,Year(Data) as Ano, Sum(Valor) as Valor from Fluxo_de_caixa_projetado_resumido where Operacao = 'À Creditar' and data > '01/12/2020' and data < '01/01/2021' and ID_Empresa = '1' Group By Day(Data),Month(Data),Year(Data) order by  Year(Data)"
     TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
        
        Do While TBAbrir.EOF = False
            Series.Points.Add TBAbrir!Dia & "/" & TBAbrir!Mes, TBAbrir!valor
        TBAbrir.MoveNext
        Loop
    End If

SaldoInicial = txtSaldoInicial.Text
'SaldoInicial = Replace(txtSaldoInicial.Text, ",", ".")
'strCaminho = Replace(txtSaldoInicial.Text, ",", ".")
   Set Series = ChartControl.Content.Series.Add("Saldo")
    Set TBAbrir = CreateObject("adodb.recordset")
    
    StrSql = "select Day(Data) as Dia, Month(Data) as Mes,Year(Data) as Ano, Sum(Valor_creditar-Valor_debitar) as Saldo from Fluxo_de_caixa_resumido where data > '01/12/2020' and data < '01/01/2021' and ID_Empresa = '1' Group By Day(Data),Month(Data),Year(Data) order by  Year(Data)"
    
    'StrSql = "select Day(Data) as Dia, Month(Data) as Mes,Year(Data) as Ano, Sum('" & strCaminho & "' +(Valor_creditar-Valor_debitar)) as Saldo from Fluxo_de_caixa_resumido where data > '01/12/2020' and data < '01/01/2021' and ID_Empresa = '1' Group By Day(Data),Month(Data),Year(Data) order by  Year(Data)"
    Debug.Print StrSql
    
     TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
        
        Do While TBAbrir.EOF = False
            Series.Points.Add TBAbrir!Dia & "/" & TBAbrir!Mes, TBAbrir!Saldo + SaldoInicial
            SaldoInicial = TBAbrir!Saldo + SaldoInicial
        TBAbrir.MoveNext
        Loop
    End If

    

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ChartControl_MouseMove(Button As Integer, Shift As Integer, x As Long, Y As Long)
On Error GoTo tratar_erro
    
    Dim Element As ChartElement
    Set Element = ChartControl.HitTest(x, Y)
    
    If TypeOf Element Is XtremeChartControl.ChartSeriesPoint Then
        Dim point As ChartSeriesPoint
        On Error Resume Next
        Set point = Element
    
        If (Not point Is Nothing) Then
            Debug.Print "Point " & point.Value(0)
        End If
    End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbAppearance_Click()
On Error GoTo tratar_erro
    
    ChartControl.Content.Appearance.SetAppearance cmbAppearance.List(cmbAppearance.ListIndex)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub cmbPalette_Click()
On Error GoTo tratar_erro
    
    ChartControl.Content.Appearance.SetPalette cmbPalette.List(cmbPalette.ListIndex)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub cmbStyle_Click()
On Error GoTo tratar_erro
  
AddSeriesCombination


ChartControl.Content.Titles(0).Text = "Gráfico de fluxo de caixa"
ChartControl.Content.Legend.Visible = True
            
If (cmbStyle.ListIndex = 0) Then
'               Set ChartControl.Content.Series(3).Style = New ChartAreaSeriesStyle
    Set ChartControl.Content.Series(2).Style = New ChartPointSeriesStyle
    Set ChartControl.Content.Series(1).Style = New ChartBarSeriesStyle
    Set ChartControl.Content.Series(0).Style = New ChartSplineSeriesStyle
ElseIf (cmbStyle.ListIndex = 1) Then
'               Set ChartControl.Content.Series(3).Style = New ChartAreaSeriesStyle
    Set ChartControl.Content.Series(2).Style = New ChartAreaSeriesStyle
    Set ChartControl.Content.Series(1).Style = New ChartAreaSeriesStyle
    Set ChartControl.Content.Series(0).Style = New ChartAreaSeriesStyle
ElseIf (cmbStyle.ListIndex = 2) Then
'               Set ChartControl.Content.Series(3).Style = New ChartSplineSeriesStyle
    Set ChartControl.Content.Series(2).Style = New ChartSplineSeriesStyle
    Set ChartControl.Content.Series(1).Style = New ChartSplineSeriesStyle
    Set ChartControl.Content.Series(0).Style = New ChartSplineSeriesStyle
ElseIf (cmbStyle.ListIndex = 3) Then
'               Set ChartControl.Content.Series(3).Style = New ChartBarSeriesStyle
'               Set ChartControl.Content.Series(2).Style = New ChartBarSeriesStyle
    Set ChartControl.Content.Series(1).Style = New ChartBarSeriesStyle
    Set ChartControl.Content.Series(0).Style = New ChartBarSeriesStyle
ElseIf (cmbStyle.ListIndex = 4) Then
'               Set ChartControl.Content.Series(3).Style = New ChartPointSeriesStyle
'               Set ChartControl.Content.Series(2).Style = New ChartPointSeriesStyle
    Set ChartControl.Content.Series(1).Style = New ChartPointSeriesStyle
    Set ChartControl.Content.Series(0).Style = New ChartPointSeriesStyle
End If
            
Dim Series As ChartSeries
For Each Series In ChartControl.Content.Series
    Series.Style.Label.Format.Category = xtpChartNumber
    Series.Style.Label.Format.DecimalPlaces = 2
Next

Set Diagram = ChartControl.Content.Series(0).Diagram
Diagram.AxisY.Title = "Valor mensal"
Diagram.AxisY.Title.Visible = True
Diagram.AxisX.Title = "Dia / Mês"
Diagram.AxisX.Title.Visible = True

Diagram.AxisY.Label.Format.Category = xtpChartNumber
Diagram.AxisY.Label.Format.DecimalPlaces = 2


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub cmdPrintPreview_Click()
On Error GoTo tratar_erro
    
    If (ChartControl.PrintPreview) Then
        ChartControl.PrintChart 0
    End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdSave_Click()
On Error GoTo tratar_erro
    
    ChartControl.SaveAsImage App.Path + "\Chart.png", ChartControl.Width, ChartControl.Height
    
    ShellExecute Me.hWnd, "open", App.Path + "\Chart.png", vbNullString, vbNullString, 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Command1_Click()
On Error GoTo tratar_erro

Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Command2_Click()
On Error GoTo tratar_erro

cmbStyle_Click

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Command3_Click()
On Error GoTo tratar_erro

Set TBAbrir = CreateObject("adodb.recordset")

StrSql = "select sum(Saldo) as SaldoInicial from tbl_Instituicoes"
TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
   If TBAbrir.EOF = False Then
       txtSaldoInicial.Text = TBAbrir!SaldoInicial
   End If
   TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro
    
    AddTitles
    
    ChartControl.Content.Legend.Visible = True
    ChartControl.Content.Legend.HorizontalAlignment = xtpChartLegendFar
    ChartControl.Content.EnableMarkup = True
    
    cmbStyle.ListIndex = 1 '17
    cmbAppearance.ListIndex = 0
    cmbPalette.ListIndex = 26

'    Set TBAbrir = CreateObject("adodb.recordset")
'
'    StrSql = "select sum(Saldo) as SaldoInicial from tbl_Instituicoes"
'     TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
'        If TBAbrir.EOF = False Then
'            txtSaldoInicial.Text = TBAbrir!SaldoInicial
'        End If
'        TBAbrir.Close
'
    

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

  ChartControl.Move ChartControl.Left, ChartControl.Top, Me.ScaleWidth - 200 - ChartControl.Top, Me.ScaleHeight - 2 * ChartControl.Top
  panelOptions.Move Me.ScaleWidth - 200, 0, 200, Me.ScaleHeight

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
