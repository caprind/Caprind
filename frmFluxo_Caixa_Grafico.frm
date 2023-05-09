VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{C215CB9A-0AE1-499F-A101-48B3C370D3DF}#19.3#0"; "Codejock.ChartControl.v19.3.0.ocx"
Begin VB.Form frmFluxo_Caixa_Grafico 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Graficos de fluxo de caixa"
   ClientHeight    =   10035
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15270
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
   ScaleHeight     =   10035
   ScaleWidth      =   15270
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   825
      Left            =   60
      TabIndex        =   1
      Top             =   990
      Width           =   15165
      Begin VB.ComboBox cmbPalette 
         Height          =   315
         ItemData        =   "frmFluxo_Caixa_Grafico.frx":0000
         Left            =   12030
         List            =   "frmFluxo_Caixa_Grafico.frx":0061
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   390
         Width           =   1725
      End
      Begin VB.ComboBox cmbAppearance 
         Height          =   315
         ItemData        =   "frmFluxo_Caixa_Grafico.frx":01AC
         Left            =   10680
         List            =   "frmFluxo_Caixa_Grafico.frx":01B9
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   390
         Width           =   1335
      End
      Begin VB.ComboBox cmbStyle 
         Height          =   315
         ItemData        =   "frmFluxo_Caixa_Grafico.frx":01D2
         Left            =   8520
         List            =   "frmFluxo_Caixa_Grafico.frx":01E5
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   390
         Width           =   2145
      End
      Begin VB.ComboBox Cmb_empresa 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "frmFluxo_Caixa_Grafico.frx":024C
         Left            =   180
         List            =   "frmFluxo_Caixa_Grafico.frx":024E
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         ToolTipText     =   "Empresa."
         Top             =   390
         Width           =   8340
      End
      Begin MSComCtl2.DTPicker msk_fltFim 
         Height          =   315
         Left            =   13770
         TabIndex        =   2
         ToolTipText     =   "Data final para pesquisa."
         Top             =   390
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
         Format          =   197263363
         CurrentDate     =   39057
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Empresa"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   150
         TabIndex        =   12
         Top             =   150
         Width           =   8355
      End
      Begin VB.Label lblPalette 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cores"
         Height          =   255
         Left            =   12630
         TabIndex        =   11
         Top             =   150
         Width           =   465
      End
      Begin VB.Label lblAppearance 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Aparência"
         Height          =   255
         Left            =   10680
         TabIndex        =   9
         Top             =   150
         Width           =   1215
      End
      Begin VB.Label lblStyle 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Estilo"
         Height          =   255
         Left            =   8520
         TabIndex        =   7
         Top             =   150
         Width           =   2175
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Filtrar até"
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
         Left            =   13905
         TabIndex        =   3
         Top             =   150
         Width           =   840
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   8265
      Left            =   13380
      TabIndex        =   13
      Top             =   1770
      Width           =   1845
      Begin VB.TextBox txtSaldoinicial 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   120
         TabIndex        =   27
         TabStop         =   0   'False
         Text            =   "0,00"
         Top             =   480
         Width           =   1575
      End
      Begin DrawSuite2022.USCheckBox chkVencidos 
         Height          =   225
         Left            =   1410
         TabIndex        =   25
         ToolTipText     =   "Considerar valores vencidos"
         Top             =   2820
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   397
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocusRect   =   0   'False
      End
      Begin VB.TextBox txtSaldo_Atual1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Text            =   "0,00"
         Top             =   7800
         Width           =   1575
      End
      Begin VB.TextBox txtSaldofinalvenc1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Text            =   "0,00"
         Top             =   3060
         Width           =   1575
      End
      Begin VB.TextBox txtSaldo_Atual 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Text            =   "0,00"
         Top             =   2400
         Width           =   1575
      End
      Begin VB.TextBox txttotaldebito 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Text            =   "0,00"
         Top             =   1770
         Width           =   1575
      End
      Begin VB.TextBox txttotalcredito 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Text            =   "0,00"
         Top             =   1110
         Width           =   1575
      End
      Begin DrawSuite2022.USCheckBox chkSaldoInicial 
         Height          =   225
         Left            =   1440
         TabIndex        =   26
         ToolTipText     =   "Considerar Saldo em caixa"
         Top             =   240
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   397
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocusRect   =   0   'False
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(= Saldo final)"
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
         Left            =   300
         TabIndex        =   24
         Top             =   7590
         Width           =   1200
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(+ Vencidos)"
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
         Left            =   165
         TabIndex        =   22
         Top             =   2850
         Width           =   1080
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(= Saldo final)"
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
         Left            =   330
         TabIndex        =   20
         Top             =   2190
         Width           =   1200
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(- Total débito)"
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
         Left            =   285
         TabIndex        =   18
         Top             =   1560
         Width           =   1290
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(+ Total crédito)"
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
         Left            =   225
         TabIndex        =   16
         Top             =   900
         Width           =   1410
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(Saldo inicial)"
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
         Left            =   240
         TabIndex        =   14
         Top             =   270
         Width           =   1140
      End
   End
   Begin XtremeChartControl.ChartControl ChartControl 
      Height          =   8145
      Left            =   60
      TabIndex        =   0
      Top             =   1860
      Width           =   13305
      _Version        =   1245187
      _ExtentX        =   23469
      _ExtentY        =   14367
      _StockProps     =   0
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
      FormHeightDT    =   10500
      FormWidthDT     =   15390
      FormScaleHeightDT=   10035
      FormScaleWidthDT=   15270
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   1005
      Left            =   60
      TabIndex        =   5
      Top             =   0
      Width           =   15165
      _ExtentX        =   26749
      _ExtentY        =   1773
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
      ButtonKey2      =   "4"
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
      ButtonEnabled3  =   0   'False
      ButtonIconSize3 =   32
      ButtonAlignment3=   2
      ButtonType3     =   1
      ButtonStyle3    =   -1
      BeginProperty ButtonFont3 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState3    =   -1
      ButtonLeft3     =   93
      ButtonTop3      =   4
      ButtonWidth3    =   2
      ButtonHeight3   =   56
      ButtonCaption4  =   "Ajuda"
      ButtonEnabled4  =   0   'False
      ButtonIconSize4 =   32
      ButtonToolTipText4=   "Ajuda (F1)"
      ButtonKey4      =   "9"
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
      ButtonLeft4     =   97
      ButtonTop4      =   2
      ButtonWidth4    =   36
      ButtonHeight4   =   21
      ButtonUseMaskColor4=   0   'False
      ButtonCaption5  =   "Sair"
      ButtonEnabled5  =   0   'False
      ButtonIconSize5 =   32
      ButtonToolTipText5=   "Sair (Esc)"
      ButtonKey5      =   "10"
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
      ButtonLeft5     =   135
      ButtonTop5      =   2
      ButtonWidth5    =   26
      ButtonHeight5   =   21
      ButtonUseMaskColor5=   0   'False
      ButtonEnabled6  =   0   'False
      ButtonIconSize6 =   32
      ButtonKey6      =   "11"
      ButtonAlignment6=   2
      BeginProperty ButtonFont6 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState6    =   5
      ButtonLeft6     =   163
      ButtonTop6      =   2
      ButtonWidth6    =   24
      ButtonHeight6   =   24
      ButtonUseMaskColor6=   0   'False
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   7290
         Top             =   270
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmFluxo_Caixa_Grafico.frx":0250
         Count           =   1
      End
   End
End
Attribute VB_Name = "frmFluxo_Caixa_Grafico"
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

Sub CreateSeriesPoint(ByVal pPointCollection As ChartSeriesPointCollection, vArg As String, nValue As Double)
On Error GoTo tratar_erro
    
    Dim pPoint As ChartSeriesPoint
    Set pPoint = pPointCollection.Add(vArg, nValue)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaSaldoInicial()
On Error GoTo tratar_erro
    
If chkSaldoInicial.Value = Checked Then
    Set TBAbrir = CreateObject("adodb.recordset")
    
    StrSql = "select Sum(Saldo) as SaldoInicial from Tbl_Instituicoes where ID_Empresa = '" & ID_empresa & "'"
    
    TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
          txtSaldoinicial.Text = Format(TBAbrir!SaldoInicial, "###,##0.00")
        End If
Else
        txtSaldoinicial.Text = "0,00"
End If

ProcFiltrar
    
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
    
    StrSql = "select Data, Sum(Valor_Debitar) as Valor from Fluxo_de_caixa_resumido where data >= '" & Date & "' and data <= '" & msk_fltFim & "' and ID_Empresa = '" & ID_empresa & "' Group By Data order by Data"
    
    TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
        
        Do While TBAbrir.EOF = False
            Series.Points.Add str(TBAbrir!Data), TBAbrir!valor
        TBAbrir.MoveNext
        Loop
    End If
    
   Set Series = ChartControl.Content.Series.Add("Á receber")
    Set TBAbrir = CreateObject("adodb.recordset")
    
     StrSql = "select Data, Sum(Valor_Creditar) as Valor from Fluxo_de_caixa_resumido where data >= '" & Date & "' and data <= '" & msk_fltFim & "' and ID_Empresa = '" & ID_empresa & "' Group By Data order by Data"
     
     TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
        
        Do While TBAbrir.EOF = False
            Series.Points.Add str(TBAbrir!Data), TBAbrir!valor
        TBAbrir.MoveNext
        Loop
    End If

    SaldoInicial = txtSaldoinicial.Text
   Set Series = ChartControl.Content.Series.Add("Saldo")
    Set TBAbrir = CreateObject("adodb.recordset")
    
    StrSql = "select Data,Sum(Valor_creditar-Valor_debitar) as Saldo from Fluxo_de_caixa_resumido where data >= '" & Date & "' and data <= '" & msk_fltFim & "' and ID_Empresa = '" & ID_empresa & "' Group By Data order by  Data"
    'Debug.print StrSql
    
     TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
        
        Do While TBAbrir.EOF = False
            Series.Points.Add str(TBAbrir!Data), TBAbrir!Saldo + SaldoInicial
            SaldoInicial = TBAbrir!Saldo + SaldoInicial
        TBAbrir.MoveNext
        Loop
    End If

    

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro
  
AddSeriesCombination


ChartControl.Content.Titles(0).Text = "Gráfico de fluxo de caixa"
ChartControl.Content.Legend.Visible = True
            
If (cmbStyle.ListIndex = 0) Then
    Set ChartControl.Content.Series(2).Style = New ChartPointSeriesStyle
    Set ChartControl.Content.Series(1).Style = New ChartBarSeriesStyle
    Set ChartControl.Content.Series(0).Style = New ChartSplineSeriesStyle
ElseIf (cmbStyle.ListIndex = 1) Then
    Set ChartControl.Content.Series(2).Style = New ChartAreaSeriesStyle
    Set ChartControl.Content.Series(1).Style = New ChartAreaSeriesStyle
    Set ChartControl.Content.Series(0).Style = New ChartAreaSeriesStyle
ElseIf (cmbStyle.ListIndex = 2) Then
    Set ChartControl.Content.Series(2).Style = New ChartSplineSeriesStyle
    Set ChartControl.Content.Series(1).Style = New ChartSplineSeriesStyle
    Set ChartControl.Content.Series(0).Style = New ChartSplineSeriesStyle
ElseIf (cmbStyle.ListIndex = 3) Then
    Set ChartControl.Content.Series(2).Style = New ChartBarSeriesStyle
    Set ChartControl.Content.Series(1).Style = New ChartBarSeriesStyle
    Set ChartControl.Content.Series(0).Style = New ChartBarSeriesStyle
ElseIf (cmbStyle.ListIndex = 4) Then
    Set ChartControl.Content.Series(2).Style = New ChartPointSeriesStyle
    Set ChartControl.Content.Series(1).Style = New ChartPointSeriesStyle
    Set ChartControl.Content.Series(0).Style = New ChartPointSeriesStyle
End If
            
Dim Series As ChartSeries
For Each Series In ChartControl.Content.Series
    Series.Style.Label.Format.Category = xtpChartNumber
    Series.Style.Label.Format.DecimalPlaces = 2
Next

Set Diagram = ChartControl.Content.Series(0).Diagram
Diagram.AxisY.Title = "Valor"
Diagram.AxisY.Title.Visible = True
Diagram.AxisX.Title = "Dia/Mês/Ano"
Diagram.AxisX.Title.Visible = True

Diagram.AxisY.Label.Format.Category = xtpChartNumber
Diagram.AxisY.Label.Format.DecimalPlaces = 2

ProcTotalCredito
ProcTotalDebito
ProcCalculaSaldo
ProcTotalVencido
ProcCalculaSaldoFinal

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCalculaSaldo()
On Error GoTo tratar_erro
Dim TotalCredito As Double
Dim TotalDebito As Double
Dim Saldo As Double

Saldo = txtSaldoinicial
TotalCredito = txttotalcredito
TotalDebito = txttotaldebito
Saldo = (Saldo + TotalCredito) - TotalDebito

txtSaldo_Atual.Text = Format(Saldo, "###,##0.00")


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCalculaSaldoFinal()
On Error GoTo tratar_erro
Dim TotalVencido As Double
Dim SaldoAtual As Double

SaldoAtual = txtSaldo_Atual.Text
TotalVencido = txtSaldofinalvenc1.Text

SaldoAtual = SaldoAtual + TotalVencido
txtSaldo_Atual1.Text = Format(SaldoAtual, "###,##0.00")


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub btnImagem_Click()
On Error GoTo tratar_erro
    
    ChartControl.SaveAsImage App.Path + "\Chart.png", ChartControl.Width, ChartControl.Height
    
    ShellExecute Me.hWnd, "open", App.Path + "\Chart.png", vbNullString, vbNullString, 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro
    
    If (ChartControl.PrintPreview) Then
        ChartControl.PrintChart 0
    End If

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

Private Sub ChartControl_MouseMove(Button As Integer, Shift As Integer, x As Long, Y As Long)
On Error GoTo tratar_erro
    
    Dim Element As ChartElement
    Set Element = ChartControl.HitTest(x, Y)
    
    If TypeOf Element Is XtremeChartControl.ChartSeriesPoint Then
        Dim point As ChartSeriesPoint
        On Error Resume Next
        Set point = Element
    
        If (Not point Is Nothing) Then
            'Debug.print "Point " & point.Value(0)
        End If
    End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkSaldoInicial_Click()
On Error GoTo tratar_erro
    
ProcCarregaSaldoInicial

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkVencidos_Click()
On Error GoTo tratar_erro
    
ProcTotalVencido
ProcCalculaSaldoFinal

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub

End Sub

Private Sub Cmb_empresa_Change()
On Error GoTo tratar_erro
    
ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & "

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_empresa_Click()
On Error GoTo tratar_erro
    
ID_empresa = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)

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
  
ProcFiltrar

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcTotalVencido()
On Error GoTo tratar_erro
Dim TotalVencido As Double

If chkVencidos.Value = Checked Then
    Set TBAbrir = CreateObject("adodb.recordset")
    
     StrSql = "select Sum(Valor_Debitar) as ValorDebitar,Sum(Valor_Creditar) as ValorCreditar from Fluxo_de_caixa_resumido where data < '" & Date & "' and ID_Empresa = '" & ID_empresa & "'"
     'Debug.print StrSql
     
     TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
        TotalVencido = TBAbrir!ValorCreditar - TBAbrir!ValorDebitar
        txtSaldofinalvenc1.Text = Format(TotalVencido, "###,##0.00")
        End If
    TBAbrir.Close
Else
txtSaldofinalvenc1.Text = "0,00"
End If

    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub ProcTotalCredito()
On Error GoTo tratar_erro

    Set TBAbrir = CreateObject("adodb.recordset")
    
     StrSql = "select Sum(Valor_Creditar) as Valor from Fluxo_de_caixa_resumido where data >= '" & Date & "' and data <= '" & msk_fltFim & "' and ID_Empresa = '" & ID_empresa & "'"
     'Debug.print StrSql
     
     TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
        txttotalcredito.Text = Format(TBAbrir!valor, "###,##0.00")
        End If
    TBAbrir.Close
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub ProcTotalDebito()
On Error GoTo tratar_erro

    Set TBAbrir = CreateObject("adodb.recordset")
    
     StrSql = "select Sum(Valor_Debitar) as Valor from Fluxo_de_caixa_resumido where data >= '" & Date & "' and data <= '" & msk_fltFim & "' and ID_Empresa = '" & ID_empresa & "'"
     
     TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
        txttotaldebito.Text = Format(TBAbrir!valor, "###,##0.00")
        End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15195, 6, True

    msk_fltFim.Value = Date + 30
    ProcCarregaComboEmpresa Cmb_empresa, False
    ID_empresa = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
    txtSaldoinicial.Text = "0,0"
     
    AddTitles
    
    ChartControl.Content.Legend.Visible = True
    ChartControl.Content.Legend.HorizontalAlignment = xtpChartLegendFar
    ChartControl.Content.EnableMarkup = True
    
    cmbStyle.ListIndex = 2 '17
    cmbAppearance.ListIndex = 0
    cmbPalette.ListIndex = 26

    ProcRemoveObjetosResize Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcFiltrar
    Case 2: ProcImprimir
    Case 4: 'ProcAjuda
    Case 5: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub txtSaldoInicial_LostFocus()

If txtSaldoinicial.Text <> "" Then
txtSaldoinicial = Format(txtSaldoinicial, "###,##0.00")
ProcFiltrar
End If

End Sub
