VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Begin VB.Form frmFaturamento_Prod_serv_exportar 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Faturamento | Nota fiscal - Exportar"
   ClientHeight    =   4995
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   5820
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   5820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   465
      Left            =   0
      TabIndex        =   30
      Top             =   0
      Width           =   5820
      _ExtentX        =   10266
      _ExtentY        =   820
      DibPicture      =   "frmFaturamento_Prod_serv_exportar.frx":0000
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
      Icon            =   "frmFaturamento_Prod_serv_exportar.frx":7180
   End
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   29
      Top             =   4590
      Width           =   5820
      _ExtentX        =   10266
      _ExtentY        =   714
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Nota fiscal"
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
      Height          =   600
      Left            =   3855
      TabIndex        =   28
      Top             =   630
      Width           =   1770
      Begin VB.OptionButton Opt_propria 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Próp."
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
         TabIndex        =   4
         Top             =   300
         Value           =   -1  'True
         Width           =   705
      End
      Begin VB.OptionButton Opt_terceiros 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Terc."
         DisabledPicture =   "frmFaturamento_Prod_serv_exportar.frx":749A
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
         Left            =   900
         TabIndex        =   5
         Top             =   300
         Width           =   675
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   2220
      TabIndex        =   27
      Top             =   1230
      Width           =   3405
      Begin VB.CheckBox Chk_periodo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Filtrar por período"
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
         Left            =   180
         TabIndex        =   8
         Top             =   270
         Width           =   1875
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Exportar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   240
      TabIndex        =   26
      Top             =   630
      Width           =   3585
      Begin VB.OptionButton Opt_NFSe 
         BackColor       =   &H00E0E0E0&
         Caption         =   "NFSe"
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
         TabIndex        =   0
         Top             =   300
         Value           =   -1  'True
         Width           =   675
      End
      Begin VB.OptionButton Opt_NFp 
         BackColor       =   &H00E0E0E0&
         Caption         =   "NFp"
         DisabledPicture =   "frmFaturamento_Prod_serv_exportar.frx":2513DC
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
         Left            =   870
         TabIndex        =   1
         Top             =   300
         Width           =   585
      End
      Begin VB.OptionButton Opt_contimatic 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Contimatic"
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
         Left            =   2400
         TabIndex        =   3
         Top             =   300
         Width           =   1065
      End
      Begin VB.OptionButton Opt_sintegra 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Sintegra"
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
         Left            =   1470
         TabIndex        =   2
         Top             =   300
         Width           =   915
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tipo da nota fiscal"
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
      Height          =   600
      Left            =   240
      TabIndex        =   25
      Top             =   1230
      Width           =   1950
      Begin VB.OptionButton Opt_saida 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Saída"
         DisabledPicture =   "frmFaturamento_Prod_serv_exportar.frx":49B31E
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
         Left            =   1110
         TabIndex        =   7
         Top             =   300
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton Opt_entrada 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Entrada"
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
         TabIndex        =   6
         Top             =   300
         Width           =   885
      End
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   4140
      Width           =   5385
      _ExtentX        =   9499
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor2      =   0
      SearchText      =   ""
      Value           =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
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
      Height          =   2295
      Left            =   240
      TabIndex        =   18
      Top             =   1830
      Width           =   5385
      Begin DrawSuite2022.USButton btnExportar 
         Height          =   465
         Left            =   3240
         TabIndex        =   31
         Top             =   1740
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   820
         Caption         =   "Exportar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
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
         ShowFocusRect   =   0   'False
         Theme           =   4
      End
      Begin VB.ComboBox Cmb_cl_forn_ate 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2490
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   15
         ToolTipText     =   "Nome do cliente/fornecedor."
         Top             =   1350
         Width           =   2715
      End
      Begin VB.ComboBox Cmb_cl_forn_de 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2490
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   13
         ToolTipText     =   "Nome do cliente/fornecedor."
         Top             =   960
         Width           =   2715
      End
      Begin VB.ComboBox Cmb_De 
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
         Height          =   330
         Left            =   1050
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   12
         ToolTipText     =   "Número da nota fiscal."
         Top             =   960
         Width           =   1425
      End
      Begin VB.ComboBox Cmb_Ate 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1050
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   14
         ToolTipText     =   "Número da nota fiscal."
         Top             =   1350
         Width           =   1425
      End
      Begin VB.ComboBox Cmb_tipo 
         BackColor       =   &H00FFFFFF&
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
         ItemData        =   "frmFaturamento_Prod_serv_exportar.frx":6E5260
         Left            =   1050
         List            =   "frmFaturamento_Prod_serv_exportar.frx":6E5267
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   10
         ToolTipText     =   "Tipo da nota fiscal."
         Top             =   570
         Width           =   4155
      End
      Begin VB.ComboBox Cmb_empresa 
         BackColor       =   &H00FFFFFF&
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
         Height          =   315
         ItemData        =   "frmFaturamento_Prod_serv_exportar.frx":6E527A
         Left            =   1050
         List            =   "frmFaturamento_Prod_serv_exportar.frx":6E527C
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   9
         ToolTipText     =   "Empresa."
         Top             =   180
         Width           =   4155
      End
      Begin MSComCtl2.DTPicker msk_fltFim 
         Height          =   315
         Left            =   1050
         TabIndex        =   17
         ToolTipText     =   "Data final."
         Top             =   1350
         Visible         =   0   'False
         Width           =   1425
         _ExtentX        =   2514
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
         Format          =   197066753
         CurrentDate     =   39057
      End
      Begin MSComCtl2.DTPicker msk_fltInicio 
         Height          =   315
         Left            =   1050
         TabIndex        =   16
         ToolTipText     =   "Data inicio."
         Top             =   960
         Visible         =   0   'False
         Width           =   1425
         _ExtentX        =   2514
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
         Format          =   197066753
         CurrentDate     =   39057
      End
      Begin VB.ComboBox Cmb_registro 
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
         Height          =   330
         ItemData        =   "frmFaturamento_Prod_serv_exportar.frx":6E527E
         Left            =   4500
         List            =   "frmFaturamento_Prod_serv_exportar.frx":6E5288
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   11
         ToolTipText     =   "Número do registro."
         Top             =   570
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Reg. :"
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
         Index           =   4
         Left            =   3930
         TabIndex        =   24
         Top             =   570
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "De :"
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
         Index           =   3
         Left            =   660
         TabIndex        =   23
         Top             =   990
         Width           =   300
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Até :"
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
         Index           =   0
         Left            =   600
         TabIndex        =   21
         Top             =   1350
         Width           =   360
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo :"
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
         Index           =   1
         Left            =   555
         TabIndex        =   20
         Top             =   570
         Width           =   405
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Empresa :"
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
         Index           =   2
         Left            =   240
         TabIndex        =   19
         Top             =   180
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmFaturamento_Prod_serv_exportar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim EntradaSaida_NF As String 'OK
Dim Especie_NF As String 'OK
Dim Serie_NF As String 'OK
Dim Numero_NF As String 'OK
Dim CodigoCLI_NF As String 'OK

Private Sub btnExportar_Click()
On Error GoTo tratar_erro

ProcExportar

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_periodo_Click()
On Error GoTo tratar_erro

If Chk_periodo.Value = 1 Then
    Cmb_de.Visible = False
    Cmb_cl_forn_de.Visible = False
    Cmb_Ate.Visible = False
    Cmb_cl_forn_ate.Visible = False
    msk_fltInicio.Visible = True
    msk_fltFim.Visible = True
Else
    Cmb_de.Visible = True
    Cmb_cl_forn_de.Visible = True
    Cmb_Ate.Visible = True
    Cmb_cl_forn_ate.Visible = True
    msk_fltInicio.Visible = False
    msk_fltFim.Visible = False
    ProcCarregaNF
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_Ate_Click()
On Error GoTo tratar_erro

With Cmb_cl_forn_ate
    .Clear
    .Enabled = True
    
    If Opt_contimatic.Value = True Then
        TextoFiltro = ""
        TextoFiltro1 = ""
        TextoFiltro2 = ""
    Else
        If Opt_propria.Value = True Then TextoFiltro = "and Aplicacao = 'P'" Else TextoFiltro = "and Aplicacao = 'T'"
        If opt_Saida.Value = True Then TextoFiltro1 = "and Int_TipoNota = 1" Else TextoFiltro1 = "and Int_TipoNota = 2"
        TextoFiltro2 = "and int_status = 1"
    End If
    
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select txt_Razao_Nome from tbl_Dados_Nota_Fiscal where ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and TipoNF = '" & Tipo & "' and Imprimir = 'False' and DtValidacao IS NOT NULL and int_NotaFiscal = '" & Cmb_Ate & "' " & TextoFiltro & " " & TextoFiltro1 & " " & TextoFiltro2 & " order by txt_Razao_Nome", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Do While TBAbrir.EOF = False
            .AddItem TBAbrir!txt_Razao_Nome
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

Private Sub Cmb_cl_forn_de_Click()
On Error GoTo tratar_erro

With Cmb_Ate
    .Clear
    .Enabled = True
    
    If Opt_contimatic.Value = True Then
        TextoFiltro = ""
        TextoFiltro1 = ""
        TextoFiltro2 = ""
    Else
        If Opt_propria.Value = True Then TextoFiltro = "and Aplicacao = 'P'" Else TextoFiltro = "and Aplicacao = 'T'"
        If opt_Saida.Value = True Then TextoFiltro1 = "and Int_TipoNota = 1" Else TextoFiltro1 = "and Int_TipoNota = 2"
        TextoFiltro2 = "and int_status = 1"
    End If
    
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select int_NotaFiscal from tbl_Dados_Nota_Fiscal where ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and TipoNF = '" & Tipo & "' and Imprimir = 'False' and DtValidacao IS NOT NULL and int_NotaFiscal >= '" & Cmb_de & "' " & TextoFiltro & " " & TextoFiltro1 & " " & TextoFiltro2 & " group by int_NotaFiscal", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Do While TBAbrir.EOF = False
            .AddItem TBAbrir!int_NotaFiscal
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

Private Sub Cmb_De_Click()
On Error GoTo tratar_erro

With Cmb_cl_forn_de
    .Clear
    .Enabled = True
    
    If Opt_contimatic.Value = True Then
        TextoFiltro = ""
        TextoFiltro1 = ""
        TextoFiltro2 = ""
    Else
        If Opt_propria.Value = True Then TextoFiltro = "and Aplicacao = 'P'" Else TextoFiltro = "and Aplicacao = 'T'"
        If opt_Saida.Value = True Then TextoFiltro1 = "and Int_TipoNota = 1" Else TextoFiltro1 = "and Int_TipoNota = 2"
        TextoFiltro2 = "and int_status = 1"
    End If
    
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select txt_Razao_Nome from tbl_Dados_Nota_Fiscal where ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and TipoNF = '" & Tipo & "' and Imprimir = 'False' and DtValidacao IS NOT NULL and int_NotaFiscal = '" & Cmb_de & "' " & TextoFiltro & " " & TextoFiltro1 & " " & TextoFiltro2 & " order by txt_Razao_Nome", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Do While TBAbrir.EOF = False
            .AddItem TBAbrir!txt_Razao_Nome
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

Private Sub Cmb_empresa_Click()
On Error GoTo tratar_erro

ProcCarregaNF
ID_empresa = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_tipo_Click()
On Error GoTo tratar_erro

ProcCarregaNF
ID_empresa = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExportar()
On Error GoTo tratar_erro

If USMsgBox("Deseja realmente exportar esta(s) nota(s) fiscal(ais) para arquivo (.txt)?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    If Chk_periodo.Value = 1 Then
        If FunVerificaDataFinal(msk_fltInicio, msk_fltFim) = False Then Exit Sub
    End If
        
    If Opt_NFSe.Value = True Then
        If TemInternet = True And ErroDriverMYSQL = False Then
            Set TBFIltro = CreateObject("adodb.recordset")
            TBFIltro.Open "Select CNPJ from Empresa where Codigo = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex), Conexao, adOpenKeyset, adLockOptimistic
            If TBFIltro.EOF = False Then
                FunAbreBDSite
                If ConexaoMySql.State = 1 Then
                    Set TBMySQL = New ADODB.Recordset
                    TBMySQL.Open "Select * From Clientes Where CNPJ = '" & TBFIltro!CNPJ & "' and NFSe_liberado = 'NÃO'", ConexaoMySql, adOpenKeyset, adLockOptimistic, adCmdText
                    If TBMySQL.EOF = False Then
                        USMsgBox ("Não é permitido exportar NFSe, pois este recurso está bloqueado."), vbCritical, "CAPRIND v5.0"
                        TBMySQL.Close
                        Exit Sub
                    End If
                    TBMySQL.Close
                End If
            End If
            TBFIltro.Close
'        ElseIf TemInternet = True Then
'                USMsgBox ("Não é permitido exportar NFSe, pois no momento estamos sem conexão com nosso servidor, favor tentar mais tarde."), vbExclamation, "CAPRIND v5.0"
'                Exit Sub
'            Else
'                USMsgBox ("Não é permitido exportar NFSe, pois não existe conexão com a internet."), vbExclamation, "CAPRIND v5.0"
'                Exit Sub
        End If
        caminho = Localrel & "\Arquivos exportados\NFSe Enviar"
    ElseIf Opt_NFp.Value = True Then
            caminho = Localrel & "\Arquivos exportados\NFp Enviar"
        ElseIf Opt_sintegra.Value = True Then
                caminho = Localrel & "\Arquivos exportados\Sintegra Enviar"
            Else
                caminho = Localrel & "\Arquivos exportados\Contimatic Enviar"
    End If
    If GerArqPastas.FolderExists(caminho) = False Then
        USMsgBox ("Não é permitido exportar, pois não foi encontrado o caminho " & caminho & ", onde será armazenado os aquivos."), vbExclamation, "CAPRIND v5.0"
        Exit Sub
    End If
    
    Acao = "exportar"
    If Cmb_tipo = "" Then
        NomeCampo = "o tipo da nota"
        ProcVerificaAcao
        Cmb_tipo.SetFocus
        Exit Sub
    End If
    If Cmb_registro.Visible = True And Cmb_registro = "" Then
        NomeCampo = "o número do registro"
        ProcVerificaAcao
        Cmb_registro.SetFocus
        Exit Sub
    End If
    If Opt_contimatic.Value = True Then
        TextoFiltro = ""
        TextoFiltro1 = ""
        TextoFiltro3 = ""
    Else
        If Opt_propria.Value = True Then TextoFiltro = "and Aplicacao = 'P'" Else TextoFiltro = "and Aplicacao = 'T'"
        If opt_Saida.Value = True Then TextoFiltro1 = "and Int_TipoNota = 1" Else TextoFiltro1 = "and Int_TipoNota = 2"
        TextoFiltro3 = "and int_status = 1 and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
    End If
           
    If Chk_periodo.Value = 1 Then
        If Opt_contimatic.Value = True Then
            TextoFiltro2 = "(tbl_Dados_Nota_Fiscal.Int_TipoNota = 1 and CONVERT(VARCHAR, tbl_Dados_Nota_Fiscal.dt_DataEmissao, 103) Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "' or tbl_Dados_Nota_Fiscal.Int_TipoNota = 2 and tbl_Dados_Nota_Fiscal.dt_Saida_Entrada Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "')"
        Else
            If opt_Saida.Value = True Then DataFiltro = "CONVERT(VARCHAR, dt_DataEmissao, 103)" Else DataFiltro = "dt_Saida_Entrada"
            TextoFiltro2 = DataFiltro & " Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "'"
        End If
    Else
        If Cmb_de = "" Then
            NomeCampo = "o número da nota fiscal"
            ProcVerificaAcao
            Cmb_de.SetFocus
            Exit Sub
        End If
        If Cmb_cl_forn_de = "" Then
            NomeCampo = "o nome do destinatário"
            ProcVerificaAcao
            Cmb_cl_forn_de.SetFocus
            Exit Sub
        End If
        If Cmb_Ate = "" Then
            NomeCampo = "o número da nota fiscal"
            ProcVerificaAcao
            Cmb_Ate.SetFocus
            Exit Sub
        End If
        If Cmb_cl_forn_ate = "" Then
            NomeCampo = "o nome do destinatário"
            ProcVerificaAcao
            Cmb_cl_forn_ate.SetFocus
            Exit Sub
        End If
        TextoFiltro2 = "(tbl_Dados_Nota_Fiscal.int_NotaFiscal = '" & Cmb_de & "' and tbl_Dados_Nota_Fiscal.txt_Razao_Nome = '" & Cmb_cl_forn_de & "' or tbl_Dados_Nota_Fiscal.int_NotaFiscal = '" & Cmb_Ate & "' and tbl_Dados_Nota_Fiscal.txt_Razao_Nome = '" & Cmb_cl_forn_ate & "') or tbl_Dados_Nota_Fiscal.int_NotaFiscal > '" & Cmb_de & "' and tbl_Dados_Nota_Fiscal.int_NotaFiscal < '" & Cmb_Ate & "'"
    End If
    
    'Verifica se é da cidade de Indaiatuba quando for NFSe
    If Opt_NFSe.Value = True Then
        Set TBproducao = CreateObject("adodb.recordset")
        TBproducao.Open "Select Cidade from Empresa where Codigo = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and (Cidade = 'Indaiatuba' or Cidade = 'INDAIATUBA')", Conexao, adOpenKeyset, adLockOptimistic
        If TBproducao.EOF = True Then
            USMsgBox ("Só é permitido exportar este arquivo para emitentes da cidade de Indaiatuba."), vbExclamation, "CAPRIND v5.0"
            TBproducao.Close
            Exit Sub
        End If
        TBproducao.Close
    End If
        
    Set TBproducao = CreateObject("adodb.recordset")
    StrSql = "Select * from tbl_Dados_Nota_Fiscal where TipoNF = '" & Tipo & "' and DtValidacao IS NOT NULL " & TextoFiltro & " " & TextoFiltro1 & " and " & TextoFiltro2 & " " & TextoFiltro3 & " order by int_NotaFiscal, Id_Int_Cliente"
    'Debug.print StrSql
    
    TBproducao.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
    If TBproducao.EOF = False Then
        TBproducao.MoveLast
        
        If Opt_contimatic.Value = True Then IDlista = TBproducao!int_NotaFiscal 'Última NF do filtro
        
        PBLista.Min = 0
        PBLista.Max = TBproducao.RecordCount
        PBLista.Value = 1
        Contador1 = 0
        TBproducao.MoveFirst
        Do While TBproducao.EOF = False
            FamiliaAntiga = FunTiraAcentosTexto(TBproducao!txt_Municipio)
            
            If IsNull(TBproducao!txt_UF) = False And TBproducao!txt_UF <> "" And TBproducao!txt_UF <> "EX" Then
                Set TBFI = CreateObject("adodb.recordset")
                TBFI.Open "Select * from CEP where Municipio = '" & FamiliaAntiga & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBFI.EOF = True Then
                    USMsgBox ("Não é permitido exportar esta nota fiscal " & TBproducao!int_NotaFiscal & " - série " & IIf(IsNull(TBproducao!Serie), "", TBproducao!Serie) & ", pois a nota esta com a cidade errada."), vbExclamation, "CAPRIND v5.0"
                    TBFI.Close
                    Exit Sub
                End If
                Set TBFI = CreateObject("adodb.recordset")
                TBFI.Open "Select * from CEP where Sigla_UF = '" & TBproducao!txt_UF & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBFI.EOF = True Then
                    USMsgBox ("Não é permitido exportar esta nota fiscal " & TBproducao!int_NotaFiscal & " - série " & IIf(IsNull(TBproducao!Serie), "", TBproducao!Serie) & ", pois a nota esta com o estado errado."), vbExclamation, "CAPRIND v5.0"
                    TBFI.Close
                    Exit Sub
                End If
                Set TBFI = CreateObject("adodb.recordset")
                TBFI.Open "Select * from CEP where Municipio = '" & FamiliaAntiga & "' and Sigla_UF = '" & TBproducao!txt_UF & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBFI.EOF = True Then
                    USMsgBox ("Não é permitido exportar esta nota fiscal " & TBproducao!int_NotaFiscal & " - série " & IIf(IsNull(TBproducao!Serie), "", TBproducao!Serie) & ", pois não existe o munícipio " & FamiliaAntiga & " no estado " & UF & " na tabela CEP."), vbExclamation, "CAPRIND v5.0"
                    TBFI.Close
                    Exit Sub
                End If
            End If
            
            'Verifica se tem país cadastrado
            Set TBClientes = CreateObject("adodb.recordset")
            If TBproducao!txt_tipocliente = "E" Then
                'Empresa
                TBClientes.Open "Select * from Empresa where Codigo = " & TBproducao!Id_Int_Cliente, Conexao, adOpenKeyset, adLockOptimistic
                If TBClientes.EOF = False Then
                    If IsNull(TBClientes!Codigo_pais) = True Or TBClientes!Codigo_pais = "" Then
                        USMsgBox ("Não é permitido exportar esta nota fiscal " & TBproducao!int_NotaFiscal & " - série " & IIf(IsNull(TBproducao!Serie), "", TBproducao!Serie) & ", pois este cliente não tem país cadastrado."), vbExclamation, "CAPRIND v5.0"
                        TBClientes.Close
                        Exit Sub
                    Else
                        Pais = TBClientes!Pais
                        Codigo_pais = TBClientes!Codigo_pais
                    End If
                End If
            ElseIf TBproducao!txt_tipocliente = "JP" Or TBproducao!txt_tipocliente = "JR" Or TBproducao!txt_tipocliente = "FP" Or TBproducao!txt_tipocliente = "FR" Then
                    'Cliente
                    TBClientes.Open "Select * from Clientes where IDcliente = " & TBproducao!Id_Int_Cliente & " and NomeRazao = '" & TBproducao!txt_Razao_Nome & "'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBClientes.EOF = False Then
                        If IsNull(TBClientes!Codigo_pais) = True Or TBClientes!Codigo_pais = "" Then
                            USMsgBox ("Não é permitido exportar esta nota fiscal " & TBproducao!int_NotaFiscal & " - série " & IIf(IsNull(TBproducao!Serie), "", TBproducao!Serie) & ", pois este cliente não tem país cadastrado."), vbExclamation, "CAPRIND v5.0"
                            TBClientes.Close
                            Exit Sub
                        Else
                            Pais = TBClientes!Pais
                            Codigo_pais = TBClientes!Codigo_pais
                        End If
                    End If
                Else
                    'Fornecedor
                    TBClientes.Open "Select * from Compras_fornecedores where IDcliente = " & TBproducao!Id_Int_Cliente & " and Nome_Razao = '" & TBproducao!txt_Razao_Nome & "'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBClientes.EOF = False Then
                        EmailTexto = IIf(IsNull(TBClientes!Email), "", TBClientes!Email)
                        If IsNull(TBClientes!Codigo_pais) = True Or TBClientes!Codigo_pais = "" Then
                            USMsgBox ("Não é permitido exportar esta nota fiscal " & TBproducao!int_NotaFiscal & " - série " & IIf(IsNull(TBproducao!Serie), "", TBproducao!Serie) & ", pois este fornecedor não tem país cadastrado."), vbExclamation, "CAPRIND v5.0"
                            TBClientes.Close
                            Exit Sub
                        Else
                            Pais = TBClientes!Pais
                            Codigo_pais = TBClientes!Codigo_pais
                        End If
                    End If
            End If
            TBClientes.Close
            
            ProcVerifLiberacao
            If Permitido = False Then
                USMsgBox ("Informe o(s) campo(s) da nota fiscal " & TBproducao!int_NotaFiscal & " - série " & IIf(IsNull(TBproducao!Serie), "", TBproducao!Serie) & " antes de exportar: " & vbCrLf & " " & Familiatext & "."), vbInformation, "CAPRIND v5.0"
                Exit Sub
            End If
            TBproducao.MoveNext
            Contador1 = Contador1 + 1
            PBLista.Value = Contador1
        Loop
        
        PBLista.Min = 0
        PBLista.Max = TBproducao.RecordCount
        PBLista.Value = 1
        Contador1 = 0
        TBproducao.MoveFirst
        
        If Opt_NFSe.Value = True Then
            'Nota fiscal de serviços
            Contador2 = 1
            Contador3 = 1
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select * from Empresa where Codigo = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex), Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False Then
Inicio:
                'Formato de codificação UTF-8
                Dim fsT As Object
                'Dim fsTsem As Object
                Set fsT = CreateObject("ADODB.Stream")
                'Set fsTsem = CreateObject("ADODB.Stream")
                
                With fsT
                    .Type = 2 'Specify stream type - we want To save text/string data.
                    .Charset = "utf-8" 'Specify charset For the source text data.
                    .Open 'Open the stream And write binary data To the object
                    Do While TBproducao.EOF = False
                        If Contador3 <= 50 Then
                            .WriteText FunCriarTXTNFSe()
                            TBproducao.MoveNext
                            Contador1 = Contador1 + 1
                            Contador3 = Contador3 + 1
                            PBLista.Value = Contador1
                        Else
                            .SaveToFile Localrel & "\Arquivos exportados\NFSe enviar\" & Format(Date, "ddmmyyyy") & Contador2 & ".txt", 2 'Save binary data To disk
                            .Close
                            Contador3 = 1
                            Contador2 = Contador2 + 1
                            GoTo Inicio
                        End If
                        .WriteText vbCrLf
                    Loop
                    .SaveToFile Localrel & "\Arquivos exportados\NFSe enviar\" & Format(Date, "ddmmyyyy") & Contador2 & ".txt", 2 'Save binary data To disk
                    .Close
                End With
                               
            End If
        ElseIf Opt_NFp.Value = True Then
                'Nota fiscal paulista
                Do While TBproducao.EOF = False
                    OF = TBproducao!int_NotaFiscal
                    Call ProcCriarTXTNFp(IIf(IsNull(TBproducao!Serie), "", TBproducao!Serie), OF, "E")
                    TBproducao.MoveNext
                    Contador1 = Contador1 + 1
                    PBLista.Value = Contador1
                Loop
            ElseIf Opt_sintegra.Value = True Then
                    'Sintegra
                    If Cmb_registro = "54" Then
                        Set ArqTXT = GerArqPastas.CreateTextFile(Localrel & "\Arquivos exportados\Sintegra enviar\Registro 54\54." & Now & ".txt", True)
                    Else
                        Set ArqTXT = GerArqPastas.CreateTextFile(Localrel & "\Arquivos exportados\Sintegra enviar\Registro 75\75.NF " & OF & " - " & Format(TBproducao!dt_DataEmissao, "dd-mm-yy") & " - " & Left(Replace(TBproducao!txt_Razao_Nome, "/", " "), 20) & ".txt", True)
                    End If
                    With ArqTXT
                        Do While TBproducao.EOF = False
                            OF = TBproducao!int_NotaFiscal
                            'Produtos
                            Contador = 1
                            Set TBProduto = CreateObject("adodb.recordset")
                            TBProduto.Open "Select * from tbl_Detalhes_Nota where ID_nota = " & TBproducao!ID, Conexao, adOpenKeyset, adLockOptimistic
                            If TBProduto.EOF = False Then
                                Do While TBProduto.EOF = False
                                    .WriteLine FunCriarTXTSintegra(IIf(IsNull(TBproducao!Serie), "", TBproducao!Serie), OF, Cmb_registro)
                                    TBProduto.MoveNext
                                Loop
                            End If
                            TBProduto.Close
                            TBproducao.MoveNext
                            Contador1 = Contador1 + 1
                            PBLista.Value = Contador1
                        Loop
                        .Close
                    End With
                ElseIf Opt_contimatic.Value = True Then
                        'Contimatic
                        Set TBFI = CreateObject("adodb.recordset")
                        TBFI.Open "Select * from Empresa where Codigo = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex), Conexao, adOpenKeyset, adLockOptimistic
                        If TBFI.EOF = False Then
                            Set ArqTXT = GerArqPastas.CreateTextFile(Localrel & "\Arquivos exportados\Contimatic enviar\" & TBFI!Apelido_contimatic & ".N" & Month(TBproducao!dt_DataEmissao), True)
                            With ArqTXT
                                Do While TBproducao.EOF = False
                                    OF = TBproducao!int_NotaFiscal
                                    .WriteLine FunCriarTXTContimatic(IIf(IsNull(TBproducao!Serie), "", TBproducao!Serie), OF, IDlista)
                                    TBproducao.MoveNext
                                    Contador1 = Contador1 + 1
                                    PBLista.Value = Contador1
                                Loop
                                .Close
                            End With
                        End If
        End If
    Else
        USMsgBox ("Não foi encontrado nenhuma nota fiscal validada para esta pesquisa."), vbExclamation, "CAPRIND v5.0"
        Exit Sub
    End If
    USMsgBox ("Exportação efetuada com sucesso para pasta " & caminho & "."), vbInformation, "CAPRIND v5.0"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Function FunCriarTXTNFSe()
On Error GoTo tratar_erro
Dim Servicoexeccliente As Boolean
Dim CidadeTomador As String, UFTomador As String, TelefoneTomador As String

CpfCnpj = ReturnNumbersOnly(IIf(IsNull(TBFI!CNPJ), "", TBFI!CNPJ))
IM = ReturnNumbersOnly(IIf(IsNull(TBFI!IM), "", TBFI!IM))
NumeroRps = TBproducao!ID
DataRPS = Format(TBproducao!dt_DataEmissao, "yyyymmdd")
DataRPSAnoMes = Format(TBproducao!dt_DataEmissao, "yyyymm")

DescServico = ""
valor = 0
Servicoexeccliente = False
Set TBProduto = CreateObject("adodb.recordset")
StrSql = "Select DNF.ID_EMPRESA, P.Cod_servico_NFSE, NFP.txt_Descricao, NFP.PCCliente, NFP.Retencao_ISSQN, NFP.vlriss, NFP.Servico_cliente, NFP.ISS from tbl_Detalhes_Nota NFP INNER JOIN projproduto P ON P.Desenho = NFP.int_Cod_Produto INNER JOIN tbl_Dados_Nota_Fiscal DNF on NFP.ID_Nota = DNF.ID where NFP.ID_Nota = " & TBproducao!ID & " and DNF.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
'Debug.print StrSql

TBProduto.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    CodServico = TBProduto!Cod_servico_NFSE
    Do While TBProduto.EOF = False
        If IsNull(TBProduto!PCCliente) = False And TBProduto!PCCliente <> "" Then DescServ = Left(TBProduto!Txt_descricao, 70) & " - Ped. " & Trim(Left(TBProduto!PCCliente, 12)) Else DescServ = Left(TBProduto!Txt_descricao, 90)
        If DescServico = "" Then DescServico = DescServ Else DescServico = DescServico & "\s\n" & DescServ
        If TBProduto!Servico_cliente = True Then Servicoexeccliente = True
        
        If TBProduto!Retencao_ISSQN = True Then valor = valor + TBProduto!VlrISS
        ISS_Serv = TBProduto!ISS
        
        TBProduto.MoveNext
    Loop
End If
TBProduto.Close

DescServico = Left(DescServico, 1950)
Set TBTotaisnota = CreateObject("adodb.recordset")
TBTotaisnota.Open "Select Valor_total_aprox_tributos, dbl_Valor_Total_Nota_Serv, Total_INSS_serv, Total_IRRF_serv, Total_CSLL_serv, Total_Cofins_serv, Total_PIS_serv from tbl_Totais_Nota where ID_Nota = " & TBproducao!ID, Conexao, adOpenKeyset, adLockOptimistic
If TBTotaisnota.EOF = False Then
    If DescServico = "" Then DescServico = "Valor tributos aprox.: " & Format(TBTotaisnota!Valor_total_aprox_tributos, "###,##0.00") Else DescServico = DescServico & "\s\nValor tributos aprox.: " & Format(TBTotaisnota!Valor_total_aprox_tributos, "###,##0.00")
    DescServico = Left(DescServico, 2000)
    
    VlrTotalServico = Format(TBTotaisnota!dbl_Valor_Total_Nota_Serv, "0.00")
    VlrTotalDeducoes = 0
End If
TBTotaisnota.Close

'TipoLancamento = "T"
'CidadeServReal = TBFI!Cidade
'UFServReal = TBFI!UF
'TipoPessoa = Left(TBproducao!txt_tipocliente, 1)

'PaisOrigemTomador = ""
CPFCNPJTomador = ""
IMTomador = ""

If CPFCNPJTomador = "" Then CPFCNPJTomador = ReturnNumbersOnly(TBproducao!txt_CNPJ_CPF)
RazaoSocialTomador = Left(TBproducao!txt_Razao_Nome, 100)
CidadeTomador = TBproducao!txt_Municipio
UFTomador = TBproducao!txt_UF

Set TBClientes = CreateObject("adodb.recordset")
TBClientes.Open "Select idTipoEmpresa, Pais, Tipo_bairro, Bairro, Tipo_endereco, Endereco, Complemento, Email, RG_IM from Clientes where IDCliente = " & TBproducao!Id_Int_Cliente & " and NomeRazao = '" & TBproducao!txt_Razao_Nome & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBClientes.EOF = False Then
    If TBClientes!idTipoEmpresa = 0 Then
        'PaisOrigemTomador = TBClientes!Pais
        If Left(TBproducao!txt_tipocliente, 1) = "J" Then CPFCNPJTomador = "99999999999999"
    End If
    If IsNull(TBClientes!RG_IM) = False And (CidadeTomador = "Indaiatuba" Or CidadeTomador = "INDAIATUBA") Then IMTomador = ReturnNumbersOnly(TBClientes!RG_IM)
    
    TipoBairro = TBClientes!Tipo_bairro
    BairroTomador = Left(TBClientes!Bairro, 50)
    TipoEndereco = TBClientes!Tipo_endereco
    EnderecoTomador = Left(TBClientes!Endereco, 50)
    ComplementoEnderecoTomador = IIf(IsNull(TBClientes!complemento), "", TBClientes!complemento)
    EmailTomador = IIf(IsNull(TBClientes!Email), "", Left(TBClientes!Email, 60))
    
    Set TBContas = CreateObject("adodb.recordset")
    TBContas.Open "Select Email from Clientes_Contatos where IDCliente = " & TBproducao!Id_Int_Cliente & " and Enviar_NFe = 'True'", Conexao, adOpenKeyset, adLockOptimistic
    If TBContas.EOF = False Then
        EmailTomador = IIf(IsNull(TBContas!Email), "", Left(Trim(TBContas!Email), 60))
    End If
    TBContas.Close
Else
    Set TBClientes = CreateObject("adodb.recordset")
    TBClientes.Open "Select idTipoEmpresa, Pais, Tipo_bairro, Bairro, Tipo_endereco, Endereco, Complemento, Email, RG_IM from Compras_fornecedores where IDCliente = " & TBproducao!Id_Int_Cliente & " and Nome_Razao = '" & TBproducao!txt_Razao_Nome & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBClientes.EOF = False Then
        If TBClientes!idTipoEmpresa = 0 Then
            'PaisOrigemTomador = TBClientes!Pais
            If Left(TBproducao!txt_tipocliente, 1) = "J" Then CPFCNPJTomador = "99999999999999"
        End If
        If IsNull(TBClientes!RG_IM) = False And (CidadeTomador = "Indaiatuba" Or CidadeTomador = "INDAIATUBA") Then IMTomador = ReturnNumbersOnly(TBClientes!RG_IM)
        
        TipoBairro = TBClientes!Tipo_bairro
        BairroTomador = Left(TBClientes!Bairro, 50)
        TipoEndereco = TBClientes!Tipo_endereco
        EnderecoTomador = Left(TBClientes!Endereco, 50)
        ComplementoEnderecoTomador = IIf(IsNull(TBClientes!complemento), "", TBClientes!complemento)
        
        Set TBContas = CreateObject("adodb.recordset")
        TBContas.Open "Select Email from Contatos_fornecedor where IdFornecedor = " & TBproducao!Id_Int_Cliente & " and Enviar_NFe = 'True'", Conexao, adOpenKeyset, adLockOptimistic
        If TBContas.EOF = False Then
            EmailTomador = IIf(IsNull(TBContas!Email), "", Left(TBContas!Email, 60))
        End If
        TBContas.Close
    End If
End If
TBClientes.Close

Select Case TipoBairro
    Case "Bairro": TipoBairroTomador = "BR"
    Case "Bosque": TipoBairroTomador = ""
    Case "Chácara": TipoBairroTomador = "CH"
    Case "Conjunto": TipoBairroTomador = ""
    Case "Desmembramento": TipoBairroTomador = "DM"
    Case "Distrito": TipoBairroTomador = "DI"
    Case "Favela": TipoBairroTomador = ""
    Case "Fazenda": TipoBairroTomador = ""
    Case "Gleba": TipoBairroTomador = ""
    Case "Horto": TipoBairroTomador = ""
    Case "Jardim": TipoBairroTomador = "JD"
    Case "Loteamento": TipoBairroTomador = "LT"
    Case "Núcleo": TipoBairroTomador = "NR"
    Case "Parque": TipoBairroTomador = "PQ"
    Case "Residencial": TipoBairroTomador = "CR"
    Case "Sítio": TipoBairroTomador = ""
    Case "Tropical": TipoBairroTomador = ""
    Case "Vila": TipoBairroTomador = "VL"
    Case "Zona": TipoBairroTomador = ""
End Select

TipoEnderecoTomador = ""
Select Case TipoEndereco
    Case "Alameda": TipoEnderecoTomador = "AL"
    Case "Avenida": TipoEnderecoTomador = "AV"
    Case "Estrada": TipoEnderecoTomador = "ET"
    Case "Praça": TipoEnderecoTomador = "PC"
    'Case "Rio": TipoEnderecoTomador = ""
    Case "Rodovia": TipoEnderecoTomador = "RD"
    Case "Rua": TipoEnderecoTomador = "R"
    'Case "Ruela": TipoEnderecoTomador = ""
    'Case "Sítio": TipoEnderecoTomador = ""
    'Case "Sub Quadra": TipoEnderecoTomador = ""
    Case "Travessa": TipoEnderecoTomador = "TV"
    'Case "Vale": TipoEnderecoTomador = ""
    'Case "Vargem": TipoEnderecoTomador = ""
    'Case "Via": TipoEnderecoTomador = ""
    'Case "Viaduto": TipoEnderecoTomador = ""
    'Case "Viela": TipoEnderecoTomador = ""
    Case "Vila": TipoEnderecoTomador = "VI"
End Select
If TipoEnderecoTomador = "" Then TipoEnderecoTomador = "R"

If IsNull(TBproducao!Numero) = True And TBproducao!Numero = "" Then NumeroTomador = 0 Else NumeroTomador = Left(ReturnNumbersOnly(TBproducao!Numero), 6)
CEPTomador = Left(ReturnNumbersOnly(TBproducao!Txt_CEP), 8)
TelefoneTomador = Left(IIf(IsNull(TBproducao!txt_Fone_Fax) = False, ReturnNumbersOnly(TBproducao!txt_Fone_Fax), 10), "0000000000")

TextoOBS = ""
Set TBContas = CreateObject("adodb.recordset")
TBContas.Open "Select txt_tipoPagto, txt_Portador_Banco, txt_Agencia, txt_Conta, dt_Vencimento, dbl_Valor, txt_parcela from tbl_Detalhes_Recebimento where ID_Nota = " & TBproducao!ID, Conexao, adOpenKeyset, adLockOptimistic
If TBContas.EOF = False Then
    TextoDadosConta = ""
    If TBContas!txt_tipopagto = "DEPÓSITO" Or TBContas!txt_tipopagto = "DOC" Or TBContas!txt_tipopagto = "TED" Then TextoDadosConta = "\s\nDADOS DA CONTA PARA PAGAMENTO:\s\n" & "BANCO: " & TBContas!txt_Portador_Banco & ", AGÊNCIA: " & TBContas!txt_Agencia & ", C/C: " & TBContas!txt_Conta
    Do While TBContas.EOF = False
        DescOBS = Format(TBContas!dt_Vencimento, "dd/mm/yy") & " R$" & Format(TBContas!dbl_Valor, "###,##0.00") & " " & TBContas!txt_Parcela
        If TextoOBS = "" Then TextoOBS = DescOBS Else TextoOBS = TextoOBS & ", " & DescOBS
        TBContas.MoveNext
    Loop
    If TextoDadosConta <> "" Then TextoOBS = TextoOBS & TextoDadosConta
End If
TBContas.Close

VlrINSS = 0
VlrIR = 0
VlrCSLL = 0
VlrCofins = 0
VlrPIS = 0
VlrOutras = 0
VlrISSRetido = Format(valor, "0.00")
VlrISS = Format(ISS_Serv, "0.0000")

NumeroNFSeSub = 0

ExigibilidadeISS = 1
RegimeEspecialTributacao = 6
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select TipoTributacao, ExigibilidadeISS, RegimeEspecialTributacao from tbl_Dados_Nota_Fiscal_NFe where ID_nota = " & TBproducao!ID, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    If TBAbrir!TipoTributacao = 7 Then
        MunicipioIncidenciaDiferente = "S"
        codigomunicipio = FunVerificaCodMunicipio(CidadeTomador, UFTomador)
    Else
        MunicipioIncidenciaDiferente = "N"
        codigomunicipio = FunVerificaCodMunicipio(TBFI!Cidade, TBFI!UF)
    End If

    ExigibilidadeISS = IIf(IsNull(TBAbrir!ExigibilidadeISS), 1, TBAbrir!ExigibilidadeISS)
    RegimeEspecialTributacao = IIf(IsNull(TBAbrir!RegimeEspecialTributacao), 6, TBAbrir!RegimeEspecialTributacao)
End If
TBAbrir.Close

If Servicoexeccliente = True Then
    CodigoMunicipioServicoRealizado = FunVerificaCodMunicipio(CidadeTomador, UFTomador)
Else
    CodigoMunicipioServicoRealizado = FunVerificaCodMunicipio(TBFI!Cidade, TBFI!UF)
End If

If codigomunicipio = "" Then
codigomunicipio = CodigoMunicipioServicoRealizado
End If


If TBFI!Simples = True Or TBFI!Simples1 = True Then OptanteSimples = 1 Else OptanteSimples = 2

TextoFinal = Contador3 & "|" & CpfCnpj & "|" & IM & "|" & NumeroRps & "|" & DataRPS & "|" & DataRPSAnoMes & "|" & CodServico & "|" & VlrTotalServico & "|" & VlrTotalDeducoes & "|" & ExigibilidadeISS & "||" & codigomunicipio & "|" & CodigoMunicipioServicoRealizado & "||" & OptanteSimples & "|2|" & IIf(valor > 0, 1, 2) & "|" & RegimeEspecialTributacao & "|" & VlrINSS & "|" & VlrIR & "|" & VlrCSLL & "|" & VlrCofins & "|" & VlrPIS & "|" & IIf(valor > 0, VlrISSRetido, "") & "|" & VlrOutras & "|" & IIf(OptanteSimples = 1, VlrISS, "") & "||"
TextoFinal1 = Left(DescServico, 2000) & "|" & Left(TextoOBS, 254) & "|" & CPFCNPJTomador & "|" & IMTomador & "|" & Left(RazaoSocialTomador, 100) & "|" & TipoEnderecoTomador & "|" & Left(EnderecoTomador, 125) & "|" & IIf(NumeroTomador = "", "", NumeroTomador) & "|" & IIf(ComplementoEnderecoTomador = "", "", Left(ComplementoEnderecoTomador, 50)) & "|" & TipoBairroTomador & "|" & BairroTomador & "|" & IIf(CPFCNPJTomador = "99999999999999", "", FunVerificaCodMunicipio(CidadeTomador, UFTomador)) & "|" & UFTomador & "|"
TextoFinal2 = IIf(CPFCNPJTomador = "99999999999999", "", "") & "|" & CEPTomador & "|" & Left(EmailTomador, 60) & "|" & TelefoneTomador & "||||||"
'Debug.print TextoFinal & TextoFinal1 & TextoFinal2

FunCriarTXTNFSe = TextoFinal & TextoFinal1 & TextoFinal2
'TipoLancamento & FunTamanhoTextoVazioDir(CidadeServReal, 50) & UFServReal & TipoPessoa & FunTamanhoTextoVazioDir(PaisOrigemTomador, 50) & FunTamanhoTextoVazioDir(CidadeTomador, 50)
'Familiatext = TextoFinal & TextoFinal1
'FunCriarTXTNFSe = FunTiraAcentosTexto(Familiatext)

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Sub ProcCriarTXTNFp(Serie As String, Nota As Long, status As String)
On Error GoTo tratar_erro

'Verifica status da nota
If TBproducao!Int_status = 2 Then
    Status_NFp = "C"
    motivo = "Cancelamento por motivos de erros de dados"
Else
    Status_NFp = "I"
    motivo = ""
End If

'Serie
Serie_NFp = Replace(TBproducao!Serie, IsNumeric(TBproducao!Serie) = False, "1")
Select Case Len(Serie_NFp)
    Case 1: Serie_NFp = "00" & Serie_NFp
    Case 2: Serie_NFp = "0" & Serie_NFp
    Case Is >= 3: Serie_NFp = Right(Serie_NFp, 3)
End Select

'Arruma data emissão e data saída (data + hora)
Data_Emissao_NFp = Format(TBproducao!dt_DataEmissao, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss")
Data_Saida_NFp = Format(TBproducao!dt_Saida_Entrada, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss")

If TBproducao!int_TipoNota = 1 Then tipo_NFp = 1 Else tipo_NFp = 0 '1 = Saída 2 = Entrada

'Cfop
If Len(TBproducao!int_CFOP) > 5 Then
    If Left(TBproducao!int_CFOP, 5) <> "5.902" And Left(TBproducao!int_CFOP, 5) <> "6.902" Then CFOP_NFp = ReturnNumbersOnly(Left(TBproducao!int_CFOP, 5))
    If Right(TBproducao!int_CFOP, 5) <> "5.902" And Right(TBproducao!int_CFOP, 5) <> "6.902" Then CFOP_NFp = ReturnNumbersOnly(Right(TBproducao!int_CFOP, 5))
Else
    CFOP_NFp = ReturnNumbersOnly(TBproducao!int_CFOP)
End If

'Cidade e UF
If IsNull(TBproducao!txt_UF) = True Or TBproducao!txt_UF = "" Or TBproducao!txt_UF = "EX" Then
    Nome_Cidade_NFp = "EXTERIOR"
    UF_NFp = "EX"
Else
    Nome_Cidade_NFp = FunTiraAcentosTexto(TBproducao!txt_Municipio)
    UF_NFp = TBproducao!txt_UF
End If

'Verifica pais
Set TBClientes = CreateObject("adodb.recordset")
If TBproducao!txt_tipocliente = "E" Then
    'Empresa
    TBClientes.Open "Select * from Empresa where Codigo = " & TBproducao!Id_Int_Cliente, Conexao, adOpenKeyset, adLockOptimistic
ElseIf TBproducao!txt_tipocliente = "JP" Or TBproducao!txt_tipocliente = "JR" Or TBproducao!txt_tipocliente = "FP" Or TBproducao!txt_tipocliente = "FR" Then
        'Cliente
        TBClientes.Open "Select * from Clientes where IDcliente = " & TBproducao!Id_Int_Cliente & " and NomeRazao = '" & TBproducao!txt_Razao_Nome & "'", Conexao, adOpenKeyset, adLockOptimistic
    Else
        'Fornecedor
        TBClientes.Open "Select * from Compras_fornecedores where IDcliente = " & TBproducao!Id_Int_Cliente & " and Nome_Razao = '" & TBproducao!txt_Razao_Nome & "'", Conexao, adOpenKeyset, adLockOptimistic
End If
If TBClientes.EOF = False Then Pais = TBClientes!Pais
TBClientes.Close

Set ArqTXT = GerArqPastas.CreateTextFile(Localrel & "\Arquivos exportados\NFp Enviar\Serie " & Serie & " - Nota " & Nota & " - Status " & status & ".txt", True)
'Gravando no arquivo
With ArqTXT
    Set TBFIltro = CreateObject("adodb.recordset")
    TBFIltro.Open "Select * from Empresa where codigo = " & TBproducao!ID_empresa, Conexao, adOpenKeyset, adLockOptimistic
    If TBFIltro.EOF = False Then
        .WriteLine ("10|1,00|" & ReturnNumbersOnly(TBFIltro!CNPJ) & "|" & Format(TBproducao!dt_DataEmissao, "dd/mm/yyyy") & "|" & Format(TBproducao!dt_DataEmissao, "dd/mm/yyyy"))
        .WriteLine ("20|" & Status_NFp & "|" & motivo & "|" & FunTiraAcentosTexto(TBproducao!txt_NaturezaOP) & "|" & Serie_NFp & "|" & TBproducao!int_NotaFiscal & "|" & Data_Emissao_NFp & "|" & Data_Saida_NFp & "|" & tipo_NFp & "|" & CFOP_NFp & "|" & ReturnNumbersOnly(TBproducao!txt_Inscr_Substituto) & "|" & ReturnNumbersOnly(TBFIltro!IM) & "|" & ReturnNumbersOnly(TBproducao!txt_CNPJ_CPF) & "|" & FunTiraAcentosTexto(TBproducao!txt_Razao_Nome) & "|" & FunTiraAcentosTexto(TBproducao!txt_Endereco) & "|" & TBproducao!Numero & "||" & FunTiraAcentosTexto(TBproducao!txt_Bairro) & "|" & Nome_Cidade_NFp & "|" & UF_NFp & "|" & ReturnNumbersOnly(TBproducao!Txt_CEP) & "|" & Pais & "|" & ReturnNumbersOnly(TBproducao!txt_Fone_Fax) & "|" & ReturnNumbersOnly(TBproducao!txt_IE_Cliente))
    End If
    TBFIltro.Close
    
    If TBproducao!Int_status = 1 Then
        'Produtos e Serviços
        Contador = 0
        Set TBProduto = CreateObject("adodb.recordset")
        TBProduto.Open "Select * from tbl_Detalhes_Nota where ID_nota = " & TBproducao!ID, Conexao, adOpenKeyset, adLockOptimistic
        If TBProduto.EOF = False Then
            Do While TBProduto.EOF = False
                If TBProduto!Tipo = "P" Then
                    .WriteLine ("30|" & FunTiraAcentosTexto(TBProduto!int_Cod_Produto) & "|" & FunTiraAcentosTexto(TBProduto!Txt_descricao) & "||" & FunTiraAcentosTexto(TBProduto!txt_Unid) & "|" & Format(TBProduto!int_Qtd, "0.0000") & "|" & Format(TBProduto!dbl_ValorUnitario, "0.0000") & "|" & Format(TBProduto!dbl_ValorTotal, "0.00") & "|" & TBProduto!txt_CST & "|" & Format(TBProduto!int_ICMS, "0.00") & "|" & Format(TBProduto!int_IPI, "0.00") & "|" & Format(TBProduto!dbl_valoripi, "0.00"))
                Else
                    Valor_ISS_Serv = IIf(IsNull(TBProduto!ISS), 0, TBProduto!ISS)
                    .WriteLine ("30|" & FunTiraAcentosTexto(TBProduto!int_Cod_Produto) & "|" & FunTiraAcentosTexto(TBProduto!Txt_descricao) & "||" & FunTiraAcentosTexto(TBProduto!txt_Unid) & "|" & Format(TBProduto!int_Qtd, "0.0000") & "|" & Format(TBProduto!dbl_ValorUnitario, "0.0000") & "|" & Format(TBProduto!dbl_ValorTotal, "0.00") & "|041|" & "0,00" & "|" & "0,00" & "|" & "0,00")
                End If
                Contador = Contador + 1
                TBProduto.MoveNext
            Loop
        End If
        TBProduto.Close
        
        'Totais da nota
        Set TBTotaisnota = CreateObject("adodb.recordset")
        TBTotaisnota.Open "Select * from tbl_Totais_Nota where ID_nota = " & TBproducao!ID, Conexao, adOpenKeyset, adLockOptimistic
        If TBTotaisnota.EOF = False Then
            Qtde = IIf(IsNull(TBTotaisnota!dbl_Valor_Total_Produtos), 0, TBTotaisnota!dbl_Valor_Total_Produtos) + IIf(IsNull(TBTotaisnota!dbl_Valor_Total_Nota_Serv), 0, TBTotaisnota!dbl_Valor_Total_Nota_Serv)
            If Permitido1 = True Then Texto = Format(TBTotaisnota!dbl_Valor_Total_Nota_Serv, "0.00") & "|" & Format(VlrISS, "0.00") & "|" & Format(TBTotaisnota!dbl_valor_total_iss, "0.00") Else Texto = "|0,00|0,00"
            .WriteLine ("40|" & Format(TBTotaisnota!dbl_Base_ICMS, "0.00") & "|" & Format(TBTotaisnota!dbl_Valor_ICMS, "0.00") & "|" & Format(TBTotaisnota!dbl_Base_ICMS_Subst, "0.00") & "|" & Format(TBTotaisnota!dbl_Valor_ICMS_Subst, "0.00") & "|" & Format(Qtde, "0.00") & "|" & Format(TBTotaisnota!dbl_Valor_Frete, "0.00") & "|" & Format(TBTotaisnota!dbl_Valor_Seguro, "0.00") & "|" & Format(TBTotaisnota!Valor_total_desconto, "0.00") & "|" & Format(TBTotaisnota!dbl_Valor_Total_IPI, "0.00") & "|" & Format(TBTotaisnota!dbl_Desp_Adicionais, "0.00") & "|" & Format(TBTotaisnota!dbl_Valor_Total_Nota, "0.00") & "|" & Texto)
        End If
        TBTotaisnota.Close
        
        'Transportadora
        Set TBTransporte = CreateObject("adodb.recordset")
        TBTransporte.Open "Select * from tbl_Dados_Transp where ID_Nota = " & TBproducao!ID, Conexao, adOpenKeyset, adLockOptimistic
        If TBTransporte.EOF = False Then
            If TBTransporte!txt_Frete_Conta = 1 Then '1 = Emitente
                Transportadora_frete = 0
            Else
                Transportadora_frete = 1 '2 = Destinatário
            End If
            If IsNull(TBTransporte!txt_CNPJ) = False And TBTransporte!txt_CNPJ <> "" Then
                Set TBFornecedor = CreateObject("adodb.recordset")
                TBFornecedor.Open "Select * from Compras_fornecedores where IDCliente = " & TBTransporte!IdIntTransp, Conexao, adOpenKeyset, adLockOptimistic
                If TBFornecedor.EOF = False Then
                    Transportadora_CNPJ = ReturnNumbersOnly(TBTransporte!txt_CNPJ)
                End If
                TBFornecedor.Close
            End If
            .WriteLine ("50|" & Transportadora_frete & "|" & Transportadora_CNPJ & "|" & FunTiraAcentosTexto(TBTransporte!txt_Razao) & "|" & ReturnNumbersOnly(TBTransporte!txt_IE) & "|" & FunTiraAcentosTexto(TBTransporte!txt_Endereco) & "|" & FunTiraAcentosTexto(TBTransporte!txt_Municipio) & "|" & TBTransporte!txt_UF & "|" & TBTransporte!txt_Placa & "|" & TBTransporte!txt_UF_Placa & "|" & TBTransporte!int_Qtd_Transp & "|" & TBTransporte!txt_Especie & "|" & TBTransporte!txt_Marca & "|" & TBTransporte!int_numero & "|" & Format(TBTransporte!dbl_Peso_Liquido, "0.000") & "|" & Format(TBTransporte!dbl_Peso_Bruto, "0.000"))
        End If
        TBTransporte.Close
        
        Permitido = False
        Texto3 = ""
        Set TBContas = CreateObject("adodb.recordset")
        TBContas.Open "Select * from tbl_Detalhes_Recebimento where ID_nota = " & TBproducao!ID & " order by ID", Conexao, adOpenKeyset, adLockOptimistic
        If TBContas.EOF = False Then
            Permitido = True
            Do While TBContas.EOF = False
                If Texto3 = "" Then
                    Texto3 = TBContas!int_NotaFiscal & " - " & TBContas!txt_Parcela & " - " & Format(TBContas!dt_Vencimento, "dd/mm/yyyy") & " - " & Format(TBContas!dbl_Valor, "0.00")
                Else
                    Texto3 = Texto3 & " / " & TBContas!int_NotaFiscal & " - " & TBContas!txt_Parcela & " - " & Format(TBContas!dt_Vencimento, "dd/mm/yyyy") & " - " & Format(TBContas!dbl_Valor, "0.00")
                End If
                TBContas.MoveNext
            Loop
        End If
        TBContas.Close
        
        Set TBControleNF = CreateObject("adodb.recordset")
        TBControleNF.Open "Select * from tbl_DadosAdicionais where ID_nota = " & TBproducao!ID, Conexao, adOpenKeyset, adLockOptimistic
        If TBControleNF.EOF = False Then
            Permitido = True
            DadosAdicionais_NFp = Replace(FunTiraAcentosTexto(Trim(TBControleNF!mem_DadosAdicionais)), "|", "")
            CorpoNota_NFp = Replace(FunTiraAcentosTexto(Trim(TBControleNF!mem_corpo)), "|", "")
            If Texto3 = "" Then
                Texto3 = "|" & DadosAdicionais_NFp & "|" & CorpoNota_NFp
            Else
                Texto3 = Texto3 & "|" & DadosAdicionais_NFp & "|" & CorpoNota_NFp
            End If
        Else
            Texto3 = Texto3 & "||"
        End If
        TBControleNF.Close
        
        If Permitido = True Then
            Texto2 = "00001"
            .WriteLine ("60|" & Texto3)
        Else
            Texto2 = "00000"
        End If
        .WriteLine ("90|00001" & "|" & Format(Contador, "00000") & "|" & "00001|00001" & "|" & Texto2)
    Else
        .WriteLine ("90|00001|00000|00000|00000|00000")
    End If
    .Close
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Function FunCriarTXTSintegra(Serie As String, Nota As Long, Registro As Integer)
On Error GoTo tratar_erro

'Data inicial e data final
If TBproducao!int_TipoNota = 1 Then 'Saída
    Data_Inicial_Sintegra = Format(TBproducao!dt_DataEmissao, "dd/mm/yyyy")
    Data_Final_Sintegra = Format(TBproducao!dt_DataEmissao, "dd/mm/yyyy")
Else
    If IsNull(TBproducao!dt_Saida_Entrada) = False And TBproducao!dt_Saida_Entrada <> "" Then
        Data_Inicial_Sintegra = ReturnNumbersOnly(Format(TBproducao!dt_Saida_Entrada, "dd/mm/yyyy"))
        Data_Final_Sintegra = ReturnNumbersOnly(Format(TBproducao!dt_Saida_Entrada, "dd/mm/yyyy"))
    Else
        Data_Inicial_Sintegra = ReturnNumbersOnly(Format(TBproducao!dt_DataEmissao, "dd/mm/yyyy"))
        Data_Final_Sintegra = ReturnNumbersOnly(Format(TBproducao!dt_DataEmissao, "dd/mm/yyyy"))
    End If
End If

Serie_Sintegra = IIf(IsNull(TBproducao!Serie), "", TBproducao!Serie) 'Serie
Nota_Sintegra = Right(OF, 6) 'Número da nota

'Verifica valores para somar na base de calculo
Frete = IIf(IsNull(TBProduto!Valor_frete), 0, TBProduto!Valor_frete) 'Frete
Seguro = IIf(IsNull(TBProduto!Valor_seguro), 0, TBProduto!Valor_seguro) 'Seguro
Acessorias = IIf(IsNull(TBProduto!Valor_acessorias), 0, TBProduto!Valor_acessorias) 'Acessorias
QuantsolicitadoN2 = IIf(IsNull(TBProduto!Valor_desconto), 0, TBProduto!Valor_desconto) 'Desconto

'CFOP
Set TBFIltro = CreateObject("adodb.recordset")
TBFIltro.Open "Select id_CFOP from tbl_NaturezaOperacao where IDCountCfop = " & TBProduto!ID_CFOP, Conexao, adOpenKeyset, adLockOptimistic
If TBFIltro.EOF = False Then
    If Len(TBFIltro!ID_CFOP) > 5 Then
        If TBProduto!retorno = True Then
            If Left(TBFIltro!ID_CFOP, 5) = "5.902" Or Left(TBFIltro!ID_CFOP, 5) = "6.902" Or Left(TBFIltro!ID_CFOP, 5) = "5.916" Or Left(TBFIltro!ID_CFOP, 5) = "6.916" Or Left(TBFIltro!ID_CFOP, 5) = "5.925" Or Left(TBFIltro!ID_CFOP, 5) = "6.925" Then
                Familiatext = Left(TBFIltro!ID_CFOP, 5)
            ElseIf Right(TBFIltro!ID_CFOP, 5) = "5.902" Or Right(TBFIltro!ID_CFOP, 5) = "6.902" Or Right(TBFIltro!ID_CFOP, 5) = "5.916" Or Right(TBFIltro!ID_CFOP, 5) = "6.916" Or Right(TBFIltro!ID_CFOP, 5) = "5.925" Or Right(TBFIltro!ID_CFOP, 5) = "6.925" Then
                    Familiatext = Right(TBFIltro!ID_CFOP, 5)
            End If
        Else
            If Left(TBFIltro!ID_CFOP, 5) <> "5.902" And Left(TBFIltro!ID_CFOP, 5) <> "6.902" And Left(TBFIltro!ID_CFOP, 5) <> "5.916" And Left(TBFIltro!ID_CFOP, 5) <> "6.916" And Left(TBFIltro!ID_CFOP, 5) <> "5.925" And Left(TBFIltro!ID_CFOP, 5) <> "6.925" Then
                Familiatext = Left(TBFIltro!ID_CFOP, 5)
            ElseIf Right(TBFIltro!ID_CFOP, 5) <> "5.902" And Right(TBFIltro!ID_CFOP, 5) <> "6.902" And Right(TBFIltro!ID_CFOP, 5) <> "5.916" And Right(TBFIltro!ID_CFOP, 5) <> "6.916" And Right(TBFIltro!ID_CFOP, 5) <> "5.925" And Right(TBFIltro!ID_CFOP, 5) <> "6.925" Then
                    Familiatext = Right(TBFIltro!ID_CFOP, 5)
            End If
        End If
    Else
        Familiatext = TBFIltro!ID_CFOP
    End If
End If
CFOP_Sintegra = ReturnNumbersOnly(Familiatext)

Cod_Produto_Sintegra = TBProduto!int_Cod_Produto 'Código interno
Set TBFIltro = CreateObject("adodb.recordset")
TBFIltro.Open "Select * from tbl_ClassificacaoFiscal where Idclass = " & TBProduto!ID_CF, Conexao, adOpenKeyset, adLockOptimistic
If TBFIltro.EOF = False Then
    CF_Sintegra = ReturnNumbersOnly(TBFIltro!IDIntClasse) 'Classificação Fiscal (Cód. NCM)
End If
Descricao_Sintegra = Left(TBProduto!Txt_descricao, 53) 'Descricao
UN_Sintegra = TBProduto!txt_Unid 'Unidade
Qtde_Sintegra = ReturnNumbersOnly(Format(TBProduto!int_Qtd, "###,##0.0000")) 'Quantidade
Vlr_Total_Sintegra = ReturnNumbersOnly(Format(TBProduto!dbl_ValorTotal, "###,##0.00")) 'Valor total
Vlr_Desconto_Sintegra = ReturnNumbersOnly(IIf(IsNull(TBProduto!Valor_desconto), 0, Format(TBProduto!Valor_desconto, "###,##0.00"))) 'Valor do desconto
VlrIPI = IIf(IsNull(TBProduto!dbl_valoripi), 0, TBProduto!dbl_valoripi) 'Valor do IPI
            
Nome = IIf(IsNull(TBProduto!Txt_CF), "", TBProduto!Txt_CF)
ProcControleImposto IIf(IsNull(TBProduto!ID_CFOP), 0, TBProduto!ID_CFOP), IIf(IsNull(TBproducao!Id_Int_Cliente), 0, TBproducao!Id_Int_Cliente)
If IsNull(TBProduto!Txt_CF) = False Then
    ProcBuscaTributos (TBProduto!Txt_CF)
    If TBproducao!Aplicacao = "T" Then
        If vlrICMS(0, 1) = 0 Then vlrICMS(0, 1) = IIf(IsNull(TBProduto!int_IPI), 0, TBProduto!int_IPI)
        If vlrICMS(0, 2) = 0 Then vlrICMS(0, 2) = IIf(IsNull(TBProduto!int_ICMS), 0, TBProduto!int_ICMS)
        If vlrICMS(0, 3) = 0 Then vlrICMS(0, 3) = IIf(IsNull(TBProduto!int_ICMS), 0, TBProduto!int_ICMS)
        If vlrICMS(0, 4) = 0 Then vlrICMS(0, 4) = IIf(IsNull(TBProduto!int_ICMS), 0, TBProduto!int_ICMS)
    End If
    ProcVerificaRegiao TBproducao!txt_UF, IIf(IsNull(TBproducao!Id_Int_Cliente), 0, TBproducao!Id_Int_Cliente), IIf(IsNull(TBproducao!txt_Razao_Nome), "", TBproducao!txt_Razao_Nome)
End If

'Base de cálculo do ICMS
PV1 = 0
ICMS = 0
If TemICMS = "SIM" Then
    ICMS = IIf(IsNull(TBProduto!int_ICMS), 0, TBProduto!int_ICMS)
    
    valor = 0
    Set TBFIltro = CreateObject("adodb.recordset")
    TBFIltro.Open "Select * from tbl_DadosAdicionais where ID_nota = " & TBproducao!ID, Conexao, adOpenKeyset, adLockOptimistic
    If TBFIltro.EOF = False Then
        valor = IIf(IsNull(TBFIltro!VlrMPA), 0, TBFIltro!VlrMPA)
    End If
    TBFIltro.Close
    
    If ICMS <> 0 Or valor <> 0 Then
        PV = IIf(IsNull(TBProduto!dbl_ValorTotal), 0, TBProduto!dbl_ValorTotal)
        PV = (PV - QuantsolicitadoN2) + Frete
        
        IntICMS = IIf(IsNull(TBProduto!int_ICMS), 0, TBProduto!int_ICMS)
        
        ProcCalculaBC TBproducao!ID_empresa, Familiatext, valor, PV, VlrIPI, SomarIPI, SomarIPIST, TemReducaoBC, False, IIf(IsNull(TBProduto!txt_CST), "", TBProduto!txt_CST), TBproducao!Aplicacao, TBproducao!Id_Int_Cliente, TBproducao!txt_Razao_Nome
        PV1 = BC + Seguro + Acessorias
    End If
End If
FamiliaAntiga = PV1
Vlr_BC_Sintegra = ReturnNumbersOnly(Format(FamiliaAntiga, "###,##0.00"))

'Base de cálculo do ICMS com subst. tributária
BCICMSCST = 0
Set TBMaquinas = CreateObject("adodb.recordset")
TBMaquinas.Open "Select * from regioes where uf = '" & TBproducao!txt_UF & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBMaquinas.EOF = False Then
    Set TBFI = CreateObject("adodb.recordset")
    TBFI.Open "Select Margem, Aliquota from CST where Id_UF = " & TBMaquinas!ID & " and CST = '" & TBProduto!txt_CST & "' and CF = '" & TBProduto!Txt_CF & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBFI.EOF = False Then
        If TBProduto!txt_CST = "010" Or TBProduto!txt_CST = "020" Or TBProduto!txt_CST = "051" Or TBProduto!txt_CST = "060" Or TBProduto!txt_CST = "070" Or TBProduto!txt_CST = "0201" Or TBProduto!txt_CST = "0202" Or TBProduto!txt_CST = "0203" Or TBProduto!txt_CST = "0900" Then
            If TBproducao!txt_UF = "MT" Then
                PV = IIf(IsNull(TBProduto!dbl_ValorUnitario), 0, TBProduto!dbl_ValorUnitario) * IIf(IsNull(TBProduto!int_Qtd), 0, TBProduto!int_Qtd)
                PV = PV - QuantsolicitadoN2
                ProcCalculaBC TBproducao!ID_empresa, Familiatext, valor, PV, IIf(IsNull(TBProduto!dbl_valoripi), 0, TBProduto!dbl_valoripi), SomarIPI, SomarIPIST, TemReducaoBC, False, IIf(IsNull(TBProduto!txt_CST), "", TBProduto!txt_CST), TBproducao!Aplicacao, TBproducao!Id_Int_Cliente, TBproducao!txt_Razao_Nome
                
                'A (Valor total da operação)
                ValorConta = BC + Frete + Seguro + Acessorias
                
                'B (Alíquota do ICMS normal)
                ICMS = vRegiao(0, 1)
                
                'C (Calcula valor de ICMS normal)
                TTICMS = (ValorConta * ICMS) / 100
                TTICMS = Format(TTICMS, "###,##0.00")
                
                'D (Alíquota do ICMS ST)
                ICMSST = IIf(IsNull(TBFI!Aliquota), 0, TBFI!Aliquota)
                
                'E (Margem do ICMS ST)
                QtdeSaida = IIf(IsNull(TBFI!Margem), 0, TBFI!Margem)
                
                'F (Valor do ICMS ST)
                ICMSCST = Format((ValorConta * QtdeSaida) / 100, "###,##0.00")
                
                'G (Base do ICMS ST)
                BCICMSCST = Format(((TTICMS + ICMSCST) / ICMSST) * 100, "###,##0.00")
            Else
                PV = IIf(IsNull(TBProduto!dbl_ValorUnitario), 0, TBProduto!dbl_ValorUnitario) * IIf(IsNull(TBProduto!int_Qtd), 0, TBProduto!int_Qtd)
                PV = PV - QuantsolicitadoN2
                ProcCalculaBC TBproducao!ID_empresa, Familiatext, valor, PV, IIf(IsNull(TBProduto!dbl_valoripi), 0, TBProduto!dbl_valoripi), SomarIPI, SomarIPIST, TemReducaoBC, False, IIf(IsNull(TBProduto!txt_CST), "", TBProduto!txt_CST), TBproducao!Aplicacao, TBproducao!Id_Int_Cliente, TBproducao!txt_Razao_Nome
                                        
                '(((Base de ICMS Reduzida + IPI) * Margem) +  (Valor total do produto + IPI))
                ValorConta = BC + Frete + Seguro + Acessorias
                QtdeSaida = IIf(IsNull(TBFI!Margem), 0, TBFI!Margem)
                BCICMSCST = ((ValorConta * QtdeSaida) / 100) + ValorConta
            End If
        End If
    End If
    TBFI.Close
End If
TBMaquinas.Close
FamiliaAntiga = BCICMSCST
Vlr_BC_Sintegra_ICMS_Sub = ReturnNumbersOnly(Format(FamiliaAntiga, "###,##0.00"))

IPI_Sintegra = ReturnNumbersOnly(Format(TBProduto!int_IPI, "###,##0.00")) 'IPI

'Valor do IPI
FamiliaAntiga = VlrIPI
Vlr_IPI_Sintegra = ReturnNumbersOnly(Format(FamiliaAntiga, "###,##0.00"))

ICMS_Sintegra = ReturnNumbersOnly(Format(TBProduto!int_ICMS, "###,##0.00")) 'ICMS

'Redução na base de cálculo do ICMS
Set TBMaquinas = CreateObject("adodb.recordset")
TBMaquinas.Open "Select * from regioes where uf = '" & TBproducao!txt_UF & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBMaquinas.EOF = False Then
    Set TBFI = CreateObject("adodb.recordset")
    TBFI.Open "Select * from tbl_ClassificacaoFiscal where txt_Ref = '" & TBProduto!Txt_CF & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBFI.EOF = False Then
        Select Case TBMaquinas!regiao
            Case "DE": Reducao_BC_Sintegra = ReturnNumbersOnly(Format(TBFI!CTDE, "###,##0.00"))
            Case "SS": Reducao_BC_Sintegra = ReturnNumbersOnly(Format(TBFI!CTSS, "###,##0.00"))
            Case "NN": Reducao_BC_Sintegra = ReturnNumbersOnly(Format(TBFI!CTNN, "###,##0.00"))
            Case "CO": Reducao_BC_Sintegra = ReturnNumbersOnly(Format(TBFI!CTCO, "###,##0.00"))
        End Select
    End If
    TBFI.Close
End If
TBMaquinas.Close

If Registro = 54 Then
    FunCriarTXTSintegra = ("54" & ReturnNumbersOnly(TBproducao!txt_CNPJ_CPF) & "01" & FunTamanhoTextoVazioDir(Serie_Sintegra, 3) & FunTamanhoTextoZeroEsq(Nota_Sintegra, 6) & CFOP_Sintegra & TBProduto!txt_CST & FunTamanhoTextoZeroEsq(Contador, 3) & FunTamanhoTextoVazioDir(Cod_Produto_Sintegra, 14) & FunTamanhoTextoZeroEsq(Qtde_Sintegra, 11) & FunTamanhoTextoZeroEsq(Vlr_Total_Sintegra, 12) & FunTamanhoTextoZeroEsq(Vlr_Desconto_Sintegra, 12) & FunTamanhoTextoZeroEsq(Vlr_BC_Sintegra, 12) & FunTamanhoTextoZeroEsq(Vlr_BC_Sintegra_ICMS_Sub, 12) & FunTamanhoTextoZeroEsq(Vlr_IPI_Sintegra, 12) & FunTamanhoTextoZeroEsq(ICMS_Sintegra, 4))
    '.WriteLine ("54" & ReturnNumbersOnly(TBproducao!txt_CNPJ_CPF) & "01" & FunTamanhoTextoVazioDir(Serie_Sintegra, 3) & FunTamanhoTextoZeroEsq(Nota_Sintegra, 6) & CFOP_Sintegra & TBProduto!txt_CST & FunTamanhoTextoZeroEsq(Contador, 3) & FunTamanhoTextoVazioDir(Cod_Produto_Sintegra, 14) & FunTamanhoTextoZeroEsq(Qtde_Sintegra, 11) & FunTamanhoTextoZeroEsq(Vlr_Total_Sintegra, 12) & FunTamanhoTextoZeroEsq(Vlr_Desconto_Sintegra, 12) & FunTamanhoTextoZeroEsq(Vlr_BC_Sintegra, 12) & FunTamanhoTextoZeroEsq(Vlr_BC_Sintegra_ICMS_Sub, 12) & FunTamanhoTextoZeroEsq(Vlr_IPI_Sintegra, 12) & FunTamanhoTextoZeroEsq(ICMS_Sintegra, 4))
Else
    FunCriarTXTSintegra = ("75" & Data_Inicial_Sintegra & Data_Final_Sintegra & FunTamanhoTextoVazioDir(Cod_Produto_Sintegra, 14) & CF_Sintegra & FunTamanhoTextoVazioDir(Descricao_Sintegra, 53) & FunTamanhoTextoVazioDir(UN_Sintegra, 6) & FunTamanhoTextoZeroEsq(IPI_Sintegra, 5) & FunTamanhoTextoZeroEsq(ICMS_Sintegra, 4) & FunTamanhoTextoZeroEsq(Reducao_BC_Sintegra, 5) & FunTamanhoTextoZeroEsq(Vlr_BC_Sintegra, 13))
    '.WriteLine ("75" & Data_Inicial_Sintegra & Data_Final_Sintegra & FunTamanhoTextoVazioDir(Cod_Produto_Sintegra, 14) & CF_Sintegra & FunTamanhoTextoVazioDir(Descricao_Sintegra, 53) & FunTamanhoTextoVazioDir(UN_Sintegra, 6) & FunTamanhoTextoZeroEsq(IPI_Sintegra, 5) & FunTamanhoTextoZeroEsq(ICMS_Sintegra, 4) & FunTamanhoTextoZeroEsq(Reducao_BC_Sintegra, 5) & FunTamanhoTextoZeroEsq(Vlr_BC_Sintegra, 13))
End If

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Function FunCriarTXTContimatic(Serie As String, Nota As Long, UltimaNota As Long)
On Error GoTo tratar_erro

If TBproducao!int_TipoNota = 1 Then Tipo_NF_Contimatic = "S" Else Tipo_NF_Contimatic = "E" 'Tipo da nota
Data_Emissao_Contimatic = Format(TBproducao!dt_DataEmissao, "ddmm") 'Data de emissão
Data_Circulacao_Contimatic = Format(TBproducao!dt_Saida_Entrada, "ddmm") 'Data de circulacao
Especie_Contimatic = "NFF  " 'Espécie
Serie_Contimatic = IIf(IsNull(TBproducao!Serie), "", TBproducao!Serie) 'Série
Nota_Contimatic = Right(Nota, 6) 'NF
Ultima_Nota_Contimatic = Right(UltimaNota, 6) 'NF

'Uf do emitente
Set TBCFOP = CreateObject("adodb.recordset")
TBCFOP.Open "select * from empresa where codigo = " & TBproducao!ID_empresa, Conexao, adOpenKeyset, adLockOptimistic
If TBCFOP.EOF = False Then
    UF_emitente_Contimatic = IIf(IsNull(TBCFOP!UF), "", TBCFOP!UF)
End If

'CFOP
Set TBCFOP = CreateObject("adodb.recordset")
TBCFOP.Open "select * from tbl_NaturezaOperacao where IDCountCFOP = " & TBproducao!ID_CFOP, Conexao, adOpenKeyset, adLockOptimistic
If TBCFOP.EOF = False Then
    CFOP_Contimatic = IIf(IsNull(TBCFOP!ID_CFOP), "", TBCFOP!ID_CFOP)
End If

'Código DIPAM e municipio
FamiliaAntiga = FunTiraAcentosTexto(TBproducao!Cidade)
Cod_Cidade_Contimatic = FunVerificaCodMunicipio(FamiliaAntiga, TBproducao!UF)
Cod_Cidade_DIPAM_Contimatic = FunVerificaCodMunicipioDIPAM(FamiliaAntiga, TBproducao!UF)

'Valor contábil (total da NF), Valor da base ST e Valor de desconto
Set TBCFOP = CreateObject("adodb.recordset")
TBCFOP.Open "select * from tbl_totais_nota where ID_nota = " & TBproducao!ID, Conexao, adOpenKeyset, adLockOptimistic
If TBCFOP.EOF = False Then
    VlrTotal_Nota_Contimatic = IIf(IsNull(TBCFOP!dbl_Valor_Total_Nota), 0, TBCFOP!dbl_Valor_Total_Nota)
    VlrTotal_IPI_Contimatic = IIf(IsNull(TBCFOP!dbl_Valor_Total_IPI), 0, TBCFOP!dbl_Valor_Total_IPI)
    If VlrTotal_IPI_Contimatic <> 0 Then
        VlrBase_IPI_Contimatic = IIf(IsNull(TBCFOP!dbl_Valor_Total_Produtos), 0, TBCFOP!dbl_Valor_Total_Produtos) + IIf(IsNull(TBCFOP!dbl_Valor_Frete), 0, TBCFOP!dbl_Valor_Frete)
    Else
        VlrBase_IPI_Contimatic = 0
    End If
    VlrBase_ICMSST_Contimatic = IIf(IsNull(TBCFOP!dbl_Base_ICMS_Subst), 0, TBCFOP!dbl_Base_ICMS_Subst)
    VlrTotal_ICMSST_Contimatic = IIf(IsNull(TBCFOP!dbl_Valor_ICMS_Subst), 0, TBCFOP!dbl_Valor_ICMS_Subst)
    VlrDesconto_Contimatic = IIf(IsNull(TBCFOP!Valor_total_desconto), 0, TBCFOP!Valor_total_desconto)
End If
            
'ICMS
Contador = 1
AliquotaICMS1_Contimatic = 0
VlrBase_ICMS1_Contimatic = 0
VlrICMS1_Contimatic = 0
VlrBase_ICMS_Isento1_Contimatic = 0
VlrBase_ICMS_Outros1_Contimatic = 0
Tipo_NF1_Contimatic = 0

AliquotaICMS2_Contimatic = 0
VlrBase_ICMS2_Contimatic = 0
VlrICMS2_Contimatic = 0
VlrBase_ICMS_Isento2_Contimatic = 0
VlrBase_ICMS_Outros2_Contimatic = 0
Tipo_NF2_Contimatic = 0

AliquotaICMS3_Contimatic = 0
VlrBase_ICMS3_Contimatic = 0
VlrICMS3_Contimatic = 0
VlrBase_ICMS_Isento3_Contimatic = 0
VlrBase_ICMS_Outros3_Contimatic = 0
Tipo_NF3_Contimatic = 0

AliquotaICMS4_Contimatic = 0
VlrBase_ICMS4_Contimatic = 0
VlrICMS4_Contimatic = 0
VlrBase_ICMS_Isento4_Contimatic = 0
VlrBase_ICMS_Outros4_Contimatic = 0
Tipo_NF4_Contimatic = 0

AliquotaICMS5_Contimatic = 0
VlrBase_ICMS5_Contimatic = 0
VlrICMS5_Contimatic = 0
VlrBase_ICMS_Isento5_Contimatic = 0
VlrBase_ICMS_Outros5_Contimatic = 0
Tipo_NF5_Contimatic = 0
Set TBCFOP = CreateObject("adodb.recordset")
TBCFOP.Open "Select int_ICMS from tbl_detalhes_nota where ID_nota = " & TBproducao!ID & " Group by int_icms", Conexao, adOpenKeyset, adLockOptimistic
If TBCFOP.EOF = False Then
    Do While TBCFOP.EOF = False
        Qtd = IIf(IsNull(TBCFOP!int_ICMS), 0, TBCFOP!int_ICMS)
        
        valor = 0
        Valor1 = 0
        Valor2 = 0
        Valor3 = 0
        Set TBProduto = CreateObject("adodb.recordset")
        TBProduto.Open "Select Sum(dbl_ValorTotal) as VlrBase_ICMS, Sum(dbl_ValorIPI) as ValorIPI, txt_CST, ID_CFOP from tbl_detalhes_nota where ID_nota = " & TBproducao!ID & " and int_ICMS = " & TBCFOP!int_ICMS & " and txt_CST is not null Group by txt_CST, ID_CFOP", Conexao, adOpenKeyset, adLockOptimistic
        If TBProduto.EOF = False Then
            Do While TBProduto.EOF = False
                If Len(TBProduto!txt_CST) = 3 Then Contador2 = 2 Else Contador2 = 3
                CST = Right(TBProduto!txt_CST, Contador2)
                
                Set TBFI = CreateObject("adodb.recordset")
                TBFI.Open "Select * from tbl_NaturezaOperacao where IDCountCFOP = " & TBProduto!ID_CFOP & " and txt_Somar = 'SIM'", Conexao, adOpenKeyset, adLockOptimistic
                
                If Contador = 1 Then
                    If CST = "00" Or CST = "10" Or CST = "20" Or CST = "101" Or CST = "102" Or CST = "201" Or CST = "202" Then
                        'Alíquota do ICMS
                        AliquotaICMS1_Contimatic = Qtd
                        If TBFI.EOF = False Then
                            'Valor da BC do ICMS com IPI
                            valor = valor + IIf(IsNull(TBProduto!VlrBase_ICMS), 0, TBProduto!VlrBase_ICMS) + IIf(IsNull(TBProduto!ValorIPI), 0, TBProduto!ValorIPI)
                            'Valor do ICMS com IPI
                            Valor1 = Valor1 + ((IIf(IsNull(TBProduto!VlrBase_ICMS), 0, TBProduto!VlrBase_ICMS) + IIf(IsNull(TBProduto!ValorIPI), 0, TBProduto!ValorIPI)) * Qtd) / 100
                        Else
                            'Valor da BC do ICMS
                            valor = valor + IIf(IsNull(TBProduto!VlrBase_ICMS), 0, TBProduto!VlrBase_ICMS)
                            'Valor do ICMS
                            Valor1 = Valor1 + (IIf(IsNull(TBProduto!VlrBase_ICMS), 0, TBProduto!VlrBase_ICMS) * Qtd) / 100
                        End If
                        VlrBase_ICMS1_Contimatic = valor
                        VlrICMS1_Contimatic = Valor1
                        'Tipo da NF
                        If CST = "00" Or CST = "101" Or CST = "102" Then
                            Tipo_NF1_Contimatic = 0
                        ElseIf CST = "10" Or CST = "201" Or CST = "202" Then
                                Tipo_NF1_Contimatic = 13
                            Else
                                Tipo_NF1_Contimatic = 16
                        End If
                        
                    ElseIf CST = "30" Or CST = "40" Or CST = "103" Or CST = "203" Then
                            'Valor da BC do ICMS Isento
                            Valor2 = Valor2 + IIf(IsNull(TBProduto!VlrBase_ICMS), 0, TBProduto!VlrBase_ICMS)
                            VlrBase_ICMS_Isento1_Contimatic = Valor2
                        Else
                            'Valor da BC do ICMS Outros
                            Valor3 = Valor3 + IIf(IsNull(TBProduto!VlrBase_ICMS), 0, TBProduto!VlrBase_ICMS)
                            VlrBase_ICMS_Outros1_Contimatic = Valor3
                    End If
                    
                    'Tipo da NF
                    If TBFI.EOF = False Then Tipo_NF1_Contimatic = 2
                    TBFI.Close
                    
                    If TBproducao!Int_status = 2 Then Tipo_NF1_Contimatic = 99
                ElseIf Contador = 2 Then
                        If CST = "00" Or CST = "10" Or CST = "20" Or CST = "101" Or CST = "102" Or CST = "201" Or CST = "202" Then
                            'Alíquota do ICMS
                            AliquotaICMS2_Contimatic = Qtd
                            If TBFI.EOF = False Then
                                'Valor da BC do ICMS com IPI
                                valor = valor + IIf(IsNull(TBProduto!VlrBase_ICMS), 0, TBProduto!VlrBase_ICMS) + IIf(IsNull(TBProduto!ValorIPI), 0, TBProduto!ValorIPI)
                                'Valor do ICMS com IPI
                                Valor1 = Valor1 + ((IIf(IsNull(TBProduto!VlrBase_ICMS), 0, TBProduto!VlrBase_ICMS) + IIf(IsNull(TBProduto!ValorIPI), 0, TBProduto!ValorIPI)) * Qtd) / 100
                            Else
                                'Valor da BC do ICMS
                                valor = valor + IIf(IsNull(TBProduto!VlrBase_ICMS), 0, TBProduto!VlrBase_ICMS)
                                'Valor do ICMS
                                Valor1 = Valor1 + (IIf(IsNull(TBProduto!VlrBase_ICMS), 0, TBProduto!VlrBase_ICMS) * Qtd) / 100
                            End If
                            VlrBase_ICMS2_Contimatic = valor
                            VlrICMS2_Contimatic = Valor1
                            'Tipo da NF
                            If CST = "00" Or CST = "101" Or CST = "102" Then
                                Tipo_NF2_Contimatic = 0
                            ElseIf CST = "10" Or CST = "201" Or CST = "202" Then
                                    Tipo_NF2_Contimatic = 13
                                Else
                                    Tipo_NF2_Contimatic = 16
                            End If
                            
                        ElseIf CST = "30" Or CST = "40" Or CST = "103" Or CST = "203" Then
                                'Valor da BC do ICMS Isento
                                Valor2 = Valor2 + IIf(IsNull(TBProduto!VlrBase_ICMS), 0, TBProduto!VlrBase_ICMS)
                                VlrBase_ICMS_Isento2_Contimatic = Valor2
                            Else
                                'Valor da BC do ICMS Outros
                                Valor3 = Valor3 + IIf(IsNull(TBProduto!VlrBase_ICMS), 0, TBProduto!VlrBase_ICMS)
                                VlrBase_ICMS_Outros2_Contimatic = Valor3
                        End If
                        
                        'Tipo da NF
                        If TBFI.EOF = False Then Tipo_NF2_Contimatic = 2
                        TBFI.Close
                        
                        If TBproducao!Int_status = 2 Then Tipo_NF2_Contimatic = 99
                    ElseIf Contador = 3 Then
                            If CST = "00" Or CST = "10" Or CST = "20" Or CST = "101" Or CST = "102" Or CST = "201" Or CST = "202" Then
                                'Alíquota do ICMS
                                AliquotaICMS3_Contimatic = Qtd
                                If TBFI.EOF = False Then
                                    'Valor da BC do ICMS com IPI
                                    valor = valor + IIf(IsNull(TBProduto!VlrBase_ICMS), 0, TBProduto!VlrBase_ICMS) + IIf(IsNull(TBProduto!ValorIPI), 0, TBProduto!ValorIPI)
                                    'Valor do ICMS com IPI
                                    Valor1 = Valor1 + ((IIf(IsNull(TBProduto!VlrBase_ICMS), 0, TBProduto!VlrBase_ICMS) + IIf(IsNull(TBProduto!ValorIPI), 0, TBProduto!ValorIPI)) * Qtd) / 100
                                Else
                                    'Valor da BC do ICMS
                                    valor = valor + IIf(IsNull(TBProduto!VlrBase_ICMS), 0, TBProduto!VlrBase_ICMS)
                                    'Valor do ICMS
                                    Valor1 = Valor1 + (IIf(IsNull(TBProduto!VlrBase_ICMS), 0, TBProduto!VlrBase_ICMS) * Qtd) / 100
                                End If
                                VlrBase_ICMS3_Contimatic = valor
                                VlrICMS3_Contimatic = Valor1
                                'Tipo da NF
                                If CST = "00" Or CST = "101" Or CST = "102" Then
                                    Tipo_NF3_Contimatic = 0
                                ElseIf CST = "10" Or CST = "201" Or CST = "202" Then
                                        Tipo_NF3_Contimatic = 13
                                    Else
                                        Tipo_NF3_Contimatic = 16
                                End If
                                
                            ElseIf CST = "30" Or CST = "40" Or CST = "103" Or CST = "203" Then
                                    'Valor da BC do ICMS Isento
                                    Valor2 = Valor2 + IIf(IsNull(TBProduto!VlrBase_ICMS), 0, TBProduto!VlrBase_ICMS)
                                    VlrBase_ICMS_Isento3_Contimatic = Valor2
                                Else
                                    'Valor da BC do ICMS Outros
                                    Valor3 = Valor3 + IIf(IsNull(TBProduto!VlrBase_ICMS), 0, TBProduto!VlrBase_ICMS)
                                    VlrBase_ICMS_Outros3_Contimatic = Valor3
                            End If
                            
                            'Tipo da NF
                            If TBFI.EOF = False Then Tipo_NF3_Contimatic = 2
                            TBFI.Close
                            
                            If TBproducao!Int_status = 2 Then Tipo_NF3_Contimatic = 99
                        ElseIf Contador = 4 Then
                                If CST = "00" Or CST = "10" Or CST = "20" Or CST = "101" Or CST = "102" Or CST = "201" Or CST = "202" Then
                                    'Alíquota do ICMS
                                    AliquotaICMS4_Contimatic = Qtd
                                    If TBFI.EOF = False Then
                                        'Valor da BC do ICMS com IPI
                                        valor = valor + IIf(IsNull(TBProduto!VlrBase_ICMS), 0, TBProduto!VlrBase_ICMS) + IIf(IsNull(TBProduto!ValorIPI), 0, TBProduto!ValorIPI)
                                        'Valor do ICMS com IPI
                                        Valor1 = Valor1 + ((IIf(IsNull(TBProduto!VlrBase_ICMS), 0, TBProduto!VlrBase_ICMS) + IIf(IsNull(TBProduto!ValorIPI), 0, TBProduto!ValorIPI)) * Qtd) / 100
                                    Else
                                        'Valor da BC do ICMS
                                        valor = valor + IIf(IsNull(TBProduto!VlrBase_ICMS), 0, TBProduto!VlrBase_ICMS)
                                        'Valor do ICMS
                                        Valor1 = Valor1 + (IIf(IsNull(TBProduto!VlrBase_ICMS), 0, TBProduto!VlrBase_ICMS) * Qtd) / 100
                                    End If
                                    VlrBase_ICMS4_Contimatic = valor
                                    VlrICMS4_Contimatic = Valor1
                                    'Tipo da NF
                                    If CST = "00" Or CST = "101" Or CST = "102" Then
                                        Tipo_NF4_Contimatic = 0
                                    ElseIf CST = "10" Or CST = "201" Or CST = "202" Then
                                            Tipo_NF4_Contimatic = 13
                                        Else
                                            Tipo_NF4_Contimatic = 16
                                    End If
                                    
                                ElseIf CST = "30" Or CST = "40" Or CST = "103" Or CST = "203" Then
                                        'Valor da BC do ICMS Isento
                                        Valor2 = Valor2 + IIf(IsNull(TBProduto!VlrBase_ICMS), 0, TBProduto!VlrBase_ICMS)
                                        VlrBase_ICMS_Isento4_Contimatic = Valor2
                                    Else
                                        'Valor da BC do ICMS Outros
                                        Valor3 = Valor3 + IIf(IsNull(TBProduto!VlrBase_ICMS), 0, TBProduto!VlrBase_ICMS)
                                        VlrBase_ICMS_Outros4_Contimatic = Valor3
                                End If
                                
                                'Tipo da NF
                                If TBFI.EOF = False Then Tipo_NF4_Contimatic = 2
                                TBFI.Close
                                
                                If TBproducao!Int_status = 2 Then Tipo_NF4_Contimatic = 99
                            Else
                                If CST = "00" Or CST = "10" Or CST = "20" Or CST = "101" Or CST = "102" Or CST = "201" Or CST = "202" Then
                                    'Alíquota do ICMS
                                    AliquotaICMS5_Contimatic = Qtd
                                    If TBFI.EOF = False Then
                                        'Valor da BC do ICMS com IPI
                                        valor = valor + IIf(IsNull(TBProduto!VlrBase_ICMS), 0, TBProduto!VlrBase_ICMS) + IIf(IsNull(TBProduto!ValorIPI), 0, TBProduto!ValorIPI)
                                        'Valor do ICMS com IPI
                                        Valor1 = Valor1 + ((IIf(IsNull(TBProduto!VlrBase_ICMS), 0, TBProduto!VlrBase_ICMS) + IIf(IsNull(TBProduto!ValorIPI), 0, TBProduto!ValorIPI)) * Qtd) / 100
                                    Else
                                        'Valor da BC do ICMS
                                        valor = valor + IIf(IsNull(TBProduto!VlrBase_ICMS), 0, TBProduto!VlrBase_ICMS)
                                        'Valor do ICMS
                                        Valor1 = Valor1 + (IIf(IsNull(TBProduto!VlrBase_ICMS), 0, TBProduto!VlrBase_ICMS) * Qtd) / 100
                                    End If
                                    VlrBase_ICMS5_Contimatic = valor
                                    VlrICMS5_Contimatic = Valor1
                                    'Tipo da NF
                                    If CST = "00" Or CST = "101" Or CST = "102" Then
                                        Tipo_NF5_Contimatic = 0
                                    ElseIf CST = "10" Or CST = "201" Or CST = "202" Then
                                            Tipo_NF5_Contimatic = 13
                                        Else
                                            Tipo_NF5_Contimatic = 16
                                    End If
                                    
                                ElseIf CST = "30" Or CST = "40" Or CST = "103" Or CST = "203" Then
                                        If TBFI.EOF = False Then
                                            'Valor da BC do ICMS Isento com IPI
                                            Valor2 = Valor2 + IIf(IsNull(TBProduto!VlrBase_ICMS), 0, TBProduto!VlrBase_ICMS) + IIf(IsNull(TBProduto!ValorIPI), 0, TBProduto!ValorIPI)
                                        Else
                                            'Valor da BC do ICMS Isento
                                            Valor2 = Valor2 + IIf(IsNull(TBProduto!VlrBase_ICMS), 0, TBProduto!VlrBase_ICMS)
                                        End If
                                        VlrBase_ICMS_Isento5_Contimatic = Valor2
                                    Else
                                        If TBFI.EOF = False Then
                                            'Valor da BC do ICMS Outros com IPI
                                            Valor3 = Valor3 + IIf(IsNull(TBProduto!VlrBase_ICMS), 0, TBProduto!VlrBase_ICMS) + IIf(IsNull(TBProduto!ValorIPI), 0, TBProduto!ValorIPI)
                                        Else
                                            'Valor da BC do ICMS Outros
                                            Valor3 = Valor3 + IIf(IsNull(TBProduto!VlrBase_ICMS), 0, TBProduto!VlrBase_ICMS)
                                        End If
                                        VlrBase_ICMS_Outros5_Contimatic = Valor3
                                End If
                                
                                'Tipo da NF
                                If TBFI.EOF = False Then Tipo_NF5_Contimatic = 2
                                TBFI.Close
                                
                                If TBproducao!Int_status = 2 Then Tipo_NF5_Contimatic = 99
                End If
                TBProduto.MoveNext
            Loop
        End If
        TBProduto.Close
        
        TBCFOP.MoveNext
        Contador = Contador + 1
    Loop
End If
    
TextoICMS1 = FunTamanhoTextoZeroEsq(ReturnNumbersOnly(Format(VlrBase_ICMS1_Contimatic, "###,##0.00")), 12) & "|" & FunTamanhoTextoZeroEsq(ReturnNumbersOnly(Format(AliquotaICMS1_Contimatic, "###,##0.0000")), 6) & "|" & FunTamanhoTextoZeroEsq(ReturnNumbersOnly(Format(VlrICMS1_Contimatic, "###,##0.00")), 12) & "|" & FunTamanhoTextoZeroEsq(ReturnNumbersOnly(Format(VlrBase_ICMS_Isento1_Contimatic, "###,##0.00")), 12) & "|" & FunTamanhoTextoZeroEsq(ReturnNumbersOnly(Format(VlrBase_ICMS_Outros1_Contimatic, "###,##0.00")), 12) & "|" & FunTamanhoTextoZeroEsq(Tipo_NF1_Contimatic, 2) & "|"
TextoICMS2 = FunTamanhoTextoZeroEsq(ReturnNumbersOnly(Format(VlrBase_ICMS2_Contimatic, "###,##0.00")), 12) & "|" & FunTamanhoTextoZeroEsq(ReturnNumbersOnly(Format(AliquotaICMS2_Contimatic, "###,##0.0000")), 6) & "|" & FunTamanhoTextoZeroEsq(ReturnNumbersOnly(Format(VlrICMS2_Contimatic, "###,##0.00")), 12) & "|" & FunTamanhoTextoZeroEsq(ReturnNumbersOnly(Format(VlrBase_ICMS_Isento2_Contimatic, "###,##0.00")), 12) & "|" & FunTamanhoTextoZeroEsq(ReturnNumbersOnly(Format(VlrBase_ICMS_Outros2_Contimatic, "###,##0.00")), 12) & "|" & FunTamanhoTextoZeroEsq(Tipo_NF2_Contimatic, 2) & "|"
TextoICMS3 = FunTamanhoTextoZeroEsq(ReturnNumbersOnly(Format(VlrBase_ICMS3_Contimatic, "###,##0.00")), 12) & "|" & FunTamanhoTextoZeroEsq(ReturnNumbersOnly(Format(AliquotaICMS3_Contimatic, "###,##0.0000")), 6) & "|" & FunTamanhoTextoZeroEsq(ReturnNumbersOnly(Format(VlrICMS3_Contimatic, "###,##0.00")), 12) & "|" & FunTamanhoTextoZeroEsq(ReturnNumbersOnly(Format(VlrBase_ICMS_Isento3_Contimatic, "###,##0.00")), 12) & "|" & FunTamanhoTextoZeroEsq(ReturnNumbersOnly(Format(VlrBase_ICMS_Outros3_Contimatic, "###,##0.00")), 12) & "|" & FunTamanhoTextoZeroEsq(Tipo_NF3_Contimatic, 2) & "|"
TextoICMS4 = FunTamanhoTextoZeroEsq(ReturnNumbersOnly(Format(VlrBase_ICMS4_Contimatic, "###,##0.00")), 12) & "|" & FunTamanhoTextoZeroEsq(ReturnNumbersOnly(Format(AliquotaICMS4_Contimatic, "###,##0.0000")), 6) & "|" & FunTamanhoTextoZeroEsq(ReturnNumbersOnly(Format(VlrICMS4_Contimatic, "###,##0.00")), 12) & "|" & FunTamanhoTextoZeroEsq(ReturnNumbersOnly(Format(VlrBase_ICMS_Isento4_Contimatic, "###,##0.00")), 12) & "|" & FunTamanhoTextoZeroEsq(ReturnNumbersOnly(Format(VlrBase_ICMS_Outros4_Contimatic, "###,##0.00")), 12) & "|" & FunTamanhoTextoZeroEsq(Tipo_NF4_Contimatic, 2) & "|"
TextoICMS5 = FunTamanhoTextoZeroEsq(ReturnNumbersOnly(Format(VlrBase_ICMS5_Contimatic, "###,##0.00")), 12) & "|" & FunTamanhoTextoZeroEsq(ReturnNumbersOnly(Format(AliquotaICMS5_Contimatic, "###,##0.0000")), 6) & "|" & FunTamanhoTextoZeroEsq(ReturnNumbersOnly(Format(VlrICMS5_Contimatic, "###,##0.00")), 12) & "|" & FunTamanhoTextoZeroEsq(ReturnNumbersOnly(Format(VlrBase_ICMS_Isento5_Contimatic, "###,##0.00")), 12) & "|" & FunTamanhoTextoZeroEsq(ReturnNumbersOnly(Format(VlrBase_ICMS_Outros5_Contimatic, "###,##0.00")), 12) & "|" & FunTamanhoTextoZeroEsq(Tipo_NF5_Contimatic, 2) & "|"
    
'BC IPI isento
VlrBase_IPI_Isento = 0
Set TBCFOP = CreateObject("adodb.recordset")
TBCFOP.Open "select Sum(dbl_ValorTotal) as VlrBase_IPI_Isento from tbl_detalhes_nota where ID_nota = " & TBproducao!ID & " and (CST_IPI = '02' or CST_IPI = '52')", Conexao, adOpenKeyset, adLockOptimistic
If TBCFOP.EOF = False Then
    VlrBase_IPI_Isento = IIf(IsNull(TBCFOP!VlrBase_IPI_Isento), 0, TBCFOP!VlrBase_IPI_Isento)
End If

'BC IPI outros
VlrBase_IPI_Outros = 0
Set TBCFOP = CreateObject("adodb.recordset")
TBCFOP.Open "select Sum(dbl_ValorTotal) as VlrBase_IPI_Outros from tbl_detalhes_nota where ID_nota = " & TBproducao!ID & " and CST_IPI <> '00' and CST_IPI <> '01' and CST_IPI <> '02' and CST_IPI <> '51' and CST_IPI <> '52'", Conexao, adOpenKeyset, adLockOptimistic
If TBCFOP.EOF = False Then
    VlrBase_IPI_Outros = IIf(IsNull(TBCFOP!VlrBase_IPI_Outros), 0, TBCFOP!VlrBase_IPI_Outros)
End If
            
'BC PIS/Cofins isento
VlrTotal_PIS_Cofins_Isento_Contimatic = 0
Set TBCFOP = CreateObject("adodb.recordset")
TBCFOP.Open "select Sum(dbl_ValorTotal) as VlrTotal_PIS_Cofins_Isento_Contimatic from tbl_detalhes_nota where ID_nota = " & TBproducao!ID & " and (CST_PIS = '07' or CST_Cofins = '07')", Conexao, adOpenKeyset, adLockOptimistic
If TBCFOP.EOF = False Then
    VlrTotal_PIS_Cofins_Isento_Contimatic = IIf(IsNull(TBCFOP!VlrTotal_PIS_Cofins_Isento_Contimatic), 0, TBCFOP!VlrTotal_PIS_Cofins_Isento_Contimatic)
End If
            
'Destinatario (IM, Complemento do endereço, código do país e código SUFRAMA)
Codigo_SUFRAMA_Contimatic = ""
Set TBCFOP = CreateObject("adodb.recordset")
If TBproducao!txt_tipocliente = "E" Then
    TBCFOP.Open "select * from Empresa where Codigo = " & TBproducao!Id_Int_Cliente, Conexao, adOpenKeyset, adLockOptimistic
    If TBCFOP.EOF = False Then
        IM_Contimatic = IIf(IsNull(TBCFOP!IM), "", TBCFOP!IM)
        Complemento_Endereco_Contimatic = Left(IIf(IsNull(TBCFOP!complemento), "", TBCFOP!complemento), 25)
        
        Set TBFI = CreateObject("adodb.recordset")
        TBFI.Open "select * from Codigos_pais where Pais = '" & TBCFOP!Pais & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBFI.EOF = False Then
            Codigo_Pais_Contimatic = IIf(IsNull(TBFI!CODIGO), "", TBFI!CODIGO)
        End If
        TBFI.Close
    End If
ElseIf TBproducao!txt_tipocliente = "JR" Or TBproducao!txt_tipocliente = "JP" Or TBproducao!txt_tipocliente = "FR" Or TBproducao!txt_tipocliente = "FP" Then
        TBCFOP.Open "select * from clientes where IDcliente = " & TBproducao!Id_Int_Cliente, Conexao, adOpenKeyset, adLockOptimistic
        If TBCFOP.EOF = False Then
            IM_Contimatic = IIf(IsNull(TBCFOP!RG_IM), "", TBCFOP!RG_IM)
            Complemento_Endereco_Contimatic = Left(IIf(IsNull(TBCFOP!complemento), "", TBCFOP!complemento), 25)
            
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "select * from Codigos_pais where Pais = '" & TBCFOP!Pais & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False Then
                Codigo_Pais_Contimatic = IIf(IsNull(TBFI!CODIGO), "", TBFI!CODIGO)
            End If
            TBFI.Close
            
            If IsNull(TBCFOP!Suframa) = False And TBCFOP!Suframa <> "" Then Codigo_SUFRAMA_Contimatic = ReturnNumbersOnly(TBCFOP!Suframa) Else Codigo_SUFRAMA_Contimatic = ""
        End If
    Else
        TBCFOP.Open "select * from compras_fornecedores where IDcliente = " & TBproducao!Id_Int_Cliente, Conexao, adOpenKeyset, adLockOptimistic
        If TBCFOP.EOF = False Then
            IM_Contimatic = ReturnNumbersOnly(IIf(IsNull(TBCFOP!RG_IM), "", TBCFOP!RG_IM))
            Complemento_Endereco_Contimatic = IIf(IsNull(TBCFOP!complemento), "", Left(TBCFOP!complemento, 25))
            
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select * from Codigos_pais where Pais = '" & TBCFOP!Pais & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False Then
                Codigo_Pais_Contimatic = IIf(IsNull(TBFI!CODIGO), "", TBFI!CODIGO)
            End If
            TBFI.Close
        End If
End If
TBCFOP.Close

Texto = "R1|" & Tipo_NF_Contimatic & "|" & Data_Emissao_Contimatic & "|" & Data_Circulacao_Contimatic & "|" & Especie_Contimatic & "|" & FunTamanhoTextoVazioDir(Serie_Contimatic, 3) & "|" & FunTamanhoTextoZeroEsq(Nota_Contimatic, 6) & "|" & FunTamanhoTextoZeroEsq(Ultima_Nota_Contimatic, 6) & "|" & UF_emitente_Contimatic & "|" & FunTamanhoTextoVazioDir(CFOP_Contimatic, 5) & "|||" & Cod_Cidade_DIPAM_Contimatic & "||" & FunTamanhoTextoZeroEsq(ReturnNumbersOnly(Format(VlrTotal_Nota_Contimatic, "###,##0.00")), 12) & "|"
Texto1 = TextoICMS1 & TextoICMS2 & TextoICMS3 & TextoICMS4 & TextoICMS5
Texto2 = FunTamanhoTextoZeroEsq(ReturnNumbersOnly(Format(VlrBase_IPI_Contimatic, "###,##0.00")), 12) & "|" & FunTamanhoTextoZeroEsq(ReturnNumbersOnly(Format(VlrTotal_IPI_Contimatic, "###,##0.00")), 12) & "|" & FunTamanhoTextoZeroEsq(ReturnNumbersOnly(Format(VlrBase_IPI_Isento, "###,##0.00")), 12) & "|" & FunTamanhoTextoZeroEsq(ReturnNumbersOnly(Format(VlrBase_IPI_Outros, "###,##0.00")), 12) & "||" & FunTamanhoTextoZeroEsq(ReturnNumbersOnly(Format(VlrBase_ICMSST_Contimatic, "###,##0.00")), 12) & "|" & FunTamanhoTextoZeroEsq(ReturnNumbersOnly(Format(VlrTotal_ICMSST_Contimatic, "###,##0.00")), 12) & "|"
TextoFinal = Texto & Texto1 & Texto2 & "|" & FunTamanhoTextoZeroEsq(ReturnNumbersOnly(Format(VlrDesconto_Contimatic, "###,##0.00")), 12) & "|||||" & FunTamanhoTextoZeroEsq(ReturnNumbersOnly(Format(VlrTotal_PIS_Cofins_Isento_Contimatic, "###,##0.00")), 12) & "|" & FunTamanhoTextoZeroEsq(ReturnNumbersOnly(TBproducao!txt_CNPJ_CPF), 14) & "|" & FunTamanhoTextoVazioDir(ReturnNumbersOnly(TBproducao!txt_IE_Cliente), 16) & "|" & IIf(IM_Contimatic <> "", FunTamanhoTextoVazioDir(IM_Contimatic, 10), "") & "|" & FunTamanhoTextoVazioDir(Left(TBproducao!txt_Razao_Nome, 35), 35) & "|" & FunTamanhoTextoVazioDir(Left(TBproducao!txt_Endereco, 50), 50) & "|" & FunTamanhoTextoVazioDir(Left(TBproducao!Numero, 6), 6) & "|"
TextoFinal2 = IIf(Complemento_Endereco_Contimatic <> "", FunTamanhoTextoVazioDir(Complemento_Endereco_Contimatic, 25), "") & "|" & FunTamanhoTextoVazioDir(Left(TBproducao!txt_Bairro, 40), 40) & "|" & ReturnNumbersOnly(TBproducao!Txt_CEP) & "|" & Cod_Cidade_Contimatic & "|" & FunTamanhoTextoVazioDir(Left(TBproducao!txt_Municipio, 40), 40) & "|" & TBproducao!txt_UF & "|" & FunTamanhoTextoZeroEsq(Codigo_Pais_Contimatic, 5) & "|" & IIf(Codigo_SUFRAMA_Contimatic <> "", FunTamanhoTextoZeroEsq(Codigo_SUFRAMA_Contimatic, 9), "") & "||"
FunCriarTXTContimatic = TextoFinal & TextoFinal2

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Private Sub ProcSair()
On Error GoTo tratar_erro

Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyF3: ProcExportar
    Case vbKeyEscape: ProcSair
End Select
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaComboEmpresa Cmb_empresa, False
Cmb_tipo = "SA - Serviços"
msk_fltInicio.Value = Date
msk_fltFim.Value = Date

'Verifica se tem o apelido da empresa do contimatic cadastrado
Set TBFI = CreateObject("adodb.recordset")
TBFI.Open "Select * from Empresa where Codigo = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and Apelido_contimatic is not null", Conexao, adOpenKeyset, adLockOptimistic
If TBFI.EOF = True Then Opt_contimatic.Enabled = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaNF()
On Error GoTo tratar_erro

Select Case Cmb_tipo
    Case "M1 - Produtos": Tipo = "M1"
    Case "SA - Serviços": Tipo = "SA"
    Case "M1SA - Produtos/Serviços": Tipo = "M1SA"
End Select
If Chk_periodo.Value = 0 Then
    Cmb_de.Clear
    Cmb_Ate.Clear
    
    If Opt_contimatic.Value = True Then
        TextoFiltro = ""
        TextoFiltro1 = ""
        TextoFiltro2 = ""
        TextoFiltro3 = ""
    Else
        If Opt_propria.Value = True Then TextoFiltro = "and Aplicacao = 'P'" Else TextoFiltro = "and Aplicacao = 'T'"
        If opt_Saida.Value = True Then TextoFiltro1 = "and Int_TipoNota = 1" Else TextoFiltro1 = "and Int_TipoNota = 2"
        TextoFiltro2 = "and int_status = 1"
    End If
    
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select int_NotaFiscal from tbl_Dados_Nota_Fiscal where ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and TipoNF = '" & Tipo & "' and Imprimir = 'False' and DtValidacao IS NOT NULL " & TextoFiltro & " " & TextoFiltro1 & " " & TextoFiltro2 & " group by int_NotaFiscal", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Do While TBAbrir.EOF = False
            Cmb_de.AddItem TBAbrir!int_NotaFiscal
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

Private Sub Opt_contimatic_Click()
On Error GoTo tratar_erro

If Opt_contimatic.Value = True Then
    Frame5.Enabled = True
    ProcCarregaTipoPadrao
    ProcDesabilitaTipoNF
    ProcEsconderRegistro
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt_entrada_Click()
On Error GoTo tratar_erro

ProcCarregaNF

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt_NFp_Click()
On Error GoTo tratar_erro

If Opt_NFp.Value = True Then
    Frame5.Enabled = True
    ProcCarregaTipoPadrao
    ProcHabilitaTipoNF
    ProcEsconderRegistro
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEsconderRegistro()
On Error GoTo tratar_erro

Label16(4).Visible = False
Cmb_registro.Visible = False
Cmb_tipo.Width = 4155

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcMostrarRegistro()
On Error GoTo tratar_erro

Label16(4).Visible = True
Cmb_registro.Visible = True
Cmb_tipo.Width = 2775

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcDesabilitaTipoNF()
On Error GoTo tratar_erro

Frame2.Enabled = False
If Opt_contimatic.Value = True Then
    opt_Entrada.Value = False
    opt_Saida.Value = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcHabilitaTipoNF()
On Error GoTo tratar_erro

Frame2.Enabled = True
opt_Entrada.Value = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt_NFSe_Click()
On Error GoTo tratar_erro

If Opt_NFSe.Value = True Then
    Opt_propria.Value = True
    Frame5.Enabled = False
    opt_Saida.Value = True
    ProcDesabilitaTipoNF
    ProcEsconderRegistro
    With Cmb_tipo
        .Clear
        .AddItem "SA - Serviços"
        .Text = "SA - Serviços"
    End With
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaTipoPadrao()
On Error GoTo tratar_erro

With Cmb_tipo
    .Clear
    .AddItem "M1 - Produtos"
    .AddItem "M1SA - Produtos/Serviços"
    .AddItem "SA - Serviços"
    .Text = "M1 - Produtos"
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub opt_Saida_Click()
On Error GoTo tratar_erro

ProcCarregaNF

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt_sintegra_Click()
On Error GoTo tratar_erro

If Opt_sintegra.Value = True Then
    Frame5.Enabled = True
    ProcCarregaTipoPadrao
    ProcHabilitaTipoNF
    ProcMostrarRegistro
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Sub ProcVerifLiberacao()
On Error GoTo tratar_erro

With frmFaturamento_Prod_Serv
    Permitido = True
    Familiatext = ""
    'Dados da nota fiscal
    Set TBFI = CreateObject("adodb.recordset")
    TBFI.Open "Select * from tbl_dados_nota_fiscal where id = " & TBproducao!ID, Conexao, adOpenKeyset, adLockOptimistic
    If TBFI.EOF = False Then
        If TBFI!Aplicacao = "P" And (TBFI!Serie = "" Or IsNull(TBFI!Serie) = True) Then
            Familiatext = "série da NF"
            Permitido = False
        End If
        If TBFI!txt_UF <> "" And IsNull(TBFI!txt_UF) = False And TBFI!txt_UF <> "EX" Then
            If TBFI!txt_CNPJ_CPF = "" Or IsNull(TBFI!txt_CNPJ_CPF) = True Then
                If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & " CNPJ do destinatário da NF" Else Familiatext = "CNPJ do destinatário da NF"
                Permitido = False
            End If
        End If
        If TBFI!Id_Int_Cliente = "" Or IsNull(TBFI!Id_Int_Cliente) = True Or TBFI!txt_Razao_Nome = "" Or IsNull(TBFI!txt_Razao_Nome) = True Then
            If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & " Destinatário da NF" Else Familiatext = "destinatário da NF"
            Permitido = False
        End If
        If TBFI!txt_Endereco = "" Or IsNull(TBFI!txt_Endereco) = True Then
            If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & " Endereço do destinatário da NF" Else Familiatext = "Endereço do destinatário da NF"
            Permitido = False
        End If
        If TBFI!Numero = "" Or IsNull(TBFI!Numero) = True Then
            If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & " Número do destinatário da NF" Else Familiatext = "Número do destinatário da NF"
            Permitido = False
        End If
        If TBFI!txt_Bairro = "" Or IsNull(TBFI!txt_Bairro) = True Then
            If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & " Bairro do destinatário da NF" Else Familiatext = "Bairro do destinatário da NF"
            Permitido = False
        End If
        If TBFI!Txt_CEP = "" Or IsNull(TBFI!Txt_CEP) = True Then
            If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & " CEP do destinatário da NF" Else Familiatext = "CEP do destinatário da NF"
            Permitido = False
        End If
        If TBFI!txt_UF = "" Or IsNull(TBFI!txt_UF) = True Then
            If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & " UF do destinatário da NF" Else Familiatext = "UF do destinatário da NF"
            Permitido = False
        End If
    End If
    
    If Opt_NFSe.Value = False Then
        'Itens da nota
        Set TBFI = CreateObject("adodb.recordset")
        TBFI.Open "Select * from tbl_Detalhes_Nota where ID_Nota = " & TBproducao!ID, Conexao, adOpenKeyset, adLockOptimistic
        Do While TBFI.EOF = False
            If TBFI!txt_CST = "" Or IsNull(TBFI!txt_CST) = True Then
                If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & " CST de ICMS do produto " & TBFI!int_Cod_Produto Else Familiatext = "CST de ICMS do produto " & TBFI!int_Cod_Produto
                Permitido = False
            End If
            If Opt_NFp.Value = True Then
                If TBFI!CST_IPI = "" Or IsNull(TBFI!CST_IPI) = True Then
                    If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & " CST de IPI do produto " & TBFI!int_Cod_Produto Else Familiatext = "CST de IPI do produto " & TBFI!int_Cod_Produto
                    Permitido = False
                End If
                If TBFI!CST_PIS = "" Or IsNull(TBFI!CST_PIS) = True Then
                    If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & " CST de PIS do produto " & TBFI!int_Cod_Produto Else Familiatext = "CST de PIS do produto " & TBFI!int_Cod_Produto
                    Permitido = False
                End If
                If TBFI!CST_Cofins = "" Or IsNull(TBFI!CST_Cofins) = True Then
                    If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & " CST de Cofins do produto " & TBFI!int_Cod_Produto Else Familiatext = "CST de Cofins do produto " & TBFI!int_Cod_Produto
                    Permitido = False
                End If
            End If
            If TBFI!ID_CF = "0" Or IsNull(TBFI!ID_CF) = True Then
                If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & " Código da classificação fiscal do produto " & TBFI!int_Cod_Produto Else Familiatext = "Código da classificação fiscal do produto " & TBFI!int_Cod_Produto
                Permitido = False
            End If
            TBFI.MoveNext
        Loop
        
        'Transportadora da nota
        If Opt_NFp.Value = True Then
            'Dados do transporte
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select * from tbl_dados_transp Where id_nota = " & TBproducao!ID, Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = True Then
                If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & " Frete por conta na transportadora" Else Familiatext = "frete por conta na transportadora"
                Permitido = False
            Else
                If TBFI!txt_Frete_Conta <> 1 Then
                    If TBFI!txt_Razao = "" Or IsNull(TBFI!txt_Razao) = True Then
                        If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & " Razão social da transportadora" Else Familiatext = "razão social da transportadora"
                        Permitido = False
                    End If
                    If TBFI!txt_Endereco = "" Or IsNull(TBFI!txt_Endereco) = True Then
                        If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & " Endereço da transportadora" Else Familiatext = "endereço da transportadora"
                        Permitido = False
                    End If
                    If TBFI!int_numero = "" Or IsNull(TBFI!int_numero) = True Then
                        If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & " Número da transportadora" Else Familiatext = "número da transportadora"
                        Permitido = False
                    End If
                    If TBFI!txt_Municipio = "" Or IsNull(TBFI!txt_Municipio) = True Then
                        If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & " Cidade da transportadora" Else Familiatext = "cidade da transportadora"
                        Permitido = False
                    End If
                    If TBFI!txt_UF = "" Or IsNull(TBFI!txt_UF) = True Then
                        If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & " UF da transportadora" Else Familiatext = "UF da transportadora"
                        Permitido = False
                    End If
                    If TBFI!txt_CNPJ = "" Or IsNull(TBFI!txt_CNPJ) = True Then
                        If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & " CNPJ da transportadora" Else Familiatext = "CNPJ da transportadora"
                        Permitido = False
                    End If
                    If TBFI!txt_Placa <> "" And IsNull(TBFI!txt_Placa) = False Then
                        If TBFI!txt_UF_Placa = "" Or IsNull(TBFI!txt_UF_Placa) = True Then
                            If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & " UF da placa do veículo da transportadora" Else Familiatext = "UF da placa do veículo da transportadora"
                            Permitido = False
                        End If
                    End If
                    If TBFI!UF_embarque = "" Or IsNull(TBFI!UF_embarque) = True Then
                        If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & " UF de embarque da transportadora" Else Familiatext = "UF de embarque da transportadora"
                        Permitido = False
                    End If
                    If TBFI!Local_embarque = "" Or IsNull(TBFI!Local_embarque) = True Then
                        If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & " Local de embarque da transportadora" Else Familiatext = "local de embarque da transportadora"
                        Permitido = False
                    End If
                End If
            End If
        End If
    Else
        'Serviços da nota
        Set TBFI = CreateObject("adodb.recordset")
        StrSql = "Select NFP.int_Cod_Produto , P.Cod_servico_NFSE from tbl_Detalhes_Nota NFP INNER JOIN projproduto P ON P.Desenho = NFP.int_Cod_Produto where NFP.ID_Nota = " & TBproducao!ID
        'Debug.print StrSql
        
        TBFI.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
        Do While TBFI.EOF = False
            If IsNull(TBFI!Cod_servico_NFSE) = True Or TBFI!Cod_servico_NFSE = "" Then
                If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & " Código NFSe do serviço " & TBFI!int_Cod_Produto Else Familiatext = "Código NFSe do serviço " & TBFI!int_Cod_Produto
                Permitido = False
            End If
            TBFI.MoveNext
        Loop
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

