VERSION 5.00
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{50D37AD9-8D3C-43DD-ADD5-7C957C951843}#1.9#0"; "FlexCell.ocx"
Begin VB.Form frmEstoque_consignado 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Estoque | Material de terceiro"
   ClientHeight    =   10035
   ClientLeft      =   1950
   ClientTop       =   1665
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10035
   ScaleWidth      =   15360
   Begin VB.Frame Frame9 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   90
      TabIndex        =   17
      Top             =   5670
      Width           =   15120
      Begin VB.TextBox txtPagIr 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   9690
         TabIndex        =   19
         ToolTipText     =   "Número da página."
         Top             =   180
         Width           =   555
      End
      Begin VB.TextBox txtNreg 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3120
         TabIndex        =   18
         Text            =   "10"
         ToolTipText     =   "Número de registros por página."
         Top             =   180
         Width           =   555
      End
      Begin DrawSuite2022.USButton cmdPagProx 
         Height          =   315
         Left            =   11910
         TabIndex        =   20
         ToolTipText     =   "Próxima página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmEstoque_consignado.frx":0000
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
         BorderColor     =   14404026
         BorderColorDown =   11632444
         BorderColorOver =   11632444
         GradientColor2  =   16777215
         GradientColor3  =   16777215
         GradientColorOver1=   16643560
         GradientColorOver2=   16576988
         GradientColorOver3=   16441780
         GradientColorOver4=   16178091
         GradientColorDown2=   16246986
         GradientColorDown3=   15189380
         GradientColorDown4=   14596208
         PicSizeH        =   19
         PicSizeW        =   19
      End
      Begin DrawSuite2022.USButton cmdPagAnt 
         Height          =   315
         Left            =   11370
         TabIndex        =   21
         ToolTipText     =   "Página anterior."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmEstoque_consignado.frx":37A4
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
         BorderColor     =   14404026
         BorderColorDown =   11632444
         BorderColorOver =   11632444
         GradientColor2  =   16777215
         GradientColor3  =   16777215
         GradientColorOver1=   16643560
         GradientColorOver2=   16576988
         GradientColorOver3=   16441780
         GradientColorOver4=   16178091
         GradientColorDown2=   16246986
         GradientColorDown3=   15189380
         GradientColorDown4=   14596208
         PicSizeH        =   19
         PicSizeW        =   19
      End
      Begin DrawSuite2022.USButton cmdPagIr 
         Height          =   315
         Left            =   10260
         TabIndex        =   22
         Top             =   180
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   556
         Caption         =   "Ir"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   14404026
         BorderColorDown =   11632444
         BorderColorOver =   11632444
         GradientColor2  =   16777215
         GradientColor3  =   16777215
         GradientColorOver1=   16643560
         GradientColorOver2=   16576988
         GradientColorOver3=   16441780
         GradientColorOver4=   16178091
         GradientColorDown2=   16246986
         GradientColorDown3=   15189380
         GradientColorDown4=   14596208
         PicSizeH        =   19
         PicSizeW        =   19
      End
      Begin DrawSuite2022.USButton cmdPagPrim 
         Height          =   315
         Left            =   10830
         TabIndex        =   23
         ToolTipText     =   "Primeira página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmEstoque_consignado.frx":72AD
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
         BorderColor     =   14404026
         BorderColorDown =   11632444
         BorderColorOver =   11632444
         GradientColor2  =   16777215
         GradientColor3  =   16777215
         GradientColorOver1=   16643560
         GradientColorOver2=   16576988
         GradientColorOver3=   16441780
         GradientColorOver4=   16178091
         GradientColorDown2=   16246986
         GradientColorDown3=   15189380
         GradientColorDown4=   14596208
         PicSizeH        =   19
         PicSizeW        =   19
      End
      Begin DrawSuite2022.USButton cmdPagUlt 
         Height          =   315
         Left            =   12450
         TabIndex        =   24
         ToolTipText     =   "Última página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmEstoque_consignado.frx":B39C
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
         BorderColor     =   14404026
         BorderColorDown =   11632444
         BorderColorOver =   11632444
         GradientColor2  =   16777215
         GradientColor3  =   16777215
         GradientColorOver1=   16643560
         GradientColorOver2=   16576988
         GradientColorOver3=   16441780
         GradientColorOver4=   16178091
         GradientColorDown2=   16246986
         GradientColorDown3=   15189380
         GradientColorDown4=   14596208
         PicSizeH        =   19
         PicSizeW        =   19
      End
      Begin VB.Label lblPaginas 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Página: 0 de: 0"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   13200
         TabIndex        =   28
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblRegistros 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nº de registros: 0"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   180
         TabIndex        =   27
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Carregar"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2430
         TabIndex        =   26
         Top             =   240
         Width           =   645
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "registros por página"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3750
         TabIndex        =   25
         Top             =   240
         Width           =   1440
      End
   End
   Begin DrawSuite2022.USButton btnFiltrar 
      Height          =   705
      Left            =   12720
      TabIndex        =   16
      Top             =   1080
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   1244
      DibPicture      =   "frmEstoque_consignado.frx":EC28
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
      PicAlign        =   7
      PicSize         =   2
      PicSizeH        =   24
      PicSizeW        =   24
      ShowFocusRect   =   0   'False
      Theme           =   4
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Texto para pesquisa"
      Height          =   900
      Left            =   8580
      TabIndex        =   14
      Top             =   990
      Width           =   4035
      Begin VB.TextBox txtTexto 
         Alignment       =   2  'Center
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
         TabIndex        =   15
         ToolTipText     =   "Texto para pesquisa."
         Top             =   390
         Width           =   3675
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Notas de saida"
      Height          =   3375
      Left            =   90
      TabIndex        =   9
      Top             =   6300
      Width           =   15105
      Begin FlexCell.Grid GridSaida 
         Height          =   3000
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   14850
         _ExtentX        =   26194
         _ExtentY        =   5292
         Appearance      =   0
         BackColor2      =   14737632
         BackColorBkg    =   16777215
         BorderColor     =   16777215
         Cols            =   1
         DefaultFontSize =   8.25
         BoldFixedCell   =   0   'False
         DisplayRowIndex =   -1  'True
         FixedRowColStyle=   2
         GridColor       =   12632256
         Rows            =   1
         SelectionMode   =   1
         MultiSelect     =   0   'False
         DateFormat      =   2
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Notas de entrada"
      Height          =   3645
      Left            =   90
      TabIndex        =   7
      Top             =   1890
      Width           =   15105
      Begin FlexCell.Grid GridEntrada 
         Height          =   3210
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   14850
         _ExtentX        =   26194
         _ExtentY        =   5662
         AllowUserResizing=   0   'False
         Appearance      =   0
         BackColor2      =   14737632
         BackColorBkg    =   16777215
         BorderColor     =   16777215
         Cols            =   1
         DefaultFontSize =   8.25
         BoldFixedCell   =   0   'False
         DisplayRowIndex =   -1  'True
         EnableVisualStyles=   0   'False
         FixedRowColStyle=   2
         GridColor       =   12632256
         Rows            =   1
         SelectionMode   =   1
         MultiSelect     =   0   'False
         DateFormat      =   2
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Opções para filtro"
      Height          =   900
      Left            =   60
      TabIndex        =   2
      Top             =   990
      Width           =   8505
      Begin VB.OptionButton optTudo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Tudo"
         Height          =   285
         Left            =   7770
         TabIndex        =   13
         Top             =   450
         Width           =   675
      End
      Begin VB.OptionButton OptSemsaldo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Sem saldo"
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   6720
         TabIndex        =   12
         Top             =   450
         Width           =   1065
      End
      Begin VB.OptionButton OptComSaldo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Com saldo"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   5670
         TabIndex        =   11
         Top             =   450
         Value           =   -1  'True
         Width           =   1065
      End
      Begin VB.ComboBox Cmb_empresa 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   150
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         ToolTipText     =   "Empresa."
         Top             =   390
         Width           =   3345
      End
      Begin VB.ComboBox cmbfiltrarpor 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3510
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Opções para filtro."
         Top             =   390
         Width           =   2085
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Empresa"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1575
         TabIndex        =   6
         Top             =   150
         Width           =   615
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filtrar por"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4200
         TabIndex        =   3
         Top             =   180
         Width           =   705
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
      FormHeightDT    =   10500
      FormWidthDT     =   15480
      FormScaleHeightDT=   10035
      FormScaleWidthDT=   15360
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   225
      Left            =   90
      TabIndex        =   1
      Top             =   9690
      Width           =   15120
      _ExtentX        =   26670
      _ExtentY        =   397
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
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   0
      TabIndex        =   4
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
      ButtonHeight3   =   54
      ButtonCaption4  =   "Ajuda"
      ButtonEnabled4  =   0   'False
      ButtonIconSize4 =   32
      ButtonToolTipText4=   "Ajuda (F1)"
      ButtonKey4      =   "4"
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
      ButtonKey5      =   "5"
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
      ButtonKey6      =   "6"
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
      ButtonState6    =   5
      ButtonLeft6     =   163
      ButtonTop6      =   2
      ButtonWidth6    =   24
      ButtonHeight6   =   24
      ButtonUseMaskColor6=   0   'False
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   1140
         Top             =   -90
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmEstoque_consignado.frx":12278
         Count           =   1
      End
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   270
      Top             =   0
      Width           =   480
   End
   Begin VB.Image imgFile 
      Height          =   240
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgFolder 
      Height          =   225
      Left            =   360
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
End
Attribute VB_Name = "frmEstoque_consignado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TBLISTA_Estoque_Movimentacao As ADODB.Recordset 'OK

Sub ProcAjuda()
On Error GoTo tratar_erro

' AbrirVideoWeb ("http://www.youtube.com/watch?v=o9mVNykTaq0&list=UUODGiDjQ-BCrxh0YXfJvoqg&index=10&feature=plcp")

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub


Private Sub ProcImprimir()
On Error GoTo tratar_erro

'NomeRel = "Faturamento_relacionamento.rpt"
'ProcImprimirRel FiltroRel_Faturamento_Relatorios_Relacionamento, ""

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub btnFiltrar_Click()
On Error GoTo tratar_erro

 ProcFiltrar

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Cmb_empresa_Click()
On Error GoTo tratar_erro

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub cmbfiltrarpor_Change()
On Error GoTo tratar_erro

txtTexto.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbfiltrarpor_Click()
On Error GoTo tratar_erro

'txtTexto.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Estoque_Movimentacao.AbsolutePage <> 2 Then
    If TBLISTA_Estoque_Movimentacao.AbsolutePage = -3 Then
        ProcExibePaginaGrid (TBLISTA_Estoque_Movimentacao.PageCount - 1)
    Else
        TBLISTA_Estoque_Movimentacao.AbsolutePage = TBLISTA_Estoque_Movimentacao.AbsolutePage - 2
        ProcExibePaginaGrid (TBLISTA_Estoque_Movimentacao.AbsolutePage)
    End If
Else
    ProcExibePaginaGrid (1)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagIr_Click()
On Error GoTo tratar_erro

If txtPagIr = "" Then Exit Sub
Quant = ReturnNumbersOnly(Right(lblPaginas.Caption, 4))
If Quant <= 1 Or txtPagIr > Quant Then Exit Sub
If txtPagIr.Text >= 1 And txtPagIr.Text <= Quant Then
    TBLISTA_Estoque_Movimentacao.AbsolutePage = txtPagIr.Text
    ProcExibePaginaGrid (TBLISTA_Estoque_Movimentacao.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Estoque_Movimentacao.AbsolutePage = 1
ProcExibePaginaGrid (TBLISTA_Estoque_Movimentacao.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Estoque_Movimentacao.AbsolutePage <> -3 Then
    If TBLISTA_Estoque_Movimentacao.AbsolutePage = 1 Then
        ProcExibePaginaGrid (2)
    Else
        ProcExibePaginaGrid (TBLISTA_Estoque_Movimentacao.AbsolutePage)
    End If
Else
    ProcExibePaginaGrid (TBLISTA_Estoque_Movimentacao.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Estoque_Movimentacao.AbsolutePage = TBLISTA_Estoque_Movimentacao.PageCount
ProcExibePaginaGrid (TBLISTA_Estoque_Movimentacao.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyF2: ProcFiltrar
    Case vbKeyF5: ProcImprimir
    Case vbKeyF1: ProcAjuda
    Case vbKeyEscape: Unload Me
End Select


Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15195, 6, True
ProcCarregaComboEmpresa Cmb_empresa, False

With cmbfiltrarpor
    .Clear
    .AddItem "Código de referência"
    .AddItem "Código interno"
    .AddItem "Emitente"
    .AddItem "Descrição"
    .AddItem "Nota fiscal"
End With

    With GridEntrada
        .AutoRedraw = False
        .AllowUserPaste = cellTextOnly
        .ExtendLastCol = True
        .DrawMode = cellOwnerDraw
        .Cols = 13
        .rows = 1
        
        .Cell(0, 1).Text = "N° nota fiscal"
        .Cell(0, 2).Text = "Emitente"
        .Cell(0, 3).Text = "Data emissão"
        .Cell(0, 4).Text = "Cód. interno"
        .Cell(0, 5).Text = "Referencia"
        .Cell(0, 6).Text = "Descriçao"
        .Cell(0, 7).Text = "Entrada"
        .Cell(0, 8).Text = "Total Saída"
        .Cell(0, 9).Text = "Saldo"
        .Column(0).Width = 20
        .Column(1).Width = 70
        .Column(2).Width = 200
        .Column(3).Width = 70
        .Column(4).Width = 85
        .Column(5).Width = 90
        .Column(6).Width = 270
        .Column(7).Width = 60
        .Column(8).Width = 60
        .Column(9).Width = 120
        .Column(10).Width = 0
        .Column(11).Width = 0
        .Column(12).Width = 10
        
        .Column(1).Alignment = cellCenterCenter
        .Column(2).Alignment = cellLeftCenter
        .Column(3).Alignment = cellCenterCenter
        .Column(4).Alignment = cellCenterCenter
        .Column(5).Alignment = cellCenterCenter
        .Column(6).Alignment = cellLeftCenter
        .Column(7).Alignment = cellRightCenter
        .Column(8).Alignment = cellRightCenter
        .Column(9).Alignment = cellRightCenter
'        .Column(10).Alignment = cellRightCenter
'        .Column(11).Alignment = cellRightCenter
        
        .AutoRedraw = True
        .Refresh
    End With

    With GridSaida
        .AutoRedraw = False
        .AllowUserPaste = cellTextOnly
        .ExtendLastCol = True
        .DrawMode = cellOwnerDraw
        .Cols = 9
        .rows = 1
        
        .Cell(0, 1).Text = "N° nota fiscal"
        .Cell(0, 2).Text = "Destinatário"
        .Cell(0, 3).Text = "Data emissão"
        .Cell(0, 4).Text = "Cód. interno"
        .Cell(0, 5).Text = "Referencia"
        .Cell(0, 6).Text = "Descriçao"
        .Cell(0, 7).Text = "Saída"
        .Cell(0, 8).Text = "Saldo"
        .Column(0).Width = 20
        .Column(1).Width = 70
        .Column(2).Width = 200
        .Column(3).Width = 70
        .Column(4).Width = 85
        .Column(5).Width = 90
        .Column(6).Width = 335
        .Column(7).Width = 60
        .Column(8).Width = 60
        
        .Column(1).Alignment = cellCenterCenter
        .Column(2).Alignment = cellLeftCenter
        .Column(3).Alignment = cellCenterCenter
        .Column(4).Alignment = cellCenterCenter
        .Column(5).Alignment = cellCenterCenter
        .Column(6).Alignment = cellLeftCenter
        .Column(7).Alignment = cellRightCenter
        .Column(8).Alignment = cellRightCenter
        .AutoRedraw = True
        .Refresh
    End With



cmbfiltrarpor = "Nota fiscal"

ProcRemoveObjetosResize Me

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub



Sub ProcCarregaConsignacaoOld()
On Error GoTo tratar_erro

'Dim arrNodes(15) As NodeData
Dim tempNode As Node
Dim intIndex, i As Integer
    
Call m_Tree.Nodes.Clear

Grid1.rows = 1

m_Row = 1
m_Col = 1
Contador1 = -1

Set TBLISTA_Estoque_Movimentacao = CreateObject("adodb.recordset")
TBLISTA_Estoque_Movimentacao.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
'TBLISTA_Estoque_Movimentacao.Open "Select * from Estoque_Consignado_entrada order by dt_DataEmissao", Conexao, adOpenKeyset, adLockOptimistic

If TBLISTA_Estoque_Movimentacao.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA_Estoque_Movimentacao.RecordCount
    PBLista.Value = 1
    Contador = 0
        Do While Not TBLISTA_Estoque_Movimentacao.EOF
          Contador1 = Contador1 + 1
          arrNodes(Contador1).Level = 0
          Saldo = TBLISTA_Estoque_Movimentacao!int_Qtd
          arrNodes(Contador1).Text = TBLISTA_Estoque_Movimentacao!int_NotaFiscal & vbTab & TBLISTA_Estoque_Movimentacao!txt_Razao_Nome & vbTab & TBLISTA_Estoque_Movimentacao!dt_DataEmissao & vbTab & TBLISTA_Estoque_Movimentacao!int_Cod_Produto & vbTab & TBLISTA_Estoque_Movimentacao!N_referencia & vbTab & TBLISTA_Estoque_Movimentacao!Txt_descricao & vbTab & Format(TBLISTA_Estoque_Movimentacao!int_Qtd, "###,##0.00") & vbTab & vbTab & Format(Saldo, "###,##0.00") & vbTab & ""
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select * from Faturamento_Relacionamento where ID_nota_relacionada = '" & TBLISTA_Estoque_Movimentacao!ID & "' and ID_produto_relacionada = '" & TBLISTA_Estoque_Movimentacao!Int_codigo & "'", Conexao, adOpenKeyset, adLockOptimistic
                'Debug.print TBAbrir.RecordCount
                Do While Not TBAbrir.EOF
                    If TBAbrir.EOF = False Then
                        ProcNivel2Consignacao 'Carrega notas de saida
                    End If
                    TBAbrir.MoveNext
                Loop
            TBLISTA_Estoque_Movimentacao.MoveNext
            Contador = Contador + 1
            PBLista.Value = Contador
        Loop
    
    With Grid1
        .AutoRedraw = False
        .AllowUserPaste = cellTextOnly
        .ExtendLastCol = True
        .DrawMode = cellOwnerDraw
        .Cols = 10
        .rows = m_Row
        
        .Cell(0, 1).Text = "N° nota fiscal"
        .Cell(0, 2).Text = "Emitente"
        .Cell(0, 3).Text = "Data emissão"
        .Cell(0, 4).Text = "Cód. interno"
        .Cell(0, 5).Text = "Referencia"
        .Cell(0, 6).Text = "Descriçao"
        .Cell(0, 7).Text = "Entrada"
        .Cell(0, 8).Text = "Saida"
        .Cell(0, 9).Text = "Saldo"
        .Column(0).Width = 20
        .Column(1).Width = 140
        .Column(2).Width = 200
        .Column(3).Width = 70
        .Column(4).Width = 85
        .Column(5).Width = 90
        .Column(6).Width = 270
        .Column(7).Width = 60
        .Column(8).Width = 60
        .Column(9).Width = 60
        
        .Column(1).Alignment = cellCenterCenter
        .Column(2).Alignment = cellLeftCenter
        .Column(3).Alignment = cellCenterCenter
        .Column(4).Alignment = cellCenterCenter
        .Column(5).Alignment = cellCenterCenter
        .Column(6).Alignment = cellLeftCenter
        .Column(7).Alignment = cellRightCenter
        .Column(8).Alignment = cellRightCenter
        .Column(9).Alignment = cellRightCenter
        
        'First node
        Set tempNode = m_Tree.Nodes.Add("")
        .AddItem arrNodes(0).Text
        
        'Other nodes
        For intIndex = 1 To Contador1 'UBound(arrNodes)
            If arrNodes(intIndex).Level = arrNodes(intIndex - 1).Level Then
                Set tempNode = tempNode.Parent.Nodes.Add("")
            ElseIf arrNodes(intIndex).Level > arrNodes(intIndex - 1).Level Then
                Set tempNode = tempNode.Nodes.Add("")
            ElseIf arrNodes(intIndex).Level < arrNodes(intIndex - 1).Level Then
                For i = arrNodes(intIndex).Level To arrNodes(intIndex - 1).Level
                    Set tempNode = tempNode.Parent
                Next
                Set tempNode = tempNode.Nodes.Add("")
            End If
            .AddItem arrNodes(intIndex).Text
        Next
        
        .AutoRedraw = True
        .Refresh
    End With
End If


Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Sub ProcCarregaEntrada()
On Error GoTo tratar_erro

'Debug.print StrSql

Set TBLISTA_Estoque_Movimentacao = CreateObject("adodb.recordset")
 TBLISTA_Estoque_Movimentacao.Open StrSql, Conexao, adOpenKeyset, adLockReadOnly

If TBLISTA_Estoque_Movimentacao.EOF = False Then
  ProcExibePaginaGrid 1
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcCorrigeForm(Entrada As Boolean)
On Error GoTo tratar_erro

With cmbfiltrarpor
    .Clear
    .AddItem "CFOP"
    .AddItem "Código de referência"
    .AddItem "Código interno"
    .AddItem "Descrição"
    .AddItem "Nota fiscal"
    
    If Entrada = True Then
        .AddItem "Emitente"
        .Text = "Emitente"
    End If
End With

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

Acao = "filtrar"

StrSql = ""
GridEntrada.rows = 1
GridSaida.rows = 1

If txtTexto.Text = "" Then
USMsgBox "Digite um texto para pesquisa", vbCritical, "CAPRIND v5.0"
txtTexto.SetFocus
Exit Sub
End If

Select Case cmbfiltrarpor.Text
    Case "Código de referência": StrSql = "Select * from Estoque_Consignado_entrada where N_Referencia like '%" & txtTexto.Text & "%'"
    Case "Código interno": StrSql = "Select * from Estoque_Consignado_entrada where int_Cod_Produto like '%" & txtTexto.Text & "%'"
    Case "Emitente": StrSql = "Select * from Estoque_Consignado_entrada where txt_Razao_Nome like '%" & txtTexto.Text & "%'"
    Case "Descrição": StrSql = "Select * from Estoque_Consignado_entrada where txt_Descricao like '%" & txtTexto.Text & "%'"
    Case "Nota fiscal": StrSql = "Select * from Estoque_Consignado_entrada where int_NotaFiscal like '%" & txtTexto.Text & "%'"
End Select

If OptComSaldo Then
StrSql = StrSql & " And Saldo > '0'  order by DataEmissao"
End If

If OptSemsaldo Then
StrSql = StrSql & " And Saldo <= '0'  order by DataEmissao"
End If

ProcCarregaEntrada

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcExibePaginaGrid(Pagina)
On Error GoTo tratar_erro

GridEntrada.rows = 1

TBLISTA_Estoque_Movimentacao.PageSize = IIf(txtNreg = "", 22, txtNreg)
TBLISTA_Estoque_Movimentacao.AbsolutePage = Pagina
TamanhoPagina = TBLISTA_Estoque_Movimentacao.PageSize
ContadorReg = 1

PBLista.Min = 0
PBLista.Max = FunVerifMaxPBListaPaginacao(TBLISTA_Estoque_Movimentacao.RecordCount - IIf(Pagina > 1, (TBLISTA_Estoque_Movimentacao.PageSize * (Pagina - 1)), 0), TBLISTA_Estoque_Movimentacao.PageSize)
PBLista.Value = 1

Do While Not TBLISTA_Estoque_Movimentacao.EOF And (ContadorReg <= TamanhoPagina)
  ContadorReg = ContadorReg + 1
  GridEntrada.AddItem TBLISTA_Estoque_Movimentacao!NotaFiscal & vbTab & TBLISTA_Estoque_Movimentacao!Emitente & vbTab & TBLISTA_Estoque_Movimentacao!DataEmissao & vbTab & TBLISTA_Estoque_Movimentacao!CODIGO & vbTab & TBLISTA_Estoque_Movimentacao!Referencia & vbTab & TBLISTA_Estoque_Movimentacao!Descricao & vbTab & Format(TBLISTA_Estoque_Movimentacao!Entrada, "###,##0.00") & vbTab & Format(TBLISTA_Estoque_Movimentacao!TotalSaida, "###,##0.00") & vbTab & Format(TBLISTA_Estoque_Movimentacao!Saldo, "###,##0.00") & vbTab & TBLISTA_Estoque_Movimentacao!ID_nota_relacionada & vbTab & TBLISTA_Estoque_Movimentacao!id_produto_relacionada & vbTab & ""
 Saldo = TBLISTA_Estoque_Movimentacao!Saldo
  If Saldo < 0 Then
  vRow = GridEntrada.rows - 1
  GridEntrada.Range(vRow, 1, vRow, 9).ForeColor = vbRed
   USMsgBox "Atenção " & pubUsuario & "!" & vbCrLf & "Existe um item com saldo menor que zero.", vbCritical, "CAPRIND v5.0"
  End If
  
  TBLISTA_Estoque_Movimentacao.MoveNext
  Contador = Contador + 1
  PBLista.Value = Contador
Loop
    
With GridEntrada
 .AutoRedraw = True
 .Refresh
End With

lblRegistros.Caption = "Nº de registros: " & TBLISTA_Estoque_Movimentacao.RecordCount

If TBLISTA_Estoque_Movimentacao.AbsolutePage = adPosBOF Then
  lblPaginas.Caption = "Página: 1 de: " & TBLISTA_Estoque_Movimentacao.PageCount
ElseIf TBLISTA_Estoque_Movimentacao.AbsolutePage = adPosEOF Then
  lblPaginas.Caption = "Página: " & TBLISTA_Estoque_Movimentacao.PageCount & " de: " & TBLISTA_Estoque_Movimentacao.PageCount
Else
  lblPaginas.Caption = "Página: " & TBLISTA_Estoque_Movimentacao.AbsolutePage - 1 & " de: " & TBLISTA_Estoque_Movimentacao.PageCount
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub GridEntrada_SelChange(ByVal FirstRow As Long, ByVal FirstCol As Long, ByVal LastRow As Long, ByVal LastCol As Long)
On Error GoTo tratar_erro

ProcCarregaGridSaida

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcCarregaGridSaida()
On Error GoTo tratar_erro
Dim TotalSaldo As Long

vRow = GridEntrada.ActiveCell.Row
GridSaida.rows = 1

If GridEntrada.Cell(vRow, 10).Text <> "" And vRow > 0 Then
ID_nota_relacionada = Int(GridEntrada.Cell(vRow, 10).Text)
id_produto_relacionada = Int(GridEntrada.Cell(vRow, 11).Text)
'Debug.print GridEntrada.Cell(vRow, 7).Text
TotalSaldo = GridEntrada.Cell(vRow, 7).Text
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from Faturamento_Relacionamento where ID_nota_relacionada = '" & ID_nota_relacionada & "' and ID_produto_relacionada = '" & id_produto_relacionada & "'", Conexao, adOpenKeyset, adLockOptimistic
          If TBAbrir.EOF = False Then
            Do While TBAbrir.EOF = False
                 StrSql = "Select * from Estoque_consignado_saida where ID_Nota = '" & TBAbrir!ID_nota & "' and ID_Produto = '" & TBAbrir!ID_Produto & "' order by Id_Nota"
                 TotalSaldo = TotalSaldo - TBAbrir!Qtde
                 Contador1 = Contador1 + 1
                  Set TBEstoque = CreateObject("adodb.recordset")
                  TBEstoque.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
                  Contador = 1
                  If TBEstoque.EOF = False Then
                  Do While TBEstoque.EOF = False
                    GridSaida.AddItem TBEstoque!NotaFiscal & vbTab & TBEstoque!Emitente & vbTab & TBEstoque!DataEmissao & vbTab & TBEstoque!CODIGO & vbTab & TBEstoque!Referencia & vbTab & TBEstoque!Descricao & vbTab & Format(TBAbrir!Qtde, "###,##0.00") & vbTab & Format(TotalSaldo, "###,##0.00")
                    Contador = Contador + 1
                    TBEstoque.MoveNext
                  Loop
                 End If
                 TBEstoque.Close
            TBAbrir.MoveNext
            Loop
         End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtTexto_Change()
On Error GoTo tratar_erro

If txtTexto.Text <> "" And cmbfiltrarpor = "Nota fiscal" Then
    VerifNumero = txtTexto.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtTexto.Text = ""
        txtTexto.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub txtTexto_LostFocus()
On Error GoTo tratar_erro

If cmbfiltrarpor = "Nota fiscal" And txtTexto <> "" Then txtTexto = FunTamanhoTextoZeroEsq(txtTexto, 9)
    
Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcFiltrar
    Case 2: ProcImprimir
    Case 4: ProcAjuda
    Case 5: Unload Me
End Select

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcCarregaComboFiltrarPor(Entrada As Boolean)
On Error GoTo tratar_erro

With cmbfiltrarpor
    .Clear
    .AddItem "CFOP"
    .AddItem "Código de referência"
    .AddItem "Código interno"
    .AddItem "Descrição"
    .AddItem "Emitente"
    .AddItem "Nota fiscal"
    .Text = "Nota fiscal"
End With

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub
