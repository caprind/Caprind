VERSION 5.00
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{50D37AD9-8D3C-43DD-ADD5-7C957C951843}#1.9#0"; "FlexCell.ocx"
Begin VB.Form frmFaturamento_Impostos 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Administrativo - Faturamento - Impostos"
   ClientHeight    =   10035
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15315
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10035
   ScaleWidth      =   15315
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Totalização impostos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   645
      Left            =   9390
      TabIndex        =   14
      Top             =   9300
      Width           =   5925
      Begin VB.TextBox txtTotalIPI 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   4410
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   210
         Width           =   1335
      End
      Begin VB.TextBox txtTotalICMS 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   1950
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   210
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Total IPI : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3660
         TabIndex        =   17
         Top             =   270
         Width           =   1545
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Total ICMS : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1020
         TabIndex        =   15
         Top             =   270
         Width           =   1545
      End
   End
   Begin FlexCell.Grid GridItens 
      Height          =   8295
      Left            =   0
      TabIndex        =   1
      Top             =   990
      Width           =   15315
      _ExtentX        =   27014
      _ExtentY        =   14631
      AllowUserReorderColumn=   -1  'True
      AllowUserSort   =   -1  'True
      Appearance      =   0
      BackColor2      =   14737632
      BackColorActiveCellSel=   12640511
      BackColorBkg    =   16777215
      BorderColor     =   12632256
      CellBorderColor =   8421504
      SelectionBorderColor=   4210752
      Cols            =   12
      DefaultFontSize =   8.25
      DisplayFocusRect=   0   'False
      DisplayRowIndex =   -1  'True
      FixedRowColStyle=   2
      GridColor       =   12632256
      ReadOnlyFocusRect=   0
      Rows            =   1
      ScrollBars      =   2
      ScrollBarStyle  =   0
      SelectionMode   =   1
      MultiSelect     =   0   'False
      DateFormat      =   2
      AllowUserPaste  =   2
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Paginação"
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
      Height          =   645
      Left            =   0
      TabIndex        =   2
      Top             =   9300
      Width           =   9375
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
         Left            =   4710
         TabIndex        =   4
         ToolTipText     =   "Número da página."
         Top             =   210
         Width           =   555
      End
      Begin VB.TextBox txtNreg 
         Alignment       =   2  'Center
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
         Left            =   2550
         TabIndex        =   3
         Text            =   "28"
         ToolTipText     =   "Número de registros por página."
         Top             =   210
         Width           =   555
      End
      Begin DrawSuite2022.USButton cmdPagProx 
         Height          =   315
         Left            =   6840
         TabIndex        =   5
         ToolTipText     =   "Próxima página."
         Top             =   210
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmFaturamento_Impostos.frx":0000
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
         Left            =   6300
         TabIndex        =   6
         ToolTipText     =   "Página anterior."
         Top             =   210
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmFaturamento_Impostos.frx":37A4
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
         Left            =   5280
         TabIndex        =   7
         Top             =   210
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
         Left            =   5760
         TabIndex        =   8
         ToolTipText     =   "Primeira página."
         Top             =   210
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmFaturamento_Impostos.frx":72AD
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
         Left            =   7380
         TabIndex        =   9
         ToolTipText     =   "Última página."
         Top             =   210
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmFaturamento_Impostos.frx":B39C
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
      Begin VB.Label lblRegistros 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nº de registros: 0"
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
         Left            =   300
         TabIndex        =   13
         Top             =   300
         Width           =   1275
      End
      Begin VB.Label lblPaginas 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Página: 0 de: 0"
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
         Left            =   8100
         TabIndex        =   12
         Top             =   270
         Width           =   1095
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Carregar"
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
         Left            =   1860
         TabIndex        =   11
         Top             =   270
         Width           =   645
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "registros por página"
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
         Left            =   3180
         TabIndex        =   10
         Top             =   270
         Width           =   1440
      End
   End
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   13020
      Top             =   180
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frmFaturamento_Impostos.frx":EC28
      Count           =   1
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15285
      _ExtentX        =   26961
      _ExtentY        =   1720
      ButtonCount     =   7
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft1     =   2
      ButtonTop1      =   2
      ButtonWidth1    =   42
      ButtonHeight1   =   21
      ButtonCaption2  =   "Salvar"
      ButtonEnabled2  =   0   'False
      ButtonIconSize2 =   32
      ButtonToolTipText2=   "Gravar alterações na lista"
      ButtonKey2      =   "2"
      ButtonAlignment2=   2
      BeginProperty ButtonFont2 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft2     =   46
      ButtonTop2      =   2
      ButtonWidth2    =   49
      ButtonHeight2   =   24
      ButtonUseMaskColor2=   0   'False
      ButtonCaption3  =   "Relatório"
      ButtonEnabled3  =   0   'False
      ButtonIconSize3 =   32
      ButtonToolTipText3=   "Relatório (F5)"
      ButtonKey3      =   "2"
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
      ButtonLeft3     =   97
      ButtonTop3      =   2
      ButtonWidth3    =   60
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
      ButtonLeft4     =   159
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft5     =   163
      ButtonTop5      =   2
      ButtonWidth5    =   41
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft6     =   206
      ButtonTop6      =   2
      ButtonWidth6    =   30
      ButtonHeight6   =   21
      ButtonUseMaskColor6=   0   'False
      ButtonEnabled7  =   0   'False
      ButtonIconSize7 =   32
      ButtonKey7      =   "6"
      ButtonAlignment7=   2
      BeginProperty ButtonFont7 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState7    =   5
      ButtonLeft7     =   238
      ButtonTop7      =   2
      ButtonWidth7    =   24
      ButtonHeight7   =   24
      ButtonUseMaskColor7=   0   'False
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
         FormWidthDT     =   15435
         FormScaleHeightDT=   10035
         FormScaleWidthDT=   15315
         ResizeFormBackground=   -1  'True
         ResizePictureBoxContents=   -1  'True
      End
   End
End
Attribute VB_Name = "frmFaturamento_Impostos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub ProcGravarLista()
On Error GoTo tratar_erro
If USMsgBox("Deseja realmente gravar as alterações feitas na lista de notas fiscais?", vbYesNo, "CAPRIND v5.0") = vbYes Then
If GridItens.rows > 0 Then
Linha = GridItens.rows - 1
Do While Linha >= 1

'Debug.print GridItens.Cell(Linha, 15).Text
IDAntigo = IIf(GridItens.Cell(Linha, 16).Text <> "", GridItens.Cell(Linha, 16).Text, 0)
'===============================================================================================
' Salvar CST ICMS
'===============================================================================================
Set TBAliquota = CreateObject("adodb.recordset")
TBAliquota.Open "Select * from tbl_Detalhes_Nota_CST_ICMS where ID_item = " & IDAntigo, Conexao, adOpenKeyset, adLockOptimistic
If TBAliquota.EOF = False Then
IntICMS = GridItens.Cell(Linha, 10).Text
If Len(GridItens.Cell(Linha, 9).Text) = 4 Then
TBAliquota!Tributacao_ICMS = Right(GridItens.Cell(Linha, 9).Text, 3)
Else
TBAliquota!Tributacao_ICMS = Right(GridItens.Cell(Linha, 9).Text, 2)
End If
CSTICMS = GridItens.Cell(Linha, 9).Text
TBAliquota!ICMS_SN = GridItens.Cell(Linha, 12).Text
TBAliquota!Valor_ICMS = GridItens.Cell(Linha, 12).Text
TBAliquota!Valor_ICMS_SN = GridItens.Cell(Linha, 12).Text
TBAliquota!Valor_BC = GridItens.Cell(Linha, 11).Text
Conexao.Execute "Update tbl_detalhes_Nota set int_ICMS = " & Replace(IntICMS, ",", ".") & " Where int_codigo = '" & IDAntigo & "'"
Conexao.Execute "Update tbl_detalhes_Nota set txt_CST = '" & CSTICMS & "' Where int_codigo = '" & IDAntigo & "'"
TBAliquota.Update
'Linha = Linha - 1
End If

TBAliquota.Close

'===============================================================================================
' Salvar CST IPI
'===============================================================================================
Set TBAliquota = CreateObject("adodb.recordset")
TBAliquota.Open "Select * from tbl_Detalhes_Nota_CST_IPI where ID_item = " & IDAntigo, Conexao, adOpenKeyset, adLockOptimistic
If TBAliquota.EOF = False Then
IntIPI = GridItens.Cell(Linha, 14).Text
Valor_IPI = GridItens.Cell(Linha, 15).Text

TBAliquota!Codigo_situacaoTributaria = GridItens.Cell(Linha, 13).Text

Conexao.Execute "Update tbl_detalhes_Nota set CST_IPI = '" & GridItens.Cell(Linha, 13).Text & "' Where int_codigo = '" & IDAntigo & "'"
Conexao.Execute "Update tbl_detalhes_Nota set int_IPI = " & Replace(IntIPI, ",", ".") & " Where int_codigo = '" & IDAntigo & "'"
Conexao.Execute "Update tbl_detalhes_Nota set dbl_ValorIPI = " & Replace(Valor_IPI, ",", ".") & " Where int_codigo = '" & IDAntigo & "'"
TBAliquota.Update

End If

TBAliquota.Close

Linha = Linha - 1

Loop

ProcCarregaLista
End If
USMsgBox "Alterações gravadas com sucesso!", vbInformation, "CAPRIND v5.0"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15200, 7, True

ProcRemoveObjetosResize Me
ProcAjustaGridItens

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub GridItens_CellChange(ByVal Row As Long, ByVal Col As Long)
On Error GoTo tratar_erro
Dim pICMS As Double
Dim BIcms As Double
Dim vICMS As Double
Dim pIPI As Double
Dim vIPI As Double

If Row > 0 Then
'================================================================
' Calcula valor ICMS
'================================================================
    If GridItens.Cell(Row, 8).Text <> "" Then
        GridItens.Cell(Row, 11).Text = GridItens.Cell(Row, 8).Text
        BIcms = GridItens.Cell(Row, 11).Text
    End If
    
    If GridItens.Cell(Row, 10).Text <> "" Then
        pICMS = GridItens.Cell(Row, 10).Text
    End If
    
    If BIcms And pICMS <> 0 Then
        vICMS = (BIcms * pICMS) / 100
        GridItens.Cell(Row, 12).Text = Format(Round(vICMS, 2), "###,##0.00") 'Round(vICMS, 2)
    End If
    
'================================================================
' Calcula valor IPI
'================================================================
    
    If GridItens.Cell(Row, 14).Text <> "" Then
        pIPI = GridItens.Cell(Row, 14).Text
    End If
    
    If BIcms And pIPI <> 0 Then
        vIPI = (BIcms * pIPI) / 100
        GridItens.Cell(Row, 15).Text = Format(Round(vIPI, 2), "###,##0.00") 'Round(vICMS, 2)
    End If
    
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: frmFaturamento_Impostos_Filtrar.Show 1
    Case 2: ProcGravarLista
    Case 3: ProcImprimir
    Case 5: 'ProcAjuda
    Case 6: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro

NomeRel = "FaturamentoNotasCST.rpt"
ProcImprimirRel FormulaRelatorio, ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Public Sub ProcAjustaGridItens()
On Error GoTo tratar_erro

With GridItens

    .AllowUserPaste = cellTextOnly
    .AllowUserResizing = True
    .ExtendLastCol = False
    .BoldFixedCell = False
    .DisplayDateTimeMask = True
    .DisplayFocusRect = False
    .SelectionMode = cellSelectionFree
    .Cols = 17
    .ScrollBars = cellScrollBarHorizontal
    
    .DrawMode = cellOwnerDraw
    .Column(0).Width = 20
    .Column(0).Alignment = cellCenterCenter
    
    .Appearance = Flat
    .ScrollBarStyle = Flat
    .FixedRowColStyle = Flat
    
    .Cell(0, 1).Text = "Data"
    .Column(1).CellType = cellTextBox
    .Column(1).Alignment = cellCenterCenter
    .Column(1).Locked = True
    .Column(1).Width = 65
    
    .Cell(0, 2).Text = "Nota fiscal"
    .Column(2).CellType = cellTextBox
    .Column(2).Alignment = cellCenterCenter
    .Column(2).Locked = True
    .Column(2).Width = 60

    .Cell(0, 3).Text = "Clasificação"
    .Column(3).CellType = cellTextBox
    .Column(3).Alignment = cellCenterCenter
    .Column(3).Locked = True
    .Column(3).Width = 70
    
    .Cell(0, 4).Text = "Fornecedor"
    .Column(4).CellType = cellTextBox
    .Column(4).Alignment = cellLeftCenter
    .Column(4).Locked = True
    .Column(4).Width = 360
    
    .Cell(0, 5).Text = "Codigo"
    .Column(5).CellType = cellTextBox
    .Column(5).Alignment = cellCenterCenter
    .Column(5).Locked = True
    .Column(5).Width = 80

    .Cell(0, 6).Text = "Descrição"
    .Column(6).CellType = cellTextBox
    .Column(6).Alignment = cellLeftCenter
    .Column(6).Locked = True
    .Column(6).Width = 365
    
    .Cell(0, 7).Text = "Valor unitário"
    .Column(7).CellType = cellTextBox
    .Column(7).Alignment = cellRightCenter
    .Column(7).Locked = True
    .Column(7).Width = 100
    
    .Cell(0, 8).Text = "Valor Total"
    .Column(8).CellType = cellTextBox
    .Column(8).Alignment = cellRightCenter
    .Column(8).Locked = True
    .Column(8).Width = 100


    .Cell(0, 9).Text = "CST ICMS"
    .Column(9).CellType = cellTextBox
    .Column(9).Alignment = cellCenterCenter
    .Column(9).Width = 85


    .Cell(0, 10).Text = "% ICMS"
    .Column(10).CellType = cellTextBox
    .Column(10).Alignment = cellCenterCenter
    .Column(10).Width = 55
    .Cell(0, 10).ForeColor = vbRed
    
    .Cell(0, 11).ForeColor = vbRed
    .Cell(0, 11).Text = "Base ICMS"
    .Column(11).CellType = cellTextBox
    .Column(11).Alignment = cellRightCenter
    .Column(11).Width = 100
    
    .Cell(0, 12).ForeColor = vbRed
    .Cell(0, 12).Text = "Valor ICMS"
    .Column(12).CellType = cellTextBox
    .Column(12).Alignment = cellRightCenter
    .Column(12).Width = 100
    
    .Cell(0, 13).ForeColor = vbRed
    .Cell(0, 13).Text = "CST IPI"
    .Column(13).CellType = cellTextBox
    .Column(13).Alignment = cellCenterCenter
    .Column(13).Width = 55
    
    .Cell(0, 14).ForeColor = vbRed
    .Cell(0, 14).Text = "% IPI"
    .Column(14).CellType = cellTextBox
    .Column(14).Alignment = cellRightCenter
    .Column(14).Width = 50
    
    .Cell(0, 15).ForeColor = vbRed
    .Cell(0, 15).Text = "Valor IPI"
    .Column(15).CellType = cellTextBox
    .Column(15).Alignment = cellRightCenter
    .Column(15).Width = 100
    
    .Cell(0, 16).ForeColor = vbRed
    .Cell(0, 16).Text = "Código"
    .Column(16).CellType = cellTextBox
    .Column(16).Alignment = cellRightCenter
    .Column(16).Width = 0
    .Column(16).Locked = True
      
    
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaLista()
On Error GoTo tratar_erro

lblRegistros.Caption = "Nº de registros: 0"
lblPaginas.Caption = "Página: 0 de: 0"
GridItens.rows = 1
If StrSql = "" Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open StrSql, Conexao, adOpenKeyset, adLockReadOnly
If TBLISTA.EOF = False Then ProcExibePagina (1)

Set TBAbrir = CreateObject("adodb.recordset")
'Debug.print StrSQLTotais
TBAbrir.Open StrSQLTotais, Conexao, adOpenKeyset, adLockReadOnly
If TBAbrir.EOF = False Then
txtTotalICMS = Format(TBAbrir!TotalICMS, "###,##0.00")
txtTotalIPI = Format(TBAbrir!TotalIPI, "###,##0.00")
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcExibePagina(Pagina)
On Error GoTo tratar_erro
Dim L As Long

GridItens.rows = 1
TBLISTA.PageSize = IIf(txtNreg = "", 28, txtNreg)

TBLISTA.AbsolutePage = Pagina
TamanhoPagina = TBLISTA.PageSize
ContadorReg = 1

Contador = 1
Contador2 = 0
TotalICMS = 0
TotalIPI = 0


Do While TBLISTA.EOF = False And (ContadorReg <= TamanhoPagina)
 'Debug.print Round(TBLISTA!ValorICMS, 2)
        With GridItens
        L = Contador
            .AddItem Contador
            .Cell(L, 1).Text = IIf(IsNull(TBLISTA!Data), "", TBLISTA!Data)
            .Cell(L, 2).Text = IIf(IsNull(TBLISTA!NotaFiscal), "", TBLISTA!NotaFiscal)
            .Cell(L, 3).Text = IIf(IsNull(TBLISTA!classificacao), "", TBLISTA!classificacao)
            .Cell(L, 4).Text = IIf(IsNull(TBLISTA!Emitente), "", TBLISTA!Emitente)
            .Cell(L, 5).Text = IIf(IsNull(TBLISTA!CODIGO), "", TBLISTA!CODIGO)
            .Cell(L, 6).Text = IIf(IsNull(TBLISTA!Descricao), "", TBLISTA!Descricao)
            .Cell(L, 7).Text = IIf(IsNull(TBLISTA!Valorunitaario), "", Format(TBLISTA!Valorunitaario, "###,##0.00"))
            .Cell(L, 8).Text = IIf(IsNull(TBLISTA!ValorTotal), "", Format(TBLISTA!ValorTotal, "###,##0.00"))
            .Cell(L, 9).Text = IIf(IsNull(TBLISTA!CSTICMS), "", TBLISTA!CSTICMS)
            .Cell(L, 10).Text = IIf(IsNull(TBLISTA!ICSM), "", TBLISTA!ICSM)
            .Cell(L, 11).Text = IIf(IsNull(TBLISTA!Bcicms), "", Format(TBLISTA!Bcicms, "###,##0.00"))
            .Cell(L, 12).Text = IIf(IsNull(TBLISTA!ValorICMS), "", Format(Round(TBLISTA!ValorICMS, 2), "###,##0.00"))
            .Cell(L, 13).Text = IIf(IsNull(TBLISTA!CSTIPI), "", TBLISTA!CSTIPI)
            .Cell(L, 14).Text = IIf(IsNull(TBLISTA!IPI), "", TBLISTA!IPI)
            .Cell(L, 15).Text = IIf(IsNull(TBLISTA!ValorIPI), "", Format(TBLISTA!ValorIPI, "###,##0.00"))
            .Cell(L, 16).Text = TBLISTA!Int_codigo
        End With
        Contador = Contador + 1

    TBLISTA.MoveNext
    ContadorReg = ContadorReg + 1
    Contador2 = Contador2 + 1
Loop

lblRegistros.Caption = "Nº de registros: " & TBLISTA.RecordCount

If TBLISTA.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "Página: 1 de: " & TBLISTA.PageCount
ElseIf TBLISTA.AbsolutePage = adPosEOF Then
        lblPaginas.Caption = "Página: " & TBLISTA.PageCount & " de: " & TBLISTA.PageCount
    Else
        lblPaginas.Caption = "Página: " & TBLISTA.AbsolutePage - 1 & " de: " & TBLISTA.PageCount
End If


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub cmdPagAnt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA.AbsolutePage <> 2 Then
    If TBLISTA.AbsolutePage = -3 Then
        ProcExibePagina (TBLISTA.PageCount - 1)
    Else
        TBLISTA.AbsolutePage = TBLISTA.AbsolutePage - 2
        ProcExibePagina (TBLISTA.AbsolutePage)
    End If
Else
    ProcExibePagina (1)
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
    TBLISTA.AbsolutePage = txtPagIr.Text
    ProcExibePagina (TBLISTA.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA.AbsolutePage = 1
ProcExibePagina (TBLISTA.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA.AbsolutePage <> -3 Then
    If TBLISTA.AbsolutePage = 1 Then
        ProcExibePagina (2)
    Else
        ProcExibePagina (TBLISTA.AbsolutePage)
    End If
Else
    ProcExibePagina (TBLISTA.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA.AbsolutePage = TBLISTA.PageCount
ProcExibePagina (TBLISTA.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

