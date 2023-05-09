VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmCFI_Localizar 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Estoque - Almoxarifado - Localizar"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8895
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   8895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   1515
      Left            =   55
      TabIndex        =   12
      Top             =   990
      Width           =   8805
      Begin VB.Frame Frame5 
         BackColor       =   &H00E0E0E0&
         Height          =   510
         Left            =   3810
         TabIndex        =   19
         Top             =   210
         Width           =   4785
         Begin VB.OptionButton Optfim 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Fim frase"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2760
            TabIndex        =   5
            Top             =   180
            Width           =   1155
         End
         Begin VB.OptionButton Optinicio 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Início frase"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   180
            TabIndex        =   3
            Top             =   180
            Value           =   -1  'True
            Width           =   1275
         End
         Begin VB.OptionButton Optmeio 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Meio frase"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1470
            TabIndex        =   4
            Top             =   180
            Width           =   1275
         End
         Begin VB.OptionButton optIgual 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Igual"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3930
            TabIndex        =   6
            Top             =   180
            Width           =   705
         End
      End
      Begin VB.TextBox txtTexto 
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
         TabIndex        =   1
         ToolTipText     =   "Texto para pesquisa."
         Top             =   1050
         Width           =   8415
      End
      Begin VB.ComboBox cmbfiltrarpor 
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
         ForeColor       =   &H00000000&
         Height          =   330
         ItemData        =   "frmCFI_Localizar.frx":0000
         Left            =   180
         List            =   "frmCFI_Localizar.frx":001C
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Opções para filtro."
         Top             =   390
         Width           =   3525
      End
      Begin VB.ComboBox cmbFamilia 
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
         Height          =   330
         ItemData        =   "frmCFI_Localizar.frx":009B
         Left            =   180
         List            =   "frmCFI_Localizar.frx":009D
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "Texto para pesquisa."
         Top             =   1050
         Visible         =   0   'False
         Width           =   8415
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filtrar por"
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
         Left            =   1522
         TabIndex        =   14
         Top             =   180
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Texto para pesquisa"
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
         Left            =   3645
         TabIndex        =   13
         Top             =   840
         Width           =   1470
      End
   End
   Begin VB.CheckBox optDevolucao_problema 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Devol. (Material c/ problemas)"
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
      Left            =   2430
      TabIndex        =   9
      ToolTipText     =   "9"
      Top             =   2790
      Width           =   2475
   End
   Begin VB.CheckBox optDevolucao 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Devolução"
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
      Left            =   1260
      TabIndex        =   8
      ToolTipText     =   "8"
      Top             =   2790
      Width           =   1065
   End
   Begin VB.CheckBox optRetirada 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Retirada"
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
      Left            =   240
      TabIndex        =   7
      ToolTipText     =   "7"
      Top             =   2790
      Width           =   915
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   675
      Left            =   55
      TabIndex        =   15
      Top             =   2520
      Width           =   8805
      Begin MSComCtl2.DTPicker txtFinal 
         Height          =   315
         Left            =   7320
         TabIndex        =   11
         ToolTipText     =   "Data final."
         Top             =   210
         Width           =   1305
         _ExtentX        =   2302
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
         Format          =   490078209
         CurrentDate     =   39057
      End
      Begin MSComCtl2.DTPicker txtInicio 
         Height          =   315
         Left            =   5430
         TabIndex        =   10
         ToolTipText     =   "Data inicio."
         Top             =   210
         Width           =   1305
         _ExtentX        =   2302
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
         Format          =   490078209
         CurrentDate     =   39057
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
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
         Left            =   5070
         TabIndex        =   17
         Top             =   240
         Width           =   300
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
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
         Left            =   6885
         TabIndex        =   16
         Top             =   240
         Width           =   360
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   55
      TabIndex        =   18
      Top             =   0
      Width           =   8805
      _ExtentX        =   15531
      _ExtentY        =   1720
      ButtonCount     =   5
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
      ButtonEnabled2  =   0   'False
      ButtonIconSize2 =   32
      ButtonAlignment2=   2
      ButtonType2     =   1
      ButtonStyle2    =   -1
      BeginProperty ButtonFont2 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState2    =   -1
      ButtonLeft2     =   40
      ButtonTop2      =   4
      ButtonWidth2    =   2
      ButtonHeight2   =   54
      ButtonCaption3  =   "Ajuda"
      ButtonEnabled3  =   0   'False
      ButtonIconSize3 =   32
      ButtonToolTipText3=   "Ajuda (F1)"
      ButtonKey3      =   "3"
      ButtonAlignment3=   2
      BeginProperty ButtonFont3 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft3     =   44
      ButtonTop3      =   2
      ButtonWidth3    =   36
      ButtonHeight3   =   21
      ButtonUseMaskColor3=   0   'False
      ButtonCaption4  =   "Sair"
      ButtonEnabled4  =   0   'False
      ButtonIconSize4 =   32
      ButtonToolTipText4=   "Sair (Esc)"
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
      ButtonLeft4     =   82
      ButtonTop4      =   2
      ButtonWidth4    =   26
      ButtonHeight4   =   21
      ButtonUseMaskColor4=   0   'False
      ButtonEnabled5  =   0   'False
      ButtonIconSize5 =   32
      ButtonKey5      =   "5"
      ButtonAlignment5=   2
      BeginProperty ButtonFont5 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState5    =   5
      ButtonLeft5     =   110
      ButtonTop5      =   2
      ButtonWidth5    =   24
      ButtonHeight5   =   24
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   7530
         Top             =   180
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmCFI_Localizar.frx":009F
         Count           =   1
      End
   End
End
Attribute VB_Name = "frmCFI_Localizar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbFamilia_Click()
On Error GoTo tratar_erro

If cmbfamilia <> "" Then txtTexto = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbfiltrarpor_Click()
On Error GoTo tratar_erro

If cmbfiltrarpor = "Família" Then
    txtTexto.Visible = False
    cmbfamilia.Visible = True
Else
    txtTexto.Visible = True
    cmbfamilia.Visible = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

With txtFinal
    If FunVerificaDataFinal(txtinicio.Value, .Value) = False Then
        .Value = Date
        .SetFocus
        Exit Sub
    End If
End With

InstFiltro = ""
InstFiltroRel = ""
DataFiltro = "C.codigo_produto IS NOT NULL"
DataFiltroRel = "NOT(ISNULL({CFI.codigo_produto}))"
Restricao_almox = ""
Restricao_almoxRel = ""
If Qualidade_Almox = True Then
    InstFiltro = " and P.Instrumento = 'True'"
    InstFiltroRel = " and {Projproduto.Instrumento} = True"
End If
If optDevolucao.Value = 1 Or optDevolucao_problema.Value = 1 Then
    If optDevolucao.Value = 1 Then
        Restricao_almox = " and C.restricao = 'False'"
        Restricao_almoxRel = " and {CFI.restricao} = False"
    Else
        Restricao_almox = " and C.restricao = 'True'"
        Restricao_almoxRel = " and {CFI.restricao} = True"
    End If
    Data_almox = "C.datadevolucao"
    Data_almoxRel = "CFI.datadevolucao"
Else
    Data_almox = "C.dataretirada"
    Data_almoxRel = "CFI.dataretirada"
End If

If optDevolucao.Value = 1 Or optDevolucao_problema.Value = 1 Or optRetirada.Value = 1 Then
    DataFiltro = Data_almox & " Between '" & Format(txtinicio.Value, "Short Date") & "' And '" & Format(txtFinal.Value, "Short Date") & "'"
    DataFiltroRel = "{" & Data_almoxRel & "} >= Date(" & Year(txtinicio.Value) & "," & Month(txtinicio.Value) & "," & Day(txtinicio.Value) & ") and {" & Data_almoxRel & "} <= Date(" & _
                            Year(txtFinal.Value) & "," & Month(txtFinal.Value) & "," & Day(txtFinal.Value) & ")"
End If

CamposFiltro = "C.IDCFI, C.Status, C.Codigo_produto, C.Descricao, C.Familia, C.IDestoque, C.LOTE, C.Funcionario, C.Quantretirada, C.Dataretirada, C.dataprevisao, C.Quantdevolvido, C.Quantdevolvidoprobl, C.Datadevolucao, C.Observacao, CM.maquina, EC.Ref"
INNERJOINTEXTO = "Select " & CamposFiltro & " from ((((CFI C LEFT JOIN CFI_itens ALI ON C.IDCFI = ALI.id_cfi) INNER JOIN projproduto P ON P.Desenho = C.Codigo_produto) LEFT JOIN Cadmaquinas CM ON CM.idmaquina = C.ID_Maquina) LEFT JOIN Estoque_controle EC ON EC.IDestoque = C.IDestoque) LEFT JOIN Projproduto_fabricante PFAB ON PFAB.Codproduto = P.codproduto"
Ordenar = "C.dataretirada, C.codigo_produto, C.lote, C.idcfi"
TextoFiltroPadrao = DataFiltro & Restricao_almox & InstFiltro & " group by " & CamposFiltro & " order by " & Ordenar
TextoFiltroPadraoRel = DataFiltroRel & Restricao_almoxRel & InstFiltroRel

With frmCFI
    If txtTexto.Visible = True And txtTexto <> "" Or cmbfamilia.Visible = True And cmbfamilia <> "" Then
        If cmbfiltrarpor = "Família" Then
            .SQL_almoxarifado = INNERJOINTEXTO & " where familia = '" & cmbfamilia & "' and " & TextoFiltroPadrao
            .FormulaRel_CFI = "{CFI.familia} = '" & cmbfamilia & "' and " & TextoFiltroPadraoRel
        Else
            Select Case cmbfiltrarpor
                Case "Código interno":
                    TextoFiltro = "C.codigo_produto"
                    TextoFiltroRel = "CFI.codigo_produto"
                Case "Código de referência":
                    TextoFiltro = "EC.Ref"
                    TextoFiltroRel = "Estoque_controle.Ref"
                Case "Número do lote":
                    TextoFiltro = "C.lote"
                    TextoFiltroRel = "CFI.lote"
                Case "Número de série":
                    TextoFiltro = "C.nserie"
                    TextoFiltroRel = "CFI.nserie"
                Case "Descrição":
                    TextoFiltro = "C.descricao"
                    TextoFiltroRel = "CFI.descricao"
                Case "Destino\aplicação":
                    TextoFiltro = "P.desenho"
                    TextoFiltroRel = "projproduto.desenho"
                Case "Part number":
                    TextoFiltro = "PFAB.Part_number"
                    TextoFiltroRel = "Projproduto_fabricante.Part_number"
            End Select
            .SQL_almoxarifado = INNERJOINTEXTO & " where " & TextoFiltro & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and " & DataFiltro & Restricao_almox & InstFiltro & " order by " & Ordenar
            .FormulaRel_CFI = "{" & TextoFiltroRel & "}" & FunVerifTipoFiltroIMFRel(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and " & TextoFiltroPadraoRel
        End If
    Else
        .SQL_almoxarifado = INNERJOINTEXTO & " where " & TextoFiltroPadrao
        .FormulaRel_CFI = TextoFiltroPadraoRel
    End If
    .ProcCarregaLista (1)
End With
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyEscape: Unload Me
    Case vbKeyF2: ProcFiltrar
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 9270, 5, True
If Qualidade_Almox = True Then
    Caption = "Qualidade - Almoxarifado - Localizar"
    ProcCarregaComboFamilia cmbfamilia, "familia <> 'Null' and Qualidade = 'True'", False
    Familiatext = "Q"
Else
    ProcCarregaComboFamilia cmbfamilia, "familia <> 'Null' and compras = 'True'", False
    Familiatext = "C"
End If

ProcFiltroPadrao cmbfiltrarpor, Optmeio, Optfim, optIgual, 0, "Produtos/Serviços", Familiatext, False
If Permitido = False Then cmbfiltrarpor = "Código interno"

txtFinal.Value = Date
txtinicio.Value = Date

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optdevolucao_Click()
On Error GoTo tratar_erro

Frame2.Enabled = False
If optDevolucao.Value = 1 Then
    optRetirada.Value = 0
    optDevolucao_problema.Value = 0
    Frame2.Enabled = True
    txtinicio.SetFocus
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optDevolucao_problema_Click()
On Error GoTo tratar_erro

Frame2.Enabled = False
If optDevolucao_problema.Value = 1 Then
    optRetirada.Value = 0
    optDevolucao.Value = 0
    Frame2.Enabled = True
    txtinicio.SetFocus
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optRetirada_Click()
On Error GoTo tratar_erro

Frame2.Enabled = False
If optRetirada.Value = 1 Then
    optDevolucao.Value = 0
    optDevolucao_problema.Value = 0
    Frame2.Enabled = True
    txtinicio.SetFocus
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtTexto_Change()
On Error GoTo tratar_erro

If txtTexto <> "" Then cmbfamilia.ListIndex = -1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcFiltrar
    'Case 3: ProcAjuda
    Case 4: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
