VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmFaturamento_Prod_Serv_Imprimir 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Faturamento - Nota fiscal - Menu impressão"
   ClientHeight    =   4005
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5715
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
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
   ScaleHeight     =   4005
   ScaleWidth      =   5715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   435
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   767
      DibPicture      =   "frmFaturamento_Prod_Serv_Imprimir.frx":0000
      CaptionOnCenter =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Icon            =   "frmFaturamento_Prod_Serv_Imprimir.frx":1C95
      ShowMaximize    =   0   'False
      ShowMinimize    =   0   'False
   End
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   19
      Top             =   3600
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   714
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2025
      Left            =   120
      TabIndex        =   14
      Top             =   1500
      Width           =   2535
      Begin VB.CheckBox chkNumeroSerie 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Numeros de série"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   180
         TabIndex        =   21
         Top             =   1620
         Width           =   2145
      End
      Begin VB.CheckBox Chk_visualizar_romaneio 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Visualizar romaneio(s)"
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
         Left            =   180
         TabIndex        =   2
         Top             =   660
         Width           =   2205
      End
      Begin VB.CheckBox Chk_romaneio 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Romaneio(s)"
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
         Left            =   180
         TabIndex        =   5
         Top             =   1380
         Width           =   2145
      End
      Begin VB.CheckBox Chk_visualizar_nota 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Visualizar nota(s)"
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
         Left            =   180
         TabIndex        =   0
         Top             =   180
         Width           =   2085
      End
      Begin VB.CheckBox Chk_notas 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Nota(s) fiscal(ais)"
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
         Left            =   180
         TabIndex        =   3
         Top             =   900
         Width           =   2085
      End
      Begin VB.CheckBox Chk_duplicatas 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Duplicata(s)"
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
         Left            =   180
         TabIndex        =   4
         Top             =   1140
         Width           =   2145
      End
      Begin VB.CheckBox Chk_visualizar_duplicata 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Visualizar duplicata(s)"
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
         Left            =   180
         TabIndex        =   1
         Top             =   420
         Width           =   2235
      End
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
      Height          =   1395
      Left            =   120
      TabIndex        =   12
      Top             =   3720
      Visible         =   0   'False
      Width           =   5385
      Begin VB.ComboBox Cmb_empresa 
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
         ItemData        =   "frmFaturamento_Prod_Serv_Imprimir.frx":1FAF
         Left            =   1050
         List            =   "frmFaturamento_Prod_Serv_Imprimir.frx":1FB1
         Locked          =   -1  'True
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   7
         ToolTipText     =   "Empresa."
         Top             =   180
         Width           =   4155
      End
      Begin VB.ComboBox Cmb_tipo 
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
         ItemData        =   "frmFaturamento_Prod_Serv_Imprimir.frx":1FB3
         Left            =   1050
         List            =   "frmFaturamento_Prod_Serv_Imprimir.frx":1FC0
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   8
         ToolTipText     =   "Tipo da nota fiscal."
         Top             =   570
         Width           =   4155
      End
      Begin VB.ComboBox Cmb_Ate 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   330
         Left            =   3780
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   11
         ToolTipText     =   "Número da nota fiscal."
         Top             =   960
         Width           =   1425
      End
      Begin VB.ComboBox Cmb_De 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   330
         Left            =   1860
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   10
         ToolTipText     =   "Número da nota fiscal."
         Top             =   960
         Width           =   1425
      End
      Begin VB.OptionButton Opt_De 
         BackColor       =   &H00E0E0E0&
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
         Left            =   1200
         TabIndex        =   9
         Top             =   990
         Width           =   615
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   16
         Top             =   180
         Width           =   825
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   510
         TabIndex        =   15
         Top             =   570
         Width           =   450
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
         Left            =   3375
         TabIndex        =   13
         Top             =   990
         Width           =   360
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   0
      TabIndex        =   17
      Top             =   450
      Width           =   5715
      _ExtentX        =   10081
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
      ButtonCaption1  =   "Relatório"
      ButtonEnabled1  =   0   'False
      ButtonIconSize1 =   32
      ButtonToolTipText1=   "Relatório (F5)"
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
      ButtonWidth1    =   51
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
      ButtonLeft2     =   55
      ButtonTop2      =   4
      ButtonWidth2    =   2
      ButtonHeight2   =   54
      ButtonCaption3  =   "Ajuda"
      ButtonEnabled3  =   0   'False
      ButtonIconSize3 =   32
      ButtonToolTipText3=   "Ajuda (F1)"
      ButtonKey3      =   "6"
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
      ButtonLeft3     =   59
      ButtonTop3      =   2
      ButtonWidth3    =   36
      ButtonHeight3   =   21
      ButtonUseMaskColor3=   0   'False
      ButtonCaption4  =   "Sair"
      ButtonEnabled4  =   0   'False
      ButtonIconSize4 =   32
      ButtonToolTipText4=   "Sair (Esc)"
      ButtonKey4      =   "7"
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
      ButtonWidth4    =   26
      ButtonHeight4   =   21
      ButtonUseMaskColor4=   0   'False
      ButtonEnabled5  =   0   'False
      ButtonIconSize5 =   32
      ButtonKey5      =   "8"
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
      ButtonState5    =   5
      ButtonLeft5     =   125
      ButtonTop5      =   2
      ButtonWidth5    =   24
      ButtonHeight5   =   24
      ButtonUseMaskColor5=   0   'False
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   3150
         Top             =   210
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmFaturamento_Prod_Serv_Imprimir.frx":1FFC
         Count           =   1
      End
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   5130
      Visible         =   0   'False
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
   Begin DrawSuite2022.USButton Cmd_avancar 
      Height          =   1905
      Left            =   2760
      TabIndex        =   6
      Top             =   1560
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   3360
      BorderColor     =   5263559
      BorderColorDisabled=   13160660
      BorderColorDown =   4013465
      BorderColorOver =   4408288
      Caption         =   "Avançar >>>"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   15.75
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
      PicAlign        =   6
      PicSizeH        =   48
      PicSizeW        =   48
      ShowFocusRect   =   0   'False
      ShowFocusRectDown=   0   'False
      Theme           =   4
   End
End
Attribute VB_Name = "frmFaturamento_Prod_Serv_Imprimir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FormulaRel_Faturamento_NF As String 'OK

Private Sub Chk_duplicatas_Click()
On Error GoTo tratar_erro

procClick

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_notas_Click()
On Error GoTo tratar_erro

procClick

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_romaneio_Click()
On Error GoTo tratar_erro

procClick

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_visualizar_duplicata_Click()
On Error GoTo tratar_erro

procClick_visualizar

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_visualizar_nota_Click()
On Error GoTo tratar_erro

procClick_visualizar

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_visualizar_romaneio_Click()
On Error GoTo tratar_erro

procClick_visualizar

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_De_Click()
On Error GoTo tratar_erro

Cmb_Ate.Clear
Cmb_Ate.Enabled = True
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select int_NotaFiscal from tbl_Dados_Nota_Fiscal where int_NotaFiscal IS NOT NULL and TipoNF = '" & Tipo & "' and int_NotaFiscal >= '" & Cmb_de & "' and Aplicacao = 'P' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and dt_DataEmissao >= '" & Format(Date - 60, "Short Date") & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    Do While TBAbrir.EOF = False
        Cmb_Ate.AddItem TBAbrir!int_NotaFiscal
        TBAbrir.MoveNext
    Loop
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_tipo_Click()
On Error GoTo tratar_erro

If Cmb_tipo = "M1 - Produtos" Then
        Tipo = "M1"
ElseIf Cmb_tipo = "SA - Serviços" Then
        Tipo = "SA"
    Else
        Tipo = "M1SA"
End If
ProcCarregaNF

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro

Acao = "executar"
If chkNumeroSerie.Value = 0 And Chk_visualizar_nota.Value = 0 And Chk_visualizar_duplicata.Value = 0 And Chk_notas.Value = 0 And Chk_duplicatas.Value = 0 And Chk_visualizar_romaneio.Value = 0 And Chk_romaneio.Value = 0 Then
    NomeCampo = "uma das opções"
    ProcVerificaAcao
    Exit Sub
End If
If Avancar = True Then
    NomeCampo = "o número da nota fiscal"
    If Opt_De.Value = True And Cmb_de = "" Then
        ProcVerificaAcao
        Cmb_de.SetFocus
        Exit Sub
    End If
    If Opt_De.Value = True And Cmb_Ate = "" Then
        ProcVerificaAcao
        Cmb_Ate.SetFocus
        Exit Sub
    End If
    
    If Opt_De.Value = True Then
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from tbl_Dados_Nota_Fiscal where TipoNF = '" & Tipo & "' and int_NotaFiscal >= '" & Cmb_de & "' and int_NotaFiscal <= '" & Cmb_Ate & "' and Aplicacao = 'P' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and dt_DataEmissao >= '" & Format(Date - 60, "Short Date") & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            PBLista.Min = 0
            PBLista.Max = TBAbrir.RecordCount
            PBLista.Value = 1
            Contador1 = 0
            Do While TBAbrir.EOF = False
                If Chk_visualizar_nota.Value = 1 Or Chk_notas.Value = 1 Then ProcImprimirNF
                
'                If Chk_visualizar_romaneio.Value = 1 Or Chk_romaneio.Value = 1 Then
'                    ProcVerifEmpresa
'                    NomeRel = "Faturamento_nota fiscal_romaneio.rpt"
'                    FormulaRel_Faturamento_NF = "{tbl_dados_nota_fiscal.id} = " & TBAbrir!ID
'                    If Chk_visualizar_romaneio.Value = 1 Then ProcImprimirRel FormulaRel_Faturamento_NF, "" Else ProcImprimirDireto FormulaRel_Faturamento_NF, ""
'                End If
                
                TBAbrir.MoveNext
                Contador1 = Contador1 + 1
                PBLista.Value = Contador1
            Loop
            If Chk_visualizar_duplicata.Value = 1 Or Chk_duplicatas.Value = 1 Then
                ProcVerifEmpresa
                NomeRel = "Faturamento_duplicata_" & NomeEmpresa & ".rpt"
                FormulaRel_Faturamento_NF = "{tbl_Dados_Nota_Fiscal.TipoNF} = '" & Tipo & "' and {tbl_Dados_Nota_Fiscal.int_NotaFiscal} >= '" & Cmb_de & "' and {tbl_Dados_Nota_Fiscal.int_NotaFiscal} <= '" & Cmb_Ate & "' and {tbl_Dados_Nota_Fiscal.int_status} = 1 and {tbl_Dados_Nota_Fiscal.Aplicacao} = 'P' and {tbl_Dados_Nota_Fiscal.ID_empresa} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
                If Chk_visualizar_duplicata.Value = 1 Then ProcImprimirRel FormulaRel_Faturamento_NF, "" Else ProcImprimirDireto FormulaRel_Faturamento_NF, ""
            End If
        End If
    End If
    If Chk_notas.Value = 1 Then ProcCarregaNF
Else
    With frmFaturamento_Prod_Serv
        If .txtId = "" Or .txtId = "0" Then
            NomeCampo = "a nota fiscal"
            ProcVerificaAcao
            Unload Me
            frm_Localizarnota.Show 1
            Exit Sub
        End If
        
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from tbl_Dados_Nota_Fiscal where ID = " & .txtId.Text & " and ID_empresa = " & .txtIDEmpresa.Text, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            If chkNumeroSerie.Value = 1 Then ProcImprimirNS
            If Chk_visualizar_nota.Value = 1 Or Chk_notas.Value = 1 Then ProcImprimirNF
            If Chk_visualizar_duplicata.Value = 1 Or Chk_duplicatas.Value = 1 Then ProcImprimirDuplicatas
            If Chk_visualizar_romaneio.Value = 1 Or Chk_romaneio.Value = 1 Then procImprimirRomaneio
        End If
    End With
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimirNF()
On Error GoTo tratar_erro

ProcVerifEmpresa
If Avancar = False Then
    frmFaturamento_Prod_Serv.ProcVerificaTipoNF False
    If TipoNF = "M1" Or TipoNF = "M1SA" Then NomeRel = "Faturamento_nota fiscal_produtos.rpt" Else NomeRel = "Faturamento_nota fiscal_servicos.rpt"
Else
    If Cmb_tipo = "M1 - Produtos" Or Cmb_tipo = "M1SA - Produtos/Serviços" Then NomeRel = "Faturamento_nota fiscal_produtos.rpt" Else NomeRel = "Faturamento_nota fiscal_servicos.rpt"
End If
FormulaRel_Faturamento_NF = "{tbl_dados_nota_fiscal.id} = " & TBAbrir!ID
If Chk_visualizar_nota.Value = 1 Then ProcImprimirRel FormulaRel_Faturamento_NF, "" Else ProcImprimirDireto FormulaRel_Faturamento_NF, ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimirNS()
On Error GoTo tratar_erro

ProcVerifEmpresa

NomeRel = "Faturamento_NSerie.rpt"

FormulaRel_Faturamento_NF = "{tbl_dados_nota_fiscal.id} = " & TBAbrir!ID

ProcImprimirRel FormulaRel_Faturamento_NF, ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub ProcImprimirDuplicatas()
On Error GoTo tratar_erro

Set TBFI = CreateObject("adodb.recordset")
TBFI.Open "Select * from tbl_Detalhes_Recebimento where id_nota = " & TBAbrir!ID & " order by id", Conexao, adOpenKeyset, adLockOptimistic
If TBFI.EOF = False Then
    Do While TBFI.EOF = False
        ProcVerifEmpresa
        NomeRel = "Faturamento_duplicata_" & NomeEmpresa & ".rpt"
        FormulaRel_Faturamento_NF = "{tbl_Detalhes_Recebimento.ID}= " & TBFI!ID
        If Chk_visualizar_duplicata.Value = 1 Then ProcImprimirRel FormulaRel_Faturamento_NF, "" Else ProcImprimirDireto FormulaRel_Faturamento_NF, ""
        TBFI.MoveNext
    Loop
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimirDireto(FormulaRel As String, FormulaRelSubReport As String)
On Error GoTo tratar_erro

ProcVerifRelPersonalizado
            
If PermitidoRel = False Then LocalrelNovo = Localrel Else LocalrelNovo = LocalRelPersonalizado
Set Report = crAPP.OpenReport(LocalrelNovo & "\" & NomeRel, crptToPrinter)
'Login SQL
Contador = Report.Database.Tables.Count
Do While Contador > 0
    Set DBTable = Report.Database.Tables(Contador)
    ProcLogonBDSQL
    Contador = Contador - 1
Loop
ProcVerifSubReport FormulaRelSubReport

Report.FormulaSyntax = crCrystalSyntaxFormula 'Configura a sintaxe da formula
Report.RecordSelectionFormula = FormulaRel 'Formula de seleção do relatório
If Chk_romaneio.Value = 1 Then
    Report.ParameterFields(1).AddCurrentValue (Qtd)
    Report.ParameterFields(2).AddCurrentValue (Quant)
    Report.ParameterFields(3).AddCurrentValue (quantidade)
End If

Report.PrintOut False 'Configura a seleção de impressora com false, enviando para impressora padrão
Set Report = Nothing 'Cancela a variavel report
Set crAPP = Nothing 'Cancela a variavel report

If Avancar = False Then
    With frmFaturamento_Prod_Serv
        .Opt_sim.Value = True
        .ListaNota.ListItems.Clear
        .ProcCarregaListaNota (IIf(ReturnNumbersOnly(Left(.lblPaginas(1).Caption, Len(.lblPaginas(1).Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(.lblPaginas(1).Caption, Len(.lblPaginas(1).Caption) - 5))))
    End With
End If

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

Private Sub Cmd_avancar_Click()
On Error GoTo tratar_erro

If Avancar = False Then
    Avancar = True
    Cmd_avancar.Caption = "<<< Recuar"
    Height = 6000
    Frame1.Visible = True
    PBLista.Visible = True
    Cmb_tipo = "M1 - Produtos"
    Opt_De.Value = True
    Cmb_empresa.SetFocus
Else
    Avancar = False
    Cmd_avancar.Caption = "Avançar >>>"
    Height = 4000
    Frame1.Visible = True
    PBLista.Visible = False
    Cmb_tipo.ListIndex = -1
    Opt_De.Value = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyF5: ProcImprimir
    Case vbKeyEscape: ProcSair
End Select
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 5715, 5, True

ProcCarregaComboEmpresa Cmb_empresa, False
If Formulario = "Faturamento/Nota fiscal/Própria" Then
    Caption = "Faturamento - Nota fiscal - Menu impressão"
ElseIf Formulario = "Faturamento/Nota fiscal/Terceiros" Then
        Caption = "Faturamento - Nota fiscal - Menu impressão"
    ElseIf Formulario = "Estoque/Ordem de faturamento" Then
            Caption = "Estoque - Ordem de faturamento - Menu impressão"
        Else
            Caption = "Estoque - Nota fiscal - Menu impressão"
End If
Avancar = False
Height = 4000

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt_De_Click()
On Error GoTo tratar_erro

If Opt_De.Value = True Then
    Cmb_de.Enabled = True
    Cmb_de.SetFocus
    ProcCarregaNF
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaNF()
On Error GoTo tratar_erro

Cmb_de.Clear
Cmb_Ate.Clear
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select int_NotaFiscal from tbl_Dados_Nota_Fiscal where int_NotaFiscal IS NOT NULL and TipoNF = '" & Tipo & "' and Aplicacao = 'P' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and dt_DataEmissao >= '" & Format(Date - 60, "Short Date") & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    Do While TBAbrir.EOF = False
        Cmb_de.AddItem TBAbrir!int_NotaFiscal
        TBAbrir.MoveNext
    Loop
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcVerifEmpresa()
On Error GoTo tratar_erro

Set TBFIltro = CreateObject("adodb.recordset")
TBFIltro.Open "Select * from Empresa where codigo = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex), Conexao, adOpenKeyset, adLockOptimistic
If TBFIltro.EOF = False Then
    NomeEmpresa = IIf(IsNull(TBFIltro!Empresa), "", Trim(Left(TBFIltro!Empresa, 10)))
End If
TBFIltro.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcImprimir
    'Case 3: ProcAjuda
    Case 4: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub procClick_visualizar()
On Error GoTo tratar_erro

If Chk_visualizar_duplicata.Value = 1 Or Chk_visualizar_nota.Value = 1 Or Chk_visualizar_romaneio.Value = 1 Then
    With Chk_notas
        .Value = 0
        .Enabled = False
    End With
    With Chk_duplicatas
        .Value = 0
        .Enabled = False
    End With
    With Chk_romaneio
        .Value = 0
        .Enabled = False
    End With
    If Chk_visualizar_romaneio.Value = 1 Then
        Cmd_avancar.Locked = True
        Avancar = True
        Cmd_avancar_Click
    Else
        Cmd_avancar.Locked = False
    End If
Else
    Chk_notas.Enabled = True
    Chk_duplicatas.Enabled = True
    Chk_romaneio.Enabled = True
    Cmd_avancar.Locked = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub procClick()
On Error GoTo tratar_erro

If Chk_notas.Value = 1 Or Chk_duplicatas = 1 Or Chk_romaneio = 1 Then
    With Chk_visualizar_nota
        .Value = 0
        .Enabled = False
    End With
    With Chk_visualizar_duplicata
        .Value = 0
        .Enabled = False
    End With
    With Chk_visualizar_romaneio
        .Value = 0
        .Enabled = False
    End With
    If Chk_romaneio.Value = 1 Then
        Cmd_avancar.Locked = True
        Avancar = True
        Cmd_avancar_Click
    Else
        Cmd_avancar.Locked = False
    End If
Else
    Chk_visualizar_nota.Enabled = True
    Chk_visualizar_duplicata.Enabled = True
    Chk_visualizar_romaneio.Enabled = True
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub procImprimirRomaneio()
On Error GoTo tratar_erro

Qtde = 0
Set TBFI = CreateObject("adodb.recordset")
TBFI.Open "Select int_Qtd_Transp from tbl_dados_transp where id_nota = " & TBAbrir!ID, Conexao, adOpenKeyset, adLockOptimistic
If TBFI.EOF = False Then
    Qtde = IIf(IsNull(TBFI!int_Qtd_Transp), 0, Format(TBFI!int_Qtd_Transp, "###,##0.0000"))
End If
TBFI.Close
If Qtde <= 0 Then
    USMsgBox ("Esta nota não possui quantidade de volume no cadastro da trasportadora, favor informar antes de imprimir."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

NomeRel = "Faturamento_nota fiscal_romaneio.rpt"
FormulaRel_Faturamento_NF = "{tbl_dados_nota_fiscal.id} = " & TBAbrir!ID
Quant = Qtde
Qtd = 1
Do While Qtde > 0

Mensagem1:
    Texto = InputBox("Favor informar o peso do volume " & Qtd & "/" & Quant & ".", , 0)
    If Texto = "" Then
        Texto = 0
    End If
    If IsNumeric(Texto) = False Then
        USMsgBox ("Só é permitido número neste campo."), vbExclamation, "CAPRIND v5.0"
        GoTo Mensagem1
    End If
    quantidade = Texto
    If quantidade <= 0 Then
        USMsgBox ("So é permitido quantidade maior que 0."), vbExclamation, "CAPRIND v5.0"
        GoTo Mensagem1
    End If
    
    If Chk_visualizar_romaneio.Value = 1 Then ProcImprimirRel_Romaneio FormulaRel_Faturamento_NF, "" Else ProcImprimirDireto FormulaRel_Faturamento_NF, ""
    Qtd = Qtd + 1
    Qtde = Qtde - 1
Loop

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcImprimirRel_Romaneio(FormulaRel As String, FormulaRelSubReport As String)
On Error GoTo tratar_erro

'Exemplo de como colocar variavel no relatorio

ProcVerifRelPersonalizado

If PermitidoRel = False Then LocalrelNovo = Localrel Else LocalrelNovo = LocalRelPersonalizado
Set Report = crAPP.OpenReport(LocalrelNovo & "\" & NomeRel)
'Login SQL
Contador = Report.Database.Tables.Count
Do While Contador > 0
    Set DBTable = Report.Database.Tables(Contador)
    ProcLogonBDSQL
    Contador = Contador - 1
Loop
ProcVerifSubReport FormulaRelSubReport

frmimprimir.CrystalActiveXReportViewer1.ReportSource = Report
Report.FormulaSyntax = crCrystalSyntaxFormula
Report.RecordSelectionFormula = FormulaRel
Report.ParameterFields(1).AddCurrentValue (Qtd)
Report.ParameterFields(2).AddCurrentValue (Quant)
Report.ParameterFields(3).AddCurrentValue (quantidade)

frmimprimir.CrystalActiveXReportViewer1.ViewReport
frmimprimir.Show 1
2:
    Set Report = Nothing
    Set crAPP = Nothing

Exit Sub
tratar_erro:
    If Err.Number = "-2147206461" Then
        USMsgBox ("Não foi encontrado o relatório " & NomeRel & " na pasta " & LocalrelNovo), vbExclamation, "CAPRIND v5.0"
        GoTo 2
    End If
    If Err.Number = "-2147483638" Then
        USMsgBox ("Não foi possível visualizar o relatório, favor reiniciar o sistema."), vbExclamation, "CAPRIND v5.0"
        GoTo 2
    End If
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
