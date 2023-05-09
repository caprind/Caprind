VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Begin VB.Form frmEstoque_fisico_abrir 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Estoque - Inventário | Localizar"
   ClientHeight    =   3510
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   7245
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
   ScaleHeight     =   3510
   ScaleWidth      =   7245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USButton btnFiltrar 
      Height          =   675
      Left            =   5610
      TabIndex        =   18
      ToolTipText     =   "Filtrar registros"
      Top             =   2160
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1191
      DibPicture      =   "frmEstoque_fisico_abrir.frx":0000
      Caption         =   "Filtrar (F2)"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   4960354
      BorderColorDisabled=   13160660
      BorderColorDown =   4210752
      BorderColorOver =   49152
      GradientColor1  =   4960354
      GradientColor2  =   4960354
      GradientColor3  =   4960354
      GradientColor4  =   4960354
      GradientColorDisabled1=   14215660
      GradientColorDisabled2=   14215660
      GradientColorDisabled3=   14215660
      GradientColorDisabled4=   14215660
      GradientColorOver1=   49152
      GradientColorOver2=   49152
      GradientColorOver3=   49152
      GradientColorOver4=   49152
      GradientColorDown1=   32768
      GradientColorDown2=   32768
      GradientColorDown3=   32768
      GradientColorDown4=   32768
      PicAlign        =   7
      Theme           =   3
   End
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   17
      Top             =   3105
      Width           =   7245
      _ExtentX        =   12779
      _ExtentY        =   714
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   405
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   7245
      _ExtentX        =   12779
      _ExtentY        =   714
      DibPicture      =   "frmEstoque_fisico_abrir.frx":3650
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
      Icon            =   "frmEstoque_fisico_abrir.frx":6CA0
   End
   Begin VB.CheckBox optPeriodo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Por período :"
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
      Left            =   600
      TabIndex        =   7
      Top             =   2460
      Width           =   1245
   End
   Begin VB.Frame Frame1 
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
      Height          =   1515
      Left            =   180
      TabIndex        =   10
      Top             =   630
      Width           =   6885
      Begin VB.Frame Frame11 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Frase"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   3795
         TabIndex        =   15
         Top             =   210
         WhatsThisHelpID =   210
         Width           =   2865
         Begin VB.OptionButton Optfim 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Fim"
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
            Height          =   255
            Left            =   1530
            TabIndex        =   5
            Top             =   210
            Width           =   585
         End
         Begin VB.OptionButton Optinicio 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Início"
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
            Height          =   255
            Left            =   180
            TabIndex        =   3
            Top             =   210
            Value           =   -1  'True
            Width           =   675
         End
         Begin VB.OptionButton Optmeio 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Meio"
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
            Height          =   255
            Left            =   870
            TabIndex        =   4
            Top             =   210
            Width           =   645
         End
         Begin VB.OptionButton optIgual 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Igual"
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
            Height          =   255
            Left            =   2130
            TabIndex        =   6
            Top             =   210
            Width           =   705
         End
      End
      Begin VB.ComboBox cmbfiltrarpor 
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
         ItemData        =   "frmEstoque_fisico_abrir.frx":6FBA
         Left            =   180
         List            =   "frmEstoque_fisico_abrir.frx":6FE8
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Opções para filtro."
         Top             =   390
         Width           =   3555
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
         Width           =   6510
      End
      Begin VB.ComboBox cmbfamilia 
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
         Left            =   180
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "Familia."
         Top             =   1050
         Visible         =   0   'False
         Width           =   6525
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
         Left            =   2700
         TabIndex        =   12
         Top             =   840
         Width           =   1470
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filtrar por"
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
         Left            =   1530
         TabIndex        =   11
         Top             =   180
         Width           =   705
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   180
      TabIndex        =   13
      Top             =   2160
      Width           =   5415
      Begin MSComCtl2.DTPicker msk_fltFim 
         Height          =   315
         Left            =   3960
         TabIndex        =   9
         ToolTipText     =   "Data de emissão da nota fiscal."
         Top             =   240
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
         Format          =   198508545
         CurrentDate     =   39057
      End
      Begin MSComCtl2.DTPicker msk_fltInicio 
         Height          =   315
         Left            =   2100
         TabIndex        =   8
         ToolTipText     =   "Data de emissão da nota fiscal."
         Top             =   240
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
         Format          =   198508545
         CurrentDate     =   39057
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "até"
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
         Left            =   3585
         TabIndex        =   14
         Top             =   330
         Width           =   240
      End
   End
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   4620
      Top             =   150
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frmEstoque_fisico_abrir.frx":7092
      Count           =   1
   End
End
Attribute VB_Name = "frmEstoque_fisico_abrir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnFiltrar_Click()
On Error GoTo tratar_erro

ProcFiltrar

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

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

txtTexto = ""
If cmbfiltrarpor = "Família" Or cmbfiltrarpor = "Grupo" Or cmbfiltrarpor = "Local de armazenamento" Or cmbfiltrarpor = "Destino" Then
    txtTexto.Visible = False
    cmbfamilia.Visible = True
    If cmbfiltrarpor = "Família" Then
        ProcCarregaComboFamilia cmbfamilia, "Familia <> 'Null'", True
    ElseIf cmbfiltrarpor = "Grupo" Then
            ProcCarregaComboGrupoFamilia cmbfamilia, "Grupo <> 'Null'", True
        ElseIf cmbfiltrarpor = "Local de armazenamento" Then
                ProcCarregaComboLA cmbfamilia, True, True
            Else
                With cmbfamilia
                    .Clear
                    .AddItem "Interno"
                    .AddItem "Terceiros"
                End With
    End If
Else
    cmbfamilia.Visible = False
    txtTexto.Visible = True
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

With frmestoque_fisico
    DataFiltro = ""
    DataFiltroRel = ""
    If optPeriodo.Value = 1 Then
        DataFiltro = " and (EF.Data) Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "'"
        DataFiltroRel = " and {Estoque_fisico.Data} >= Date(" & Year(msk_fltInicio.Value) & "," & Month(msk_fltInicio.Value) & "," & Day(msk_fltInicio.Value) & ") and {Estoque_fisico.Data} <= Date(" & Year(msk_fltFim.Value) & "," & Month(msk_fltFim.Value) & "," & Day(msk_fltFim.Value) & ")"
    End If
    
    CamposFiltro = "EF.ID,EF.IDestoque, EF.ID_empresa, EF.Data, EF.Cod_ref, EF.valor_unitario, EF.Qtde_estoque, EF.qtde_fisica, EF.Etiqueta, EF.DtValidacao, P.Desenho, P.Descricao, P.Unidade"
    CamposFiltroTotal = "Sum(EF.valor_unitario) as Valor_Cofins_Prod, Sum(EF.Qtde_estoque) as Valor_Cofins_Serv, Sum(EF.valor_unitario * EF.Qtde_estoque) as Valor_CSLL_Prod, Sum(EF.Qtde_fisica) as Valor_CSLL_Serv, Sum(EF.valor_unitario * EF.Qtde_fisica) as Valor_INSS_Serv, Sum(EF.Qtde_estoque - EF.Qtde_fisica) as Valor_IPI, Sum(EF.valor_unitario * (EF.Qtde_estoque - EF.Qtde_fisica)) as Valor_IRPJ_Prod"
    INNERJOINTEXTOPADRAO = "(((Estoque_fisico EF LEFT JOIN projproduto P on EF.Codproduto = P.Codproduto) LEFT JOIN item_aplicacoes IA ON IA.Codproduto = P.Codproduto) LEFT JOIN Projfamilia PF on PF.Familia = P.Classe) LEFT JOIN Projproduto_fabricante PFAB ON PFAB.Codproduto = P.codproduto"
    INNERJOINTEXTO = "Select " & CamposFiltro & " from " & INNERJOINTEXTOPADRAO
    INNERJOINTEXTOSUM = "Select " & CamposFiltroTotal & " from " & INNERJOINTEXTOPADRAO
    TextoFiltroPadrao1 = "EF.ID_empresa = " & .Cmb_empresa.ItemData(.Cmb_empresa.ListIndex) & DataFiltro
    TextoFiltroPadrao = TextoFiltroPadrao1 & " group by " & CamposFiltro & " order by EF.Data desc"
    TextoFiltroPadraoRel = "{Estoque_fisico.ID_empresa} = " & .Cmb_empresa.ItemData(.Cmb_empresa.ListIndex) & DataFiltroRel
        
    If txtTexto.Visible = True And txtTexto.Text <> "" Or cmbfamilia.Visible = True And cmbfamilia <> "" Then
        If cmbfiltrarpor = "Família" Then
                .Sql_Estoque_Fisico_Localizar = INNERJOINTEXTO & " where P.Classe = '" & cmbfamilia & "' and " & TextoFiltroPadrao
                .Sql_Estoque_Fisico_LocalizarTotal = INNERJOINTEXTOSUM & " where P.Classe = '" & cmbfamilia & "' and " & TextoFiltroPadrao1
                .FormulaRel_Estoque_Fisico = "{projproduto.classe} = '" & cmbfamilia & "' and " & TextoFiltroPadraoRel
            ElseIf cmbfiltrarpor = "Grupo" Then
                    .Sql_Estoque_Fisico_Localizar = INNERJOINTEXTO & " where PF.Grupo = '" & cmbfamilia & "' and " & TextoFiltroPadrao
                    .Sql_Estoque_Fisico_LocalizarTotal = INNERJOINTEXTOSUM & " where PF.Grupo = '" & cmbfamilia & "' and " & TextoFiltroPadrao1
                    .FormulaRel_Estoque_Fisico = "{Projfamilia.Grupo} = '" & cmbfamilia & "' and " & TextoFiltroPadraoRel
                ElseIf cmbfiltrarpor = "Local de armazenamento" Then
                        .Sql_Estoque_Fisico_Localizar = INNERJOINTEXTO & " where EF.Local_armaz = '" & cmbfamilia & "' and " & TextoFiltroPadrao
                        .Sql_Estoque_Fisico_LocalizarTotal = INNERJOINTEXTOSUM & " where EF.Local_armaz = '" & cmbfamilia & "' and " & TextoFiltroPadrao1
                        .FormulaRel_Estoque_Fisico = "{Estoque_fisico.Local_armaz} = '" & cmbfamilia & "' and " & TextoFiltroPadraoRel
                    ElseIf cmbfiltrarpor = "Destino" Then
                            .Sql_Estoque_Fisico_Localizar = INNERJOINTEXTO & " where EF.Destino = '" & cmbfamilia & "' and " & TextoFiltroPadrao
                            .Sql_Estoque_Fisico_LocalizarTotal = INNERJOINTEXTOSUM & " where EF.Destino = '" & cmbfamilia & "' and " & TextoFiltroPadrao1
                            .FormulaRel_Estoque_Fisico = "{Estoque_fisico.Destino} = '" & cmbfamilia & "' and " & TextoFiltroPadraoRel
                    ElseIf cmbfiltrarpor = "RE" Then
                            .Sql_Estoque_Fisico_Localizar = INNERJOINTEXTO & " where EF.IDestoque = " & txtTexto & " and " & TextoFiltroPadrao
                            .Sql_Estoque_Fisico_LocalizarTotal = INNERJOINTEXTOSUM & " where EF.IDestoque = " & txtTexto & " and " & TextoFiltroPadrao1
                            .FormulaRel_Estoque_Fisico = "{Estoque_fisico.IDestoque} = " & txtTexto & " and " & TextoFiltroPadraoRel
                        Else
                            Select Case cmbfiltrarpor
                                Case "Código interno": TextoFiltro = "P.Desenho"
                                Case "Código de referência": TextoFiltro = "IA.n_referencia"
                                Case "Descrição": TextoFiltro = "P.Descricao"
                                Case "Etiqueta": TextoFiltro = "EF.Etiqueta"
                                Case "Responsável": TextoFiltro = "EF.Responsavel"
                                Case "Lote": TextoFiltro = "EF.Lote"
                                Case "Corrida": TextoFiltro = "EF.Corrida"
                                Case "Certificado": TextoFiltro = "EF.Certificado"
                                Case "Part number": TextoFiltro = "PFAB.Part_number"
                            End Select
                            .Sql_Estoque_Fisico_Localizar = INNERJOINTEXTO & " where " & TextoFiltro & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and " & TextoFiltroPadrao
                            .Sql_Estoque_Fisico_LocalizarTotal = INNERJOINTEXTOSUM & " where " & TextoFiltro & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and " & TextoFiltroPadrao1
                            
                            If Left(TextoFiltro, 2) = "P." Then
                                TextoFiltroRel = Replace(TextoFiltro, "P.", "projproduto.")
                            ElseIf Left(TextoFiltro, 2) = "EF" Then
                                    TextoFiltroRel = Replace(TextoFiltro, "EF.", "Estoque_fisico.")
                                ElseIf Left(TextoFiltro, 2) = "PF" Then
                                        TextoFiltroRel = Replace(TextoFiltro, "PF.", "Projproduto_fabricante.")
                                    Else
                                        TextoFiltroRel = Replace(TextoFiltro, "IA.", "item_aplicacoes.")
                            End If
                            .FormulaRel_Estoque_Fisico = "{" & TextoFiltroRel & "}" & FunVerifTipoFiltroIMFRel(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and " & TextoFiltroPadraoRel
        End If
    Else
        .Sql_Estoque_Fisico_Localizar = INNERJOINTEXTO & " where " & TextoFiltroPadrao
        'Debug.print .Sql_Estoque_Fisico_Localizar
        
        .Sql_Estoque_Fisico_LocalizarTotal = INNERJOINTEXTOSUM & " where " & TextoFiltroPadrao1
        .FormulaRel_Estoque_Fisico = TextoFiltroPadraoRel
    End If
    'Debug.print .Sql_Estoque_Fisico_Localizar
    .ProcAtualizalista (1)


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
    Case vbKeyEscape: ProcSair
    Case vbKeyF2: ProcFiltrar
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcFiltroPadrao cmbfiltrarpor, Optmeio, Optfim, optIgual, frmestoque_fisico.Cmb_empresa.ItemData(frmestoque_fisico.Cmb_empresa.ListIndex), "Produtos/Serviços", "T", True
If Permitido = False Then cmbfiltrarpor = "Código interno"

msk_fltFim.Value = Date
msk_fltInicio.Value = Date

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

Private Sub optPeriodo_Click()
On Error GoTo tratar_erro

If optPeriodo.Value = 1 Then
    Frame2.Enabled = True
    msk_fltInicio.SetFocus
Else
    msk_fltInicio.Value = Date
    msk_fltFim.Value = Date
    Frame2.Enabled = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
