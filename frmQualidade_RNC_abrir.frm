VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmQualidade_RNC_abrir 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Qualidade - RNC - Localizar"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8925
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   8925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Chk_implementacao 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Pr. implementação"
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
      Left            =   1410
      TabIndex        =   8
      Top             =   2730
      Width           =   1935
   End
   Begin VB.CheckBox Chk_emissao 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Emissão"
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
      TabIndex        =   7
      Top             =   2730
      Width           =   1005
   End
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   5820
      Top             =   150
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frmQualidade_RNC_abrir.frx":0000
      Count           =   1
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   1515
      Left            =   55
      TabIndex        =   11
      Top             =   990
      Width           =   8805
      Begin VB.Frame Frame11 
         BackColor       =   &H00E0E0E0&
         Height          =   510
         Left            =   3810
         TabIndex        =   17
         Top             =   210
         Width           =   4785
         Begin VB.OptionButton optIgual 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Igual"
            Height          =   255
            Left            =   3930
            TabIndex        =   6
            Top             =   180
            Width           =   705
         End
         Begin VB.OptionButton Optmeio 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Meio frase"
            Height          =   255
            Left            =   1470
            TabIndex        =   4
            Top             =   180
            Width           =   1275
         End
         Begin VB.OptionButton Optinicio 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Início frase"
            Height          =   255
            Left            =   180
            TabIndex        =   3
            Top             =   180
            Value           =   -1  'True
            Width           =   1275
         End
         Begin VB.OptionButton Optfim 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Fim frase"
            Height          =   255
            Left            =   2760
            TabIndex        =   5
            Top             =   180
            Width           =   1155
         End
      End
      Begin VB.ComboBox cmbfiltrarpor 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         ItemData        =   "frmQualidade_RNC_abrir.frx":21F1
         Left            =   180
         List            =   "frmQualidade_RNC_abrir.frx":2216
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Opções para filtro."
         Top             =   390
         Width           =   3525
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
         ItemData        =   "frmQualidade_RNC_abrir.frx":22A5
         Left            =   180
         List            =   "frmQualidade_RNC_abrir.frx":22A7
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "Familia."
         Top             =   1050
         Width           =   8415
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
         TabIndex        =   12
         Top             =   180
         Width           =   840
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   55
      TabIndex        =   16
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
         Name            =   "Tahoma"
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
      ButtonUseMaskColor2=   0   'False
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
      ButtonEnabled5  =   0   'False
      ButtonIconSize5 =   32
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
      ButtonLeft5     =   110
      ButtonTop5      =   2
      ButtonWidth5    =   24
      ButtonHeight5   =   24
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   585
      Left            =   55
      TabIndex        =   14
      Top             =   2490
      Width           =   8805
      Begin MSComCtl2.DTPicker msk_fltFim 
         Height          =   315
         Left            =   7200
         TabIndex        =   10
         ToolTipText     =   "Data final para pesquisa."
         Top             =   180
         Width           =   1395
         _ExtentX        =   2461
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
         Format          =   489816065
         CurrentDate     =   39057
      End
      Begin MSComCtl2.DTPicker msk_fltInicio 
         Height          =   315
         Left            =   5430
         TabIndex        =   9
         ToolTipText     =   "Data início para pesquisa."
         Top             =   180
         Width           =   1395
         _ExtentX        =   2461
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
         Format          =   489816067
         CurrentDate     =   39057
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "à"
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
         Left            =   6945
         TabIndex        =   15
         Top             =   240
         Width           =   90
      End
   End
End
Attribute VB_Name = "frmQualidade_RNC_abrir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Chk_emissao_Click()
On Error GoTo tratar_erro

If Chk_emissao.Value = 1 Then
    Chk_implementacao.Value = 0
    Frame2.Enabled = True
    msk_fltInicio.SetFocus
Else
    Frame2.Enabled = False
    msk_fltInicio.Value = Date
    msk_fltFim.Value = Date
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_implementacao_Click()
On Error GoTo tratar_erro

If Chk_implementacao.Value = 1 Then
    Chk_emissao.Value = 0
    Frame2.Enabled = True
    msk_fltInicio.SetFocus
Else
    Frame2.Enabled = False
    msk_fltInicio.Value = Date
    msk_fltFim.Value = Date
End If

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

With cmbfamilia
    If cmbfiltrarpor = "Família" Or cmbfiltrarpor = "Status" Or cmbfiltrarpor = "Eficaz" Then
        txtTexto.Visible = False
        .Visible = True
        .Clear
        If cmbfiltrarpor = "Família" Then
            ProcCarregaComboFamilia cmbfamilia, "familia <> 'Null'", True
        ElseIf cmbfiltrarpor = "Eficaz" Then
                .AddItem ""
                .AddItem "Sim"
                .AddItem "Não"
            Else
                .AddItem "Ativa"
                .AddItem "Cancelada"
                .AddItem "Concluida"
                .AddItem "Enviada"
        End If
    Else
        txtTexto.Visible = True
        .Visible = False
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

ProcExcluirDadosProducaoRelatoriosTotal

With frmQualidade_RNC
    If Chk_emissao.Value = 1 Then DataTexto = "CQ_RNC.Data" Else DataTexto = "CQ_RNC.Data4"
    If Chk_emissao.Value = 1 Or Chk_implementacao.Value = 1 Then
        DataFiltro = "(" & DataTexto & ") Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "'"
        DataFiltroRel = "{" & DataTexto & "} >= Date(" & Year(msk_fltInicio.Value) & "," & Month(msk_fltInicio.Value) & "," & Day(msk_fltInicio.Value) & ") and {" & DataTexto & "} <= Date(" & _
                                Year(msk_fltFim.Value) & "," & Month(msk_fltFim.Value) & "," & Day(msk_fltFim.Value) & ")"
    Else
        DataFiltro = "CQ_RNC.Desenho is not null"
        DataFiltroRel = "{CQ_RNC.Desenho} <> 'Null'"
    End If
    
    If txtTexto.Visible = True And txtTexto <> "" Or cmbfamilia.Visible = True And cmbfamilia <> "" Then
        If cmbfiltrarpor = "Família" Then
            .StrSql_CQ_RNC_Localizar = "Select CQ_RNC.* FROM CQ_RNC INNER JOIN projproduto ON CQ_RNC.desenho = projproduto.desenho where projproduto.classe = '" & cmbfamilia & "' and " & DataFiltro & " order by CQ_RNC.Data desc, CQ_RNC.ID desc"
            .FormulaRel_CQ_RNC = "{projproduto.classe} = '" & cmbfamilia & "' and " & DataFiltroRel & " and {Producao_Relatorios_Total.Responsavel} = '" & pubUsuario & "' and {Producao_Relatorios_Total.Modulo} = '" & Formulario & "'"
        ElseIf cmbfiltrarpor = "Status" Then
                .StrSql_CQ_RNC_Localizar = "Select * from CQ_RNC where Status = '" & cmbfamilia & "' and " & DataFiltro & " order by Data desc, ID desc"
                .FormulaRel_CQ_RNC = "{CQ_RNC.Status} = '" & cmbfamilia & "' and " & DataFiltroRel & " and {Producao_Relatorios_Total.Responsavel} = '" & pubUsuario & "' and {Producao_Relatorios_Total.Modulo} = '" & Formulario & "'"
            ElseIf cmbfiltrarpor = "Eficaz" Then
                    .StrSql_CQ_RNC_Localizar = "Select * FROM CQ_RNC where Eficaz = '" & cmbfamilia & "' and " & DataFiltro & " order by Data desc, ID desc"
                    .FormulaRel_CQ_RNC = "{CQ_RNC.Eficaz} = '" & cmbfamilia & "' and " & DataFiltroRel & " and {Producao_Relatorios_Total.Responsavel} = '" & pubUsuario & "' and {Producao_Relatorios_Total.Modulo} = '" & Formulario & "'"
                ElseIf cmbfiltrarpor = "Descrição" Then
                        If Optinicio.Value = True Then
                            .StrSql_CQ_RNC_Localizar = "Select CQ_RNC.* FROM CQ_RNC INNER JOIN projproduto ON CQ_RNC.desenho = projproduto.desenho where projproduto.descricao like '" & txtTexto & "%' and " & DataFiltro & " order by CQ_RNC.Data desc, CQ_RNC.ID desc"
                            .FormulaRel_CQ_RNC = "{projproduto.descricao} like '" & txtTexto & "*' and " & DataFiltroRel & " and {Producao_Relatorios_Total.Responsavel} = '" & pubUsuario & "' and {Producao_Relatorios_Total.Modulo} = '" & Formulario & "'"
                        End If
                        If Optmeio.Value = True Then
                            .StrSql_CQ_RNC_Localizar = "Select CQ_RNC.* FROM CQ_RNC INNER JOIN projproduto ON CQ_RNC.desenho = projproduto.desenho where projproduto.descricao like '%" & txtTexto & "%' and " & DataFiltro & " order by CQ_RNC.Data desc, CQ_RNC.ID desc"
                            .FormulaRel_CQ_RNC = "{projproduto.descricao} like '*" & txtTexto & "*' and " & DataFiltroRel & " and {Producao_Relatorios_Total.Responsavel} = '" & pubUsuario & "' and {Producao_Relatorios_Total.Modulo} = '" & Formulario & "'"
                        End If
                        If Optfim.Value = True Then
                            .StrSql_CQ_RNC_Localizar = "Select CQ_RNC.* FROM CQ_RNC INNER JOIN projproduto ON CQ_RNC.desenho = projproduto.desenho where projproduto.descricao like '%" & txtTexto & "' and " & DataFiltro & " order by CQ_RNC.Data desc, CQ_RNC.ID desc"
                            .FormulaRel_CQ_RNC = "{projproduto.descricao} like '*" & txtTexto & "' and " & DataFiltroRel & " and {Producao_Relatorios_Total.Responsavel} = '" & pubUsuario & "' and {Producao_Relatorios_Total.Modulo} = '" & Formulario & "'"
                        End If
                        If optIgual.Value = True Then
                            .StrSql_CQ_RNC_Localizar = "Select CQ_RNC.* FROM CQ_RNC INNER JOIN projproduto ON CQ_RNC.desenho = projproduto.desenho where projproduto.descricao = '" & txtTexto & "' and " & DataFiltro & " order by CQ_RNC.Data desc, CQ_RNC.ID desc"
                            .FormulaRel_CQ_RNC = "{projproduto.descricao} = '" & txtTexto & "' and " & DataFiltroRel & " and {Producao_Relatorios_Total.Responsavel} = '" & pubUsuario & "' and {Producao_Relatorios_Total.Modulo} = '" & Formulario & "'"
                        End If
                    ElseIf cmbfiltrarpor = "Fornecedor" Then
                            If Optinicio.Value = True Then
                                .StrSql_CQ_RNC_Localizar = "Select CQ_RNC.* FROM CQ_RNC INNER JOIN Compras_fornecedores ON CQ_RNC.id_forn = Compras_fornecedores.idcliente where Compras_fornecedores.nome_razao like '" & txtTexto & "%' and CQ_RNC.tipo = 'F' and " & DataFiltro & " order by CQ_RNC.Data desc, CQ_RNC.ID desc"
                                .FormulaRel_CQ_RNC = "{Compras_fornecedores.nome_razao} like '" & txtTexto & "*' and {CQ_RNC.tipo} = 'F' and " & DataFiltroRel & " and {Producao_Relatorios_Total.Responsavel} = '" & pubUsuario & "' and {Producao_Relatorios_Total.Modulo} = '" & Formulario & "'"
                            End If
                            If Optmeio.Value = True Then
                                .StrSql_CQ_RNC_Localizar = "Select CQ_RNC.* FROM CQ_RNC INNER JOIN Compras_fornecedores ON CQ_RNC.id_forn = Compras_fornecedores.idcliente where Compras_fornecedores.nome_razao like '%" & txtTexto & "%' and CQ_RNC.tipo = 'F' and " & DataFiltro & " order by CQ_RNC.Data desc, CQ_RNC.ID desc"
                                .FormulaRel_CQ_RNC = "{Compras_fornecedores.nome_razao} like '*" & txtTexto & "*' and {CQ_RNC.tipo} = 'F' and " & DataFiltroRel & " and {Producao_Relatorios_Total.Responsavel} = '" & pubUsuario & "' and {Producao_Relatorios_Total.Modulo} = '" & Formulario & "'"
                            End If
                            If Optfim.Value = True Then
                                .StrSql_CQ_RNC_Localizar = "Select CQ_RNC.* FROM CQ_RNC INNER JOIN Compras_fornecedores ON CQ_RNC.id_forn = Compras_fornecedores.idcliente where Compras_fornecedores.nome_razao like '%" & txtTexto & "' and CQ_RNC.tipo = 'F' and " & DataFiltro & " order by CQ_RNC.Data desc, CQ_RNC.ID desc"
                                .FormulaRel_CQ_RNC = "{Compras_fornecedores.nome_razao} like '*" & txtTexto & "' and {CQ_RNC.tipo} = 'F' and " & DataFiltroRel & " and {Producao_Relatorios_Total.Responsavel} = '" & pubUsuario & "' and {Producao_Relatorios_Total.Modulo} = '" & Formulario & "'"
                            End If
                            If optIgual.Value = True Then
                                .StrSql_CQ_RNC_Localizar = "Select CQ_RNC.* FROM CQ_RNC INNER JOIN Compras_fornecedores ON CQ_RNC.id_forn = Compras_fornecedores.idcliente where Compras_fornecedores.nome_razao = '" & txtTexto & "' and CQ_RNC.tipo = 'F' and " & DataFiltro & " order by CQ_RNC.Data desc, CQ_RNC.ID desc"
                                .FormulaRel_CQ_RNC = "{Compras_fornecedores.nome_razao} = '" & txtTexto & "' and {CQ_RNC.tipo} = 'F' and " & DataFiltroRel & " and {Producao_Relatorios_Total.Responsavel} = '" & pubUsuario & "' and {Producao_Relatorios_Total.Modulo} = '" & Formulario & "'"
                            End If
                        ElseIf cmbfiltrarpor = "Cliente" Then
                                If Optinicio.Value = True Then
                                    .StrSql_CQ_RNC_Localizar = "Select CQ_RNC.* FROM CQ_RNC INNER JOIN Clientes ON CQ_RNC.id_forn = Clientes.idcliente where Clientes.nomerazao like '" & txtTexto & "%' and CQ_RNC.tipo = 'C' and " & DataFiltro & " order by CQ_RNC.Data desc, CQ_RNC.ID desc"
                                    .FormulaRel_CQ_RNC = "{Clientes.nomerazao} like '" & txtTexto & "*' and {CQ_RNC.tipo} = 'C' and " & DataFiltroRel & " and {Producao_Relatorios_Total.Responsavel} = '" & pubUsuario & "' and {Producao_Relatorios_Total.Modulo} = '" & Formulario & "'"
                                End If
                                If Optmeio.Value = True Then
                                    .StrSql_CQ_RNC_Localizar = "Select CQ_RNC.* FROM CQ_RNC INNER JOIN Clientes ON CQ_RNC.id_forn = Clientes.idcliente where Clientes.nomerazao like '%" & txtTexto & "%' and CQ_RNC.tipo = 'C' and " & DataFiltro & " order by CQ_RNC.Data desc, CQ_RNC.ID desc"
                                    .FormulaRel_CQ_RNC = "{Clientes.nomerazao} like '*" & txtTexto & "*' and {CQ_RNC.tipo} = 'C' and " & DataFiltroRel & " and {Producao_Relatorios_Total.Responsavel} = '" & pubUsuario & "' and {Producao_Relatorios_Total.Modulo} = '" & Formulario & "'"
                                End If
                                If Optfim.Value = True Then
                                    .StrSql_CQ_RNC_Localizar = "Select CQ_RNC.* FROM CQ_RNC INNER JOIN Clientes ON CQ_RNC.id_forn = Clientes.idcliente where Clientes.nomerazao like '%" & txtTexto & "' and CQ_RNC.tipo = 'C' and " & DataFiltro & " order by CQ_RNC.Data desc, CQ_RNC.ID desc"
                                    .FormulaRel_CQ_RNC = "{Clientes.nomerazao} like '*" & txtTexto & "' and {CQ_RNC.tipo} = 'C' and " & DataFiltroRel & "and {Producao_Relatorios_Total.Responsavel} = '" & pubUsuario & "' and {Producao_Relatorios_Total.Modulo} = '" & Formulario & "'"
                                End If
                                If optIgual.Value = True Then
                                    .StrSql_CQ_RNC_Localizar = "Select CQ_RNC.* FROM CQ_RNC INNER JOIN Clientes ON CQ_RNC.id_forn = Clientes.idcliente where Clientes.nomerazao = '" & txtTexto & "' and CQ_RNC.tipo = 'C' and " & DataFiltro & " order by CQ_RNC.Data desc, CQ_RNC.ID desc"
                                    .FormulaRel_CQ_RNC = "{Clientes.nomerazao} = '" & txtTexto & "' and {CQ_RNC.tipo} = 'C' and " & DataFiltroRel & "and {Producao_Relatorios_Total.Responsavel} = '" & pubUsuario & "' and {Producao_Relatorios_Total.Modulo} = '" & Formulario & "'"
                                End If
                            Else
                                Select Case cmbfiltrarpor
                                    Case "RNC": TextoFiltro = "ID_texto"
                                    Case "Código interno": TextoFiltro = "desenho"
                                    Case "Doc. de referência": TextoFiltro = "Documento_ref"
                                    Case "Acompanhamento": TextoFiltro = "Acompanhamento"
                                    Case "Verificação da eficácia": TextoFiltro = "Eficacia"
                                End Select
                                If Optinicio.Value = True Then
                                    .StrSql_CQ_RNC_Localizar = "Select * from CQ_RNC where " & TextoFiltro & " like '" & txtTexto.Text & "%' and " & DataFiltro & " order by Data desc, ID desc"
                                    .FormulaRel_CQ_RNC = "{CQ_RNC." & TextoFiltro & "} like '*" & txtTexto & "' and " & DataFiltroRel & "and {Producao_Relatorios_Total.Responsavel} = '" & pubUsuario & "' and {Producao_Relatorios_Total.Modulo} = '" & Formulario & "'"
                                End If
                                If Optmeio.Value = True Then
                                    .StrSql_CQ_RNC_Localizar = "Select * from CQ_RNC where " & TextoFiltro & " like '%" & txtTexto.Text & "%' and " & DataFiltro & " order by Data desc, ID desc"
                                    .FormulaRel_CQ_RNC = "{CQ_RNC." & TextoFiltro & "} like '*" & txtTexto & "*' and " & DataFiltroRel & "and {Producao_Relatorios_Total.Responsavel} = '" & pubUsuario & "' and {Producao_Relatorios_Total.Modulo} = '" & Formulario & "'"
                                End If
                                If Optfim.Value = True Then
                                    .StrSql_CQ_RNC_Localizar = "Select * from CQ_RNC where " & TextoFiltro & " like '%" & txtTexto.Text & "' and " & DataFiltro & " order by Data desc, ID desc"
                                    .FormulaRel_CQ_RNC = "{CQ_RNC." & TextoFiltro & "} like '" & txtTexto & "*' and " & DataFiltroRel & "and {Producao_Relatorios_Total.Responsavel} = '" & pubUsuario & "' and {Producao_Relatorios_Total.Modulo} = '" & Formulario & "'"
                                End If
                                If optIgual.Value = True Then
                                    .StrSql_CQ_RNC_Localizar = "Select * from CQ_RNC where " & TextoFiltro & " = '" & txtTexto.Text & "' and " & DataFiltro & " order by Data desc, ID desc"
                                    .FormulaRel_CQ_RNC = "{CQ_RNC." & TextoFiltro & "} = '" & txtTexto & "' and " & DataFiltroRel & "and {Producao_Relatorios_Total.Responsavel} = '" & pubUsuario & "' and {Producao_Relatorios_Total.Modulo} = '" & Formulario & "'"
                                End If
        End If
    Else
        .StrSql_CQ_RNC_Localizar = "Select * from CQ_RNC where " & DataFiltro & " order by Data desc, ID desc"
        .FormulaRel_CQ_RNC = DataFiltroRel & "and {Producao_Relatorios_Total.Responsavel} = '" & pubUsuario & "' and {Producao_Relatorios_Total.Modulo} = '" & Formulario & "'"
    End If
    .ProcCarregaLista (1)
End With
Unload Me

Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from Producao_Relatorios_Total where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then TBGravar.AddNew
TBGravar!Data_inicial = msk_fltInicio.Value
TBGravar!Data_final = msk_fltFim.Value
TBGravar!Responsavel = pubUsuario
TBGravar!Modulo = Formulario
TBGravar.Update
TBGravar.Close

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

ProcCarregaToolBar1 Me, 8805, 5, True
cmbfiltrarpor = "RNC"
txtTexto.Visible = True
cmbfamilia.Visible = False
msk_fltInicio.Value = Date
msk_fltFim.Value = Date

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
    Case 4: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
