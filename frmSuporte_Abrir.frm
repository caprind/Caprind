VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmSuporte_Abrir 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Outros - Suporte técnico - Localizar"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   315
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
   ScaleHeight     =   3690
   ScaleWidth      =   8925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Cmb_empresa 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   330
      ItemData        =   "frmSuporte_Abrir.frx":0000
      Left            =   1170
      List            =   "frmSuporte_Abrir.frx":0002
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Empresa."
      Top             =   1110
      Width           =   7485
   End
   Begin VB.CheckBox Chk_prazo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Prazo"
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
      Left            =   1440
      TabIndex        =   9
      Top             =   3270
      Width           =   795
   End
   Begin VB.CheckBox Chk_conclusao 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Conclusão"
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
      Left            =   2400
      TabIndex        =   10
      Top             =   3270
      Width           =   1305
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
      Left            =   270
      TabIndex        =   4
      Top             =   3270
      Width           =   1005
   End
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   6900
      Top             =   120
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frmSuporte_Abrir.frx":0004
      Count           =   1
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   1515
      Left            =   55
      TabIndex        =   13
      Top             =   1470
      Width           =   8805
      Begin VB.Frame Frame11 
         BackColor       =   &H00E0E0E0&
         Height          =   510
         Left            =   3810
         TabIndex        =   21
         Top             =   210
         WhatsThisHelpID =   210
         Width           =   4785
         Begin VB.OptionButton optIgual 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Igual"
            Height          =   255
            Left            =   3930
            TabIndex        =   8
            Top             =   180
            Width           =   705
         End
         Begin VB.OptionButton Optmeio 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Meio frase"
            Height          =   255
            Left            =   1470
            TabIndex        =   6
            Top             =   180
            Width           =   1275
         End
         Begin VB.OptionButton Optinicio 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Início frase"
            Height          =   255
            Left            =   180
            TabIndex        =   5
            Top             =   180
            Value           =   -1  'True
            Width           =   1275
         End
         Begin VB.OptionButton Optfim 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Fim frase"
            Height          =   255
            Left            =   2760
            TabIndex        =   7
            Top             =   180
            Width           =   1155
         End
      End
      Begin VB.ComboBox cmbfiltrarpor 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         ItemData        =   "frmSuporte_Abrir.frx":21F1
         Left            =   180
         List            =   "frmSuporte_Abrir.frx":2204
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
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
         TabIndex        =   2
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
         ItemData        =   "frmSuporte_Abrir.frx":2235
         Left            =   180
         List            =   "frmSuporte_Abrir.frx":2237
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
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
         TabIndex        =   15
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
         TabIndex        =   14
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
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   675
      Left            =   55
      TabIndex        =   17
      Top             =   2970
      Width           =   8805
      Begin MSComCtl2.DTPicker msk_fltFim 
         Height          =   315
         Left            =   7320
         TabIndex        =   12
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
         Format          =   489422849
         CurrentDate     =   39057
      End
      Begin MSComCtl2.DTPicker msk_fltInicio 
         Height          =   315
         Left            =   5430
         TabIndex        =   11
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
         Format          =   489422849
         CurrentDate     =   39057
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
         Height          =   285
         Left            =   6915
         TabIndex        =   19
         Top             =   240
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
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
         Height          =   285
         Left            =   5070
         TabIndex        =   18
         Top             =   240
         Width           =   300
      End
   End
   Begin VB.Label Label44 
      AutoSize        =   -1  'True
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
      Left            =   270
      TabIndex        =   20
      Top             =   1110
      Width           =   825
   End
End
Attribute VB_Name = "frmSuporte_Abrir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Chk_conclusao_Click()
On Error GoTo tratar_erro

If Chk_conclusao.Value = 1 Then
    Chk_emissao.Value = 0
    Chk_prazo.Value = 0
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

Private Sub Chk_emissao_Click()
On Error GoTo tratar_erro

If Chk_emissao.Value = 1 Then
    Chk_prazo.Value = 0
    Chk_conclusao.Value = 0
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

Private Sub Chk_prazo_Click()
On Error GoTo tratar_erro

If Chk_prazo.Value = 1 Then
    Chk_emissao.Value = 0
    Chk_conclusao.Value = 0
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
    If cmbfiltrarpor = "Status" Or cmbfiltrarpor = "Tipo" Or cmbfiltrarpor = "Solicitante" Then
        txtTexto.Visible = False
        .Visible = True
        .Clear
        If cmbfiltrarpor = "Status" Then
            .AddItem "Aberto"
            .AddItem "Em análise"
            .AddItem "Orçado"
            .AddItem "Aprovado"
            .AddItem "Em desenvolvimento"
            .AddItem "Concluído"
            .AddItem "Cancelado"
        ElseIf cmbfiltrarpor = "Tipo" Then
                .AddItem "Bug"
                .AddItem "Configuração"
                .AddItem "Customização"
                .AddItem "Dúvida"
                .AddItem "Instalação"
                .AddItem "Outros"
                .AddItem "Reindexação"
            Else
                FunAbreBDSite
                Set TBClientes = New ADODB.Recordset
                TBClientes.Open "Select U.* from usuarios U INNER JOIN Chamado C ON U.ID = C.ID_usuario where C.ID_cliente = " & FunVerifIDEmpresaSite(Cmb_empresa.ItemData(Cmb_empresa.ListIndex)) & " group by U.ID, U.Usuario order by U.Usuario", ConexaoMySql, adOpenKeyset, adLockOptimistic
                If TBClientes.EOF = False Then
                    Do While TBClientes.EOF = False
                        .AddItem TBClientes!Usuario
                        .ItemData(.NewIndex) = TBClientes!ID
                        TBClientes.MoveNext
                    Loop
                End If
                TBClientes.Close
                FunFechaBDSite
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

With frmSuporte
    If Chk_emissao.Value = 1 Then
        DataTexto = "Data"
    ElseIf Chk_prazo.Value = 1 Then
            DataTexto = "Prazo"
        Else
            DataTexto = "Data_conclusao"
    End If
    
    DataFiltro = ""
    DataFiltroRel = ""
    If Chk_emissao.Value = 1 Or Chk_prazo.Value = 1 Or Chk_conclusao.Value = 1 Then
        DataFiltro = " and " & DataTexto & " Between '" & Format(msk_fltInicio.Value, "yyyy/mm/dd") & "' And '" & Format(msk_fltFim.Value, "yyyy/mm/dd") & "'"
        DataFiltroRel = " and {Chamado." & DataTexto & "} >= Date(" & Year(msk_fltInicio) & "," & Month(msk_fltInicio) & "," & Day(msk_fltInicio) & ") and {Chamado." & DataTexto & "} <= Date(" & Year(msk_fltFim) & "," & Month(msk_fltFim) & "," & Day(msk_fltFim) & ")"
    End If
    If txtTexto.Visible = True And txtTexto <> "" Or cmbfamilia.Visible = True And cmbfamilia <> "" Then
        If cmbfiltrarpor = "Status" Then
            .Sql_Atendimento_Localizar = "Select * from Chamado where ID_cliente = " & FunVerifIDEmpresaSite(Cmb_empresa.ItemData(Cmb_empresa.ListIndex)) & " and Status = '" & cmbfamilia & "'" & DataFiltro & " order by ID desc"
            .FormulaRel_Suporte_Tecnico = "{Chamado.ID_cliente} = " & FunVerifIDEmpresaSite(Cmb_empresa.ItemData(Cmb_empresa.ListIndex)) & " and {Chamado.Status} = '" & cmbfamilia & "'" & DataFiltroRel
        ElseIf cmbfiltrarpor = "Tipo" Then
                If cmbfamilia = "Customização" Then
                    TextoFiltro = "Tipo = 'C'"
                    TextoFiltroRel = "{Chamado.Tipo} = 'C'"
                Else
                    TextoFiltro = "Tipo = 'B'"
                    TextoFiltroRel = "{Chamado.Tipo} = 'B'"
                End If
                .Sql_Atendimento_Localizar = "Select * from Chamado where ID_cliente = " & FunVerifIDEmpresaSite(Cmb_empresa.ItemData(Cmb_empresa.ListIndex)) & " and " & TextoFiltro & DataFiltro & " order by ID desc"
                .FormulaRel_Suporte_Tecnico = "{Chamado.ID_cliente} = " & FunVerifIDEmpresaSite(Cmb_empresa.ItemData(Cmb_empresa.ListIndex)) & " and " & TextoFiltroRel & DataFiltroRel
            ElseIf cmbfiltrarpor = "Solicitante" Then
                    .Sql_Atendimento_Localizar = "Select CH.* from (Chamado CH INNER JOIN Clientes C ON CH.ID_Cliente = C.ID) INNER JOIN usuarios U ON U.ID = CH.ID_usuario where CH.ID_cliente = " & FunVerifIDEmpresaSite(Cmb_empresa.ItemData(Cmb_empresa.ListIndex)) & " and CH.ID_usuario = " & cmbfamilia.ItemData(cmbfamilia.ListIndex) & DataFiltro & " order by CH.ID desc"
                    .FormulaRel_Suporte_Tecnico = "{Chamado.ID_cliente} = " & FunVerifIDEmpresaSite(Cmb_empresa.ItemData(Cmb_empresa.ListIndex)) & " and {Chamado.ID_usuario} = " & cmbfamilia.ItemData(cmbfamilia.ListIndex) & DataFiltroRel
                Else
                    If cmbfiltrarpor = "Código" Then TextoFiltro = "Numero" Else TextoFiltro = "Proposta"
                    .Sql_Atendimento_Localizar = "Select * from Chamado where ID_cliente = " & FunVerifIDEmpresaSite(Cmb_empresa.ItemData(Cmb_empresa.ListIndex)) & " and " & TextoFiltro & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & DataFiltro & " order by ID desc"
                    .FormulaRel_Suporte_Tecnico = "{Chamado.ID_cliente} = " & FunVerifIDEmpresaSite(Cmb_empresa.ItemData(Cmb_empresa.ListIndex)) & " and {Chamado." & TextoFiltro & "}" & FunVerifTipoFiltroIMFRel(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & DataFiltroRel
        End If
    Else
        .Sql_Atendimento_Localizar = "Select * from Chamado where ID_cliente = " & FunVerifIDEmpresaSite(Cmb_empresa.ItemData(Cmb_empresa.ListIndex)) & DataFiltro & " order by ID desc"
        .FormulaRel_Suporte_Tecnico = "{Chamado.ID_cliente} = " & FunVerifIDEmpresaSite(Cmb_empresa.ItemData(Cmb_empresa.ListIndex)) & DataFiltroRel
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

ProcCarregaToolBar1 Me, 8805, 5, True
ProcCarregaComboEmpresa Cmb_empresa, False
cmbfiltrarpor = "Código"
msk_fltFim.Value = Date
msk_fltInicio.Value = Date
txtTexto.Visible = True
cmbfamilia.Visible = False

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
