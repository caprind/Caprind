VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmCompras_Requisicao_abrir 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Outros - Solicitação - Localizar"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8895
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
   ScaleHeight     =   3720
   ScaleWidth      =   8895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Cmb_empresa 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   330
      ItemData        =   "frmCompras_Requisicao_abrir.frx":0000
      Left            =   1170
      List            =   "frmCompras_Requisicao_abrir.frx":0002
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Empresa."
      Top             =   1110
      Width           =   7545
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
      TabIndex        =   8
      Top             =   3270
      Width           =   1005
   End
   Begin VB.CheckBox Chk_autorizacao 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Autorização"
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
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   1515
      Left            =   55
      TabIndex        =   12
      Top             =   1470
      Width           =   8805
      Begin VB.Frame Frame11 
         BackColor       =   &H00E0E0E0&
         Height          =   510
         Left            =   3810
         TabIndex        =   20
         Top             =   210
         WhatsThisHelpID =   210
         Width           =   4785
         Begin VB.OptionButton optIgual 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Igual"
            Height          =   255
            Left            =   3930
            TabIndex        =   7
            Top             =   180
            Width           =   705
         End
         Begin VB.OptionButton Optmeio 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Meio frase"
            Height          =   255
            Left            =   1470
            TabIndex        =   5
            Top             =   180
            Width           =   1275
         End
         Begin VB.OptionButton Optinicio 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Início frase"
            Height          =   255
            Left            =   180
            TabIndex        =   4
            Top             =   180
            Value           =   -1  'True
            Width           =   1275
         End
         Begin VB.OptionButton Optfim 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Fim frase"
            Height          =   255
            Left            =   2760
            TabIndex        =   6
            Top             =   180
            Width           =   1155
         End
      End
      Begin VB.ComboBox cmbfiltrarpor 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         ItemData        =   "frmCompras_Requisicao_abrir.frx":0004
         Left            =   180
         List            =   "frmCompras_Requisicao_abrir.frx":0020
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
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
         Left            =   180
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "Familia."
         Top             =   1050
         Visible         =   0   'False
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
         TabIndex        =   14
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
         Left            =   1537
         TabIndex        =   13
         Top             =   180
         Width           =   840
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   675
      Left            =   55
      TabIndex        =   15
      Top             =   3000
      Width           =   8805
      Begin MSComCtl2.DTPicker msk_fltFim 
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
         Format          =   488570881
         CurrentDate     =   39057
      End
      Begin MSComCtl2.DTPicker msk_fltInicio 
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
         Format          =   488570881
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
         TabIndex        =   17
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
         TabIndex        =   16
         Top             =   240
         Width           =   300
      End
   End
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   4620
      Top             =   150
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frmCompras_Requisicao_abrir.frx":0071
      Count           =   1
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
      TabIndex        =   19
      Top             =   1110
      Width           =   825
   End
End
Attribute VB_Name = "frmCompras_Requisicao_abrir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Chk_autorizacao_Click()
On Error GoTo tratar_erro

If Chk_autorizacao.Value = 1 Then
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

Private Sub Chk_emissao_Click()
On Error GoTo tratar_erro

If Chk_emissao.Value = 1 Then
    Chk_autorizacao.Value = 0
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

If cmbfiltrarpor = "Status" Or cmbfiltrarpor = "Família" Then
    txtTexto.Visible = False
    With cmbfamilia
        .Visible = True
        .Clear
        If cmbfiltrarpor = "Status" Then
            .AddItem "ABERTA"
            .AddItem "CANCELADA"
            .AddItem "LIBERADA"
        Else
            ProcCarregaComboFamilia cmbfamilia, "familia <> 'Null' and compras = 'True'", False
        End If
    End With
Else
    txtTexto.Visible = True
    cmbfamilia.Visible = False
    
    If txtTexto <> "" And (cmbfiltrarpor = "Ordem" Or cmbfiltrarpor = "OS") Then
        VerifNumero = txtTexto
        ProcVerificaNumero
        If VerifNumero = False Then
            txtTexto = ""
            txtTexto.SetFocus
            Exit Sub
        End If
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

If Chk_autorizacao.Value = 1 Or Chk_emissao.Value = 1 Then
    If Chk_emissao.Value = 1 Then Data_Solicitacao = "Compras_requisicao.Data_Solicitacao" Else Data_Solicitacao = "Compras_requisicao.Data_Autorizacao"
    DataFiltro = "" & Data_Solicitacao & " Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "'"
Else
    DataFiltro = "Compras_requisicao.Status <> 'Null'"
End If

Campos = "Compras_requisicao.ID_Requisicao, Compras_requisicao.ID_empresa, Compras_requisicao.Requisicaotexto, Compras_requisicao.Data_Solicitacao, Compras_requisicao.solicitado, Compras_requisicao.setorsolic, Compras_requisicao.Data_autorizacao, Compras_requisicao.Autorizado, Compras_requisicao.setorautor, Compras_requisicao.Status, Compras_requisicao.DtValidacao"

With frmCompras_Requisicao
    If txtTexto.Visible = True And txtTexto <> "" Or cmbfamilia.Visible = True And cmbfamilia <> "" Then
        If cmbfiltrarpor = "Status" Then
            .StrSql_solicitacao = "Select " & Campos & " from Compras_requisicao where status = '" & cmbfamilia & "' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and " & DataFiltro & " group by " & Campos & " order by id_requisicao desc"
        ElseIf cmbfiltrarpor = "Família" Then
                .StrSql_solicitacao = "Select " & Campos & " FROM Compras_requisicao INNER JOIN Compras_pedido_lista ON Compras_requisicao.ID_Requisicao = Compras_pedido_lista.ID_Requisicao where Compras_pedido_lista.Familia = '" & cmbfamilia & "' and compras_requisicao.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and " & DataFiltro & " group by " & Campos & " order by Compras_requisicao.ID_Requisicao desc"
            Else
                Select Case cmbfiltrarpor
                    Case "Solicitação": TextoFiltro = "Compras_requisicao.Requisicaotexto"
                    Case "Código interno": TextoFiltro = "Compras_pedido_lista.desenho"
                    Case "Descrição": TextoFiltro = "Compras_pedido_lista.descricao"
                    Case "Detalhe": TextoFiltro = "Compras_pedido_lista.Detalheitem"
                    Case "Ordem": TextoFiltro = "Compras_pedido_lista.Ordem"
                    Case "OS": TextoFiltro = "Compras_pedido_lista.OS"
                End Select
                If cmbfiltrarpor = "Solicitação" Then
                    If Optinicio.Value = True Then .StrSql_solicitacao = "Select " & Campos & " from Compras_requisicao where " & TextoFiltro & " like '" & txtTexto.Text & "%' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and " & DataFiltro & " group by " & Campos & " order by id_requisicao desc"
                    If Optmeio.Value = True Then .StrSql_solicitacao = "Select " & Campos & " from Compras_requisicao where " & TextoFiltro & " like '%" & txtTexto.Text & "%' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and " & DataFiltro & " group by " & Campos & " order by id_requisicao desc"
                    If Optfim.Value = True Then .StrSql_solicitacao = "Select " & Campos & " from Compras_requisicao where " & TextoFiltro & " like '%" & txtTexto.Text & "' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and " & DataFiltro & " group by " & Campos & " order by id_requisicao desc"
                    If optIgual.Value = True Then .StrSql_solicitacao = "Select " & Campos & " from Compras_requisicao where " & TextoFiltro & " = '" & txtTexto.Text & "' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and " & DataFiltro & " group by " & Campos & " order by id_requisicao desc"
                Else
                    If cmbfiltrarpor = "Ordem" Or cmbfiltrarpor = "OS" Then
                        .StrSql_solicitacao = "Select " & Campos & " FROM Compras_requisicao INNER JOIN Compras_pedido_lista ON Compras_requisicao.ID_Requisicao = Compras_pedido_lista.ID_Requisicao where " & TextoFiltro & " = " & txtTexto & " and compras_requisicao.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and " & DataFiltro & " group by " & Campos & " order by Compras_requisicao.ID_Requisicao desc"
                    Else
                        If Optinicio.Value = True Then .StrSql_solicitacao = "Select " & Campos & " FROM Compras_requisicao INNER JOIN Compras_pedido_lista ON Compras_requisicao.ID_Requisicao = Compras_pedido_lista.ID_Requisicao where " & TextoFiltro & " like '" & txtTexto & "%' and compras_requisicao.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and " & DataFiltro & " group by " & Campos & " order by Compras_requisicao.ID_Requisicao desc"
                        If Optmeio.Value = True Then .StrSql_solicitacao = "Select " & Campos & " FROM Compras_requisicao INNER JOIN Compras_pedido_lista ON Compras_requisicao.ID_Requisicao = Compras_pedido_lista.ID_Requisicao where " & TextoFiltro & " like '%" & txtTexto & "%' and compras_requisicao.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and " & DataFiltro & " group by " & Campos & " order by Compras_requisicao.ID_Requisicao desc"
                        If Optfim.Value = True Then .StrSql_solicitacao = "Select " & Campos & " FROM Compras_requisicao INNER JOIN Compras_pedido_lista ON Compras_requisicao.ID_Requisicao = Compras_pedido_lista.ID_Requisicao where " & TextoFiltro & " like '%" & txtTexto & "' and compras_requisicao.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and " & DataFiltro & " group by " & Campos & " order by Compras_requisicao.ID_Requisicao desc"
                        If optIgual.Value = True Then .StrSql_solicitacao = "Select " & Campos & " FROM Compras_requisicao INNER JOIN Compras_pedido_lista ON Compras_requisicao.ID_Requisicao = Compras_pedido_lista.ID_Requisicao where " & TextoFiltro & " = '" & txtTexto & "' and compras_requisicao.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and " & DataFiltro & " group by " & Campos & " order by Compras_requisicao.ID_Requisicao desc"
                    End If
                End If
        End If
    Else
        .StrSql_solicitacao = "Select " & Campos & " from Compras_requisicao where ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and " & DataFiltro & " group by " & Campos & " order by id_requisicao desc"
    End If
    .ProcCarregaLista_Req (1)
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
    Case vbKeyF2: ProcFiltrar
    'Case vbKeyF1: ProcAjuda
    Case vbKeyEscape: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 8805, 5, True

cmbfiltrarpor = "Solicitação"
ProcCarregaComboEmpresa Cmb_empresa, False
msk_fltInicio = Date
msk_fltFim = Date

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

If txtTexto <> "" Then
    cmbfamilia.ListIndex = -1
    If cmbfiltrarpor = "Ordem" Or cmbfiltrarpor = "OS" Then
        VerifNumero = txtTexto
        ProcVerificaNumero
        If VerifNumero = False Then
            txtTexto = ""
            txtTexto.SetFocus
            Exit Sub
        End If
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
    Case 1: ProcFiltrar
    'Case 3: ProcAjuda
    Case 4: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

