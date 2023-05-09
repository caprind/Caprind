VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmCompras_programacao_abrir 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Administrativo - Compras - Programação - Localizar"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10410
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   10410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   1515
      Left            =   60
      TabIndex        =   14
      Top             =   1440
      Width           =   10305
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0E0&
         Height          =   510
         Left            =   5340
         TabIndex        =   21
         Top             =   210
         Width           =   4785
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
            TabIndex        =   5
            Top             =   180
            Width           =   705
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
            TabIndex        =   3
            Top             =   180
            Width           =   1275
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
            TabIndex        =   2
            Top             =   180
            Value           =   -1  'True
            Width           =   1275
         End
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
            TabIndex        =   4
            Top             =   180
            Width           =   1155
         End
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
         ItemData        =   "frmCompras_programacao_abrir.frx":0000
         Left            =   180
         List            =   "frmCompras_programacao_abrir.frx":0016
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "Opções para filtro."
         Top             =   390
         Width           =   5085
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
         TabIndex        =   6
         ToolTipText     =   "Texto para pesquisa."
         Top             =   1050
         Width           =   9945
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
         MousePointer    =   99  'Custom
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   7
         ToolTipText     =   "Familia."
         Top             =   1050
         Width           =   9945
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
         Left            =   4417
         TabIndex        =   16
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
         Left            =   2302
         TabIndex        =   15
         Top             =   180
         Width           =   840
      End
   End
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
      ItemData        =   "frmCompras_programacao_abrir.frx":006A
      Left            =   1200
      List            =   "frmCompras_programacao_abrir.frx":006C
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Empresa."
      Top             =   1110
      Width           =   7215
   End
   Begin VB.CheckBox Chk_data_programa 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Programa"
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
      Height          =   225
      Left            =   270
      TabIndex        =   9
      Top             =   3240
      Width           =   1155
   End
   Begin VB.CheckBox Chk_compra_conf 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Compra confirmada"
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
      Height          =   255
      Left            =   8520
      TabIndex        =   8
      Top             =   1155
      Width           =   1695
   End
   Begin VB.CheckBox Chk_data_programacao 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Programação"
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
      Height          =   225
      Left            =   1560
      TabIndex        =   10
      Top             =   3240
      Width           =   1455
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   13
      Top             =   0
      Width           =   10305
      _ExtentX        =   18177
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
      ButtonUseMaskColor5=   0   'False
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   9090
         Top             =   150
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmCompras_programacao_abrir.frx":006E
         Count           =   1
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   675
      Left            =   60
      TabIndex        =   17
      Top             =   2970
      Width           =   10305
      Begin MSComCtl2.DTPicker txtData_fim 
         Height          =   315
         Left            =   8820
         TabIndex        =   12
         ToolTipText     =   "Final."
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
         Format          =   133496833
         CurrentDate     =   39057
      End
      Begin MSComCtl2.DTPicker txtData_inicio 
         Height          =   315
         Left            =   6930
         TabIndex        =   11
         ToolTipText     =   "Inicio."
         Top             =   210
         Width           =   1335
         _ExtentX        =   2355
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
         Format          =   133496833
         CurrentDate     =   39057
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   8415
         TabIndex        =   19
         Top             =   240
         Width           =   360
      End
      Begin VB.Label Label3 
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
         Left            =   6570
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
      Left            =   300
      TabIndex        =   20
      Top             =   1110
      Width           =   825
   End
End
Attribute VB_Name = "frmCompras_programacao_abrir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Chk_data_programa_Click()
On Error GoTo tratar_erro

If Chk_data_programa.Value = 1 Then
    Chk_data_programacao.Value = 0
    Frame1.Enabled = True
    txtData_inicio.SetFocus
Else
    Frame1.Enabled = False
    txtData_inicio.Value = Date
    txtData_fim.Value = Date
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

With txtData_inicio
    If FunVerificaDataFinal(txtData_fim.Value, .Value) = False Then
        .Value = Date
        .SetFocus
        Exit Sub
    End If
End With
With frmCompras_programacao
    TextoFiltroVenda = ""
    INNERJOINTEXTOREF = ""
    INNERJOINTEXTODATA = ""
    DataFiltro = ""
    
    If Chk_compra_conf.Value = 1 Or Chk_data_programacao.Value = 1 Or cmbfiltrarpor = "Código de referência" And txtTexto <> "" Then INNERJOINTEXTODATA = "INNER JOIN compras_programacao CPR ON CPR.ID = CP.ID"
        
    If Chk_compra_conf.Value = 1 Then TextoFiltroVenda = "and CPR.Firme = 'True'"
        
    If Chk_data_programa.Value = 1 Or Chk_data_programacao.Value = 1 Then
        If Chk_data_programa.Value = 1 Then
            DataFiltro = "and CP.Data Between '" & Format(txtData_inicio.Value, "Short Date") & "' And '" & Format(txtData_fim.Value, "Short Date") & "'"
        Else
            DataFiltro = "and CPR.data_inicio >= '" & txtData_inicio.Value & "' and CPR.data_fim <= '" & txtData_fim.Value & "'"
        End If
    End If
    
    CamposFiltro = "CP.ID, CP.Data, CP.Responsavel, CP.Programa, CP.ProgramaTexto, CP.Rev, CP.Data_rev, CP.Status, F.Nome_Razao"
    If txtTexto.Visible = True And txtTexto <> "" Or cmbfamilia.Visible = True And cmbfamilia <> "" Then
        If cmbfiltrarpor = "Família" Then
            .Sql_Programacao_Compras_Localizar = "Select " & CamposFiltro & " FROM compras_programa CP INNER JOIN Compras_fornecedores F ON CP.id_Forn = F.IDCliente INNER JOIN compras_programa_item CP ON CP.ID = CP.ID INNER JOIN projproduto P ON P.desenho = CP.Codigo " & INNERJOINTEXTODATA & " where CP.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and P.Classe = '" & cmbfamilia & "' " & DataFiltro & " " & TextoFiltroVenda & " group by " & CamposFiltro & " order by CP.Programa desc, CP.Rev desc"
        ElseIf cmbfiltrarpor = "Programa" Or cmbfiltrarpor = "Fornecedor" Then
                Select Case cmbfiltrarpor
                    Case "Programa": TextoFiltro = "CP.ProgramaTexto"
                    Case "Fornecedor": TextoFiltro = "F.Nome_Razao"
                End Select
                If Optinicio.Value = True Then .Sql_Programacao_Compras_Localizar = "Select " & CamposFiltro & " FROM compras_programa CP INNER JOIN Compras_fornecedores F ON CP.id_Forn = F.IDCliente " & INNERJOINTEXTODATA & " where CP.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and " & TextoFiltro & " like '" & txtTexto & "%' " & DataFiltro & " " & TextoFiltroVenda & " group by " & CamposFiltro & " order by CP.Programa desc, CP.Rev desc"
                If Optmeio.Value = True Then .Sql_Programacao_Compras_Localizar = "Select " & CamposFiltro & " FROM compras_programa CP INNER JOIN Compras_fornecedores F ON CP.id_Forn = F.IDCliente " & INNERJOINTEXTODATA & " where CP.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and " & TextoFiltro & " like '%" & txtTexto & "%' " & DataFiltro & " " & TextoFiltroVenda & " group by " & CamposFiltro & " order by CP.Programa desc, CP.Rev desc"
                If Optfim.Value = True Then .Sql_Programacao_Compras_Localizar = "Select " & CamposFiltro & " FROM compras_programa CP INNER JOIN Compras_fornecedores F ON CP.id_Forn = F.IDCliente " & INNERJOINTEXTODATA & " where CP.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and " & TextoFiltro & " like '%" & txtTexto & "' " & DataFiltro & " " & TextoFiltroVenda & " group by " & CamposFiltro & " order by CP.Programa desc, CP.Rev desc"
                If optIgual.Value = True Then .Sql_Programacao_Compras_Localizar = "Select " & CamposFiltro & " FROM compras_programa CP INNER JOIN Compras_fornecedores F ON CP.id_Forn = F.IDCliente " & INNERJOINTEXTODATA & " where CP.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and " & TextoFiltro & " = '" & txtTexto & "' " & DataFiltro & " " & TextoFiltroVenda & " group by " & CamposFiltro & " order by CP.Programa desc, CP.Rev desc"
            Else
                Select Case cmbfiltrarpor
                    Case "Código de referência":
                        INNERJOINTEXTOREF = "INNER JOIN item_aplicacoes IA ON IA.Codproduto = P.Codproduto"
                        TextoFiltro = "IA.n_referencia"
                    Case "Pedido cliente": TextoFiltro = "CPR.PCCliente"
                    Case "Código interno": TextoFiltro = "P.Desenho"
                    Case "Descrição": TextoFiltro = "P.Descricao_tecnica"
                End Select
                If Optinicio.Value = True Then .Sql_Programacao_Compras_Localizar = "Select " & CamposFiltro & " FROM compras_programa CP INNER JOIN Compras_fornecedores F ON CP.id_Forn = F.IDCliente INNER JOIN compras_programa_item CP ON CP.ID = CP.ID INNER JOIN projproduto P ON P.desenho = CP.Codigo " & INNERJOINTEXTOREF & " " & INNERJOINTEXTODATA & " where CP.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and " & TextoFiltro & " like '" & txtTexto & "%' " & DataFiltro & " " & TextoFiltroVenda & " group by " & CamposFiltro & " order by CP.Programa desc, CP.Rev desc"
                If Optmeio.Value = True Then .Sql_Programacao_Compras_Localizar = "Select " & CamposFiltro & " FROM compras_programa CP INNER JOIN Compras_fornecedores F ON CP.id_Forn = F.IDCliente INNER JOIN compras_programa_item CP ON CP.ID = CP.ID INNER JOIN projproduto P ON P.desenho = CP.Codigo " & INNERJOINTEXTOREF & " " & INNERJOINTEXTODATA & " where CP.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and " & TextoFiltro & " like '%" & txtTexto & "%' " & DataFiltro & " " & TextoFiltroVenda & " group by " & CamposFiltro & " order by CP.Programa desc, CP.Rev desc"
                If Optfim.Value = True Then .Sql_Programacao_Compras_Localizar = "Select " & CamposFiltro & " FROM compras_programa CP INNER JOIN Compras_fornecedores F ON CP.id_Forn = F.IDCliente INNER JOIN compras_programa_item CP ON CP.ID = CP.ID INNER JOIN projproduto P ON P.desenho = CP.Codigo " & INNERJOINTEXTOREF & " " & INNERJOINTEXTODATA & " where CP.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and " & TextoFiltro & " like '%" & txtTexto & "' " & DataFiltro & " " & TextoFiltroVenda & " group by " & CamposFiltro & " order by CP.Programa desc, CP.Rev desc"
                If optIgual.Value = True Then .Sql_Programacao_Compras_Localizar = "Select " & CamposFiltro & " FROM compras_programa CP INNER JOIN Compras_fornecedores F ON CP.id_Forn = F.IDCliente INNER JOIN compras_programa_item CP ON CP.ID = CP.ID INNER JOIN projproduto P ON P.desenho = CP.Codigo " & INNERJOINTEXTOREF & " " & INNERJOINTEXTODATA & " where CP.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and " & TextoFiltro & " = '" & txtTexto & "' " & DataFiltro & " " & TextoFiltroVenda & " group by " & CamposFiltro & " order by CP.Programa desc, CP.Rev desc"
        End If
    Else
        .Sql_Programacao_Compras_Localizar = "Select " & CamposFiltro & " FROM compras_programa CP INNER JOIN Compras_fornecedores F ON CP.id_Forn = F.IDCliente " & INNERJOINTEXTODATA & " where CP.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & DataFiltro & " " & TextoFiltroVenda & " group by " & CamposFiltro & " order by CP.Programa desc, CP.Rev desc"
    End If
    .ProcCarregaLista (1)
End With
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_data_programacao_Click()
On Error GoTo tratar_erro

If Chk_data_programacao.Value = 1 Then
    Chk_data_programa.Value = 0
    Frame1.Enabled = True
    txtData_inicio.SetFocus
Else
    Frame1.Enabled = False
    txtData_inicio.Value = Date
    txtData_fim.Value = Date
End If

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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyF2: ProcFiltrar
    'Case vbKeyF1: ProcAjuda
    Case vbKeyEscape: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 10305, 5, True
ProcCarregaComboEmpresa Cmb_empresa, False
cmbfiltrarpor = "Programa"
ProcCarregaComboFamilia cmbfamilia, "familia <> 'Null' and Compras = 'True'", True
txtData_inicio.Value = Date
txtData_fim.Value = Date

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
