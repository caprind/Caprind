VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form frmLiquido_pedido 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Qualidade - Ensaios - Líquido penetrante - Localizar pedido"
   ClientHeight    =   6390
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   8925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   11085
      MouseIcon       =   "frmLiquido_pedido.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "frmLiquido_pedido.frx":0152
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Cancela e fecha formulário (Esc)"
      Top             =   165
      Width           =   570
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3705
      Left            =   60
      TabIndex        =   5
      Top             =   2400
      Width           =   8805
      _ExtentX        =   15531
      _ExtentY        =   6535
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "T"
         Text            =   "Pedido"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Object.Tag             =   "N"
         Text            =   "Rev."
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Object.Tag             =   "D"
         Text            =   "Dt. emissão"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Status"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Object.Tag             =   "N"
         Text            =   "Valor total"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Object.Tag             =   "T"
         Text            =   "Cliente"
         Object.Width           =   4789
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Object.Tag             =   "T"
         Text            =   "Ped. cliente"
         Object.Width           =   2117
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   1515
      Left            =   55
      TabIndex        =   8
      Top             =   870
      Width           =   8805
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0E0&
         Height          =   510
         Left            =   4620
         TabIndex        =   10
         Top             =   210
         Width           =   3975
         Begin VB.OptionButton Optmeio 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Meio frase"
            Height          =   255
            Left            =   1470
            TabIndex        =   2
            Top             =   180
            Width           =   1275
         End
         Begin VB.OptionButton Optinicio 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Início frase"
            Height          =   255
            Left            =   180
            TabIndex        =   1
            Top             =   180
            Value           =   -1  'True
            Width           =   1275
         End
         Begin VB.OptionButton Optfim 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Fim frase"
            Height          =   255
            Left            =   2760
            TabIndex        =   3
            Top             =   180
            Width           =   1155
         End
      End
      Begin VB.ComboBox cmbfiltrarpor 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         ItemData        =   "frmLiquido_pedido.frx":1194
         Left            =   180
         List            =   "frmLiquido_pedido.frx":11B3
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Opções para filtro."
         Top             =   390
         Width           =   4365
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
         TabIndex        =   4
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
         TabIndex        =   6
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
         Left            =   3652
         TabIndex        =   11
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
         Left            =   1942
         TabIndex        =   9
         Top             =   180
         Width           =   840
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   0
      TabIndex        =   12
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft1     =   2
      ButtonTop1      =   2
      ButtonWidth1    =   42
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
      ButtonLeft2     =   46
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft3     =   50
      ButtonTop3      =   2
      ButtonWidth3    =   41
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft4     =   93
      ButtonTop4      =   2
      ButtonWidth4    =   30
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
      ButtonLeft5     =   125
      ButtonTop5      =   2
      ButtonWidth5    =   24
      ButtonHeight5   =   24
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   6150
         Top             =   150
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmLiquido_pedido.frx":1236
         Count           =   1
      End
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   60
      TabIndex        =   13
      Top             =   6120
      Width           =   8835
      _ExtentX        =   15584
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
      SearchText      =   "Atualizando..."
      Value           =   0
   End
End
Attribute VB_Name = "frmLiquido_pedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StrSql_Liquido_Pedido As String 'OK

Private Sub cmbFamilia_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
If cmbfamilia <> "" Then txtTexto = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbfiltrarpor_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
If cmbfiltrarpor = "Família" Then
    txtTexto.Visible = False
    cmbfamilia.Visible = True
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLocalizar()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
If txtTexto.Visible = True And txtTexto <> "" Or cmbfamilia.Visible = True And cmbfamilia <> "" Then
    If cmbfiltrarpor = "Pedido" Then
        If Optinicio.Value = True Then StrSql_Liquido_Pedido = "Select * from vendas_Proposta where Ncotacao like '" & txtTexto & "%' and Tipo = 'PE' or Ncotacao like '" & txtTexto & "%' and tipo = 'PRPE' order by ordenarproposta, cotacao desc"
        If Optmeio.Value = True Then StrSql_Liquido_Pedido = "Select * from vendas_Proposta where Ncotacao like '%" & txtTexto & "%' and Tipo = 'PE' or Ncotacao like '%" & txtTexto & "%' and tipo = 'PRPE' order by ordenarproposta, cotacao desc"
        If Optfim.Value = True Then StrSql_Liquido_Pedido = "Select * from vendas_Proposta where Ncotacao like '%" & txtTexto & "' and Tipo = 'PE' or Ncotacao like '%" & txtTexto & "' and tipo = 'PRPE' order by ordenarproposta, cotacao desc"
    End If
    If cmbfiltrarpor = "S/ referencia" Then
        If Optinicio.Value = True Then StrSql_Liquido_Pedido = "Select * from vendas_Proposta where Referente like '" & txtTexto & "%' and Tipo = 'PE' or Referente like '" & txtTexto & "%' and tipo = 'PRPE' order by ordenarproposta, cotacao desc"
        If Optmeio.Value = True Then StrSql_Liquido_Pedido = "Select * from vendas_Proposta where Referente like '%" & txtTexto & "%' and Tipo = 'PE' or Referente like '%" & txtTexto & "%' and tipo = 'PRPE' order by ordenarproposta, cotacao desc"
        If Optfim.Value = True Then StrSql_Liquido_Pedido = "Select * from vendas_Proposta where Referente like '%" & txtTexto & "' and Tipo = 'PE' or Referente like '%" & txtTexto & "' and tipo = 'PRPE' order by ordenarproposta, cotacao desc"
    End If
    If cmbfiltrarpor = "Pedido cliente" Then
        If Optinicio.Value = True Then StrSql_Liquido_Pedido = "Select vendas_Proposta.* FROM vendas_Proposta INNER JOIN vendas_carteira ON vendas_Proposta.cotacao = vendas_carteira.cotacao where vendas_carteira.PCCliente like '" & txtTexto & "%' and vendas_Proposta.Tipo = 'PE' or vendas_carteira.PCCliente like '" & txtTexto & "%' and vendas_Proposta.Tipo = 'PRPE' order by vendas_Proposta.ordenarproposta, vendas_Proposta.cotacao desc"
        If Optmeio.Value = True Then StrSql_Liquido_Pedido = "Select vendas_Proposta.* FROM vendas_Proposta INNER JOIN vendas_carteira ON vendas_Proposta.cotacao = vendas_carteira.cotacao where vendas_carteira.PCCliente like '%" & txtTexto & "%' and vendas_Proposta.Tipo = 'PE' or vendas_carteira.PCCliente like '%" & txtTexto & "%' and vendas_Proposta.Tipo = 'PRPE' order by vendas_Proposta.ordenarproposta, vendas_Proposta.cotacao desc"
        If Optfim.Value = True Then StrSql_Liquido_Pedido = "Select vendas_Proposta.* FROM vendas_Proposta INNER JOIN vendas_carteira ON vendas_Proposta.cotacao = vendas_carteira.cotacao where vendas_carteira.PCCliente like '%" & txtTexto & "' and vendas_Proposta.Tipo = 'PE' or vendas_carteira.PCCliente like '%" & txtTexto & "' and vendas_Proposta.Tipo = 'PRPE' order by vendas_Proposta.ordenarproposta, vendas_Proposta.cotacao desc"
    End If
    If cmbfiltrarpor = "Cliente" Then
        If Optinicio.Value = True Then StrSql_Liquido_Pedido = "Select * from vendas_Proposta where Cliente like '" & txtTexto & "%' and Tipo = 'PE' or Cliente like '" & txtTexto & "%' and tipo = 'PRPE' order by ordenarproposta, cotacao desc"
        If Optmeio.Value = True Then StrSql_Liquido_Pedido = "Select * from vendas_Proposta where Cliente like '%" & txtTexto & "%' and Tipo = 'PE' or Cliente like '%" & txtTexto & "%' and tipo = 'PRPE' order by ordenarproposta, cotacao desc"
        If Optfim.Value = True Then StrSql_Liquido_Pedido = "Select * from vendas_Proposta where Cliente like '%" & txtTexto & "' and Tipo = 'PE' or Cliente like '%" & txtTexto & "' and tipo = 'PRPE' order by ordenarproposta, cotacao desc"
    End If
    If cmbfiltrarpor = "Código interno" Then
        If Optinicio.Value = True Then StrSql_Liquido_Pedido = "Select vendas_Proposta.* FROM vendas_Proposta INNER JOIN vendas_carteira ON vendas_Proposta.cotacao = vendas_carteira.cotacao where vendas_carteira.Desenho like '" & txtTexto & "%' and vendas_Proposta.Tipo = 'PE' or vendas_carteira.Desenho like '" & txtTexto & "%' and vendas_Proposta.Tipo = 'PRPE' order by vendas_Proposta.ordenarproposta, vendas_Proposta.cotacao desc"
        If Optmeio.Value = True Then StrSql_Liquido_Pedido = "Select vendas_Proposta.* FROM vendas_Proposta INNER JOIN vendas_carteira ON vendas_Proposta.cotacao = vendas_carteira.cotacao where vendas_carteira.Desenho like '%" & txtTexto & "%' and vendas_Proposta.Tipo = 'PE' or vendas_carteira.Desenho like '%" & txtTexto & "%' and vendas_Proposta.Tipo = 'PRPE' order by vendas_Proposta.ordenarproposta, vendas_Proposta.cotacao desc"
        If Optfim.Value = True Then StrSql_Liquido_Pedido = "Select vendas_Proposta.* FROM vendas_Proposta INNER JOIN vendas_carteira ON vendas_Proposta.cotacao = vendas_carteira.cotacao where vendas_carteira.Desenho like '%" & txtTexto & "' and vendas_Proposta.Tipo = 'PE' or vendas_carteira.Desenho like '%" & txtTexto & "' and vendas_Proposta.Tipo = 'PRPE' order by vendas_Proposta.ordenarproposta, vendas_Proposta.cotacao desc"
    End If
    If cmbfiltrarpor = "Código de referência" Then
        If Optinicio.Value = True Then StrSql_Liquido_Pedido = "Select vendas_Proposta.* FROM vendas_Proposta INNER JOIN vendas_carteira ON vendas_Proposta.cotacao = vendas_carteira.cotacao where vendas_carteira.n_referencia like '" & txtTexto & "%' and vendas_Proposta.Tipo = 'PE' or vendas_carteira.n_referencia like '" & txtTexto & "%' and vendas_Proposta.Tipo = 'PRPE' order by vendas_Proposta.ordenarproposta, vendas_Proposta.cotacao desc"
        If Optmeio.Value = True Then StrSql_Liquido_Pedido = "Select vendas_Proposta.* FROM vendas_Proposta INNER JOIN vendas_carteira ON vendas_Proposta.cotacao = vendas_carteira.cotacao where vendas_carteira.n_referencia like '%" & txtTexto & "%' and vendas_Proposta.Tipo = 'PE' or vendas_carteira.n_referencia like '%" & txtTexto & "%' and vendas_Proposta.Tipo = 'PRPE' order by vendas_Proposta.ordenarproposta, vendas_Proposta.cotacao desc"
        If Optfim.Value = True Then StrSql_Liquido_Pedido = "Select vendas_Proposta.* FROM vendas_Proposta INNER JOIN vendas_carteira ON vendas_Proposta.cotacao = vendas_carteira.cotacao where vendas_carteira.n_referencia like '%" & txtTexto & "' and vendas_Proposta.Tipo = 'PE' or vendas_carteira.n_referencia like '%" & txtTexto & "' and vendas_Proposta.Tipo = 'PRPE' order by vendas_Proposta.ordenarproposta, vendas_Proposta.cotacao desc"
    End If
    If cmbfiltrarpor = "Descrição" Then
        If Optinicio.Value = True Then StrSql_Liquido_Pedido = "Select vendas_Proposta.* FROM vendas_Proposta INNER JOIN vendas_carteira ON vendas_Proposta.cotacao = vendas_carteira.cotacao where vendas_carteira.Descricao_tecnica like '" & txtTexto & "%' and vendas_Proposta.Tipo = 'PE' or vendas_carteira.Descricao_tecnica like '" & txtTexto & "%' and vendas_Proposta.Tipo = 'PRPE' order by vendas_Proposta.ordenarproposta, vendas_Proposta.cotacao desc"
        If Optmeio.Value = True Then StrSql_Liquido_Pedido = "Select vendas_Proposta.* FROM vendas_Proposta INNER JOIN vendas_carteira ON vendas_Proposta.cotacao = vendas_carteira.cotacao where vendas_carteira.Descricao_tecnica like '%" & txtTexto & "%' and vendas_Proposta.Tipo = 'PE' or vendas_carteira.Descricao_tecnica like '%" & txtTexto & "%' and vendas_Proposta.Tipo = 'PRPE' order by vendas_Proposta.ordenarproposta, vendas_Proposta.cotacao desc"
        If Optfim.Value = True Then StrSql_Liquido_Pedido = "Select vendas_Proposta.* FROM vendas_Proposta INNER JOIN vendas_carteira ON vendas_Proposta.cotacao = vendas_carteira.cotacao where vendas_carteira.Descricao_tecnica like '%" & txtTexto & "' and vendas_Proposta.Tipo = 'PE' or vendas_carteira.Descricao_tecnica like '%" & txtTexto & "' and vendas_Proposta.Tipo = 'PRPE' order by vendas_Proposta.ordenarproposta, vendas_Proposta.cotacao desc"
    End If
    If cmbfiltrarpor = "Descrição comercial" Then
        If Optinicio.Value = True Then StrSql_Liquido_Pedido = "Select vendas_Proposta.* FROM vendas_Proposta INNER JOIN vendas_carteira ON vendas_Proposta.cotacao = vendas_carteira.cotacao where vendas_carteira.Descricao like '" & txtTexto & "%' and vendas_Proposta.Tipo = 'PE' or vendas_carteira.Descricao like '" & txtTexto & "%' and vendas_Proposta.Tipo = 'PRPE' order by vendas_Proposta.ordenarproposta, vendas_Proposta.cotacao desc"
        If Optmeio.Value = True Then StrSql_Liquido_Pedido = "Select vendas_Proposta.* FROM vendas_Proposta INNER JOIN vendas_carteira ON vendas_Proposta.cotacao = vendas_carteira.cotacao where vendas_carteira.Descricao like '%" & txtTexto & "%' and vendas_Proposta.Tipo = 'PE' or vendas_carteira.Descricao like '%" & txtTexto & "%' and vendas_Proposta.Tipo = 'PRPE' order by vendas_Proposta.ordenarproposta, vendas_Proposta.cotacao desc"
        If Optfim.Value = True Then StrSql_Liquido_Pedido = "Select vendas_Proposta.* FROM vendas_Proposta INNER JOIN vendas_carteira ON vendas_Proposta.cotacao = vendas_carteira.cotacao where vendas_carteira.Descricao like '%" & txtTexto & "' and vendas_Proposta.Tipo = 'PE' or vendas_carteira.Descricao like '%" & txtTexto & "' and vendas_Proposta.Tipo = 'PRPE' order by vendas_Proposta.ordenarproposta, vendas_Proposta.cotacao desc"
    End If
    If cmbfiltrarpor = "Família" Then StrSql_Liquido_Pedido = "Select vendas_Proposta.* FROM vendas_Proposta INNER JOIN vendas_carteira ON vendas_Proposta.cotacao = vendas_carteira.cotacao where vendas_carteira.Familia = '" & cmbfamilia & "' and vendas_Proposta.Tipo = 'PE' or vendas_carteira.Familia = '" & cmbfamilia & "' and vendas_Proposta.Tipo = 'PRPE' order by vendas_Proposta.ordenarproposta, vendas_Proposta.cotacao desc"
Else
    StrSql_Liquido_Pedido = "Select * from vendas_Proposta where Tipo = 'PE' or tipo = 'PRPE' order by ordenarproposta, cotacao desc"
End If
ProcCarregaLista

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyEscape: Unload Me
    Case vbKeyReturn: ListView1_DblClick
    Case vbKeyF2: ProcLocalizar
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 8805, 5, True
If Ultrasom = True Then frmLiquido_pedido.Caption = "Qualidade - Ensaios - Ultra-som - Localizar pedido"
If Liquido = True Then frmLiquido_pedido.Caption = "Qualidade - Ensaios - Líquido penetrante - Localizar pedido"
cmbfiltrarpor = "Pedido"
Optinicio.Value = True
txtTexto.Visible = True
cmbfamilia.Visible = False
ProcCarregaComboFamilia cmbfamilia, "familia <> 'Null' and vendas = 'True'", False

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

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView ListView1, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListView1_DblClick()
On Error GoTo tratar_erro

If ListView1.ListItems.Count = 0 Then Exit Sub
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from vendas_proposta where cotacao = " & ListView1.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    If Liquido = True Then
        With frmLiquido
            .txtImagem = ""
            .txtdescricao = ""
            .txtdesenho = ""
            .txtid_cliente = ""
            .txtCliente = ""
            .txtPedido_interno = ""
            .txtPedido_cliente = ""
            .txtPedido_interno = IIf(IsNull(TBAbrir!Ncotacao), "", TBAbrir!Ncotacao)
            .txtPedido_cliente = IIf(IsNull(TBAbrir!afm), "", TBAbrir!afm)
            .txtid_cliente = IIf(IsNull(TBAbrir!IDCliente), "", TBAbrir!IDCliente)
            .txtCliente = IIf(IsNull(TBAbrir!Cliente), "", TBAbrir!Cliente)
        End With
    End If
    If Ultrasom = True Then
        With frmUltraSom
            .txtImagem = ""
            .txtdescricao = ""
            .txtdesenho = ""
            .txtid_cliente = ""
            .txtCliente = ""
            .txtPedido_interno = ""
            .txtPedido_cliente = ""
            .txtPedido_interno = IIf(IsNull(TBAbrir!Ncotacao), "", TBAbrir!Ncotacao)
            .txtPedido_cliente = IIf(IsNull(TBAbrir!afm), "", TBAbrir!afm)
            .txtid_cliente = IIf(IsNull(TBAbrir!IDCliente), "", TBAbrir!IDCliente)
            .txtCliente = IIf(IsNull(TBAbrir!Cliente), "", TBAbrir!Cliente)
        End With
    End If
    If Ultrasom = False And Liquido = False Then
        With frmCertificado_qualidade
            .SSTab1.Tab = 0
            .ProcLimpacampos_ultra
            .ProcCarregalista_ultra
            .txtdescricao = ""
            .txtdesenho = ""
            .txtid_cliente = ""
            .txtCliente = ""
            .txtPedido_interno = ""
            .txtPedido_cliente = ""
            .txtPedido_interno = IIf(IsNull(TBAbrir!Ncotacao), "", TBAbrir!Ncotacao)
            .txtPedido_cliente = IIf(IsNull(TBAbrir!afm), "", TBAbrir!afm)
            .txtid_cliente = IIf(IsNull(TBAbrir!IDCliente), "", TBAbrir!IDCliente)
            .txtCliente = IIf(IsNull(TBAbrir!Cliente), "", TBAbrir!Cliente)
        End With
    End If
End If
TBAbrir.Close
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaLista()
On Error GoTo tratar_erro

Cotacao = 0
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open StrSql_Liquido_Pedido, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    TBLISTA.MoveLast
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    contador = 0
    TBLISTA.MoveFirst
    Do While TBLISTA.EOF = False
        If Cotacao <> TBLISTA!Cotacao Then
            With ListView1.ListItems
                .Add , , TBLISTA!Cotacao
                .Item(.Count).SubItems(1) = TBLISTA!Ncotacao
                .Item(.Count).SubItems(2) = TBLISTA!Revisao
                .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Data), "", Format(TBLISTA!Data, "dd/mm/yy"))
                .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!status), "", TBLISTA!status)
                .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!dbl_valor_total), "", Format(TBLISTA!dbl_valor_total, "###,##0.00"))
                .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA!Cliente), "", Trim(TBLISTA!Cliente))
                .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA!afm), "", TBLISTA!afm)
            End With
        End If
        Cotacao = TBLISTA!Cotacao
        TBLISTA.MoveNext
        contador = contador + 1
        PBLista.Value = contador
    Loop
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optfim_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optinicio_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optmeio_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtTexto_Change()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
If txtTexto <> "" Then cmbfamilia.ListIndex = -1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcLocalizar
    'Case 3: ProcAjuda
    Case 4: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

