VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_trocaduplicata_novo 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Administrativo - Financeiro - Desconto de duplicata - Novo"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10890
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   10890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   9150
      Top             =   210
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frm_trocaduplicata_novo.frx":0000
      Count           =   1
   End
   Begin VB.CheckBox chkVencimento 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Vencimento"
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
      TabIndex        =   5
      Top             =   2820
      Width           =   1485
   End
   Begin VB.CheckBox chkEmissao 
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
      TabIndex        =   4
      Top             =   2820
      Width           =   1005
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Height          =   1515
      Left            =   55
      TabIndex        =   8
      Top             =   990
      Width           =   10800
      Begin VB.Frame Frame11 
         BackColor       =   &H00E0E0E0&
         Height          =   510
         Left            =   5820
         TabIndex        =   16
         Top             =   210
         WhatsThisHelpID =   210
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
            TabIndex        =   20
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
            TabIndex        =   19
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
            TabIndex        =   18
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
            TabIndex        =   17
            Top             =   180
            Width           =   705
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
         ItemData        =   "frm_trocaduplicata_novo.frx":2D89
         Left            =   180
         List            =   "frm_trocaduplicata_novo.frx":2D9F
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Opções para filtro."
         Top             =   390
         Width           =   5535
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
         Height          =   330
         Left            =   180
         TabIndex        =   1
         ToolTipText     =   "Texto para pesquisa."
         Top             =   1050
         Width           =   10425
      End
      Begin VB.ComboBox cmbTexto 
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
         ItemData        =   "frm_trocaduplicata_novo.frx":2DF6
         Left            =   180
         List            =   "frm_trocaduplicata_novo.frx":2DF8
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "Familia."
         Top             =   1050
         Width           =   10425
      End
      Begin VB.Label Label45 
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
         Left            =   2527
         TabIndex        =   10
         Top             =   180
         Width           =   840
      End
      Begin VB.Label Label3 
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
         Left            =   4657
         TabIndex        =   9
         Top             =   840
         Width           =   1470
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   675
      Left            =   55
      TabIndex        =   11
      Top             =   2520
      Width           =   10800
      Begin MSComCtl2.DTPicker msk_fltFim 
         Height          =   315
         Left            =   9300
         TabIndex        =   7
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
         Format          =   183304193
         CurrentDate     =   39057
      End
      Begin MSComCtl2.DTPicker msk_fltInicio 
         Height          =   315
         Left            =   7410
         TabIndex        =   6
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
         Format          =   183304193
         CurrentDate     =   39057
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
         Height          =   195
         Left            =   7050
         TabIndex        =   13
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
         Left            =   8895
         TabIndex        =   12
         Top             =   240
         Width           =   360
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   14
      Top             =   0
      Width           =   10800
      _ExtentX        =   19050
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
      ButtonCaption2  =   "Adicionar"
      ButtonEnabled2  =   0   'False
      ButtonIconSize2 =   32
      ButtonToolTipText2=   "Adicionar selecionados (F3)"
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
      ButtonWidth2    =   52
      ButtonHeight2   =   21
      ButtonUseMaskColor2=   0   'False
      ButtonEnabled3  =   0   'False
      ButtonIconSize3 =   32
      ButtonAlignment3=   2
      ButtonType3     =   1
      ButtonStyle3    =   -1
      BeginProperty ButtonFont3 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState3    =   -1
      ButtonLeft3     =   94
      ButtonTop3      =   4
      ButtonWidth3    =   2
      ButtonHeight3   =   54
      ButtonUseMaskColor3=   0   'False
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
      ButtonLeft4     =   98
      ButtonTop4      =   2
      ButtonWidth4    =   36
      ButtonHeight4   =   21
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
      ButtonLeft5     =   136
      ButtonTop5      =   2
      ButtonWidth5    =   26
      ButtonHeight5   =   21
      ButtonEnabled6  =   0   'False
      ButtonIconSize6 =   32
      ButtonKey6      =   "6"
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
      ButtonLeft6     =   164
      ButtonTop6      =   2
      ButtonWidth6    =   24
      ButtonHeight6   =   24
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   60
      TabIndex        =   15
      Top             =   6240
      Width           =   10800
      _ExtentX        =   19050
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
   Begin MSComctlLib.ListView Lista 
      Height          =   3015
      Left            =   60
      TabIndex        =   3
      Top             =   3210
      Width           =   10800
      _ExtentX        =   19050
      _ExtentY        =   5318
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Object.Tag             =   "D"
         Text            =   "Dt. emissão"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Object.Tag             =   "D"
         Text            =   "Dt. venc."
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Object.Tag             =   "N"
         Text            =   "Valor"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Nº docto."
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Object.Tag             =   "N"
         Text            =   "Nota fiscal"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Object.Tag             =   "T"
         Text            =   "Parcela"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Object.Tag             =   "T"
         Text            =   "Cliente"
         Object.Width           =   5345
      EndProperty
   End
End
Attribute VB_Name = "frm_trocaduplicata_novo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StrSql_Contas_Receber_Desconto_Duplicata_Novo As String 'OK

Private Sub chkEmissao_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
If chkEmissao.Value = 1 Then
    chkVencimento.Value = 0
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

Private Sub chkVencimento_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
If chkVencimento.Value = 1 Then
    chkEmissao.Value = 0
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

Private Sub cmbfiltrarpor_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
txtTexto.Visible = True
If cmbfiltrarpor = "Conta contábil" Or cmbfiltrarpor = "Instituição" Or cmbfiltrarpor = "Cliente" Then
    txtTexto.Visible = False
    cmbTexto.Visible = True
    
    Texto = ""
    cmbTexto.Clear
    Set TBLISTA = CreateObject("adodb.recordset")
    Select Case cmbfiltrarpor
        Case "Conta contábil":
            Set TBLISTA = CreateObject("adodb.recordset")
            TBLISTA.Open "Select tbl_familia.int_codfamilia, tbl_familia.Codigo, tbl_familia.txt_descricao from (tbl_contas_receber INNER JOIN familia_financeiro ON tbl_contas_receber.IdIntConta = familia_financeiro.IDConta) INNER JOIN tbl_familia ON tbl_familia.int_codfamilia = familia_financeiro.ID_PC where familia_financeiro.tipoconta = 'R' and tbl_contas_receber.Idtrocatitulo = 0 and tbl_contas_receber.ID_empresa = " & frm_trocaduplicata.Cmb_empresa.ItemData(frm_trocaduplicata.Cmb_empresa.ListIndex) & " and tbl_contas_receber.Logsit = 'N' Group by tbl_familia.int_codfamilia, tbl_familia.Codigo, tbl_familia.txt_descricao", Conexao, adOpenKeyset, adLockOptimistic
            If TBLISTA.EOF = False Then
                Do While TBLISTA.EOF = False
                    cmbTexto.AddItem TBLISTA!Txt_descricao & " - " & TBLISTA!CODIGO
                    cmbTexto.ItemData(cmbTexto.NewIndex) = TBLISTA!int_codfamilia
                    TBLISTA.MoveNext
                Loop
            End If
            TBLISTA.Close
        Case "Cliente"
            Set TBLISTA = CreateObject("adodb.recordset")
            TBLISTA.Open "Select IDcliente, Nome_Razao from tbl_contas_receber where Nome_Razao is not null and Idtrocatitulo = 0 and ID_empresa = " & frm_trocaduplicata.Cmb_empresa.ItemData(frm_trocaduplicata.Cmb_empresa.ListIndex) & " and Logsit = 'N' Group by IDcliente, Nome_Razao", Conexao, adOpenKeyset, adLockOptimistic
            If TBLISTA.EOF = False Then
                Do While TBLISTA.EOF = False
                    cmbTexto.AddItem TBLISTA!Nome_Razao
                    cmbTexto.ItemData(cmbTexto.NewIndex) = TBLISTA!IDCliente
                    TBLISTA.MoveNext
                Loop
            End If
            TBLISTA.Close
        Case "Instituição":
            Set TBLISTA = CreateObject("adodb.recordset")
            TBLISTA.Open "Select txt_Descricao from tbl_Instituicoes where ID_empresa = " & frm_trocaduplicata.Cmb_empresa.ItemData(frm_trocaduplicata.Cmb_empresa.ListIndex) & " order by txt_Descricao", Conexao, adOpenKeyset, adLockOptimistic
            If TBLISTA.EOF = False Then
                Do While TBLISTA.EOF = False
                    If IsNull(TBLISTA!Txt_descricao) = False And TBLISTA!Txt_descricao <> "" Then
                        If Texto <> TBLISTA!Txt_descricao Then cmbTexto.AddItem Trim(TBLISTA!Txt_descricao)
                        Texto = TBLISTA!Txt_descricao
                    End If
                    TBLISTA.MoveNext
                Loop
            End If
            TBLISTA.Close
    End Select
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbTexto_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Lista
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                'Verifica se a conta está vencida
                Dataini = .ListItems.Item(InitFor).ListSubItems(2)
                DataFim = frm_trocaduplicata.txtData
                If Dataini < DataFim Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView Lista, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With Lista
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True And .ListItems.Item(InitFor) = Item Then
            'Verifica se a conta está vencida
            Dataini = .ListItems.Item(InitFor).ListSubItems(2)
            DataFim = frm_trocaduplicata.txtData
            If Dataini < DataFim Then
                USMsgBox ("Não é permitido descontar esta duplicata, pois o mesma já está vencida."), vbExclamation, "CAPRIND v5.0"
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            
            'Verifica limite de desconto no banco
            valor = Lista.SelectedItem.ListSubItems(3)
            Set TBReceber = CreateObject("adodb.recordset")
            TBReceber.Open "Select * from tbl_Instituicoes where txt_Descricao = '" & frm_trocaduplicata.txtlocaltroca & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBReceber.EOF = False Then
                If valor + TBReceber!Limite_utilizado > TBReceber!Limite_desconto Then
                    USMsgBox ("Não é permitido descontar essa duplicata pois o limite para desconto hoje é de " & Format(TBReceber!Limite_desconto - TBReceber!Limite_utilizado, "###,##0.00") & "."), vbInformation, "CAPRIND v5.0"
                    If USMsgBox("Deseja prosseguir assim mesmo?", vbYesNo, "CAPRIND v5.0") = vbNo Then
                        TBReceber.Close
                        Exit Sub
                    End If
                End If
            End If
            TBReceber.Close
        End If
    Next InitFor
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub msk_fltFim_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 10800, 6, True

cmbfiltrarpor = "Cliente"
msk_fltInicio.Value = Date
msk_fltFim.Value = Date

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro
  
Select Case KeyCode
    Case vbKeyF2: ProcFiltrar
    Case vbKeyF3: ProcAdicionar
    'Case vbKeyF1: ProcAjuda
    Case vbKeyEscape: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
 
Private Sub ProcFiltrar()
On Error GoTo tratar_erro

With msk_fltFim
    If FunVerificaDataFinal(msk_fltInicio.Value, .Value) = False Then
        .Value = Date
        .SetFocus
        Exit Sub
    End If
End With
If chkEmissao.Value = 1 Then Data_receber = "CR.Emissao" Else Data_receber = "CR.Vencimento"
If chkVencimento.Value = 1 Or chkEmissao.Value = 1 Then DataFiltro = Data_receber & " Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "'" Else DataFiltro = "CR.Nome_Razao <> 'Null'"
With frm_trocaduplicata
    If txtTexto.Visible = True And txtTexto <> "" Or cmbTexto.Visible = True And cmbTexto <> "" Then
        If cmbfiltrarpor = "Conta contábil" Then
            StrSql_Contas_Receber_Desconto_Duplicata_Novo = "Select CR.* from tbl_Contas_receber CR INNER JOIN familia_financeiro FF on CR.IdIntConta = FF.idconta WHERE FF.ID_PC = " & cmbTexto.ItemData(cmbTexto.ListIndex) & " and FF.tipoconta = 'R' AND CR.logsit = 'N' and CR.bloqueado = 'False' and " & DataFiltro & " and CR.ID_empresa = " & .Cmb_empresa.ItemData(.Cmb_empresa.ListIndex) & " and CR.Idtrocatitulo = 0 order by " & Data_receber & ", CR.IdIntConta"
        ElseIf cmbfiltrarpor = "Instituição" Then
                StrSql_Contas_Receber_Desconto_Duplicata_Novo = "Select * FROM tbl_Contas_RECEBER CR WHERE Banco = '" & cmbTexto & "' and logsit='N' and bloqueado = 'False' and " & DataFiltro & " and ID_empresa = " & .Cmb_empresa.ItemData(.Cmb_empresa.ListIndex) & " and Idtrocatitulo = 0 order by " & Data_receber & ", IdIntConta"
            ElseIf cmbfiltrarpor = "Cliente" Then
                    StrSql_Contas_Receber_Desconto_Duplicata_Novo = "Select * FROM tbl_Contas_receber CR WHERE IDcliente = " & cmbTexto.ItemData(cmbTexto.ListIndex) & " and logsit='N' and bloqueado = 'False' and " & DataFiltro & " and ID_empresa = " & .Cmb_empresa.ItemData(.Cmb_empresa.ListIndex) & " and Idtrocatitulo = 0 order by " & Data_receber & ", IDIntconta"
                ElseIf cmbfiltrarpor = "Pedido interno" Then
                        StrSql_Contas_Receber_Desconto_Duplicata_Novo = "Select CR.* FROM tbl_Contas_RECEBER CR INNER JOIN tbl_proposta_nota P on CR.NFiscal = P.NF WHERE P.proposta" & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and CR.logsit='N' and CR.bloqueado = 'False' and " & DataFiltro & " and CR.ID_empresa = " & .Cmb_empresa.ItemData(.Cmb_empresa.ListIndex) & " and Idtrocatitulo = 0 order by " & Data_receber & ", CR.IdIntConta"
                    ElseIf cmbfiltrarpor = "Pedido cliente" Then
                            StrSql_Contas_Receber_Desconto_Duplicata_Novo = "Select CR.* FROM (tbl_Contas_RECEBER CR INNER JOIN vendas_proposta VP ON CR.proposta = VP.Ncotacao) INNER JOIN vendas_carteira VC ON VC.Cotacao = VP.Cotacao WHERE VC.PCCliente" & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and CR.logsit='N' and CR.bloqueado = 'False' and " & DataFiltro & " and CR.ID_empresa = " & .Cmb_empresa.ItemData(.Cmb_empresa.ListIndex) & " and CR.Idtrocatitulo = 0 order by " & Data_receber & ", CR.IdIntConta"
                        Else
                            Select Case cmbfiltrarpor
                                Case "Nota fiscal":
                                    TextoFiltro = "Nfiscal"
                                    If txtTexto <> "" Then txtTexto = FunTamanhoTextoZeroEsq(txtTexto, 9)
                                Case "Cliente": TextoFiltro = "Nome_Razao"
                            End Select
                            StrSql_Contas_Receber_Desconto_Duplicata_Novo = "Select * from tbl_Contas_RECEBER CR where " & TextoFiltro & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and logsit='N' and bloqueado = 'False' and " & DataFiltro & " and ID_empresa = " & .Cmb_empresa.ItemData(.Cmb_empresa.ListIndex) & " and Idtrocatitulo = 0 order by " & Data_receber & ", IdIntConta"
                            
        End If
    Else
        StrSql_Contas_Receber_Desconto_Duplicata_Novo = "Select * FROM tbl_Contas_receber CR WHERE logsit = 'N' and bloqueado = 'False' and " & DataFiltro & " and ID_empresa = " & .Cmb_empresa.ItemData(.Cmb_empresa.ListIndex) & " and Idtrocatitulo = 0 order by " & Data_receber & ", IDIntconta"
    End If
    ProcCarregaLista
End With

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

Sub ProcCarregaLista()
On Error GoTo tratar_erro

Dataini = 0
Lista.ListItems.Clear
If StrSql_Contas_Receber_Desconto_Duplicata_Novo <> "" Then
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open StrSql_Contas_Receber_Desconto_Duplicata_Novo, Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
        TBProduto.MoveLast
        PBLista.Min = 0
        PBLista.Max = TBProduto.RecordCount
        PBLista.Value = 1
        Contador = 0
        TBProduto.MoveFirst
        Do While TBProduto.EOF = False
            With Lista.ListItems.Add(, , TBProduto!IDintconta)
            Set TBContas = CreateObject("adodb.recordset")
                TBContas.Open "Select * from troca_titulo_valores where N_conta = " & TBProduto!IDintconta, Conexao, adOpenKeyset, adLockOptimistic
                If TBContas.EOF = True Then
                    .SubItems(1) = IIf(IsNull(TBProduto!emissao), "", Format(TBProduto!emissao, "dd/mm/yy"))
                    .SubItems(2) = IIf(IsNull(TBProduto!Vencimento), "", Format(TBProduto!Vencimento, "dd/mm/yy"))
                    .SubItems(3) = IIf(IsNull(TBProduto!valor), "", Format(Trim(TBProduto!valor), "###,##0.00"))
                    .SubItems(4) = IIf(IsNull(TBProduto!txt_ndocumento), "", TBProduto!txt_ndocumento)
                    .SubItems(5) = IIf(IsNull(TBProduto!NFiscal), "", TBProduto!NFiscal)
                    .SubItems(6) = IIf(IsNull(TBProduto!Parcela), "", TBProduto!Parcela)
                    .SubItems(7) = IIf(IsNull(TBProduto!Nome_Razao), "", TBProduto!Nome_Razao)
                    
                    Dataini = Format(TBProduto!Vencimento, "dd/mm/yy")
                    If Date > Dataini Then
                        .ForeColor = vbRed
                        .ListSubItems(1).ForeColor = vbRed
                        .ListSubItems(2).ForeColor = vbRed
                        .ListSubItems(3).ForeColor = vbRed
                        .ListSubItems(4).ForeColor = vbRed
                        .ListSubItems(5).ForeColor = vbRed
                        .ListSubItems(6).ForeColor = vbRed
                        .ListSubItems(7).ForeColor = vbRed
                    End If
                End If
                TBContas.Close
            End With
            ValorTotal = ValorTotal + TBProduto!valor
            TBProduto.MoveNext
            Contador = Contador + 1
            PBLista.Value = Contador
        Loop
        
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcEnviaDados()
On Error GoTo tratar_erro

'ValorTotal = 0
With frm_trocaduplicata
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from tbl_contas_receber where IDIntconta = " & Lista.ListItems.Item(InitFor), Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        TBAbrir!IDtrocatitulo = .txtBordero
        TBAbrir!titulodesc = True
        TBAbrir!status = "DUPLICATA DESCONTADA EM ABERTO"
        TBAbrir.Update
        
        Set TBContas = CreateObject("adodb.recordset")
        TBContas.Open "Select * from troca_titulo_valores where N_conta = " & Lista.ListItems.Item(InitFor), Conexao, adOpenKeyset, adLockOptimistic
        If TBContas.EOF = True Then TBContas.AddNew
        'ValorTotal = IIf(IsNull(TBAbrir!Valor), 0, TBAbrir!Valor)
        '.txtvalortitulo = ValorTotal
        
        'Verifica prazo medio
        If IsNull(TBAbrir!Vencimento) = False And TBAbrir!Vencimento >= .txtData Then
            Dataini = TBAbrir!Vencimento
            DataFim = frm_trocaduplicata.txtData
            Data = Dataini - DataFim
        Else
            Data = 0
        End If
        
        ElapsedTime (Data)
        TBContas!Prazo = D
        
        TBContas!valor_pis = IIf(.txtPIS = "", 0, .txtPIS)
        .ProcCalculaPIS
        TBContas!valor_cofins = IIf(.txtCofins = "", 0, .txtCofins)
        .ProcCalculaCofins
        TBContas!valor_retido = .txtretido
        TBContas!valor_enviado = .txtenviado
        TBContas!IDduplicata = .txtBordero
        TBContas!n_conta = Lista.ListItems.Item(InitFor)
        TBContas.Update
        TBContas.Close
        
        '==================================
        Modulo = "Financeiro/Contas à receber/Desconto de duplicata"
        Evento = "Nova duplicata"
        ID_documento = Lista.ListItems.Item(InitFor)
        Documento = "Borderô: " & .txtBordero
        Documento1 = "Documento: " & Lista.ListItems(InitFor).SubItems(4)
        ProcGravaEvento
        '==================================
    End If
    TBAbrir.Close
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub msk_fltInicio_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optfim_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optinicio_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optmeio_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtTexto_Change()
On Error GoTo tratar_erro

Lista.ListItems.Clear
If cmbfiltrarpor = "Nota fiscal" And txtTexto <> "" Then
    VerifNumero = txtTexto
    ProcVerificaNumero
    If VerifNumero = False Then
        txtTexto = ""
        txtTexto.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtTexto_LostFocus()
On Error GoTo tratar_erro

If cmbfiltrarpor = "Nota fiscal" And txtTexto <> "" Then txtTexto = FunTamanhoTextoZeroEsq(txtTexto, 9)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcFiltrar
    Case 2: ProcAdicionar
    'Case 4: ProcAjuda
    Case 5: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcAdicionar()
On Error GoTo tratar_erro

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With Lista
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente adicionar esta(s) duplicata(s)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            
            frm_trocaduplicata.ProcLimpaValores
            ProcEnviaDados
        End If
    Next InitFor
End With

If Permitido = False Then
    USMsgBox ("Informe a(s) duplicatas(s) antes de adicionar."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Duplicata(s) adicionada(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    With frm_trocaduplicata
        .ProcCalculaNDuplicatasPMedio
        .ProcCarregaLista
        .ProcGravarTotais
        .ProcAtualizaLimiteUtil
        
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from troca_titulo where ID = " & .txtBordero, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            .ProcCarregaTotais
        End If
        TBAbrir.Close
        .ProcLimpaValores
    End With
End If
ProcCarregaLista

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
