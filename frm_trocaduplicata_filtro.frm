VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frm_trocaduplicata_filtro 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Administrativo - Financeiro - Desconto de duplicata - Localizar"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10890
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   3690
   ScaleWidth      =   10890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      ItemData        =   "frm_trocaduplicata_filtro.frx":0000
      Left            =   1170
      List            =   "frm_trocaduplicata_filtro.frx":0002
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Empresa."
      Top             =   1110
      Width           =   9495
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
      TabIndex        =   9
      Top             =   3240
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
      TabIndex        =   8
      Top             =   3240
      Width           =   1005
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Height          =   1515
      Left            =   55
      TabIndex        =   12
      Top             =   1470
      Width           =   10800
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Height          =   510
         Left            =   5820
         TabIndex        =   20
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
            TabIndex        =   7
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
            TabIndex        =   5
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
            TabIndex        =   4
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
            TabIndex        =   6
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
         ItemData        =   "frm_trocaduplicata_filtro.frx":0004
         Left            =   180
         List            =   "frm_trocaduplicata_filtro.frx":001D
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "Opções para filtro."
         Top             =   390
         Width           =   5565
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
         TabIndex        =   2
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
         ItemData        =   "frm_trocaduplicata_filtro.frx":0083
         Left            =   180
         List            =   "frm_trocaduplicata_filtro.frx":0085
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "Familia."
         Top             =   1050
         Width           =   10425
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
         TabIndex        =   14
         Top             =   840
         Width           =   1470
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
         Left            =   2542
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
      Width           =   10800
      Begin MSComCtl2.DTPicker msk_fltFim 
         Height          =   315
         Left            =   9300
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
         Format          =   489226241
         CurrentDate     =   39057
      End
      Begin MSComCtl2.DTPicker msk_fltInicio 
         Height          =   315
         Left            =   7410
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
         Format          =   489226241
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
         Height          =   195
         Left            =   8895
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
         Height          =   195
         Left            =   7050
         TabIndex        =   16
         Top             =   240
         Width           =   300
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   55
      TabIndex        =   19
      Top             =   0
      Width           =   10800
      _ExtentX        =   19050
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
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   7560
         Top             =   120
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frm_trocaduplicata_filtro.frx":0087
         Count           =   1
      End
   End
   Begin VB.Label Label27 
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
      TabIndex        =   18
      Top             =   1110
      Width           =   825
   End
End
Attribute VB_Name = "frm_trocaduplicata_filtro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chkEmissao_Click()
On Error GoTo tratar_erro

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

txtTexto.Visible = True
If cmbfiltrarpor = "Conta contábil" Or cmbfiltrarpor = "Instituição" Or cmbfiltrarpor = "Cliente" Or cmbfiltrarpor = "Local do desconto" Then
    txtTexto.Visible = False
    cmbTexto.Visible = True
    
    Texto = ""
    cmbTexto.Clear
    Set TBLISTA = CreateObject("adodb.recordset")
    Select Case cmbfiltrarpor
        Case "Conta contábil":
            Set TBLISTA = CreateObject("adodb.recordset")
            TBLISTA.Open "Select tbl_familia.int_codfamilia, tbl_familia.Codigo, tbl_familia.txt_descricao from (tbl_contas_receber INNER JOIN familia_financeiro ON CR.IdIntConta = familia_financeiro.IDConta) INNER JOIN tbl_familia ON tbl_familia.int_codfamilia = familia_financeiro.ID_PC where familia_financeiro.tipoconta = 'R' and CR.Idtrocatitulo <> 0 and CR.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " Group by tbl_familia.int_codfamilia, tbl_familia.Codigo, tbl_familia.txt_descricao", Conexao, adOpenKeyset, adLockOptimistic
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
            TBLISTA.Open "Select IDcliente, Nome_Razao from tbl_contas_receber where Nome_Razao is not null and Idtrocatitulo <> 0 and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " Group by IDcliente, Nome_Razao", Conexao, adOpenKeyset, adLockOptimistic
            If TBLISTA.EOF = False Then
                Do While TBLISTA.EOF = False
                    cmbTexto.AddItem TBLISTA!Nome_Razao
                    cmbTexto.ItemData(cmbTexto.NewIndex) = TBLISTA!IDCliente
                    TBLISTA.MoveNext
                Loop
            End If
            TBLISTA.Close
        Case "Local do desconto":
            Set TBLISTA = CreateObject("adodb.recordset")
            TBLISTA.Open "Select local_troca from troca_titulo group by local_troca", Conexao, adOpenKeyset, adLockOptimistic
            If TBLISTA.EOF = False Then
                Do While TBLISTA.EOF = False
                    If IsNull(TBLISTA!local_troca) = False And TBLISTA!local_troca <> "" Then cmbTexto.AddItem Trim(TBLISTA!local_troca)
                    TBLISTA.MoveNext
                Loop
            End If
            TBLISTA.Close
    End Select
End If
If cmbfiltrarpor = "Borderô" Then chkVencimento.Visible = False Else chkVencimento.Visible = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 10980, 5, True

cmbfiltrarpor = "Borderô"
msk_fltInicio.Value = Date
msk_fltFim.Value = Date
ProcCarregaComboEmpresa Cmb_empresa, False

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
 
Private Sub ProcFiltrar()
On Error GoTo tratar_erro

With msk_fltFim
    If FunVerificaDataFinal(msk_fltInicio.Value, .Value) = False Then
        .Value = Date
        .SetFocus
        Exit Sub
    End If
End With
If chkEmissao.Value = 1 Then
    If cmbfiltrarpor = "Borderô" Or cmbfiltrarpor = "Local do desconto" Then Data_receber = "TT.Data" Else Data_receber = "CR.Emissao"
Else
    Data_receber = "CR.Vencimento"
End If
DataFiltro = ""
If chkVencimento.Value = 1 Or chkEmissao.Value = 1 Then DataFiltro = " and " & Data_receber & " Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "'"
CamposFiltro = "TT.ID, TT.Data, TT.Responsavel, TT.banco_recebedor, TT.local_troca, TT.Vlrtotalresgatado"
INNERJOINTEXTO = "Select " & CamposFiltro & " from ((((troca_titulo TT LEFT JOIN tbl_Contas_receber CR ON CR.Idtrocatitulo = TT.ID) LEFT JOIN familia_financeiro FF ON FF.idconta = CR.IdIntConta) LEFT JOIN tbl_proposta_nota PN ON PN.ID_nota = CR.ID_nota) LEFT JOIN vendas_proposta VP ON VP.Ncotacao = PN.proposta) LEFT JOIN vendas_carteira VC ON VC.Cotacao = VP.Cotacao"
TextoFiltroPadrao = "TT.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & DataFiltro & " group by " & CamposFiltro & " order by TT.ID desc"

With frm_trocaduplicata
    If txtTexto.Visible = True And txtTexto <> "" Or cmbTexto.Visible = True And cmbTexto <> "" Then
        If cmbfiltrarpor = "Conta contábil" Then
            .StrSql_Desconto_Duplicata = INNERJOINTEXTO & " where FF.ID_PC = " & cmbTexto.ItemData(cmbTexto.ListIndex) & " and FF.tipoconta = 'R' and " & TextoFiltroPadrao
        ElseIf cmbfiltrarpor = "Local do desconto" Then
                .StrSql_Desconto_Duplicata = INNERJOINTEXTO & " where TT.local_troca = '" & cmbTexto & "' and " & TextoFiltroPadrao
            ElseIf cmbfiltrarpor = "Cliente" Then
                    .StrSql_Desconto_Duplicata = INNERJOINTEXTO & " where CR.IDcliente = " & cmbTexto.ItemData(cmbTexto.ListIndex) & " and " & TextoFiltroPadrao
                ElseIf cmbfiltrarpor = "Borderô" Then
                        .StrSql_Desconto_Duplicata = INNERJOINTEXTO & " where TT.id = " & txtTexto & " and " & TextoFiltroPadrao
                    Else
                        Select Case cmbfiltrarpor
                            Case "Pedido interno": TextoFiltro = "PN.proposta"
                            Case "Pedido cliente": TextoFiltro = "VC.PCCliente"
                            Case "Nota fiscal": TextoFiltro = "CR.Nfiscal"
                        End Select
                        .StrSql_Desconto_Duplicata = INNERJOINTEXTO & " where " & TextoFiltro & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and " & TextoFiltroPadrao
        End If
    Else
        .StrSql_Desconto_Duplicata = INNERJOINTEXTO & " where " & TextoFiltroPadrao
    End If
    .procCarregalistaPrincipal 1
End With
Unload Me

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
    'Case 3: ProcAjuda
    Case 4: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
