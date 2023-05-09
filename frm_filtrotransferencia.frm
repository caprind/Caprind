VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frm_filtrotransferencia 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Instituições - Localizar movimentação financeira"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8895
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_filtrotransferencia.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   8895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Chk_periodo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Período de :"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   4320
      TabIndex        =   3
      Top             =   2760
      Width           =   1185
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Height          =   1515
      Left            =   60
      TabIndex        =   13
      Top             =   990
      Width           =   8805
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Height          =   510
         Left            =   3840
         TabIndex        =   16
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
            TabIndex        =   9
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
            TabIndex        =   7
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
            TabIndex        =   6
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
            TabIndex        =   8
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
         ItemData        =   "frm_filtrotransferencia.frx":1042
         Left            =   180
         List            =   "frm_filtrotransferencia.frx":105B
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Opções para filtro."
         Top             =   390
         Width           =   3585
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
         ItemData        =   "frm_filtrotransferencia.frx":10E3
         Left            =   180
         List            =   "frm_filtrotransferencia.frx":10E5
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "Texto para pesquisa."
         Top             =   1050
         Width           =   8445
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
         Visible         =   0   'False
         Width           =   8445
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Texto para pesquisa"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3667
         TabIndex        =   15
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
         Left            =   1552
         TabIndex        =   14
         Top             =   180
         Width           =   840
      End
   End
   Begin VB.Frame frm_filtro 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   60
      TabIndex        =   10
      Top             =   2520
      Width           =   8805
      Begin MSComCtl2.DTPicker txtde 
         Height          =   315
         Left            =   5520
         TabIndex        =   4
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
         Format          =   489226243
         CurrentDate     =   39057
      End
      Begin MSComCtl2.DTPicker txta 
         Height          =   315
         Left            =   7230
         TabIndex        =   5
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
         Format          =   489226241
         CurrentDate     =   39057
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "à"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   7035
         TabIndex        =   11
         Top             =   210
         Width           =   105
      End
   End
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   3120
      Top             =   150
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frm_filtrotransferencia.frx":10E7
      Count           =   1
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
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
   End
End
Attribute VB_Name = "frm_filtrotransferencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Chk_periodo_Click()
On Error GoTo tratar_erro

If Chk_periodo.Value = 1 Then
    frm_filtro.Enabled = True
    txtDe.SetFocus
Else
    frm_filtro.Enabled = False
    txtDe.Value = Date
    txta.Value = Date
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbfiltrarpor_Click()
On Error GoTo tratar_erro

Optinicio.Value = True
If cmbfiltrarpor = "Valor" Or cmbfiltrarpor = "Documento" Then
    txtTexto.Visible = True
    cmbTexto.Visible = False
    If cmbfiltrarpor = "Documento" Then
        Optfim.Value = True
    ElseIf cmbfiltrarpor = "Valor" Then
            If txtTexto <> "" Then
                VerifNumero = txtTexto
                ProcVerificaNumero
                If VerifNumero = False Then
                    txtTexto = ""
                    txtTexto.SetFocus
                    Exit Sub
                End If
            End If
    End If
Else
    txtTexto.Visible = False
    cmbTexto.Visible = True
    With cmbTexto
        .Clear
        .AddItem ""
        Set TBLISTA = CreateObject("adodb.recordset")
        If cmbfiltrarpor = "Conta contábil" Or cmbfiltrarpor = "Código da conta contábil" Then
            If frm_Instituicoes.SSTab3.Tab = 0 Then
                INNERJOINTEXTO = "IT.id_transf = FF.IDConta"
                TextoFiltroTipo = "FF.Deposito_transf = 'True' and (IT.Tipo = 'T' or IT.Tipo = 'D')"
            Else
                INNERJOINTEXTO = "IT.IDintconta = FF.IDConta and IT.Tipo = FF.Tipoconta"
                TextoFiltroTipo = "FF.Deposito_transf = 'False' and (IT.Tipo = 'P' or IT.Tipo = 'R')"
            End If
            Set TBLISTA = CreateObject("adodb.recordset")
            TBLISTA.Open "Select F.int_codfamilia, F.Codigo, F.txt_descricao from familia_financeiro FF INNER JOIN tbl_instituicoes_transf IT ON " & INNERJOINTEXTO & " INNER JOIN tbl_familia F ON F.int_codfamilia = FF.ID_PC where IT.id_banco_rem = " & frm_Instituicoes.txtCodBanco & " and " & TextoFiltroTipo & " Group by F.int_codfamilia, F.Codigo, F.txt_descricao", Conexao, adOpenKeyset, adLockOptimistic
            If TBLISTA.EOF = False Then
                Do While TBLISTA.EOF = False
                    If cmbfiltrarpor = "Conta contábil" Then .AddItem TBLISTA!Txt_descricao & " - " & TBLISTA!CODIGO Else .AddItem TBLISTA!CODIGO & " - " & TBLISTA!Txt_descricao
                    .ItemData(cmbTexto.NewIndex) = TBLISTA!int_codfamilia
                    TBLISTA.MoveNext
                Loop
            End If
            TBLISTA.Close
        Else
            If cmbfiltrarpor = "Instituição bancária recebedora" Then
                Set TBLISTA = CreateObject("adodb.recordset")
                TBLISTA.Open "Select id_banco_rec, banco_recebedor from tbl_instituicoes_transf where id_banco_rem = " & frm_Instituicoes.txtCodBanco & " and  banco_recebedor IS NOT NULL Group by id_banco_rec, banco_recebedor", Conexao, adOpenKeyset, adLockOptimistic
                If TBLISTA.EOF = False Then
                    Do While TBLISTA.EOF = False
                        .AddItem TBLISTA!banco_recebedor
                        .ItemData(cmbTexto.NewIndex) = TBLISTA!id_banco_rec
                        TBLISTA.MoveNext
                    Loop
                End If
                TBLISTA.Close
            Else
                If frm_Instituicoes.SSTab3.Tab = 0 Then TextoFiltroTipo = "(Tipo = 'T' or Tipo = 'D')" Else TextoFiltroTipo = "(Tipo = 'P' or Tipo = 'R')"
                Set TBLISTA = CreateObject("adodb.recordset")
                TBLISTA.Open "Select Formabaixa from tbl_instituicoes_transf where id_banco_rem = " & frm_Instituicoes.txtCodBanco & " and " & TextoFiltroTipo & " and Formabaixa IS NOT NULL Group by Formabaixa", Conexao, adOpenKeyset, adLockOptimistic
                If TBLISTA.EOF = False Then
                    Do While TBLISTA.EOF = False
                        .AddItem TBLISTA!FormaBaixa
                        TBLISTA.MoveNext
                    Loop
                End If
                TBLISTA.Close
            End If
        End If
    End With
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 8805, 5, True
With cmbfiltrarpor
    Select Case frm_Instituicoes.SSTab3.Tab
        Case 0: Caption = "Instituições - Localizar depósito ou tranferência"
        Case 1: Caption = "Instituições - Localizar saque"
        Case 2: Caption = "Instituições - Localizar tarifa"
    End Select
    .Clear
    .AddItem "Valor"
    If frm_Instituicoes.SSTab3.Tab = 0 Or frm_Instituicoes.SSTab3.Tab = 2 Then
        .AddItem "Código da conta contábil"
        .AddItem "Conta contábil"
        If frm_Instituicoes.SSTab3.Tab = 0 Then
            .AddItem "Forma da movimentação"
            .AddItem "Instituição bancária recebedora"
            .AddItem "Documento"
        Else
            .AddItem "Forma da baixa"
        End If
        .Text = "Conta contábil"
    Else
        .Text = "Valor"
    End If
End With
txtDe.Value = Date
txta.Value = Date

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

With txta
    If FunVerificaDataFinal(txtDe.Value, .Value) = False Then
        .Value = Date
        .SetFocus
        Exit Sub
    End If
End With

With frm_Instituicoes
    DataFiltro = ""
    DataFiltroRel = ""
    If Chk_periodo.Value = 1 Then
        DataFiltro = " and IT.data_transf Between '" & Format(txtDe.Value, "Short Date") & "' And '" & Format(txta.Value, "Short Date") & "'"
        DataFiltroRel = " and {tbl_instituicoes_transf.data_transf}>=Date(" & Year(txtDe.Value) & "," & Month(txtDe.Value) & "," & Day(txtDe.Value) & ") and {tbl_instituicoes_transf.data_transf}<= Date(" & Year(txta.Value) & "," & Month(txta.Value) & "," & Day(txta.Value) & ")"
    End If
    Select Case .SSTab3.Tab
        Case 0:
            TextoFiltroTipo = "(IT.Tipo = 'T' or IT.Tipo = 'D')"
            TextoFiltroTipoRel = "({tbl_instituicoes_transf.Tipo} = 'T' or {tbl_instituicoes_transf.Tipo} = 'D')"
            If cmbfiltrarpor = "Conta contábil" Or cmbfiltrarpor = "Código da conta contábil" Then
                If cmbTexto <> "" Then
                    TextoFiltroTipo = TextoFiltroTipo & " and FF.Deposito_transf = 'True'"
                    TextoFiltroTipoRel = TextoFiltroTipoRel & " and {Familia_financeiro.Deposito_transf} = True"
                End If
            End If
            INNERJOINTEXTO1 = "IT.id_transf = FF.IDConta"
        Case 1:
            TextoFiltroTipo = "IT.Tipo = 'S'"
            TextoFiltroTipoRel = "{tbl_instituicoes_transf.Tipo} = 'S'"
        Case 2:
            TextoFiltroTipo = "(IT.Tipo = 'P' or IT.Tipo = 'R')"
            TextoFiltroTipoRel = "({tbl_instituicoes_transf.Tipo} = 'P' or {tbl_instituicoes_transf.Tipo} = 'R')"
            If cmbfiltrarpor = "Conta contábil" Or cmbfiltrarpor = "Código da conta contábil" Then
                If cmbTexto <> "" Then TextoFiltroTipo = TextoFiltroTipo & " and FF.Deposito_transf = 'False'"
                TextoFiltroTipoRel = TextoFiltroTipoRel & " and {Familia_financeiro.Deposito_transf} = False"
            End If
            INNERJOINTEXTO1 = "IT.IDintconta = FF.IDConta"
    End Select
    TextoFiltroPadrao = "IT.id_banco_rem = " & .txtCodBanco & DataFiltro & " and " & TextoFiltroTipo & " order by IT.data_transf desc, IT.id_transf"
    TextoFiltroPadraoRel = "{tbl_instituicoes_transf.id_banco_rem} = " & .txtCodBanco & DataFiltroRel & " and " & TextoFiltroTipoRel
    INNERJOINTEXTO = "Select * from tbl_instituicoes_transf IT"
    
    If txtTexto.Visible = True And txtTexto <> "" Or cmbTexto.Visible = True And cmbTexto <> "" Then
        If cmbTexto.Visible = True Then
            If cmbfiltrarpor = "Conta contábil" Or cmbfiltrarpor = "Código da conta contábil" Then
                INNERJOINTEXTO = "Select IT.* from familia_financeiro FF INNER JOIN tbl_instituicoes_transf IT ON " & INNERJOINTEXTO1 & IIf(.SSTab3.Tab = 2, " and IT.Tipo = FF.Tipoconta", "") & " INNER JOIN tbl_familia F ON F.int_codfamilia = FF.ID_PC"
                TextoFiltro = "FF.ID_PC = " & cmbTexto.ItemData(cmbTexto.ListIndex)
                .FormulaRel_Instituicao = "{familia_financeiro.ID_PC} = " & cmbTexto.ItemData(cmbTexto.ListIndex) & " and " & TextoFiltroPadraoRel
            Else
                If cmbfiltrarpor = "Instituição bancária recebedora" Then CampoFiltro = "banco_recebedor" Else CampoFiltro = "FormaBaixa"
                TextoFiltro = CampoFiltro & " = '" & cmbTexto & "'"
                .FormulaRel_Instituicao = "{familia_financeiro." & CampoFiltro & "} = '" & cmbTexto & "' and " & TextoFiltroPadraoRel
            End If
        Else
            If cmbfiltrarpor = "Valor" Then
                valor = txtTexto
                NovoValor = Replace(valor, ",", ".")
                TextoFiltro = "Valor_transf = " & NovoValor
                .FormulaRel_Instituicao = "{familia_financeiro.Valor_transf} = " & NovoValor & " and " & TextoFiltroPadraoRel
            Else
                TextoFiltro = "IT.NDoctoBaixa" & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto)
                .FormulaRel_Instituicao = "{familia_financeiro.NDoctoBaixa}" & FunVerifTipoFiltroIMFRel(Optinicio, Optmeio, Optfim, optIgual, txtTexto)
            End If
        End If
    Else
        TextoFiltro = ""
        .FormulaRel_Instituicao = TextoFiltroPadraoRel
    End If
    
    TextoFiltro1 = INNERJOINTEXTO & " where " & IIf(TextoFiltro = "", TextoFiltro, TextoFiltro & " and ") & TextoFiltroPadrao
    Select Case .SSTab3.Tab
        Case 0:
            .Instituicao_Localizar_Transf = TextoFiltro1
            .ProcCarregaListaTransf
        Case 1:
            .Instituicao_Localizar_Saque = TextoFiltro1
            .ProcCarregaListaSaque
        Case 2:
            .Instituicao_Localizar_Tarifa = TextoFiltro1
            .ProcCarregaListaTarifa
    End Select
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

If cmbfiltrarpor = "Valor" And txtTexto <> "" Then
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

If cmbfiltrarpor = "Valor" And txtTexto <> "" Then txtTexto = Format(txtTexto, "###,##0.00")

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
