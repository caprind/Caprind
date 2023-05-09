VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form frmFaturamento_Prod_Serv_pesoItens 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Faturamento | Peso bruto dos itens da nota"
   ClientHeight    =   6465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6765
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6465
   ScaleWidth      =   6765
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USCheckBox chkLiquido 
      Height          =   285
      Left            =   390
      TabIndex        =   5
      Top             =   4980
      Width           =   2625
      _ExtentX        =   4630
      _ExtentY        =   503
      Caption         =   "Carregar valor no peso liquido"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocusRect   =   -1  'True
      Value           =   1
   End
   Begin DrawSuite2022.USLabel USLabel1 
      Height          =   195
      Left            =   3300
      Top             =   5010
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   344
      Caption         =   "Peso bruto total:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      NoHTMLCaption   =   "Peso bruto total:"
   End
   Begin DrawSuite2022.USTextBoxEx txtPesototal 
      Height          =   375
      Left            =   4650
      TabIndex        =   4
      Top             =   4950
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      Alignment       =   2
      AutoFormatDate  =   -1  'True
      BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Decimals        =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontDisabled {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontOver {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16512
      MaskType        =   1
   End
   Begin DrawSuite2022.USButton btnCarregar 
      Height          =   405
      Left            =   4650
      TabIndex        =   3
      Top             =   5460
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   714
      BorderColor     =   5263559
      BorderColorDisabled=   13160660
      BorderColorDown =   4013465
      BorderColorOver =   4408288
      Caption         =   "Carregar peso bruto"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
      Theme           =   4
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   405
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6765
      _ExtentX        =   11933
      _ExtentY        =   741
      DibPicture      =   "frmFaturamento_Prod_Serv_PesoItens.frx":0000
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
      Icon            =   "frmFaturamento_Prod_Serv_PesoItens.frx":D817
      ShowMaximize    =   0   'False
      ShowMinimize    =   0   'False
   End
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   6060
      Width           =   6765
      _ExtentX        =   11933
      _ExtentY        =   714
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   4065
      Left            =   330
      TabIndex        =   2
      Top             =   690
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   7170
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
      Appearance      =   0
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "N"
         Text            =   "Código"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Peso Unitario"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Quantidade"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Peso bruto"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmFaturamento_Prod_Serv_pesoItens"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCarregar_Click()
On Error GoTo tratar_erro

If USMsgBox("Deseja realmente carregar o valor do peso bruto na NFe?", vbYesNo, "CAPRIND v5.0") = vbYes Then

    If Formulario <> "Estoque/Ordem de faturamento" Then
        frmFaturamento_Prod_Serv.txtTransp_pesoBruto.Text = txtPesototal.Text
    Else
        frmEstoque_Ordem_Faturamento.txtTransp_pesoBruto.Text = txtPesototal.Text
    End If


    If chkLiquido.Value = Checked Then
        If Formulario <> "Estoque/Ordem de faturamento" Then
            frmFaturamento_Prod_Serv.txtTransp_pesoliquido.Text = txtPesototal.Text
        Else
            frmEstoque_Ordem_Faturamento.txtTransp_pesoliquido.Text = txtPesototal.Text
        End If
    End If
End If

Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")

If Formulario <> "Estoque/Ordem de faturamento" Then
    StrSql = "Select DNF.ID,DTNF.int_Cod_Produto,PP.pBruto as PesoUnitario, DTNF.int_Qtd as Quantidade, sum(DTNF.int_Qtd*PP.Pbruto)as PesoBruto from tbl_Dados_Nota_Fiscal DNF inner join tbl_Detalhes_Nota DTNF on DTNF.ID_Nota = DNF.id inner Join projproduto PP on PP.codproduto = DTNF.codproduto where DNF.ID = '" & frmFaturamento_Prod_Serv.txtid.Text & "' group By DTNF.Int_codigo,DTNF.int_Qtd, PP.pBruto, DTNF.int_Cod_Produto, DNF.ID"
Else
    StrSql = "Select DNF.ID,DTNF.int_Cod_Produto,PP.pBruto as PesoUnitario, DTNF.int_Qtd as Quantidade, sum(DTNF.int_Qtd*PP.Pbruto)as PesoBruto from tbl_Dados_Nota_Fiscal DNF inner join tbl_Detalhes_Nota DTNF on DTNF.ID_Nota = DNF.id inner Join projproduto PP on PP.codproduto = DTNF.codproduto where DNF.ID = '" & frmEstoque_Ordem_Faturamento.txtid.Text & "' group By DTNF.Int_codigo,DTNF.int_Qtd, PP.pBruto, DTNF.int_Cod_Produto, DNF.ID"
End If

TBLISTA.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic

If TBLISTA.EOF = False Then
    Do While TBLISTA.EOF = False
        With Lista.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!int_Cod_Produto), "", TBLISTA!int_Cod_Produto)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!PesoUnitario), "", Format(TBLISTA!PesoUnitario, "###,##0.0000"))
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!quantidade), "", Format(TBLISTA!quantidade, "###,##0.0000"))
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!PesoBruto), "", Format(TBLISTA!PesoBruto, "###,##0.0000"))
            PesoBruto = PesoBruto + TBLISTA!PesoBruto
            TBLISTA.MoveNext
            contador = contador + 1
        End With
    Loop
End If
TBLISTA.Close

txtPesototal = Format(PesoBruto, "###,##0.0000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
