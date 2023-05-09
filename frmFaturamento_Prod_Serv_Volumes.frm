VERSION 5.00
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFaturamento_Prod_Serv_Volumes 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Faturamento | Lista de itens x embalagem"
   ClientHeight    =   6240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6525
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
   ScaleHeight     =   6240
   ScaleWidth      =   6525
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtTotalVolumes 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   4470
      TabIndex        =   4
      Top             =   4770
      Width           =   1845
   End
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   1
      Top             =   5835
      Width           =   6525
      _ExtentX        =   11509
      _ExtentY        =   714
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   435
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6525
      _ExtentX        =   11509
      _ExtentY        =   767
      DibPicture      =   "frmFaturamento_Prod_Serv_Volumes.frx":0000
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
      Icon            =   "frmFaturamento_Prod_Serv_Volumes.frx":1CAD
   End
   Begin DrawSuite2022.USLabel USLabel1 
      Height          =   195
      Left            =   2760
      TabIndex        =   5
      Top             =   4860
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   344
      Caption         =   "Total de embalagens:"
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
      NoHTMLCaption   =   "Total de embalagens:"
   End
   Begin DrawSuite2022.USButton btnCarregar 
      Height          =   405
      Left            =   4500
      TabIndex        =   2
      Top             =   5310
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   714
      Caption         =   "Carregar total volumes"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   5263559
      BorderColorDisabled=   13160660
      BorderColorDown =   4013465
      BorderColorOver =   4408288
      GradientColor1  =   5263559
      GradientColor2  =   5263559
      GradientColor3  =   5263559
      GradientColor4  =   5263559
      GradientColorDisabled1=   13160660
      GradientColorDisabled2=   13160660
      GradientColorDisabled3=   13160660
      GradientColorDisabled4=   13160660
      GradientColorOver1=   4408288
      GradientColorOver2=   4408288
      GradientColorOver3=   4408288
      GradientColorOver4=   4408288
      GradientColorDown1=   4013465
      GradientColorDown2=   4013465
      GradientColorDown3=   4013465
      GradientColorDown4=   4013465
      ShowFocusRect   =   0   'False
      Theme           =   4
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   4065
      Left            =   180
      TabIndex        =   3
      Top             =   570
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
         Text            =   "itens x emb."
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
         Text            =   "Total emb."
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmFaturamento_Prod_Serv_Volumes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnCarregar_Click()
On Error GoTo tratar_erro

If USMsgBox("Deseja realmente carregar o numero de volumes?", vbYesNo, "CAPRIND v5.0") = vbYes Then

        frmFaturamento_Prod_Serv.txtTransp_qtd.Text = Int(txtTotalVolumes.Text)

End If

Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro
Dim Embalagens As Double

Lista.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")

StrSql = "Select DTN.Int_codigo, DTN.int_Cod_Produto as Desenho, int_Qtd as Quantidade, PP.Qtde_embalagem as TTemb from tbl_Detalhes_Nota DTN inner join projproduto PP on PP.Desenho = DTN.int_Cod_Produto where DTN.ID_Nota = '" & frmFaturamento_Prod_Serv.txtId.Text & "' group By DTN.Int_codigo,DTN.int_Qtd, PP.Qtde_embalagem, DTN.int_Cod_Produto, DTN.Int_codigo"
'Debug.print StrSql

TBLISTA.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic

If TBLISTA.EOF = False Then
    Do While TBLISTA.EOF = False
        With Lista.ListItems
            .Add , , TBLISTA!Int_codigo
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Desenho), "", TBLISTA!Desenho)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!TTEmb), "", Format(TBLISTA!TTEmb, "###,##0.00"))
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!quantidade), "", Format(TBLISTA!quantidade, "###,##0.00"))
            Embalagens = "0"
            If TBLISTA!quantidade <> "" And TBLISTA!TTEmb <> "" Then
                If TBLISTA!quantidade <> "0" And TBLISTA!TTEmb <> "0" Then
                    Embalagens = TBLISTA!quantidade / TBLISTA!TTEmb
                End If
            End If
            
            .Item(.Count).SubItems(4) = Format(Embalagens, "###,##0.00")
            
            TotalVolumes = TotalVolumes + Embalagens
            TBLISTA.MoveNext
            Contador = Contador + 1
        End With
    Loop
End If
TBLISTA.Close

txtTotalVolumes = Format(TotalVolumes, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

