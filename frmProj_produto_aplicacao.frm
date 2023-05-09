VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form frmProj_produto_aplicacao 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Engenharia | Aplicação"
   ClientHeight    =   5640
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6870
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5640
   ScaleWidth      =   6870
   StartUpPosition =   1  'CenterOwner
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   5
      Top             =   5235
      Width           =   6870
      _ExtentX        =   12118
      _ExtentY        =   714
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Dados do item"
      Height          =   915
      Left            =   150
      TabIndex        =   1
      Top             =   540
      Width           =   6495
      Begin DrawSuite2022.USLabel USLabel1 
         Height          =   195
         Left            =   660
         Top             =   270
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   344
         Caption         =   "Código"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
         NoHTMLCaption   =   "Código"
      End
      Begin VB.TextBox txtDescricao 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   1680
         TabIndex        =   3
         Top             =   480
         Width           =   4575
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   1425
      End
      Begin DrawSuite2022.USLabel USLabel2 
         Height          =   195
         Left            =   3660
         Top             =   270
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   344
         Caption         =   "Descrição"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
         NoHTMLCaption   =   "Descrição"
      End
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6870
      _ExtentX        =   12118
      _ExtentY        =   661
      DibPicture      =   "frmProj_produto_aplicacao.frx":0000
      CaptionDelimiter=   "|"
      CaptionOnCenter =   -1  'True
      EnableMaximizeButton=   0   'False
      EnableMinimizeButton=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Icon            =   "frmProj_produto_aplicacao.frx":3650
      ShowMaximize    =   0   'False
      ShowMinimize    =   0   'False
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   3435
      Left            =   150
      TabIndex        =   4
      Top             =   1470
      Width           =   6500
      _ExtentX        =   11456
      _ExtentY        =   6059
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Object.Width           =   442
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Object.Tag             =   "T"
         Text            =   "Cód. interno"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Descrição"
         Object.Width           =   6880
      EndProperty
   End
End
Attribute VB_Name = "frmProj_produto_aplicacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Error GoTo tratar_erro
txtCodigo.Text = Desenho
txtdescricao.Text = DesenhoProduto

contador = 1
StrSql = "select pc.Desenho as CodMaterial, pc.descricao as descMaterial , pp.Desenho as codproduto, pp.descricao as descproduto from Projconjunto as pc inner join projproduto as pp on pc.codproduto= pp.codproduto where pc.Desenho = '" & Desenho & "'"

Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
Do While TBLISTA.EOF = False

    With Lista.ListItems
        .Add , , contador
        .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Codproduto), "", TBLISTA!Codproduto)
        .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!descproduto), "", TBLISTA!descproduto)
    End With
contador = contador + 1
TBLISTA.MoveNext
Loop
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
