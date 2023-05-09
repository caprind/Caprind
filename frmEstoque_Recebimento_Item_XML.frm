VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form frmEstoque_Recebimento_Item_XML 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'Nenhum
   Caption         =   "Faturamento | Nota fiscal | Importação XML - CAPRIND v5.0"
   ClientHeight    =   8775
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8070
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEstoque_Recebimento_Item_XML.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8775
   ScaleWidth      =   8070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Centralizar no Mestre
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Dados do item para importação do XML"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8025
      Left            =   210
      TabIndex        =   1
      Top             =   630
      Width           =   7635
      Begin VB.TextBox txtPedido 
         Alignment       =   2  'Centralizar
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2310
         TabIndex        =   6
         Top             =   2460
         Width           =   1005
      End
      Begin DrawSuite2022.USLabel USLabel2 
         Height          =   585
         Left            =   1560
         Top             =   360
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   1032
         Caption         =   $"frmEstoque_Recebimento_Item_XML.frx":000C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   128
         NoHTMLCaption   =   $"frmEstoque_Recebimento_Item_XML.frx":00BB
      End
      Begin DrawSuite2022.USButton btnImportar 
         Height          =   615
         Left            =   5880
         TabIndex        =   5
         Top             =   2430
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1085
         DibPicture      =   "frmEstoque_Recebimento_Item_XML.frx":016C
         BorderColor     =   5263559
         BorderColorDisabled=   13160660
         BorderColorDown =   4013465
         BorderColorOver =   4408288
         Caption         =   "Importar item"
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
         ToolTipTitle    =   "Continuar a importação..."
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   2  'Centralizar
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2310
         TabIndex        =   3
         Top             =   2010
         Width           =   5055
      End
      Begin VB.TextBox txtdescricao 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   2310
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1290
         Width           =   5055
      End
      Begin DrawSuite2022.USLabel USLabel1 
         Height          =   195
         Left            =   330
         Top             =   2040
         Width           =   3225
         _ExtentX        =   5689
         _ExtentY        =   344
         Caption         =   "Informe o código do item :"
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
         NoHTMLCaption   =   "Informe o código do item :"
      End
      Begin DrawSuite2022.USButton BtnPedido 
         Height          =   345
         Left            =   3330
         TabIndex        =   7
         Top             =   2460
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   609
         DibPicture      =   "frmEstoque_Recebimento_Item_XML.frx":3554
         BorderColor     =   5263559
         BorderColorDisabled=   13160660
         BorderColorDown =   4013465
         BorderColorOver =   4408288
         Caption         =   "Listar itens"
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
         ToolTipTitle    =   "Continuar a importação..."
      End
      Begin DrawSuite2022.USLabel USLabel3 
         Height          =   195
         Left            =   840
         Top             =   2490
         Width           =   3225
         _ExtentX        =   5689
         _ExtentY        =   344
         Caption         =   "Pedido de compra :"
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
         NoHTMLCaption   =   "Pedido de compra :"
      End
      Begin MSComctlLib.ListView Listprod 
         Height          =   4635
         Left            =   90
         TabIndex        =   8
         Top             =   3210
         Width           =   7425
         _ExtentX        =   13097
         _ExtentY        =   8176
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
         Appearance      =   0
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
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Tag             =   "T"
            Text            =   "Cód. interno"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Descricão"
            Object.Width           =   9596
         EndProperty
      End
      Begin DrawSuite2022.USAlphaImage USAlphaImage1 
         Height          =   990
         Left            =   450
         Top             =   240
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   1746
         Image           =   "frmEstoque_Recebimento_Item_XML.frx":6BA4
         Props           =   5
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparente
         Caption         =   "Descricao do item no XML :"
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
         Left            =   300
         TabIndex        =   4
         Top             =   1290
         Width           =   2895
      End
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   465
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8070
      _ExtentX        =   14235
      _ExtentY        =   820
      DibPicture      =   "frmEstoque_Recebimento_Item_XML.frx":97C2
      CaptionOnCenter =   -1  'True
      EnableMaximizeButton=   0   'False
      EnableMinimizeButton=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowMaximize    =   0   'False
      ShowMinimize    =   0   'False
   End
End
Attribute VB_Name = "frmEstoque_Recebimento_Item_XML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnImportar_Click()
On Error GoTo tratar_erro

If txtCodigo.Text <> "" Then
Cod_produto = txtCodigo.Text
Unload Me
Else
USMsgBox "Informe o codigo interno do item a ser importado", vbInformation, "CAPRIND v5."
Exit Sub
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub BtnPedido_Click()
On Error GoTo tratar_erro

Listprod.ListItems.Clear

Set TBLISTA = CreateObject("adodb.recordset")
StrSql = "Select CPL.idlista,CP.Pedido, CPL.Desenho, CPL.Descricao from Compras_pedido CP inner join Compras_pedido_lista CPL on CP.IDPedido = CPL.IDPedido Where Pedido = '" & txtPedido.Text & "'"
TBLISTA.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    Contador = 0
    Do While TBLISTA.EOF = False
        With Listprod.ListItems
            .Add , , TBLISTA!IDlista
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Desenho), "", TBLISTA!Desenho)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Descricao), "", TBLISTA!Descricao)
        End With
        TBLISTA.MoveNext
        Contador = Contador + 1
    Loop
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
