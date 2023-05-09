VERSION 5.00
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Begin VB.Form frmProdAlterarRM 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Requisição | Alterar quantidade"
   ClientHeight    =   3300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4905
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3300
   ScaleWidth      =   4905
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   330
      TabIndex        =   2
      Top             =   630
      Width           =   4125
      Begin DrawSuite2022.USButton btnGravar 
         Height          =   735
         Left            =   330
         TabIndex        =   6
         ToolTipText     =   "Gravar nova quantidade na requisição de materiais"
         Top             =   1470
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   1296
         DibPicture      =   "frmProdAlterarRM.frx":0000
         Caption         =   "Gravar alteração"
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
         PicAlign        =   8
         PicSize         =   2
         PicSizeH        =   24
         PicSizeW        =   24
         ShowFocusRect   =   0   'False
         Theme           =   4
      End
      Begin VB.TextBox txtQTNova 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   1920
         TabIndex        =   0
         Top             =   840
         Width           =   1755
      End
      Begin VB.TextBox txtQTAtual 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H00000080&
         Height          =   345
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   330
         Width           =   1755
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Quantidade nova:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   330
         TabIndex        =   4
         Top             =   930
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Quantidade atual:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   300
         TabIndex        =   3
         Top             =   420
         Width           =   2295
      End
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   405
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4905
      _ExtentX        =   8652
      _ExtentY        =   714
      DibPicture      =   "frmProdAlterarRM.frx":8A05
      CaptionDelimiter=   "|"
      CaptionOnCenter =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Icon            =   "frmProdAlterarRM.frx":ECE9
   End
End
Attribute VB_Name = "frmProdAlterarRM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnGravar_Click()
On Error GoTo tratar_erro

If USMsgBox("Deseja realmente alterar a quantidade do item na requisição?", vbYesNo, "CAPRIND v5.0") = vbNo Then
Exit Sub
End If
QTNova = Replace(txtQTNova.Text, ".", "")
QTNova = Replace(QTNova, ",", ".")
'USMsgBox QTNova

StrSql = "Update Producaomaterial set Quantidade = '" & QTNova & "', Requisitado = '" & QTNova & "' Where idMateriaprima = '" & frmprod.ListaRequisicao.SelectedItem & "'"
'Debug.print StrSql


Conexao.Execute ("Update Producaomaterial set Quantidade = '" & QTNova & "', Requisitado = '" & QTNova & "' Where idMateriaprima = '" & frmprod.ListaRequisicao.SelectedItem & "'")
USMsgBox "Quantidade alterada com sucesso!", vbInformation, "CAPRIND v5.0"
frmprod.ProcCarregaListaRequisicao
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

txtQTAtual.Text = frmprod.ListaRequisicao.SelectedItem.ListSubItems.Item(4).Text

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtQTNova_LostFocus()
On Error GoTo tratar_erro

txtQTNova.Text = Format(txtQTNova.Text, "###,##0.00000000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
