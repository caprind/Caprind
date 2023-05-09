VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmCQ_Certificado_Abrir 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'Nenhum
   Caption         =   "CAPRIND v5.0 | CQ | Certificado"
   ClientHeight    =   3135
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   1  'Centralizar no Mestre
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   2445
      Left            =   180
      TabIndex        =   1
      Top             =   540
      Width           =   4245
      Begin DrawSuite2022.USButton btnFiltrar 
         Height          =   465
         Left            =   570
         TabIndex        =   4
         Top             =   1680
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   820
         DibPicture      =   "frmCQ_Certificado_Abrir.frx":0000
         BorderColor     =   5263559
         BorderColorDisabled=   13160660
         BorderColorDown =   4013465
         BorderColorOver =   4408288
         Caption         =   "Filtrar lote (OP)"
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
      Begin VB.TextBox txtLote 
         Alignment       =   2  'Centralizar
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1860
         TabIndex        =   2
         Top             =   1140
         Width           =   1905
      End
      Begin DrawSuite2022.USLabel USLabel1 
         Height          =   480
         Left            =   510
         Top             =   390
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   847
         Caption         =   "Digite o numero do lote (OP) e clique no botão filtrar."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   128
         NoHTMLCaption   =   $"frmCQ_Certificado_Abrir.frx":3650
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparente
         Caption         =   "N° do lote (OP) : "
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
         Left            =   600
         TabIndex        =   3
         Top             =   1200
         Width           =   1245
      End
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   714
      DibPicture      =   "frmCQ_Certificado_Abrir.frx":368C
      CaptionDelimiter=   "|"
      CaptionOnCenter =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Icon            =   "frmCQ_Certificado_Abrir.frx":6CDC
      ShowMaximize    =   0   'False
      ShowMinimize    =   0   'False
   End
End
Attribute VB_Name = "frmCQ_Certificado_Abrir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnFiltrar_Click()
On Error GoTo tratar_erro

If txtLote.Text = "" Then
USMsgBox "Digite o numero da ordem de produção", vbCritical, "CAPRIND v5.0"
Exit Sub
End If

Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "select * from producao where ordem = " & txtLote.Text & "", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then

With frmCQ_Certificado_Analise
Var1 = Year(Format(TBLISTA!data, "dd/mm/yy"))
Var1 = Right(Var1, 2)
Var1 = Var1 & Month(TBLISTA!data)

Select Case Len(TBLISTA!Ordem)
Case 1:
Var1 = Var1 & "000000" & TBLISTA!Ordem
Case 2:
Var1 = Var1 & "00000" & TBLISTA!Ordem
Case 3:
Var1 = Var1 & "0000" & TBLISTA!Ordem
Case 4:
Var1 = Var1 & "000" & TBLISTA!Ordem
Case 5:
Var1 = Var1 & "00" & TBLISTA!Ordem
Case 6:
Var1 = Var1 & "0" & TBLISTA!Ordem
Case 7:
Var1 = Var1 & TBLISTA!Ordem
End Select

.txtLote = Var1
.txtProduto = TBLISTA!Desenho
.txtDescricao = TBLISTA!Produto
.txtQuant_Env = TBLISTA!Quant
.txtAnalista = pubUsuario
.txtData = Date
End With

Else
USMsgBox "Não existe esse lote no sistema", vbInformation, "CAPRIND v5.0"
Exit Sub
End If
TBLISTA.Close
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcBuscaLote()
On Error GoTo tratar_erro



Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub ProcNovo()
On Error GoTo tratar_erro

'ProcLimpaCampos
'ProcGerarCodigoCA
'ProcBuscaLote


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

'txtLote.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
