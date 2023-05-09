VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmCertificado_abrir 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Qualidade - Controle de certificados - Localizar"
   ClientHeight    =   2520
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   8925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   7710
      Top             =   240
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frmCertificado_abrir.frx":0000
      Count           =   1
   End
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
      MouseIcon       =   "frmCertificado_abrir.frx":21F3
      MousePointer    =   99  'Custom
      Picture         =   "frmCertificado_abrir.frx":2345
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Cancela e fecha formulário (Esc)"
      Top             =   165
      Width           =   570
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   1515
      Left            =   55
      TabIndex        =   7
      Top             =   990
      Width           =   8805
      Begin VB.Frame Frame11 
         BackColor       =   &H00E0E0E0&
         Height          =   510
         Left            =   3810
         TabIndex        =   11
         Top             =   210
         WhatsThisHelpID =   210
         Width           =   4785
         Begin VB.OptionButton optIgual 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Igual"
            Height          =   255
            Left            =   3930
            TabIndex        =   5
            Top             =   180
            Width           =   705
         End
         Begin VB.OptionButton Optmeio 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Meio frase"
            Height          =   255
            Left            =   1470
            TabIndex        =   3
            Top             =   180
            Width           =   1275
         End
         Begin VB.OptionButton Optinicio 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Início frase"
            Height          =   255
            Left            =   180
            TabIndex        =   2
            Top             =   180
            Value           =   -1  'True
            Width           =   1275
         End
         Begin VB.OptionButton Optfim 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Fim frase"
            Height          =   255
            Left            =   2760
            TabIndex        =   4
            Top             =   180
            Width           =   1155
         End
      End
      Begin VB.ComboBox cmbfiltrarpor 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         ItemData        =   "frmCertificado_abrir.frx":3387
         Left            =   180
         List            =   "frmCertificado_abrir.frx":339D
         MouseIcon       =   "frmCertificado_abrir.frx":33F3
         MousePointer    =   99  'Custom
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Opções para filtro."
         Top             =   390
         Width           =   3555
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
         MouseIcon       =   "frmCertificado_abrir.frx":36FD
         TabIndex        =   1
         ToolTipText     =   "Texto para pesquisa."
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
         Left            =   3645
         TabIndex        =   9
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
         Left            =   1537
         TabIndex        =   8
         Top             =   180
         Width           =   840
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   55
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft3     =   44
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
      ButtonLeft4     =   87
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
      ButtonLeft5     =   119
      ButtonTop5      =   2
      ButtonWidth5    =   24
      ButtonHeight5   =   24
   End
End
Attribute VB_Name = "frmCertificado_abrir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

With frmCertificado
    .Lista.ListItems.Clear
    If txtTexto <> "" Then
        If cmbfiltrarpor = "Certificado" Then
            If Optinicio.Value = True Then .StrSql_Certificado = "Select * from certificado where Certificado like '" & txtTexto & "%' order by Ordem"
            If Optmeio.Value = True Then .StrSql_Certificado = "Select * from certificado where Certificado like '%" & txtTexto & "%' order by Ordem"
            If Optfim.Value = True Then .StrSql_Certificado = "Select * from certificado where Certificado like '%" & txtTexto & "' order by Ordem"
            If optIgual.Value = True Then .StrSql_Certificado = "Select * from certificado where Certificado = '" & txtTexto & "' order by Ordem"
        End If
        If cmbfiltrarpor = "Ordem" Then
            If Optinicio.Value = True Then .StrSql_Certificado = "Select * from certificado where Ordem like '" & txtTexto & "%' order by Ordem"
            If Optmeio.Value = True Then .StrSql_Certificado = "Select * from certificado where Ordem like '%" & txtTexto & "%' order by Ordem"
            If Optfim.Value = True Then .StrSql_Certificado = "Select * from certificado where Ordem like '%" & txtTexto & "' order by Ordem"
            If optIgual.Value = True Then .StrSql_Certificado = "Select * from certificado where Ordem = '" & txtTexto & "' order by Ordem"
        End If
        If cmbfiltrarpor = "Pedido interno" Then
            If Optinicio.Value = True Then .StrSql_Certificado = "Select certificado.* FROM Certificado INNER JOIN Producao ON Certificado.Ordem = Producao.Ordem where Producao.Lista like '" & txtTexto & "%' order by Certificado.Ordem"
            If Optmeio.Value = True Then .StrSql_Certificado = "Select certificado.* FROM Certificado INNER JOIN Producao ON Certificado.Ordem = Producao.Ordem where Producao.Lista like '%" & txtTexto & "%' order by Certificado.Ordem"
            If Optfim.Value = True Then .StrSql_Certificado = "Select certificado.* FROM Certificado INNER JOIN Producao ON Certificado.Ordem = Producao.Ordem where Producao.Lista like '%" & txtTexto & "' order by Certificado.Ordem"
            If optIgual.Value = True Then .StrSql_Certificado = "Select certificado.* FROM Certificado INNER JOIN Producao ON Certificado.Ordem = Producao.Ordem where Producao.Lista = '" & txtTexto & "' order by Certificado.Ordem"
        End If
        If cmbfiltrarpor = "Código interno" Then
            If Optinicio.Value = True Then .StrSql_Certificado = "Select certificado.* FROM Certificado INNER JOIN Producao ON Certificado.Ordem = Producao.Ordem where Producao.Desenho like '" & txtTexto & "%' order by Certificado.Ordem"
            If Optmeio.Value = True Then .StrSql_Certificado = "Select certificado.* FROM Certificado INNER JOIN Producao ON Certificado.Ordem = Producao.Ordem where Producao.Desenho like '%" & txtTexto & "%' order by Certificado.Ordem"
            If Optfim.Value = True Then .StrSql_Certificado = "Select certificado.* FROM Certificado INNER JOIN Producao ON Certificado.Ordem = Producao.Ordem where Producao.Desenho like '%" & txtTexto & "' order by Certificado.Ordem"
            If optIgual.Value = True Then .StrSql_Certificado = "Select certificado.* FROM Certificado INNER JOIN Producao ON Certificado.Ordem = Producao.Ordem where Producao.Desenho = '" & txtTexto & "' order by Certificado.Ordem"
        End If
        If cmbfiltrarpor = "Código referencia" Then
            If Optinicio.Value = True Then .StrSql_Certificado = "Select certificado.* FROM Certificado INNER JOIN Producao ON Certificado.Ordem = Producao.Ordem where Producao.N_Referencia like '" & txtTexto & "%' order by Certificado.Ordem"
            If Optmeio.Value = True Then .StrSql_Certificado = "Select certificado.* FROM Certificado INNER JOIN Producao ON Certificado.Ordem = Producao.Ordem where Producao.N_Referencia like '%" & txtTexto & "%' order by Certificado.Ordem"
            If Optfim.Value = True Then .StrSql_Certificado = "Select certificado.* FROM Certificado INNER JOIN Producao ON Certificado.Ordem = Producao.Ordem where Producao.N_Referencia like '%" & txtTexto & "' order by Certificado.Ordem"
            If optIgual.Value = True Then .StrSql_Certificado = "Select certificado.* FROM Certificado INNER JOIN Producao ON Certificado.Ordem = Producao.Ordem where Producao.N_Referencia = '" & txtTexto & "' order by Certificado.Ordem"
        End If
        If cmbfiltrarpor = "Descrição" Then
            If Optinicio.Value = True Then .StrSql_Certificado = "Select certificado.* FROM Certificado INNER JOIN Producao ON Certificado.Ordem = Producao.Ordem where Producao.Produto like '" & txtTexto & "%' order by Certificado.Ordem"
            If Optmeio.Value = True Then .StrSql_Certificado = "Select certificado.* FROM Certificado INNER JOIN Producao ON Certificado.Ordem = Producao.Ordem where Producao.Produto like '%" & txtTexto & "%' order by Certificado.Ordem"
            If Optfim.Value = True Then .StrSql_Certificado = "Select certificado.* FROM Certificado INNER JOIN Producao ON Certificado.Ordem = Producao.Ordem where Producao.Produto like '%" & txtTexto & "' order by Certificado.Ordem"
            If optIgual.Value = True Then .StrSql_Certificado = "Select certificado.* FROM Certificado INNER JOIN Producao ON Certificado.Ordem = Producao.Ordem where Producao.Produto = '" & txtTexto & "' order by Certificado.Ordem"
        End If
    Else
        .StrSql_Certificado = "Select * from Certificado order by Ordem"
    End If
    .ProcCarregaLista
End With
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyEscape: Unload Me
    Case vbKeyF2: ProcFiltrar
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 8805, 5, True
cmbfiltrarpor = "Certificado"
Optinicio.Value = True

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
    Case 4: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
