VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form frmNumeroSerieUtilizados 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Numero de série Utilizados"
   ClientHeight    =   6885
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5580
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNumeroSerieUtilizados.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   5580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   405
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   5580
      _ExtentX        =   9843
      _ExtentY        =   714
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
      Icon            =   "frmNumeroSerieUtilizados.frx":000C
      ShowMaximize    =   0   'False
      ShowMinimize    =   0   'False
   End
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   13
      Top             =   6480
      Width           =   5580
      _ExtentX        =   9843
      _ExtentY        =   714
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   645
      Left            =   150
      TabIndex        =   1
      Top             =   5130
      Width           =   5265
      Begin VB.TextBox txtPagIr 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   180
         TabIndex        =   2
         ToolTipText     =   "Número da página."
         Top             =   210
         Width           =   555
      End
      Begin DrawSuite2022.USButton cmdPagProx 
         Height          =   315
         Left            =   2400
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Próxima página."
         Top             =   210
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmNumeroSerieUtilizados.frx":0028
         BorderColor     =   14404026
         BorderColorDown =   11632444
         BorderColorOver =   11632444
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColor2  =   16777215
         GradientColor3  =   16777215
         GradientColorDown2=   16246986
         GradientColorDown3=   15189380
         GradientColorDown4=   14596208
         GradientColorOver1=   16643560
         GradientColorOver2=   16576988
         GradientColorOver3=   16441780
         GradientColorOver4=   16178091
         PicSizeH        =   19
         PicSizeW        =   19
         ShowFocusRect   =   0   'False
         ShowFocusRectDown=   0   'False
      End
      Begin DrawSuite2022.USButton cmdPagAnt 
         Height          =   315
         Left            =   1860
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Página anterior."
         Top             =   210
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmNumeroSerieUtilizados.frx":37CC
         BorderColor     =   14404026
         BorderColorDown =   11632444
         BorderColorOver =   11632444
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColor2  =   16777215
         GradientColor3  =   16777215
         GradientColorDown2=   16246986
         GradientColorDown3=   15189380
         GradientColorDown4=   14596208
         GradientColorOver1=   16643560
         GradientColorOver2=   16576988
         GradientColorOver3=   16441780
         GradientColorOver4=   16178091
         PicSizeH        =   19
         PicSizeW        =   19
         ShowFocusRect   =   0   'False
         ShowFocusRectDown=   0   'False
      End
      Begin DrawSuite2022.USButton cmdPagIr 
         Height          =   315
         Left            =   750
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   210
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   556
         BorderColor     =   14404026
         BorderColorDown =   11632444
         BorderColorOver =   11632444
         Caption         =   "Ir"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColor2  =   16777215
         GradientColor3  =   16777215
         GradientColorDown2=   16246986
         GradientColorDown3=   15189380
         GradientColorDown4=   14596208
         GradientColorOver1=   16643560
         GradientColorOver2=   16576988
         GradientColorOver3=   16441780
         GradientColorOver4=   16178091
         PicSizeH        =   19
         PicSizeW        =   19
         ShowFocusRect   =   0   'False
         ShowFocusRectDown=   0   'False
      End
      Begin DrawSuite2022.USButton cmdPagPrim 
         Height          =   315
         Left            =   1320
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Primeira página."
         Top             =   210
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmNumeroSerieUtilizados.frx":72D5
         BorderColor     =   14404026
         BorderColorDown =   11632444
         BorderColorOver =   11632444
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColor2  =   16777215
         GradientColor3  =   16777215
         GradientColorDown2=   16246986
         GradientColorDown3=   15189380
         GradientColorDown4=   14596208
         GradientColorOver1=   16643560
         GradientColorOver2=   16576988
         GradientColorOver3=   16441780
         GradientColorOver4=   16178091
         PicSizeH        =   19
         PicSizeW        =   19
         ShowFocusRect   =   0   'False
         ShowFocusRectDown=   0   'False
      End
      Begin DrawSuite2022.USButton cmdPagUlt 
         Height          =   315
         Left            =   2940
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Última página."
         Top             =   210
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmNumeroSerieUtilizados.frx":B3C4
         BorderColor     =   14404026
         BorderColorDown =   11632444
         BorderColorOver =   11632444
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColor2  =   16777215
         GradientColor3  =   16777215
         GradientColorDown2=   16246986
         GradientColorDown3=   15189380
         GradientColorDown4=   14596208
         GradientColorOver1=   16643560
         GradientColorOver2=   16576988
         GradientColorOver3=   16441780
         GradientColorOver4=   16178091
         PicSizeH        =   19
         PicSizeW        =   19
         ShowFocusRect   =   0   'False
         ShowFocusRectDown=   0   'False
      End
      Begin VB.Label lblPaginas 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Página: 0 de: 0"
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
         Left            =   3630
         TabIndex        =   8
         Top             =   300
         Width           =   1245
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   705
      Left            =   150
      TabIndex        =   9
      Top             =   5670
      Width           =   5265
      Begin VB.Label TTProd 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Produzido :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   3570
         TabIndex        =   12
         Top             =   330
         Width           =   1215
      End
      Begin VB.Label TTNC 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total NC : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   1980
         TabIndex        =   11
         Top             =   330
         Width           =   765
      End
      Begin VB.Label TTOK 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Aprovado : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   90
         TabIndex        =   10
         Top             =   330
         Width           =   1260
      End
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   4485
      Left            =   120
      TabIndex        =   0
      Top             =   630
      Width           =   5310
      _ExtentX        =   9366
      _ExtentY        =   7911
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Text            =   "Item"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Object.Tag             =   "N"
         Text            =   "Número de série"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Status"
         Object.Width           =   4233
      EndProperty
   End
End
Attribute VB_Name = "frmNumeroSerieUtilizados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub ProcExibePagina(Pagina As Integer)
On Error GoTo tratar_erro

With Lista
.ListItems.Clear

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from ProducaoFases_Codigos where IDProducao = '" & frmSituacao_Producao.Lista.SelectedItem & "' and IDProducao IS NOT NULL", Conexao, adOpenKeyset, adLockOptimistic
  If TBAbrir.EOF = False Then
  
TBAbrir.PageSize = 20
TBAbrir.AbsolutePage = Pagina
TamanhoPagina = TBAbrir.PageSize
ContadorReg = 1
  

    Contador2 = 1
        Do While TBAbrir.EOF = False And (ContadorReg <= TamanhoPagina)
            .ListItems.Add , , Contador2
            .ListItems.Item(Contador2).SubItems(1) = TBAbrir!CODIGO
            .ListItems.Item(Contador2).SubItems(2) = TBAbrir!status
            Contador2 = Contador2 + 1
            ContadorReg = ContadorReg + 1
            TBAbrir.MoveNext
        Loop
        

    If TBAbrir.AbsolutePage = adPosBOF Then
       lblPaginas.Caption = "Página: 1 de: " & TBAbrir.PageCount
    ElseIf TBAbrir.AbsolutePage = adPosEOF Then
            lblPaginas.Caption = "Página: " & TBAbrir.PageCount & " de: " & TBAbrir.PageCount
        Else
            lblPaginas.Caption = "Página: " & TBAbrir.AbsolutePage - 1 & " de: " & TBAbrir.PageCount
    End If
        
  End If
End With


Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcExibePagina 1
TotalOK = 0
TotalNC = 0

Set TBAcessos = CreateObject("adodb.recordset")
TBAcessos.Open "Select Count(Status) AS TTOK from ProducaoFases_Codigos where IDProducao = '" & frmSituacao_Producao.Lista.SelectedItem & "' and Status = 'APROVADO'", Conexao, adOpenKeyset, adLockOptimistic
  If TBAcessos.EOF = False Then
        TTOK.Caption = "Total aprovado: " & TBAcessos!TTOK
        TotalOK = TBAcessos!TTOK
        Else
         TTOK.Caption = "Total aprovado: 0"
    End If
TBAcessos.Close

Set TBAcessos = CreateObject("adodb.recordset")
TBAcessos.Open "Select Count(Status) AS TTNC from ProducaoFases_Codigos where IDProducao = '" & frmSituacao_Producao.Lista.SelectedItem & "' and Status = 'NÃO CONFORME'", Conexao, adOpenKeyset, adLockOptimistic
  If TBAcessos.EOF = False Then
        TTNC.Caption = "Total rejeitado: " & TBAcessos!TTNC
        TotalNC = TBAcessos!TTNC
        Else
        TTNC.Caption = "Total rejeitado: 0"
    End If
TBAcessos.Close

Totalprod = TotalOK + TotalNC
TTProd.Caption = "Total prod: " & Totalprod

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyEscape: Unload Me
End Select

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub cmdPagAnt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBAbrir.AbsolutePage <> 2 Then
    If TBAbrir.AbsolutePage = -3 Then
        ProcExibePagina (TBAbrir.PageCount - 1)
    Else
        TBAbrir.AbsolutePage = TBAbrir.AbsolutePage - 2
        ProcExibePagina (TBAbrir.AbsolutePage)
    End If
Else
    ProcExibePagina (1)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagIr_Click()
On Error GoTo tratar_erro

If txtPagIr = "" Then Exit Sub
Quant = ReturnNumbersOnly(Right(lblPaginas.Caption, 4))
If Quant <= 1 Or txtPagIr > Quant Then Exit Sub
If txtPagIr.Text >= 1 And txtPagIr.Text <= Quant Then
    TBAbrir.AbsolutePage = txtPagIr.Text
    ProcExibePagina (TBAbrir.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBAbrir.AbsolutePage = 1
ProcExibePagina (TBAbrir.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBAbrir.AbsolutePage <> -3 Then
    If TBAbrir.AbsolutePage = 1 Then
        ProcExibePagina (2)
    Else
        ProcExibePagina (TBAbrir.AbsolutePage)
    End If
Else
    ProcExibePagina (TBAbrir.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBAbrir.AbsolutePage = TBAbrir.PageCount
ProcExibePagina (TBAbrir.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
