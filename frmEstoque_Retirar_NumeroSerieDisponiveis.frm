VERSION 5.00
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEstoque_Retirar_NumeroSerieDisponiveis 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Numero de s�rie dispon�veis"
   ClientHeight    =   6300
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
   Icon            =   "frmEstoque_Retirar_NumeroSerieDisponiveis.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   5580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   10
      Top             =   5895
      Width           =   5580
      _ExtentX        =   9843
      _ExtentY        =   714
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   435
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   5580
      _ExtentX        =   9843
      _ExtentY        =   767
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
      Icon            =   "frmEstoque_Retirar_NumeroSerieDisponiveis.frx":000C
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   645
      Left            =   150
      TabIndex        =   1
      Top             =   5070
      Width           =   5265
      Begin VB.TextBox txtPagIr 
         Alignment       =   2  'Center
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
         ToolTipText     =   "N�mero da p�gina."
         Top             =   210
         Width           =   555
      End
      Begin DrawSuite2022.USButton cmdPagProx 
         Height          =   315
         Left            =   2400
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Pr�xima p�gina."
         Top             =   210
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmEstoque_Retirar_NumeroSerieDisponiveis.frx":0028
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
         BorderColor     =   14404026
         BorderColorDown =   11632444
         BorderColorOver =   11632444
         GradientColor2  =   16777215
         GradientColor3  =   16777215
         GradientColorOver1=   16643560
         GradientColorOver2=   16576988
         GradientColorOver3=   16441780
         GradientColorOver4=   16178091
         GradientColorDown2=   16246986
         GradientColorDown3=   15189380
         GradientColorDown4=   14596208
         PicSizeH        =   19
         PicSizeW        =   19
         ShowFocusRect   =   0   'False
      End
      Begin DrawSuite2022.USButton cmdPagAnt 
         Height          =   315
         Left            =   1860
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "P�gina anterior."
         Top             =   210
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmEstoque_Retirar_NumeroSerieDisponiveis.frx":37CC
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
         BorderColor     =   14404026
         BorderColorDown =   11632444
         BorderColorOver =   11632444
         GradientColor2  =   16777215
         GradientColor3  =   16777215
         GradientColorOver1=   16643560
         GradientColorOver2=   16576988
         GradientColorOver3=   16441780
         GradientColorOver4=   16178091
         GradientColorDown2=   16246986
         GradientColorDown3=   15189380
         GradientColorDown4=   14596208
         PicSizeH        =   19
         PicSizeW        =   19
         ShowFocusRect   =   0   'False
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
         BorderColor     =   14404026
         BorderColorDown =   11632444
         BorderColorOver =   11632444
         GradientColor2  =   16777215
         GradientColor3  =   16777215
         GradientColorOver1=   16643560
         GradientColorOver2=   16576988
         GradientColorOver3=   16441780
         GradientColorOver4=   16178091
         GradientColorDown2=   16246986
         GradientColorDown3=   15189380
         GradientColorDown4=   14596208
         PicSizeH        =   19
         PicSizeW        =   19
         ShowFocusRect   =   0   'False
      End
      Begin DrawSuite2022.USButton cmdPagPrim 
         Height          =   315
         Left            =   1320
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Primeira p�gina."
         Top             =   210
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmEstoque_Retirar_NumeroSerieDisponiveis.frx":72D5
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
         BorderColor     =   14404026
         BorderColorDown =   11632444
         BorderColorOver =   11632444
         GradientColor2  =   16777215
         GradientColor3  =   16777215
         GradientColorOver1=   16643560
         GradientColorOver2=   16576988
         GradientColorOver3=   16441780
         GradientColorOver4=   16178091
         GradientColorDown2=   16246986
         GradientColorDown3=   15189380
         GradientColorDown4=   14596208
         PicSizeH        =   19
         PicSizeW        =   19
         ShowFocusRect   =   0   'False
      End
      Begin DrawSuite2022.USButton cmdPagUlt 
         Height          =   315
         Left            =   2940
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "�ltima p�gina."
         Top             =   210
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmEstoque_Retirar_NumeroSerieDisponiveis.frx":B3C4
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
         BorderColor     =   14404026
         BorderColorDown =   11632444
         BorderColorOver =   11632444
         GradientColor2  =   16777215
         GradientColor3  =   16777215
         GradientColorOver1=   16643560
         GradientColorOver2=   16576988
         GradientColorOver3=   16441780
         GradientColorOver4=   16178091
         GradientColorDown2=   16246986
         GradientColorDown3=   15189380
         GradientColorDown4=   14596208
         PicSizeH        =   19
         PicSizeW        =   19
         ShowFocusRect   =   0   'False
      End
      Begin VB.Label lblPaginas 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "P�gina: 0 de: 0"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3690
         TabIndex        =   8
         Top             =   270
         Width           =   1545
      End
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   4485
      Left            =   120
      TabIndex        =   0
      Top             =   540
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
         Text            =   "N�mero de s�rie"
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
Attribute VB_Name = "frmEstoque_Retirar_NumeroSerieDisponiveis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cmd_F3_Click()
On Error GoTo tratar_erro

If USMsgBox("Deseja realmente gravar esses dados informados?", vbYesNo, "CAPRIND v5.0") = vbNo Then
 Exit Sub
End If

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from CQ_NC_FABRICA_Serie where Codigo = '" & CODIGO & "'", Conexao, adOpenKeyset, adLockOptimistic
Contador = 1
With GridSerie
Linha = .rows - 1
    For InitFor = 1 To Linha
      If .Cell(Contador, 2).Text = "" Then
      USMsgBox "Informe o numero de serie", vbCritical, "GERPROD | COLETOR DE DADOS"
      .Cell(Contador, 2).SetFocus
      TBAbrir.Close
      Exit Sub
      End If
        NumeroSerie = .Cell(Contador, 2).Text
        If TBAbrir.EOF = True Then
            TBAbrir.AddNew
        End If
        TBAbrir!CODIGO = CODIGO
        TBAbrir!NumeroSerie = NumeroSerie
        TBAbrir!IDProducao = IDProducao
        TBAbrir.Update
        Contador = Contador + 1
        Linha = Linha - 1
        TBAbrir.MoveNext
    Next InitFor
End With

USMsgBox "Dados gravados com sucesso", vbInformation, "CAPRIND v5.0"
Unload Me

TBAbrir.Close

Exit Sub
tratar_erro:
    MsgBox ("Descri��o do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcExibePagina(Pagina As Integer)
On Error GoTo tratar_erro

With Lista
.ListItems.Clear

StrSql = "Select * from Producao_Etiquetas where Ordem = '" & Ordem & "' AND ID_ProdutoNota is null ORDER BY N_serie"

Set TBAbrir = CreateObject("adodb.recordset")
''Debug.print StrSQL

TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic


If TBAbrir.EOF = False Then
  
TBAbrir.PageSize = 20
TBAbrir.AbsolutePage = Pagina
TamanhoPagina = TBAbrir.PageSize
ContadorReg = 1

    Contador2 = 1
    
        Do While TBAbrir.EOF = False And (ContadorReg <= TamanhoPagina)
        'Verifica se est� provado
        StrSql = "Select N_Serie, Status from Producao_rastreavel where N_Serie = '" & TBAbrir!N_Serie & "'"
        Set TBEtiqueta = CreateObject("adodb.recordset")
        TBEtiqueta.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
            If TBEtiqueta.EOF = False Then
                    .ListItems.Add , , Contador2
                    .ListItems.Item(Contador2).SubItems(1) = TBAbrir!N_Serie
                    .ListItems.Item(Contador2).SubItems(2) = TBEtiqueta!status
                Contador2 = Contador2 + 1
                ContadorReg = ContadorReg + 1
            End If
            TBEtiqueta.Close
             TBAbrir.MoveNext
        Loop
        
    If TBAbrir.AbsolutePage = adPosBOF Then
       lblPaginas.Caption = "P�gina: 1 de: " & TBAbrir.PageCount
    ElseIf TBAbrir.AbsolutePage = adPosEOF Then
            lblPaginas.Caption = "P�gina: " & TBAbrir.PageCount & " de: " & TBAbrir.PageCount
        Else
            lblPaginas.Caption = "P�gina: " & TBAbrir.AbsolutePage - 1 & " de: " & TBAbrir.PageCount
    End If
        
  End If


End With
'TBAbrir.Close

Exit Sub
tratar_erro:
    MsgBox ("Descri��o do erro : " + Error()), vbCritical
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcExibePagina 1

Exit Sub
tratar_erro:
    USMsgBox ("Descri��o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyF3: Cmd_F3_Click
    Case vbKeyEscape: Unload Me
    Case vbKeyReturn:
    If Lista.ListItems.Count > 0 Then
        If USMsgBox("Deseja utilizar o numero de s�rie " & Lista.SelectedItem.ListSubItems.Item(1).Text & " no apontamento?", vbYesNo, "GERPROD") = vbYes Then
            frmNumeroSerieOK.txtNSerie = Lista.SelectedItem.ListSubItems.Item(1).Text
            ProcExibePagina 1
        End If
    End If
    
End Select

Exit Sub
tratar_erro:
    MsgBox ("Descri��o do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Lista_DblClick()
On Error GoTo tratar_erro

If Lista.ListItems.Count > 0 Then
    frmEstoque_Retirar_NumeroSerie.txtNSerie = Lista.SelectedItem.ListSubItems.Item(1).Text
    Unload Me
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descri��o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
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
    USMsgBox ("Descri��o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
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
    USMsgBox ("Descri��o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBAbrir.AbsolutePage = 1
ProcExibePagina (TBAbrir.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descri��o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
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
    USMsgBox ("Descri��o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub cmdPagUlt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBAbrir.AbsolutePage = TBAbrir.PageCount
ProcExibePagina (TBAbrir.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descri��o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub



