VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2014.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form frmSugestoes 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Processos | Sugestões de melhorias"
   ClientHeight    =   7680
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6540
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7680
   ScaleWidth      =   6540
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2014.USButton btnNovo 
      Height          =   405
      Left            =   240
      TabIndex        =   10
      ToolTipText     =   "Criar nova sugestão"
      Top             =   570
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   714
      DibPicture      =   "frmSugestoes.frx":0000
      BorderColor     =   1154291
      BorderColorDisabled=   13160660
      BorderColorDown =   16576
      BorderColorOver =   8438015
      Caption         =   "Novo (Insert)"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
      GradientColor1  =   1154291
      GradientColor2  =   1154291
      GradientColor3  =   1154291
      GradientColor4  =   1154291
      GradientColorDisabled1=   14215660
      GradientColorDisabled2=   14215660
      GradientColorDisabled3=   14215660
      GradientColorDisabled4=   14215660
      GradientColorDown1=   16576
      GradientColorDown2=   16576
      GradientColorDown3=   16576
      GradientColorDown4=   16576
      GradientColorOver1=   8438015
      GradientColorOver2=   8438015
      GradientColorOver3=   8438015
      GradientColorOver4=   8438015
      ShowFocusRect   =   0   'False
      ShowFocusRectDown=   0   'False
      Theme           =   5
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Dados da sugestão de melhoria da fase do processo"
      Height          =   2655
      Left            =   210
      TabIndex        =   3
      Top             =   1170
      Width           =   6135
      Begin VB.TextBox txtSugestao 
         Height          =   1185
         Left            =   210
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   0
         Top             =   1260
         Width           =   5685
      End
      Begin VB.TextBox txtResponsavel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   570
         Width           =   4575
      End
      Begin VB.TextBox txtData 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   315
         Left            =   210
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   570
         Width           =   1095
      End
      Begin VB.TextBox txtID 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   315
         Left            =   210
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   570
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sugestão para melhoria da fase do processo"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   1455
         TabIndex        =   8
         Top             =   1050
         Width           =   3195
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Responsável"
         Height          =   195
         Left            =   3150
         TabIndex        =   5
         Top             =   360
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data"
         Height          =   195
         Left            =   585
         TabIndex        =   4
         Top             =   360
         Width           =   345
      End
   End
   Begin DrawSuite2014.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   2
      Top             =   7275
      Width           =   6540
      _ExtentX        =   11536
      _ExtentY        =   714
   End
   Begin DrawSuite2014.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   465
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6540
      _ExtentX        =   11536
      _ExtentY        =   820
      DibPicture      =   "frmSugestoes.frx":09CE
      CaptionDelimiter=   "|"
      CaptionOnCenter =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Icon            =   "frmSugestoes.frx":1647
      ShowMaximize    =   0   'False
      ShowMinimize    =   0   'False
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   3285
      Left            =   210
      TabIndex        =   9
      Top             =   3870
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   5794
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Data"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Responsável"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Sugestão"
         Object.Width           =   12348
      EndProperty
   End
   Begin DrawSuite2014.USButton btnSalvar 
      Height          =   405
      Left            =   1650
      TabIndex        =   11
      ToolTipText     =   "Salvar dados da sugestão"
      Top             =   570
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   714
      DibPicture      =   "frmSugestoes.frx":1961
      BorderColor     =   4960354
      BorderColorDisabled=   13160660
      BorderColorDown =   4210752
      BorderColorOver =   49152
      Caption         =   "Salvar (F3)"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
      GradientColor1  =   4960354
      GradientColor2  =   4960354
      GradientColor3  =   4960354
      GradientColor4  =   4960354
      GradientColorDisabled1=   14215660
      GradientColorDisabled2=   14215660
      GradientColorDisabled3=   14215660
      GradientColorDisabled4=   14215660
      GradientColorDown1=   32768
      GradientColorDown2=   32768
      GradientColorDown3=   32768
      GradientColorDown4=   32768
      GradientColorOver1=   49152
      GradientColorOver2=   49152
      GradientColorOver3=   49152
      GradientColorOver4=   49152
      ShowFocusRect   =   0   'False
      ShowFocusRectDown=   0   'False
      Theme           =   3
   End
   Begin DrawSuite2014.USButton btnExcluir 
      Height          =   405
      Left            =   3060
      TabIndex        =   12
      ToolTipText     =   "Excluir sugestão"
      Top             =   570
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   714
      DibPicture      =   "frmSugestoes.frx":20F3
      BorderColor     =   5263559
      BorderColorDisabled=   13160660
      BorderColorDown =   4013465
      BorderColorOver =   4408288
      Caption         =   "Excluir (F4)"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
      ShowFocusRect   =   0   'False
      ShowFocusRectDown=   0   'False
      Theme           =   4
   End
End
Attribute VB_Name = "frmSugestoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ProcApagar()
On Error GoTo tratar_erro

If txtID.Text = "" Then
    USMsgBox "Escolha uma sugestão para excluir!", vbInformation, "CAPRIND v5.0"
    Exit Sub
End If

If txtResponsavel.Text <> Operador Then
    USMsgBox "Usuario não autorizado a excluir essa sugestão!", vbInformation, "CAPRIND v5.0"
    Exit Sub
End If

If USMsgBox("Deseja realmente excluir essa sugestão?", vbYesNo, "CAPRIND v5.0") = vbNo Then
Exit Sub
End If


StrSql = "Delete from Fases_Sugestao where ID = '" & txtID.Text & "'"

Conexao.Execute StrSql
USMsgBox "Sugestão excluida com sucesso!", vbInformation, "CAPRIND v5.0"

ProcNovo
ProcCarregaLista

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub btnExcluir_Click()
On Error GoTo tratar_erro

ProcApagar

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub btnNovo_Click()
On Error GoTo tratar_erro

ProcNovo

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcNovo()
On Error GoTo tratar_erro

txtID.Text = 0
txtData.Text = Date
txtResponsavel.Text = Operador
txtSugestao.Text = ""
txtSugestao.SetFocus

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Lista_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

ProcPuxaDados

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcPuxaDados()
On Error GoTo tratar_erro

If Lista.ListItems.Count > 0 Then
    txtID.Text = Lista.SelectedItem
    txtData.Text = Lista.SelectedItem.ListSubItems(1)
    txtResponsavel.Text = Lista.SelectedItem.ListSubItems(2)
    txtSugestao.Text = Lista.SelectedItem.ListSubItems(3)
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcSalvar()
On Error GoTo tratar_erro

If txtSugestao.Text = "" Then
    USMsgBox "Digite sua sugestão para salvar!", vbInformation, "CAPRIND v5.0"
    Exit Sub
End If

If txtResponsavel.Text <> Operador Then
    USMsgBox "Usuario não autorizado a modificar essa sugestão!", vbInformation, "CAPRIND v5.0"
    Exit Sub
End If


    Set TBproducao = CreateObject("adodb.recordset")
    TBproducao.Open "Select * from Fases_Sugestao where ID = " & txtID.Text, Conexao, adOpenKeyset, adLockOptimistic
    If TBproducao.EOF = True Then
    TBproducao.AddNew
    TBproducao!status = 1
    End If
    
    
    TBproducao!IDFase = IDFase
    TBproducao!Data = txtData.Text
    TBproducao!Responsavel = txtResponsavel.Text
    TBproducao!Sugestao = txtSugestao.Text
    TBproducao.Update
    TBproducao.Close
    USMsgBox "Dados da sugestão gravados com sucesso!", vbInformation, "CAPRIND v5.0"
    ProcCarregaLista
    ProcNovo
    

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcCarregaLista()
On Error GoTo tratar_erro

Lista.ListItems.Clear

Set TBLISTA = CreateObject("adodb.recordset")

StrSql = "Select * from Fases_Sugestao where IDFase = '" & IDFase & "'"

TBLISTA.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    
    Do While TBLISTA.EOF = False
        With Lista.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = Format(TBLISTA!Data, "dd/mm/yy")
            .Item(.Count).SubItems(2) = TBLISTA!Responsavel
            .Item(.Count).SubItems(3) = TBLISTA!Sugestao
        End With
        TBLISTA.MoveNext
    Loop
  TBLISTA.Close
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub btnSalvar_Click()
On Error GoTo tratar_erro

ProcSalvar
ProcNovo
ProcCarregaLista

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyInsert: ProcNovo
    Case vbKeyF3: ProcSalvar
    Case vbKeyF4: ProcApagar
    Case vbKeyEscape:
        Unload Me
End Select

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

OS = frmProducao.Txt_OS.Text

Set TBFase = CreateObject("adodb.recordset")
TBFase.Open "Select IDFase from OrdemServico where IDproducao = " & OS, Conexao, adOpenKeyset, adLockOptimistic
If TBFase.EOF = False Then
IDFase = TBFase!IDFase
End If
TBFase.Close

ProcCarregaLista

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub
