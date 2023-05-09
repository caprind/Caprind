VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmEstoque_Retirar_NumeroSerie 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Caprind - Baixa de estoque por numero de série"
   ClientHeight    =   5295
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7005
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEstoque_Retirar_NumeroSerie.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5295
   ScaleWidth      =   7005
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChkRegistro 
      BackColor       =   &H00000000&
      Caption         =   "(F1) Registro automático"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1275
      Left            =   5130
      MaskColor       =   &H00004000&
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   600
      Width           =   1725
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Anexado a nota fiscal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1245
      Left            =   2970
      TabIndex        =   7
      Top             =   600
      Width           =   2115
      Begin VB.TextBox txtAnexado 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   555
         Left            =   180
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   "0"
         ToolTipText     =   "Quantidade conforme."
         Top             =   480
         Width           =   1680
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Informe o numero de série"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1215
      Left            =   240
      TabIndex        =   6
      Top             =   630
      Width           =   2655
      Begin VB.TextBox txtNSerie 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   180
         MaxLength       =   12
         TabIndex        =   0
         ToolTipText     =   "Digite o numero de série do item."
         Top             =   420
         Width           =   2265
      End
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   465
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   7005
      _ExtentX        =   12356
      _ExtentY        =   741
      DibPicture      =   "frmEstoque_Retirar_NumeroSerie.frx":0E42
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
      Icon            =   "frmEstoque_Retirar_NumeroSerie.frx":AAF3
      ShowMaximize    =   0   'False
      ShowMinimize    =   0   'False
   End
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   4
      Top             =   4890
      Width           =   7005
      _ExtentX        =   12356
      _ExtentY        =   714
   End
   Begin DrawSuite2022.USButton cmdExcluir 
      Height          =   1215
      Left            =   3510
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Excluir numero de serie do item"
      Top             =   3540
      Visible         =   0   'False
      Width           =   3285
      _ExtentX        =   5794
      _ExtentY        =   2143
      BorderColor     =   5263559
      BorderColorDisabled=   13160660
      BorderColorDown =   4013465
      BorderColorOver =   4408288
      Caption         =   "(F4) Excluir número de série"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
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
      PicAlign        =   8
      ShowFocusRect   =   0   'False
      ShowFocusRectDown=   0   'False
      Theme           =   4
      ToolTipTitle    =   "GERPROD"
   End
   Begin DrawSuite2022.USButton btnDisponiveis 
      Height          =   1455
      Left            =   240
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Gravar evento."
      Top             =   1920
      Width           =   3225
      _ExtentX        =   5689
      _ExtentY        =   2566
      DibPicture      =   "frmEstoque_Retirar_NumeroSerie.frx":AE0D
      BorderColorDown =   15048022
      BorderColorOver =   15381630
      Caption         =   "(F3) Disponíveis"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicAlign        =   8
      PicSize         =   5
      PicSizeH        =   50
      PicSizeW        =   50
      ShowFocusRect   =   0   'False
      ShowFocusRectDown=   0   'False
      ToolTipTitle    =   "GERPROD"
   End
   Begin DrawSuite2022.USButton btnUtilizados 
      Height          =   1455
      Left            =   3510
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Gravar evento."
      Top             =   1920
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   2566
      DibPicture      =   "frmEstoque_Retirar_NumeroSerie.frx":14FBA
      BorderColorDown =   15048022
      BorderColorOver =   15381630
      Caption         =   "(F6) Utilizados"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicAlign        =   8
      PicSize         =   5
      PicSizeH        =   50
      PicSizeW        =   50
      ShowFocusRect   =   0   'False
      ShowFocusRectDown=   0   'False
      ToolTipTitle    =   "GERPROD"
   End
   Begin DrawSuite2022.USButton cmdGravar 
      Height          =   1215
      Left            =   240
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   "Excluir numero de serie do item"
      Top             =   3540
      Visible         =   0   'False
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   2143
      BorderColor     =   4960354
      BorderColorDisabled=   13160660
      BorderColorDown =   4210752
      BorderColorOver =   49152
      Caption         =   "(F2) Anexar número de série"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
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
      PicAlign        =   8
      ShowFocusRect   =   0   'False
      ShowFocusRectDown=   0   'False
      Theme           =   3
      ToolTipTitle    =   "GERPROD"
   End
End
Attribute VB_Name = "frmEstoque_Retirar_NumeroSerie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnDisponiveis_Click()
On Error GoTo tratar_erro

    frmEstoque_Retirar_NumeroSerieDisponiveis.Show 1

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub btnUtilizados_Click()
On Error GoTo tratar_erro

    frmEstoque_Retirar_NumeroSerieUtilizados.Show 1

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ChkRegistro_Click()
On Error GoTo tratar_erro

If ChkRegistro.Value = 0 Then
ChkRegistro.BackColor = &H0
ChkRegistro.ForeColor = &HFFFFFF
ChkRegistro.Caption = "(F1) Registro automático"
cmdGravar.Visible = False
cmdExcluir.Visible = False
frmEstoque_Retirar_NumeroSerie.Height = 3875
Else
ChkRegistro.BackColor = &H119CF3
ChkRegistro.ForeColor = &H80000012
ChkRegistro.Caption = "(F1) Registro manual"
cmdGravar.Visible = True
cmdExcluir.Visible = True
frmEstoque_Retirar_NumeroSerie.Height = 5295
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcAnexar()
On Error GoTo tratar_erro

If txtNSerie.Text = "" Then
    USMsgBox "É obrigatorio informar o numero de série do item antes de gravar", vbCritical, "CAPRIND v5.0"
    Exit Sub
End If

Contador = frmestoque_Retirar.txtRetirar

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Producao_Etiquetas where N_Serie = '" & txtNSerie.Text & "'", Conexao, adOpenKeyset, adLockOptimistic

If TBAbrir.EOF = False Then
TBAbrir!ID_nota = ID_nota
TBAbrir!ID_produtonota = ID_produto_nota
TBAbrir.Update
TBAbrir.Close
txtNSerie.Text = ""
End If

ProcAtualizaTotaisAnexado

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcExcluir()
On Error GoTo tratar_erro


If txtNSerie.Text <> "" Then

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Producao_Etiquetas where N_serie = '" & txtNSerie & "' and ID_ProdutoNota is not null", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = True Then
USMsgBox "Numero de série " & txtNSerie.Text & " inexistente nos apontamentos!", vbInformation, "CAPRIND v5.0"
Exit Sub
End If

If USMsgBox("Deseja realmente excluir o numero de série " & txtNSerie.Text & " ?", vbYesNo, "CAPRIND v5.0") = vbYes Then
NumeroSerie = txtNSerie.Text
Conexao.Execute "update Producao_etiquetas SET ID_Nota = NULL , ID_ProdutoNota = NULL where N_serie = '" & txtNSerie & "'"
USMsgBox "Numero de série " & txtNSerie.Text & " excluido com sucesso !", vbInformation, "CAPRIND v5.0"
txtNSerie.Text = ""
End If

ProcAtualizaTotaisAnexado

End If


Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub cmdExcluir_Click()
On Error GoTo tratar_erro

ProcExcluir

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub cmdGravar_Click()
On Error GoTo tratar_erro

    ProcAnexar
    
Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyF1:
    
    If ChkRegistro.Value = 0 Then
        ChkRegistro.Value = 1
    Else
        ChkRegistro.Value = 0
    End If
    
    Case vbKeyF2: ProcAnexar
    Case vbKeyF3: frmNumeroSerieDisponiveis.Show 1
    Case vbKeyF4: ProcExcluir
    Case vbKeyF6: frmNumeroSerieUtilizados.Show 1
    Case vbKeyEscape: Unload Me
End Select

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcAtualizaTotaisAnexado()
On Error GoTo tratar_erro
  
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select Count(ID_Nota) AS TotalAnexado from Producao_Etiquetas where ID_Nota = '" & ID_nota & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        txtAnexado = TBAbrir!TotalAnexado
    End If
    TBAbrir.Close

    
Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub


Private Sub Form_Load()
On Error GoTo tratar_erro

    ChkRegistro_Click
    ProcAtualizaTotaisAnexado

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtNSerie_Change()
On Error GoTo tratar_erro

If Len(txtNSerie.Text) = 8 Then

If txtNSerie.Text = "" Then
    USMsgBox "É obrigatorio informar o numero de série do item antes de gravar", vbCritical, "CAPRIND v5.0"
    Exit Sub
End If


status = "APROVADO"
    
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from Producao_rastreavel where N_serie = '" & txtNSerie.Text & "' and Status = 'NÃO CONFORME'", Conexao, adOpenKeyset, adLockOptimistic

    If TBAbrir.EOF = False And ChkRegistro.Value = 0 Then
        USMsgBox "Numero de série informado está com Status Não conforme." & vbCrLf & "Informe um numero de série válido", vbCritical, "CAPRIND v5.0"
        txtNSerie.Text = ""
        txtNSerie.SetFocus
        Exit Sub
    End If
    
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from Producao_rastreavel where N_serie = '" & txtNSerie.Text & "' and Status <> 'NÃO CONFORME'", Conexao, adOpenKeyset, adLockOptimistic

    If TBAbrir.EOF = True Then
        USMsgBox "Numero de série informado não existe." & vbCrLf & "Informe um numero de série válido", vbCritical, "CAPRIND v5.0"
        txtNSerie.Text = ""
        txtNSerie.SetFocus
        Exit Sub
    End If
     
  
    If ChkRegistro.Value = 0 Then
        ProcAnexar
    End If
    
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtAnexado_Change()
On Error GoTo tratar_erro

If IsNumeric(txtAnexado) = True Then
    TotalAnexado = txtAnexado.Text
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
