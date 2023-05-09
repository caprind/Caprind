VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmGermaqfer_bloqTurno 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "PCP - Postos de trabalho - Bloquear/Desbloquear"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6450
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   6450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optTurno 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Por turno:"
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
      Left            =   1350
      TabIndex        =   1
      Top             =   1185
      Width           =   1035
   End
   Begin VB.OptionButton Opt_registro 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Registro"
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
      Left            =   255
      TabIndex        =   0
      Top             =   1185
      Value           =   -1  'True
      Width           =   915
   End
   Begin VB.ComboBox cmbturno 
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
      Height          =   330
      ItemData        =   "frmGermaqfer_bloqTurno.frx":0000
      Left            =   2415
      List            =   "frmGermaqfer_bloqTurno.frx":0010
      Locked          =   -1  'True
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Turno."
      Top             =   1140
      Width           =   885
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   3015
      Left            =   55
      TabIndex        =   6
      Top             =   1440
      Width           =   6330
      Begin VB.TextBox txtStatus 
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
         Left            =   1290
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   180
         Width           =   4815
      End
      Begin VB.TextBox txtObservacoes 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   1995
         Left            =   1290
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Observações."
         Top             =   900
         Width           =   4815
      End
      Begin VB.TextBox txtResponsavel 
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
         Left            =   1290
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   540
         Width           =   4815
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Status :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   9
         Top             =   180
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Observações :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   150
         TabIndex        =   8
         Top             =   900
         Width           =   1065
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Responsável :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   7
         Top             =   540
         Width           =   1035
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   55
      TabIndex        =   10
      Top             =   30
      Width           =   6330
      _ExtentX        =   11165
      _ExtentY        =   1720
      ButtonCount     =   7
      GradientColor2  =   14737632
      GradientColorOverRight1=   16315633
      GradientColorOverRight2=   15195350
      GripperColor    =   15195350
      IsStrech        =   -1  'True
      RightColor1     =   0
      RightColor2     =   0
      ShowEndPanel    =   0   'False
      Theme           =   1
      ButtonCaption1  =   "Salvar"
      ButtonEnabled1  =   0   'False
      ButtonIconSize1 =   32
      ButtonToolTipText1=   "Salvar (F3)"
      ButtonKey1      =   "3"
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
      ButtonWidth1    =   38
      ButtonHeight1   =   21
      ButtonUseMaskColor1=   0   'False
      ButtonCaption2  =   "Bloquear"
      ButtonEnabled2  =   0   'False
      ButtonToolTipText2=   "Bloquear (F6)"
      ButtonKey2      =   "2"
      ButtonAlignment2=   2
      BeginProperty ButtonFont2 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft2     =   42
      ButtonTop2      =   2
      ButtonWidth2    =   50
      ButtonHeight2   =   21
      ButtonUseMaskColor2=   0   'False
      ButtonCaption3  =   "Desbloquear"
      ButtonEnabled3  =   0   'False
      ButtonIconSize3 =   32
      ButtonToolTipText3=   "Desbloquear (F7)"
      ButtonKey3      =   "3"
      ButtonAlignment3=   2
      BeginProperty ButtonFont3 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft3     =   94
      ButtonTop3      =   2
      ButtonWidth3    =   68
      ButtonHeight3   =   21
      ButtonUseMaskColor3=   0   'False
      ButtonEnabled4  =   0   'False
      ButtonIconSize4 =   32
      ButtonAlignment4=   2
      ButtonType4     =   1
      ButtonStyle4    =   -1
      BeginProperty ButtonFont4 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState4    =   -1
      ButtonLeft4     =   164
      ButtonTop4      =   4
      ButtonWidth4    =   2
      ButtonHeight4   =   54
      ButtonCaption5  =   "Ajuda"
      ButtonEnabled5  =   0   'False
      ButtonIconSize5 =   32
      ButtonToolTipText5=   "Ajuda (F1)"
      ButtonKey5      =   "5"
      ButtonAlignment5=   2
      BeginProperty ButtonFont5 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft5     =   168
      ButtonTop5      =   2
      ButtonWidth5    =   36
      ButtonHeight5   =   21
      ButtonUseMaskColor5=   0   'False
      ButtonCaption6  =   "Sair"
      ButtonEnabled6  =   0   'False
      ButtonIconSize6 =   32
      ButtonToolTipText6=   "Sair (Esc)"
      ButtonKey6      =   "6"
      ButtonAlignment6=   2
      BeginProperty ButtonFont6 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft6     =   206
      ButtonTop6      =   2
      ButtonWidth6    =   26
      ButtonHeight6   =   21
      ButtonUseMaskColor6=   0   'False
      ButtonEnabled7  =   0   'False
      ButtonIconSize7 =   32
      ButtonKey7      =   "7"
      ButtonAlignment7=   2
      BeginProperty ButtonFont7 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState7    =   5
      ButtonLeft7     =   234
      ButtonTop7      =   2
      ButtonWidth7    =   24
      ButtonHeight7   =   24
      ButtonUseMaskColor7=   0   'False
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   4770
         Top             =   180
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmGermaqfer_bloqTurno.frx":0020
         Count           =   1
      End
   End
End
Attribute VB_Name = "frmGermaqfer_bloqTurno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ProcBloquear()
On Error GoTo tratar_erro

If txtStatus.Text = "Bloqueado" Then
    USMsgBox ("O turno " & frmGermaqfer.cmbturno & " da semana " & frmGermaqfer.cmbdia & " já esta bloqueado."), vbInformation, "CAPRIND v5.0"
    Exit Sub
End If
txtStatus.Text = "Bloqueado"
txtResponsavel.Text = pubUsuario
With txtObservacoes
    .Locked = False
    .TabStop = True
    .Text = ""
    .SetFocus
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcDesbloquear()
On Error GoTo tratar_erro

If txtStatus.Text = "Liberado" Then
    USMsgBox ("O turno " & frmGermaqfer.cmbturno & " da semana " & frmGermaqfer.cmbdia & " já esta liberado."), vbInformation, "CAPRIND v5.0"
    Exit Sub
End If
txtStatus.Text = "Liberado"
txtResponsavel.Text = pubUsuario
With txtObservacoes
    .Locked = False
    .TabStop = True
    .Text = ""
    .SetFocus
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvar()
On Error GoTo tratar_erro

Acao = "salvar"
If txtStatus = "" Then
    NomeCampo = "o status"
    ProcVerificaAcao
    Exit Sub
End If
If txtStatus.Text = "Liberado" Then
    Campo = "Bloqueado = 'False'"
    Acao = "desbloquear"
    Evento = "Liberar turno"
Else
    Campo = "Bloqueado = 'True'"
    Acao = "bloquear"
    Evento = "Bloquear turno"
End If

Set TBTempo = CreateObject("adodb.recordset")
With frmGermaqfer
    If Opt_registro.Value = True Then
        If .txtStatus1 = "" Then
            NomeCampo = "o registro"
            ProcVerificaAcao
            Exit Sub
        End If
        Conexao.Execute "Update CadmaqTurnos Set " & Campo & ", Data_bloq = '" & Date & "', obs_bloq = '" & txtObservacoes & "', responsavel_bloq = '" & txtResponsavel & "' where maquina = '" & .txtmaquina & "' and diasemana = '" & .cmbdia & "' and turno = " & .cmbturno
        .ProcRecalculaTempoTotalDia
        
        TBTempo.Open "Select * from CadMaqTurnos where maquina = '" & .txtmaquina & "' and diasemana = '" & .cmbdia & "' and turno = " & .cmbturno, Conexao, adOpenKeyset, adLockOptimistic
    Else
        If cmbturno = "" Then
            NomeCampo = "o turno"
            ProcVerificaAcao
            cmbturno.SetFocus
            Exit Sub
        End If
        Conexao.Execute "Update CadmaqTurnos Set " & Campo & ", Data_bloq = '" & Date & "', obs_bloq = '" & txtObservacoes & "', responsavel_bloq = '" & txtResponsavel & "' where maquina = '" & .txtmaquina & "' and turno = " & cmbturno
        ProcRecalculaTotalDia_Todos
        
        TBTempo.Open "Select * from CadMaqTurnos where maquina = '" & .txtmaquina & "' and turno = " & cmbturno, Conexao, adOpenKeyset, adLockOptimistic
    End If
    
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    If TBTempo.EOF = False Then
        Do While TBTempo.EOF = False
            '==================================
            Modulo = "PCP/Postos de trabalho"
            ID_documento = TBTempo!CODIGO
            Documento = "Código do posto de trabalho: " & .txtmaquina.Text
            Documento1 = "Dia da semana: " & TBTempo!Diasemana & " - Truno: " & TBTempo!Turno
            ProcGravaEvento
            '==================================
            TBTempo.MoveNext
        Loop
    End If
    TBTempo.Close
    
    .ProcCarregaTurnos
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
    Case vbKeyF3: ProcSalvar
    Case vbKeyF6: ProcBloquear
    Case vbKeyF7: ProcDesbloquear
    'Case vbKeyF1: ProcAjuda
    Case vbKeyEscape: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 6330, 6, True
ProcPuxaDadosSemanaTurno

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcPuxaDadosSemanaTurno()
On Error GoTo tratar_erro

With frmGermaqfer
    If .txtResponsavel1 <> "" Then
        Set TBClientes = CreateObject("adodb.recordset")
        TBClientes.Open "Select * from CadmaqTurnos where maquina = '" & .txtmaquina.Text & "' and diasemana = '" & .cmbdia.Text & "' and turno = " & .cmbturno.Text, Conexao, adOpenKeyset, adLockOptimistic
        If TBClientes!Bloqueado = True Then
            txtStatus.Text = "Bloqueado"
        Else
            txtStatus.Text = "Liberado"
        End If
        txtObservacoes.Text = IIf(IsNull(TBClientes!obs_bloq), "", TBClientes!obs_bloq)
        txtResponsavel.Text = IIf(IsNull(TBClientes!responsavel_bloq), "", TBClientes!responsavel_bloq)
        TBClientes.Close
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt_registro_Click()
On Error GoTo tratar_erro

ProcPuxaDadosSemanaTurno
With cmbturno
    .ListIndex = -1
    .Locked = False
    .TabStop = True
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optTurno_Click()
On Error GoTo tratar_erro

With cmbturno
    .Locked = False
    .TabStop = True
    .SetFocus
End With
txtStatus = ""
txtResponsavel = ""
txtObservacoes = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcRecalculaTotalDia_Todos()
On Error GoTo tratar_erro

DataFim = 0
With frmGermaqfer
    Set TBAfericao = CreateObject("adodb.recordset")
    TBAfericao.Open "Select maquina, DiaSemana from CadmaqTurnos where maquina = '" & .txtmaquina.Text & "' and turno = " & cmbturno & " Group by maquina, DiaSemana", Conexao, adOpenKeyset, adLockOptimistic
    If TBAfericao.EOF = False Then
        Do While TBAfericao.EOF = False
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select * from cadmaqturnos where maquina = '" & TBAfericao!maquina & "' and diasemana = '" & TBAfericao!Diasemana & "' and turno = " & cmbturno & " and bloqueado = 'False'", Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                Do While TBAbrir.EOF = False
                    Dataini = Left(TBAbrir!TotalTurno, 8)
                    DataFim = Format(DataFim + Dataini, "hh:mm:ss")
                    TBAbrir.MoveNext
                Loop
            End If
            TBAbrir.Close
            Conexao.Execute "Update cadmaqturnos Set TotalDia = '" & Format(DataFim, "hh:mm:ss") & "' where maquina = '" & TBAfericao!maquina & "' and diasemana = '" & TBAfericao!Diasemana & "' and turno = " & cmbturno
            TBAfericao.MoveNext
        Loop
    End If
    TBAfericao.Close
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcSalvar
    Case 2: ProcBloquear
    Case 3: ProcDesbloquear
    'Case 5: ProcAjuda
    Case 6: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
