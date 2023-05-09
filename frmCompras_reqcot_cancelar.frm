VERSION 5.00
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Begin VB.Form frmCompras_reqcot_cancelar 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cotação - Cancelar"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4230
   ClipControls    =   0   'False
   ControlBox      =   0   'False
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
   Icon            =   "frmCompras_reqcot_cancelar.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   4230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   3030
      Top             =   180
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frmCompras_reqcot_cancelar.frx":0442
      Count           =   1
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   55
      TabIndex        =   2
      Top             =   0
      Width           =   4125
      _ExtentX        =   7276
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
      ButtonCaption1  =   "Cancelar"
      ButtonEnabled1  =   0   'False
      ButtonIconSize1 =   32
      ButtonToolTipText1=   "Cancelar (F4)"
      ButtonKey1      =   "1"
      ButtonAlignment1=   2
      BeginProperty ButtonFont1 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft1     =   2
      ButtonTop1      =   2
      ButtonWidth1    =   57
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
      ButtonLeft2     =   61
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
      ButtonLeft3     =   65
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
      ButtonLeft4     =   108
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
      ButtonLeft5     =   140
      ButtonTop5      =   2
      ButtonWidth5    =   24
      ButtonHeight5   =   24
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Motivo do cancelamento da cotação :"
      Height          =   2115
      Left            =   55
      TabIndex        =   1
      Top             =   1020
      Width           =   4125
      Begin VB.TextBox txtmotivo 
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
         Height          =   1725
         Left            =   210
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   0
         ToolTipText     =   "Motivo do cancelamento da cotação."
         Top             =   270
         Width           =   3705
      End
   End
End
Attribute VB_Name = "frmCompras_reqcot_cancelar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyEscape: Unload Me
    Case vbKeyF4: ProcExcluir
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 4125, 5, True
If frmcompras_reqcot.txtStatus = "CANCELADA" Then
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from Compras_Cotacao where id_cotacao = " & Cont, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        txtMotivo = IIf(IsNull(TBAbrir!motivo), "", TBAbrir!motivo)
    End If
    TBAbrir.Close
    txtMotivo.Locked = True
    txtMotivo.TabStop = False
Else
    txtMotivo.Locked = False
    txtMotivo.TabStop = True
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluir()
On Error GoTo tratar_erro

If txtMotivo.Locked = True Then
    USMsgBox ("A cotação já está cancelada."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If txtMotivo.Text = "" Then
    USMsgBox ("Informe o motivo antes de alterar o status para cancelado."), vbExclamation, "CAPRIND v5.0"
    txtMotivo.SetFocus
    Exit Sub
End If
If USMsgBox("Deseja realmente alterar o status para cancelado?", vbYesNo) = vbYes Then
        With frmcompras_reqcot
        Cont = .txtidcot
        Conexao.Execute "Update compras_pedido_lista Set Status_item = 'CANCELADO' where id_cotacao = " & Cont
        Conexao.Execute "Update compras_cotacao Set Statuscotacao = 'CANCELADA', deletou = '" & pubUsuario & "', datacancelada = '" & Date & "', motivo = '" & txtMotivo & "' where id_cotacao = " & Cont
        '==================================
        Modulo = "Compras/Cotação"
        Evento = "Cancelar"
        ID_documento = Cont
        Documento = "Nº cotação: " & .txtidcotacao
        Documento1 = ""
        ProcGravaEvento
        '==================================
        USMsgBox ("Cotação cancelada com sucesso."), vbInformation, "CAPRIND v5.0"
        .txtStatus = "CANCELADA"
    End With
    Unload Me
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcExcluir
    'Case 3: ProcAjuda
    Case 4: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
