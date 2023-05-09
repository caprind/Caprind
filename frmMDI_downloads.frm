VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmMDI_downloads 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Downloads - Nota fiscal"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   10770
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   10770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Centralziar na Tela
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Local"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3825
      Left            =   55
      TabIndex        =   4
      Top             =   840
      Width           =   5535
      Begin VB.FileListBox Lista_arquivos 
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
         Height          =   3450
         Left            =   3000
         MultiSelect     =   2  'Extended
         TabIndex        =   7
         Top             =   225
         Width           =   2355
      End
      Begin VB.DirListBox Lista_pastas 
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
         Height          =   2970
         Left            =   195
         TabIndex        =   6
         Top             =   705
         Width           =   2745
      End
      Begin VB.DriveListBox Lista_driver 
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
         Left            =   189
         TabIndex        =   5
         Top             =   230
         Width           =   2745
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Procam"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3825
      Left            =   6690
      TabIndex        =   2
      Top             =   840
      Width           =   4035
      Begin VB.ListBox Lista_Procam 
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
         Height          =   3420
         Left            =   180
         MultiSelect     =   2  'Extended
         TabIndex        =   3
         Top             =   225
         Width           =   3645
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   790
      Left            =   8010
      TabIndex        =   0
      Top             =   0
      Width           =   2715
      Begin VB.Label Lbl_status 
         Alignment       =   2  'Centralizar
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Pronto"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   180
         TabIndex        =   1
         Top             =   330
         Width           =   2355
      End
   End
   Begin InetCtlsObjects.Inet itcFTP 
      Left            =   5820
      Top             =   900
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   2
      RemotePort      =   21
      URL             =   "ftp://"
   End
   Begin DrawSuite2022.USButton Cmd_conectar 
      Height          =   705
      Left            =   6780
      TabIndex        =   11
      Top             =   90
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   1244
      DibPicture      =   "frmMDI_downloads.frx":0000
      BorderColor     =   14404026
      BorderColorDown =   11632444
      BorderColorOver =   11632444
      Caption         =   "Conectar"
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
      HandPointer     =   0   'False
      PicAlign        =   0
      PicSize         =   5
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   790
      Left            =   55
      TabIndex        =   8
      Top             =   0
      Width           =   6675
      Begin VB.ComboBox Cmb_empresa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         ItemData        =   "frmMDI_downloads.frx":3B1B
         Left            =   210
         List            =   "frmMDI_downloads.frx":3B1D
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   9
         ToolTipText     =   "Empresa."
         Top             =   340
         Width           =   6285
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparente
         Caption         =   "Empresa"
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
         Left            =   2985
         TabIndex        =   10
         Top             =   150
         Width           =   735
      End
   End
   Begin DrawSuite2022.USButton Cmd_baixar 
      Height          =   465
      Left            =   5652
      TabIndex        =   12
      Top             =   2610
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   820
      DibPicture      =   "frmMDI_downloads.frx":3B1F
      BorderColor     =   14404026
      BorderColorDown =   11632444
      BorderColorOver =   11632444
      Caption         =   "Download"
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
      PicAlign        =   8
      PicSize         =   5
      State           =   3
   End
End
Attribute VB_Name = "frmMDI_downloads"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Diretorio As String 'OK

Private Sub Cmd_conectar_Click()
On Error GoTo tratar_erro

Acao = "baixar"
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select CNPJ from Empresa where Codigo = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex), Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    If IsNull(TBAbrir!CNPJ) = True Or TBAbrir!CNPJ = "" Then
        USMsgBox ("É necessário cadastrar o CNPJ da empresa antes de conectar."), vbExclamation, "CAPRIND v5.0"
        TBAbrir.Close
        Exit Sub
    End If
    Familiatext = ReturnNumbersOnly(TBAbrir!CNPJ)
End If
TBAbrir.Close
If Downloads_NF = True Then FamiliaAntiga = "NF" Else FamiliaAntiga = "Boletos"
If Left(Cmd_conectar.Caption, 1) = "C" Then ProcConectar Else ProcDesconectar

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcConectar()
On Error GoTo tratar_erro

With itcFTP
    .url = "caprind.com.br"
    .UserName = "caprind1"
    .Password = "cap0902loc"
End With
ProcCarregaListaProcam
ProcExecutarComando "cd " & "public_html", True
ProcExecutarComando "cd " & "Arquivos", True
ProcExecutarComando "cd " & "Clientes", True
ProcExecutarComando "cd " & Familiatext, True
ProcExecutarComando "cd " & FamiliaAntiga, True

Diretorio = "http://www.caprind.com.br/Arquivos/Clientes/" & Familiatext & "/" & FamiliaAntiga

Frame1.Enabled = False
Cmd_baixar.Enabled = True
Frame4.Enabled = True
Cmd_conectar.Caption = "Desconectar"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcDesconectar()
On Error GoTo tratar_erro
    
itcFTP.Execute , "Quit"
ProcExecutarComando "Quit", False
Lista_Procam.Clear
Frame1.Enabled = True
Cmd_baixar.Enabled = False
Frame4.Enabled = False
Cmd_conectar.Caption = "Conectar"
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmd_baixar_Click()
On Error GoTo tratar_erro
Dim NomeArquivo As String, Origem As String, Destino As String

Permitido = False
Permitido1 = False
Contador = 0
For Contador = 0 To Lista_Procam.ListCount - 1
    If Lista_Procam.Selected(Contador) = True Then
        Permitido = True
        NomeArquivo = Lista_Procam.List(Contador)
        Origem = Diretorio & "/" & NomeArquivo
        If Len(Lista_pastas.Path) > 3 Then Destino = Lista_pastas.Path & "\" & NomeArquivo Else Destino = Lista_pastas.Path & NomeArquivo
        FunBaixarArquivoNET Origem, Destino
        Lista_Procam.Selected(Contador) = False
    End If
Next Contador
If Permitido = False Then
    USMsgBox ("Informe o(s) arquivos antes de fazer o download."), vbExclamation, "CAPRIND v5.0"
ElseIf Permitido1 = True Then
        USMsgBox ("Download efetuado com sucesso."), vbInformation, "CAPRIND v5.0"
        Lista_arquivos.Refresh
    Else
        USMsgBox ("Ocorreu um erro ao efetuar o download."), vbExclamation, "CAPRIND v5.0"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyEscape: ProcSair
End Select
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_pastas_Change()
On Error GoTo tratar_erro

Lista_arquivos.Path = Lista_pastas.Path

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_driver_Change()
On Error GoTo tratar_erro
    
Lista_pastas.Path = Lista_driver.Drive

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

If Downloads_NF = True Then Caption = "Downloads - Nota fiscal" Else Caption = "Downloads - Boletos"
ProcCarregaComboEmpresa Cmb_empresa, False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo tratar_erro
    
ProcSair
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub itcFTP_StateChanged(ByVal State As Integer)
On Error GoTo tratar_erro

With Lbl_status
    Select Case State
        Case icResolvingHost: .Caption = "Resolvendo host..."
        Case icHostResolved: .Caption = "Host resolvido"
        Case icConnecting: .Caption = "Conectando..."
        Case icConnected: .Caption = "Conectado"
        Case icRequesting: .Caption = "Requisitando..."
        Case icRequestSent: .Caption = "Requisição enviada"
        Case icReceivingResponse: .Caption = "Recebendo..."
        Case icResponseReceived: .Caption = "Resposta recebida"
        Case icDisconnecting: .Caption = "Desconectando..."
        Case icDisconnected: .Caption = "Desconectado"
        Case icError: .Caption = itcFTP.ResponseInfo
        Case icResponseCompleted: .Caption = "Operacao completa"
    End Select
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExecutarComando(ByVal Operacao As String, ByVal Ld As Boolean)
On Error GoTo tratar_erro

With itcFTP
    If .StillExecuting Then .Cancel
    .Execute , Operacao
End With
ProcTerminarComando
If Ld = True Then
    ProcCarregaListaProcam
    ProcTerminarComando
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcTerminarComando()
On Error GoTo tratar_erro
    
Do While itcFTP.StillExecuting
    DoEvents
Loop

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaListaProcam()
On Error GoTo tratar_erro
Dim Data1 As Variant
Dim Inicio As Integer, Length As Integer

Contador = 1
Inicio = 1
Lista_Procam.Clear
ProcExecutarComando "dir", False
Do
    data = itcFTP.GetChunk(1024, icString)
    DoEvents
    For Contador = 1 To Len(data)
        If Mid(data, Contador, 1) = Chr(13) Then
            If Length > 0 And Mid(data, Inicio, Length) <> "./" Then Lista_Procam.AddItem Mid(data, Inicio, Length)
            Inicio = Contador + 2
            Length = -1
        Else
            Length = Length + 1
       End If
    Next Contador
Loop While LenB(data) > 0
ProcExecutarComando "pwd", False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

If Cmd_baixar.Enabled = True Then itcFTP.Execute , "Quit"
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
