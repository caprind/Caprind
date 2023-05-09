VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Begin VB.Form frmBackup_apontamentos 
   Appearance      =   0  'Flat
   BackColor       =   &H00F9F9F9&
   BorderStyle     =   0  'None
   Caption         =   "Backup - Apontamentos"
   ClientHeight    =   2760
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   3690
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   3690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   405
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   3690
      _ExtentX        =   6509
      _ExtentY        =   714
      DibPicture      =   "frmBackup_apontamentos.frx":0000
      CaptionOnCenter =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Icon            =   "frmBackup_apontamentos.frx":108C4
   End
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   5
      Top             =   2355
      Width           =   3690
      _ExtentX        =   6509
      _ExtentY        =   714
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
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
      Height          =   615
      Left            =   30
      TabIndex        =   1
      Top             =   1440
      Width           =   3615
      Begin MSComCtl2.DTPicker Msk_data 
         Height          =   315
         Left            =   2250
         TabIndex        =   0
         Top             =   180
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   16777215
         CalendarForeColor=   0
         CalendarTitleBackColor=   8421504
         CalendarTitleForeColor=   16777215
         CalendarTrailingForeColor=   255
         Format          =   179896323
         CurrentDate     =   39057
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Ordem(ns) concluída(s) até:"
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
         Index           =   0
         Left            =   180
         TabIndex        =   2
         Top             =   270
         Width           =   2010
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   30
      TabIndex        =   3
      Top             =   450
      Width           =   3615
      _ExtentX        =   6376
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
      ButtonCaption1  =   "Salvar"
      ButtonEnabled1  =   0   'False
      ButtonIconSize1 =   32
      ButtonToolTipText1=   "Salvar (F3)"
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
      ButtonWidth1    =   44
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
      ButtonLeft2     =   48
      ButtonTop2      =   4
      ButtonWidth2    =   2
      ButtonHeight2   =   54
      ButtonCaption3  =   "Ajuda"
      ButtonEnabled3  =   0   'False
      ButtonIconSize3 =   32
      ButtonToolTipText3=   "Ajuda (F1)"
      ButtonKey3      =   "5"
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
      ButtonLeft3     =   52
      ButtonTop3      =   2
      ButtonWidth3    =   41
      ButtonHeight3   =   21
      ButtonUseMaskColor3=   0   'False
      ButtonCaption4  =   "Sair"
      ButtonEnabled4  =   0   'False
      ButtonIconSize4 =   32
      ButtonToolTipText4=   "Sair (Esc)"
      ButtonKey4      =   "6"
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
      ButtonLeft4     =   95
      ButtonTop4      =   2
      ButtonWidth4    =   30
      ButtonHeight4   =   21
      ButtonUseMaskColor4=   0   'False
      ButtonEnabled5  =   0   'False
      ButtonIconSize5 =   32
      ButtonKey5      =   "7"
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
      ButtonLeft5     =   127
      ButtonTop5      =   2
      ButtonWidth5    =   24
      ButtonHeight5   =   24
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   2190
         Top             =   120
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmBackup_apontamentos.frx":10BDE
         Count           =   1
      End
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   30
      TabIndex        =   4
      Top             =   2070
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor2      =   0
      SearchText      =   ""
      Value           =   0
   End
End
Attribute VB_Name = "frmBackup_apontamentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

If USMsgBox("Deseja realmente gerar backup dos apontamentos?", vbYesNo, "Caprind") = vbYes Then
    Set TBUsuarios = CreateObject("adodb.recordset")
    TBUsuarios.Open "Select U.IDUsuario from usuarios U INNER JOIN acessos A ON A.IDUsuario = U.IDUsuario where U.usuario = '" & pubUsuario & "' and A.Acesso = 'PCP/Gerenciamento de ordem/Validar resultados' and A.Validacao = 'True'", Conexao, adOpenKeyset, adLockOptimistic
    If TBUsuarios.EOF = True Then
        USMsgBox ("Não é permitido gerar o backup, pois o usuário " & pubUsuario & " não tem autorização para validar o resultado da(s) ordem(ns)."), vbExclamation, "CAPRIND v5.0"
        TBUsuarios.Close
        Exit Sub
    End If
    TBUsuarios.Close
    
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select P.Ordem from Producao P INNER JOIN ProducaoFases PF ON P.Ordem = PF.Ordem where P.Pronta = 'SIM' and P.Dataentrega <= '" & msk_data & "' and P.DtValidacao IS NOT NULL Group by P.Ordem", Conexao, adOpenKeyset, adLockReadOnly
    If TBAbrir.EOF = False Then
        PBLista.Min = 0
        PBLista.Max = TBAbrir.RecordCount
        PBLista.Value = 1
        Contador = 0
        Do While TBAbrir.EOF = False
            'Salva dados na tabela ProducaoFases_backup
            Conexao.Execute "INSERT INTO ProducaoFases_Backup (Ordem, IDFASE, CodigoDesc, quantidade, Descricao, Fase, maquina, Usuario, TempoInicio, TempoFinal, TempoTotal, Pronto, Dias, Preparacao, Execucao, Data, Quant, Reprovada, OS, Turno, TempoTotalSeg, QTCD) Select Ordem, IDFASE, CodigoDesc, quantidade, Descricao, Fase, maquina, Usuario, TempoInicio, TempoFinal, TempoTotal, Pronto, Dias, Preparacao, Execucao, Data, Quant, Reprovada, OS, Turno, TempoTotalSeg, QTCD from ProducaoFases where Ordem = " & TBAbrir!Ordem & " order by Tempoinicio"
            
            'Altera a ordem para backup
            Conexao.Execute "Update Producao Set Ap_backup = 'True', DtValidacao_custo = '" & Date & "', RespValidacao_custo = '" & pubUsuario & "' where Ordem = " & TBAbrir!Ordem
            
            'Altera ID da produção na NC
            Conexao.Execute "Update CQNCF set CQNCF.idproducao = PFB.IDproducao from (CQ_NC_FABRICA CQNCF INNER JOIN ProducaoFases PF on CQNCF.IDproducao = PF.IDproducao) INNER JOIN ProducaoFases_Backup PFB ON PFB.OS = PF.OS and PFB.Tempoinicio = PF.Tempoinicio where PF.Ordem = " & TBAbrir!Ordem & " and PF.Reprovada <> 0"
            
            'Altera ID da produção na manutenção
            Conexao.Execute "Update MD set MD.idproducao2 = PFB.IDproducao from (Manutencao_data MD INNER JOIN ProducaoFases PF on MD.IDproducao2 = PF.IDproducao) INNER JOIN ProducaoFases_Backup PFB ON PFB.OS = PF.OS and PFB.Tempoinicio = PF.Tempoinicio where PF.Ordem = " & TBAbrir!Ordem
            
            'Apaga dados da tabela ProducaoFases
            Conexao.Execute "DELETE from ProducaoFases where Ordem = " & TBAbrir!Ordem
            
            'Salva dados na tabela ProducaoFases_Totalizacao_backup
            Conexao.Execute "INSERT INTO ProducaoFases_Totalizacao_Backup (Ordem, OS, Fase, Data, Usuario, maquina, Turno, Pronto, Preparacao, Execucao, QTNC, QTOK, TPUTIL, TEUTIL, TETTUTIL, CRLOTE, CRPECA, CPLOTE, CPPECA, Eficiencia, Totalprod, Eficiencia_prep, Eficiencia_exec, Valor_hs_prep, Valor_hs_exec) Select Ordem, OS, Fase, Data, Usuario, maquina, Turno, Pronto, Preparacao, Execucao, QTNC, QTOK, TPUTIL, TEUTIL, TETTUTIL, CRLOTE, CRPECA, CPLOTE, CPPECA, Eficiencia, Totalprod, Eficiencia_prep, Eficiencia_exec, Valor_hs_prep, Valor_hs_exec from ProducaoFases_Totalizacao where Ordem = " & TBAbrir!Ordem
                
            'Apaga dados da tabela ProducaoFases_Totalizacao
            Conexao.Execute "DELETE from ProducaoFases_Totalizacao where Ordem = " & TBAbrir!Ordem
            
            TBAbrir.MoveNext
            Contador = Contador + 1
            PBLista.Value = Contador
        Loop
    End If
    TBAbrir.Close
    USMsgBox ("Backup efetuado com sucesso."), vbInformation, "CAPRIND v5.0"
    '==================================
    Modulo = Formulario
    Evento = "Criar backup apontamentos"
    ID_documento = 0
    Documento = "Até: " & msk_data
    Documento1 = ""
    ProcGravaEvento
    '==================================
    Unload Me
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyF3: ProcSalvar
    Case vbKeyEscape: ProcSair
End Select
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 3615, 5, True
Formulario = "Configuração do sistema/Criar backup/Apontamentos"
msk_data.Value = Date

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcSalvar
    'Case 3: ProcAjuda
    Case 4: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
