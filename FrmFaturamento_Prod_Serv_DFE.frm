VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form FrmFaturamento_Prod_Serv_DFE 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Nota Fiscal | Baixar Danfe e XML (DFE)"
   ClientHeight    =   5055
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5640
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5055
   ScaleWidth      =   5640
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   9
      Top             =   4650
      Width           =   5640
      _ExtentX        =   9948
      _ExtentY        =   714
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   1515
      Left            =   420
      TabIndex        =   5
      Top             =   1650
      Width           =   4815
      Begin VB.TextBox txtD2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   320
         Left            =   210
         TabIndex        =   10
         Top             =   1050
         Width           =   4395
      End
      Begin VB.TextBox txtchNFe 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   210
         MaxLength       =   44
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Chave de acesso NFe."
         Top             =   450
         Width           =   4395
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Pasta Danfe e XML"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   1740
         TabIndex        =   11
         Top             =   840
         Width           =   1350
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Chave de acesso"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   1792
         TabIndex        =   7
         Top             =   240
         Width           =   1230
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Opções para download"
      Height          =   885
      Left            =   420
      TabIndex        =   1
      Top             =   600
      Width           =   4815
      Begin DrawSuite2022.USOptionButton opt1 
         Height          =   525
         Left            =   1785
         TabIndex        =   2
         Top             =   270
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   926
         Caption         =   "DANFE"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocusRect   =   0   'False
      End
      Begin DrawSuite2022.USOptionButton opt2 
         Height          =   525
         Left            =   2910
         TabIndex        =   3
         Top             =   270
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   926
         Caption         =   "XML"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocusRect   =   0   'False
      End
      Begin DrawSuite2022.USOptionButton opt3 
         Height          =   525
         Left            =   240
         TabIndex        =   4
         Top             =   270
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   926
         Caption         =   "DANFE e XML"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocusRect   =   0   'False
         Value           =   -1  'True
      End
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   435
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5640
      _ExtentX        =   9948
      _ExtentY        =   741
      DibPicture      =   "FrmFaturamento_Prod_Serv_DFE.frx":0000
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
      Icon            =   "FrmFaturamento_Prod_Serv_DFE.frx":29A1
      ShowMaximize    =   0   'False
      ShowMinimize    =   0   'False
   End
   Begin DrawSuite2022.USButton cmdBaixar 
      Height          =   825
      Left            =   420
      TabIndex        =   8
      ToolTipText     =   "Executar download Danfe e XML"
      Top             =   3330
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   1455
      DibPicture      =   "FrmFaturamento_Prod_Serv_DFE.frx":2CBB
      BorderColor     =   5263559
      BorderColorDisabled=   13160660
      BorderColorDown =   4013465
      BorderColorOver =   4408288
      Caption         =   "Baixar Danfe e XML"
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
      PicAlign        =   7
      PicSize         =   2
      PicSizeH        =   24
      PicSizeW        =   24
      ShowFocusRect   =   0   'False
      ShowFocusRectDown=   0   'False
      Theme           =   4
      ToolTipTitle    =   "CAPRIND v5.0"
   End
   Begin DrawSuite2022.USButton cmdD2 
      Height          =   825
      Left            =   2880
      TabIndex        =   12
      ToolTipText     =   "Abrir pasta DANFE-XML..."
      Top             =   3330
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   1455
      DibPicture      =   "FrmFaturamento_Prod_Serv_DFE.frx":565C
      BorderColor     =   4960354
      BorderColorDisabled=   13160660
      BorderColorDown =   4210752
      BorderColorOver =   49152
      Caption         =   "Abrir pasta Danfe e XML"
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
      PicAlign        =   8
      PicSize         =   2
      PicSizeH        =   24
      PicSizeW        =   24
      ShowFocusRect   =   0   'False
      ShowFocusRectDown=   0   'False
      Theme           =   3
      ToolTipTitle    =   "CAPRIND v5.0"
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   4290
      Width           =   5385
      _ExtentX        =   9499
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
      SearchText      =   "Atualizando..."
      Value           =   0
   End
End
Attribute VB_Name = "FrmFaturamento_Prod_Serv_DFE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ProcBaixarTodas()
On Error GoTo tratar_erro
Dim resposta As String

  Set TBAbrir_NFe = CreateObject("adodb.recordset")
    StrSql = "select DNF.dt_DataEmissao, DNF.ID_empresa, DNFE.Chave_acesso from tbl_Dados_Nota_Fiscal DNF inner Join tbl_Dados_Nota_Fiscal_NFe DNFE on DNF.ID = DNFE.ID_nota where DNFE.Chave_acesso <> '' and DNF.ID_empresa = '" & IDempresa & "' and cast(DNF.dt_DataEmissao as date) between cast( dateadd (day,-90,getdate()) as date) and cast (getdate() as date) and DNF.int_TipoNota = '1'"
    TBAbrir_NFe.Open StrSql, Conexao, adOpenKeyset, adLockReadOnly
    
    
    If TBAbrir_NFe.EOF = False Then
    
    PBLista.Min = 0
    PBLista.Max = TBAbrir_NFe.RecordCount
    PBLista.Value = 1
    Contador = 0
      Do While TBAbrir_NFe.EOF = False
        chNNfe = TBAbrir_NFe!Chave_acesso
        resposta = NFeAPI.downloadNFeAndSave(chNNfe, tpAmb, "X", DiretorioXMLDanfe, False)
        resposta = NFeAPI.downloadNFeAndSave(chNNfe, tpAmb, "P", DiretorioXMLDanfe, False)
        TBAbrir_NFe.MoveNext
        Contador = Contador + 1
        PBLista.Value = Contador
        FrmFaturamento_Prod_Serv_DFE.Refresh
      Loop
   status = LerDadosJSON(resposta, "status", "", "")
   End If
   
   If status = "200" Then
   USMsgBox "Documento(s) fiscal(is) baixado(s) com sucesso!" & vbCrLf & "Verifique o(s) documento(s) e mova-o(s) para a pasta correta." & vbCrLf & "Lembre-se de sempre deixar vazia a pasta de downloads.", vbInformation, "CAPRIND v5.0"
   ShellExecute 0, "open", DiretorioXMLDanfe, "", "", vbNormalFocus
   End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub cmdBaixar_Click()
On Error GoTo tratar_erro
Dim resposta As String

If USMsgBox("Deseja realmente baixar o(s) documento(s) fiscal(is)?", vbYesNo, "CAPRIND v5.0") = vbYes Then
   If Len(txtchNFe.Text) < 44 Then
      ProcBaixarTodas
      Exit Sub
   End If
   
   chNNfe = txtchNFe.Text
   frmFaturamento_Prod_Serv_NFe_NS.ProcCriarPastaDanfe
    frmFaturamento_Prod_Serv_NFe_NS.ProcCriarPastaXML
      
   If NFCe = False Then
   resposta = NFeAPI.downloadNFeAndSave(chNNfe, tpAmb, "X", DiretorioDanfe, False)
   resposta = NFeAPI.downloadNFeAndSave(chNNfe, tpAmb, "P", DiretorioXML, False)
   Else
    resposta = NFCe_downloadESalvar(chNNfe, tpAmb, DiretorioDanfe, False)
    resposta = NFCe_downloadESalvar(chNNfe, tpAmb, DiretorioXML, False)
   End If
   status = LerDadosJSON(resposta, "status", "", "")
   
   If status = "200" Or status = "100" Then
   USMsgBox "Documento(s) fiscal(is) baixado(s) com sucesso!", vbInformation, "CAPRIND v5.0"
   ShellExecute 0, "open", DiretorioXMLDanfe, "", "", vbNormalFocus
   Else
   USMsgBox "Ocorreu um erro ao baixar documentos, por favor tente mais tarde.", vbInformation, "CAPRIND v5.0"
   End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub cmdD2_Click()
On Error GoTo tratar_erro

  ShellExecute 0, "open", DiretorioXMLDanfe, "", "", vbNormalFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub
