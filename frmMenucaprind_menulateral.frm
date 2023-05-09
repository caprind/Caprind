VERSION 5.00
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Begin VB.Form frmMenucaprind_menulateral 
   Appearance      =   0  'Flat
   BackColor       =   &H00815135&
   ClientHeight    =   10680
   ClientLeft      =   540
   ClientTop       =   -45
   ClientWidth     =   19080
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
   Icon            =   "frmMenucaprind_menulateral.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10680
   ScaleWidth      =   19080
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   465
      Left            =   -1110
      TabIndex        =   0
      Top             =   2880
      Width           =   1005
   End
   Begin DrawSuite2022.USSysTray USSysTray1 
      Left            =   990
      Top             =   90
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   510
      Top             =   60
   End
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   30
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   99
      ResizeFonts     =   0   'False
      ScreenHeight    =   1080
      ScreenWidth     =   2560
      ScreenHeightDT  =   1080
      ScreenWidthDT   =   1920
      AutoResizeOnLoad=   0   'False
      ApplicationName =   "Active Resize Control Professional"
      FormHeightDT    =   10800
      FormWidthDT     =   19200
      FormScaleHeightDT=   10680
      FormScaleWidthDT=   19080
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin DrawSuite2022.USButton Btn1 
      Height          =   585
      Left            =   15360
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Clique aqui para acessar nossos tutoriais online"
      Top             =   9570
      Width           =   585
      _ExtentX        =   1032
      _ExtentY        =   1032
      DibPicture      =   "frmMenucaprind_menulateral.frx":0CCA
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   8474933
      BorderColorDisabled=   8474933
      BorderColorDown =   8474933
      BorderColorOver =   8474933
      GradientColor1  =   8474933
      GradientColor2  =   8474933
      GradientColor3  =   8474933
      GradientColor4  =   8474933
      GradientColorDisabled1=   14215660
      GradientColorDisabled2=   14215660
      GradientColorDisabled3=   14215660
      GradientColorDisabled4=   14215660
      GradientColorOver1=   12632319
      GradientColorOver2=   255
      GradientColorOver3=   255
      GradientColorOver4=   192
      GradientColorDown1=   0
      GradientColorDown2=   0
      GradientColorDown3=   0
      GradientColorDown4=   0
      PicAlign        =   7
      PicSize         =   3
      PicSizeH        =   32
      PicSizeW        =   32
      Theme           =   3
      ToolTipForeColor=   4210752
      ToolTipIcon     =   1
      ToolTipTitle    =   "Caprind Gestão Industrial"
   End
   Begin DrawSuite2022.USButton Btn2 
      Height          =   555
      Left            =   16170
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Clique aqui para acessar nosso site"
      Top             =   9570
      Width           =   585
      _ExtentX        =   1032
      _ExtentY        =   979
      DibPicture      =   "frmMenucaprind_menulateral.frx":1223
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   8474933
      BorderColorDisabled=   8474933
      BorderColorDown =   8474933
      BorderColorOver =   8474933
      ForeColor       =   8474933
      ForeColorDisabled=   8474933
      ForeColorOver   =   8474933
      ForeColorDown   =   8474933
      GradientColor1  =   8474933
      GradientColor2  =   8474933
      GradientColor3  =   8474933
      GradientColor4  =   8474933
      GradientColorDisabled1=   14215660
      GradientColorDisabled2=   14215660
      GradientColorDisabled3=   14215660
      GradientColorDisabled4=   14215660
      GradientColorOver1=   12632319
      GradientColorOver2=   255
      GradientColorOver3=   255
      GradientColorOver4=   192
      GradientColorDown1=   0
      GradientColorDown2=   0
      GradientColorDown3=   0
      GradientColorDown4=   0
      PicAlign        =   7
      PicSize         =   3
      PicSizeH        =   32
      PicSizeW        =   32
      Theme           =   3
      ToolTipForeColor=   4210752
      ToolTipIcon     =   1
   End
   Begin DrawSuite2022.USButton btn3 
      Height          =   555
      Left            =   16980
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Clique aqui para acessar nosso Facebook"
      Top             =   9570
      Width           =   585
      _ExtentX        =   1032
      _ExtentY        =   979
      DibPicture      =   "frmMenucaprind_menulateral.frx":171C
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   8474933
      BorderColorDisabled=   8474933
      BorderColorDown =   8474933
      BorderColorOver =   8474933
      GradientColor1  =   8474933
      GradientColor2  =   8474933
      GradientColor3  =   8474933
      GradientColor4  =   8474933
      GradientColorDisabled1=   14215660
      GradientColorDisabled2=   14215660
      GradientColorDisabled3=   14215660
      GradientColorDisabled4=   14215660
      GradientColorOver1=   12632319
      GradientColorOver2=   255
      GradientColorOver3=   255
      GradientColorOver4=   192
      GradientColorDown1=   0
      GradientColorDown2=   0
      GradientColorDown3=   0
      GradientColorDown4=   0
      PicAlign        =   7
      PicSize         =   3
      PicSizeH        =   32
      PicSizeW        =   32
      Theme           =   3
      ToolTipForeColor=   4210752
      ToolTipIcon     =   1
   End
   Begin DrawSuite2022.USButton btn4 
      Height          =   555
      Left            =   17790
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Clique aqui para acessar nosso canal no Youtube"
      Top             =   9570
      Width           =   585
      _ExtentX        =   1032
      _ExtentY        =   979
      DibPicture      =   "frmMenucaprind_menulateral.frx":1BEB
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   8474933
      BorderColorDisabled=   8474933
      BorderColorDown =   8474933
      BorderColorOver =   8474933
      ForeColor       =   8474933
      ForeColorDisabled=   8474933
      ForeColorOver   =   8474933
      ForeColorDown   =   8474933
      GradientColor1  =   8474933
      GradientColor2  =   8474933
      GradientColor3  =   8474933
      GradientColor4  =   8474933
      GradientColorDisabled1=   14215660
      GradientColorDisabled2=   14215660
      GradientColorDisabled3=   14215660
      GradientColorDisabled4=   14215660
      GradientColorOver1=   12632319
      GradientColorOver2=   255
      GradientColorOver3=   255
      GradientColorOver4=   192
      GradientColorDown1=   0
      GradientColorDown2=   0
      GradientColorDown3=   0
      GradientColorDown4=   0
      PicAlign        =   7
      PicSize         =   3
      PicSizeH        =   32
      PicSizeW        =   32
      Theme           =   3
      ToolTipForeColor=   4210752
      ToolTipIcon     =   1
   End
   Begin VB.Label lblMin 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1725
      Left            =   0
      TabIndex        =   5
      Top             =   660
      Width           =   585
   End
   Begin DrawSuite2022.USAlphaImage ImgAbrir 
      Height          =   1485
      Left            =   0
      TabIndex        =   9
      Top             =   720
      Visible         =   0   'False
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   2619
      Image           =   "frmMenucaprind_menulateral.frx":222F
      Props           =   5
   End
   Begin DrawSuite2022.USAlphaImage ImgFechar 
      Height          =   1485
      Left            =   0
      TabIndex        =   10
      Top             =   720
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   2619
      Image           =   "frmMenucaprind_menulateral.frx":3602
      Props           =   5
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Conheça nossas redes sociais"
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   15990
      TabIndex        =   8
      Top             =   10230
      Width           =   2130
   End
   Begin DrawSuite2022.USAlphaImage USAlphaImage1 
      Height          =   3435
      Left            =   6060
      TabIndex        =   11
      Top             =   3690
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   6059
      Image           =   "frmMenucaprind_menulateral.frx":4A9A
      ShadowDepth     =   10
   End
   Begin VB.Label lblAvisodiario2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Clique aqui para visualizar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   15660
      MouseIcon       =   "frmMenucaprind_menulateral.frx":1408F
      MousePointer    =   99  'Custom
      TabIndex        =   7
      ToolTipText     =   "Atenção ! Você tem nova(s) tarefa(s). Clique abaixo para visualizar."
      Top             =   420
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.Label lblAvisoDiário1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Atenção usuario! Você tem novo(s) avisos(s)"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   14955
      TabIndex        =   6
      ToolTipText     =   "Atenção ! Você tem nova(s) tarefa(s). Clique abaixo para visualizar."
      Top             =   180
      Visible         =   0   'False
      Width           =   3255
   End
End
Attribute VB_Name = "frmMenucaprind_menulateral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Mostraaviso As Boolean

Private Sub Btn1_Click()
On Error GoTo tratar_erro

Dim iret As Long
iret = ShellExecute(Me.hWnd, vbNullString, "http://www.caprind.com.br/Tutoriais", vbNullString, "c:\", SW_SHOWNORMAL)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Btn2_Click()
On Error GoTo tratar_erro

Dim iret As Long
iret = ShellExecute(Me.hWnd, vbNullString, "http://www.caprind.com.br", vbNullString, "c:\", SW_SHOWNORMAL)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub btn3_Click()
On Error GoTo tratar_erro

Dim iret As Long
iret = ShellExecute(Me.hWnd, vbNullString, "https://www.facebook.com/CaprindERP/", vbNullString, "c:\", SW_SHOWNORMAL)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub btn4_Click()
On Error GoTo tratar_erro

    Dim iret As Long
    iret = ShellExecute(Me.hWnd, vbNullString, "https://www.youtube.com/channel/UCQdWE_CO-BL_LD-fQmW06sg", vbNullString, "c:\", SW_SHOWNORMAL)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub btnSA_Click()
On Error GoTo tratar_erro
caminho = "C:\Program Files (x86)\Caprind\SuporteRemotoCaprind-idce3xzmxt.exe"

If FunVefificaModuloLocacao(True, True, True) = False Then Exit Sub

Formulario = "Suporte/Solicitação de atendimento"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
If TemInternet = True And ErroDriverMYSQL = False Then
    If FunVerificaManutencaoAtiva = False Then Exit Sub
    If FileOrDirExists(caminho) = True Then
     ProcAbrirArquivo (caminho) 'FrmAtendimento.Show
    Else
    USMsgBox "Arquivo de suporte a conexão remota não encontrado na pasta " & caminho & "!", vbCritical, "CAPRIND v5.0"
    End If
Else
    If TemInternet = False Then MsgTexto = "não foi encontrado conexão com a internet" Else MsgTexto = "no momento estamos sem conexão com nosso servidor, favor tentar mais tarde"
    USMsgBox ("Não é permitido abrir este módulo, pois " & MsgTexto & "."), vbExclamation, "CAPRIND v5.0"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro
Mostraaviso = False
    
        With Btn1
        .BorderColor = RGB(53, 81, 129)
        
        .GradientColorOver1 = RGB(53, 81, 129)
        .GradientColorOver2 = RGB(53, 81, 129)
        .GradientColorOver3 = RGB(53, 81, 129)
        .GradientColorOver4 = RGB(53, 81, 129)
       ' .BorderColorOver = RGB(255, 255, 255)
        End With
        
        With Btn2
        .BorderColor = RGB(53, 81, 129)
        
        .GradientColorOver1 = RGB(53, 81, 129)
        .GradientColorOver2 = RGB(53, 81, 129)
        .GradientColorOver3 = RGB(53, 81, 129)
        .GradientColorOver4 = RGB(53, 81, 129)
       ' .BorderColorOver = RGB(255, 255, 255)
        End With
        
        With btn3
        .BorderColor = RGB(53, 81, 129)
        
        .GradientColorOver1 = RGB(53, 81, 129)
        .GradientColorOver2 = RGB(53, 81, 129)
        .GradientColorOver3 = RGB(53, 81, 129)
        .GradientColorOver4 = RGB(53, 81, 129)
       ' .BorderColorOver = RGB(255, 255, 255)
        End With
        
         With btn4
        .BorderColor = RGB(53, 81, 129)
        
        .GradientColorOver1 = RGB(53, 81, 129)
        .GradientColorOver2 = RGB(53, 81, 129)
        .GradientColorOver3 = RGB(53, 81, 129)
        .GradientColorOver4 = RGB(53, 81, 129)
        '.BorderColorOver = RGB(255, 255, 255)
        End With
       
        
    
    Caption = "CÓPIA REGISTRADA"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub lblAvisodiario2_Click()
On Error GoTo tratar_erro

    lblAvisoDiário1.Visible = False
    lblAvisodiario2.Visible = False
    FrmMDI_AvisoDiario.Show
    Timer1.Enabled = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub lblMin_Click()
On Error GoTo tratar_erro

If ImgFechar.Visible = True Then
    ImgAbrir.Visible = True
    ImgFechar.Visible = False
    frmMDI.picMain.Visible = False
Else
    ImgAbrir.Visible = False
    ImgFechar.Visible = True
    frmMDI.picMain.Visible = True
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub lbltutoriais_Click()
On Error GoTo tratar_erro

    Dim iret As Long
    iret = ShellExecute(Me.hWnd, vbNullString, "http://www.caprind.com.br/Tutoriais", vbNullString, "c:\", SW_SHOWNORMAL)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Timer1_Timer()
On Error GoTo tratar_erro

    lblAvisoDiário1.Visible = Not lblAvisoDiário1.Visible
    lblAvisodiario2.Visible = Not lblAvisodiario2.Visible

    'frmMDI.StatusBar1.Panels(4).Visible = Not frmMDI.StatusBar1.Panels(4).Visible

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USSysTray1_Click()
On Error GoTo tratar_erro

    lblAvisoDiário1.Visible = False
    lblAvisodiario2.Visible = False
    FrmMDI_AvisoDiario.Show
    'Timer1.Enabled = False
    'USSysTray1.SysTrayRemoveIcon
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub

End Sub

