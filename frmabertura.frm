VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{84147065-0227-424E-827F-9E79B1DA5D8B}#21.0#0"; "kftp.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Begin VB.Form frmabertura 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00815135&
   BorderStyle     =   0  'None
   ClientHeight    =   11505
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   19200
   ClipControls    =   0   'False
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmabertura.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   11505
   ScaleWidth      =   19200
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtSenha 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   8085
      PasswordChar    =   "*"
      TabIndex        =   2
      ToolTipText     =   "Digite sua senha de usuário do sistema."
      Top             =   8865
      Visible         =   0   'False
      Width           =   3075
   End
   Begin VB.TextBox txtUsuario 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
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
      Height          =   330
      Left            =   8085
      TabIndex        =   1
      ToolTipText     =   "Digite seu nome de usuário do sistema."
      Top             =   8452
      Visible         =   0   'False
      Width           =   3075
   End
   Begin VB.TextBox cmbBanco 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00BC8452&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   8085
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Nome do banco de dados"
      Top             =   8040
      Visible         =   0   'False
      Width           =   3075
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   19200
      _ExtentX        =   33867
      _ExtentY        =   688
      BackColor       =   12354642
      CaptionOnCenter =   -1  'True
      EnableMaximizeButton=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   2250
      Top             =   4770
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3060
      Top             =   4770
   End
   Begin VB.Timer Timer1 
      Interval        =   11000
      Left            =   3720
      Top             =   4800
   End
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   660
      Top             =   4770
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   26
      ScreenHeight    =   1080
      ScreenWidth     =   1920
      ScreenHeightDT  =   1080
      ScreenWidthDT   =   1920
      AutoResizeOnLoad=   0   'False
      ApplicationName =   "Active Resize Control Professional"
      FormHeightDT    =   11505
      FormWidthDT     =   19200
      FormScaleHeightDT=   11505
      FormScaleWidthDT=   19200
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin KFTPActiveX.kftp kftp 
      Height          =   600
      Left            =   960
      TabIndex        =   6
      Top             =   -840
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   1058
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   165
      Left            =   0
      TabIndex        =   11
      Top             =   11340
      Width           =   19230
      _ExtentX        =   33920
      _ExtentY        =   291
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BarColor1       =   8388608
      BarColor2       =   14737632
      BorderColor     =   16777215
      SearchText      =   "Atualizando..."
      ShowText        =   0   'False
      Theme           =   6
      Value           =   0
   End
   Begin DrawSuite2022.USButton Cmd_chat 
      Height          =   495
      Left            =   9645
      TabIndex        =   4
      ToolTipText     =   "Precisa de ajuda? Clique aqui para permitir uma conexão remota"
      Top             =   9240
      Visible         =   0   'False
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   873
      DibPicture      =   "frmabertura.frx":0CCA
      Caption         =   "Ajuda?"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   5263559
      BorderColorDown =   2039646
      BorderColorOver =   3026574
      GradientColor1  =   5263559
      GradientColor2  =   5263559
      GradientColor3  =   5263559
      GradientColor4  =   5263559
      GradientColorOver1=   3026574
      GradientColorOver2=   3026574
      GradientColorOver3=   3026574
      GradientColorOver4=   3026574
      GradientColorDown1=   2039646
      GradientColorDown2=   2039646
      GradientColorDown3=   2039646
      GradientColorDown4=   2039646
      PicSize         =   1
      ShowFocusRect   =   0   'False
      Theme           =   4
   End
   Begin DrawSuite2022.USButton cmdAcessar 
      Height          =   495
      Left            =   8085
      TabIndex        =   3
      ToolTipText     =   "Clique aqui para acessar o sistema."
      Top             =   9240
      Visible         =   0   'False
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   873
      DibPicture      =   "frmabertura.frx":1C0F4
      Caption         =   "Entrar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   4538168
      BorderColorDown =   1249549
      BorderColorOver =   2893857
      GradientColor1  =   4538168
      GradientColor2  =   4538168
      GradientColor3  =   4538168
      GradientColor4  =   4538168
      GradientColorOver1=   2893857
      GradientColorOver2=   2893857
      GradientColorOver3=   2893857
      GradientColorOver4=   2893857
      GradientColorDown1=   1249549
      GradientColorDown2=   1249549
      GradientColorDown3=   1249549
      GradientColorDown4=   1249549
      PicSize         =   1
      ShowFocusRect   =   0   'False
      Theme           =   6
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Digite seus dados abaixo para acesso ao sistema"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   270
      Left            =   6810
      TabIndex        =   16
      Top             =   7590
      Visible         =   0   'False
      Width           =   5655
   End
   Begin DrawSuite2022.USAlphaImage Cmd_novo_local_bd 
      Height          =   405
      Left            =   11160
      TabIndex        =   5
      Top             =   8010
      Visible         =   0   'False
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   714
      Image           =   "frmabertura.frx":1EC25
      Props           =   5
   End
   Begin VB.Label L2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Senha"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   270
      Left            =   6420
      TabIndex        =   15
      Top             =   8835
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label L1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Usuário"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   270
      Left            =   6420
      TabIndex        =   14
      Top             =   8460
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label LblBanco 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Empresa"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   270
      Left            =   6420
      TabIndex        =   13
      Top             =   8040
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblLocalBanco 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   11160
      MouseIcon       =   "frmabertura.frx":39C85
      MousePointer    =   99  'Custom
      TabIndex        =   12
      ToolTipText     =   "Clique aqui para acessar as configurações do banco de dados SQL Server ( F8 )"
      Top             =   7950
      Width           =   465
   End
   Begin VB.Label lblVersaoatual 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "5.0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   435
      Left            =   3120
      TabIndex        =   10
      Top             =   690
      Width           =   495
   End
   Begin DrawSuite2022.USAlphaImage USAlphaImage1 
      Height          =   3255
      Left            =   30
      TabIndex        =   17
      Top             =   3870
      Width           =   19080
      _ExtentX        =   33655
      _ExtentY        =   5741
      Image           =   "frmabertura.frx":39F8F
      MaskColor       =   93548617
      ShadowColor     =   93548617
   End
   Begin VB.Label lblano 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright 1999 - 2018 Caprind Sistemas ®. Todos os direitos reservados."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   120
      TabIndex        =   8
      Top             =   10800
      Width           =   18960
   End
   Begin VB.Label lbldireitos 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Este programa é protegido por leis de direitos autorais (Copyright) e tratados internacionais."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   105
      TabIndex        =   7
      Top             =   11010
      Width           =   19050
   End
End
Attribute VB_Name = "frmabertura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'===============================================
'Muda cor do fundo do TEXT box
'===============================================
Public Sub MudaCor(FormX As Form)
On Error GoTo tratar_erro

Dim i As Integer

    For i = 0 To FormX.Controls.Count - 1
        If TypeOf FormX.Controls(i) Is TextBox Or TypeOf FormX.Controls(i) Is ComboBox Or TypeOf FormX.Controls(i) Is MaskEdBox Then
            With FormX.Controls(i)
            .BackColor = vbWhite
            .ForeColor = vbBlack
            End With
        End If
    Next

    If TypeOf FormX.ActiveControl Is TextBox Or TypeOf FormX.ActiveControl Is ComboBox Or TypeOf FormX.ActiveControl Is MaskEdBox Then
        FormX.ActiveControl.BackColor = RGB(128, 146, 170)
    End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub MostraLogon()
On Error GoTo tratar_erro

L1.Visible = True
L2.Visible = True
LblBanco.Visible = True
'Shape1.Visible = True
lblInfo.Visible = True

txtUsuario.Visible = True
txtSenha.Visible = True
cmbBanco.Visible = True
Cmd_novo_local_bd.Visible = True
cmdAcessar.Visible = True
Cmd_chat.Visible = True

PBLista.Visible = False
USForm1.Visible = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub


Private Sub lblLocalBanco_Click()
On Error GoTo tratar_erro

If Atualizando = True Then Exit Sub
frmOpcoesGeral2.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Timer1_Timer()
On Error GoTo tratar_erro

MostraLogon
txtUsuario.SetFocus
Timer1.Enabled = False
Timer2.Enabled = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Timer2_Timer()
On Error GoTo tratar_erro

If situacao = 0 Then
    PBLista.Value = PBLista.Value + 2
    situacao = 1
    Exit Sub
End If
If situacao = 1 Then
    situacao = 0
    Exit Sub
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub cmbBanco_Change()
On Error GoTo tratar_erro

If cmbBanco.Text = Nome_banco Then
    Var = NomeServidor
    Var1 = NomeServidor1
    Var2 = NomeServidor2
    Var3 = NomeServidor3
    
    VarE = Nome_banco
    VarE1 = Nome_banco1
    VarE2 = Nome_banco2
    VarE3 = Nome_banco3
    
    VarR = Localrel
    VarR1 = Localrel1
    VarR2 = Localrel2
    VarR3 = Localrel3
    
    VarU = Usuario_banco
    VarU1 = Usuario_banco1
    VarU2 = Usuario_banco2
    VarU3 = Usuario_banco3
    
    VarS = Senha_banco
    VarS1 = Senha_banco1
    VarS2 = Senha_banco2
    VarS3 = Senha_banco3
    
    VarLAC = LocalAntigoCaprind
    VarLAC1 = LocalAntigoCaprind1
    VarLAC2 = LocalAntigoCaprind2
    VarLAC3 = LocalAntigoCaprind3
    
    VarLNC = LocalNovoCaprind
    VarLNC1 = LocalNovoCaprind1
    VarLNC2 = LocalNovoCaprind2
    VarLNC3 = LocalNovoCaprind3
    
    VarLAG = LocalAntigoGerprod
    VarLAG1 = LocalAntigoGerprod1
    VarLAG2 = LocalAntigoGerprod2
    VarLAG3 = LocalAntigoGerprod3
    
    VarLNG = LocalNovoGerprod
    VarLNG1 = LocalNovoGerprod1
    VarLNG2 = LocalNovoGerprod2
    VarLNG3 = LocalNovoGerprod3
End If
If cmbBanco.Text = Nome_banco1 Then
    Var = NomeServidor1
    Var1 = NomeServidor2
    Var2 = NomeServidor3
    Var3 = NomeServidor
    
    VarE = Nome_banco1
    VarE1 = Nome_banco2
    VarE2 = Nome_banco3
    VarE3 = Nome_banco
    
    VarR = Localrel1
    VarR1 = Localrel2
    VarR2 = Localrel3
    VarR3 = Localrel
    
    VarU = Usuario_banco1
    VarU1 = Usuario_banco2
    VarU2 = Usuario_banco3
    VarU3 = Usuario_banco
    
    VarS = Senha_banco1
    VarS1 = Senha_banco2
    VarS2 = Senha_banco3
    VarS3 = Senha_banco
    
    VarLAC = LocalAntigoCaprind1
    VarLAC1 = LocalAntigoCaprind2
    VarLAC2 = LocalAntigoCaprind3
    VarLAC3 = LocalAntigoCaprind
    
    VarLNC = LocalNovoCaprind1
    VarLNC1 = LocalNovoCaprind2
    VarLNC2 = LocalNovoCaprind3
    VarLNC3 = LocalNovoCaprind
    
    VarLAG = LocalAntigoGerprod1
    VarLAG1 = LocalAntigoGerprod2
    VarLAG2 = LocalAntigoGerprod3
    VarLAG3 = LocalAntigoGerprod
    
    VarLNG = LocalNovoGerprod1
    VarLNG1 = LocalNovoGerprod2
    VarLNG2 = LocalNovoGerprod3
    VarLNG3 = LocalNovoGerprod
End If
If cmbBanco.Text = Nome_banco2 Then
    Var = NomeServidor2
    Var1 = NomeServidor3
    Var2 = NomeServidor
    Var3 = NomeServidor1
    
    VarE = Nome_banco2
    VarE1 = Nome_banco3
    VarE2 = Nome_banco
    VarE3 = Nome_banco1
    
    VarR = Localrel2
    VarR1 = Localrel3
    VarR2 = Localrel
    VarR3 = Localrel1
    
    VarU = Usuario_banco2
    VarU1 = Usuario_banco3
    VarU2 = Usuario_banco
    VarU3 = Usuario_banco1
    
    VarS = Senha_banco2
    VarS1 = Senha_banco3
    VarS2 = Senha_banco
    VarS3 = Senha_banco1
    
    VarLAC = LocalAntigoCaprind2
    VarLAC1 = LocalAntigoCaprind3
    VarLAC2 = LocalAntigoCaprind
    VarLAC3 = LocalAntigoCaprind1
    
    VarLNC = LocalNovoCaprind2
    VarLNC1 = LocalNovoCaprind3
    VarLNC2 = LocalNovoCaprind
    VarLNC3 = LocalNovoCaprind1
    
    VarLAG = LocalAntigoGerprod2
    VarLAG1 = LocalAntigoGerprod3
    VarLAG2 = LocalAntigoGerprod
    VarLAG3 = LocalAntigoGerprod1
    
    VarLNG = LocalNovoGerprod2
    VarLNG1 = LocalNovoGerprod3
    VarLNG2 = LocalNovoGerprod
    VarLNG3 = LocalNovoGerprod1
End If
If cmbBanco.Text = Nome_banco3 Then
    Var = NomeServidor3
    Var1 = NomeServidor
    Var2 = NomeServidor1
    Var3 = NomeServidor2
    
    VarE = Nome_banco3
    VarE1 = Nome_banco
    VarE2 = Nome_banco1
    VarE3 = Nome_banco2
    
    VarR = Localrel3
    VarR1 = Localrel
    VarR2 = Localrel1
    VarR3 = Localrel2
    
    VarU = Usuario_banco3
    VarU1 = Usuario_banco
    VarU2 = Usuario_banco1
    VarU3 = Usuario_banco2
    
    VarS = Senha_banco3
    VarS1 = Senha_banco
    VarS2 = Senha_banco1
    VarS3 = Senha_banco2
    
    VarLAC = LocalAntigoCaprind3
    VarLAC1 = LocalAntigoCaprind
    VarLAC2 = LocalAntigoCaprind1
    VarLAC3 = LocalAntigoCaprind2
    
    VarLNC = LocalNovoCaprind3
    VarLNC1 = LocalNovoCaprind
    VarLNC2 = LocalNovoCaprind1
    VarLNC3 = LocalNovoCaprind2
    
    VarLAG = LocalAntigoGerprod3
    VarLAG1 = LocalAntigoGerprod
    VarLAG2 = LocalAntigoGerprod1
    VarLAG3 = LocalAntigoGerprod2
    
    VarLNG = LocalNovoGerprod3
    VarLNG1 = LocalNovoGerprod
    VarLNG2 = LocalNovoGerprod1
    VarLNG3 = LocalNovoGerprod2
End If
NomeServidor = Var
NomeServidor1 = Var1
NomeServidor2 = Var2
NomeServidor3 = Var3

Nome_banco = VarE
Nome_banco1 = VarE1
Nome_banco2 = VarE2
Nome_banco3 = VarE3

Localrel = VarR
Localrel1 = VarR1
Localrel2 = VarR2
Localrel3 = VarR3

Usuario_banco = VarU
Usuario_banco1 = VarU1
Usuario_banco2 = VarU2
Usuario_banco3 = VarU3
    
Senha_banco = VarS
Senha_banco1 = VarS1
Senha_banco2 = VarS2
Senha_banco3 = VarS3

LocalAntigoCaprind = VarLAC
LocalAntigoCaprind1 = VarLAC1
LocalAntigoCaprind2 = VarLAC2
LocalAntigoCaprind3 = VarLAC3

LocalNovoCaprind = VarLNC
LocalNovoCaprind1 = VarLNC1
LocalNovoCaprind2 = VarLNC2
LocalNovoCaprind3 = VarLNC3

LocalAntigoGerprod = VarLAG
LocalAntigoGerprod1 = VarLAG1
LocalAntigoGerprod2 = VarLAG2
LocalAntigoGerprod3 = VarLAG3

LocalNovoGerprod = VarLNG
LocalNovoGerprod1 = VarLNG1
LocalNovoGerprod2 = VarLNG2
LocalNovoGerprod3 = VarLNG3

SaveSetting "Procam", "CaprindSQL", "NomeServidor", NomeServidor
SaveSetting "Procam", "CaprindSQL", "NomeServidor1", NomeServidor1
SaveSetting "Procam", "CaprindSQL", "NomeServidor2", NomeServidor2
SaveSetting "Procam", "CaprindSQL", "NomeServidor3", NomeServidor3
SaveSetting "Procam", "CaprindSQL", "Nome_banco", Nome_banco
SaveSetting "Procam", "CaprindSQL", "Nome_banco1", Nome_banco1
SaveSetting "Procam", "CaprindSQL", "Nome_banco2", Nome_banco2
SaveSetting "Procam", "CaprindSQL", "Nome_banco3", Nome_banco3
SaveSetting "Procam", "CaprindSQL", "Localrel", Localrel
SaveSetting "Procam", "CaprindSQL", "Localrel1", Localrel1
SaveSetting "Procam", "CaprindSQL", "Localrel2", Localrel2
SaveSetting "Procam", "CaprindSQL", "Localrel3", Localrel3
SaveSetting "Procam", "CaprindSQL", "Usuario_banco", Usuario_banco
SaveSetting "Procam", "CaprindSQL", "Usuario_banco1", Usuario_banco1
SaveSetting "Procam", "CaprindSQL", "Usuario_banco2", Usuario_banco2
SaveSetting "Procam", "CaprindSQL", "Usuario_banco3", Usuario_banco3
SaveSetting "Procam", "CaprindSQL", "Senha_banco", Senha_banco
SaveSetting "Procam", "CaprindSQL", "Senha_banco1", Senha_banco1
SaveSetting "Procam", "CaprindSQL", "Senha_banco2", Senha_banco2
SaveSetting "Procam", "CaprindSQL", "Senha_banco3", Senha_banco3
SaveSetting "Procam", "CaprindSQL", "LocalAntigoCaprind", LocalAntigoCaprind
SaveSetting "Procam", "CaprindSQL", "LocalAntigoCaprind1", LocalAntigoCaprind1
SaveSetting "Procam", "CaprindSQL", "LocalAntigoCaprind2", LocalAntigoCaprind2
SaveSetting "Procam", "CaprindSQL", "LocalAntigoCaprind3", LocalAntigoCaprind3
SaveSetting "Procam", "CaprindSQL", "LocalNovoCaprind", LocalNovoCaprind
SaveSetting "Procam", "CaprindSQL", "LocalNovoCaprind1", LocalNovoCaprind1
SaveSetting "Procam", "CaprindSQL", "LocalNovoCaprind2", LocalNovoCaprind2
SaveSetting "Procam", "CaprindSQL", "LocalNovoCaprind3", LocalNovoCaprind3
SaveSetting "Procam", "CaprindSQL", "LocalAntigoGerprod", LocalAntigoGerprod
SaveSetting "Procam", "CaprindSQL", "LocalAntigoGerprod1", LocalAntigoGerprod1
SaveSetting "Procam", "CaprindSQL", "LocalAntigoGerprod2", LocalAntigoGerprod2
SaveSetting "Procam", "CaprindSQL", "LocalAntigoGerprod3", LocalAntigoGerprod3
SaveSetting "Procam", "CaprindSQL", "LocalNovoGerprod", LocalNovoGerprod
SaveSetting "Procam", "CaprindSQL", "LocalNovoGerprod1", LocalNovoGerprod1
SaveSetting "Procam", "CaprindSQL", "LocalNovoGerprod2", LocalNovoGerprod2
SaveSetting "Procam", "CaprindSQL", "LocalNovoGerprod3", LocalNovoGerprod3

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub cmbbanco_Click()
On Error GoTo tratar_erro

If cmbBanco.Text = Nome_banco Then
    Var = NomeServidor
    Var1 = NomeServidor1
    Var2 = NomeServidor2
    Var3 = NomeServidor3
    
    VarE = Nome_banco
    VarE1 = Nome_banco1
    VarE2 = Nome_banco2
    VarE3 = Nome_banco3
    
    VarR = Localrel
    VarR1 = Localrel1
    VarR2 = Localrel2
    VarR3 = Localrel3
    
    VarU = Usuario_banco
    VarU1 = Usuario_banco1
    VarU2 = Usuario_banco2
    VarU3 = Usuario_banco3
    
    VarS = Senha_banco
    VarS1 = Senha_banco1
    VarS2 = Senha_banco2
    VarS3 = Senha_banco3
    
    VarLAC = LocalAntigoCaprind
    VarLAC1 = LocalAntigoCaprind1
    VarLAC2 = LocalAntigoCaprind2
    VarLAC3 = LocalAntigoCaprind3
    
    VarLNC = LocalNovoCaprind
    VarLNC1 = LocalNovoCaprind1
    VarLNC2 = LocalNovoCaprind2
    VarLNC3 = LocalNovoCaprind3
    
    VarLAG = LocalAntigoGerprod
    VarLAG1 = LocalAntigoGerprod1
    VarLAG2 = LocalAntigoGerprod2
    VarLAG3 = LocalAntigoGerprod3
    
    VarLNG = LocalNovoGerprod
    VarLNG1 = LocalNovoGerprod1
    VarLNG2 = LocalNovoGerprod2
    VarLNG3 = LocalNovoGerprod3
End If
If cmbBanco.Text = Nome_banco1 Then
    Var = NomeServidor1
    Var1 = NomeServidor2
    Var2 = NomeServidor3
    Var3 = NomeServidor
    
    VarE = Nome_banco1
    VarE1 = Nome_banco2
    VarE2 = Nome_banco3
    VarE3 = Nome_banco
    
    VarR = Localrel1
    VarR1 = Localrel2
    VarR2 = Localrel3
    VarR3 = Localrel
    
    VarU = Usuario_banco1
    VarU1 = Usuario_banco2
    VarU2 = Usuario_banco3
    VarU3 = Usuario_banco
    
    VarS = Senha_banco1
    VarS1 = Senha_banco2
    VarS2 = Senha_banco3
    VarS3 = Senha_banco
    
    VarLAC = LocalAntigoCaprind1
    VarLAC1 = LocalAntigoCaprind2
    VarLAC2 = LocalAntigoCaprind3
    VarLAC3 = LocalAntigoCaprind
    
    VarLNC = LocalNovoCaprind1
    VarLNC1 = LocalNovoCaprind2
    VarLNC2 = LocalNovoCaprind3
    VarLNC3 = LocalNovoCaprind
    
    VarLAG = LocalAntigoGerprod1
    VarLAG1 = LocalAntigoGerprod2
    VarLAG2 = LocalAntigoGerprod3
    VarLAG3 = LocalAntigoGerprod
    
    VarLNG = LocalNovoGerprod1
    VarLNG1 = LocalNovoGerprod2
    VarLNG2 = LocalNovoGerprod3
    VarLNG3 = LocalNovoGerprod
End If
If cmbBanco.Text = Nome_banco2 Then
    Var = NomeServidor2
    Var1 = NomeServidor3
    Var2 = NomeServidor
    Var3 = NomeServidor1
    
    VarE = Nome_banco2
    VarE1 = Nome_banco3
    VarE2 = Nome_banco
    VarE3 = Nome_banco1
    
    VarR = Localrel2
    VarR1 = Localrel3
    VarR2 = Localrel
    VarR3 = Localrel1
    
    VarU = Usuario_banco2
    VarU1 = Usuario_banco3
    VarU2 = Usuario_banco
    VarU3 = Usuario_banco1
    
    VarS = Senha_banco2
    VarS1 = Senha_banco3
    VarS2 = Senha_banco
    VarS3 = Senha_banco1
    
    VarLAC = LocalAntigoCaprind2
    VarLAC1 = LocalAntigoCaprind3
    VarLAC2 = LocalAntigoCaprind
    VarLAC3 = LocalAntigoCaprind1
    
    VarLNC = LocalNovoCaprind2
    VarLNC1 = LocalNovoCaprind3
    VarLNC2 = LocalNovoCaprind
    VarLNC3 = LocalNovoCaprind1
    
    VarLAG = LocalAntigoGerprod2
    VarLAG1 = LocalAntigoGerprod3
    VarLAG2 = LocalAntigoGerprod
    VarLAG3 = LocalAntigoGerprod1
    
    VarLNG = LocalNovoGerprod2
    VarLNG1 = LocalNovoGerprod3
    VarLNG2 = LocalNovoGerprod
    VarLNG3 = LocalNovoGerprod1
End If
If cmbBanco.Text = Nome_banco3 Then
    Var = NomeServidor3
    Var1 = NomeServidor
    Var2 = NomeServidor1
    Var3 = NomeServidor2
    
    VarE = Nome_banco3
    VarE1 = Nome_banco
    VarE2 = Nome_banco1
    VarE3 = Nome_banco2
    
    VarR = Localrel3
    VarR1 = Localrel
    VarR2 = Localrel1
    VarR3 = Localrel2
    
    VarU = Usuario_banco3
    VarU1 = Usuario_banco
    VarU2 = Usuario_banco1
    VarU3 = Usuario_banco2
    
    VarS = Senha_banco3
    VarS1 = Senha_banco
    VarS2 = Senha_banco1
    VarS3 = Senha_banco2
    
    VarLAC = LocalAntigoCaprind3
    VarLAC1 = LocalAntigoCaprind
    VarLAC2 = LocalAntigoCaprind1
    VarLAC3 = LocalAntigoCaprind2
    
    VarLNC = LocalNovoCaprind3
    VarLNC1 = LocalNovoCaprind
    VarLNC2 = LocalNovoCaprind1
    VarLNC3 = LocalNovoCaprind2
    
    VarLAG = LocalAntigoGerprod3
    VarLAG1 = LocalAntigoGerprod
    VarLAG2 = LocalAntigoGerprod1
    VarLAG3 = LocalAntigoGerprod2
    
    VarLNG = LocalNovoGerprod3
    VarLNG1 = LocalNovoGerprod
    VarLNG2 = LocalNovoGerprod1
    VarLNG3 = LocalNovoGerprod2
End If
NomeServidor = Var
NomeServidor1 = Var1
NomeServidor2 = Var2
NomeServidor3 = Var3

Nome_banco = VarE
Nome_banco1 = VarE1
Nome_banco2 = VarE2
Nome_banco3 = VarE3

Localrel = VarR
Localrel1 = VarR1
Localrel2 = VarR2
Localrel3 = VarR3

Usuario_banco = VarU
Usuario_banco1 = VarU1
Usuario_banco2 = VarU2
Usuario_banco3 = VarU3
    
Senha_banco = VarS
Senha_banco1 = VarS1
Senha_banco2 = VarS2
Senha_banco3 = VarS3

LocalAntigoCaprind = VarLAC
LocalAntigoCaprind1 = VarLAC1
LocalAntigoCaprind2 = VarLAC2
LocalAntigoCaprind3 = VarLAC3

LocalNovoCaprind = VarLNC
LocalNovoCaprind1 = VarLNC1
LocalNovoCaprind2 = VarLNC2
LocalNovoCaprind3 = VarLNC3

LocalAntigoGerprod = VarLAG
LocalAntigoGerprod1 = VarLAG1
LocalAntigoGerprod2 = VarLAG2
LocalAntigoGerprod3 = VarLAG3

LocalNovoGerprod = VarLNG
LocalNovoGerprod1 = VarLNG1
LocalNovoGerprod2 = VarLNG2
LocalNovoGerprod3 = VarLNG3

SaveSetting "Procam", "CaprindSQL", "NomeServidor", NomeServidor
SaveSetting "Procam", "CaprindSQL", "NomeServidor1", NomeServidor1
SaveSetting "Procam", "CaprindSQL", "NomeServidor2", NomeServidor2
SaveSetting "Procam", "CaprindSQL", "NomeServidor3", NomeServidor3
SaveSetting "Procam", "CaprindSQL", "Nome_banco", Nome_banco
SaveSetting "Procam", "CaprindSQL", "Nome_banco1", Nome_banco1
SaveSetting "Procam", "CaprindSQL", "Nome_banco2", Nome_banco2
SaveSetting "Procam", "CaprindSQL", "Nome_banco3", Nome_banco3
SaveSetting "Procam", "CaprindSQL", "Localrel", Localrel
SaveSetting "Procam", "CaprindSQL", "Localrel1", Localrel1
SaveSetting "Procam", "CaprindSQL", "Localrel2", Localrel2
SaveSetting "Procam", "CaprindSQL", "Localrel3", Localrel3
SaveSetting "Procam", "CaprindSQL", "Usuario_banco", Usuario_banco
SaveSetting "Procam", "CaprindSQL", "Usuario_banco1", Usuario_banco1
SaveSetting "Procam", "CaprindSQL", "Usuario_banco2", Usuario_banco2
SaveSetting "Procam", "CaprindSQL", "Usuario_banco3", Usuario_banco3
SaveSetting "Procam", "CaprindSQL", "Senha_banco", Senha_banco
SaveSetting "Procam", "CaprindSQL", "Senha_banco1", Senha_banco1
SaveSetting "Procam", "CaprindSQL", "Senha_banco2", Senha_banco2
SaveSetting "Procam", "CaprindSQL", "Senha_banco3", Senha_banco3
SaveSetting "Procam", "CaprindSQL", "LocalAntigoCaprind", LocalAntigoCaprind
SaveSetting "Procam", "CaprindSQL", "LocalAntigoCaprind1", LocalAntigoCaprind1
SaveSetting "Procam", "CaprindSQL", "LocalAntigoCaprind2", LocalAntigoCaprind2
SaveSetting "Procam", "CaprindSQL", "LocalAntigoCaprind3", LocalAntigoCaprind3
SaveSetting "Procam", "CaprindSQL", "LocalNovoCaprind", LocalNovoCaprind
SaveSetting "Procam", "CaprindSQL", "LocalNovoCaprind1", LocalNovoCaprind1
SaveSetting "Procam", "CaprindSQL", "LocalNovoCaprind2", LocalNovoCaprind2
SaveSetting "Procam", "CaprindSQL", "LocalNovoCaprind3", LocalNovoCaprind3
SaveSetting "Procam", "CaprindSQL", "LocalAntigoGerprod", LocalAntigoGerprod
SaveSetting "Procam", "CaprindSQL", "LocalAntigoGerprod1", LocalAntigoGerprod1
SaveSetting "Procam", "CaprindSQL", "LocalAntigoGerprod2", LocalAntigoGerprod2
SaveSetting "Procam", "CaprindSQL", "LocalAntigoGerprod3", LocalAntigoGerprod3
SaveSetting "Procam", "CaprindSQL", "LocalNovoGerprod", LocalNovoGerprod
SaveSetting "Procam", "CaprindSQL", "LocalNovoGerprod1", LocalNovoGerprod1
SaveSetting "Procam", "CaprindSQL", "LocalNovoGerprod2", LocalNovoGerprod2
SaveSetting "Procam", "CaprindSQL", "LocalNovoGerprod3", LocalNovoGerprod3

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Cmd_chat_Click()
On Error GoTo tratar_erro
Dim Comando As String
Dim appName As String
'================================================================
'Fecha Team Viewer se estiver aberto
'================================================================
appName = "TeamViewer.exe"
Comando = "TASKKILL -F -IM " & appName
Shell Comando
'================================================================

caminho = App.Path & "\TeamViewerQS.exe"

If USMsgBox("Deseja realmente solicitar uma conexão remota em sua máquina?", vbYesNo, "CAPRIND v5.0") = vbNo Then
    Exit Sub
End If

'================================================================
'Verifica se tem internet disponível
'================================================================
If IsInternetOnline = True Then
    If FileOrDirExists(caminho) = False Then
        Atualizacao_TeamViewerQS = True
        Frm_atualizacao_sistema.Show 1
    Else
        ProcAbrirArquivo (caminho)
    End If
Else
    If IsInternetOnline = False Then
        MsgTexto = "não foi encontrado conexão com a internet"
    Else
        MsgTexto = "no momento estamos sem conexão com nosso servidor, favor tentar mais tarde"
    End If
        
    USMsgBox ("Não é permitido abrir este módulo, pois " & MsgTexto & "."), vbExclamation, "CAPRIND v5.0"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub cmdAcessar_Click()
On Error GoTo tratar_erro


'ProcUsuariosLiberados

'If FunVerifAtualizacaoObrigatoria(False, True) = False Then
    If Atualizando = True Then Exit Sub
    If cmbBanco.Text <> "" Then
        If FunAbreBD() = False Then
        'If FunAbreBDWEB() = False Then
            USMsgBox ("Não foi possível conectar com o banco de dados do " & NomeCampo & ", favor verificar as configurações do sistema."), vbExclamation, "CAPRIND v5.0"
            frmOpcoesGeral2.Show 1
            Exit Sub
        Else
            ProcEntrar
        End If
    Else
        USMsgBox ("É necessário configurar os dados para conexão com o banco de dados."), vbExclamation, "CAPRIND v5.0"
        frmOpcoesGeral2.Show 1
    End If
'End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub cmdSair_Click()
On Error GoTo tratar_erro

If Atualizando = True Then Exit Sub
ProcSair

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Sub ProcEntrar()
On Error GoTo tratar_erro


If cmbBanco.Text = "" Then
    USMsgBox ("Informe o nome do banco de dados antes de efetuar o logon."), vbExclamation, "CAPRIND v5.0"
    cmbBanco.SetFocus
    Exit Sub
End If
If txtSenha.Text = "" Or txtUsuario.Text = "" Then
    USMsgBox ("Informe o usuário e senha antes de efetuar o logon."), vbExclamation, "CAPRIND v5.0"
    If txtUsuario.Text = "" Then
        txtUsuario.SetFocus
        Exit Sub
    End If
    If txtSenha.Text = "" Then
        txtSenha.SetFocus
        Exit Sub
    End If
End If

Contador = 0

'Verifica se a versão do exe é menor que a do banco de dados
Quant = ReturnNumbersOnly(App.Major & "." & App.Minor & "." & App.Revision & ".txt")
Quant = ReturnNumbersOnly(lblVersaoatual.Caption)
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select Versao from Versao where Versao IS NOT NULL", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    VersaoNova = ReturnNumbersOnly(IIf(TBAbrir!versao = "", 0, TBAbrir!versao))
    If Quant < VersaoNova Then
        'usMsgbox ("O sistema está desatualizado e será encerrado."), vbExclamation, "CAPRIND v5.0"
        'TBAbrir.Close
        'End
    End If
End If


ProcVerificaInternet True, False 'Verifica conexão com a internet

If TemInternet = True And ErroDriverMYSQL = False Then
    'Verifica se tem versão atualização disponível para baixar
    FunAbreBDSite
    If ConexaoMySql.State = 1 Then
        Set TBMySQL = New ADODB.Recordset
        TBMySQL.Open "Select * From Atualizacao_liberada", ConexaoMySql, adOpenKeyset, adLockOptimistic, adCmdText
        If TBMySQL.EOF = False Then
            VersaoNova = ReturnNumbersOnly(IIf(TBMySQL!versao = "", 0, TBMySQL!versao))
            If Quant < VersaoNova Then
                USMsgBox ("Existe uma atualização disponível para ser baixada, solicite ao administrador do sistema."), vbInformation, "CAPRIND v5.0"
                TBMySQL.Close
            End If
            If Quant > VersaoNova Then
                versaoatualtexto = Format(VersaoNova, "@.@.@@@")
                versaonovatexto = Format(Quant, "@.@.@@@")
                
                Set TBAbrir = CreateObject("adodb.recordset")
                   TBAbrir.Open "Select cnpj from Empresa", Conexao, adOpenKeyset, adLockOptimistic
                   If TBAbrir.EOF = False Then
                       Do While TBAbrir.EOF = False
                           If ConexaoMySql.State = 1 Then ConexaoMySql.Execute "Update Clientes Set Versao_anterior = '" & versaoatualtexto & "', Versao = '" & versaonovatexto & "', Data_atualizacao = '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "' where CNPJ = '" & TBAbrir!CNPJ & "'"
                           TBAbrir.MoveNext
                       Loop
                   End If
                   TBAbrir.Close
            End If
        End If
    End If
End If

Texto = ""
Numero = 0
Numero1 = Len(NomeServidor)
Hora = 0
If Numero1 <> 1 Then
    Do While Numero1 <> 0
        If Texto = "\" Then GoTo Pula
        Texto = Left(NomeServidor, (Numero + 1))
        Texto = Right(Texto, Len(Texto) - Numero)
        Numero = Numero + 1
        Numero1 = Numero1 - 1
    Loop
End If
Pula:
Familiatext = Left(NomeServidor, Numero - 1)
Dataini = Format(FunHoraServidor("\\" & Familiatext), "dd/mm/yyyy")
If Dataini <> Date Then
    If Dataini > Date Then MsgTexto = "maior" Else MsgTexto = "menor"
    USMsgBox ("A data do computador está " & MsgTexto & " que a data do servidor, favor arrumar antes de logar no sistema."), vbExclamation, "CAPRIND v5.0"
    End
End If

ProcLogonOutSemUtilizacao 'Verifica e apaga logon com a data menor que a atual

If DS.FileOrDirExists(Localrel) = False Then
    If USMsgBox("Não foi encontrado o caminho " & Localrel & " onde está armazenado os relatórios, se efetuar o login desta forma os relatórios não vão funcionar. " & vbCrLf & "Deseja prosseguir mesmo assim?", vbYesNo, "CAPRIND v5.0") = vbNo Then End
End If

Set TBUsuarios = CreateObject("adodb.recordset")
TBUsuarios.Open "Select * FROM Usuarios WHERE Usuario = '" & txtUsuario & "' AND Senha = '" & txtSenha & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBUsuarios.EOF = False Then
    pubUsuario = TBUsuarios!Usuario
    pubIDUsuario = TBUsuarios!IDUsuario
    pubNome = TBUsuarios!Nome
    pubSetor = TBUsuarios!Setor
    pubSenha = TBUsuarios!Senha
    pubEmail = IIf(IsNull(TBUsuarios!Email), "", TBUsuarios!Email)
    If TBUsuarios!Bloqueado = True Then
        USMsgBox ("Não será possível efetuar o logon, pois o usuário " & txtUsuario.Text & " está bloqueado."), vbExclamation, "CAPRIND v5.0"
        TBUsuarios.Close
        Exit Sub
    End If
    If IsNull(TBUsuarios!DtExpiracao) = False And TBUsuarios!DtExpiracao <> "" Then
        If TBUsuarios!DtExpiracao <= Date Then
            USMsgBox ("A data de utilização do Caprind para o usuario " & pubUsuario & " venceu em " & Format(TBUsuarios!DtExpiracao, "dd/mm/yy") & "."), vbExclamation, "CAPRIND v5.0"
            TBUsuarios.Close
            Exit Sub
        End If
    End If
Else

    With frmabertura
        .txtUsuario.Enabled = True
        .txtSenha.Enabled = True
        .cmbBanco.Enabled = True
        .Cmd_novo_local_bd.Enabled = True
        .Cmd_chat.Enabled = True
        .cmdAcessar.Enabled = True
    End With

    USMsgBox ("Nome de usuário ou senha inválida."), vbExclamation, "CAPRIND v5.0"
    txtSenha.SetFocus
    TBUsuarios.Close
    Exit Sub
End If
TBUsuarios.Close

'Criar este codigo para caminho dos arquivos avi
If Localrel <> "" Then caminho = Localrel

TemInternet = DS.IsInternetOnline
'======================================================================
' Importante, verificação de cliente na base e dados MySQL
'======================================================================
If TemInternet = True And ErroDriverMYSQL = False Then
    FunValidarCliente 'Valida cliente com internet
Else
    FunValidarClienteSemInternet 'Valida cliente sem internet
End If
'======================================================================

FunLogonIn True
Unload Me
frmMDI.Show
frmMDI.Caption = VarE & " - CAPRIND - v" & App.Major & "." & App.Minor & "." & App.Revision & " - Menu Principal"
ProcMontaStatusBar

Unload frmabertura

'Pega ID da empresa cadastrada
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * FROM Empresa", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
IDEmp = TBAbrir!CODIGO
IDempresa = TBAbrir!CODIGO
CNPJ_Empresa = TBAbrir!CNPJ
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub


Private Sub ProcInfLicenca()
On Error GoTo tratar_erro

'qtde_solicitada = ""
'Mensagem1:
'qtde_solicitada = InputBox("Favor informar o número de licença.")
'If qtde_solicitada = "" Then Exit Sub
'If IsNumeric(qtde_solicitada) = False Then
'    usMsgbox ("Só é permitido número neste campo."), vbExclamation, "CAPRIND v5.0"
'    GoTo Mensagem1
'End If
'Qtde = qtde_solicitada
'If Qtde <= 0 Then
'    usMsgbox ("So é permitido número maior que 0."), vbExclamation, "CAPRIND v5.0"
'    GoTo Mensagem1
'End If
'Set TBExecucao = CreateObject("adodb.recordset")
'TBExecucao.Open "Select * FROM Keylock", Conexao, adOpenKeyset, adLockOptimistic
'If TBExecucao.EOF = True Then TBExecucao.AddNew
'TBExecucao!Licenca = Qtde
'TBExecucao.Update
'TBExecucao.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Sub ProcSair()
On Error GoTo tratar_erro

If USMsgBox("Deseja realmente abandonar o Caprind?", vbYesNo, "Caprind") = vbYes Then
    End
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Form_Activate()
On Error GoTo tratar_erro


CompactarRepararBanco = False
With cmbBanco
    If Nome_banco <> "" Then
        .Text = Nome_banco
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro
   
Select Case KeyCode
    Case vbKeyF8: lblLocalBanco_Click
    'Case vbKeyF11: InfLicenca = True
    Case vbKeyF9: Cmd_chat_Click
    Case vbKeyF11:
        If Atualizando = True Then Exit Sub
        If USMsgBox("Deseja realmente reiniciar a utilização do banco de dados?", vbYesNo, "CAPRIND v5.0") = vbYes Then
            If InputBox("Informe a primeira senha que tenha pemissão para realizar esta operação.") = "280362" Then
                If InputBox("Informe a segunda senha que tenha pemissão para realizar esta operação.") = "pro0902loc" Then ProcApagarDadosBD
            End If
        End If
    Case vbKeyReturn:
    
        If cmdAcessar.Enabled = True Then
        cmdAcessar_Click
        End If
        
    'Case vbKeyF1: Cmd_ajuda_Click
    Case vbKeyEscape: cmdSair_Click
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro
Revisao = App.Revision
lblVersaoatual.Caption = "Versão " & App.Major & "." & App.Minor & "." & IIf(Len(Revisao) = 3, Revisao, "0" & Revisao)

lblano.Caption = "Copyright 1999 - " & Year(Date) & " Caprind Sistemas ®. Todos os direitos reservados."

CompactarRepararBanco = False

'===========================================
' VERIFICA SE O SISTEMA ESTÁ ATUALIZADO
'===========================================
    ProcVerifAtualizacao
'===========================================

Me.BackColor = RGB(53, 81, 129)
PBLista.BarColor2 = RGB(53, 81, 129)
Me.Refresh
PBLista.BarColor1 = RGB(53, 81, 129)
PBLista.BorderColor = RGB(53, 81, 129)

USForm1.BackColor = RGB(53, 81, 129)
USForm1.BackColorDisabled = RGB(53, 81, 129)
USForm1.BorderColor = RGB(53, 81, 129)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcUsuariosLiberados()
On Error GoTo tratar_erro
Dim Qtlicencas_caprind_Liberadas As Integer
Dim Qtlicencas_gerprod_Liberadas As Integer
Dim Qtlicencas_caprind_Cadastradas As Integer
Dim Qtlicencas_gerprod_Cadastradas As Integer


'=============================================================
'Verifica número de usuarios liberados na nuvem
'=============================================================

If IsInternetOnline = True And ErroDriverMYSQL = False Then
    Set TBComponente = CreateObject("adodb.recordset")
    TBComponente.Open "Select CNPJ from Empresa", Conexao, adOpenKeyset, adLockOptimistic
    If TBComponente.EOF = False Then
        FunAbreBDSite
        If ConexaoMySql.State = 1 Then
            Set TBMySQL = New ADODB.Recordset
            TBMySQL.Open "Select Licencas, Licencas_gerprod, Modulo From Clientes Where CNPJ = '" & TBComponente!CNPJ & "'", ConexaoMySql, adOpenKeyset, adLockOptimistic, adCmdText
            With TBMySQL
                If .EOF = False Then

                    If IsNull(.Fields!Licencas) = False And .Fields!Licencas <> "" Then
                        Qtlicencas_caprind_Liberadas = .Fields!Licencas
                    End If
                    
                    If IsNull(.Fields!Licencas_gerprod) = False And .Fields!Licencas_gerprod <> "" Then
                        Qtlicencas_gerprod_Liberadas = .Fields!Licencas_gerprod
                    End If
                    
                End If
            End With
        End If
    End If
    TBComponente.Close
End If

'=============================================================
'Verifica número de usuarios liberados na nuvem
'=============================================================
StrSql = "select count(Distinct US.idusuario) as Licencas_Caprind from Usuarios US inner join Acessos AC on AC.IDUsuario = US.IDUsuario where US.Bloqueado = 'False' and US.Usuario <> 'Procam' and US.Usuario <> 'ADMIN'"
    Set TBComponente = CreateObject("adodb.recordset")
    TBComponente.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
    If TBComponente.EOF = False Then
        If IsNull(TBComponente!Licencas_caprind) = False And TBComponente!Licencas_caprind <> "" Then
            Qtlicencas_caprind_Cadastradas = TBComponente!Licencas_caprind
        End If
    End If
    TBComponente.Close

If Qtlicencas_caprind_Liberadas < Qtlicencas_caprind_Cadastradas Then
USMsgBox "Atenção, numero de usuarios cadastrados e liberados no sistema foi excedido!" & vbCrLf & "Usuarios liberados : " & Qtlicencas_caprind_Liberadas & vbCrLf & "Usuarios cadastrados : " & Qtlicencas_caprind_Cadastradas & vbCrLf & "Avise ao administrador do sistema! ", vbCritical, "CAPRIND v5.0"

End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub

End Sub
Private Sub ProcApagarDadosBD()
On Error GoTo tratar_erro

'Data da última alteração: 24/01/18
FunAbreBD
Conexao.Execute "DELETE from Acessos Where IDUsuario <> 1 and IDUsuario <> 2"
Conexao.Execute "UPDATE Acessos Set Data = '" & Date & "', responsavel = 'PROCAM'"
Conexao.Execute "UPDATE CodigoDesc Set Data = '" & Date & "', responsavel = 'PROCAM', Obs_status = NULL, Resp_status = NULL"
Conexao.Execute "UPDATE Usuarios Set Usuario = 'PROCAM', Codigo = NULL, Senha = 'pro0902loc', Nome = 'PROCAM', Setor = 'ADMINISTRADOR', Data = '" & Date & "', responsavel = 'PROCAM' where Usuario = 'CLIENTE' or Usuario = 'PROCAM'"
Conexao.Execute "UPDATE Usuarios Set Codigo = NULL, Setor = 'ADMINISTRADOR', Data = '" & Date & "', responsavel = 'PROCAM' where Usuario = 'ADMIN'"
Conexao.Execute "DELETE Usuarios where Usuario <> 'PROCAM' and Usuario <> 'ADMIN'"
Conexao.Execute "DBCC CheckIdent('Usuarios',Reseed,1)"
Conexao.Execute "UPDATE Unidade_Medida Set Data = '" & Date & "', responsavel = 'PROCAM'"
Conexao.Execute "UPDATE Moeda Set Data = '" & Date & "', responsavel = 'PROCAM'"
Conexao.Execute "UPDATE Regioes Set Data = '" & Date & "', responsavel = 'PROCAM'"
Conexao.Execute "UPDATE Tbl_familia Set Data = '" & Date & "', responsavel = 'PROCAM'"
Conexao.Execute "UPDATE Tbl_FormaPagto Set Data = '" & Date & "', responsavel = 'PROCAM'"
Conexao.Execute "UPDATE tbl_ClassificacaoFiscal Set Data = '" & Date & "', responsavel = 'PROCAM', DtValidacao = NULL, RespValidacao = NULL"
Conexao.Execute "UPDATE tbl_NaturezaOperacao Set Data = '" & Date & "', responsavel = 'PROCAM', DtValidacao = NULL, RespValidacao = NULL"
Conexao.Execute "Truncate Table Afericao"
Conexao.Execute "DBCC CheckIdent('Afericao',Reseed,1)"
Conexao.Execute "Truncate Table Backup_configuracoes"
Conexao.Execute "DBCC CheckIdent('Backup_configuracoes',Reseed,1)"
Conexao.Execute "Truncate Table Backup_historico"
Conexao.Execute "DBCC CheckIdent('Backup_historico',Reseed,1)"
Conexao.Execute "Truncate Table Boleto_instrucoes"
Conexao.Execute "DBCC CheckIdent('Boleto_instrucoes',Reseed,1)"
Conexao.Execute "Truncate Table CadmaqTurnos"
Conexao.Execute "DBCC CheckIdent('CadmaqTurnos',Reseed,1)"
Conexao.Execute "Truncate Table CadMaquinas"
Conexao.Execute "DBCC CheckIdent('CadMaquinas',Reseed,1)"
Conexao.Execute "Truncate Table CadMaquinas_acessorios"
Conexao.Execute "DBCC CheckIdent('CadMaquinas_acessorios',Reseed,1)"
Conexao.Execute "Truncate Table CadMaquinas_grupo"
Conexao.Execute "DBCC CheckIdent('CadMaquinas_grupo',Reseed,1)"
Conexao.Execute "Truncate Table CadMaquinas_instrucoes"
Conexao.Execute "DBCC CheckIdent('CadMaquinas_instrucoes',Reseed,1)"
Conexao.Execute "Truncate Table CadMaquinas_Monitor"
Conexao.Execute "DBCC CheckIdent('CadMaquinas_Monitor',Reseed,1)"
Conexao.Execute "Truncate Table CargaMaq_Total"
Conexao.Execute "DBCC CheckIdent('CargaMaq_Total',Reseed,1)"
Conexao.Execute "Truncate Table CargaMaq_TotalMaq"
Conexao.Execute "DBCC CheckIdent('CargaMaq_TotalMaq',Reseed,1)"
Conexao.Execute "Truncate Table CC_realizado"
Conexao.Execute "DBCC CheckIdent('CC_realizado',Reseed,1)"
Conexao.Execute "Truncate Table Certificado"
Conexao.Execute "DBCC CheckIdent('Certificado',Reseed,1)"
Conexao.Execute "Truncate Table Certificado_Analise"
Conexao.Execute "DBCC CheckIdent('Certificado_Analise',Reseed,1)"
Conexao.Execute "Truncate Table Certificado_qualidade"
Conexao.Execute "DBCC CheckIdent('Certificado_qualidade',Reseed,1)"
Conexao.Execute "Truncate Table Certificado_Quimica"
Conexao.Execute "DBCC CheckIdent('Certificado_Quimica',Reseed,1)"
Conexao.Execute "Truncate Table Certificados"
Conexao.Execute "DBCC CheckIdent('Certificados',Reseed,1)"
Conexao.Execute "Truncate Table CFI"
Conexao.Execute "DBCC CheckIdent('CFI',Reseed,1)"
Conexao.Execute "Truncate Table CFI_Itens"
Conexao.Execute "DBCC CheckIdent('CFI_Itens',Reseed,1)"
Conexao.Execute "Truncate Table Cheques_Cancelados"
Conexao.Execute "DBCC CheckIdent('Cheques_Cancelados',Reseed,1)"
Conexao.Execute "Truncate Table Cheques_Relatorios"
Conexao.Execute "DBCC CheckIdent('Cheques_Relatorios',Reseed,1)"
Conexao.Execute "Truncate Table Clientes"
Conexao.Execute "DBCC CheckIdent('Clientes',Reseed,1)"
Conexao.Execute "Truncate Table clientes_cobranca"
Conexao.Execute "DBCC CheckIdent('clientes_cobranca',Reseed,1)"
Conexao.Execute "Truncate Table Clientes_Contatos"
Conexao.Execute "DBCC CheckIdent('Clientes_Contatos',Reseed,1)"
Conexao.Execute "Truncate Table Clientes_DadosComerciais"
Conexao.Execute "DBCC CheckIdent('Clientes_DadosComerciais',Reseed,1)"
Conexao.Execute "Truncate Table clientes_entrega"
Conexao.Execute "DBCC CheckIdent('clientes_entrega',Reseed,1)"
Conexao.Execute "Truncate Table Clientes_grupos"
Conexao.Execute "DBCC CheckIdent('Clientes_grupos',Reseed,1)"
Conexao.Execute "Truncate Table Clientes_Impostos"
Conexao.Execute "DBCC CheckIdent('Clientes_Impostos',Reseed,1)"
Conexao.Execute "Truncate Table Compras_comercial"
Conexao.Execute "DBCC CheckIdent('Compras_comercial',Reseed,1)"
Conexao.Execute "Truncate Table Compras_Cotacao"
Conexao.Execute "DBCC CheckIdent('Compras_Cotacao',Reseed,1)"
Conexao.Execute "Truncate Table Compras_fornecedores"
Conexao.Execute "DBCC CheckIdent('Compras_fornecedores',Reseed,1)"
Conexao.Execute "Truncate Table Compras_fornecedores_banco"
Conexao.Execute "DBCC CheckIdent('Compras_fornecedores_banco',Reseed,1)"
Conexao.Execute "Truncate Table Compras_fornecedores_familia"
Conexao.Execute "DBCC CheckIdent('Compras_fornecedores_familia',Reseed,1)"
Conexao.Execute "Truncate Table Compras_fornecedores_segmentos"
Conexao.Execute "DBCC CheckIdent('Compras_fornecedores_segmentos',Reseed,1)"
Conexao.Execute "Truncate Table Compras_pedido"
Conexao.Execute "DBCC CheckIdent('Compras_pedido',Reseed,1)"
Conexao.Execute "Truncate Table Compras_pedido_lista"
Conexao.Execute "DBCC CheckIdent('Compras_pedido_lista',Reseed,1)"
Conexao.Execute "Truncate Table Compras_pedido_lista_custo"
Conexao.Execute "DBCC CheckIdent('Compras_pedido_lista_custo',Reseed,1)"
Conexao.Execute "Truncate Table Compras_pedido_lista_empenhos"
Conexao.Execute "DBCC CheckIdent('Compras_pedido_lista_empenhos',Reseed,1)"
Conexao.Execute "Truncate Table Compras_programa"
Conexao.Execute "DBCC CheckIdent('Compras_programa',Reseed,1)"
Conexao.Execute "Truncate Table Compras_programa_item"
Conexao.Execute "DBCC CheckIdent('Compras_programa_item',Reseed,1)"
Conexao.Execute "Truncate Table Compras_programacao"
Conexao.Execute "DBCC CheckIdent('Compras_programacao',Reseed,1)"
Conexao.Execute "Truncate Table Compras_recebimento"
Conexao.Execute "DBCC CheckIdent('Compras_recebimento',Reseed,1)"
Conexao.Execute "Truncate Table Compras_Recebimento_Relatorios"
Conexao.Execute "DBCC CheckIdent('Compras_Recebimento_Relatorios',Reseed,1)"
Conexao.Execute "Truncate Table Compras_requisicao"
Conexao.Execute "DBCC CheckIdent('Compras_requisicao',Reseed,1)"
Conexao.Execute "Truncate Table Contatos_fornecedor"
Conexao.Execute "DBCC CheckIdent('Contatos_fornecedor',Reseed,1)"
Conexao.Execute "Truncate Table Controle_projetos"
Conexao.Execute "DBCC CheckIdent('Controle_projetos',Reseed,1)"
Conexao.Execute "Truncate Table Controle_projetos_historico"
Conexao.Execute "DBCC CheckIdent('Controle_projetos_historico',Reseed,1)"
Conexao.Execute "Truncate Table Cotacao_fornecedor"
Conexao.Execute "DBCC CheckIdent('Cotacao_fornecedor',Reseed,1)"
Conexao.Execute "Truncate Table Cotacao_item"
Conexao.Execute "DBCC CheckIdent('Cotacao_item',Reseed,1)"
Conexao.Execute "Truncate Table CQ_NC_FABRICA"
Conexao.Execute "DBCC CheckIdent('CQ_NC_FABRICA',Reseed,1)"
Conexao.Execute "Truncate Table CQ_NC_FABRICA_causa"
Conexao.Execute "DBCC CheckIdent('CQ_NC_FABRICA_causa',Reseed,1)"
Conexao.Execute "Truncate Table CQ_NC_FABRICA_origem"
Conexao.Execute "DBCC CheckIdent('CQ_NC_FABRICA_origem',Reseed,1)"
Conexao.Execute "Truncate Table CQ_RNC"
Conexao.Execute "DBCC CheckIdent('CQ_RNC',Reseed,1)"
Conexao.Execute "Truncate Table CQ_RNC_Classificacao"
Conexao.Execute "DBCC CheckIdent('CQ_RNC_Classificacao',Reseed,1)"
Conexao.Execute "Truncate Table CQ_RNC_documentos"
Conexao.Execute "DBCC CheckIdent('CQ_RNC_documentos',Reseed,1)"
Conexao.Execute "Truncate Table CQ_SA"
Conexao.Execute "DBCC CheckIdent('CQ_SA',Reseed,1)"
Conexao.Execute "Truncate Table CQ_SA_Equipe"
Conexao.Execute "DBCC CheckIdent('CQ_SA_Equipe',Reseed,1)"
Conexao.Execute "Truncate Table CQ_SD"
Conexao.Execute "DBCC CheckIdent('CQ_SD',Reseed,1)"
Conexao.Execute "Truncate Table CQ_Sistema"
Conexao.Execute "DBCC CheckIdent('CQ_Sistema',Reseed,1)"
Conexao.Execute "Truncate Table CQ_Sistema_Movimentacoes"
Conexao.Execute "DBCC CheckIdent('CQ_Sistema_Movimentacoes',Reseed,1)"
Conexao.Execute "Truncate Table CQ_Sistema_Tipo"
Conexao.Execute "DBCC CheckIdent('CQ_Sistema_Tipo',Reseed,1)"
Conexao.Execute "Truncate Table CST"
Conexao.Execute "DBCC CheckIdent('CST',Reseed,1)"
Conexao.Execute "Truncate Table Custos"
Conexao.Execute "DBCC CheckIdent('Custos',Reseed,1)"
Conexao.Execute "Truncate Table Custos_familias"
Conexao.Execute "DBCC CheckIdent('Custos_familias',Reseed,1)"
Conexao.Execute "Truncate Table CustosJustificativa"
Conexao.Execute "DBCC CheckIdent('CustosJustificativa',Reseed,1)"
Conexao.Execute "Truncate Table Empresa"
Conexao.Execute "DBCC CheckIdent('Empresa',Reseed,1)"
Conexao.Execute "Truncate Table Empresa_armazenamento_PDF"
Conexao.Execute "DBCC CheckIdent('Empresa_armazenamento_PDF',Reseed,1)"
Conexao.Execute "Truncate Table Empresa_CNAE_atividade"
Conexao.Execute "DBCC CheckIdent('Empresa_CNAE_atividade',Reseed,1)"
Conexao.Execute "Truncate Table Empresa_email"
Conexao.Execute "DBCC CheckIdent('Empresa_email',Reseed,1)"
Conexao.Execute "Truncate Table Empresa_Filtros"
Conexao.Execute "DBCC CheckIdent('Empresa_Filtros',Reseed,1)"
Conexao.Execute "Truncate Table Empresa_armazenamento_PDF"
Conexao.Execute "DBCC CheckIdent('Empresa_armazenamento_PDF',Reseed,1)"
Conexao.Execute "Truncate Table Estoque_Controle"
Conexao.Execute "DBCC CheckIdent('Estoque_Controle',Reseed,1)"
Conexao.Execute "Truncate Table Estoque_Controle_Empenho_Vendas"
Conexao.Execute "DBCC CheckIdent('Estoque_Controle_Empenho_Vendas',Reseed,1)"
Conexao.Execute "Truncate Table Estoque_controle_recebimento"
Conexao.Execute "DBCC CheckIdent('Estoque_controle_recebimento',Reseed,1)"
Conexao.Execute "Truncate Table Estoque_fisico"
Conexao.Execute "DBCC CheckIdent('Estoque_fisico',Reseed,1)"
Conexao.Execute "Truncate Table Estoque_Localarmazenamento"
Conexao.Execute "DBCC CheckIdent('Estoque_Localarmazenamento',Reseed,1)"

Conexao.Execute "Truncate Table Estoque_Localarmazenamento_criar"
Conexao.Execute "DBCC CheckIdent('Estoque_Localarmazenamento_criar',Reseed,1)"
Conexao.Execute "INSERT INTO Estoque_Localarmazenamento_criar (Data, Responsavel, Descricao, Status, DtValidacao, RespValidacao) VALUES ('" & Date & "', '" & PROCAM & "', 'SERVIÇOS', 'Liberado', '" & Now & "', 'PROCAM')"
Conexao.Execute "INSERT INTO Estoque_Localarmazenamento_criar (Data, Responsavel, Descricao, Status, DtValidacao, RespValidacao) VALUES ('" & Date & "', '" & PROCAM & "', 'RETORNO DE MERCADORIA', 'Liberado', '" & Now & "', 'PROCAM')"
Conexao.Execute "INSERT INTO Estoque_Localarmazenamento_criar (Data, Responsavel, Descricao, Status, DtValidacao, RespValidacao) VALUES ('" & Date & "', '" & PROCAM & "', 'INDUSTRIALIZAÇÃO', 'Liberado', '" & Now & "', 'PROCAM')"
Conexao.Execute "INSERT INTO Estoque_Localarmazenamento_criar (Data, Responsavel, Descricao, Status, DtValidacao, RespValidacao) VALUES ('" & Date & "', '" & PROCAM & "', 'ESTOQUE PADRÃO', 'Liberado', '" & Now & "', 'PROCAM')"

Conexao.Execute "Truncate Table Estoque_movimentacao"
Conexao.Execute "DBCC CheckIdent('Estoque_movimentacao',Reseed,1)"
Conexao.Execute "Truncate Table Estoque_relatorios"
Conexao.Execute "DBCC CheckIdent('Estoque_relatorios',Reseed,1)"
Conexao.Execute "Truncate Table Etiqueta"
Conexao.Execute "DBCC CheckIdent('Etiqueta',Reseed,1)"
Conexao.Execute "Truncate Table Fabricante_marca"
Conexao.Execute "DBCC CheckIdent('Fabricante_marca',Reseed,1)"
Conexao.Execute "Truncate Table Familia_financeiro"
Conexao.Execute "DBCC CheckIdent('Familia_financeiro',Reseed,1)"
Conexao.Execute "Truncate Table Fases"
Conexao.Execute "DBCC CheckIdent('Fases',Reseed,1)"
Conexao.Execute "Truncate Table Fases_revisao"
Conexao.Execute "DBCC CheckIdent('Fases_revisao',Reseed,1)"
Conexao.Execute "Truncate Table Faturamento_Relacionamento"
Conexao.Execute "DBCC CheckIdent('Faturamento_Relacionamento',Reseed,1)"
Conexao.Execute "Truncate Table Feriados"
Conexao.Execute "DBCC CheckIdent('Feriados',Reseed,1)"
Conexao.Execute "Truncate Table Ferramentas"
Conexao.Execute "DBCC CheckIdent('Ferramentas',Reseed,1)"
Conexao.Execute "Truncate Table Funcionarios"
Conexao.Execute "DBCC CheckIdent('Funcionarios',Reseed,1)"
Conexao.Execute "Truncate Table Funcionarios_atestados"
Conexao.Execute "DBCC CheckIdent('Funcionarios_atestados',Reseed,1)"
Conexao.Execute "Truncate Table Funcionarios_aumentos"
Conexao.Execute "DBCC CheckIdent('Funcionarios_aumentos',Reseed,1)"
Conexao.Execute "Truncate Table Funcionarios_cursos"
Conexao.Execute "DBCC CheckIdent('Funcionarios_cursos',Reseed,1)"
Conexao.Execute "Truncate Table Funcionarios_dependentes"
Conexao.Execute "DBCC CheckIdent('Funcionarios_dependentes',Reseed,1)"
Conexao.Execute "Truncate Table Funcionarios_descricao"
Conexao.Execute "DBCC CheckIdent('Funcionarios_descricao',Reseed,1)"
Conexao.Execute "Truncate Table Funcionarios_divisao"
Conexao.Execute "DBCC CheckIdent('Funcionarios_divisao',Reseed,1)"
Conexao.Execute "Truncate Table Funcionarios_documentos"
Conexao.Execute "DBCC CheckIdent('Funcionarios_documentos',Reseed,1)"
Conexao.Execute "Truncate Table Funcionarios_ferias"
Conexao.Execute "DBCC CheckIdent('Funcionarios_ferias',Reseed,1)"
Conexao.Execute "Truncate Table Funcionarios_funcao"
Conexao.Execute "DBCC CheckIdent('Funcionarios_funcao',Reseed,1)"
Conexao.Execute "Truncate Table Funcionarios_obs"
Conexao.Execute "DBCC CheckIdent('Funcionarios_obs',Reseed,1)"
Conexao.Execute "Truncate Table Funcionarios_sindicato"
Conexao.Execute "DBCC CheckIdent('Funcionarios_sindicato',Reseed,1)"
Conexao.Execute "Truncate Table Funcionarios_tipo"
Conexao.Execute "DBCC CheckIdent('Funcionarios_tipo',Reseed,1)"
Conexao.Execute "Truncate Table Funcionarios_turno"
Conexao.Execute "DBCC CheckIdent('Funcionarios_turno',Reseed,1)"
Conexao.Execute "Truncate Table Grupo_op"
Conexao.Execute "DBCC CheckIdent('Grupo_op',Reseed,1)"
Conexao.Execute "Truncate Table hist_proc"
Conexao.Execute "DBCC CheckIdent('hist_proc',Reseed,1)"
Conexao.Execute "Truncate Table HistProc"
Conexao.Execute "DBCC CheckIdent('HistProc',Reseed,1)"
Conexao.Execute "Truncate Table Impostos"
Conexao.Execute "DBCC CheckIdent('Impostos',Reseed,1)"
Conexao.Execute "Truncate Table Impostos_FaturamentoMes"
Conexao.Execute "DBCC CheckIdent('Impostos_FaturamentoMes',Reseed,1)"
'Conexao.Execute "Truncate Table Impostos_TabelaDAS"
'Conexao.Execute "DBCC CheckIdent('Impostos_TabelaDAS',Reseed,1)"
Conexao.Execute "Truncate Table Instrumentos"
Conexao.Execute "DBCC CheckIdent('Instrumentos',Reseed,1)"
Conexao.Execute "Truncate Table Intervalos"
Conexao.Execute "DBCC CheckIdent('Intervalos',Reseed,1)"
Conexao.Execute "Truncate Table item_aplicacoes"
Conexao.Execute "DBCC CheckIdent('item_aplicacoes',Reseed,1)"
Conexao.Execute "Truncate Table KeyLock"
Conexao.Execute "DBCC CheckIdent('KeyLock',Reseed,1)"
Conexao.Execute "Truncate Table Liquido_penetrante"
Conexao.Execute "DBCC CheckIdent('Liquido_penetrante',Reseed,1)"
Conexao.Execute "Truncate Table Logon"
Conexao.Execute "DBCC CheckIdent('Logon',Reseed,1)"
Conexao.Execute "Truncate Table Manutencao"
Conexao.Execute "DBCC CheckIdent('Manutencao',Reseed,1)"
Conexao.Execute "Truncate Table Manutencao_Checklist"
Conexao.Execute "DBCC CheckIdent('Manutencao_Checklist',Reseed,1)"
Conexao.Execute "Truncate Table Manutencao_data"
Conexao.Execute "DBCC CheckIdent('Manutencao_data',Reseed,1)"
Conexao.Execute "Truncate Table Manutencao_defeito"
Conexao.Execute "DBCC CheckIdent('Manutencao_defeito',Reseed,1)"
Conexao.Execute "Truncate Table Manutencao_descricao"
Conexao.Execute "DBCC CheckIdent('Manutencao_descricao',Reseed,1)"
Conexao.Execute "Truncate Table Mascara"
Conexao.Execute "DBCC CheckIdent('Mascara',Reseed,1)"
Conexao.Execute "Truncate Table Medicao"
Conexao.Execute "DBCC CheckIdent('Medicao',Reseed,1)"
Conexao.Execute "Truncate Table Medicaodimensao"
Conexao.Execute "DBCC CheckIdent('Medicaodimensao',Reseed,1)"
Conexao.Execute "Truncate Table Medicaodimensao_Familia"
Conexao.Execute "DBCC CheckIdent('Medicaodimensao_Familia',Reseed,1)"
Conexao.Execute "Truncate Table Medicaodimensao_instrumentos"
Conexao.Execute "DBCC CheckIdent('Medicaodimensao_instrumentos',Reseed,1)"
Conexao.Execute "Truncate Table Medicaodimensao_peca"
Conexao.Execute "DBCC CheckIdent('Medicaodimensao_peca',Reseed,1)"
Conexao.Execute "Truncate Table Medicaodimensao_peca_relatorios"
Conexao.Execute "DBCC CheckIdent('Medicaodimensao_peca_relatorios',Reseed,1)"
Conexao.Execute "Truncate Table Minuta"
Conexao.Execute "DBCC CheckIdent('Minuta',Reseed,1)"
Conexao.Execute "Truncate Table Minuta_notas"
Conexao.Execute "DBCC CheckIdent('Minuta_notas',Reseed,1)"
Conexao.Execute "Truncate Table NF_Carta_Correcao"
Conexao.Execute "DBCC CheckIdent('NF_Carta_Correcao',Reseed,1)"
Conexao.Execute "Truncate Table Norma"
Conexao.Execute "DBCC CheckIdent('Norma',Reseed,1)"
Conexao.Execute "Truncate Table Ordemservico"
Conexao.Execute "DBCC CheckIdent('Ordemservico',Reseed,1)"
Conexao.Execute "Truncate Table Ordemservico_HoraUtilizadaporDia"
Conexao.Execute "DBCC CheckIdent('Ordemservico_HoraUtilizadaporDia',Reseed,1)"
Conexao.Execute "Truncate Table Ordemservico_maq_utilizadas"
Conexao.Execute "DBCC CheckIdent('Ordemservico_maq_utilizadas',Reseed,1)"
Conexao.Execute "Truncate Table Outros_SolicitacaoPCP"
Conexao.Execute "DBCC CheckIdent('Outros_SolicitacaoPCP',Reseed,1)"
Conexao.Execute "Truncate Table PCP_programacao"
Conexao.Execute "DBCC CheckIdent('PCP_programacao',Reseed,1)"
Conexao.Execute "Truncate Table PCP_programacao_ordem"
Conexao.Execute "DBCC CheckIdent('PCP_programacao_ordem',Reseed,1)"
Conexao.Execute "Truncate Table Plano"
Conexao.Execute "DBCC CheckIdent('Plano',Reseed,1)"
Conexao.Execute "Truncate Table Plano_de_contas_totalizacao"
Conexao.Execute "DBCC CheckIdent('Plano_de_contas_totalizacao',Reseed,1)"
Conexao.Execute "Truncate Table Plano_revisao"
Conexao.Execute "DBCC CheckIdent('Plano_revisao',Reseed,1)"
Conexao.Execute "Truncate Table Planodimensao"
Conexao.Execute "DBCC CheckIdent('Planodimensao',Reseed,1)"
Conexao.Execute "Truncate Table Planodimensao_instrumentos"
Conexao.Execute "DBCC CheckIdent('Planodimensao_instrumentos',Reseed,1)"
Conexao.Execute "Truncate Table Processos"
Conexao.Execute "DBCC CheckIdent('Processos',Reseed,1)"
Conexao.Execute "Truncate Table Producao"
Conexao.Execute "DBCC CheckIdent('Producao',Reseed,1)"
Conexao.Execute "Truncate Table Producao_etiquetas"
Conexao.Execute "DBCC CheckIdent('Producao_etiquetas',Reseed,1)"
Conexao.Execute "Truncate Table Producao_NF_Consignada"
Conexao.Execute "DBCC CheckIdent('Producao_outras_despesas',Reseed,1)"
Conexao.Execute "Truncate Table Producao_outras_despesas"
Conexao.Execute "DBCC CheckIdent('Producao_NF_Consignada',Reseed,1)"
Conexao.Execute "Truncate Table Producao_pedidos"
Conexao.Execute "DBCC CheckIdent('Producao_pedidos',Reseed,1)"
Conexao.Execute "Truncate Table Producao_Relatorios"
Conexao.Execute "DBCC CheckIdent('Producao_Relatorios',Reseed,1)"
Conexao.Execute "Truncate Table Producao_Relatorios_Detalhes"
Conexao.Execute "DBCC CheckIdent('Producao_Relatorios_Detalhes',Reseed,1)"
Conexao.Execute "Truncate Table Producao_Relatorios_Total"
Conexao.Execute "DBCC CheckIdent('Producao_Relatorios_Total',Reseed,1)"
Conexao.Execute "Truncate Table ProducaoFases"
Conexao.Execute "DBCC CheckIdent('ProducaoFases',Reseed,1)"
Conexao.Execute "Truncate Table ProducaoFases_OS"
Conexao.Execute "DBCC CheckIdent('ProducaoFases_OS',Reseed,1)"
Conexao.Execute "Truncate Table ProducaoFases_Backup"
Conexao.Execute "DBCC CheckIdent('ProducaoFases_Backup',Reseed,1)"
Conexao.Execute "Truncate Table ProducaoFases_Totalizacao"
Conexao.Execute "DBCC CheckIdent('ProducaoFases_Totalizacao',Reseed,1)"
Conexao.Execute "Truncate Table ProducaoFases_Totalizacao_Backup"
Conexao.Execute "DBCC CheckIdent('ProducaoFases_Totalizacao_Backup',Reseed,1)"
Conexao.Execute "Truncate Table Producaomaterial"
Conexao.Execute "DBCC CheckIdent('Producaomaterial',Reseed,1)"
Conexao.Execute "Truncate Table Programas"
Conexao.Execute "DBCC CheckIdent('Programas',Reseed,1)"
Conexao.Execute "Truncate Table Projconjunto"
Conexao.Execute "DBCC CheckIdent('Projconjunto',Reseed,1)"
Conexao.Execute "Truncate Table Projconjunto_desc_versao"
Conexao.Execute "DBCC CheckIdent('Projconjunto_desc_versao',Reseed,1)"
Conexao.Execute "Truncate Table Projfamilia"
Conexao.Execute "DBCC CheckIdent('Projfamilia',Reseed,1)"
Conexao.Execute "Truncate Table projfamilia_grupo"
Conexao.Execute "DBCC CheckIdent('projfamilia_grupo',Reseed,1)"
Conexao.Execute "Truncate Table projproduto"
Conexao.Execute "DBCC CheckIdent('projproduto',Reseed,1)"
Conexao.Execute "Truncate Table Projproduto_clientes"
Conexao.Execute "DBCC CheckIdent('Projproduto_clientes',Reseed,1)"
Conexao.Execute "Truncate Table Projproduto_descricao"
Conexao.Execute "DBCC CheckIdent('Projproduto_descricao',Reseed,1)"
Conexao.Execute "Truncate Table projproduto_documentos"
Conexao.Execute "DBCC CheckIdent('projproduto_documentos',Reseed,1)"
Conexao.Execute "Truncate Table Projproduto_fabricante"
Conexao.Execute "DBCC CheckIdent('Projproduto_fabricante',Reseed,1)"
Conexao.Execute "Truncate Table Projproduto_fornecedor"
Conexao.Execute "DBCC CheckIdent('Projproduto_fornecedor',Reseed,1)"
Conexao.Execute "Truncate Table Projproduto_revisao"
Conexao.Execute "DBCC CheckIdent('Projproduto_revisao',Reseed,1)"
Conexao.Execute "Truncate Table Projproduto_similar"
Conexao.Execute "DBCC CheckIdent('Projproduto_similar',Reseed,1)"
Conexao.Execute "Truncate Table Qualidade_revisao_relatorios"
Conexao.Execute "DBCC CheckIdent('Qualidade_revisao_relatorios',Reseed,1)"
Conexao.Execute "Truncate Table Qualidade_revisao_relatorios_subreports"
Conexao.Execute "DBCC CheckIdent('Qualidade_revisao_relatorios_subreports',Reseed,1)"
Conexao.Execute "Truncate Table QualidadePPAP"
Conexao.Execute "DBCC CheckIdent('QualidadePPAP',Reseed,1)"
Conexao.Execute "Truncate Table QualidadePPAP_FMEA"
Conexao.Execute "DBCC CheckIdent('QualidadePPAP_FMEA',Reseed,1)"
Conexao.Execute "Truncate Table QualidadePPAP_FMEA_EfeitoFalha"
Conexao.Execute "DBCC CheckIdent('QualidadePPAP_FMEA_EfeitoFalha',Reseed,1)"
Conexao.Execute "Truncate Table QualidadePPAP_FMEA_Fases"
Conexao.Execute "DBCC CheckIdent('QualidadePPAP_FMEA_Fases',Reseed,1)"
Conexao.Execute "Truncate Table QualidadePPAP_FMEA_ModoFalha"
Conexao.Execute "DBCC CheckIdent('QualidadePPAP_FMEA_ModoFalha',Reseed,1)"
Conexao.Execute "Truncate Table QualidadePPAP_Padrao"
Conexao.Execute "DBCC CheckIdent('QualidadePPAP_Padrao',Reseed,1)"
Conexao.Execute "Truncate Table QualidadePPAP_PlanoControle"
Conexao.Execute "DBCC CheckIdent('QualidadePPAP_PlanoControle',Reseed,1)"
Conexao.Execute "Truncate Table QualidadePPAP_PlanoControle_Dimensoes"
Conexao.Execute "DBCC CheckIdent('QualidadePPAP_PlanoControle_Dimensoes',Reseed,1)"
Conexao.Execute "Truncate Table Requisicao_materiais"
Conexao.Execute "DBCC CheckIdent('Requisicao_materiais',Reseed,1)"
Conexao.Execute "Truncate Table Requisicao_materiais_lista"
Conexao.Execute "DBCC CheckIdent('Requisicao_materiais_lista',Reseed,1)"
Conexao.Execute "Truncate Table Suporte"
Conexao.Execute "DBCC CheckIdent('Suporte',Reseed,1)"
Conexao.Execute "Truncate Table Segmentos"
Conexao.Execute "DBCC CheckIdent('Segmentos',Reseed,1)"
Conexao.Execute "Truncate Table Tabela_conversao_unidade"
Conexao.Execute "DBCC CheckIdent('Tabela_conversao_unidade',Reseed,1)"
Conexao.Execute "Truncate Table tbl_contas_antecipacao"
Conexao.Execute "DBCC CheckIdent('tbl_contas_antecipacao',Reseed,1)"
Conexao.Execute "Truncate Table tbl_contas_devolucao"
Conexao.Execute "DBCC CheckIdent('tbl_contas_devolucao',Reseed,1)"
Conexao.Execute "Truncate Table tbl_contas_receber"
Conexao.Execute "DBCC CheckIdent('tbl_contas_receber',Reseed,1)"
Conexao.Execute "Truncate Table tbl_contas_receber_duplicatas"
Conexao.Execute "DBCC CheckIdent('tbl_contas_receber_duplicatas',Reseed,1)"
Conexao.Execute "Truncate Table tbl_Contas_Varias"
Conexao.Execute "DBCC CheckIdent('tbl_Contas_Varias',Reseed,1)"
Conexao.Execute "Truncate Table tbl_ContasPagar"
Conexao.Execute "DBCC CheckIdent('tbl_ContasPagar',Reseed,1)"
Conexao.Execute "Truncate Table tbl_ContasPagar_Saque"
Conexao.Execute "DBCC CheckIdent('tbl_ContasPagar_Saque',Reseed,1)"
Conexao.Execute "Truncate Table tbl_ContasPagar_Tipo_Docto"
Conexao.Execute "DBCC CheckIdent('tbl_ContasPagar_Tipo_Docto',Reseed,1)"
Conexao.Execute "Truncate Table tbl_Dados_Nota_Fiscal"
Conexao.Execute "DBCC CheckIdent('tbl_Dados_Nota_Fiscal',Reseed,1)"
Conexao.Execute "Truncate Table tbl_Dados_Nota_Fiscal_NFe"
Conexao.Execute "DBCC CheckIdent('tbl_Dados_Nota_Fiscal_NFe',Reseed,1)"
Conexao.Execute "Truncate Table tbl_Dados_Transp"
Conexao.Execute "DBCC CheckIdent('tbl_Dados_Transp',Reseed,1)"
Conexao.Execute "Truncate Table tbl_DadosAdicionais"
Conexao.Execute "DBCC CheckIdent('tbl_DadosAdicionais',Reseed,1)"
Conexao.Execute "Truncate Table tbl_DadosAdicionais_padrao"
Conexao.Execute "DBCC CheckIdent('tbl_DadosAdicionais_padrao',Reseed,1)"
Conexao.Execute "Truncate Table tbl_Detalhes_Nota"
Conexao.Execute "DBCC CheckIdent('tbl_Detalhes_Nota',Reseed,1)"
Conexao.Execute "Truncate Table tbl_Detalhes_Nota_CST_Cofins"
Conexao.Execute "DBCC CheckIdent('tbl_Detalhes_Nota_CST_Cofins',Reseed,1)"
Conexao.Execute "Truncate Table tbl_Detalhes_Nota_CST_ICMS"
Conexao.Execute "DBCC CheckIdent('tbl_Detalhes_Nota_CST_ICMS',Reseed,1)"
Conexao.Execute "Truncate Table tbl_Detalhes_Nota_CST_IPI"
Conexao.Execute "DBCC CheckIdent('tbl_Detalhes_Nota_CST_IPI',Reseed,1)"
Conexao.Execute "Truncate Table tbl_Detalhes_Nota_CST_PIS"
Conexao.Execute "DBCC CheckIdent('tbl_Detalhes_Nota_CST_PIS',Reseed,1)"
Conexao.Execute "Truncate Table tbl_Detalhes_Nota_NFe"
Conexao.Execute "DBCC CheckIdent('tbl_Detalhes_Nota_NFe',Reseed,1)"
Conexao.Execute "Truncate Table tbl_Detalhes_Nota_pedidos"
Conexao.Execute "DBCC CheckIdent('tbl_Detalhes_Nota_pedidos',Reseed,1)"
Conexao.Execute "Truncate Table tbl_Detalhes_Recebimento"
Conexao.Execute "DBCC CheckIdent('tbl_Detalhes_Recebimento',Reseed,1)"
Conexao.Execute "Truncate Table tbl_Detalhes_Recebimento_Nboletos"
Conexao.Execute "DBCC CheckIdent('tbl_Detalhes_Recebimento_Nboletos',Reseed,1)"
Conexao.Execute "Truncate Table tbl_Fluxo_de_caixa"
Conexao.Execute "DBCC CheckIdent('tbl_Fluxo_de_caixa',Reseed,1)"
Conexao.Execute "Truncate Table tbl_Fluxo_de_caixa_saldos"
Conexao.Execute "DBCC CheckIdent('tbl_Fluxo_de_caixa_saldos',Reseed,1)"
Conexao.Execute "Truncate Table Tbl_Fluxo_de_Caixa2"
Conexao.Execute "DBCC CheckIdent('Tbl_Fluxo_de_Caixa2',Reseed,1)"
Conexao.Execute "Truncate Table tbl_Instituicoes"
Conexao.Execute "DBCC CheckIdent('tbl_Instituicoes',Reseed,1)"
Conexao.Execute "Truncate Table tbl_instituicoes_transf"
Conexao.Execute "DBCC CheckIdent('tbl_instituicoes_transf',Reseed,1)"
Conexao.Execute "Truncate Table tbl_NaturezaOperacao_CST"
Conexao.Execute "DBCC CheckIdent('tbl_NaturezaOperacao_CST',Reseed,1)"
Conexao.Execute "Truncate Table tbl_proposta_nota"
Conexao.Execute "DBCC CheckIdent('tbl_proposta_nota',Reseed,1)"
Conexao.Execute "Truncate Table tbl_Totais_Nota"
Conexao.Execute "DBCC CheckIdent('tbl_Totais_Nota',Reseed,1)"
Conexao.Execute "Truncate Table Tipodimensao"
Conexao.Execute "DBCC CheckIdent('Tipodimensao',Reseed,1)"
Conexao.Execute "Truncate Table troca_titulo"
Conexao.Execute "DBCC CheckIdent('troca_titulo',Reseed,1)"
Conexao.Execute "Truncate Table Troca_titulo_relatorio"
Conexao.Execute "DBCC CheckIdent('Troca_titulo_relatorio',Reseed,1)"
Conexao.Execute "Truncate Table troca_titulo_valores"
Conexao.Execute "DBCC CheckIdent('troca_titulo_valores',Reseed,1)"
Conexao.Execute "Truncate Table troca_titulo_ValoresImpostos"
Conexao.Execute "DBCC CheckIdent('troca_titulo_ValoresImpostos',Reseed,1)"
Conexao.Execute "Truncate Table UltraSom"
Conexao.Execute "DBCC CheckIdent('UltraSom',Reseed,1)"
Conexao.Execute "Truncate Table UltraSom_inspetores"
Conexao.Execute "DBCC CheckIdent('UltraSom_inspetores',Reseed,1)"
Conexao.Execute "Truncate Table Usuarios_monitor_trabalho"
Conexao.Execute "DBCC CheckIdent('Usuarios_monitor_trabalho',Reseed,1)"
Conexao.Execute "Truncate Table Usuarios_Setor"
Conexao.Execute "DBCC CheckIdent('Usuarios_Setor',Reseed,1)"
Conexao.Execute "Truncate Table Usuarios_Setor_Consolidacao"
Conexao.Execute "DBCC CheckIdent('Usuarios_Setor_Consolidacao',Reseed,1)"
Conexao.Execute "Truncate Table Usuarios_Setor_Previsao"
Conexao.Execute "DBCC CheckIdent('Usuarios_Setor_Previsao',Reseed,1)"
Conexao.Execute "Truncate Table Usuarios_Setor_Responsavel"
Conexao.Execute "DBCC CheckIdent('Usuarios_Setor_Responsavel',Reseed,1)"
Conexao.Execute "Truncate Table Vendas_analise"
Conexao.Execute "DBCC CheckIdent('Vendas_analise',Reseed,1)"
Conexao.Execute "Truncate Table Vendas_analise_descricao_checklist"
Conexao.Execute "DBCC CheckIdent('Vendas_analise_descricao_checklist',Reseed,1)"
Conexao.Execute "Truncate Table Vendas_analise_ProdutosProcessos"
Conexao.Execute "DBCC CheckIdent('Vendas_analise_ProdutosProcessos',Reseed,1)"
Conexao.Execute "Truncate Table Vendas_analise_setores"
Conexao.Execute "DBCC CheckIdent('Vendas_analise_setores',Reseed,1)"
Conexao.Execute "Truncate Table Vendas_analise_setores_checklist"
Conexao.Execute "DBCC CheckIdent('Vendas_analise_setores_checklist',Reseed,1)"
Conexao.Execute "Truncate Table Vendas_analise_tipo"
Conexao.Execute "DBCC CheckIdent('Vendas_analise_tipo',Reseed,1)"
Conexao.Execute "Truncate Table vendas_carteira"
Conexao.Execute "DBCC CheckIdent('vendas_carteira',Reseed,1)"
Conexao.Execute "Truncate Table vendas_carteira_alteracoes"
Conexao.Execute "DBCC CheckIdent('vendas_carteira_alteracoes',Reseed,1)"
Conexao.Execute "Truncate Table vendas_carteira_composicao"
Conexao.Execute "DBCC CheckIdent('vendas_carteira_composicao',Reseed,1)"
Conexao.Execute "Truncate Table vendas_comercial"
Conexao.Execute "DBCC CheckIdent('vendas_comercial',Reseed,1)"
Conexao.Execute "Truncate Table Vendas_programa"
Conexao.Execute "DBCC CheckIdent('Vendas_programa',Reseed,1)"
Conexao.Execute "Truncate Table Vendas_programa_item"
Conexao.Execute "DBCC CheckIdent('Vendas_programa_item',Reseed,1)"
Conexao.Execute "Truncate Table Vendas_programacao"
Conexao.Execute "DBCC CheckIdent('Vendas_programacao',Reseed,1)"
Conexao.Execute "Truncate Table vendas_proposta"
Conexao.Execute "DBCC CheckIdent('vendas_proposta',Reseed,1)"
Conexao.Execute "Truncate Table vendas_proposta_dadoscomerciais_padrao"
Conexao.Execute "DBCC CheckIdent('vendas_proposta_dadoscomerciais_padrao',Reseed,1)"
Conexao.Execute "Truncate Table vendas_proposta_previsaopgto"
Conexao.Execute "DBCC CheckIdent('vendas_proposta_previsaopgto',Reseed,1)"
Conexao.Execute "Truncate Table Vendas_Tele"
Conexao.Execute "DBCC CheckIdent('Vendas_Tele',Reseed,1)"
Conexao.Execute "Truncate Table Vendas_Vendedores"
Conexao.Execute "DBCC CheckIdent('Vendas_Vendedores',Reseed,1)"
Conexao.Execute "Truncate Table Vendas_Vendedores_Clientes"
Conexao.Execute "DBCC CheckIdent('Vendas_Vendedores_Clientes',Reseed,1)"
Conexao.Execute "Truncate Table Vendas_Vendedores_Produto"
Conexao.Execute "DBCC CheckIdent('Vendas_Vendedores_Produto',Reseed,1)"
USMsgBox ("Banco de dados reiniciado com sucesso."), vbInformation, "CAPRIND v5.0"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcVerifAtualizacao()
On Error GoTo tratar_erro

LocalAntigoCaprind = GetSetting("Procam", "CaprindSQL", "LocalAntigoCaprind")
LocalNovoCaprind = GetSetting("Procam", "CaprindSQL", "LocalNovoCaprind")
If LocalAntigoCaprind <> "" And LocalNovoCaprind <> "" Then
    CaminhoAnt = Left((LocalAntigoCaprind), Len(LocalAntigoCaprind) - 12)
    CaminhoNovo = Left((LocalNovoCaprind), Len(LocalNovoCaprind) - 12)
    Set Fso = CreateObject("Scripting.FileSystemObject")
    Set Fsu = CreateObject("Scripting.FileSystemObject")
    NomeCampo = "Caprind.exe na pasta " & CaminhoAnt
    Set f = Fso.GetFile(LocalAntigoCaprind)
    NomeCampo = "Caprind.exe na pasta " & CaminhoNovo
    Set FU = Fsu.GetFile(LocalNovoCaprind)
    If f.DateLastModified < FU.DateLastModified Then
        USMsgBox ("O sistema Caprind está desatualizado e será atualizado automaticamente."), vbInformation, "CAPRIND v5.0"
        NomeCampo = "SincCaprind.exe na pasta " & CaminhoNovo
        Shell CaminhoNovo & "\SincCaprind.exe", vbNormalFocus
        End
    End If
End If

Exit Sub
tratar_erro:
    If Err.Number = "53" Then
        USMsgBox ("Não será possível atualizar o sistema, pois não foi encontrado o arquivo " & NomeCampo & "."), vbExclamation, "CAPRIND v5.0"
    Else
        USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    End If
End Sub

Private Sub ProcVerifAlteracaoData()
On Error GoTo tratar_erro

If ActiveLock1.LastRunDate > Now Then
    USMsgBox ("Ocorreu uma alteração na data do sistema operacional, o sistema Caprind será encerrado."), vbCritical, "CAPRIND v5.0"
    End
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro
  
CompactarRepararBanco = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub


Private Sub txtSenha_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus txtSenha
MudaCor Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub txtSenha_LostFocus()
On Error GoTo tratar_erro

MudaCor Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub txtUsuario_Change()
On Error GoTo tratar_erro

Call FunTextoMaiusculoDigitar(txtUsuario)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub txtUsuario_GotFocus()
On Error GoTo tratar_erro

MudaCor Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub txtUsuario_LostFocus()
On Error GoTo tratar_erro
  
If FunAbreBD() = False Then
    USMsgBox ("Não foi possível conectar com o banco de dados do " & NomeCampo & ", favor verificar as configurações do sistema."), vbExclamation, "CAPRIND v5.0"
    frmOpcoesGeral2.Show 1
    Exit Sub
Else
  
If txtUsuario <> "" Then

txtUsuario.Text = UCase(txtUsuario.Text)
MudaCor Me

Set TBUsuarios = CreateObject("adodb.recordset")
TBUsuarios.Open "Select Senha from usuarios where Usuario = '" & txtUsuario.Text & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBUsuarios.EOF = False Then
    If TBUsuarios!Senha = "" Or IsNull(TBUsuarios!Senha) = True Then
        frmUsuarios_Senha.Show
    End If
End If
TBUsuarios.Close
Else
'USMsgBox "Digite seu nome de usuario.", vbInformation, "CAPRIND V5.0"
'txtUsuario.SetFocus
End If


End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

