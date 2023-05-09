VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{84147065-0227-424E-827F-9E79B1DA5D8B}#21.0#0"; "kftp.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2014.ocx"
Begin VB.Form frmLogon 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00F9F9F9&
   BorderStyle     =   0  'None
   ClientHeight    =   11370
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   19200
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   7.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogon.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   11370
   ScaleWidth      =   19200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cmbBanco 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmLogon.frx":0CCA
      Left            =   8850
      List            =   "frmLogon.frx":0CCC
      Style           =   2  'Dropdown List
      TabIndex        =   2
      ToolTipText     =   "Banco de dados."
      Top             =   8775
      Width           =   2535
   End
   Begin VB.TextBox txtUsuario 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
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
      Height          =   360
      Left            =   8850
      TabIndex        =   0
      ToolTipText     =   "Nome do usuário."
      Top             =   7695
      Width           =   2565
   End
   Begin VB.TextBox txtSenha 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
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
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   8850
      PasswordChar    =   "*"
      TabIndex        =   1
      ToolTipText     =   "Senha do usuário."
      Top             =   8205
      Width           =   2565
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   1440
      Top             =   855
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   450
      Top             =   1170
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   99
      ScreenHeight    =   768
      ScreenWidth     =   1366
      ScreenHeightDT  =   1080
      ScreenWidthDT   =   1920
      AutoResizeOnLoad=   0   'False
      ApplicationName =   "Active Resize Control Professional"
      FormHeightDT    =   11370
      FormWidthDT     =   19200
      FormScaleHeightDT=   11370
      FormScaleWidthDT=   19200
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin KFTPActiveX.kftp kftp 
      Height          =   600
      Left            =   1080
      TabIndex        =   8
      Top             =   -675
      Visible         =   0   'False
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   1058
   End
   Begin DrawSuite2014.USButton Cmd_novo_local_bd 
      Height          =   315
      Left            =   11430
      TabIndex        =   12
      ToolTipText     =   "Configurações do banco de dados (F8)"
      Top             =   8775
      Width           =   345
      _ExtentX        =   609
      _ExtentY        =   556
      BorderColor     =   8421504
      BorderColorDisabled=   0
      BorderColorDown =   15048022
      BorderColorOver =   15381630
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
      GradientColor2  =   14737632
      GradientColor3  =   12632256
      GradientColor4  =   12632256
      PicNormal       =   "frmLogon.frx":0CCE
      Theme           =   1
   End
   Begin DrawSuite2014.USButton Cmd_ajuda 
      Height          =   315
      Left            =   11820
      TabIndex        =   13
      ToolTipText     =   "Ajuda (F1)"
      Top             =   8775
      Width           =   345
      _ExtentX        =   609
      _ExtentY        =   556
      BorderColor     =   8421504
      BorderColorDisabled=   0
      BorderColorDown =   15048022
      BorderColorOver =   15381630
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
      GradientColor2  =   14737632
      GradientColor3  =   12632256
      GradientColor4  =   12632256
      PicNormal       =   "frmLogon.frx":1268
      Theme           =   1
   End
   Begin DrawSuite2014.USButton cmdsair 
      Height          =   495
      Left            =   10020
      TabIndex        =   4
      Top             =   9360
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   873
      BorderColor     =   5263559
      BorderColorDisabled=   13160660
      BorderColorDown =   4013465
      BorderColorOver =   4408288
      Caption         =   "Sair"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
      Theme           =   4
   End
   Begin DrawSuite2014.USButton Cmd_chat 
      Height          =   345
      Left            =   8400
      TabIndex        =   5
      Top             =   9945
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   609
      BorderColor     =   0
      BorderColorDisabled=   13160660
      BorderColorDown =   4210752
      BorderColorOver =   8421504
      Caption         =   "Precisa de ajuda?"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Rounded MT Bold"
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
      GradientColor1  =   0
      GradientColor2  =   0
      GradientColor3  =   0
      GradientColor4  =   0
      GradientColorDisabled1=   13160660
      GradientColorDisabled2=   13160660
      GradientColorDisabled3=   13160660
      GradientColorDisabled4=   13160660
      GradientColorDown1=   4210752
      GradientColorDown2=   4210752
      GradientColorDown3=   4210752
      GradientColorDown4=   4210752
      GradientColorOver1=   8421504
      GradientColorOver2=   8421504
      GradientColorOver3=   8421504
      GradientColorOver4=   8421504
      Theme           =   6
   End
   Begin DrawSuite2014.USButton cmdAcessar 
      Height          =   495
      Left            =   8400
      TabIndex        =   3
      Top             =   9360
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   873
      BorderColor     =   4960354
      BorderColorDisabled=   13160660
      BorderColorDown =   4210752
      BorderColorOver =   49152
      Caption         =   "Entrar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
      Theme           =   3
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Senha:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   8190
      TabIndex        =   17
      Top             =   8235
      Width           =   510
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Recuperar senha"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   195
      Left            =   11505
      TabIndex        =   16
      Top             =   8280
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Usuário:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   8100
      TabIndex        =   15
      Top             =   7740
      Width           =   600
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Banco de dados:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   7515
      TabIndex        =   14
      Top             =   8820
      Width           =   1200
   End
   Begin DrawSuite2014.USAlphaImage USAlphaImage1 
      Height          =   2685
      Left            =   3500
      Top             =   3915
      Width           =   10560
      _ExtentX        =   18627
      _ExtentY        =   4736
      Image           =   "frmLogon.frx":1802
   End
   Begin VB.Label lbldireitos 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   $"frmLogon.frx":15979
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Left            =   4725
      TabIndex        =   11
      Top             =   11070
      Width           =   10290
   End
   Begin VB.Label lblVersaoatual 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2.5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   17640
      TabIndex        =   10
      Top             =   90
      Width           =   1215
   End
   Begin VB.Label lblano 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright 1999 - 2018 Caprind Sistemas ®. Todos os direitos reservados."
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Left            =   7785
      TabIndex        =   9
      Top             =   10800
      Width           =   4665
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Informe os seus dados abaixo para efetuar o logon no sistema."
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   7440
      TabIndex        =   7
      Top             =   6795
      Width           =   5475
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "www.caprind.com.br"
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
      Height          =   195
      Left            =   12465
      TabIndex        =   6
      Top             =   5445
      Width           =   1500
   End
   Begin VB.Image Image1 
      Height          =   11520
      Left            =   -135
      Picture         =   "frmLogon.frx":15A0A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   19470
   End
End
Attribute VB_Name = "frmLogon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Cmd_chat_Click()
On Error GoTo tratar_erro

If Atualizando = True Then Exit Sub
If MsgBox("Deseja realmente entrar no chat online?", vbQuestion + vbYesNo) = vbYes Then
    If cmbBanco = "" Then
        MsgBox ("Informe o banco de dados antes de entrar no chat online."), vbExclamation
        cmbBanco.SetFocus
        Exit Sub
    End If
    ProcVerificaInternet False, True
    If TemInternet = True Then
        If FunVerifHorarioChat = False Then Exit Sub
        FunAbreBD
        If FunVerificaManutencaoAtiva = False Then Exit Sub
'        Chat = True
'        Video_ajuda = False
'        With Frm_web
'            .Web.Navigate "http://www.caprind.com.br/Suporte/chat.php"
'            .Show 1
'        End With
        Set ie = New InternetExplorer
        ie.Navigate "http://www.caprind.com.br/Suporte/chat.php"
        ie.Visible = True
    Else
        MsgBox ("Não é permitido entrar no chat online, pois não foi encontrado conexão com a internet."), vbExclamation
    End If
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Cmd_novo_local_bd_Click()
On Error GoTo tratar_erro

If Atualizando = True Then Exit Sub
frmOpcoesGeral2.Show 1

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub cmdAcessar_Click()
On Error GoTo tratar_erro

If Atualizando = True Then Exit Sub
If cmbBanco.Text <> "" Then
    If FunAbreBD() = False Then
        MsgBox ("Não foi possível conectar com o banco de dados do " & NomeCampo & ", favor verificar as configurações do sistema."), vbExclamation
        frmOpcoesGeral2.Show 1
        Exit Sub
    Else
        ProcEntrar
    End If
Else
    MsgBox ("É necessário configurar os dados para conexão com o banco de dados."), vbExclamation
    frmOpcoesGeral2.Show 1
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub cmdSair_Click()
On Error GoTo tratar_erro

If Atualizando = True Then Exit Sub
ProcSair

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Sub ProcEntrar()
On Error GoTo tratar_erro

If cmbBanco.Text = "" Then
    MsgBox ("Informe o nome do banco de dados antes de efetuar o logon."), vbExclamation
    cmbBanco.SetFocus
    Exit Sub
End If
If txtSenha.Text = "" Or txtUsuario.Text = "" Then
    MsgBox ("Informe o usuário e senha antes de efetuar o logon."), vbExclamation
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
Quant = FunSóNumeros(App.Major & "." & App.Minor & "." & App.Revision & ".txt")
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select Versao from Versao where Versao IS NOT NULL", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    VersaoNova = FunSóNumeros(IIf(TBAbrir!Versao = "", 0, TBAbrir!Versao))
    If Quant < VersaoNova Then
        'MsgBox ("O sistema está desatualizado e será encerrado."), vbExclamation
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
            VersaoNova = FunSóNumeros(IIf(TBMySQL!Versao = "", 0, TBMySQL!Versao))
            If Quant < VersaoNova Then
                MsgBox ("Existe uma atualização disponível para ser baixada, solicite ao administrador do sistema."), vbInformation
                TBMySQL.Close
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
    MsgBox ("A data do computador está " & MsgTexto & " que a data do servidor, favor arrumar antes de logar no sistema."), vbExclamation
    End
End If

ProcLogonOutSemUtilizacao 'Verifica e apaga logon com a data menor que a atual

If GerArqPastas.FolderExists(Localrel) = False Then
    If MsgBox("Não foi encontrado o caminho " & Localrel & " onde está armazenado os relatórios, se efetuar o login desta forma os relatórios não vão funcionar. " & vbCrLf & "Deseja prosseguir mesmo assim?", vbQuestion + vbYesNo) = vbNo Then End
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
        MsgBox ("Não será possível efetuar o logon, pois o usuário " & txtUsuario.Text & " está bloqueado."), vbExclamation
        TBUsuarios.Close
        Exit Sub
    End If
    If IsNull(TBUsuarios!DtExpiracao) = False And TBUsuarios!DtExpiracao <> "" Then
        If TBUsuarios!DtExpiracao <= Date Then
            MsgBox ("A data de utilização do Caprind para o usuario " & pubUsuario & " venceu em " & Format(TBUsuarios!DtExpiracao, "dd/mm/yy") & "."), vbExclamation
            TBUsuarios.Close
            Exit Sub
        End If
    End If
Else
    MsgBox ("Nome de usuário ou senha inválida."), vbExclamation
    txtSenha.SetFocus
    TBUsuarios.Close
    Exit Sub
End If
TBUsuarios.Close

'Criar este codigo para caminho dos arquivos avi
If Localrel <> "" Then Caminho = Localrel

If TemInternet = True And ErroDriverMYSQL = False Then FunValidarCliente Else FunValidarClienteSemInternet
FunLogonIn True
Unload Me
frmMDI.Show
Unload frmabertura

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcInfLicenca()
On Error GoTo tratar_erro

'qtde_solicitada = ""
'Mensagem1:
'qtde_solicitada = InputBox("Favor informar o número de licença.")
'If qtde_solicitada = "" Then Exit Sub
'If IsNumeric(qtde_solicitada) = False Then
'    MsgBox ("Só é permitido número neste campo."), vbExclamation
'    GoTo Mensagem1
'End If
'Qtde = qtde_solicitada
'If Qtde <= 0 Then
'    MsgBox ("So é permitido número maior que 0."), vbExclamation
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
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Sub ProcSair()
On Error GoTo tratar_erro

If MsgBox("Deseja realmente abandonar o Caprind?", vbQuestion + vbYesNo, "Caprind") = vbYes Then
    End
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Form_Activate()
On Error GoTo tratar_erro

CompactarRepararBanco = False
With cmbBanco
    .Clear
    If Nome_banco <> "" Then
        .AddItem Nome_banco
        .Text = Nome_banco
    End If
    If Nome_banco1 <> "" Then .AddItem Nome_banco1
    If Nome_banco2 <> "" Then .AddItem Nome_banco2
    If Nome_banco3 <> "" Then .AddItem Nome_banco3
End With

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro
   
Select Case KeyCode
    Case vbKeyF8: Cmd_novo_local_bd_Click
    'Case vbKeyF11: InfLicenca = True
    Case vbKeyF9: Cmd_chat_Click
    Case vbKeyF11:
        If Atualizando = True Then Exit Sub
        If MsgBox("Deseja realmente reiniciar a utilização do banco de dados?", vbQuestion + vbYesNo) = vbYes Then
            If InputBox("Informe a primeira senha que tenha pemissão para realizar esta operação.") = "280362" Then
                If InputBox("Informe a segunda senha que tenha pemissão para realizar esta operação.") = "pro0902loc" Then ProcApagarDadosBD
            End If
        End If
    Case vbKeyReturn: If cmdAcessar.Enabled = True Then cmdAcessar_Click
    'Case vbKeyF1: Cmd_ajuda_Click
    Case vbKeyEscape: cmdSair_Click
End Select

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

lblVersaoatual.Caption = "v" & App.Major & "." & App.Minor & "." & App.Revision
lblano.Caption = "Copyright 1999 - " & Year(Date) & " Caprind Sistemas ®. Todos os direitos reservados."

CompactarRepararBanco = False
ProcVerifAtualizacao 'Verifica versão do sistema para atualização

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
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
MsgBox ("Banco de dados reiniciado com sucesso."), vbInformation

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
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
        MsgBox ("O sistema Caprind está desatualizado e será atualizado automaticamente."), vbInformation
        NomeCampo = "SincCaprind.exe na pasta " & CaminhoNovo
        Shell CaminhoNovo & "\SincCaprind.exe", vbNormalFocus
        End
    End If
End If
1:

Exit Sub
tratar_erro:
    If Err.Number = "53" Then
        MsgBox ("Não será possível atualizar o sistema, pois não foi encontrado o arquivo " & NomeCampo & "."), vbExclamation
        GoTo 1
    End If
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcVerifAlteracaoData()
On Error GoTo tratar_erro

If ActiveLock1.LastRunDate > Now Then
    MsgBox ("Ocorreu uma alteração na data do sistema operacional, o sistema Caprind será encerrado."), vbCritical
    End
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro
  
CompactarRepararBanco = False
frmLogon.Refresh
'frmLogon.Frame1.Refresh

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub txtSenha_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus txtSenha

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub txtUsuario_Change()
On Error GoTo tratar_erro

Call FunTextoMaiusculoDigitar(txtUsuario)

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub txtUsuario_LostFocus()
On Error GoTo tratar_erro
    
If txtUsuario <> "" Then txtUsuario.Text = UCase(txtUsuario.Text)

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub
