VERSION 5.00
Begin VB.Form frmLicensa 
   BackColor       =   &H8007000D&
   Caption         =   "- Administrador de licensa do Caprind v3.01."
   ClientHeight    =   5085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7080
   ClipControls    =   0   'False
   ForeColor       =   &H00004040&
   Icon            =   "FrmLicensa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Palette         =   "FrmLicensa.frx":030A
   PaletteMode     =   2  'Custom
   Picture         =   "FrmLicensa.frx":1AD67C
   ScaleHeight     =   5085
   ScaleWidth      =   7080
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtserial 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2160
      TabIndex        =   4
      ToolTipText     =   "Numero de série do Caprind 2002."
      Top             =   2640
      Width           =   2600
   End
   Begin VB.TextBox txtsenha 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      IMEMode         =   3  'DISABLE
      Left            =   2640
      MaxLength       =   10
      TabIndex        =   2
      ToolTipText     =   "Senha de liberação do Caprind 2002."
      Top             =   4320
      Width           =   1692
   End
   Begin VB.CommandButton cmdsair 
      Caption         =   "Sair"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   6360
      TabIndex        =   1
      ToolTipText     =   "Fecha o gerenciador de licensas."
      Top             =   4680
      Width           =   732
   End
   Begin VB.CommandButton cmdlibera 
      Caption         =   "Libera"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   6360
      TabIndex        =   0
      ToolTipText     =   "Libera o aplicativo por mais dez utilizações."
      Top             =   4200
      Width           =   732
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Digite a senha de liberação do Caprind  v3.01."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2280
      TabIndex        =   6
      Top             =   480
      Width           =   2655
   End
   Begin VB.Label lblserial 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nº série :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1200
      TabIndex        =   5
      Top             =   2640
      Width           =   1110
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Senha :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1800
      TabIndex        =   3
      Top             =   4320
      Width           =   924
   End
End
Attribute VB_Name = "frmLicensa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdlibera_Click()
On Error GoTo tratar_erro

If txtSenha.Text = "" Then
    USMsgBox ("Informe a senha antes de liberar."), vbInformation, "CAPRIND v5.0"
    txtSenha.SetFocus
    Exit Sub
End If

Open "C:\Procam\licensa.dat" For Random As #1 Len = 60
Get #1, 1, licensa
Select Case licensa.Serie
    Case 1
        If licensa.Senha = txtSenha.Text Then
            licensa.Contador = 1
            Put #1, 1, licensa
            RegModif = False
            Close #1
            USMsgBox ("Liberação efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
            cmdlibera.Enabled = False
        Else
            USMsgBox ("Senha incorreta."), vbInformation, "CAPRIND v5.0"
            Close #1
            txtSenha.Text = ""
            txtSenha.SetFocus
            Exit Sub
        End If
    Case 2
        If licensa.Senha = txtSenha.Text Then
            licensa.Contador = 2
            Put #1, 1, licensa
            RegModif = False
            Close #1
            USMsgBox ("Liberação efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
            cmdlibera.Enabled = False
        Else
            USMsgBox ("Senha incorreta."), vbInformation, "CAPRIND v5.0"
            Close #1
            txtSenha.Text = ""
            txtSenha.SetFocus
            Exit Sub
        End If
    Case 3
        If licensa.Senha = txtSenha.Text Then
            licensa.Contador = 3
            Put #1, 1, licensa
            RegModif = False
            Close #1
            USMsgBox ("Liberação efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
            cmdlibera.Enabled = False
        Else
            USMsgBox ("Senha incorreta."), vbInformation, "CAPRIND v5.0"
            Close #1
            txtSenha.Text = ""
            txtSenha.SetFocus
            Exit Sub
        End If
    Case 4
        If licensa.Senha = txtSenha.Text Then
            licensa.Contador = 4
            Put #1, 1, licensa
            RegModif = False
            Close #1
            USMsgBox ("Liberação efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
            cmdlibera.Enabled = False
        Else
            USMsgBox ("Senha incorreta."), vbInformation, "CAPRIND v5.0"
            Close #1
            txtSenha.Text = ""
            txtSenha.SetFocus
            Exit Sub
        End If
    Case 5
        If licensa.Senha = txtSenha.Text Then
            licensa.Contador = 5
            Put #1, 1, licensa
            RegModif = False
            Close #1
            USMsgBox ("Liberação efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
            cmdlibera.Enabled = False
        Else
            USMsgBox ("Senha incorreta."), vbInformation, "CAPRIND v5.0"
            Close #1
            txtSenha.Text = ""
            txtSenha.SetFocus
            Exit Sub
        End If
    Case 6
        If licensa.Senha = txtSenha.Text Then
            licensa.Contador = 6
            Put #1, 1, licensa
            RegModif = False
            Close #1
            USMsgBox ("Liberação efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
            cmdlibera.Enabled = False
        Else
            USMsgBox ("Senha incorreta."), vbInformation, "CAPRIND v5.0"
            Close #1
            txtSenha.Text = ""
            txtSenha.SetFocus
            Exit Sub
        End If
    Case 7
        If licensa.Senha = txtSenha.Text Then
            licensa.Contador = 7
            Put #1, 1, licensa
            RegModif = False
            Close #1
            USMsgBox ("Liberação efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
            cmdlibera.Enabled = False
        Else
            USMsgBox ("Senha incorreta."), vbInformation, "CAPRIND v5.0"
            Close #1
            txtSenha.Text = ""
            txtSenha.SetFocus
            Exit Sub
        End If
    Case 8
        If licensa.Senha = txtSenha.Text Then
            licensa.Contador = 8
            Put #1, 1, licensa
            RegModif = False
            Close #1
            USMsgBox ("Liberação efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
            cmdlibera.Enabled = False
        Else
            USMsgBox ("Senha incorreta."), vbInformation, "CAPRIND v5.0"
            Close #1
            txtSenha.Text = ""
            txtSenha.SetFocus
            Exit Sub
        End If
    Case 9
        If licensa.Senha = txtSenha.Text Then
            licensa.Contador = 9
            Put #1, 1, licensa
            RegModif = False
            Close #1
            USMsgBox ("Liberação efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
            cmdlibera.Enabled = False
        Else
            USMsgBox ("Senha incorreta."), vbInformation, "CAPRIND v5.0"
            Close #1
            txtSenha.Text = ""
            txtSenha.SetFocus
            Exit Sub
        End If
    Case 9999
        If licensa.Senha = txtSenha.Text Then
            licensa.Contador = 9999
            Put #1, 1, licensa
            RegModif = False
            Close #1
            USMsgBox ("Liberação efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
            cmdlibera.Enabled = False
        Else
            USMsgBox ("Senha incorreta."), vbInformation, "CAPRIND v5.0"
            Close #1
            txtSenha.Text = ""
            txtSenha.SetFocus
            Exit Sub
        End If
End Select

'If licensa.serie = 9999 And licensa.Senha = txtsenha.Text Then
    'licensa.contador = 9999
    'Put #1, 1, licensa
    'RegModif = False
    'Close #1
    'usMsgbox ("Liberação efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    'cmdlibera.Enabled = False
'Else
    'usMsgbox (" - SENHA NÃO AUTORIZADA, CHAME A PROCAM PROGRAMAÇÃO C.N.C"), vbCritical, "CAPRIND v5.0"
    'licensa.serie = 1
    'Put #1, 1, licensa
    'RegModif = False
    'Close #1
    'txtsenha.Text = ""
    'txtsenha.SetFocus
    'Exit Sub
'End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdSair_Click()
On Error GoTo tratar_erro

If USMsgBox("Deseja realmente sair do administrador?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    Unload Me
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro
Dim Diretorio 'OK

Open "C:\arq.log" For Random As #2 Len = 300
Get #2, 1, Serial
Serial.Serial = "1JKRL78J0K"
Serial.Texto = "Este arquivo não pode ser apagado."
Put #2, 1, Serial
Close #2
SetAttr "c:\arq.log", vbHidden
Diretorio = Dir("C:\Procam", vbDirectory + vbHidden)
If Diretorio = "" Then
    MkDir "C:\Procam"
End If
Open "C:\Procam\licensa.dat" For Random As #1 Len = 60
Get #1, 1, licensa

Select Case licensa.Serie
    Case 0
        Label2.Caption = "Informe a senha para liberar o Caprind."
        licensa.Senha = "C3DF32EAH7"
        licensa.licensa = "AWSJRTHSW0"
        txtserial.Text = licensa.licensa
        licensa.Serie = 1
    Case 1
        Label2.Caption = "Informe a senha para liberar o Caprind."
        licensa.Senha = "WHRTS680J4"
        licensa.licensa = "SNDJKETUOE"
        txtserial.Text = licensa.licensa
        licensa.Serie = 2
    Case 2
        Label2.Caption = "Informe a senha para liberar o Caprind."
        licensa.Senha = "SHYETH97JL"
        licensa.licensa = "DJEUSLOTPS"
        txtserial.Text = licensa.licensa
        licensa.Serie = 3
    Case 3
        Label2.Caption = "Informe a senha para liberar o Caprind."
        licensa.Senha = "H9EKSL8K9X"
        licensa.licensa = "SJKLRUSNMO"
        txtserial.Text = licensa.licensa
        licensa.Serie = 4
    Case 4
        Label2.Caption = "Informe a senha para liberar o Caprind."
        licensa.Senha = "KMSHE8PKME"
        licensa.licensa = "KDMSNOTHNQ"
        txtserial.Text = licensa.licensa
        licensa.Serie = 5
    Case 5
        Label2.Caption = "Informe a senha para liberar o Caprind."
        licensa.Senha = "K9ND3K8DYC"
        licensa.licensa = "ÇRMDNSO984"
        txtserial.Text = licensa.licensa
        licensa.Serie = 6
    Case 6
        Label2.Caption = "Informe a senha para liberar o Caprind."
        licensa.Senha = "SL8JNFT29K"
        licensa.licensa = "KDMSO03K8F"
        txtserial.Text = licensa.licensa
        licensa.Serie = 7
    Case 7
        Label2.Caption = "Informe a senha para liberar o Caprind."
        licensa.Senha = "K8FISHT9ME"
        licensa.licensa = "KDMS90ELÇ9"
        txtserial.Text = licensa.licensa
        licensa.Serie = 8
    Case 8
        Label2.Caption = "Informe a senha para liberar o Caprind."
        licensa.Senha = "J9N4HSGYR0"
        licensa.licensa = "KDMS74KMO4"
        txtserial.Text = licensa.licensa
        licensa.Serie = 9
    Case 9
        Label2.Caption = "Informe a senha para liberar o Caprind."
        licensa.Senha = "JS8TYHD8JN"
        licensa.licensa = "KDMS85JSPE"
        txtserial.Text = licensa.licensa
        licensa.Serie = 10
    Case 9999
        Label2.Caption = "O Caprind já foi liberado definitivamente."
        txtserial.Enabled = False
        txtSenha.Enabled = False
        cmdlibera.Enabled = False
        Close #1
        Exit Sub
End Select

Put #1, 1, licensa
RegModif = False
Close #1
SetAttr "c:\procam", vbHidden
SetAttr "c:\procam\licensa.dat", vbHidden
txtSenha.Text = ""
txtSenha.SetFocus
Exit Sub

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
