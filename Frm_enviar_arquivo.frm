VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form Frm_enviar_arquivo 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Centralziar na Tela
   Begin DrawSuite2022.USGroupBox Download 
      Height          =   1905
      Left            =   -30
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   3360
      BorderColor     =   14404026
      Caption         =   "Envio de arquivos em anexo..."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   10236447
      GradientColor1  =   16643823
      GradientColor2  =   16115420
      GradientHeader1 =   16643823
      GradientHeader2 =   16181984
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   2760
         Top             =   1980
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.PictureBox cur_hst 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5E6DC&
         BorderStyle     =   0  'Nenhum
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   165
         ScaleHeight     =   255
         ScaleWidth      =   4275
         TabIndex        =   1
         Top             =   1530
         Width           =   4275
         Begin VB.PictureBox cur_bar 
            BackColor       =   &H00BE8342&
            Height          =   255
            Left            =   0
            ScaleHeight     =   195
            ScaleWidth      =   75
            TabIndex        =   2
            Top             =   0
            Width           =   135
         End
      End
      Begin DrawSuite2022.USButton Cmd_enviar 
         Height          =   375
         Left            =   150
         TabIndex        =   3
         Top             =   1080
         Width           =   4305
         _ExtentX        =   7594
         _ExtentY        =   661
         BorderColor     =   14404026
         BorderColorDown =   11632444
         BorderColorOver =   11632444
         Caption         =   "Enviar arquivo"
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColorDown2=   16246986
         GradientColorDown3=   15189380
         GradientColorDown4=   14596208
         GradientColorOver1=   16643560
         GradientColorOver2=   16576988
         GradientColorOver3=   16441780
         GradientColorOver4=   16178091
         State           =   3
      End
      Begin DrawSuite2022.USButton cmd_Caminho 
         Height          =   315
         Left            =   4110
         TabIndex        =   4
         ToolTipText     =   "Localizar arquivo anexo"
         Top             =   690
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   556
         BorderColor     =   14404026
         BorderColorDown =   11632444
         BorderColorOver =   11632444
         Caption         =   "..."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColorDown2=   16246986
         GradientColorDown3=   15189380
         GradientColorDown4=   14596208
         GradientColorOver1=   16643560
         GradientColorOver2=   16576988
         GradientColorOver3=   16441780
         GradientColorOver4=   16178091
      End
      Begin DrawSuite2022.USTextBoxEx txt_Caminho 
         Height          =   315
         Left            =   150
         TabIndex        =   5
         Top             =   690
         Width           =   3945
         _ExtentX        =   6959
         _ExtentY        =   556
         AutoFormatDate  =   -1  'True
         BorderColor     =   14404026
         BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontDisabled {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontOver {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
         MaxLength       =   0
      End
   End
   Begin VB.PictureBox kftp 
      Height          =   600
      Left            =   0
      ScaleHeight     =   540
      ScaleWidth      =   3135
      TabIndex        =   6
      Top             =   0
      Width           =   3195
   End
End
Attribute VB_Name = "Frm_enviar_arquivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmd_Caminho_Click()
On Error GoTo tratar_erro

With CommonDialog1
    .Filter = "(*.*) | *.*"
    .InitDir = App.Path
    .DefaultExt = "*.*"
    .ShowOpen
    Anexo = .filename
    Nome_anexo = .FileTitle
End With

If Anexo = "" Then
    USMsgBox ("Você não escolheu nenhum arquivo para enviar!"), vbInformation, "CAPRIND v5.0"
    Cmd_enviar.Enabled = False
    Exit Sub
End If

    txt_Caminho = Anexo
    Cmd_enviar.Enabled = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_enviar_Click()
On Error GoTo tratar_erro

Arquivo_local = IIf(Anexo = "", "", Anexo)
Arquivo_Site = "/procamonline/web/arquivos/" & Nome_anexo

'If Not kftp.UploadFile(Arquivo_Site, Arquivo_local, False) Then
    'usMsgbox kftp.LastError
'Else
    'ProcEnviarArquivo (IDAtendimento)
    'usMsgbox ("Arquivo enviado com sucesso!")
'End If


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub connect_btn_Click()
On Error GoTo tratar_erro


'If Not kftp.Connect("ftp.procamonline.com.br", "procamonline", "pro0802loc", Val("21")) Then
    'usMsgbox kftp.LastError
'Else
'    List
'End If


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Function FunProg(FunProg_bar As PictureBox, FunProg_hst As PictureBox, pc As Byte)

FunProg_bar.Width = Int((pc / 100) * FunProg_hst.Width)
FunProg_bar.Visible = (pc <> 0)
DoEvents
FunProg_bar.Refresh

End Function

Private Sub DELE_btn_Click()
On Error GoTo tratar_erro

If files_lst.ListIndex < 0 Then Exit Sub
'If Not kftp.DeleteFile(files_lst.List(files_lst.ListIndex)) Then
    'usMsgbox kftp.LastError
'Else
   ' List
'End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

Call FunProg(cur_bar, cur_hst, 0)
'kftp.DisableRESTCommand
'kftp.LogFile = "kftp.txt"
'connect_btn_Click

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo tratar_erro

'kftp.Disconnect

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub kftp_FunProgress(percent As Byte)
On Error GoTo tratar_erro

FunProg cur_bar, cur_hst, percent

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub size_btn_Click()
On Error GoTo tratar_erro

Dim size As Long
If files_lst.ListIndex < 0 Then Exit Sub
local_txt.Text = Trim(local_txt.Text)
If Len(local_txt.Text) < 2 Then Exit Sub
If Dir(local_txt.Text, vbDirectory) = vbNullString Then Exit Sub
If Right(local_txt.Text, 1) <> "\" Then local_txt.Text = local_txt.Text & "\"

'If Not kftp.GetFileSize(size, files_lst.List(files_lst.ListIndex), passive_chk.Value) Then
    'usMsgbox kftp.LastError
'Else
    'usMsgbox size & " bytes"
'End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

