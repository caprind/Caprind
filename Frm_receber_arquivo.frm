VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form Frm_receber_arquivo 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Centralziar na Tela
   Begin DrawSuite2022.USGroupBox Download 
      Height          =   2535
      Left            =   -30
      TabIndex        =   39
      Top             =   0
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   4471
      BorderColor     =   14404026
      Caption         =   "Recebimento de arquivos"
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
      Begin VB.PictureBox cur_hst 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5E6DC&
         BorderStyle     =   0  'Nenhum
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   165
         ScaleHeight     =   255
         ScaleWidth      =   4275
         TabIndex        =   44
         Top             =   1650
         Width           =   4275
         Begin VB.PictureBox cur_bar 
            BackColor       =   &H00BE8342&
            Height          =   255
            Left            =   0
            ScaleHeight     =   195
            ScaleWidth      =   75
            TabIndex        =   45
            Top             =   0
            Width           =   135
         End
      End
      Begin DrawSuite2022.USButton Cmd_Receber 
         Height          =   315
         Left            =   3960
         TabIndex        =   40
         Top             =   1230
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   556
         BorderColor     =   14404026
         BorderColorDown =   11632444
         BorderColorOver =   11632444
         Caption         =   "Receber arquivo"
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
         PicAlign        =   3
         PicSize         =   2
         PicSizeH        =   24
         PicSizeW        =   24
         State           =   3
      End
      Begin DrawSuite2022.USButton cmd_Caminho 
         Height          =   315
         Left            =   3540
         TabIndex        =   41
         Top             =   1230
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
      Begin DrawSuite2022.USLabel USLabel1 
         Height          =   195
         Left            =   330
         Top             =   450
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   344
         Caption         =   "Arquivo anexo..."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin DrawSuite2022.USTextBoxEx txt_Arquivo 
         Height          =   315
         Left            =   180
         TabIndex        =   42
         Top             =   630
         Width           =   4275
         _ExtentX        =   7541
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
      Begin DrawSuite2022.USTextBoxEx txt_Caminho 
         Height          =   315
         Left            =   150
         TabIndex        =   43
         Top             =   1230
         Width           =   3405
         _ExtentX        =   6006
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
      Begin DrawSuite2022.USLabel USLabel2 
         Height          =   195
         Left            =   300
         Top             =   1050
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   344
         Caption         =   "Salvar em"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   240
         Top             =   2580
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComDlg.CommonDialog CommonDialog2 
         Left            =   1530
         Top             =   2640
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Modo de transferência"
      Height          =   735
      Left            =   0
      TabIndex        =   29
      Top             =   7560
      Width           =   3255
      Begin VB.CommandButton ascii_btn 
         Caption         =   "Ascii"
         Height          =   375
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton binary_btn 
         Caption         =   "Binary"
         Height          =   375
         Left            =   1920
         TabIndex        =   30
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Progress"
      Height          =   1065
      Left            =   30
      TabIndex        =   21
      Top             =   1470
      Width           =   3615
      Begin VB.CommandButton RETR_btn 
         Caption         =   "Download de arquivo"
         Height          =   375
         Left            =   120
         TabIndex        =   37
         Top             =   510
         Width           =   3375
      End
      Begin VB.PictureBox kftp 
         Height          =   600
         Left            =   0
         ScaleHeight     =   540
         ScaleWidth      =   3135
         TabIndex        =   38
         Top             =   120
         Width           =   3195
      End
      Begin VB.Label lastcommand_txt 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   120
         TabIndex        =   32
         Top             =   600
         Width           =   45
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Diretorios"
      Height          =   2445
      Left            =   3660
      TabIndex        =   12
      Top             =   30
      Width           =   3615
      Begin VB.ListBox dir_lst 
         Height          =   1815
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   3375
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Comandos"
      Height          =   4935
      Left            =   0
      TabIndex        =   2
      Top             =   2520
      Width           =   3255
      Begin VB.TextBox rnto_txt 
         Height          =   285
         Left            =   1680
         TabIndex        =   36
         Text            =   "newname.txt"
         Top             =   3600
         Width           =   1335
      End
      Begin VB.TextBox rnfr_txt 
         Height          =   285
         Left            =   120
         TabIndex        =   35
         Top             =   3600
         Width           =   1335
      End
      Begin VB.TextBox newdir_txt 
         Height          =   285
         Left            =   120
         TabIndex        =   34
         Text            =   "my folder"
         Top             =   1200
         Width           =   3015
      End
      Begin VB.CommandButton ren_btn 
         Caption         =   "Renomear arquivo (RNFR+RNTO)"
         Height          =   375
         Left            =   120
         TabIndex        =   33
         Top             =   3120
         Width           =   3015
      End
      Begin VB.CommandButton size_btn 
         Caption         =   "Size"
         Height          =   375
         Left            =   2610
         TabIndex        =   28
         Top             =   4440
         Width           =   525
      End
      Begin VB.CommandButton backCWD_btn 
         Caption         =   ".."
         Height          =   375
         Left            =   2760
         TabIndex        =   27
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton chmod_btn 
         Caption         =   "Acesso total (CHMOD 777)"
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   2160
         Width           =   3015
      End
      Begin VB.CommandButton RMD_btn 
         Caption         =   "Removee diretório (RMD)"
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   1680
         Width           =   3015
      End
      Begin VB.CommandButton STOR_btn 
         Caption         =   "Upload de arquivo (STOR)"
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   3960
         Width           =   3015
      End
      Begin VB.CommandButton DELE_btn 
         Caption         =   "Apagar arquivo (DELE)"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   2640
         Width           =   3015
      End
      Begin VB.CommandButton MKD_btn 
         Caption         =   "Criar novo diretório (MKD)"
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   3015
      End
      Begin VB.CommandButton CWD_btn 
         Caption         =   "Mudar diretório de trabalho(CWD)"
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Lista de arquivos"
      Height          =   3495
      Left            =   3360
      TabIndex        =   1
      Top             =   2280
      Width           =   3615
      Begin VB.ListBox files_lst 
         Height          =   2985
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   3375
      End
   End
   Begin VB.Frame settings_frm 
      Caption         =   "Configurações"
      Height          =   2535
      Left            =   0
      TabIndex        =   0
      Top             =   2130
      Width           =   3255
      Begin VB.CommandButton abort_btn 
         Caption         =   "Abortar"
         Height          =   375
         Left            =   1320
         TabIndex        =   24
         Top             =   2040
         Width           =   675
      End
      Begin VB.TextBox local_txt 
         Height          =   285
         Left            =   2160
         TabIndex        =   23
         Text            =   "c:\"
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton disconnect_btn 
         Caption         =   "Desconectar"
         Height          =   375
         Left            =   2040
         TabIndex        =   17
         Top             =   2040
         Width           =   1095
      End
      Begin VB.CheckBox passive_chk 
         Caption         =   "Passive Mode"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   1320
         Width           =   3015
      End
      Begin VB.CommandButton connect_btn 
         Caption         =   "Conectar"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox pass_txt 
         Height          =   285
         Left            =   1440
         TabIndex        =   10
         Text            =   "pro0802loc"
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox login_txt 
         Height          =   285
         Left            =   1440
         TabIndex        =   8
         Text            =   "procamonline"
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox port_txt 
         Height          =   285
         Left            =   2640
         TabIndex        =   6
         Text            =   "21"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox server_txt 
         Height          =   285
         Left            =   840
         TabIndex        =   4
         Text            =   "ftp.procamonline.com.br"
         Top             =   240
         Width           =   1245
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Diretório local download:"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   1680
         Width           =   1740
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Password:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Login :"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Port :"
         Height          =   195
         Left            =   2160
         TabIndex        =   5
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Servidor :"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   675
      End
   End
End
Attribute VB_Name = "Frm_receber_arquivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'Option Explicit
'Private Type OPENFILENAME
'    lStructSize As Long
'    hwndOwner As Long
'    hInstance As Long
'    lpstrFilter As String
'    lpstrCustomFilter As String
'    nMaxCustFilter As Long
'    nFilterIndex As Long
'    lpstrFile As String
'    nMaxFile As Long
'    lpstrFileTitle As String
'    nMaxFileTitle As Long
'    lpstrInitialDir As String
'    lpstrTitle As String
'    flags As Long
'    nFileOffset As Integer
'    nFileExtension As Integer
'    lpstrDefExt As String
'    lCustData1 as Long
'    lpfnHook As Long
'    lpTemplateName As String
'End Type
'
'Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
'Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'Private Const OFN_ALLOWMULTISELECT = 512
'Private Const OFN_EXPLORER = 524288
'Private Const OFN_FILEMUSTEXIST = 4096
'Private Const OFN_HIDEREADONLY = 4
'Private Sub abort_btn_Click()
''If Not kftp.Abort Then
'    'usMsgbox kftp.LastError
''End If
'End Sub
'
'Private Sub ascii_btn_Click()
''kftp.ChangeTransfertMode Ascii
'End Sub
'
'Private Sub backCWD_btn_Click()
''If Not kftp.ChangeWorkingDir("../") Then
'    'usMsgbox kftp.LastError
''Else
'    'List
''End If
'
'End Sub
'
'Private Sub binary_btn_Click()
''kftp.ChangeTransfertMode Binary
'End Sub
'Private Sub chmod_btn_Click()
'If files_lst.ListIndex < 0 Then Exit Sub
''If Not kftp.CHMOD(files_lst.List(files_lst.ListIndex), 777) Then
'    'usMsgbox kftp.LastError
''Else
'    'List
''End If
'End Sub
'
'Private Sub cmd_Caminho_Click()
'On Error GoTo tratar_erro
'
'With CommonDialog1
'    .Filter = "(*.*) | *.*"
'    .InitDir = App.Path
'    .DefaultExt = "*.*"
'    .filename = txt_Arquivo
'    .ShowSave
'     'Anexo = .filename
'End With
'
''If Anexo = "" Then
'    'usMsgbox ("Você não escolheu nenhum caminho para salvar o arquivo!"), vbInformation, "CAPRIND v5.0"
'    'Cmd_Receber.Enabled = False
'    'Exit Sub
''End If
'
'    'txt_Caminho = Anexo
'    'Cmd_Receber.Enabled = True
'
'
'Exit Sub
'tratar_erro:
'    usMsgbox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
'    Exit Sub
'End Sub
'
'Private Sub Cmd_Receber_Click()
'
'If txt_Arquivo.Text = "" Then Exit Sub
'
'txt_Caminho.Text = Trim(txt_Caminho.Text)
'If Len(txt_Caminho.Text) < 2 Then Exit Sub
'
''If Not kftp.DownloadFile(txt_Arquivo, txt_Caminho, 0, passive_chk.Value) Then
'    'usMsgbox kftp.LastError
''Else
''usMsgbox ("Arquivo recebido com sucesso!"), vbInformation, "CAPRIND v5.0"
''ProcReceberArquivo (IDAtendimento)
''End If
'
'cur_bar.Width = 0
'
'End Sub
'
'Private Sub connect_btn_Click()
'
''If Not kftp.Connect(server_txt.Text, login_txt.Text, pass_txt.Text, Val(port_txt.Text)) Then
'    'usMsgbox kftp.LastError
''Else
'    'List
''End If
'
'End Sub
'Public Function Prog(prog_bar As PictureBox, prog_hst As PictureBox, pc As Byte)
'prog_bar.Width = Int((pc / 100) * prog_hst.Width)
'prog_bar.Visible = (pc <> 0)
'DoEvents
'prog_bar.Refresh
'End Function
'Private Sub List()
'Dim i As Integer, j As Integer
'Dim ttab() As String
'Dim entry() As String
'Dim link() As String
'Dim Data1 as String
'Dim NewEntry As String
'
''dir_lst.Clear
''files_lst.Clear
''
''
''If Not kftp.ListContent(data, passive_chk.Value, VBCRLF_Separated_Preformatted_Array) Then
''    usMsgbox kftp.LastError
''    Exit Sub
''End If
''
''
''ttab = Split(data, vbCrLf)
''For i = LBound(ttab) To UBound(ttab)
''    entry = Split(ttab(i), Space(1))
''    NewEntry = vbNullString
''    For j = 1 To UBound(entry)
''        NewEntry = NewEntry & Space(1) & entry(j)
''    Next j
''
''    If Split(ttab(i), Space(1))(0) = "DIR" Then
''        dir_lst.AddItem Trim(NewEntry)
''    Else
''        If Split(ttab(i), Space(1))(0) = "LINK" Then
''            If InStr(1, NewEntry, "->") > 0 Then
''                link = Split(NewEntry, "->")
''                NewEntry = link(0)
''            End If
''            dir_lst.AddItem Trim(NewEntry)
''        Else
''            files_lst.AddItem Trim(NewEntry)
''        End If
''    End If
''Next i
'
'End Sub
'Private Sub CWD_btn_Click()
'If dir_lst.ListIndex < 0 Then Exit Sub
'
''If Not kftp.ChangeWorkingDir("web/Arquivos") Then
'
''If Not kftp.ChangeWorkingDir(dir_lst.List(dir_lst.ListIndex)) Then
'    'usMsgbox kftp.LastError
''Else
'
'    'List
'
''End If
'End Sub
'Private Sub DELE_btn_Click()
'If files_lst.ListIndex < 0 Then Exit Sub
''If Not kftp.DeleteFile(files_lst.List(files_lst.ListIndex)) Then
''    usMsgbox kftp.LastError
''Else
''    List
''End If
'End Sub
'Private Sub dir_lst_DblClick()
'CWD_btn_Click
'End Sub
'
'Private Sub disconnect_btn_Click()
''If Not kftp.Disconnect Then
''    usMsgbox kftp.LastError
''End If
'End Sub
'
'Private Sub files_lst_Click()
'If files_lst.ListIndex < 0 Then Exit Sub
'rnfr_txt.Text = files_lst.List(files_lst.ListIndex)
'End Sub
'
'Private Sub Form_Load()
'Call Prog(cur_bar, cur_hst, 0)
'kftp.DisableRESTCommand
'kftp.LogFile = "kftp.txt"
'
'connect_btn_Click
'txt_Arquivo = Nome_anexo
'
'    If Not kftp.ChangeWorkingDir("web/Arquivos") Then
'        usMsgbox kftp.LastError
'    Else
'        List
'    End If
'
'End Sub
'
'Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'kftp.Disconnect
'End Sub
'Private Sub kftp_Command(Message As String)
'lastcommand_txt.Caption = Message
'End Sub
'Private Sub kftp_Progress(percent As Byte)
'Prog cur_bar, cur_hst, percent
'End Sub
'Private Sub MKD_btn_Click()
'If Len(newdir_txt.Text) = 0 Then
'    usMsgbox "Empty name for new directory"
'    Exit Sub
'End If
'
'If Not kftp.CreateNewDirectory(newdir_txt.Text) Then
'    usMsgbox kftp.LastError
'Else
'    List
'End If
'End Sub
'
'Private Sub ren_btn_Click()
'If Len(rnfr_txt.Text) * Len(rnto_txt.Text) = 0 Then Exit Sub
'If Not kftp.RenameFile(rnfr_txt.Text, rnto_txt.Text) Then
'    usMsgbox kftp.LastError
'Else
'    List
'End If
'
'End Sub
'
'Private Sub RETR_btn_Click()
'
'If Nome_anexo = "" Then Exit Sub
'local_txt.Text = Trim(local_txt.Text)
'If Len(local_txt.Text) < 2 Then Exit Sub
'If Dir(local_txt.Text, vbDirectory) = vbNullString Then Exit Sub
'If Right(local_txt.Text, 1) <> "\" Then local_txt.Text = local_txt.Text & "\"
'
''If Not kftp.DownloadFile(files_lst.List(files_lst.ListIndex), local_txt.Text & files_lst.List(files_lst.ListIndex), 0, passive_chk.Value) Then
'If Not kftp.DownloadFile(Nome_anexo, local_txt.Text & Nome_anexo, 0, passive_chk.Value) Then
'    usMsgbox kftp.LastError
'End If
'cur_bar.Width = 0
'
'End Sub
'
'Private Sub RMD_btn_Click()
'If dir_lst.ListIndex < 0 Then Exit Sub
'If Not kftp.DeleteDirectory(dir_lst.List(dir_lst.ListIndex)) Then
'    usMsgbox kftp.LastError
'Else
'    List
'End If
'End Sub
'Private Sub size_btn_Click()
'Dim size As Long
'If files_lst.ListIndex < 0 Then Exit Sub
'local_txt.Text = Trim(local_txt.Text)
'If Len(local_txt.Text) < 2 Then Exit Sub
'If Dir(local_txt.Text, vbDirectory) = vbNullString Then Exit Sub
'If Right(local_txt.Text, 1) <> "\" Then local_txt.Text = local_txt.Text & "\"
'
'If Not kftp.GetFileSize(size, files_lst.List(files_lst.ListIndex), passive_chk.Value) Then
'    usMsgbox kftp.LastError
'Else
'    usMsgbox size & " bytes"
'End If
'End Sub
'
'Private Sub STOR_btn_Click()
'Dim a As Long
'Dim ofn As OPENFILENAME
'
'ofn.hwndOwner = Me.hwnd
'ofn.hInstance = App.hInstance
'ofn.lpstrFilter = "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
'ofn.lpstrFile = Space$(511)
'ofn.nMaxFile = 512
'ofn.lpstrFileTitle = Space$(511)
'ofn.nMaxFileTitle = 512
'ofn.lpstrInitialDir = App.Path
'ofn.lpstrTitle = App.ProductName
'ofn.flags = OFN_EXPLORER Or OFN_FILEMUSTEXIST Or OFN_HIDEREADONLY
'ofn.lStructSize = Len(ofn)
'
'Me.Visible = False
'a = GetOpenFileName(ofn)
'Me.Visible = True
'
'If a = 0 Then Exit Sub
'
'If Not kftp.UploadFile(Trim(ofn.lpstrFileTitle), Trim(ofn.lpstrFile), passive_chk.Value) Then
'    usMsgbox kftp.LastError
'Else
'    List
'End If
'End Sub
