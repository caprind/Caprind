VERSION 5.00
Object = "{84147065-0227-424E-827F-9E79B1DA5D8B}#21.0#0"; "kftp.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form Frm_atualizacao_sistema 
   BackColor       =   &H00F9F9F9&
   BorderStyle     =   0  'None
   Caption         =   "Suporte - Atualização | Caprind e Gerprod"
   ClientHeight    =   7725
   ClientLeft      =   6615
   ClientTop       =   -75
   ClientWidth     =   9360
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Frm_atualizacao_sistema.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7725
   ScaleWidth      =   9360
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   9360
      _ExtentX        =   16510
      _ExtentY        =   741
      CaptionDelimiter=   "|"
      CaptionOnCenter =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Icon            =   "Frm_atualizacao_sistema.frx":000C
   End
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   12
      Top             =   7320
      Width           =   9360
      _ExtentX        =   16510
      _ExtentY        =   714
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Check1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2790
      TabIndex        =   3
      Top             =   2565
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Check1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2790
      TabIndex        =   2
      Top             =   3075
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Check1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2790
      TabIndex        =   1
      Top             =   2805
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CheckBox passive_chk 
      Caption         =   "Passive Mode"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3000
      TabIndex        =   11
      Top             =   4530
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.PictureBox cur_hst 
      Appearance      =   0  'Flat
      BackColor       =   &H00F5E6DC&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2790
      ScaleHeight     =   315
      ScaleWidth      =   3645
      TabIndex        =   9
      Top             =   3390
      Width           =   3645
      Begin VB.PictureBox cur_bar 
         BackColor       =   &H00BE8342&
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
         Left            =   0
         ScaleHeight     =   255
         ScaleWidth      =   75
         TabIndex        =   10
         Top             =   0
         Width           =   135
      End
   End
   Begin KFTPActiveX.kftp kftp 
      Height          =   600
      Left            =   5790
      TabIndex        =   8
      Top             =   1260
      Visible         =   0   'False
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   1058
   End
   Begin VB.TextBox Txt_zip 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5160
      TabIndex        =   7
      Top             =   1290
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   4710
      Top             =   1260
   End
   Begin DrawSuite2022.USAlphaImage imageTeamViewer 
      Height          =   1500
      Left            =   1170
      Top             =   3750
      Width           =   5940
      _ExtentX        =   10478
      _ExtentY        =   2646
      Image           =   "Frm_atualizacao_sistema.frx":0028
      Props           =   5
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Executando sistema de atualização"
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
      Height          =   225
      Left            =   3060
      TabIndex        =   6
      Top             =   3060
      Visible         =   0   'False
      Width           =   3465
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Descompactando arquivos"
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
      Height          =   225
      Left            =   3060
      TabIndex        =   5
      Top             =   2790
      Visible         =   0   'False
      Width           =   3465
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Baixando arquivos"
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
      Height          =   225
      Left            =   3060
      TabIndex        =   4
      Top             =   2520
      Width           =   3465
   End
   Begin VB.Label Lbl_progresso 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1200
      Left            =   2790
      TabIndex        =   0
      Top             =   3720
      Width           =   3645
   End
   Begin VB.Image Image1 
      Height          =   6900
      Left            =   0
      Picture         =   "Frm_atualizacao_sistema.frx":52477
      Top             =   480
      Width           =   9465
   End
   Begin VB.Image Image2 
      Height          =   6900
      Left            =   0
      Picture         =   "Frm_atualizacao_sistema.frx":127399
      Top             =   480
      Visible         =   0   'False
      Width           =   9465
   End
End
Attribute VB_Name = "Frm_atualizacao_sistema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
On Error GoTo tratar_erro

    If Atualizacao_TeamViewerQS = True Then
            USForm1.Caption = "Suporte | Download - Team ViewerQS"
            Image1.Visible = False
            Image2.Visible = False
            imageTeamViewer.Visible = True
    End If

    If Atualizacao_TeamViewer = True Then
            USForm1.Caption = "Suporte | Download - Team Viewer"
            Image1.Visible = False
            Image2.Visible = False
            imageTeamViewer.Visible = True
    End If
    
    If Atualizacao_versao = True Then
            USForm1.Caption = "Suporte | Atualização servidor"
            imageTeamViewer.Visible = False
            Image3.Visible = True
            Image4.Visible = True
    End If

Call FunProgBarKFTP(cur_bar, cur_hst, 0)

With kftp
    .DisableRESTCommand
    If Atualizacao_versao = True Then
        If FunConectaKFTP(kftp, "/public_html/Sistemas/Atualizacao", True) = False Then
            Unload Me
            Exit Sub
        End If
    End If
    
    If Atualizacao_TeamViewer = True Or Atualizacao_TeamViewerQS = True Then
        If FunConectaKFTP(kftp, "public_html/Arquivos", True) = False Then
            Unload Me
            Exit Sub
        End If
    End If
End With

InicializaZip Me, Txt_zip
Timer1.Enabled = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Timer1_Timer()
On Error GoTo tratar_erro

ProcIniciaAtualizacao

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcIniciaAtualizacao()
On Error GoTo tratar_erro

With GerArqPastas
            If Atualizacao_TeamViewer = True Then
                                     
                    If FileOrDirExists(App.Path & "\TeamViewer_Setup.exe") = True Then
                        FileDelete (App.Path & "\TeamViewer_Setup.exe")
                    End If
                    
                    
                    If FunDownloadKFTP(kftp, "TeamViewer_Setup.exe", App.Path & "\TeamViewer_Setup.exe") = False Then
                        Unload Me
                        Timer1.Enabled = False
                        Exit Sub
                    End If
                    Label1.Visible = False
                    Check1.Visible = True
                    Label2.Visible = True
                    Check2.Visible = True
                    Label3.Visible = True
                    DS.FileExecute (App.Path & "\TeamViewer_Setup.exe")
                    Check3.Visible = True
                    Timer1.Enabled = False
                ElseIf Atualizacao_TeamViewerQS = True Then
                                     
                    If FileOrDirExists(App.Path & "\TeamViewerQS.exe") = True Then
                        FileDelete (App.Path & "\TeamViewerQS.exe")
                    End If
                    
                    
                    If FunDownloadKFTP(kftp, "TeamViewerQS.exe", App.Path & "\TeamViewerQS.exe") = False Then
                        Unload Me
                        Timer1.Enabled = False
                        Exit Sub
                    End If
                    Label1.Visible = False
                    Check1.Visible = True
                    Label2.Visible = True
                    Check2.Visible = True
                    Label3.Visible = True
                    DS.FileExecute (App.Path & "\TeamViewerQS.exe")
                    Check3.Visible = True
                    Timer1.Enabled = False
                
                Else
                    If .FileExists(App.Path & "\" & FamiliaAntiga) = True Then FileDelete (App.Path & "\" & FamiliaAntiga)
                    If .FileExists(App.Path & "\" & Replace(FamiliaAntiga, ".zip", ".exe")) = True Then FileDelete (App.Path & "\" & Replace(FamiliaAntiga, ".zip", ".exe"))
                    
                    If FunDownloadKFTP(kftp, FamiliaAntiga, App.Path & "\" & FamiliaAntiga) = False Then
                        Unload Me
                        Timer1.Enabled = False
                        Exit Sub
                    End If
                    Check1.Visible = True
                    
                    Label2.Visible = True
                    If Check1.Value = 1 Then Descompacta App.Path & "\" & FamiliaAntiga, "*.*", App.Path, True
                    If Left(Lbl_progresso, 4) <> "Erro" Then Check2.Visible = True
                    
                    If Check2.Visible = True Then
                        Label3.Visible = True
                        DS.FileExecute (App.Path & "\" & Replace(FamiliaAntiga, ".zip", ".exe"))
                        Check3.Visible = True
                    End If
                    
                    Timer1.Enabled = False
                    
                    If .FileExists(App.Path & "\" & FamiliaAntiga) = True Then FileDelete (App.Path & "\" & FamiliaAntiga)
                    If .FileExists(App.Path & "\" & Replace(FamiliaAntiga, ".zip", ".exe")) = True Then FileDelete (App.Path & "\" & Replace(FamiliaAntiga, ".zip", ".exe"))
                End If
End With
kftp.Disconnect
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_zip_Change()
On Error GoTo tratar_erro

Lbl_progresso = TipoAção(Val(GetAction(Txt_zip))) & " "
Lbl_progresso = Lbl_progresso & GetFileName(Txt_zip) & " -> "
Lbl_progresso = Lbl_progresso & GetPercentComplete(Txt_zip) & "%"
DoEvents

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub kftp_Progress(percent As Byte)
On Error GoTo tratar_erro

FunProgBarKFTP cur_bar, cur_hst, percent

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
