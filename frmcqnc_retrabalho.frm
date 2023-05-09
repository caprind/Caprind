VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmcqnc_retrabalho 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "PCP - Não conformidade - Retrabalho"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   12675
   ClipControls    =   0   'False
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
   Icon            =   "frmcqnc_retrabalho.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   12675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3480
      Top             =   210
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Instruções de trabalho"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2505
      Left            =   55
      TabIndex        =   10
      Top             =   990
      Width           =   12555
      Begin VB.PictureBox Cor_fonte 
         BackColor       =   &H00000000&
         Height          =   195
         Left            =   10380
         ScaleHeight     =   135
         ScaleWidth      =   765
         TabIndex        =   13
         Top             =   1170
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.CommandButton cmdsimbolos 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Símbolos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   945
         Left            =   11340
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Inserir símbolos especiais nas instruções de trabalho."
         Top             =   510
         Width           =   1005
      End
      Begin VB.CommandButton Cmd_abrir_instrucao 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   945
         Left            =   9240
         Picture         =   "frmcqnc_retrabalho.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Localizar instruções de trabalho."
         Top             =   510
         Width           =   1005
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0E0&
         Height          =   1035
         Left            =   7740
         TabIndex        =   14
         Top             =   420
         Width           =   1395
         Begin VB.CheckBox Chk_negrito 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Negrito"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   90
            TabIndex        =   1
            Top             =   210
            Width           =   915
         End
         Begin VB.CheckBox Chk_italico 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Itálico"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   90
            TabIndex        =   2
            Top             =   465
            Width           =   915
         End
         Begin VB.CheckBox Chk_sublinhado 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Sublinhado"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   90
            TabIndex        =   3
            Top             =   720
            Width           =   1245
         End
      End
      Begin VB.ComboBox Cmb_fonte 
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
         Left            =   7740
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   7
         ToolTipText     =   "Fonte."
         Top             =   1770
         Width           =   3825
      End
      Begin VB.ComboBox Cmb_tamanho_fonte 
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
         ItemData        =   "frmcqnc_retrabalho.frx":010E
         Left            =   11610
         List            =   "frmcqnc_retrabalho.frx":0110
         Style           =   2  'Dropdown List
         TabIndex        =   8
         ToolTipText     =   "Tamanho."
         Top             =   1770
         Width           =   735
      End
      Begin RichTextLib.RichTextBox Txt_instrucoes 
         Height          =   2115
         Left            =   180
         TabIndex        =   0
         ToolTipText     =   "Instruções de trabalho."
         Top             =   270
         Width           =   7395
         _ExtentX        =   13044
         _ExtentY        =   3731
         _Version        =   393217
         BorderStyle     =   0
         ScrollBars      =   2
         TextRTF         =   $"frmcqnc_retrabalho.frx":0112
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CommandButton Cmd_cor 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cor"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   945
         Left            =   10290
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Mudar cor das instruções de trabalho."
         Top             =   510
         Width           =   1005
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Fonte"
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
         Index           =   31
         Left            =   9412
         TabIndex        =   16
         Top             =   1560
         Width           =   480
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Tam."
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
         Index           =   32
         Left            =   11760
         TabIndex        =   15
         Top             =   1560
         Width           =   420
      End
   End
   Begin VB.Frame Frame24 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Observações"
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
      Height          =   2505
      Left            =   55
      TabIndex        =   11
      Top             =   3540
      Width           =   12555
      Begin VB.TextBox Txt_obs 
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
         Height          =   2055
         Left            =   210
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         ToolTipText     =   "Observações."
         Top             =   270
         Width           =   12135
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   12
      Top             =   0
      Width           =   12555
      _ExtentX        =   22146
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft1     =   2
      ButtonTop1      =   2
      ButtonWidth1    =   38
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
      ButtonLeft2     =   42
      ButtonTop2      =   4
      ButtonWidth2    =   2
      ButtonHeight2   =   54
      ButtonCaption3  =   "Ajuda"
      ButtonEnabled3  =   0   'False
      ButtonIconSize3 =   32
      ButtonToolTipText3=   "Ajuda (F1)"
      ButtonKey3      =   "3"
      ButtonAlignment3=   2
      BeginProperty ButtonFont3 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft3     =   46
      ButtonTop3      =   2
      ButtonWidth3    =   36
      ButtonHeight3   =   21
      ButtonUseMaskColor3=   0   'False
      ButtonCaption4  =   "Sair"
      ButtonEnabled4  =   0   'False
      ButtonIconSize4 =   32
      ButtonToolTipText4=   "Sair (Esc)"
      ButtonKey4      =   "4"
      ButtonAlignment4=   2
      BeginProperty ButtonFont4 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft4     =   84
      ButtonTop4      =   2
      ButtonWidth4    =   26
      ButtonHeight4   =   21
      ButtonUseMaskColor4=   0   'False
      ButtonEnabled5  =   0   'False
      ButtonIconSize5 =   32
      ButtonKey5      =   "5"
      ButtonAlignment5=   2
      BeginProperty ButtonFont5 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState5    =   5
      ButtonLeft5     =   112
      ButtonTop5      =   2
      ButtonWidth5    =   24
      ButtonHeight5   =   24
      ButtonUseMaskColor5=   0   'False
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   11340
         Top             =   300
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmcqnc_retrabalho.frx":0190
         Count           =   1
      End
   End
End
Attribute VB_Name = "frmcqnc_retrabalho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ProcSalvar()
On Error GoTo tratar_erro

If Txt_instrucoes.Text = "" Then
    NomeCampo = "a instrução de trabalho"
    ProcVerificaAcao
    Txt_instrucoes.SetFocus
    Exit Sub
End If
FamiliaAntiga = Txt_instrucoes
Familiatext = Txt_obs
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_italico_Click()
On Error GoTo tratar_erro

Txt_instrucoes.SelItalic = Chk_italico

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_negrito_Click()
On Error GoTo tratar_erro

Txt_instrucoes.SelBold = Chk_negrito

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_sublinhado_Click()
On Error GoTo tratar_erro

Txt_instrucoes.SelUnderline = Chk_sublinhado

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_fonte_Click()
On Error GoTo tratar_erro

Txt_instrucoes.SelFontName = Cmb_fonte

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_tamanho_fonte_Click()
On Error GoTo tratar_erro

Txt_instrucoes.SelFontSize = Cmb_tamanho_fonte

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_abrir_instrucao_Click()
On Error GoTo tratar_erro

'If cmbMaquina = "" Then
'    usMsgbox ("Informe o posto de trabalho antes de localizar as instruções de trabalho."), vbExclamation, "CAPRIND v5.0"
'    cmbMaquina.SetFocus
'    Exit Sub
'End If
Processos_instrucoes = False
RNC_Nao_Conformidade = True
FrmInstrucoes_trabalho.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_cor_Click()
On Error GoTo tratar_erro

With CommonDialog1
    .Color = Cor_fonte.BackColor
    .ShowColor
End With
Cor_fonte.BackColor = CommonDialog1.Color
Txt_instrucoes.SelColor = CommonDialog1.Color

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdsimbolos_Click()
On Error GoTo tratar_erro

RNC_Nao_Conformidade = True
frmsimbolos.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro
    
Select Case KeyCode
    Case vbKeyF3: ProcSalvar
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 12555, 5, True

ProcCarregaComboFontes Cmb_fonte
ProcCarregaComboTamanhoFonte Cmb_tamanho_fonte, 8, 16
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from ordemservico where IDproducao = " & IIf(frmcqnc.cmbOS = "", 0, frmcqnc.cmbOS), Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    Txt_instrucoes = IIf(IsNull(TBAbrir!descfase), "", TBAbrir!descfase)
    Txt_obs = IIf(IsNull(TBAbrir!Obs), "", TBAbrir!Obs)
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_instrucoes_Change()
On Error GoTo tratar_erro

With Txt_instrucoes
    Cor_fonte.BackColor = IIf(IsNull(.SelColor), Cor_fonte.BackColor, .SelColor)
    Chk_negrito.Value = IIf(IsNull(.SelBold), 2, Abs(.SelBold))
    Chk_italico.Value = IIf(IsNull(.SelItalic), 2, Abs(.SelItalic))
    Chk_sublinhado.Value = IIf(IsNull(.SelUnderline), 2, Abs(.SelUnderline))
    Cmb_fonte = IIf(IsNull(.SelFontName), "", .SelFontName)
1:
    Cmb_tamanho_fonte = IIf(IsNull(.SelFontSize), "", .SelFontSize)
End With

Exit Sub
tratar_erro:
    If Err.Number = 383 Then
        Cmb_tamanho_fonte.AddItem IIf(IsNull(Txt_instrucoes.SelFontSize), "", Txt_instrucoes.SelFontSize)
        GoTo 1
    End If
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcSalvar
    'Case 3: ProcAjuda
    Case 4: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
