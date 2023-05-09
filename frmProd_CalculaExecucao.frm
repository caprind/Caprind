VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmProd_CalculaExecucao 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'Nenhum
   Caption         =   "Engenharia | Calcular execução"
   ClientHeight    =   4125
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5505
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4125
   ScaleWidth      =   5505
   StartUpPosition =   2  'Centralziar na Tela
   Begin VB.Frame Frame29 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2955
      Left            =   210
      TabIndex        =   8
      Top             =   540
      Width           =   5055
      Begin DrawSuite2022.USTextBoxEx txtPcHora 
         Height          =   375
         Left            =   3450
         TabIndex        =   1
         Top             =   1170
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   661
         Alignment       =   2
         AutoFormatDate  =   -1  'True
         BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CurrencyChar    =   ""
         Decimals        =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontDisabled {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontOver {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FormatType      =   2
         MaskType        =   1
         MaxLength       =   29
         NumberOnly      =   -1  'True
         Text            =   "0,0000"
      End
      Begin VB.TextBox txtUN 
         Alignment       =   2  'Centralizar
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   3450
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   13
         ToolTipText     =   "Total de peças por tempo de execução prevista."
         Top             =   270
         Width           =   1155
      End
      Begin DrawSuite2022.USButton btnCalcular 
         Height          =   675
         Left            =   360
         TabIndex        =   2
         Top             =   2070
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1191
         DibPicture      =   "frmProd_CalculaExecucao.frx":0000
         BorderColor     =   4960354
         BorderColorDisabled=   13160660
         BorderColorDown =   4210752
         BorderColorOver =   49152
         Caption         =   "Calcular"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
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
         PicAlign        =   7
         ShowFocusRect   =   0   'False
         ShowFocusRectDown=   0   'False
         Theme           =   3
      End
      Begin VB.TextBox txtPcHora2 
         Alignment       =   2  'Centralizar
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   4170
         MaxLength       =   15
         TabIndex        =   4
         ToolTipText     =   "Total de peças por tempo de execução prevista."
         Top             =   2970
         Width           =   1155
      End
      Begin VB.TextBox TxtA3 
         Alignment       =   2  'Centralizar
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3450
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   "00:00:00"
         Top             =   1620
         Width           =   1155
      End
      Begin MSMask.MaskEdBox txtexecucao 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "H:mm:ss"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   4
         EndProperty
         Height          =   330
         Left            =   3450
         TabIndex        =   0
         ToolTipText     =   "Tempo de execução previsto."
         Top             =   750
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   582
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         AutoTab         =   -1  'True
         MaxLength       =   9
         MouseIcon       =   "frmProd_CalculaExecucao.frx":0561
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "###:##:##"
         PromptChar      =   "_"
      End
      Begin DrawSuite2022.USButton BtnCarregar 
         Height          =   675
         Left            =   2550
         TabIndex        =   3
         Top             =   2070
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1191
         DibPicture      =   "frmProd_CalculaExecucao.frx":087B
         BorderColor     =   5263559
         BorderColorDisabled=   13160660
         BorderColorDown =   4013465
         BorderColorOver =   4408288
         Caption         =   "Carregar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
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
         ShowFocusRect   =   0   'False
         ShowFocusRectDown=   0   'False
         Theme           =   4
      End
      Begin VB.Label lblUN 
         Alignment       =   1  'Alinhar à Direita
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparente
         Caption         =   "Unidade :"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   180
         TabIndex        =   12
         Top             =   330
         Width           =   3135
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Alinhar à Direita
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparente
         Caption         =   "Tempo por unidade :"
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   66
         Left            =   180
         TabIndex        =   11
         Top             =   765
         Width           =   3135
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Alinhar à Direita
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparente
         Caption         =   "Unidade(s) por tempo :"
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   67
         Left            =   1305
         TabIndex        =   10
         Top             =   1215
         Width           =   2010
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Alinhar à Direita
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparente
         Caption         =   "Tempo de execução por unidade :"
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   68
         Left            =   375
         TabIndex        =   9
         Top             =   1650
         Width           =   2940
      End
   End
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   7
      Top             =   3720
      Width           =   5505
      _ExtentX        =   9710
      _ExtentY        =   714
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   405
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   5505
      _ExtentX        =   9710
      _ExtentY        =   714
      DibPicture      =   "frmProd_CalculaExecucao.frx":A328
      CaptionDelimiter=   "|"
      CaptionOnCenter =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Icon            =   "frmProd_CalculaExecucao.frx":A889
      ShowMaximize    =   0   'False
      ShowMinimize    =   0   'False
   End
End
Attribute VB_Name = "frmProd_CalculaExecucao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCalcular_Click()
On Error GoTo tratar_erro

ProcCalculaExecucaoPeca

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCalculaExecucaoPeca()
On Error GoTo tratar_erro

txtexecucao.PromptInclude = False
If Len(txtexecucao.Text) < 7 Then
    txtexecucao.PromptInclude = True
    Exit Sub
End If
txtexecucao.PromptInclude = True
If txtexecucao > "023:59:59" Then
    ProcFormataHora (txtexecucao)
    Familiatext = DataResultado
    TxtA3 = FunCalculaSegPC(Familiatext, txtPcHora)
Else
    TxtA3 = FunCalculaSegPC(txtexecucao, txtPcHora)
End If
TxtA3 = FormataTempo(TxtA3.Text)

txtexecucao.PromptInclude = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub btnCarregar_Click()
On Error GoTo tratar_erro

If USMsgBox("Deseja carregar esse tempo de execução calculado?", vbYesNo, "CAPRIND v5.0") = vbYes Then
frmprod.txtexecucao.Text = "0" & TxtA3.Text
frmprod.txtPcHora.Text = 1
Unload Me
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

txtUN.Text = frmprod.txtUN.Text

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
