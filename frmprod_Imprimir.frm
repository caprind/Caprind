VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmprod_Imprimir 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "PCP | Gerenciamento de ordem - Menu impressão"
   ClientHeight    =   5745
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   6405
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   6405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   6405
      _ExtentX        =   11298
      _ExtentY        =   873
      DibPicture      =   "frmprod_Imprimir.frx":0000
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
      Icon            =   "frmprod_Imprimir.frx":2126
      ShowMaximize    =   0   'False
      ShowMinimize    =   0   'False
   End
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   23
      Top             =   5340
      Width           =   6405
      _ExtentX        =   11298
      _ExtentY        =   714
   End
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   5010
      Top             =   210
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frmprod_Imprimir.frx":2440
      Count           =   1
   End
   Begin DrawSuite2022.USButton Cmd_avancar 
      Height          =   3645
      Left            =   4590
      TabIndex        =   17
      Top             =   1620
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   6429
      BorderColor     =   5263559
      BorderColorDisabled=   13160660
      BorderColorDown =   4013465
      BorderColorOver =   4408288
      Caption         =   "Avançar >>>>"
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
      PicSizeH        =   48
      PicSizeW        =   48
      Theme           =   4
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   90
      TabIndex        =   20
      Top             =   1560
      Width           =   4425
      Begin VB.CheckBox Chk_etiqueta_selecionadas 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Imprimir etiqueta individual (selecionadas)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   180
         TabIndex        =   11
         Top             =   2886
         Width           =   3975
      End
      Begin VB.CheckBox Chk_etiqueta 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Imprimir etiqueta individual"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   180
         TabIndex        =   10
         Top             =   2640
         Width           =   3345
      End
      Begin VB.CheckBox chkPersonalizado_selecionadas 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Imprimir rel. personalizado (selecionadas)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   180
         TabIndex        =   12
         Top             =   3132
         Width           =   3885
      End
      Begin VB.CheckBox chkRm_selecionadas 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Imprimir RM (selecionadas)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   180
         TabIndex        =   7
         Top             =   1902
         Width           =   3885
      End
      Begin VB.CheckBox chkOrdem_selecionadas 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Imprimir ordem(ns) (selecionadas)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   180
         TabIndex        =   1
         Top             =   426
         Width           =   3885
      End
      Begin VB.CheckBox Chk_ordem_rm_resumido 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Imprimir ordem(ns) e RM (resumido)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   180
         TabIndex        =   3
         Top             =   918
         Width           =   3885
      End
      Begin VB.CheckBox chkOrdem_rm_selecionadas 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Imprimir ordem(ns) e RM (selecionadas)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   180
         TabIndex        =   4
         Top             =   1164
         Width           =   3885
      End
      Begin VB.CheckBox Chk_ordem_manual 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Imprimir ordem(ns) p/ apontamento manual"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   180
         TabIndex        =   5
         Top             =   1410
         Width           =   4155
      End
      Begin VB.CheckBox Chk_frequencia 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Imprimir frequencia(s) de medição"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   180
         TabIndex        =   9
         Top             =   2394
         Width           =   3285
      End
      Begin VB.CheckBox Chk_plano 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Imprimir plano(s) de inspeção"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   180
         TabIndex        =   8
         Top             =   2148
         Width           =   2985
      End
      Begin VB.CheckBox Chk_visualizar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Visualizando impressão"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   180
         TabIndex        =   13
         Top             =   3390
         Width           =   2295
      End
      Begin VB.CheckBox Chk_ordem 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Imprimir ordem(ns)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   180
         TabIndex        =   0
         Top             =   180
         Width           =   1995
      End
      Begin VB.CheckBox Chk_ordem_rm 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Imprimir ordem(ns) e RM"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   180
         TabIndex        =   2
         Top             =   672
         Width           =   2475
      End
      Begin VB.CheckBox Chk_rm 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Imprimir RM"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   180
         TabIndex        =   6
         Top             =   1656
         Width           =   2235
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   705
      Left            =   90
      TabIndex        =   18
      Top             =   5310
      Visible         =   0   'False
      Width           =   6195
      Begin VB.ComboBox Cmb_Ate 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3690
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   16
         ToolTipText     =   "Número da ordem."
         Top             =   240
         Width           =   2325
      End
      Begin VB.ComboBox Cmb_De 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   810
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   15
         ToolTipText     =   "Número da ordem."
         Top             =   240
         Width           =   2325
      End
      Begin VB.OptionButton Opt_De 
         BackColor       =   &H00E0E0E0&
         Caption         =   "De :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   180
         TabIndex        =   14
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Até :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   3255
         TabIndex        =   19
         Top             =   270
         Width           =   360
      End
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   90
      TabIndex        =   21
      Top             =   6030
      Visible         =   0   'False
      Width           =   6195
      _ExtentX        =   10927
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
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   0
      TabIndex        =   22
      Top             =   510
      Width           =   6405
      _ExtentX        =   11298
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
      ButtonCaption1  =   "Relatório"
      ButtonEnabled1  =   0   'False
      ButtonIconSize1 =   32
      ButtonToolTipText1=   "Relatório (F5)"
      ButtonKey1      =   "1"
      ButtonAlignment1=   2
      BeginProperty ButtonFont1 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft1     =   2
      ButtonTop1      =   2
      ButtonWidth1    =   60
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
      ButtonLeft2     =   64
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft3     =   68
      ButtonTop3      =   2
      ButtonWidth3    =   41
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft4     =   111
      ButtonTop4      =   2
      ButtonWidth4    =   30
      ButtonHeight4   =   21
      ButtonUseMaskColor4=   0   'False
      ButtonEnabled5  =   0   'False
      ButtonIconSize5 =   32
      ButtonKey5      =   "5"
      ButtonAlignment5=   2
      BeginProperty ButtonFont5 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState5    =   5
      ButtonLeft5     =   143
      ButtonTop5      =   2
      ButtonWidth5    =   24
      ButtonHeight5   =   24
   End
End
Attribute VB_Name = "frmprod_Imprimir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FormulaRel_Ordem1 As String 'OK

Private Sub Chk_etiqueta_Click()
On Error GoTo tratar_erro

ProcHabDesabVisualizandoImpressao
If Chk_etiqueta.Value = 1 Then ProcDesabChkSelecionadas

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_etiqueta_selecionadas_Click()
On Error GoTo tratar_erro

ProcHabDesabVisualizandoImpressao
If Chk_etiqueta_selecionadas.Value = 1 Then
    ProcCorrigeFormChkSelecionada
Else
    Cmd_avancar.Enabled = True
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_frequencia_Click()
On Error GoTo tratar_erro

ProcHabDesabVisualizandoImpressao
If Chk_frequencia.Value = 1 Then ProcDesabChkSelecionadas

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_ordem_Click()
On Error GoTo tratar_erro

ProcHabDesabVisualizandoImpressao
If Chk_ordem.Value = 1 Then ProcDesabChkSelecionadas

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_ordem_manual_Click()
On Error GoTo tratar_erro

ProcHabDesabVisualizandoImpressao
If Chk_ordem_manual.Value = 1 Then ProcDesabChkSelecionadas

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_ordem_rm_Click()
On Error GoTo tratar_erro

ProcHabDesabVisualizandoImpressao
If Chk_ordem_rm.Value = 1 Then ProcDesabChkSelecionadas

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_ordem_rm_resumido_Click()
On Error GoTo tratar_erro

ProcHabDesabVisualizandoImpressao
If Chk_ordem_rm_resumido.Value = 1 Then ProcDesabChkSelecionadas
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_plano_Click()
On Error GoTo tratar_erro

ProcHabDesabVisualizandoImpressao
If Chk_plano.Value = 1 Then ProcDesabChkSelecionadas

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_rm_Click()
On Error GoTo tratar_erro

ProcHabDesabVisualizandoImpressao
If Chk_rm.Value = 1 Then ProcDesabChkSelecionadas

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcDesabChkSelecionadas()
On Error GoTo tratar_erro

chkOrdem_selecionadas.Value = 0
chkOrdem_rm_selecionadas.Value = 0
chkRm_selecionadas.Value = 0
Chk_etiqueta_selecionadas.Value = 0
chkPersonalizado_selecionadas.Value = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcHabDesabVisualizandoImpressao()
On Error GoTo tratar_erro

With Chk_visualizar
    If Chk_ordem.Value = 0 And chkOrdem_selecionadas.Value = 0 And Chk_ordem_rm.Value = 0 And Chk_ordem_rm_resumido.Value = 0 And chkOrdem_rm_selecionadas.Value = 0 And Chk_ordem_manual.Value = 0 And Chk_rm.Value = 0 And chkRm_selecionadas.Value = 0 And Chk_plano.Value = 0 And Chk_frequencia.Value = 0 And Chk_etiqueta.Value = 0 And Chk_etiqueta_selecionadas.Value = 0 And chkPersonalizado_selecionadas.Value = 0 Then
        .Value = 0
        .Enabled = False
    Else
        .Enabled = True
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkOrdem_rm_selecionadas_Click()
On Error GoTo tratar_erro

ProcHabDesabVisualizandoImpressao
If chkOrdem_rm_selecionadas.Value = 1 Then
    chkOrdem_selecionadas.Value = 0
    chkRm_selecionadas.Value = 0
    chkPersonalizado_selecionadas.Value = 0
    ProcCorrigeFormChkSelecionada
Else
    Cmd_avancar.Enabled = True
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkOrdem_selecionadas_Click()
On Error GoTo tratar_erro

ProcHabDesabVisualizandoImpressao
If chkOrdem_selecionadas.Value = 1 Then
    chkOrdem_rm_selecionadas.Value = 0
    chkRm_selecionadas.Value = 0
    chkPersonalizado_selecionadas.Value = 0
    ProcCorrigeFormChkSelecionada
Else
    Cmd_avancar.Enabled = True
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkPersonalizado_selecionadas_Click()
On Error GoTo tratar_erro

ProcHabDesabVisualizandoImpressao
If chkPersonalizado_selecionadas.Value = 1 Then
    chkOrdem_selecionadas.Value = 0
    chkOrdem_rm_selecionadas.Value = 0
    chkRm_selecionadas.Value = 0
    ProcCorrigeFormChkSelecionada
Else
    Cmd_avancar.Enabled = True
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkRm_selecionadas_Click()
On Error GoTo tratar_erro

ProcHabDesabVisualizandoImpressao
If chkRm_selecionadas.Value = 1 Then
    chkOrdem_selecionadas.Value = 0
    chkOrdem_rm_selecionadas.Value = 0
    chkPersonalizado_selecionadas.Value = 0
    ProcCorrigeFormChkSelecionada
Else
    Cmd_avancar.Enabled = True
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_De_Click()
On Error GoTo tratar_erro

Cmb_Ate.Clear
Cmb_Ate.Enabled = True
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select Ordem from Producao where Concluida = 'False' and Ordem > " & Cmb_de, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    Do While TBAbrir.EOF = False
        Cmb_Ate.AddItem TBAbrir!Ordem
        TBAbrir.MoveNext
    Loop
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcRelatorio()
On Error GoTo tratar_erro

If Chk_visualizar.Value = 1 Then Acao = "visualizar impressão" Else Acao = "imprimir"
If Chk_ordem.Value = 0 And chkOrdem_selecionadas.Value = 0 And Chk_ordem_rm.Value = 0 And Chk_ordem_rm_resumido.Value = 0 And chkOrdem_rm_selecionadas.Value = 0 And Chk_ordem_manual.Value = 0 And Chk_rm.Value = 0 And chkRm_selecionadas.Value = 0 And Chk_plano.Value = 0 And Chk_frequencia.Value = 0 And Chk_etiqueta.Value = 0 And Chk_etiqueta_selecionadas.Value = 0 And chkPersonalizado_selecionadas.Value = 0 Then
    NomeCampo = "uma das opções"
    ProcVerificaAcao
    Exit Sub
End If
If Avancar = True Then
    NomeCampo = "o número da ordem"
    If Opt_De.Value = True And Cmb_de = "" Then
        ProcVerificaAcao
        Cmb_de.SetFocus
        Exit Sub
    End If
    If Opt_De.Value = True And Cmb_Ate = "" Then
        ProcVerificaAcao
        Cmb_Ate.SetFocus
        Exit Sub
    End If
    
    If Opt_De.Value = True Then
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from Producao where Concluida = 'False' and Ordem >= " & Cmb_de & " and Ordem <= " & Cmb_Ate, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            PBLista.Min = 0
            PBLista.Max = TBAbrir.RecordCount
            PBLista.Value = 1
            Contador1 = 0
            Do While TBAbrir.EOF = False
            'TBAbrir.MovePrevious
            Ordem = TBAbrir!Ordem
                If Chk_ordem.Value = 1 Then procImprimirOrdem
                If Chk_ordem_rm.Value = 1 Then ProcImprimirOrdemRM
                If Chk_ordem_rm_resumido.Value = 1 Then ProcImprimirOrdemRMRes
                If Chk_ordem_manual.Value = 1 Then ProcImprimirOrdemRMManual
                If Chk_rm.Value = 1 Then ProcImprimirRM
                If Chk_plano.Value = 1 Then ProcImprimirPlano
                If Chk_frequencia.Value = 1 Then ProcImprimirFrequencia
                If Chk_etiqueta.Value = 1 Then ProcImprimirEtiqueta
                TBAbrir.MoveNext
                Contador1 = Contador1 + 1
                PBLista.Value = Contador1
            Loop
        End If
        TBAbrir.Close
    End If
Else
    With frmprod
        If chkOrdem_selecionadas.Value = 1 Or chkOrdem_rm_selecionadas.Value = 1 Or chkRm_selecionadas.Value = 1 Or Chk_etiqueta_selecionadas.Value = 1 Or chkPersonalizado_selecionadas.Value = 1 Then
            FormulaRel_Ordem1 = ""
            Permitido = False
            With .Lista
                For InitFor = 1 To .ListItems.Count
                    If .ListItems.Item(InitFor).Checked = True Then
                        ProcSalvarViaOrdem .ListItems.Item(InitFor), True
                        If FormulaRel_Ordem1 = "" Then
                            FormulaRel_Ordem1 = "{Producao.Ordem} = " & .ListItems.Item(InitFor)
                        Else
                            FormulaRel_Ordem1 = FormulaRel_Ordem1 & " or " & "{Producao.Ordem} = " & .ListItems.Item(InitFor)
                        End If
                        Permitido = True
                    End If
                Next InitFor
            End With
            If Permitido = True Then
                ProcImprimirSelecionadas
                If Chk_etiqueta_selecionadas.Value = 1 Then
                    With .Lista
                        For InitFor = 1 To .ListItems.Count
                            If .ListItems.Item(InitFor).Checked = True Then
                                Set TBAbrir = CreateObject("adodb.recordset")
                                TBAbrir.Open "Select Ordem, Quant from Producao where Ordem = " & .ListItems.Item(InitFor), Conexao, adOpenKeyset, adLockOptimistic
                                If TBAbrir.EOF = False Then
                                    frmprod.MRP_Qtcopias_Etiq = 0
                                    frmprod.ProcCriarEtiquetas TBAbrir!Ordem, TBAbrir!Quant
                                    OF = TBAbrir!Ordem
                                    If USMsgBox("Deseja imprimir as etiquetas individuais?", vbYesNo, "CAPRIND v5.0") = vbYes Then
                                        frmprod_imprimir_etiqueta.Show 1
                                    End If
                                End If
                                TBAbrir.Close
                            End If
                        Next InitFor
                    End With
                End If
            Else
                If Chk_visualizar.Value = 1 Then Texto = "visualizar impressão" Else Texto = "imprimir"
                USMsgBox ("Informe a(s) ordem(ns) antes de " & Texto & "."), vbExclamation, "CAPRIND v5.0"
            End If
        Else
            If .txtof = "" Or .txtof = "0" Then
                NomeCampo = "a ordem"
                ProcVerificaAcao
                Unload Me
                frmProd_imp_filtro.Show 1
                Exit Sub
            End If
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select * from Producao where Ordem = " & .txtof, Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                If Chk_ordem.Value = 1 Then procImprimirOrdem
                If Chk_ordem_rm.Value = 1 Then ProcImprimirOrdemRM
                If Chk_ordem_rm_resumido.Value = 1 Then ProcImprimirOrdemRMRes
                If Chk_ordem_manual.Value = 1 Then ProcImprimirOrdemRMManual
                If Chk_rm.Value = 1 Then ProcImprimirRM
                If Chk_plano.Value = 1 Then ProcImprimirPlano
                If Chk_frequencia.Value = 1 Then ProcImprimirFrequencia
                If Chk_etiqueta.Value = 1 Then ProcImprimirEtiqueta
            End If
            TBAbrir.Close
        End If
    End With
End If
FormulaRel_Ordem1 = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procImprimirOrdem()
On Error GoTo tratar_erro

NomeRel = "Pcp_ordem.rpt"
If Avancar = False Then FormulaRel_Ordem1 = "{Producao.Ordem} = " & frmprod.txtof Else FormulaRel_Ordem1 = "{Producao.Ordem} = " & TBAbrir!Ordem
ProcImprimir

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimirOrdemRM()
On Error GoTo tratar_erro

NomeRel = "Pcp_ordem e rm.rpt"
If Avancar = False Then FormulaRel_Ordem1 = "{Producao.Ordem} = " & frmprod.txtof Else FormulaRel_Ordem1 = "{Producao.Ordem} = " & Ordem 'TBAbrir!Ordem
ProcImprimir

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimirOrdemRMRes()
On Error GoTo tratar_erro

NomeRel = "Pcp_ordem e rm_resumido.rpt"
If Avancar = False Then FormulaRel_Ordem1 = "{Producao.Ordem} = " & frmprod.txtof Else FormulaRel_Ordem1 = "{Producao.Ordem} = " & TBAbrir!Ordem
ProcImprimir

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimirSelecionadas()
On Error GoTo tratar_erro

If chkOrdem_selecionadas.Value = 1 Then
    NomeRel = "Pcp_ordem_selecionadas.rpt"
    ProcImprimir
End If
If chkOrdem_rm_selecionadas.Value = 1 Then
    NomeRel = "Pcp_ordem e rm_selecionadas.rpt"
    ProcImprimir
End If
If chkRm_selecionadas.Value = 1 Then
    NomeRel = "Pcp_rm_selecionadas.rpt"
    ProcImprimir
End If
If chkPersonalizado_selecionadas.Value = 1 Then
    NomeRel = "Pcp_ordem_personalizado_selecionadas.rpt"
    ProcImprimir
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimirOrdemRMManual()
On Error GoTo tratar_erro

NomeRel = "Pcp_ordem_apontamento manual.rpt"
If Avancar = False Then FormulaRel_Ordem1 = "{Producao.Ordem} = " & frmprod.txtof Else FormulaRel_Ordem1 = "{Producao.Ordem} = " & TBAbrir!Ordem
ProcImprimir

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimirRM()
On Error GoTo tratar_erro

NomeRel = "Pcp_rm.rpt"
If Avancar = False Then FormulaRel_Ordem1 = "{Producao.Ordem} = " & frmprod.txtof Else FormulaRel_Ordem1 = "{Producao.Ordem} = " & TBAbrir!Ordem
ProcImprimir

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimirPlano()
On Error GoTo tratar_erro

NomeRel = "Pcp_plano inspecao.rpt"
If Avancar = False Then FormulaRel_Ordem1 = "{ordemservico.Ordem} = " & frmprod.txtof & " and {Planodimensao.PCP} = True" Else FormulaRel_Ordem1 = "{ordemservico.Ordem} = " & TBAbrir!Ordem & " and {Planodimensao.PCP} = True"
ProcImprimir

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimirFrequencia()
On Error GoTo tratar_erro

NomeRel = "Pcp_plano inspecao_frequencia de medicao.rpt"
If Avancar = False Then FormulaRel_Ordem1 = "{Producao.Ordem} = " & frmprod.txtof & " and {Planodimensao.Freq} <> '" & Null & "' and {Planodimensao.PCP} = True" Else FormulaRel_Ordem1 = "{Producao.Ordem} = " & TBAbrir!Ordem & " and {Planodimensao.Freq} <> '" & Null & "' and {Planodimensao.PCP} = True"
ProcImprimir

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimirEtiqueta()
On Error GoTo tratar_erro

With frmprod
    .MRP_Qtcopias_Etiq = 0
    If Avancar = False Then
        .ProcCriarEtiquetas .txtof, frmprod.txtQuantidade
        OF = frmprod.txtof
    Else
        .ProcCriarEtiquetas TBAbrir!Ordem, TBAbrir!Quant
        OF = TBAbrir!Ordem
    End If
End With

frmprod_imprimir_etiqueta.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro

If Chk_visualizar.Value = 1 Then
    ProcImprimirRel FormulaRel_Ordem1, ""
Else
    ProcImprimirDireto FormulaRel_Ordem1, ""
    If Avancar = False Then ProcSalvarViaOrdem frmprod.txtof, True Else ProcSalvarViaOrdem TBAbrir!Ordem, True
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_avancar_Click()
On Error GoTo tratar_erro

If Avancar = False Then
    Avancar = True
    Cmd_avancar.Caption = "<<<< Recuar"
    Height = 6720
    Frame1.Visible = True
    PBLista.Visible = True
    Opt_De.Value = True
Else
    Avancar = False
    Cmd_avancar.Caption = "Avançar >>>>"
    Height = 5750
    Frame1.Visible = False
    PBLista.Visible = False
    Opt_De.Value = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyF5: ProcRelatorio
    Case vbKeyEscape: ProcSair
End Select
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 6405, 5, True
Avancar = False
Height = 5750

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt_De_Click()
On Error GoTo tratar_erro

If Opt_De.Value = True Then
    Cmb_de.Enabled = True
    Cmb_de.SetFocus
    ProcCarregaOrdem
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaOrdem()
On Error GoTo tratar_erro

Cmb_de.Clear
Cmb_Ate.Clear
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select Ordem from Producao where Concluida = 'False'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    Do While TBAbrir.EOF = False
        Cmb_de.AddItem TBAbrir!Ordem
        TBAbrir.MoveNext
    Loop
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcRelatorio
    'Case 3: ProcAjuda
    Case 4: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCorrigeFormChkSelecionada()
On Error GoTo tratar_erro

Avancar = False
Cmd_avancar.Caption = "Avançar >>>>"
Cmd_avancar.Enabled = False
Height = 5220
Frame1.Visible = False
PBLista.Visible = False
Opt_De.Value = False
Chk_ordem.Value = 0
Chk_ordem_rm.Value = 0
Chk_ordem_rm_resumido.Value = 0
Chk_ordem_manual.Value = 0
Chk_rm.Value = 0
Chk_plano.Value = 0
Chk_frequencia.Value = 0
Chk_etiqueta.Value = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
