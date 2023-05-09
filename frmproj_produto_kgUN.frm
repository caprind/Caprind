VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmproj_produto_kgUN 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cálculo de Kg/un"
   ClientHeight    =   1845
   ClientLeft      =   1890
   ClientTop       =   1365
   ClientWidth     =   8655
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
   Icon            =   "frmproj_produto_kgUN.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   8655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Custo unitário"
      Height          =   855
      Left            =   55
      TabIndex        =   11
      Top             =   3660
      Width           =   11745
      Begin VB.TextBox txtmaodeobra 
         Alignment       =   2  'Center
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
         Height          =   315
         Left            =   180
         Locked          =   -1  'True
         MaxLength       =   50
         MousePointer    =   99  'Custom
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Total processo."
         Top             =   420
         Width           =   1605
      End
      Begin VB.TextBox txtmaterial 
         Alignment       =   2  'Center
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
         Height          =   315
         Left            =   2136
         Locked          =   -1  'True
         MaxLength       =   50
         MousePointer    =   99  'Custom
         TabIndex        =   6
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Total material."
         Top             =   420
         Width           =   1605
      End
      Begin VB.TextBox txtimpostos 
         Alignment       =   2  'Center
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
         Height          =   315
         Left            =   8004
         Locked          =   -1  'True
         MaxLength       =   50
         MousePointer    =   99  'Custom
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Total impostos."
         Top             =   420
         Width           =   1605
      End
      Begin VB.TextBox txtvalortotal 
         Alignment       =   2  'Center
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
         Height          =   315
         Left            =   9960
         Locked          =   -1  'True
         MaxLength       =   50
         MousePointer    =   99  'Custom
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   "0,0000"
         ToolTipText     =   "Valor total."
         Top             =   420
         Width           =   1605
      End
      Begin VB.TextBox txtlucro 
         Alignment       =   2  'Center
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
         Height          =   315
         Left            =   4092
         Locked          =   -1  'True
         MaxLength       =   50
         MousePointer    =   99  'Custom
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Total lucro."
         Top             =   420
         Width           =   1605
      End
      Begin VB.TextBox txtcomissao 
         Alignment       =   2  'Center
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
         Height          =   315
         Left            =   6048
         Locked          =   -1  'True
         MaxLength       =   50
         MousePointer    =   99  'Custom
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Total comissão."
         Top             =   420
         Width           =   1605
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total processo"
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
         Index           =   2
         Left            =   360
         TabIndex        =   23
         Top             =   210
         Width           =   1245
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+"
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
         Index           =   3
         Left            =   1893
         TabIndex        =   22
         Top             =   450
         Width           =   135
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total material"
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
         Index           =   4
         Left            =   2338
         TabIndex        =   21
         Top             =   210
         Width           =   1200
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+"
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
         Index           =   5
         Left            =   6980
         TabIndex        =   20
         Top             =   450
         Width           =   135
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total impostos"
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
         Index           =   6
         Left            =   8176
         TabIndex        =   19
         Top             =   210
         Width           =   1260
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "="
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
         Index           =   7
         Left            =   9717
         TabIndex        =   18
         Top             =   450
         Width           =   135
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor unitário"
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
         Index           =   8
         Left            =   10192
         TabIndex        =   17
         Top             =   210
         Width           =   1140
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+"
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
         Index           =   9
         Left            =   3849
         TabIndex        =   16
         Top             =   450
         Width           =   135
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total lucro"
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
         Index           =   10
         Left            =   4444
         TabIndex        =   15
         Top             =   210
         Width           =   900
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total comissão"
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
         Index           =   11
         Left            =   6213
         TabIndex        =   14
         Top             =   210
         Width           =   1275
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+"
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
         Index           =   12
         Left            =   5805
         TabIndex        =   13
         Top             =   450
         Width           =   135
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+"
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
         Index           =   13
         Left            =   7761
         TabIndex        =   12
         Top             =   450
         Width           =   135
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   825
      Left            =   55
      TabIndex        =   24
      Top             =   990
      Width           =   8565
      Begin VB.TextBox txtLargura 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1890
         MaxLength       =   50
         TabIndex        =   1
         Text            =   "0,00000"
         ToolTipText     =   "Largura."
         Top             =   390
         Width           =   1455
      End
      Begin VB.TextBox txtComprimento 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   5310
         MaxLength       =   50
         TabIndex        =   3
         Text            =   "0,00000"
         ToolTipText     =   "Comprimento."
         Top             =   390
         Width           =   1455
      End
      Begin VB.TextBox txtdensidade 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   180
         MaxLength       =   10
         TabIndex        =   0
         Text            =   "0,00000"
         ToolTipText     =   "Densidade."
         Top             =   390
         Width           =   1455
      End
      Begin VB.TextBox txtEspessura 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3600
         MaxLength       =   50
         TabIndex        =   2
         Text            =   "0,00000"
         ToolTipText     =   "Espessura."
         Top             =   390
         Width           =   1455
      End
      Begin VB.TextBox txtKgUN 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   7020
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   4
         TabStop         =   0   'False
         Text            =   "0,00000"
         ToolTipText     =   "Kg/unidade."
         Top             =   390
         Width           =   1365
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "="
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Left            =   6840
         TabIndex        =   33
         Top             =   465
         Width           =   105
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Left            =   1710
         TabIndex        =   32
         Top             =   465
         Width           =   75
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Left            =   3420
         TabIndex        =   31
         Top             =   465
         Width           =   75
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Comprimento"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5565
         TabIndex        =   30
         Top             =   180
         Width           =   945
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Largura"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2340
         TabIndex        =   29
         Top             =   180
         Width           =   555
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Densidade"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   532
         TabIndex        =   28
         Top             =   180
         Width           =   750
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Left            =   5130
         TabIndex        =   27
         Top             =   465
         Width           =   75
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kg/unidade"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   7297
         TabIndex        =   26
         Top             =   180
         Width           =   810
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Espessura"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3960
         TabIndex        =   25
         Top             =   180
         Width           =   735
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   34
      Top             =   0
      Width           =   8565
      _ExtentX        =   15108
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
      ButtonKey3      =   "6"
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
      ButtonKey4      =   "7"
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
      ButtonKey5      =   "8"
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
         Left            =   7110
         Top             =   150
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmproj_produto_kgUN.frx":030A
         Count           =   1
      End
   End
End
Attribute VB_Name = "frmproj_produto_kgUN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ProcSair()
On Error GoTo tratar_erro

Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvar()
On Error GoTo tratar_erro

ValorTotal = IIf(txtKgUN = "", 0, txtKgUN)
If ValorTotal < 0 Then
    USMsgBox ("O Kg/un não pode ser menor que zero."), vbExclamation, "CAPRIND v5.0"
    txtKgUN.SetFocus
    Exit Sub
End If
frmproj_produto.txtPeso = txtKgUN
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro
  
Select Case KeyCode
    Case vbKeyF3: ProcSalvar
    Case vbKeyEscape: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCalculaTotal()
On Error GoTo tratar_erro

qt = IIf(txtdensidade = "", 0, txtdensidade)
Qtde = IIf(txtLargura = "", 0, txtLargura)
Qtd = IIf(txtespessura = "", 0, txtespessura)
ValorTotal = IIf(txtComprimento = "", 0, txtComprimento)
txtKgUN = Format((qt * (Qtde / 1000) * (Qtd / 1000) * ValorTotal) / 1000, "###,##0.0000000000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 8565, 5, True

If frmproj_produto.cmbun.Text = "M³" Then
    Label4.Caption = frmproj_produto.cmbun.Text & " x Unidade"
    txtKgUN.ToolTipText = frmproj_produto.cmbun.Text & " x Unidade"
    frmproj_produto_kgUN.Caption = "Cálculo " & frmproj_produto.cmbun.Text & " x Unidade"
    With txtdensidade
        .Text = 1
        .Locked = True
        .TabStop = False
    End With
    txtespessura.Text = frmproj_produto.txtespessura
    txtLargura.Text = frmproj_produto.txtLargura
End If
If frmproj_produto.cmbunkg.Text = "Mt/L" Then
    With txtComprimento
        .Text = 1000
        .Locked = True
        .TabStop = False
    End With
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtComprimento_Change()
On Error GoTo tratar_erro

If txtComprimento.Text <> "" Then
    VerifNumero = txtComprimento.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtComprimento.Text = ""
        txtComprimento.SetFocus
        Exit Sub
    End If
End If
ProcCalculaTotal

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtComprimento_GotFocus()
On Error GoTo tratar_erro

If txtComprimento = "0,00000" Then txtComprimento = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtComprimento_LostFocus()
On Error GoTo tratar_erro

txtComprimento.Text = IIf(txtComprimento = "", "0,00000", Format(txtComprimento.Text, "###,##0.0000000000"))

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtdensidade_Change()
On Error GoTo tratar_erro

If txtdensidade.Text <> "" Then
    VerifNumero = txtdensidade.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtdensidade.Text = ""
        txtdensidade.SetFocus
        Exit Sub
    End If
End If
ProcCalculaTotal

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtdensidade_GotFocus()
On Error GoTo tratar_erro

If txtdensidade = "0,00000" Then txtdensidade = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtdensidade_LostFocus()
On Error GoTo tratar_erro

txtdensidade.Text = IIf(txtdensidade = "", "0,00000", Format(txtdensidade.Text, "###,##0.0000000000"))

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtEspessura_Change()
On Error GoTo tratar_erro

If txtespessura.Text <> "" Then
    VerifNumero = txtespessura.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtespessura.Text = ""
        txtespessura.SetFocus
        Exit Sub
    End If
End If
ProcCalculaTotal

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtEspessura_GotFocus()
On Error GoTo tratar_erro

If txtespessura = "0,00000" Then txtespessura = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtEspessura_LostFocus()
On Error GoTo tratar_erro

txtespessura.Text = IIf(txtespessura = "", "0,00000", Format(txtespessura.Text, "###,##0.0000000000"))

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtLargura_Change()
On Error GoTo tratar_erro

If txtLargura.Text <> "" Then
    VerifNumero = txtLargura.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtLargura.Text = ""
        txtLargura.SetFocus
        Exit Sub
    End If
End If
ProcCalculaTotal

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtLargura_GotFocus()
On Error GoTo tratar_erro

If txtLargura = "0,00000" Then txtLargura = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtLargura_LostFocus()
On Error GoTo tratar_erro

txtLargura.Text = IIf(txtLargura = "", "0,00000", Format(txtLargura.Text, "###,##0.0000000000"))

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcSalvar
    'Case 3: ProcAjuda
    Case 4: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
