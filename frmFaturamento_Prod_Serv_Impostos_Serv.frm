VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmFaturamento_Prod_Serv_Impostos_Serv 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Nota fiscal | Impostos sobre serviço"
   ClientHeight    =   4905
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   5025
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
   ScaleHeight     =   4905
   ScaleWidth      =   5025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   33
      Top             =   4500
      Width           =   5025
      _ExtentX        =   8864
      _ExtentY        =   714
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   405
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Width           =   5025
      _ExtentX        =   8864
      _ExtentY        =   714
      DibPicture      =   "frmFaturamento_Prod_Serv_Impostos_Serv.frx":0000
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
      Icon            =   "frmFaturamento_Prod_Serv_Impostos_Serv.frx":7180
      ShowMaximize    =   0   'False
      ShowMinimize    =   0   'False
   End
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   2670
      Top             =   180
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frmFaturamento_Prod_Serv_Impostos_Serv.frx":749A
      Count           =   1
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
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
      Height          =   2415
      Left            =   390
      TabIndex        =   19
      Top             =   1710
      Width           =   4230
      Begin VB.CheckBox chkReter_ISSQN 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Reter imposto"
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
         Left            =   2730
         TabIndex        =   11
         Top             =   1320
         Width           =   1305
      End
      Begin VB.TextBox txtVlr_ISSQN 
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
         Height          =   315
         Left            =   1665
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor total de ISSQN."
         Top             =   1260
         Width           =   945
      End
      Begin VB.TextBox txtISSQN 
         Alignment       =   2  'Center
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
         Height          =   315
         Left            =   780
         MaxLength       =   50
         TabIndex        =   9
         ToolTipText     =   "Porcentagem de imposto."
         Top             =   1260
         Width           =   555
      End
      Begin VB.CheckBox chkReter_IRRF 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Reter imposto"
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
         Left            =   2730
         TabIndex        =   17
         Top             =   2040
         Width           =   1305
      End
      Begin VB.CheckBox chkReter_INSS 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Reter imposto"
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
         Left            =   2730
         TabIndex        =   14
         Top             =   1680
         Width           =   1305
      End
      Begin VB.CheckBox chkReter_CSLL 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Reter imposto"
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
         Left            =   2730
         TabIndex        =   8
         Top             =   960
         Width           =   1305
      End
      Begin VB.CheckBox chkReter_Cofins 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Reter imposto"
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
         Left            =   2730
         TabIndex        =   5
         Top             =   600
         Width           =   1305
      End
      Begin VB.CheckBox chkReter_PIS 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Reter imposto"
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
         Left            =   2730
         TabIndex        =   2
         Top             =   240
         Width           =   1305
      End
      Begin VB.TextBox txtVlr_PIS 
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
         Height          =   315
         Left            =   1665
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   1
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor total de PIS."
         Top             =   180
         Width           =   945
      End
      Begin VB.TextBox txtVlr_Cofins 
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
         Height          =   315
         Left            =   1665
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   4
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor total de cofins."
         Top             =   540
         Width           =   945
      End
      Begin VB.TextBox txtPIS 
         Alignment       =   2  'Center
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
         Height          =   315
         Left            =   780
         MaxLength       =   50
         TabIndex        =   0
         ToolTipText     =   "Porcentagem de imposto."
         Top             =   180
         Width           =   555
      End
      Begin VB.TextBox txtCofins 
         Alignment       =   2  'Center
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
         Height          =   315
         Left            =   780
         MaxLength       =   50
         TabIndex        =   3
         ToolTipText     =   "Porcentagem de imposto."
         Top             =   540
         Width           =   555
      End
      Begin VB.TextBox txtVlr_CSLL 
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
         Height          =   315
         Left            =   1665
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor total de CSLL."
         Top             =   900
         Width           =   945
      End
      Begin VB.TextBox txtCSLL 
         Alignment       =   2  'Center
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
         Height          =   315
         Left            =   780
         MaxLength       =   50
         TabIndex        =   6
         ToolTipText     =   "Porcentagem de imposto."
         Top             =   900
         Width           =   555
      End
      Begin VB.TextBox txtINSS 
         Alignment       =   2  'Center
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
         Height          =   315
         Left            =   780
         MaxLength       =   50
         TabIndex        =   12
         ToolTipText     =   "Porcentagem de imposto."
         Top             =   1620
         Width           =   555
      End
      Begin VB.TextBox txtVlr_INSS 
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
         Height          =   315
         Left            =   1665
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   13
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor total de INSS."
         Top             =   1620
         Width           =   945
      End
      Begin VB.TextBox txtVlr_IRRF 
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
         Height          =   315
         Left            =   1665
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   16
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor total de IRRF."
         Top             =   1980
         Width           =   945
      End
      Begin VB.TextBox txtIRRF 
         Alignment       =   2  'Center
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
         Height          =   315
         Left            =   780
         MaxLength       =   50
         TabIndex        =   15
         ToolTipText     =   "Porcentagem de imposto."
         Top             =   1980
         Width           =   555
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Left            =   1410
         TabIndex        =   31
         Top             =   1320
         Width           =   165
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ISSQN :"
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
         TabIndex        =   30
         Top             =   1260
         Width           =   570
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "PIS :"
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
         Left            =   405
         TabIndex        =   29
         Top             =   180
         Width           =   345
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Cofins :"
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
         Left            =   195
         TabIndex        =   28
         Top             =   540
         Width           =   555
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CSLL :"
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
         Left            =   300
         TabIndex        =   27
         Top             =   900
         Width           =   450
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "INSS :"
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
         Left            =   300
         TabIndex        =   26
         Top             =   1650
         Width           =   450
      End
      Begin VB.Label Label37 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Left            =   1410
         TabIndex        =   25
         Top             =   240
         Width           =   165
      End
      Begin VB.Label Label38 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Left            =   1410
         TabIndex        =   24
         Top             =   600
         Width           =   165
      End
      Begin VB.Label Label39 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Left            =   1410
         TabIndex        =   23
         Top             =   960
         Width           =   165
      End
      Begin VB.Label Label40 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Left            =   1410
         TabIndex        =   22
         Top             =   1710
         Width           =   165
      End
      Begin VB.Label Label36 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Left            =   1410
         TabIndex        =   21
         Top             =   2070
         Width           =   165
      End
      Begin VB.Label Label47 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "IRRF :"
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
         Left            =   285
         TabIndex        =   20
         Top             =   2010
         Width           =   465
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   -30
      TabIndex        =   18
      Top             =   420
      Width           =   5190
      _ExtentX        =   9155
      _ExtentY        =   1720
      ButtonCount     =   6
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
      ButtonCaption2  =   "Excluir"
      ButtonEnabled2  =   0   'False
      ButtonIconSize2 =   32
      ButtonToolTipText2=   "Excluir (F4)"
      ButtonKey2      =   "2"
      ButtonAlignment2=   2
      BeginProperty ButtonFont2 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft2     =   42
      ButtonTop2      =   2
      ButtonWidth2    =   39
      ButtonHeight2   =   21
      ButtonUseMaskColor2=   0   'False
      ButtonEnabled3  =   0   'False
      ButtonIconSize3 =   32
      ButtonAlignment3=   2
      ButtonType3     =   1
      ButtonStyle3    =   -1
      BeginProperty ButtonFont3 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState3    =   -1
      ButtonLeft3     =   83
      ButtonTop3      =   4
      ButtonWidth3    =   2
      ButtonHeight3   =   54
      ButtonCaption4  =   "Ajuda"
      ButtonEnabled4  =   0   'False
      ButtonIconSize4 =   32
      ButtonToolTipText4=   "Ajuda (F1)"
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
      ButtonLeft4     =   87
      ButtonTop4      =   2
      ButtonWidth4    =   36
      ButtonHeight4   =   21
      ButtonUseMaskColor4=   0   'False
      ButtonCaption5  =   "Sair"
      ButtonEnabled5  =   0   'False
      ButtonIconSize5 =   32
      ButtonToolTipText5=   "Sair (Esc)"
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
      ButtonLeft5     =   125
      ButtonTop5      =   2
      ButtonWidth5    =   26
      ButtonHeight5   =   21
      ButtonUseMaskColor5=   0   'False
      ButtonEnabled6  =   0   'False
      ButtonIconSize6 =   32
      ButtonKey6      =   "6"
      ButtonAlignment6=   2
      BeginProperty ButtonFont6 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState6    =   5
      ButtonLeft6     =   153
      ButtonTop6      =   2
      ButtonWidth6    =   24
      ButtonHeight6   =   24
      ButtonUseMaskColor6=   0   'False
   End
End
Attribute VB_Name = "frmFaturamento_Prod_Serv_Impostos_Serv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyF3: ProcSalvar
    Case vbKeyF4: ProcExcluir
    Case vbKeyEscape: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 5200, 6, True
ProcLimpaVariaveisPrincipais
If Formulario = "Estoque/Ordem de faturamento" Then Caption = "Ordem de fat. - Impostos sobre serviço" Else Caption = "Nota fiscal - Impostos sobre serviço"
Qtde = frmFaturamento_Prod_Serv.txtVlrtotalserv
ProcCarregaImposto
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

Direitos
ProcLimpaVariaveisPrincipais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtCofins_Change()
On Error GoTo tratar_erro

If txtCofins.Text <> "" Then
    VerifNumero = txtCofins.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtCofins.Text = ""
        txtCofins.SetFocus
        Exit Sub
    End If
End If
ProcCalculaImposto

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtCSLL_Change()
On Error GoTo tratar_erro

If txtCSLL.Text <> "" Then
    VerifNumero = txtCSLL.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtCSLL.Text = ""
        txtCSLL.SetFocus
        Exit Sub
    End If
End If
ProcCalculaImposto

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtINSS_Change()
On Error GoTo tratar_erro

If txtINSS.Text <> "" Then
    VerifNumero = txtINSS.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtINSS.Text = ""
        txtINSS.SetFocus
        Exit Sub
    End If
End If
ProcCalculaImposto

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtIRRF_Change()
On Error GoTo tratar_erro

If txtIRRF.Text <> "" Then
    VerifNumero = txtIRRF.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtIRRF.Text = ""
        txtIRRF.SetFocus
        Exit Sub
    End If
End If
ProcCalculaImposto

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtISSQN_Change()
On Error GoTo tratar_erro

If txtISSQN.Text <> "" Then
    VerifNumero = txtISSQN.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtISSQN.Text = ""
        txtISSQN.SetFocus
        Exit Sub
    End If
End If
ProcCalculaImposto

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtpis_Change()
On Error GoTo tratar_erro

If txtPIS.Text <> "" Then
    VerifNumero = txtPIS.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtPIS.Text = ""
        txtPIS.SetFocus
        Exit Sub
    End If
End If
ProcCalculaImposto

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcSalvar
    Case 2: ProcExcluir
    'Case 4: ProcAjuda
    Case 5: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvar()
On Error GoTo tratar_erro

With frmFaturamento_Prod_Serv
    If FunVerifValidacaoRegistro("salvar", .txtDtValidacao, IIf(.txtNFiscal = "", "ordem de faturamento", "nota fiscal"), "os impostos", False) = False Then Exit Sub
    
    Set TBGravar = CreateObject("adodb.recordset")
    TBGravar.Open "Select * from tbl_detalhes_Nota where Int_codigo = " & .txtidservico, Conexao, adOpenKeyset, adLockOptimistic
    If TBGravar.EOF = False Then
        'PIS
        TBGravar!PIS_Serv = IIf(txtPIS = "", 0, txtPIS)
        TBGravar!Total_PIS_serv = IIf(txtVlr_PIS = "", 0, txtVlr_PIS)
        If chkReter_PIS.Value = 1 Then TBGravar!Retencao_PIS = True Else TBGravar!Retencao_PIS = False
        'Cofins
        TBGravar!Cofins_Serv = IIf(txtCofins = "", 0, txtCofins)
        TBGravar!Total_Cofins_serv = IIf(txtVlr_Cofins = "", 0, txtVlr_Cofins)
        If chkReter_Cofins.Value = 1 Then TBGravar!Retencao_Cofins = True Else TBGravar!Retencao_Cofins = False
        'CSLL
        TBGravar!CSLL_Serv = IIf(txtCSLL = "", 0, txtCSLL)
        TBGravar!Total_CSLL_serv = IIf(txtVlr_CSLL = "", 0, txtVlr_CSLL)
        If chkReter_CSLL.Value = 1 Then TBGravar!Retencao_CSLL = True Else TBGravar!Retencao_CSLL = False
        'ISSQN
        TBGravar!ISS = IIf(txtISSQN = "", 0, txtISSQN)
        TBGravar!VlrISS = IIf(txtVlr_ISSQN = "", 0, txtVlr_ISSQN)
        If chkReter_ISSQN.Value = 1 Then TBGravar!Retencao_ISSQN = True Else TBGravar!Retencao_ISSQN = False
        'INSS
        TBGravar!INSS_Serv = IIf(txtINSS = "", 0, txtINSS)
        TBGravar!Total_INSS_serv = IIf(txtVlr_INSS = "", 0, txtVlr_INSS)
        If chkReter_INSS.Value = 1 Then TBGravar!Retencao_INSS = True Else TBGravar!Retencao_INSS = False
        'IRRF
        TBGravar!IRRF_Serv = IIf(txtIRRF = "", 0, txtIRRF)
        TBGravar!Total_IRRF_serv = IIf(txtVlr_IRRF = "", 0, txtVlr_IRRF)
        If chkReter_IRRF.Value = 1 Then TBGravar!Retencao_IRRF = True Else TBGravar!Retencao_IRRF = False
        
        TBGravar!Imposto_manual_serv = True
        TBGravar.Update
    End If
    TBGravar.Close
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    '==================================
    Modulo = Formulario
    Evento = "Alterar impostos do serviço"
    ID_documento = .txtidservico
    .ProcVerificaTipoNF False
    If .txtNFiscal = "" Then NomeCampo = "N° ordem: " & .txtID Else NomeCampo = "N° nota: " & .txtNFiscal
    Documento = NomeCampo & " - Tipo: " & TipoNF & " - Série: " & .txtSerie
    Documento1 = "Cód. interno: " & .txtcodServ
    ProcGravaEvento
    '==================================
    .ProcCarregaListaServicos
    .ProcGravarTotaisNota
    .ProcCarregaTotaisNota IIf(.txtID = "", 0, .txtID)
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluir()
On Error GoTo tratar_erro

With frmFaturamento_Prod_Serv
    If FunVerifValidacaoRegistro("excluir", .txtDtValidacao, IIf(.txtNFiscal = "", "ordem de faturamento", "nota fiscal"), "os impostos", False) = False Then Exit Sub
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select * from tbl_detalhes_Nota where Int_codigo = " & .txtidservico, Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
        TBProduto!Retencao_PIS = False
        TBProduto!Retencao_Cofins = False
        TBProduto!Retencao_CSLL = False
        TBProduto!Retencao_ISSQN = False
        TBProduto!Retencao_INSS = False
        TBProduto!Retencao_IRRF = False
        TBProduto!Imposto_manual_serv = False
        .ProcSalvarImpostosServ
        TBProduto.Update
    End If
    TBProduto.Close
    USMsgBox ("Alteração excluir com sucesso."), vbInformation, "CAPRIND v5.0"
    '==================================
    Modulo = Formulario
    Evento = "Excluir alteração dos impostos do serviço"
    ID_documento = .txtidservico
    .ProcVerificaTipoNF False
    If .txtNFiscal = "" Then NomeCampo = "N° ordem: " & .txtID Else NomeCampo = "N° nota: " & .txtNFiscal
    Documento = NomeCampo & " - Tipo: " & TipoNF & " - Série: " & .txtSerie
    Documento1 = "Cód. interno: " & .txtcodServ
    ProcGravaEvento
    '==================================
    ProcCarregaImposto
    .ProcCarregaListaServicos
    .ProcGravarTotaisNota
    .ProcCarregaTotaisNota IIf(.txtID = "", 0, .txtID)
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaImposto()
On Error GoTo tratar_erro

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from tbl_detalhes_Nota where Int_codigo = " & frmFaturamento_Prod_Serv.txtidservico, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    'PIS
    txtPIS = IIf(IsNull(TBAbrir!PIS_Serv), "", TBAbrir!PIS_Serv)
    txtVlr_PIS = IIf(IsNull(TBAbrir!Total_PIS_serv), "", Format(TBAbrir!Total_PIS_serv, "###,##0.00"))
    If TBAbrir!Retencao_PIS = True Then chkReter_PIS.Value = 1 Else chkReter_PIS.Value = 0
    'Cofins
    txtCofins = IIf(IsNull(TBAbrir!Cofins_Serv), "", TBAbrir!Cofins_Serv)
    txtVlr_Cofins = IIf(IsNull(TBAbrir!Total_Cofins_serv), "", Format(TBAbrir!Total_Cofins_serv, "###,##0.00"))
    If TBAbrir!Retencao_Cofins = True Then chkReter_Cofins.Value = 1 Else chkReter_Cofins.Value = 0
    'CSLL
    txtCSLL = IIf(IsNull(TBAbrir!CSLL_Serv), "", TBAbrir!CSLL_Serv)
    txtVlr_CSLL = IIf(IsNull(TBAbrir!Total_CSLL_serv), "", Format(TBAbrir!Total_CSLL_serv, "###,##0.00"))
    If TBAbrir!Retencao_CSLL = True Then chkReter_CSLL.Value = 1 Else chkReter_CSLL.Value = 0
    'ISSQN
    txtISSQN = IIf(IsNull(TBAbrir!ISS), "", TBAbrir!ISS)
    txtVlr_ISSQN = IIf(IsNull(TBAbrir!VlrISS), "", Format(TBAbrir!VlrISS, "###,##0.00"))
    If TBAbrir!Retencao_ISSQN = True Then chkReter_ISSQN.Value = 1 Else chkReter_ISSQN.Value = 0
    'INSS
    txtINSS = IIf(IsNull(TBAbrir!INSS_Serv), "", TBAbrir!INSS_Serv)
    txtVlr_INSS = IIf(IsNull(TBAbrir!Total_INSS_serv), "", Format(TBAbrir!Total_INSS_serv, "###,##0.00"))
    If TBAbrir!Retencao_INSS = True Then chkReter_INSS.Value = 1 Else chkReter_INSS.Value = 0
    'IRRF
    txtIRRF = IIf(IsNull(TBAbrir!IRRF_Serv), "", TBAbrir!IRRF_Serv)
    txtVlr_IRRF = IIf(IsNull(TBAbrir!Total_IRRF_serv), "", Format(TBAbrir!Total_IRRF_serv, "###,##0.00"))
    If TBAbrir!Retencao_IRRF = True Then chkReter_IRRF.Value = 1 Else chkReter_IRRF.Value = 0
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCalculaImposto()
On Error GoTo tratar_erro

'PIS
qt = 0
Qtd = IIf(txtPIS = "", 0, txtPIS)
If Qtd > 0 And Qtde > 0 Then qt = (Qtde * Qtd) / 100
txtVlr_PIS = Format(qt, "###,##0.00")
'Cofins
qt = 0
Qtd = IIf(txtCofins = "", 0, txtCofins)
If Qtd > 0 And Qtde > 0 Then qt = (Qtde * Qtd) / 100
txtVlr_Cofins = Format(qt, "###,##0.00")
'CSLL
qt = 0
Qtd = IIf(txtCSLL = "", 0, txtCSLL)
If Qtd > 0 And Qtde > 0 Then qt = (Qtde * Qtd) / 100
txtVlr_CSLL = Format(qt, "###,##0.00")
'ISSQN
qt = 0
Qtd = IIf(txtISSQN = "", 0, txtISSQN)
If Qtd > 0 And Qtde > 0 Then qt = (Qtde * Qtd) / 100
txtVlr_ISSQN = Format(qt, "###,##0.00")
'INSS
qt = 0
Qtd = IIf(txtINSS = "", 0, txtINSS)
If Qtd > 0 And Qtde > 0 Then qt = (Qtde * Qtd) / 100
txtVlr_INSS = Format(qt, "###,##0.00")
'IRRF
qt = 0
Qtd = IIf(txtIRRF = "", 0, txtIRRF)
If Qtd > 0 And Qtde > 0 Then qt = (Qtde * Qtd) / 100
txtVlr_IRRF = Format(qt, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
