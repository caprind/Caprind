VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form frmprod_alterarRM 
   BackColor       =   &H00E0E0E0&
   Caption         =   "PCP - Requisição da ordem"
   ClientHeight    =   10035
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15270
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmprod_alterarRM.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10035
   ScaleWidth      =   15270
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   99
      ScreenHeight    =   768
      ScreenWidth     =   1366
      ScreenHeightDT  =   1080
      ScreenWidthDT   =   1920
      AutoResizeOnLoad=   0   'False
      ApplicationName =   "Active Resize Control Professional"
      FormHeightDT    =   10500
      FormWidthDT     =   15390
      FormScaleHeightDT=   10035
      FormScaleWidthDT=   15270
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin VB.CheckBox Chk_servico 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Serviço"
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
      Left            =   4410
      TabIndex        =   6
      Top             =   1080
      Width           =   825
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   90
      TabIndex        =   53
      Top             =   9180
      Width           =   15195
      Begin VB.CommandButton cmdDesenho 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   2280
         Picture         =   "frmprod_alterarRM.frx":1042
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Localizar produtos."
         Top             =   375
         Width           =   315
      End
      Begin VB.TextBox txtDesenho_Similar 
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   180
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   26
         TabStop         =   0   'False
         ToolTipText     =   "Código interno."
         Top             =   375
         Width           =   2085
      End
      Begin VB.TextBox txtDesc_Similar 
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
         Left            =   2700
         Locked          =   -1  'True
         MaxLength       =   150
         TabIndex        =   28
         TabStop         =   0   'False
         ToolTipText     =   "Descrição."
         Top             =   375
         Width           =   8435
      End
      Begin VB.TextBox txtQtde_Similar 
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
         Left            =   12615
         MaxLength       =   50
         TabIndex        =   31
         ToolTipText     =   "Quantidade."
         Top             =   375
         Width           =   1185
      End
      Begin VB.TextBox txtQtde_PC_Similar 
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
         Left            =   13815
         MaxLength       =   50
         TabIndex        =   32
         ToolTipText     =   "Quantidade de peças."
         Top             =   375
         Width           =   1185
      End
      Begin VB.TextBox txtUn_Similar 
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   11145
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   29
         ToolTipText     =   "Unidade de estoque."
         Top             =   375
         Width           =   720
      End
      Begin VB.TextBox txtUN_com_Similar 
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   11880
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   30
         ToolTipText     =   "Unidade comercial."
         Top             =   375
         Width           =   720
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição"
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
         Left            =   6577
         TabIndex        =   59
         Top             =   180
         Width           =   690
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantidade"
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
         Left            =   12780
         TabIndex        =   58
         Top             =   180
         Width           =   840
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Un. est."
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
         Left            =   11228
         TabIndex        =   57
         Top             =   180
         Width           =   585
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Un. com."
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
         Left            =   11933
         TabIndex        =   56
         Top             =   180
         Width           =   645
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantidade PÇ"
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
         Left            =   13860
         TabIndex        =   55
         Top             =   180
         Width           =   1080
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cód. interno"
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
         Left            =   772
         TabIndex        =   54
         Top             =   180
         Width           =   900
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
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
      Height          =   825
      Left            =   90
      TabIndex        =   45
      Top             =   8340
      Width           =   15195
      Begin VB.TextBox Txt_ID 
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
         Left            =   3513
         Locked          =   -1  'True
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   375
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.TextBox txtQtde 
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
         Left            =   12615
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   24
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade."
         Top             =   375
         Width           =   1185
      End
      Begin VB.TextBox txtDescricao 
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
         Left            =   3513
         Locked          =   -1  'True
         MaxLength       =   150
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   "Descrição."
         Top             =   375
         Width           =   7620
      End
      Begin VB.TextBox txtDesenho 
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1684
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   20
         TabStop         =   0   'False
         ToolTipText     =   "Código interno."
         Top             =   375
         Width           =   1815
      End
      Begin VB.TextBox txtUN 
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   11145
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   "Unidade de estoque."
         Top             =   375
         Width           =   720
      End
      Begin VB.TextBox txtOrdem 
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
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "Ordem."
         Top             =   375
         Width           =   1490
      End
      Begin VB.TextBox txtQtde_PC 
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
         Left            =   13815
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   25
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade de peças."
         Top             =   375
         Width           =   1185
      End
      Begin VB.TextBox txtUn_com 
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   11880
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         ToolTipText     =   "Unidade comercial."
         Top             =   375
         Width           =   720
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantidade"
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
         Left            =   12780
         TabIndex        =   52
         Top             =   180
         Width           =   840
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cód. interno"
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
         Left            =   2137
         TabIndex        =   51
         Top             =   180
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição"
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
         Left            =   6978
         TabIndex        =   50
         Top             =   180
         Width           =   690
      End
      Begin VB.Label Label34 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Un. est."
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
         Left            =   11213
         TabIndex        =   49
         Top             =   180
         Width           =   585
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ordem"
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
         Left            =   633
         TabIndex        =   48
         Top             =   180
         Width           =   585
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantidade PÇ"
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
         Left            =   13860
         TabIndex        =   47
         Top             =   180
         Width           =   1080
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Un. com."
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
         Left            =   11918
         TabIndex        =   46
         Top             =   180
         Width           =   645
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   90
      TabIndex        =   41
      Top             =   7440
      Width           =   15195
      Begin VB.TextBox txtPagIr 
         Height          =   315
         Left            =   9540
         TabIndex        =   13
         ToolTipText     =   "Número da página."
         Top             =   180
         Width           =   555
      End
      Begin VB.TextBox txtNreg 
         Alignment       =   2  'Center
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
         Left            =   3780
         TabIndex        =   12
         Text            =   "30"
         ToolTipText     =   "Número de registros por página."
         Top             =   180
         Width           =   555
      End
      Begin DrawSuite2022.USButton cmdPagProx 
         Height          =   315
         Left            =   11760
         TabIndex        =   17
         ToolTipText     =   "Próxima página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmprod_alterarRM.frx":1144
         BorderColor     =   14404026
         BorderColorDown =   11632444
         BorderColorOver =   11632444
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
         GradientColor2  =   16777215
         GradientColor3  =   16777215
         GradientColorDown2=   16246986
         GradientColorDown3=   15189380
         GradientColorDown4=   14596208
         GradientColorOver1=   16643560
         GradientColorOver2=   16576988
         GradientColorOver3=   16441780
         GradientColorOver4=   16178091
         PicSizeH        =   19
         PicSizeW        =   19
      End
      Begin DrawSuite2022.USButton cmdPagAnt 
         Height          =   315
         Left            =   11220
         TabIndex        =   16
         ToolTipText     =   "Página anterior."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmprod_alterarRM.frx":48E8
         BorderColor     =   14404026
         BorderColorDown =   11632444
         BorderColorOver =   11632444
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
         GradientColor2  =   16777215
         GradientColor3  =   16777215
         GradientColorDown2=   16246986
         GradientColorDown3=   15189380
         GradientColorDown4=   14596208
         GradientColorOver1=   16643560
         GradientColorOver2=   16576988
         GradientColorOver3=   16441780
         GradientColorOver4=   16178091
         PicSizeH        =   19
         PicSizeW        =   19
      End
      Begin DrawSuite2022.USButton cmdPagIr 
         Height          =   315
         Left            =   10110
         TabIndex        =   14
         Top             =   180
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   556
         BorderColor     =   14404026
         BorderColorDown =   11632444
         BorderColorOver =   11632444
         Caption         =   "Ir"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColor2  =   16777215
         GradientColor3  =   16777215
         GradientColorDown2=   16246986
         GradientColorDown3=   15189380
         GradientColorDown4=   14596208
         GradientColorOver1=   16643560
         GradientColorOver2=   16576988
         GradientColorOver3=   16441780
         GradientColorOver4=   16178091
         PicSizeH        =   19
         PicSizeW        =   19
      End
      Begin DrawSuite2022.USButton cmdPagPrim 
         Height          =   315
         Left            =   10680
         TabIndex        =   15
         ToolTipText     =   "Primeira página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmprod_alterarRM.frx":83F1
         BorderColor     =   14404026
         BorderColorDown =   11632444
         BorderColorOver =   11632444
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
         GradientColor2  =   16777215
         GradientColor3  =   16777215
         GradientColorDown2=   16246986
         GradientColorDown3=   15189380
         GradientColorDown4=   14596208
         GradientColorOver1=   16643560
         GradientColorOver2=   16576988
         GradientColorOver3=   16441780
         GradientColorOver4=   16178091
         PicSizeH        =   19
         PicSizeW        =   19
      End
      Begin DrawSuite2022.USButton cmdPagUlt 
         Height          =   315
         Left            =   12300
         TabIndex        =   18
         ToolTipText     =   "Última página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmprod_alterarRM.frx":C4E0
         BorderColor     =   14404026
         BorderColorDown =   11632444
         BorderColorOver =   11632444
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
         GradientColor2  =   16777215
         GradientColor3  =   16777215
         GradientColorDown2=   16246986
         GradientColorDown3=   15189380
         GradientColorDown4=   14596208
         GradientColorOver1=   16643560
         GradientColorOver2=   16576988
         GradientColorOver3=   16441780
         GradientColorOver4=   16178091
         PicSizeH        =   19
         PicSizeW        =   19
      End
      Begin VB.Label lblPaginas 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Página: 0 de: 0"
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
         Left            =   13050
         TabIndex        =   44
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblRegistros 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nº de registros: 0"
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
         TabIndex        =   43
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Carregar               registros por página"
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
         Left            =   3090
         TabIndex        =   42
         Top             =   240
         Width           =   2760
      End
   End
   Begin VB.CheckBox ChkData1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Prazo entrega"
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
      Left            =   13680
      TabIndex        =   8
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CheckBox ChkData 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Dt. emissão"
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
      Left            =   12450
      TabIndex        =   7
      Top             =   1080
      Width           =   1155
   End
   Begin VB.CheckBox chkMontagem 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Subconjunto"
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
      Left            =   1650
      TabIndex        =   4
      Top             =   1080
      Width           =   1245
   End
   Begin VB.CheckBox chkExpedicao 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Produto final"
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
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CheckBox chkFabricacao 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Componente"
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
      Left            =   3030
      TabIndex        =   5
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Height          =   885
      Left            =   90
      TabIndex        =   34
      Top             =   1290
      Width           =   12015
      Begin VB.Frame Frame5 
         BackColor       =   &H00E0E0E0&
         Height          =   510
         Left            =   2400
         TabIndex        =   61
         Top             =   240
         Width           =   4785
         Begin VB.OptionButton optIgual 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Igual"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3930
            TabIndex        =   65
            Top             =   180
            Width           =   705
         End
         Begin VB.OptionButton Optmeio 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Meio frase"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1470
            TabIndex        =   64
            Top             =   180
            Width           =   1275
         End
         Begin VB.OptionButton Optinicio 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Início frase"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   180
            TabIndex        =   63
            Top             =   180
            Value           =   -1  'True
            Width           =   1275
         End
         Begin VB.OptionButton Optfim 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Fim frase"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2760
            TabIndex        =   62
            Top             =   180
            Width           =   1155
         End
      End
      Begin VB.ComboBox cmbfiltrarpor 
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
         ItemData        =   "frmprod_alterarRM.frx":FD6C
         Left            =   180
         List            =   "frmprod_alterarRM.frx":FDA0
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Opções para filtro."
         Top             =   390
         Width           =   2145
      End
      Begin VB.TextBox txtTexto 
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
         Left            =   7260
         TabIndex        =   1
         ToolTipText     =   "Texto para pesquisa."
         Top             =   390
         Width           =   4545
      End
      Begin VB.ComboBox cmbfamilia 
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
         Height          =   330
         Left            =   7260
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "Texto para pesquisa."
         Top             =   390
         Visible         =   0   'False
         Width           =   4545
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Texto para pesquisa"
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
         Left            =   8797
         TabIndex        =   36
         Top             =   180
         Width           =   1470
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filtrar por"
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
         Left            =   832
         TabIndex        =   35
         Top             =   180
         Width           =   840
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   885
      Left            =   12120
      TabIndex        =   33
      Top             =   1290
      Width           =   3165
      Begin MSComCtl2.DTPicker msk_fltFim 
         Height          =   315
         Left            =   1590
         TabIndex        =   10
         ToolTipText     =   "Data final para pesquisa."
         Top             =   390
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   16777215
         CalendarForeColor=   0
         CalendarTitleBackColor=   8421504
         CalendarTitleForeColor=   16777215
         CalendarTrailingForeColor=   255
         Format          =   490733569
         CurrentDate     =   39057
      End
      Begin MSComCtl2.DTPicker msk_fltInicio 
         Height          =   315
         Left            =   180
         TabIndex        =   9
         ToolTipText     =   "Data início para pesquisa."
         Top             =   390
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   16777215
         CalendarForeColor=   0
         CalendarTitleBackColor=   8421504
         CalendarTitleForeColor=   16777215
         CalendarTrailingForeColor=   255
         Format          =   490733571
         CurrentDate     =   39057
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Até"
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
         Left            =   2160
         TabIndex        =   40
         Top             =   180
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "De"
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
         Left            =   780
         TabIndex        =   39
         Top             =   180
         Width           =   195
      End
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   60
      TabIndex        =   37
      Top             =   8070
      Width           =   15195
      _ExtentX        =   26802
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
      SearchText      =   ""
      Value           =   0
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   90
      TabIndex        =   38
      Top             =   0
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   1720
      ButtonCount     =   6
      GradientColor2  =   14737632
      GradientColorOverRight1=   16315633
      GradientColorOverRight2=   15195350
      GripperColor    =   15195350
      IsStrech        =   -1  'True
      RightColor1     =   15195350
      RightColor2     =   16315633
      ShowEndPanel    =   0   'False
      ShowGripper     =   0   'False
      Theme           =   1
      ButtonCaption1  =   "Filtrar"
      ButtonEnabled1  =   0   'False
      ButtonIconSize1 =   32
      ButtonToolTipText1=   "Filtrar (F2)"
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
      ButtonWidth1    =   36
      ButtonHeight1   =   21
      ButtonUseMaskColor1=   0   'False
      ButtonCaption2  =   "Salvar"
      ButtonEnabled2  =   0   'False
      ButtonIconSize2 =   32
      ButtonToolTipText2=   "Salvar (F3)"
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
      ButtonLeft2     =   40
      ButtonTop2      =   2
      ButtonWidth2    =   38
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
      ButtonLeft3     =   80
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
      ButtonLeft4     =   84
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
      ButtonLeft5     =   122
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
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState6    =   5
      ButtonLeft6     =   150
      ButtonTop6      =   2
      ButtonWidth6    =   24
      ButtonHeight6   =   24
      ButtonUseMaskColor6=   0   'False
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   7920
         Top             =   180
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmprod_alterarRM.frx":FEA5
         Count           =   1
      End
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   5205
      Left            =   90
      TabIndex        =   11
      Top             =   2190
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   9181
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   0
      BackColor       =   16777215
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   14
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "N"
         Text            =   "Cód. interno"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Descrição"
         Object.Width           =   4472
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Object.Tag             =   "N"
         Text            =   "Qtde."
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Un."
         Object.Width           =   970
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Object.Tag             =   "N"
         Text            =   "Qtde. PÇ"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Object.Tag             =   "N"
         Text            =   "Ordem"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Object.Tag             =   "D"
         Text            =   "Prazo"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Object.Tag             =   "T"
         Text            =   "Cód. interno"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   9
         Object.Tag             =   "T"
         Text            =   "Rev."
         Object.Width           =   970
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Object.Tag             =   "T"
         Text            =   "Cód. ref."
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Object.Tag             =   "T"
         Text            =   "Descrição"
         Object.Width           =   4472
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Object.Tag             =   "T"
         Text            =   "Cliente"
         Object.Width           =   3590
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Object.Tag             =   "N"
         Text            =   "IDempresa"
         Object.Width           =   0
      EndProperty
   End
End
Attribute VB_Name = "frmprod_alterarRM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StrSql_PCPRM As String 'OK
Dim TBLISTA_PCPRM  As ADODB.Recordset 'OK

Private Sub Chk_servico_Click()
On Error GoTo tratar_erro

ProcLimparCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ChkData_Click()
On Error GoTo tratar_erro

ProcLimparCampos
If chkData.Value = 1 Then
    ChkData1.Value = 0
    Frame3.Enabled = True
    msk_fltInicio.SetFocus
Else
    Frame3.Enabled = False
    msk_fltInicio.Value = Date
    msk_fltFim.Value = Date
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ChkData1_Click()
On Error GoTo tratar_erro

ProcLimparCampos
If ChkData1.Value = 1 Then
    chkData.Value = 0
    Frame3.Enabled = True
    msk_fltInicio.SetFocus
Else
    msk_fltInicio.Value = Date
    msk_fltFim.Value = Date
    Frame3.Enabled = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkExpedicao_Click()
On Error GoTo tratar_erro

ProcLimparCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkFabricacao_Click()
On Error GoTo tratar_erro

ProcLimparCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkMontagem_Click()
On Error GoTo tratar_erro

ProcLimparCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbFamilia_Click()
On Error GoTo tratar_erro

ProcLimparCampos
If cmbfamilia <> "" Then txtTexto = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbfiltrarpor_Click()
On Error GoTo tratar_erro

ProcLimparCampos
If (cmbfiltrarpor = "Ordem" Or cmbfiltrarpor = "OS") And txtTexto <> "" Then
    VerifNumero = txtTexto
    ProcVerificaNumero
    If VerifNumero = False Then
        txtTexto = ""
        txtTexto.SetFocus
        Exit Sub
    End If
End If
With cmbfamilia
    If cmbfiltrarpor = "Família" Or cmbfiltrarpor = "Família do material" Or cmbfiltrarpor = "Prioridade" Or cmbfiltrarpor = "Reposição" Then
        txtTexto.Visible = False
        .Visible = True
        .Clear
        .AddItem ""
        If cmbfiltrarpor = "Família" Then
            ProcCarregaComboFamilia cmbfamilia, "familia <> 'Null' and Fabricacao = 'True'", True
        ElseIf cmbfiltrarpor = "Família do material" Then
                ProcCarregaComboFamilia cmbfamilia, "familia <> 'Null' and Compras = 'True'", True
            ElseIf cmbfiltrarpor = "Prioridade" Then
                    .AddItem "Urgente"
                    .AddItem "Normal"
                ElseIf cmbfiltrarpor = "Reposição" Then
                        .AddItem "Sim"
                        .AddItem "Não"
        End If
    Else
        txtTexto.Visible = True
        .Visible = False
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdDesenho_Click()
On Error GoTo tratar_erro

PCP_AlterarRM = True
Frm_necessidade_Item.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_PCPRM.AbsolutePage <> 2 Then
    If TBLISTA_PCPRM.AbsolutePage = -3 Then
        ProcExibePagina (TBLISTA_PCPRM.PageCount - 1)
    Else
        TBLISTA_PCPRM.AbsolutePage = TBLISTA_PCPRM.AbsolutePage - 2
        ProcExibePagina (TBLISTA_PCPRM.AbsolutePage)
    End If
Else
    ProcExibePagina (1)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagIr_Click()
On Error GoTo tratar_erro

If txtPagIr = "" Then Exit Sub
Quant = ReturnNumbersOnly(Right(lblPaginas.Caption, 4))
If Quant <= 1 Or txtPagIr > Quant Then Exit Sub
If txtPagIr.Text >= 1 And txtPagIr.Text <= Quant Then
    TBLISTA_PCPRM.AbsolutePage = txtPagIr.Text
    ProcExibePagina (TBLISTA_PCPRM.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_PCPRM.AbsolutePage = 1
ProcExibePagina (TBLISTA_PCPRM.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_PCPRM.AbsolutePage <> -3 Then
    If TBLISTA_PCPRM.AbsolutePage = 1 Then
        ProcExibePagina (2)
    Else
        ProcExibePagina (TBLISTA_PCPRM.AbsolutePage)
    End If
Else
    ProcExibePagina (TBLISTA_PCPRM.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_PCPRM.AbsolutePage = TBLISTA_PCPRM.PageCount
ProcExibePagina (TBLISTA_PCPRM.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro
  
Select Case KeyCode
    'Case vbKeyF1: ProcAjuda
    Case vbKeyEscape: ProcSair
    Case vbKeyF2: ProcFiltrar
    Case vbKeyF3: ProcSalvar
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15192, 6, True

Formulario = "PCP/Requisição da ordem"
ProcLimpaVariaveisPrincipais
cmbfiltrarpor = "Código interno"
msk_fltInicio.Value = Date
msk_fltFim.Value = Date

ProcRemoveObjetosResize Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

With msk_fltFim
    If FunVerificaDataFinal(msk_fltInicio.Value, .Value) = False Then
        .Value = Date
        .SetFocus
        Exit Sub
    End If
End With

TipoFiltro = ""
If chkFabricacao.Value = 1 Then TipoFiltro = "PROD.tipo = 'F'"
If chkMontagem.Value = 1 Then
    If TipoFiltro <> "" Then TipoFiltro = TipoFiltro & " or PROD.tipo = 'M'" Or TipoFiltro = "PROD.tipo = 'M'"
End If
If chkExpedicao.Value = 1 Then
    If TipoFiltro <> "" Then TipoFiltro = TipoFiltro & " or PROD.tipo = 'E'" Or TipoFiltro = "PROD.tipo = 'E'"
End If
If Chk_servico.Value = 1 Then
    If TipoFiltro <> "" Then TipoFiltro = TipoFiltro & " or PROD.tipo = 'S'" Or TipoFiltro = "PROD.tipo = 'S'"
End If
If TipoFiltro <> "" Then TipoFiltro = "(" & TipoFiltro & ")"

DataFiltro = ""
If chkData.Value = 1 Or ChkData1.Value = 1 Then
    If chkData.Value = 1 Then Pesquisa_ordem = "PROD.data"
    If ChkData1.Value = 1 Then Pesquisa_ordem = "ENDOP.Prazo"
    DataFiltro = "and " & Pesquisa_ordem & " Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "'"
End If

CamposFiltro = "ENDOP.ID, ENDOP.Desenho, ENDOP.descricao, ENDOP.Unidade, ENDOP.Requisitado, ENDOP.Requisitado_PC, ENDOP.Ordem_req, ENDOP.Prazo, ENDOP.Desenho_doc_req, ENDOP.Revitem_doc_req, ENDOP.Ref_doc_req, ENDOP.Descricao_doc_req, ENDOP.Cliente_doc_req, ENDOP.ID_empresa"
INNERJOINTEXTO = "Select " & CamposFiltro & " from ((((((((Estoque_necessidade_detalhado_OP ENDOP LEFT JOIN Producao PROD ON PROD.Ordem = ENDOP.ID) LEFT JOIN Projproduto P ON P.Desenho = ENDOP.Desenho_doc_req) LEFT JOIN item_aplicacoes IA ON IA.codproduto = ENDOP.codproduto) LEFT JOIN Ordemservico OS ON OS.Ordem = PROD.Ordem) LEFT JOIN Producao_pedidos PPE ON PPE.Ordem = PROD.Ordem) LEFT JOIN Vendas_carteira VC ON VC.codigo = PPE.IDCarteira) LEFT JOIN Vendas_proposta VP ON VP.Cotacao = VC.Cotacao) LEFT JOIN Outros_SolicitacaoPCP OSP ON OSP.ID = VC.ID_solicitacao) LEFT JOIN Projproduto_fabricante PFAB ON PFAB.Codproduto = P.codproduto"
TextoFiltroPadrao = TipoFiltro & " group by " & CamposFiltro & " order by ENDOP.ID"

If txtTexto.Visible = True And txtTexto <> "" Or cmbfamilia.Visible = True And cmbfamilia <> "" Then
    If Left(cmbfiltrarpor, 7) = "Família" Then
        If cmbfiltrarpor = "Família" Then TextoFiltro = "ENDOP.Classe" Else TextoFiltro = "P.Classe"
        StrSql_PCPRM = INNERJOINTEXTO & " where " & TextoFiltro & " = '" & cmbfamilia & "' " & IIf(TipoFiltro = "", TextoFiltroPadrao, "and " & TextoFiltroPadrao)
    ElseIf cmbfiltrarpor = "Prioridade" Then
            If cmbfamilia = "Urgente" Then TextoFiltro = "PROD.IMPREQ = 'True'" Else TextoFiltro = "PROD.IMPREQ = 'False'"
            StrSql_PCPRM = INNERJOINTEXTO & " where " & TextoFiltro & IIf(TipoFiltro = "", TextoFiltroPadrao, " and " & TextoFiltroPadrao)
        ElseIf cmbfiltrarpor = "Reposição" Then
                If cmbfamilia = "Sim" Then TextoFiltro = "PROD.Reposicao = 'True'" Else TextoFiltro = "PROD.Reposicao = 'False'"
                StrSql_PCPRM = INNERJOINTEXTO & " where " & TextoFiltro & IIf(TipoFiltro = "", TextoFiltroPadrao, " and " & TextoFiltroPadrao)
            Else
                Select Case cmbfiltrarpor
                    Case "Código interno": TextoFiltro = "ENDOP.Desenho"
                    Case "Código interno do produto": TextoFiltro = "ENDOP.Desenho_doc_req"
                    Case "Código de referência": TextoFiltro = "IA.n_referencia"
                    Case "Código de referência do produto": TextoFiltro = "ENDOP.Ref_doc_req"
                    Case "Descrição": TextoFiltro = "ENDOP.Descricao"
                    Case "Descrição do produto": TextoFiltro = "ENDOP.Descricao_doc_req"
                    Case "Cliente": TextoFiltro = "ENDOP.Cliente_doc_req"
                    Case "Pedido interno": TextoFiltro = "VP.Ncotacao"
                    Case "Solicitação de produção": TextoFiltro = "OSP.Requisicaotexto"
                    Case "Part number": TextoFiltro = "PFAB.Part_number"
                    Case "Ordem de produção": TextoFiltro = "ENDOP.ID"
                    Case "OS": TextoFiltro = "OS.Idproducao"
                End Select
                StrSql_PCPRM = INNERJOINTEXTO & " where " & TextoFiltro & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & IIf(TipoFiltro = "", TextoFiltroPadrao, " and " & TextoFiltroPadrao)
    End If
Else
    StrSql_PCPRM = INNERJOINTEXTO & IIf(TipoFiltro = "", TextoFiltroPadrao, " where " & TextoFiltroPadrao)
End If
ProcCarregaLista (1)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

Formulario = "PCP/Requisição da ordem"
ProcLimpaVariaveisPrincipais

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

Private Sub Lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView Lista, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With Lista
    Txt_ID = .SelectedItem
    txtOrdem = .SelectedItem.SubItems(6)
    txtdesenho = .SelectedItem.SubItems(1)
    txtQtde = .SelectedItem.SubItems(3)
    txtQtde_PC = .SelectedItem.SubItems(5)
End With

Set TBItem = CreateObject("adodb.recordset")
TBItem.Open "select Unidade, unidade_com, descricao from projproduto where desenho = '" & txtdesenho & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBItem.EOF = False Then
    txtUN = IIf(IsNull(TBItem!Unidade), "", TBItem!Unidade)
    txtUn_com = IIf(IsNull(TBItem!Unidade_com), "", TBItem!Unidade_com)
    txtdescricao = IIf(IsNull(TBItem!Descricao), "", TBItem!Descricao)
End If
TBItem.Close

If FunVerifMovimentacaoEstPC(Lista.SelectedItem.ListSubItems(13)) = True And txtQtde_PC > 0 Then
    With txtQtde_Similar
        .Locked = True
        .TabStop = False
    End With
    With txtQtde_PC_Similar
        .Locked = False
        .TabStop = True
    End With
Else
    With txtQtde_Similar
        .Locked = False
        .TabStop = True
    End With
    With txtQtde_PC_Similar
        .Locked = True
        .TabStop = False
    End With
End If
Frame1.Enabled = True
Frame6.Enabled = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub msk_fltInicio_Click()
On Error GoTo tratar_erro

ProcLimparCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optfim_Click()
On Error GoTo tratar_erro

ProcLimparCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optinicio_Click()
On Error GoTo tratar_erro

ProcLimparCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optmeio_Click()
On Error GoTo tratar_erro

ProcLimparCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtNreg_Change()
On Error GoTo tratar_erro

If txtNreg <> "" Then
    VerifNumero = txtNreg
    ProcVerificaNumero
    If VerifNumero = False Then
        txtNreg = ""
        txtNreg.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtPagIr_Change()
On Error GoTo tratar_erro

If txtPagIr <> "" Then
    VerifNumero = txtPagIr
    ProcVerificaNumero
    If VerifNumero = False Then
        txtPagIr = ""
        txtPagIr.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtQtde_PC_Similar_Change()
On Error GoTo tratar_erro

If txtQtde_PC_Similar.Text <> "" Then
    VerifNumero = txtQtde_PC_Similar.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtQtde_PC_Similar.Text = ""
        txtQtde_PC_Similar.SetFocus
        Exit Sub
    End If
    If txtQtde_PC_Similar.Locked = False Then ProcCalculaQtdesSimilar txtDesenho_Similar, txtQtde_Similar, txtQtde_PC_Similar
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtQtde_Similar_Change()
On Error GoTo tratar_erro

If txtQtde_Similar.Text <> "" Then
    VerifNumero = txtQtde_Similar.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtQtde_Similar.Text = ""
        txtQtde_Similar.SetFocus
        Exit Sub
    End If
    If txtQtde_Similar.Locked = False Then ProcCalculaQtdesSimilar txtDesenho_Similar, txtQtde_Similar, txtQtde_PC_Similar
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtTexto_Change()
On Error GoTo tratar_erro

ProcLimparCampos
If txtTexto <> "" Then
    If cmbfiltrarpor = "Ordem" Or cmbfiltrarpor = "OS" Then
        VerifNumero = txtTexto
        ProcVerificaNumero
        If VerifNumero = False Then
            txtTexto = ""
            txtTexto.SetFocus
            Exit Sub
        End If
    End If
    cmbfamilia.ListIndex = -1
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaLista(Pagina As Integer)
On Error GoTo tratar_erro

If StrSql_PCPRM = "" Then Exit Sub
lblRegistros.Caption = "Nº de registros: 0"
lblPaginas.Caption = "Página: 0 de: 0"
Lista.ListItems.Clear
If StrSql_PCPRM = "" Then Exit Sub
Set TBLISTA_PCPRM = CreateObject("adodb.recordset")
TBLISTA_PCPRM.Open StrSql_PCPRM, Conexao, adOpenKeyset, adLockReadOnly
If TBLISTA_PCPRM.EOF = False Then ProcExibePagina (Pagina)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina(Pagina)
On Error GoTo tratar_erro

Lista.ListItems.Clear
TBLISTA_PCPRM.PageSize = IIf(txtNreg = "", 30, txtNreg)
TBLISTA_PCPRM.AbsolutePage = Pagina
TamanhoPagina = TBLISTA_PCPRM.PageSize
ContadorReg = 1

PBLista.Min = 0
PBLista.Max = FunVerifMaxPBListaPaginacao(TBLISTA_PCPRM.RecordCount - IIf(Pagina > 1, (TBLISTA_PCPRM.PageSize * (Pagina - 1)), 0), TBLISTA_PCPRM.PageSize)
PBLista.Value = 1
contador = 0
Do While TBLISTA_PCPRM.EOF = False And (ContadorReg <= TamanhoPagina)
    With Lista.ListItems
        .Add , , TBLISTA_PCPRM!ID
        .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA_PCPRM!Desenho), "", TBLISTA_PCPRM!Desenho)
        .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA_PCPRM!Descricao), "", TBLISTA_PCPRM!Descricao)
        .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA_PCPRM!Requisitado), "0,0000", Format(TBLISTA_PCPRM!Requisitado, "###,##0.0000"))
        .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA_PCPRM!Unidade), "", TBLISTA_PCPRM!Unidade)
        .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA_PCPRM!Requisitado_PC), 0, TBLISTA_PCPRM!Requisitado_PC)
        .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA_PCPRM!Ordem_req), "", TBLISTA_PCPRM!Ordem_req)
        .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA_PCPRM!Prazo), "", Format(TBLISTA_PCPRM!Prazo, "dd/mm/yy"))
        .Item(.Count).SubItems(8) = IIf(IsNull(TBLISTA_PCPRM!Desenho_doc_req), "", TBLISTA_PCPRM!Desenho_doc_req)
        .Item(.Count).SubItems(9) = IIf(IsNull(TBLISTA_PCPRM!Revitem_doc_req), "", TBLISTA_PCPRM!Revitem_doc_req)
        .Item(.Count).SubItems(10) = IIf(IsNull(TBLISTA_PCPRM!Ref_doc_req), "", TBLISTA_PCPRM!Ref_doc_req)
        .Item(.Count).SubItems(11) = IIf(IsNull(TBLISTA_PCPRM!Descricao_doc_req), "", TBLISTA_PCPRM!Descricao_doc_req)
        .Item(.Count).SubItems(12) = IIf(IsNull(TBLISTA_PCPRM!Cliente_doc_req), "", TBLISTA_PCPRM!Cliente_doc_req)
        .Item(.Count).SubItems(13) = IIf(IsNull(TBLISTA_PCPRM!ID_empresa), "", TBLISTA_PCPRM!ID_empresa)
    End With
    TBLISTA_PCPRM.MoveNext
    ContadorReg = ContadorReg + 1
    contador = contador + 1
    PBLista.Value = contador
Loop
lblRegistros.Caption = "Nº de registros: " & TBLISTA_PCPRM.RecordCount
If TBLISTA_PCPRM.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "Página: 1 de: " & TBLISTA_PCPRM.PageCount
ElseIf TBLISTA_PCPRM.AbsolutePage = adPosEOF Then
        lblPaginas.Caption = "Página: " & TBLISTA_PCPRM.PageCount & " de: " & TBLISTA_PCPRM.PageCount
    Else
        lblPaginas.Caption = "Página: " & TBLISTA_PCPRM.AbsolutePage - 1 & " de: " & TBLISTA_PCPRM.PageCount
End If


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimparCampos()
On Error GoTo tratar_erro

Lista.ListItems.Clear
txtOrdem = ""
txtdesenho = ""
txtdescricao = ""
txtUN = ""
txtUn_com = ""
txtQtde = ""
txtQtde_PC = ""
txtDesenho_Similar = ""
txtDesc_Similar = ""
txtUn_Similar = ""
txtUN_com_Similar = ""
txtQtde_Similar = ""
txtQtde_PC_Similar = ""
Frame1.Enabled = False
Frame6.Enabled = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcFiltrar
    Case 2: ProcSalvar
    'Case 4: ProcAjuda
    Case 5: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvar()
On Error GoTo tratar_erro

If txtOrdem = "" Then Exit Sub

Acao = "salvar"
If txtDesenho_Similar = "" Then
    NomeCampo = "o código interno"
    ProcVerificaAcao
    cmdDesenho_Click
    Exit Sub
End If
If txtQtde_Similar.Locked = False Then
    Total = IIf(txtQtde_Similar = "", 0, txtQtde_Similar)
    If Total <= 0 Then
        NomeCampo = "a quantidade"
        ProcVerificaAcao
        txtQtde_Similar.SetFocus
        Exit Sub
    End If
Else
    Total = IIf(txtQtde_PC_Similar = "", 0, txtQtde_PC_Similar)
    If Total <= 0 Then
        NomeCampo = "a quantidade de peças"
        ProcVerificaAcao
        txtQtde_PC_Similar.SetFocus
        Exit Sub
    End If
End If
If FunAlterarProdSimiliarOrdem(Lista.SelectedItem.ListSubItems(13), txtDesenho_Similar, txtOrdem, txtdesenho, IIf(txtQtde_Similar = "", 0, txtQtde_Similar), IIf(txtQtde_PC_Similar = "", 0, txtQtde_PC_Similar), False) = True Then
    USMsgBox ("Requisição de material da ordem alterada com sucesso."), vbInformation, "CAPRIND v5.0"
    '==================================
    Modulo = Formulario
    Evento = "Alterar requisição de material da ordem"
    ID_documento = Txt_ID
    Documento = "Ordem:" & txtOrdem & " - Cód. interno de: " & txtdesenho
    Documento1 = "Ordem:" & txtOrdem & " - Cód. interno para: " & txtDesenho_Similar
    ProcGravaEvento
    '==================================
    ProcCarregaLista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
Else
    USMsgBox ("Não foi alterado a requisição do material desta ordem, pois não está ativo o recurso para produtos similares no cadastro da empresa."), vbExclamation, "CAPRIND v5.0"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
