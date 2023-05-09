VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form frmFinanceiro_Relatorios 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Administrativo - Financeiro - Relatórios - Histórico"
   ClientHeight    =   10035
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15270
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10035
   ScaleWidth      =   15270
   WindowState     =   2  'Maximized
   Begin VB.Frame frm_Receberxpagar 
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
      Height          =   825
      Left            =   60
      TabIndex        =   62
      Top             =   8850
      Visible         =   0   'False
      Width           =   15195
      Begin VB.TextBox txtSaldo1 
         Alignment       =   1  'Right Justify
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
         Left            =   13005
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   77
         TabStop         =   0   'False
         ToolTipText     =   "Total a receber."
         Top             =   390
         Width           =   2010
      End
      Begin VB.TextBox Txt_pagar1 
         Alignment       =   1  'Right Justify
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
         Left            =   10965
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   66
         TabStop         =   0   'False
         ToolTipText     =   "Total recebido."
         Top             =   390
         Width           =   2010
      End
      Begin VB.TextBox Txt_receber1 
         Alignment       =   1  'Right Justify
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
         Left            =   4890
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   65
         TabStop         =   0   'False
         ToolTipText     =   "Total a receber."
         Top             =   390
         Width           =   2010
      End
      Begin VB.TextBox txtDescontado2 
         Alignment       =   1  'Right Justify
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
         Left            =   6915
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   64
         TabStop         =   0   'False
         ToolTipText     =   "Total descontado."
         Top             =   390
         Width           =   2010
      End
      Begin VB.TextBox txtTotal_receber2 
         Alignment       =   1  'Right Justify
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
         Left            =   8940
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   63
         TabStop         =   0   'False
         ToolTipText     =   "Total geral."
         Top             =   390
         Width           =   2010
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo"
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
         Left            =   13778
         TabIndex        =   79
         Top             =   180
         Width           =   465
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total a pagar"
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
         Left            =   11408
         TabIndex        =   70
         Top             =   180
         Width           =   1125
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total geral a receber"
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
         Left            =   9060
         TabIndex        =   69
         Top             =   180
         Width           =   1770
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total a receber"
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
         Left            =   5250
         TabIndex        =   68
         Top             =   180
         Width           =   1290
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total descontado"
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
         Left            =   7170
         TabIndex        =   67
         Top             =   180
         Width           =   1500
      End
   End
   Begin VB.Frame Frm_receberxrecebido 
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
      Height          =   825
      Left            =   60
      TabIndex        =   43
      Top             =   8850
      Visible         =   0   'False
      Width           =   15195
      Begin VB.TextBox txtTotal_receber1 
         Alignment       =   1  'Right Justify
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
         Left            =   10980
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   29
         TabStop         =   0   'False
         ToolTipText     =   "Total geral."
         Top             =   390
         Width           =   2010
      End
      Begin VB.TextBox txtDescontado1 
         Alignment       =   1  'Right Justify
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
         Left            =   8950
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   28
         TabStop         =   0   'False
         ToolTipText     =   "Total descontado."
         Top             =   390
         Width           =   2010
      End
      Begin VB.TextBox Txt_receber 
         Alignment       =   1  'Right Justify
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
         Left            =   6935
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   27
         TabStop         =   0   'False
         ToolTipText     =   "Total a receber."
         Top             =   390
         Width           =   2010
      End
      Begin VB.TextBox Txt_recebido 
         Alignment       =   1  'Right Justify
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
         Left            =   13005
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   30
         TabStop         =   0   'False
         ToolTipText     =   "Total recebido."
         Top             =   390
         Width           =   2010
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total descontado"
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
         Left            =   9205
         TabIndex        =   56
         Top             =   180
         Width           =   1500
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total a receber"
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
         Left            =   7295
         TabIndex        =   45
         Top             =   180
         Width           =   1290
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total geral a receber"
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
         Left            =   11100
         TabIndex        =   57
         Top             =   180
         Width           =   1770
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total recebido"
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
         Left            =   13403
         TabIndex        =   44
         Top             =   180
         Width           =   1215
      End
   End
   Begin VB.Frame Frm_pagarxpago 
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
      Height          =   825
      Left            =   60
      TabIndex        =   59
      Top             =   8850
      Visible         =   0   'False
      Width           =   15195
      Begin VB.TextBox Txt_pago 
         Alignment       =   1  'Right Justify
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
         Left            =   13005
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   26
         TabStop         =   0   'False
         ToolTipText     =   "Total pago."
         Top             =   390
         Width           =   2010
      End
      Begin VB.TextBox Txt_pagar 
         Alignment       =   1  'Right Justify
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
         Left            =   10980
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   25
         TabStop         =   0   'False
         ToolTipText     =   "Total a pagar."
         Top             =   390
         Width           =   2010
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total pago"
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
         Left            =   13560
         TabIndex        =   61
         Top             =   180
         Width           =   900
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total a pagar"
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
         Left            =   11423
         TabIndex        =   60
         Top             =   180
         Width           =   1125
      End
   End
   Begin VB.Frame Frm_pagar_pago_recebido 
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
      Height          =   825
      Left            =   55
      TabIndex        =   36
      Top             =   8850
      Width           =   15195
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
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
         Left            =   13005
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   "Total geral."
         Top             =   390
         Width           =   2010
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total geral"
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
         Left            =   13545
         TabIndex        =   37
         Top             =   180
         Width           =   930
      End
   End
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
   Begin VB.ComboBox Cmb_empresa 
      BackColor       =   &H00FFFFFF&
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
      Height          =   315
      ItemData        =   "frmFinanceiro_Relatorios.frx":0000
      Left            =   1140
      List            =   "frmFinanceiro_Relatorios.frx":0002
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   54
      ToolTipText     =   "Empresa."
      Top             =   1110
      Width           =   10530
   End
   Begin VB.ComboBox Cmb_tipo 
      BackColor       =   &H00FFFFFF&
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
      Height          =   315
      ItemData        =   "frmFinanceiro_Relatorios.frx":0004
      Left            =   12300
      List            =   "frmFinanceiro_Relatorios.frx":0020
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   53
      ToolTipText     =   "Opções para filtro."
      Top             =   1110
      Width           =   2775
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Height          =   1325
      Left            =   13140
      TabIndex        =   48
      Top             =   1470
      Width           =   2115
      Begin VB.ComboBox cmbPor 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         ItemData        =   "frmFinanceiro_Relatorios.frx":009A
         Left            =   630
         List            =   "frmFinanceiro_Relatorios.frx":00A1
         Style           =   2  'Dropdown List
         TabIndex        =   10
         ToolTipText     =   "Por."
         Top             =   210
         Width           =   1305
      End
      Begin VB.ComboBox Cmb_mes_de 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         ItemData        =   "frmFinanceiro_Relatorios.frx":00AA
         Left            =   630
         List            =   "frmFinanceiro_Relatorios.frx":00D2
         Style           =   2  'Dropdown List
         TabIndex        =   13
         ToolTipText     =   "Mês de."
         Top             =   570
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.ComboBox Cmb_ano_de 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         ItemData        =   "frmFinanceiro_Relatorios.frx":0113
         Left            =   1260
         List            =   "frmFinanceiro_Relatorios.frx":0115
         Style           =   2  'Dropdown List
         TabIndex        =   14
         ToolTipText     =   "Ano de."
         Top             =   570
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.ComboBox Cmb_mes_ate 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         ItemData        =   "frmFinanceiro_Relatorios.frx":0117
         Left            =   630
         List            =   "frmFinanceiro_Relatorios.frx":013F
         Style           =   2  'Dropdown List
         TabIndex        =   15
         ToolTipText     =   "Mês até."
         Top             =   930
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.ComboBox Cmb_ano_ate 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         ItemData        =   "frmFinanceiro_Relatorios.frx":0180
         Left            =   1260
         List            =   "frmFinanceiro_Relatorios.frx":0182
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   16
         ToolTipText     =   "Ano até."
         Top             =   930
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.ComboBox Cmb_ano_de1 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         ItemData        =   "frmFinanceiro_Relatorios.frx":0184
         Left            =   630
         List            =   "frmFinanceiro_Relatorios.frx":0186
         Style           =   2  'Dropdown List
         TabIndex        =   17
         ToolTipText     =   "Ano de."
         Top             =   570
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.ComboBox Cmb_ano_ate1 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         ItemData        =   "frmFinanceiro_Relatorios.frx":0188
         Left            =   630
         List            =   "frmFinanceiro_Relatorios.frx":018A
         Style           =   2  'Dropdown List
         TabIndex        =   18
         ToolTipText     =   "Ano até."
         Top             =   930
         Visible         =   0   'False
         Width           =   1305
      End
      Begin MSComCtl2.DTPicker msk_fltInicio 
         Height          =   315
         Left            =   630
         TabIndex        =   11
         ToolTipText     =   "Data inicio."
         Top             =   570
         Width           =   1305
         _ExtentX        =   2302
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
         Format          =   133496833
         CurrentDate     =   39799
      End
      Begin MSComCtl2.DTPicker msk_fltFim 
         Height          =   315
         Left            =   630
         TabIndex        =   12
         ToolTipText     =   "Data final."
         Top             =   930
         Width           =   1305
         _ExtentX        =   2302
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
         Format          =   133496833
         CurrentDate     =   39799
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
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
         Left            =   180
         TabIndex        =   51
         Top             =   990
         Width           =   360
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
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
         Left            =   240
         TabIndex        =   50
         Top             =   630
         Width           =   300
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Por :"
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
         TabIndex        =   49
         Top             =   278
         Width           =   345
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   52
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
      RightColor1     =   0
      RightColor2     =   0
      ShowEndPanel    =   0   'False
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
      ButtonCaption2  =   "Relatório"
      ButtonEnabled2  =   0   'False
      ButtonIconSize2 =   32
      ButtonToolTipText2=   "Relatório (F5)"
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
      ButtonWidth2    =   51
      ButtonHeight2   =   21
      ButtonUseMaskColor2=   0   'False
      ButtonEnabled3  =   0   'False
      ButtonIconSize3 =   32
      ButtonAlignment3=   2
      ButtonType3     =   1
      ButtonStyle3    =   -1
      BeginProperty ButtonFont3 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState3    =   -1
      ButtonLeft3     =   93
      ButtonTop3      =   4
      ButtonWidth3    =   2
      ButtonHeight3   =   54
      ButtonUseMaskColor3=   0   'False
      ButtonCaption4  =   "Ajuda"
      ButtonEnabled4  =   0   'False
      ButtonIconSize4 =   32
      ButtonToolTipText4=   "Ajuda (F1)"
      ButtonKey4      =   "3"
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
      ButtonLeft4     =   97
      ButtonTop4      =   2
      ButtonWidth4    =   36
      ButtonHeight4   =   21
      ButtonCaption5  =   "Sair"
      ButtonEnabled5  =   0   'False
      ButtonIconSize5 =   32
      ButtonToolTipText5=   "Sair (Esc)"
      ButtonKey5      =   "4"
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
      ButtonLeft5     =   135
      ButtonTop5      =   2
      ButtonWidth5    =   26
      ButtonHeight5   =   21
      ButtonEnabled6  =   0   'False
      ButtonIconSize6 =   32
      BeginProperty ButtonFont6 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState6    =   5
      ButtonLeft6     =   163
      ButtonTop6      =   2
      ButtonWidth6    =   24
      ButtonHeight6   =   24
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   10740
         Top             =   60
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmFinanceiro_Relatorios.frx":018C
         Count           =   1
      End
   End
   Begin VB.Frame Frame4 
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
      Height          =   1325
      Left            =   3240
      TabIndex        =   31
      Top             =   1470
      Width           =   9885
      Begin VB.CheckBox Chk_mostrar_todasCC 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Mostrar todas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   8190
         TabIndex        =   9
         Top             =   960
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.OptionButton optEmisao 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Emissão"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   1800
         TabIndex        =   6
         Top             =   930
         Value           =   -1  'True
         Width           =   1005
      End
      Begin VB.OptionButton optPgto_receb 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Baixa"
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
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   4530
         TabIndex        =   8
         Top             =   930
         Width           =   1485
      End
      Begin VB.OptionButton optVencimento 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Vencimento"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   3000
         TabIndex        =   7
         Top             =   930
         Width           =   1305
      End
      Begin VB.ComboBox cmbTexto 
         BackColor       =   &H00FFFFFF&
         Height          =   330
         ItemData        =   "frmFinanceiro_Relatorios.frx":2F80
         Left            =   2790
         List            =   "frmFinanceiro_Relatorios.frx":2F82
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         ToolTipText     =   "Texto para pesquisa."
         Top             =   480
         Width           =   6915
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
         ItemData        =   "frmFinanceiro_Relatorios.frx":2F84
         Left            =   180
         List            =   "frmFinanceiro_Relatorios.frx":2FA0
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         ToolTipText     =   "Opções para filtro."
         Top             =   480
         Width           =   2595
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pesquisar por :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   420
         TabIndex        =   42
         Top             =   930
         Width           =   1080
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filtrar por"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   1050
         TabIndex        =   33
         Top             =   270
         Width           =   705
      End
      Begin VB.Label Label9 
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
         Left            =   5512
         TabIndex        =   32
         Top             =   270
         Width           =   1470
      End
   End
   Begin VB.Frame Frame5 
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
      Height          =   1325
      Left            =   1770
      TabIndex        =   34
      Top             =   1470
      Width           =   1455
      Begin VB.OptionButton optDetalhado 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Detalhado"
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
         Left            =   180
         TabIndex        =   2
         Top             =   480
         Value           =   -1  'True
         Width           =   1185
      End
      Begin VB.OptionButton optResumido 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Resumido"
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
         Left            =   180
         TabIndex        =   3
         Top             =   750
         Width           =   1155
      End
   End
   Begin VB.Frame Frame7 
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
      Height          =   1325
      Left            =   60
      TabIndex        =   35
      Top             =   1470
      Width           =   1695
      Begin VB.OptionButton Opt_individual 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Individual"
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
         Left            =   180
         TabIndex        =   0
         ToolTipText     =   "0"
         Top             =   480
         Value           =   -1  'True
         Width           =   1185
      End
      Begin VB.OptionButton Opt_comparativo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Comparativo"
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
         Left            =   180
         TabIndex        =   1
         Top             =   750
         Width           =   1425
      End
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   60
      TabIndex        =   58
      Top             =   9720
      Width           =   11385
      _ExtentX        =   20082
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
   Begin VB.Frame Frm_receber 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   60
      TabIndex        =   38
      Top             =   8850
      Visible         =   0   'False
      Width           =   15195
      Begin VB.TextBox txtReceber 
         Alignment       =   1  'Right Justify
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
         Left            =   8960
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   "Total a receber."
         Top             =   390
         Width           =   2010
      End
      Begin VB.TextBox txtDescontado 
         Alignment       =   1  'Right Justify
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
         Left            =   10980
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   23
         TabStop         =   0   'False
         ToolTipText     =   "Total descontado."
         Top             =   390
         Width           =   2010
      End
      Begin VB.TextBox txtTotal_receber 
         Alignment       =   1  'Right Justify
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
         Left            =   13005
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   24
         TabStop         =   0   'False
         ToolTipText     =   "Total geral."
         Top             =   390
         Width           =   2010
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total a receber"
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
         Left            =   9305
         TabIndex        =   41
         Top             =   180
         Width           =   1320
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total descontado"
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
         Left            =   11235
         TabIndex        =   40
         Top             =   180
         Width           =   1500
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total geral"
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
         Left            =   13545
         TabIndex        =   39
         Top             =   180
         Width           =   930
      End
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   6045
      Left            =   60
      TabIndex        =   19
      Top             =   2805
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   10663
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   12
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Object.Tag             =   "D"
         Text            =   "Dt. emissão"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Object.Tag             =   "D"
         Text            =   "Dt. vencto."
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Object.Tag             =   "N"
         Text            =   "Valor"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Nº documento"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Tag             =   "T"
         Text            =   "Nota fiscal"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Object.Tag             =   "T"
         Text            =   "Parcela"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Object.Tag             =   "T"
         Text            =   "Razão social"
         Object.Width           =   6800
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Object.Tag             =   "N"
         Text            =   "Vlr. baixa"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   9
         Object.Tag             =   "D"
         Text            =   "Dt. baixa"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Object.Tag             =   "T"
         Text            =   "Doc. baixa"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   11
         Object.Tag             =   "T"
         Text            =   "Tipo"
         Object.Width           =   1587
      EndProperty
   End
   Begin MSComctlLib.ListView Lista1 
      Height          =   6045
      Left            =   60
      TabIndex        =   20
      Top             =   2805
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   10663
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Frame frm_Recebidoxpago 
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
      Height          =   825
      Left            =   60
      TabIndex        =   71
      Top             =   8850
      Visible         =   0   'False
      Width           =   15195
      Begin VB.TextBox txtSaldo 
         Alignment       =   1  'Right Justify
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
         Left            =   13005
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   76
         TabStop         =   0   'False
         ToolTipText     =   "Total a receber."
         Top             =   390
         Width           =   2010
      End
      Begin VB.TextBox Txt_recebido1 
         Alignment       =   1  'Right Justify
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
         Left            =   8940
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   73
         TabStop         =   0   'False
         ToolTipText     =   "Total a receber."
         Top             =   390
         Width           =   2010
      End
      Begin VB.TextBox Txt_pago1 
         Alignment       =   1  'Right Justify
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
         Left            =   10972
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   72
         TabStop         =   0   'False
         ToolTipText     =   "Total recebido."
         Top             =   390
         Width           =   2010
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo"
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
         Left            =   13778
         TabIndex        =   78
         Top             =   180
         Width           =   465
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total recebido"
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
         Left            =   9338
         TabIndex        =   75
         Top             =   180
         Width           =   1215
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total pago"
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
         Left            =   11527
         TabIndex        =   74
         Top             =   180
         Width           =   900
      End
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo :"
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
      Left            =   11760
      TabIndex        =   55
      Top             =   1110
      Width           =   405
   End
   Begin VB.Label Lbl_relatorio 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Registros encontrados: 0000 - 00:00:00"
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
      Left            =   11610
      TabIndex        =   47
      Top             =   9750
      Width           =   3315
   End
   Begin VB.Label Label44 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Empresa :"
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
      Left            =   240
      TabIndex        =   46
      Top             =   1110
      Width           =   720
   End
End
Attribute VB_Name = "frmFinanceiro_Relatorios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub ProcAjuda()
On Error GoTo tratar_erro

FunAbrirVideoWeb ("www.youtube.com/watch?v=eS72kdWIrB4&list=UUODGiDjQ-BCrxh0YXfJvoqg&index=22&feature=plcp")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Chk_mostrar_todasCC_Click()
On Error GoTo tratar_erro

ProcLimpaListaeCampos
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Cmb_ano_ate1_Click()
On Error GoTo tratar_erro

ProcLimpaListaeCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Cmb_ano_de_Click()
On Error GoTo tratar_erro

Cmb_ano_ate = Cmb_ano_de
ProcLimpaListaeCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Cmb_ano_de1_Click()
On Error GoTo tratar_erro

ProcLimpaListaeCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Cmb_empresa_Click()
On Error GoTo tratar_erro

ProcLimpaListaeCampos
ProcCarregaComboTexto

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Cmb_mes_ate_Click()
On Error GoTo tratar_erro

ProcLimpaListaeCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Cmb_mes_de_Click()
On Error GoTo tratar_erro

ProcLimpaListaeCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Cmb_tipo_Click()
On Error GoTo tratar_erro

ProcOrdenaTudo

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub cmbPor_Click()
On Error GoTo tratar_erro

ProcLimpaListaeCampos
ProcMostrarEsconderCombosData

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 And Lista1.ListItems.Count = 0 Then Exit Sub
frmFinanceiro_Relatorios_Menu_Impressao.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

ProcExcluirDadosProducaoRelatorios
ProcExcluirDadosProducaoRelatoriosTotal
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub cmbTexto_Click()
On Error GoTo tratar_erro

ProcLimpaListaeCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyF2: ProcFiltrar
    Case vbKeyF5: ProcImprimir
    Case vbKeyF1: ProcAjuda
    Case vbKeyEscape: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Sub ProcCarregaLista()
On Error GoTo tratar_erro

Select Case Cmb_tipo
    Case "A pagar": TextoColuna = "Vlr. total pagar"
    Case "A receber": TextoColuna = "Vlr. total receber"
    Case "A pagar e pagas": TextoColuna = "Vlr. total pagar  |  pago"
    Case "Pagas": TextoColuna = "Vlr. total pago"
    Case "Recebidas": TextoColuna = "Vlr. total recebido"
    Case "A receber e recebidas": TextoColuna = "Vlr. total receber  |  receb."
    Case "Recebidas e pagas": TextoColuna = "Vlr. total receb.  |  pagas"
    Case "A receber e a pagar": TextoColuna = "Vlr. total receber  |  pagar"
End Select
Familiatext = ""
Contador1 = 1
Posicao = 0
Lista.ListItems.Clear
Lista1.ListItems.Clear
If TBLISTA.EOF = False Then
    If optDetalhado.Value = True Then Posicao = TBLISTA.RecordCount
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    contador = 0
    Do While TBLISTA.EOF = False
        If optDetalhado.Value = True Then
            With Lista.ListItems
                .Add , , TBLISTA!ID
                .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Data), "", Format(TBLISTA!Data, "dd/mm/yy")) 'Data emissão
                .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Data6), "", Format(TBLISTA!Data6, "dd/mm/yy")) 'Data vencimento
                .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!qtdeOK), "", Format(TBLISTA!qtdeOK, "###,##0.00")) 'Valor a receber / pagar
                .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!Totalhsutil), "", TBLISTA!Totalhsutil) 'Número documento
                .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!Data2), "", TBLISTA!Data2) 'NF
                .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA!Data1), "", TBLISTA!Data1) 'Parcela
                .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA!DescEvento), "", TBLISTA!DescEvento) 'Cliente/Fornecedor
                .Item(.Count).SubItems(8) = IIf(IsNull(TBLISTA!qtdeNC), "", Format(TBLISTA!qtdeNC, "###,##0.00")) 'Valor recebido / pago
                .Item(.Count).SubItems(9) = IIf(IsNull(TBLISTA!Data5), "", Format(TBLISTA!Data5, "dd/mm/yy")) 'Data recbto. / pgto.
                .Item(.Count).SubItems(10) = IIf(IsNull(TBLISTA!Data4), "", TBLISTA!Data4) 'Numero docto. baixa
                If Cmb_tipo = "A receber e a pagar" Or Cmb_tipo = "Recebidas e pagas" Then .Item(.Count).SubItems(11) = IIf(IsNull(TBLISTA!Totalhsprev), "", IIf(TBLISTA!Totalhsprev = "R", "Receber", "Pagar")) 'R para receber e P para pagar
            End With
        Else
            If TBLISTA!maquina <> "" Then
                With Lista1.ListItems
                    Contador1 = 1
                    If cmbPor = "Dia" Then
                        Do While Lista1.ColumnHeaders(Contador1 + 1).Text <> Format(TBLISTA!Execucaoprev, "dd/mm/yy")
                            Contador1 = Contador1 + 1
                        Loop
                    ElseIf cmbPor = "Mês" Then
                            Do While Lista1.ColumnHeaders(Contador1 + 1).Text <> TBLISTA!Execucaoprev
                                Contador1 = Contador1 + 1
                            Loop
                        Else
                            Do While Lista1.ColumnHeaders(Contador1 + 1).Text <> TBLISTA!Execucaoprev
                                Contador1 = Contador1 + 1
                            Loop
                    End If
Pula:
                    If TBLISTA!maquina <> Familiatext Then
                        .Add , , TBLISTA!maquina
                        Posicao = Posicao + 1
                    End If
                    
                    If Cmb_tipo = "A pagar" Or Cmb_tipo = "A receber" Then
                        valor = IIf(IsNull(TBLISTA!qtdeOK), 0, TBLISTA!qtdeOK)
                        Valor1 = IIf(IsNull(TBLISTA!Qtdetotalprod), 0, TBLISTA!Qtdetotalprod)
                    ElseIf Cmb_tipo = "Pagas" Or Cmb_tipo = "Recebidas" Then
                            valor = IIf(IsNull(TBLISTA!qtdeNC), 0, TBLISTA!qtdeNC)
                            Valor1 = IIf(IsNull(TBLISTA!OS), 0, TBLISTA!OS)
                        Else
                            valor = IIf(IsNull(TBLISTA!qtdeOK), 0, TBLISTA!qtdeOK) 'A pagar / receber
                            Valor1 = IIf(IsNull(TBLISTA!qtdeNC), 0, TBLISTA!qtdeNC) 'Pago / recebido
                            Valor2 = IIf(IsNull(TBLISTA!Qtdetotalprod), 0, TBLISTA!Qtdetotalprod) 'Total a pagar / receber
                            Valor3 = IIf(IsNull(TBLISTA!OS), 0, TBLISTA!OS) 'Total pago / recebido
                    End If
                    
                    If Cmb_tipo = "A pagar" Or Cmb_tipo = "A receber" Or Cmb_tipo = "Pagas" Or Cmb_tipo = "Recebidas" Then
                        .Item(.Count).SubItems(Contador1) = Format(valor, "###,##0.00")
                    
                        'Carrega valor total
                        Contador1 = 1
                        Do While Lista1.ColumnHeaders(Contador1 + 1).Text <> TextoColuna
                            Contador1 = Contador1 + 1
                        Loop
                        
                        .Item(.Count).SubItems(Contador1) = Format(Valor1, "###,##0.00")
                    Else
                        .Item(.Count).SubItems(Contador1) = Format(valor, "###,##0.00") & "  |  " & Format(Valor1, "###,##0.00")
                    
                        'Carrega valor total
                        Contador1 = 1
                        Do While Lista1.ColumnHeaders(Contador1 + 1).Text <> TextoColuna
                            Contador1 = Contador1 + 1
                        Loop
                        
                        .Item(.Count).SubItems(Contador1) = Format(Valor2, "###,##0.00") & "  |  " & Format(Valor3, "###,##0.00")
                    End If
                End With
            End If
        End If
        Familiatext = IIf(IsNull(TBLISTA!maquina), "", TBLISTA!maquina)
        TBLISTA.MoveNext
        contador = contador + 1
        PBLista.Value = contador
    Loop
    If optDetalhado.Value = True Then Else
End If

valor = 0
Valor1 = 0
Valor2 = 0
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Producao_Relatorios_Total where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    valor = IIf(IsNull(TBLISTA!QtdeProduzida), 0, TBLISTA!QtdeProduzida) 'Total descontado
    Valor1 = IIf(IsNull(TBLISTA!QtdePrevista), 0, TBLISTA!QtdePrevista) 'Total receber / pagar
    Valor2 = IIf(IsNull(TBLISTA!qtdeNC), 0, TBLISTA!qtdeNC) 'Total recebido / pago
    
    'Pagar / Pago / Recebido
    If Cmb_tipo = "A pagar" Then
        txtTotal = Format(Valor1, "###,##0.00")
    ElseIf Cmb_tipo = "Pagas" Or Cmb_tipo = "Recebidas" Then
        txtTotal = Format(Valor2, "###,##0.00")
    End If
    
    'Receber
    If Cmb_tipo = "A receber" Then
        txtDescontado = Format(valor, "###,##0.00")
        txtReceber = Format(Valor1, "###,##0.00")
        txtTotal_receber = Format(valor + Valor1, "###,##0.00")
    End If
    
    'Pagar e pago / receber e recebido
    If Cmb_tipo = "A pagar e pagas" Then
        Txt_pagar = Format(Valor1, "###,##0.00")
        Txt_pago = Format(Valor2, "###,##0.00")
    ElseIf Cmb_tipo = "A receber e recebidas" Then
        Txt_receber = Format(Valor1, "###,##0.00")
        txtDescontado1 = Format(valor, "###,##0.00")
        txtTotal_receber1 = Format(Valor1 + valor, "###,##0.00")
        Txt_recebido = Format(Valor2, "###,##0.00")
    End If
    
    If Cmb_tipo = "Recebidas e pagas" Then
        Txt_recebido1 = Format(Valor1, "###,##0.00")
        Txt_pago1 = Format(Valor2, "###,##0.00")
        txtSaldo = Format(Valor1 - Valor2, "###,##0.00")
    ElseIf Cmb_tipo = "A receber e a pagar" Then
        Txt_receber1 = Format(Valor1, "###,##0.00")
        txtDescontado2 = Format(valor, "###,##0.00")
        txtTotal_receber2 = Format(Valor1 + valor, "###,##0.00")
        Txt_pagar1 = Format(Valor2, "###,##0.00")
        txtSaldo1 = Format(Valor1 - Valor2, "###,##0.00")
    End If
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcLimpaListaeCampos()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista1.ListItems.Clear
ProcLimpaCamposTotais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Sub ProcLimpaCamposTotais()
On Error GoTo tratar_erro

Lbl_relatorio.Caption = "Registros encontrados: 0000 - 00:00:00"

'Pagar / Pago / Recebido
txtTotal = ""

'Pagar x pagas
Txt_pagar = ""
Txt_pago = ""

'Receber
txtReceber = ""
txtDescontado = ""
txtTotal_receber = ""

'Receber x recebidas
Txt_receber = ""
txtDescontado1 = ""
txtTotal_receber1 = ""
Txt_recebido = ""

'Receber x pagar
Txt_receber1 = ""
txtDescontado2 = ""
txtTotal_receber2 = ""
Txt_pagar1 = ""
txtSaldo = ""

'Recebido x pago
Txt_recebido1 = ""
Txt_pago1 = ""
txtSaldo1 = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15195, 6, True

Formulario = "Financeiro/Relatórios/Histórico"
Direitos
ProcLimpaVariaveisPrincipais
ProcCarregaComboEmpresa Cmb_empresa, False
Cmb_tipo = "A pagar"
ProcOrdenaTudo
msk_fltInicio.Value = Date
msk_fltFim.Value = Date
Cmb_ano_de.Clear
Cmb_ano_de1.Clear
Cmb_ano_ate.Clear
Cmb_ano_ate1.Clear
AnoAtual = 2005
Do While AnoAtual <> (Year(Date) + 4)
    Cmb_ano_de.AddItem AnoAtual
    Cmb_ano_de1.AddItem AnoAtual
    Cmb_ano_ate.AddItem AnoAtual
    Cmb_ano_ate1.AddItem AnoAtual
    AnoAtual = AnoAtual + 1
Loop
Cmb_ano_ate = Year(Date)
Cmb_ano_ate1 = Year(Date)
Cmb_ano_de = Year(Date)
Cmb_ano_de1 = Year(Date)
cmbPor = "Dia"
With cmbfiltrarpor
    .Text = "Fornecedor"
    .Refresh
End With
cmbTexto.Refresh

ProcRemoveObjetosResize Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

Formulario = "Financeiro/Relatórios/Histórico"
Direitos
ProcLimpaVariaveisPrincipais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView Lista, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Lista1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView Lista1, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub msk_fltFim_Click()
On Error GoTo tratar_erro

ProcLimpaListaeCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub cmbfiltrarpor_Click()
On Error GoTo tratar_erro

ProcCarregaComboTexto

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcCarregaComboTexto()
On Error GoTo tratar_erro

Desenho = ""
cmbTexto.Clear
Texto = ""
Logsit = ""
ProcArrumaLista
Chk_mostrar_todasCC.Visible = False

If cmbfiltrarpor = "" Then Exit Sub

If cmbfiltrarpor = "Conta contábil" Then
    If Opt_comparativo.Value = True Then Chk_mostrar_todasCC.Visible = True
    
    If Cmb_tipo = "A receber" Or Cmb_tipo = "Recebidas" Or Cmb_tipo = "A receber e recebidas" Then
        NomeTabela = "tbl_contas_receber"
        Tipo = "AND familia_financeiro.tipoconta = 'R'"
    ElseIf Cmb_tipo = "Recebidas e pagas" Or Cmb_tipo = "A receber e a pagar" Then
        NomeTabela = "Financeiro_relatorios_historico_detalhado"
        Tipo = ""
    Else
        NomeTabela = "tbl_ContasPagar"
        Tipo = "AND familia_financeiro.tipoconta = 'P'"
    End If

    Select Case Cmb_tipo
        Case "A receber", "A pagar", "A receber e a pagar": Logsit = "familia_financeiro.Pago_recebido = 'False'"
        Case "Recebidas", "Pagas", "Recebidas e pagas": Logsit = "familia_financeiro.Pago_recebido = 'True'"
        Case "A receber e recebidas", "A pagar e pagas": Logsit = "(familia_financeiro.Pago_recebido = 'False' or familia_financeiro.Pago_recebido = 'True')"
    End Select

    Set TBLISTA = CreateObject("adodb.recordset")
    TBLISTA.Open "Select tbl_familia.int_codfamilia, tbl_familia.Codigo, tbl_familia.txt_descricao from (" & NomeTabela & " INNER JOIN familia_financeiro ON " & NomeTabela & ".IdIntConta = familia_financeiro.IDConta) INNER JOIN tbl_familia ON tbl_familia.int_codfamilia = familia_financeiro.ID_PC where Familia_financeiro.Deposito_transf <> 'True'" & Tipo & " and " & Logsit & " and " & NomeTabela & ".ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " Group by tbl_familia.int_codfamilia, tbl_familia.Codigo, tbl_familia.txt_descricao", Conexao, adOpenKeyset, adLockOptimistic
    If TBLISTA.EOF = False Then
        Do While TBLISTA.EOF = False
            cmbTexto.AddItem TBLISTA!CODIGO & " - " & TBLISTA!Txt_descricao
            cmbTexto.ItemData(cmbTexto.NewIndex) = TBLISTA!int_codfamilia
            TBLISTA.MoveNext
        Loop
    End If
    TBLISTA.Close
Else
    Select Case Cmb_tipo
        Case "A receber", "A pagar", "A receber e a pagar": Logsit = "LogSit = 'N'"
        Case "Recebidas", "Pagas", "Recebidas e pagas": Logsit = "LogSit = 'S'"
        Case "A receber e recebidas", "A pagar e pagas": Logsit = "(LogSit = 'S' or LogSit = 'N')"
    End Select

    If Cmb_tipo = "A receber" Or Cmb_tipo = "Recebidas" Or Cmb_tipo = "A receber e recebidas" Then
        If cmbfiltrarpor = "Status" Then
            With cmbTexto
                If Cmb_tipo = "A receber" Then
                    .AddItem "TÍTULO EM ABERTO"
                    .AddItem "TÍTULO RECEBIDO PARCIAL"
                    .AddItem "DUPLICATA DESCONTADA EM ABERTO"
                    .AddItem "BLOQUEADA"
                Else
                    .AddItem "TÍTULO LIQUIDADO"
                    .AddItem "TÍTULO RECEBIDO PARCIAL"
                    .AddItem "TÍTULO RECEBIDO PARCIAL LIQUIDADO"
                    .AddItem "DUPLICATA DESCONTADA LIQUIDADA"
                    .AddItem "DUPLICATA DESCONTADA RECOMPRADA"
                End If
            End With
        Else
            Select Case cmbfiltrarpor
                Case "Documento": TextoFiltro = "txt_NDocumento"
                Case "Nota fiscal": TextoFiltro = "Nfiscal"
                Case "Cliente": TextoFiltro = "Nome_Razao"
                Case "Tipo do doc.": TextoFiltro = "Tipo_doc"
                Case "Instituição": TextoFiltro = "Banco"
                Case "Documento baixa": TextoFiltro = "txt_NDocumento"
            End Select
            
            Set TBLISTA = CreateObject("adodb.recordset")
            TBLISTA.Open "Select " & TextoFiltro & " as NomeCampo from tbl_contas_receber where " & TextoFiltro & " IS NOT NULL and Bloqueado = 'False' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and " & Logsit & " Group by " & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
            If TBLISTA.EOF = False Then
                Do While TBLISTA.EOF = False
                    If TBLISTA!NomeCampo <> "" Then cmbTexto.AddItem TBLISTA!NomeCampo
                    TBLISTA.MoveNext
                Loop
            End If
            TBLISTA.Close
        End If
    ElseIf Cmb_tipo = "A receber e a pagar" Or Cmb_tipo = "Recebidas e pagas" Then
        Select Case cmbfiltrarpor
            Case "Documento": TextoFiltro = "txt_NDocumento"
            Case "Cliente/Fornecedor": TextoFiltro = "Razao"
            Case "Tipo do doc.": TextoFiltro = "Tipo_doc"
            Case "Instituição": TextoFiltro = "Banco"
            Case "Documento baixa": TextoFiltro = "NDoctoBaixa"
            Case "Forma de pagamento": TextoFiltro = "FormaBaixa"
        End Select
        
        Set TBLISTA = CreateObject("adodb.recordset")
        TBLISTA.Open "Select " & TextoFiltro & " as NomeCampo from Financeiro_relatorios_historico_detalhado where " & TextoFiltro & " IS NOT NULL and Bloqueado = 'False' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and " & Logsit & " Group by " & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
        If TBLISTA.EOF = False Then
            Do While TBLISTA.EOF = False
                If TBLISTA!NomeCampo <> "" Then cmbTexto.AddItem TBLISTA!NomeCampo
                TBLISTA.MoveNext
            Loop
        End If
        TBLISTA.Close
    Else
        If cmbfiltrarpor = "Status" Then
            With cmbTexto
                If Cmb_tipo = "A pagar e pagas" Then
                    .AddItem "TÍTULO LIQUIDADO"
                    .AddItem "TÍTULO PAGO PARCIAL"
                    .AddItem "TÍTULO PAGO PARCIAL LIQUIDADO"
                Else
                    .AddItem "TÍTULO EM ABERTO"
                    .AddItem "TÍTULO PAGO PARCIAL"
                    .AddItem "BLOQUEADA"
                End If
            End With
        Else
            Select Case cmbfiltrarpor
                Case "Documento": TextoFiltro = "txt_NDocumento"
                Case "Fornecedor": TextoFiltro = "txt_Fornecedor"
                Case "Tipo do doc.": TextoFiltro = "Class_conta"
                Case "Instituição": TextoFiltro = "Banco"
                Case "Documento baixa": TextoFiltro = "NDoctoBaixa"
                Case "Forma de pagamento": TextoFiltro = "FormaBaixa"
            End Select
            Set TBLISTA = CreateObject("adodb.recordset")
            TBLISTA.Open "Select (" & TextoFiltro & ") as NomeCampo from tbl_contaspagar where " & TextoFiltro & " <> 'Null' and Bloqueado = 'False' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and " & Logsit & " and " & TextoFiltro & " <> N'' Group by " & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
            If TBLISTA.EOF = False Then
                Do While TBLISTA.EOF = False
                    cmbTexto.AddItem TBLISTA!NomeCampo
                    TBLISTA.MoveNext
                Loop
            End If
            TBLISTA.Close
        End If
    End If
End If
ProcLimpaListaeCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

Acao = "filtrar"
If cmbfiltrarpor = "" Then
    NomeCampo = "o filtro para pesquisa"
    ProcVerificaAcao
    cmbfiltrarpor.SetFocus
    Exit Sub
End If
If Opt_individual.Value = True And cmbTexto = "" Then
    NomeCampo = "o texto para pesquisa"
    ProcVerificaAcao
    cmbTexto.SetFocus
    Exit Sub
End If

If cmbPor = "Dia" Then
    With msk_fltFim
        If FunVerificaDataFinal(msk_fltInicio.Value, .Value) = False Then
            .Value = Date
            .SetFocus
            Exit Sub
        End If
    End With
End If
If optResumido.Value = True Then
    If cmbPor = "Mês" Then
        If Cmb_mes_de = "" Then
            NomeCampo = "o mês"
            ProcVerificaAcao
            Cmb_mes_de.SetFocus
            Exit Sub
        End If
        If Cmb_mes_ate = "" Then
            NomeCampo = "o mês"
            ProcVerificaAcao
            Cmb_mes_ate.SetFocus
            Exit Sub
        End If
        If Cmb_ano_de = "" Then
            NomeCampo = "o ano"
            ProcVerificaAcao
            Cmb_ano_de.SetFocus
            Exit Sub
        End If
        qt = FunVerificaMes(Cmb_mes_de)
        Qtd = FunVerificaMes(Cmb_mes_ate)
        If Qtd < qt And Cmb_ano_ate <= Cmb_ano_de Then
            USMsgBox ("O mês final não pode ser menor que o mês inicial."), vbExclamation, "CAPRIND v5.0"
            Cmb_mes_ate = Cmb_mes_de
        End If
    ElseIf cmbPor = "Ano" Then
            If Cmb_ano_de1 = "" Then
                NomeCampo = "o ano"
                ProcVerificaAcao
                Cmb_ano_de1.SetFocus
                Exit Sub
            End If
            If Cmb_ano_ate1 = "" Then
                NomeCampo = "o ano"
                ProcVerificaAcao
                Cmb_ano_ate1.SetFocus
                Exit Sub
            End If
            Qtd = Cmb_ano_de1
            qt = Cmb_ano_ate1
            If qt < Qtd Then
                USMsgBox ("O ano final não pode ser menor que o ano inicial."), vbExclamation, "CAPRIND v5.0"
                Cmb_ano_ate1 = Cmb_ano_de1
            End If
    End If
End If

Inicio = Time
Desenho = ""
ProcLimpaCamposTotais
ProcAbrirTabelas
If optResumido.Value = True Then
    ProcCriaColunas
    
    'Soma e grava o total geral
    Set TBLISTA = CreateObject("adodb.recordset")
    TBLISTA.Open "Select maquina, Sum(QtdeOK) as Valor, Sum(QtdeNC) as Valor1 from Producao_relatorios where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "' Group by Maquina", Conexao, adOpenKeyset, adLockOptimistic
    If TBLISTA.EOF = False Then
        Do While TBLISTA.EOF = False
            valor = IIf(IsNull(TBLISTA!valor), 0, TBLISTA!valor) 'Valor receber / pagar
            Valor1 = IIf(IsNull(TBLISTA!Valor1), 0, TBLISTA!Valor1) 'Valor recebido / pago
            NovoValor = Replace(valor, ",", ".")
            NovoValor1 = Replace(Valor1, ",", ".")
            Conexao.Execute "Update Producao_relatorios Set Qtdetotalprod = " & NovoValor & ", OS = " & NovoValor1 & " where Maquina = '" & TBLISTA!maquina & "'"
            TBLISTA.MoveNext
        Loop
    End If
    TBLISTA.Close
End If
If Permitido = True Then ProcGravarTotalizacoes

Set TBLISTA = CreateObject("adodb.recordset")
If optDetalhado.Value = True And Opt_individual.Value = True Then
    TBLISTA.Open "Select * from Producao_Relatorios where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "' order by Maquina", Conexao, adOpenKeyset, adLockReadOnly
Else
    TBLISTA.Open "Select * from Producao_Relatorios where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "' order by Maquina", Conexao, adOpenKeyset, adLockReadOnly
End If
ProcCarregaLista

intervalo = Time
ElapsedTime (intervalo - Inicio)
Lbl_relatorio.Caption = "Registros encontrados: " & FunTamanhoTextoZeroEsq(Posicao, 4) & " - " & HoraTotal

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcAbrirTabelas()
On Error GoTo tratar_erro

'Deleta registros e adiciona novos
ProcExcluirDadosProducaoRelatorios
ProcExcluirDadosProducaoRelatoriosTotal
Conexao.Execute "DELETE from Plano_de_contas_totalizacao where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'"

Select Case Cmb_tipo
    Case "A receber", "A pagar", "A receber e a pagar": Logsit = "LogSit = 'N'"
    Case "Recebidas", "Pagas", "Recebidas e pagas": Logsit = "LogSit = 'S'"
    Case "A receber e recebidas", "A pagar e pagas": Logsit = "(LogSit = 'N' or LogSit = 'S')"
End Select

If optEmisao.Value = True Then
    If Cmb_tipo = "A pagar" Or Cmb_tipo = "Pagas" Or Cmb_tipo = "A pagar e pagas" Then
        Data_pesquisa = "dt_Emissao"
    Else
        Data_pesquisa = "Emissao"
    End If
ElseIf optVencimento.Value = True Then
    If Cmb_tipo = "A pagar" Or Cmb_tipo = "Pagas" Or Cmb_tipo = "A pagar e pagas" Then
        Data_pesquisa = "dt_Pagamento"
    Else
        Data_pesquisa = "Vencimento"
    End If
Else
    If Cmb_tipo = "A pagar" Or Cmb_tipo = "Pagas" Or Cmb_tipo = "A pagar e pagas" Then
        Data_pesquisa = "DataBaixa"
    Else
        Data_pesquisa = "Data_pagamento"
    End If
End If

If Cmb_tipo = "A receber" Or Cmb_tipo = "Recebidas" Or Cmb_tipo = "A receber e recebidas" Then
    TabelaFiltro = "Financeiro_relatorios_historico_detalhado_receber"
    TipoConta = " and FF.tipoconta = 'R'"
ElseIf Cmb_tipo = "A receber e a pagar" Or Cmb_tipo = "Recebidas e pagas" Then
    TabelaFiltro = "Financeiro_relatorios_historico_detalhado"
    TipoConta = ""
Else
    TabelaFiltro = "Financeiro_relatorios_historico_detalhado_pagar"
    TipoConta = " and FF.tipoconta = 'P'"
End If

Select Case cmbfiltrarpor
    Case "Documento": FamiliaAntiga = "txt_NDocumento"
    Case "Nota fiscal": FamiliaAntiga = "Nfiscal"
    Case "Cliente": FamiliaAntiga = "Nome_Razao"
    Case "Fornecedor": FamiliaAntiga = "txt_Fornecedor"
    Case "Cliente/Fornecedor": FamiliaAntiga = "Razao"
    Case "Tipo do doc.":
        Select Case Cmb_tipo
            Case "A receber", "Recebidas", "A receber e recebidas", "A receber e a pagar", "Recebidas e pagas": FamiliaAntiga = "Tipo_doc"
            Case "A pagar", "Pagas", "A pagar e pagas": FamiliaAntiga = "Class_conta"
        End Select
    Case "Status": FamiliaAntiga = "status"
    Case "Instituição": FamiliaAntiga = "Banco"
    Case "Documento baixa":
        Select Case Cmb_tipo
            Case "A receber", "Recebidas", "A receber e recebidas": FamiliaAntiga = "txt_NDocumento"
            Case "A pagar", "Pagas", "A pagar e pagas", "Recebidas e pagas": FamiliaAntiga = "NDoctoBaixa"
        End Select
    Case "Forma de pagamento": FamiliaAntiga = "FormaBaixa"
End Select

If Cmb_tipo = "A receber e a pagar" Or Cmb_tipo = "Recebidas e pagas" Then onamais = "and F.Tipo = FF.tipoconta" Else onamais = ""

Set TBCarteira = CreateObject("adodb.recordset")
If optDetalhado.Value = True Then
    If cmbfiltrarpor = "Conta contábil" Then
        TBCarteira.Open "Select F.*, FF.valor as quantidade, FF.Pago_recebido from " & TabelaFiltro & " F inner join familia_financeiro FF on F.IdIntConta = FF.idconta " & onamais & " WHERE FF.ID_PC = " & cmbTexto.ItemData(cmbTexto.ListIndex) & TipoConta & " and FF.Deposito_transf <> 'True' and F.bloqueado = 'False' and " & Data_pesquisa & " Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "' and F.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and " & Logsit & " order by FF.ID_PC", Conexao, adOpenKeyset, adLockReadOnly
    Else
        TBCarteira.Open "Select * FROM " & TabelaFiltro & " where " & FamiliaAntiga & " = '" & cmbTexto & "' and Bloqueado = 'False' and " & Data_pesquisa & " Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and " & Logsit & " order by " & Data_pesquisa & ", " & FamiliaAntiga, Conexao, adOpenKeyset, adLockReadOnly
    End If
Else
    If optEmisao.Value = True Then
        Data_pesquisa1 = "Mes_emissao"
        Data_pesquisa2 = "Ano_emissao"
    ElseIf optVencimento.Value = True Then
        Data_pesquisa1 = "Mes_venc"
        Data_pesquisa2 = "Ano_venc"
    Else
        Data_pesquisa1 = "Mes_pag"
        Data_pesquisa2 = "Ano_pag"
    End If
    
    If cmbfiltrarpor = "Conta contábil" Then
        Select Case cmbPor
            Case "Dia": DataFiltro = Data_pesquisa & " Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "'"
            Case "Mês":
                qt = FunVerificaMes(Cmb_mes_de)
                Qtd = FunVerificaMes(Cmb_mes_ate)
                MesX = qt
                MesX1 = Qtd
                DataFiltro = Data_pesquisa1 & " >= " & qt & " and " & Data_pesquisa2 & " >= " & Cmb_ano_de & " and " & Data_pesquisa1 & " <= " & Qtd & " and " & Data_pesquisa2 & " <= " & Cmb_ano_ate
            Case "Ano": DataFiltro = Data_pesquisa2 & " >= " & Cmb_ano_de1 & " and " & Data_pesquisa2 & " <= " & Cmb_ano_ate1
        End Select

        Set TBFamilia = CreateObject("adodb.recordset")
        If Opt_individual = True Then
            TBFamilia.Open "Select F.*, FF.ID_PC, FF.valor, FF.Pago_recebido, FF.TipoConta from " & TabelaFiltro & " F inner join familia_financeiro FF on F.IdIntConta = FF.idconta " & onamais & " WHERE FF.ID_PC = " & cmbTexto.ItemData(cmbTexto.ListIndex) & TipoConta & " and FF.Deposito_transf <> 'True' and F.bloqueado = 'False' and " & DataFiltro & " and F.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and " & Logsit & " Order by FF.ID_PC, FF.Pago_recebido", Conexao, adOpenKeyset, adLockReadOnly
        Else
            TBFamilia.Open "Select F.*, FF.ID_PC, FF.valor, FF.Pago_recebido, FF.TipoConta from " & TabelaFiltro & " F inner join familia_financeiro FF on F.IdIntConta = FF.idconta " & onamais & " WHERE FF.Deposito_transf <> 'True' " & TipoConta & " and F.bloqueado = 'False' and " & DataFiltro & " and F.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and " & Logsit & " Order by FF.ID_PC, FF.Pago_recebido", Conexao, adOpenKeyset, adLockReadOnly
        End If
        ProcVerifPC
        TBFamilia.Close
        
        filtrovalor = "valor"
        filtrovalor2 = "Valor_pgto_receb"
        
        'Tem lá em cima a verificação das datas, mas preciso ver de novo porque quando é resumido ele não pega da mesma tabela
        If optEmisao.Value = True Then
            Data_pesquisa = "Emissao"
        ElseIf optVencimento.Value = True Then
            Data_pesquisa = "Vencimento"
        Else
            Data_pesquisa = "Pagamento_receb"
            Data_pesquisa1 = "Mes_pgto_receb"
            Data_pesquisa2 = "Ano_pgto_receb"
        End If
    Else
        If Cmb_tipo = "Pagas" Or Cmb_tipo = "A pagar" Or Cmb_tipo = "A pagar e pagas" Then
            If Cmb_tipo = "Pagas" Then
                filtrovalor = "Valorpago"
            Else
                filtrovalor = "dbl_valorpagto"
            End If
            filtrovalor2 = "Valorpago"
        Else
            If Cmb_tipo = "Recebidas" Then
                filtrovalor = "valortitulorecebido"
            ElseIf Cmb_tipo = "Recebidas e pagas" Then
                filtrovalor = "vlrBaixa"
            Else
                filtrovalor = "valor"
            End If
            
            If Cmb_tipo = "Recebidas e pagas" Then
                filtrovalor2 = "vlrBaixa"
            ElseIf Cmb_tipo = "A receber e a pagar" Then
                filtrovalor2 = "valor"
            Else
                filtrovalor2 = "valortitulorecebido"
            End If
        End If
    End If
        
    Par1 = ""
    Permitido = False
    Select Case cmbPor
        Case "Dia":
            Dataini = msk_fltInicio
            DataFim = msk_fltFim
            Do While Dataini <= DataFim
                If Permitido = False Then Par1 = "[" & Dataini & "]" Else Par1 = Par1 & " , [" & Dataini & "]"
                Permitido = True
                Dataini = Dataini + 1
            Loop
            Pesquisa = "(" & Data_pesquisa & ") Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "'"
            Pesquisa1 = "PIVOT (Sum(" & filtrovalor & ") for " & Data_pesquisa & " In (" & Par1 & "))"
            Pesquisa3 = "PIVOT (Sum(" & filtrovalor2 & ") for " & Data_pesquisa & " In (" & Par1 & "))"
            Pesquisa2 = Data_pesquisa
        Case "Mês":
            qt = FunVerificaMes(Cmb_mes_de)
            Qtd = FunVerificaMes(Cmb_mes_ate)
            MesX = qt
            MesX1 = Qtd
            Do While qt <= Qtd
                If Permitido = False Then Par1 = "[" & qt & "]" Else Par1 = Par1 & ", [" & qt & "]"
                Permitido = True
                qt = qt + 1
            Loop
            Pesquisa = "Month(" & Data_pesquisa & ") >= '" & MesX & "' and Year(" & Data_pesquisa & ") >= '" & Cmb_ano_de & "' and Month(" & Data_pesquisa & ") <= '" & MesX1 & "' and Year(" & Data_pesquisa & ") <= '" & Cmb_ano_ate & "'"
            Pesquisa1 = "PIVOT (Sum(" & filtrovalor & ") for " & Data_pesquisa1 & " In (" & Par1 & "))"
            Pesquisa3 = "PIVOT (Sum(" & filtrovalor2 & ") for " & Data_pesquisa1 & " In (" & Par1 & "))"
            Pesquisa2 = Data_pesquisa1
        Case "Ano":
            qt = Cmb_ano_de1
            Qtd = Cmb_ano_ate1
            Do While qt <= Qtd
                If Permitido = False Then Par1 = "[" & qt & "]" Else Par1 = Par1 & ", [" & qt & "]"
                Permitido = True
                qt = qt + 1
            Loop
            Pesquisa = "Year(" & Data_pesquisa & ") >= '" & Cmb_ano_de1 & "' and Year(" & Data_pesquisa & ") <= '" & Cmb_ano_ate1 & "'"
            Pesquisa1 = "PIVOT (Sum(" & filtrovalor & ") for " & Data_pesquisa2 & " In (" & Par1 & "))"
            Pesquisa3 = "PIVOT (Sum(" & filtrovalor2 & ") for " & Data_pesquisa2 & " In (" & Par1 & "))"
            Pesquisa2 = Data_pesquisa2
    End Select
    
    If Cmb_tipo = "A receber e a pagar" Or Cmb_tipo = "Recebidas e pagas" Then
        
        FiltroResumido = "Tipo = 'R' and " & Logsit
        FiltroResumido1 = "Tipo = 'P' and " & Logsit
        If cmbfiltrarpor = "Conta contábil" Then
            FiltroResumido = FiltroResumido & " and TBL.Destino = 'R'"
            'FiltroResumido = Logsit & " and TBL.Destino = 'R'"
            'FiltroResumido1 = Logsit & " and TBL.Destino = 'P'"
            FiltroResumido1 = FiltroResumido1 & " and TBL.Destino = 'P'"
        End If
    ElseIf Cmb_tipo = "A receber e recebidas" Or Cmb_tipo = "A pagar e pagas" Then
        FiltroResumido = "LogSit = 'N'"
        FiltroResumido1 = "LogSit = 'S'"
    Else
        FiltroResumido = Logsit
    End If
    
    If cmbfiltrarpor = "Conta contábil" Then
        If Opt_individual.Value = True Then filtroIndiv = "ID_PC = " & cmbTexto.ItemData(cmbTexto.ListIndex) & " and " Else filtroIndiv = ""
        
        TBCarteira.Open "SELECT Codigo, Descricao, LogSit, Nivel, " & Par1 & " From (Select P.Codigo, P.Descricao, P.LogSit, P.Nivel, " & Pesquisa2 & ", P.Valor from Plano_de_contas_totalizacao P INNER JOIN tbl_familia TBL ON TBL.int_codfamilia = P.ID_PC Where " & filtroIndiv & Pesquisa & " and P.Modulo = '" & Formulario & "' and P.Responsavel = '" & pubUsuario & "' and " & FiltroResumido & ") p " & Pesquisa1 & " as pvt", Conexao, adOpenKeyset, adLockReadOnly
        If Cmb_tipo = "A receber e recebidas" Or Cmb_tipo = "A pagar e pagas" Or Cmb_tipo = "A receber e a pagar" Or Cmb_tipo = "Recebidas e pagas" Then
            'Aqui usa o filtro duplicado
            Set TBGravar = CreateObject("adodb.recordset")
            TBGravar.Open "SELECT Codigo, Descricao, LogSit, Nivel, " & Par1 & " From (Select P.Codigo, P.Descricao, P.LogSit, P.Nivel, " & Pesquisa2 & ", P.Valor_pgto_receb from Plano_de_contas_totalizacao P INNER JOIN tbl_familia TBL ON TBL.int_codfamilia = P.ID_PC Where " & filtroIndiv & Pesquisa & " and P.Modulo = '" & Formulario & "' and P.Responsavel = '" & pubUsuario & "' and " & FiltroResumido1 & ") p " & Pesquisa3 & " as pvt", Conexao, adOpenKeyset, adLockReadOnly
        End If
    Else
        If Opt_individual.Value = True Then filtroIndiv = FamiliaAntiga & " = '" & cmbTexto & "' and " Else filtroIndiv = ""
        
        TBCarteira.Open "SELECT " & FamiliaAntiga & ", LogSit, " & Par1 & " From (Select " & FamiliaAntiga & ", LogSit, " & Pesquisa2 & ", " & filtrovalor & " from " & TabelaFiltro & " Where " & filtroIndiv & Pesquisa & " and Bloqueado = 'False' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and " & FiltroResumido & ") p " & Pesquisa1 & " as pvt", Conexao, adOpenKeyset, adLockReadOnly
        If Cmb_tipo = "A receber e recebidas" Or Cmb_tipo = "A pagar e pagas" Or Cmb_tipo = "A receber e a pagar" Or Cmb_tipo = "Recebidas e pagas" Then
            'Aqui usa o filtro duplicado
            Set TBGravar = CreateObject("adodb.recordset")
            TBGravar.Open "SELECT " & FamiliaAntiga & ", LogSit, " & Par1 & " From (Select " & FamiliaAntiga & ", LogSit, " & Pesquisa2 & ", " & filtrovalor2 & " from " & TabelaFiltro & " Where " & filtroIndiv & Pesquisa & " and Bloqueado = 'False' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and " & FiltroResumido1 & ") p " & Pesquisa3 & " as pvt", Conexao, adOpenKeyset, adLockReadOnly
        End If
    End If
End If
ProcFiltrar1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcFiltrar1()
On Error GoTo tratar_erro

valor = 0
Valor1 = 0
Valor2 = 0
If TBCarteira.EOF = False Then
    Permitido = True
    PBLista.Min = 0
    PBLista.Max = TBCarteira.RecordCount
    PBLista.Value = 1
    contador = 0
    Do While TBCarteira.EOF = False
        Set TBProdutividade = CreateObject("adodb.recordset")
        If Opt_individual.Value = True And optDetalhado.Value = True Then
            TBProdutividade.Open "Select * from Producao_Relatorios", Conexao, adOpenKeyset, adLockOptimistic
            If Cmb_tipo = "A receber" Or Cmb_tipo = "Recebidas" Or Cmb_tipo = "A receber e recebidas" Then
                ProcEnviaDadosDetalhadoRec
            ElseIf Cmb_tipo = "A receber e a pagar" Or Cmb_tipo = "Recebidas e pagas" Then
                ProcEnviaDadosDetalhadoRecPag
            Else
                ProcEnviaDadosDetalhadoPag
            End If
        Else
            ProcCriarResumido
        End If
        TBCarteira.MoveNext
        contador = contador + 1
        PBLista.Value = contador
    Loop
End If
TBCarteira.Close

If optResumido.Value = True Then
    'If Cmb_tipo = "Pagas" Or Cmb_tipo = "Recebidas" Or Cmb_tipo = "A pagar e pagas" Or Cmb_tipo = "A receber e recebidas" Then
    If Cmb_tipo = "A pagar e pagas" Or Cmb_tipo = "A receber e recebidas" Or Cmb_tipo = "A receber e a pagar" Or Cmb_tipo = "Recebidas e pagas" Then
        valor = 0
        Valor1 = 0
        Valor2 = 0
        If TBGravar.EOF = False Then
            Permitido = True
            TBGravar.MoveLast
            PBLista.Min = 0
            PBLista.Max = TBGravar.RecordCount
            PBLista.Value = 1
            contador = 0
            TBGravar.MoveFirst
            Do While TBGravar.EOF = False
                ProcCriarResumido1
                TBGravar.MoveNext
                contador = contador + 1
                PBLista.Value = contador
            Loop
        End If
        TBGravar.Close
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcEnviaDadosDetalhadoPag()
On Error GoTo tratar_erro

TBProdutividade.AddNew
Select Case Cmb_tipo
    Case "A pagar": TBProdutividade!Nota = "P1"
    Case "Pagas": TBProdutividade!Nota = "P2"
    Case "A pagar e pagas": TBProdutividade!Nota = "P3"
End Select
TBProdutividade!Responsavel = pubUsuario
TBProdutividade!Modulo = Formulario
TBProdutividade!maquina = cmbTexto
TBProdutividade!Ordem = TBCarteira!IDintconta 'ID da conta
If TBCarteira!Logsit = "N" Then TBProdutividade!Tipo = "N" Else TBProdutividade!Tipo = "S"
TBProdutividade!Data = IIf(IsNull(TBCarteira!Dt_emissao), Null, Format(TBCarteira!Dt_emissao, "dd/mm/yy")) 'Data emissão
TBProdutividade!Data6 = IIf(IsNull(TBCarteira!dt_Pagamento), Null, Format(TBCarteira!dt_Pagamento, "dd/mm/yy")) 'Data vencimento

If cmbfiltrarpor = "Conta contábil" Then
    Valor_Cofins_Prod = IIf(IsNull(TBCarteira!quantidade), 0, TBCarteira!quantidade)
    Valor_Cofins_Serv = IIf(IsNull(TBCarteira!quantidade), 0, TBCarteira!quantidade)
Else
    Valor_Cofins_Prod = IIf(IsNull(TBCarteira!dbl_valorpagto), 0, TBCarteira!dbl_valorpagto)
    Valor_Cofins_Serv = IIf(IsNull(TBCarteira!ValorPago), 0, TBCarteira!ValorPago)
End If

TBProdutividade!qtdeOK = Valor_Cofins_Prod 'Total a pagar
If TBCarteira!Logsit = "N" Then Valor1 = Valor1 + TBProdutividade!qtdeOK 'Total a pagar
    
If TBCarteira!Logsit = "S" Then
    TBProdutividade!qtdeNC = Valor_Cofins_Serv 'Total pago
    Valor2 = Valor2 + TBProdutividade!qtdeNC 'Total pago
Else
    TBProdutividade!qtdeNC = 0
End If

TBProdutividade!Totalhsutil = IIf(IsNull(TBCarteira!txt_ndocumento), "", TBCarteira!txt_ndocumento) 'Número do documento
TBProdutividade!Data1 = IIf(IsNull(TBCarteira!txt_Parcela), "", TBCarteira!txt_Parcela) 'Parcela
TBProdutividade!DescEvento = IIf(IsNull(TBCarteira!Txt_fornecedor), "", TBCarteira!Txt_fornecedor) 'Fornecedor
TBProdutividade!Data5 = IIf(IsNull(TBCarteira!DataBaixa), Null, Format(TBCarteira!DataBaixa, "dd/mm/yy")) 'Data pagamento
TBProdutividade!Data4 = IIf(IsNull(TBCarteira!NDoctoBaixa), "", TBCarteira!NDoctoBaixa) 'Documento baixa

If optEmisao.Value = True Then
    TBProdutividade!Data7 = IIf(IsNull(TBCarteira!Dt_emissao), Null, Format(TBCarteira!Dt_emissao, "dd/mm/yy"))
ElseIf optVencimento.Value = True Then
        TBProdutividade!Data7 = IIf(IsNull(TBCarteira!dt_Pagamento), Null, Format(TBCarteira!dt_Pagamento, "dd/mm/yy"))
    Else
        TBProdutividade!Data7 = IIf(IsNull(TBCarteira!DataBaixa), Null, Format(TBCarteira!DataBaixa, "dd/mm/yy"))
End If

TBProdutividade.Update
TBProdutividade.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcEnviaDadosDetalhadoRec()
On Error GoTo tratar_erro

TBProdutividade.AddNew
Select Case Cmb_tipo
    Case "A receber": TBProdutividade!Nota = "R1"
    Case "Recebidas": TBProdutividade!Nota = "R2"
    Case "A receber e recebidas": TBProdutividade!Nota = "R3"
End Select
TBProdutividade!Responsavel = pubUsuario
TBProdutividade!Modulo = Formulario
TBProdutividade!maquina = cmbTexto
TBProdutividade!Ordem = TBCarteira!IDintconta 'ID da conta
If TBCarteira!Logsit = "N" Then TBProdutividade!Tipo = "N" Else TBProdutividade!Tipo = "S"
TBProdutividade!Data = IIf(IsNull(TBCarteira!emissao), Null, Format(TBCarteira!emissao, "dd/mm/yy")) 'Data emissão
TBProdutividade!Data6 = IIf(IsNull(TBCarteira!Vencimento), Null, Format(TBCarteira!Vencimento, "dd/mm/yy")) 'Data vencimento

If cmbfiltrarpor = "Conta contábil" Then
    Valor_Cofins_Prod = IIf(IsNull(TBCarteira!quantidade), 0, TBCarteira!quantidade)
    Valor_Cofins_Serv = IIf(IsNull(TBCarteira!quantidade), 0, TBCarteira!quantidade)
Else
    Valor_Cofins_Prod = IIf(IsNull(TBCarteira!valor), 0, TBCarteira!valor)
    Valor_Cofins_Serv = IIf(IsNull(TBCarteira!valortitulorecebido), 0, TBCarteira!valortitulorecebido)
End If

TBProdutividade!qtdeOK = Valor_Cofins_Prod 'Total a receber
If TBCarteira!status = "DUPLICATA DESCONTADA EM ABERTO" Then
    valor = valor + TBProdutividade!qtdeOK 'Total decontado
Else
    If TBCarteira!Logsit = "N" Then Valor1 = Valor1 + TBProdutividade!qtdeOK 'Total a receber
End If

If TBCarteira!Logsit = "S" Then
    TBProdutividade!qtdeNC = Valor_Cofins_Serv 'Total recebido
    Valor2 = Valor2 + TBProdutividade!qtdeNC 'Total recebido
Else
    TBProdutividade!qtdeNC = 0
End If

TBProdutividade!Totalhsutil = IIf(IsNull(TBCarteira!txt_ndocumento), "", TBCarteira!txt_ndocumento) 'Número do documento
TBProdutividade!Data2 = IIf(IsNull(TBCarteira!NFiscal), "", TBCarteira!NFiscal) 'NF
TBProdutividade!Data1 = IIf(IsNull(TBCarteira!Parcela), "", TBCarteira!Parcela) 'Parcela
TBProdutividade!DescEvento = IIf(IsNull(TBCarteira!Nome_Razao), "", TBCarteira!Nome_Razao) 'Cliente
TBProdutividade!Data5 = IIf(IsNull(TBCarteira!Data_pagamento), Null, Format(TBCarteira!Data_pagamento, "dd/mm/yy")) 'Data recebimento
TBProdutividade!Data4 = IIf(IsNull(TBCarteira!NDoctoBaixa), "", TBCarteira!NDoctoBaixa) 'Documento baixa

If optEmisao.Value = True Then
    TBProdutividade!Data7 = IIf(IsNull(TBCarteira!emissao), Null, Format(TBCarteira!emissao, "dd/mm/yy"))
ElseIf optVencimento.Value = True Then
        TBProdutividade!Data7 = IIf(IsNull(TBCarteira!Vencimento), Null, Format(TBCarteira!Vencimento, "dd/mm/yy"))
    Else
        TBProdutividade!Data7 = IIf(IsNull(TBCarteira!Data_pagamento), Null, Format(TBCarteira!Data_pagamento, "dd/mm/yy"))
End If

TBProdutividade.Update
TBProdutividade.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcEnviaDadosDetalhadoRecPag()
On Error GoTo tratar_erro

TBProdutividade.AddNew
Select Case Cmb_tipo
    Case "A receber e a pagar": TBProdutividade!Nota = "RP1"
    Case "Recebidas e pagas": TBProdutividade!Nota = "RP2"
End Select
TBProdutividade!Responsavel = pubUsuario
TBProdutividade!Modulo = Formulario
TBProdutividade!maquina = cmbTexto
TBProdutividade!Ordem = TBCarteira!IDintconta 'ID da conta
If TBCarteira!Logsit = "N" Then TBProdutividade!Tipo = "N" Else TBProdutividade!Tipo = "S"
TBProdutividade!Data = IIf(IsNull(TBCarteira!emissao), Null, Format(TBCarteira!emissao, "dd/mm/yy")) 'Data emissão
TBProdutividade!Data6 = IIf(IsNull(TBCarteira!Vencimento), Null, Format(TBCarteira!Vencimento, "dd/mm/yy")) 'Data vencimento

If cmbfiltrarpor = "Conta contábil" Then
    Valor_Cofins_Prod = IIf(IsNull(TBCarteira!quantidade), 0, TBCarteira!quantidade) 'a receber/a pagar
    If TBCarteira!Logsit = "S" Then Valor_Cofins_Serv = IIf(IsNull(TBCarteira!quantidade), 0, TBCarteira!quantidade) 'recebido/pago
Else
    Valor_Cofins_Prod = IIf(IsNull(TBCarteira!valor), 0, TBCarteira!valor) 'a receber/a pagar
    If TBCarteira!Logsit = "S" Then Valor_Cofins_Serv = IIf(IsNull(TBCarteira!vlrBaixa), 0, TBCarteira!vlrBaixa) 'recebido/pago
End If

TBProdutividade!qtdeOK = Valor_Cofins_Prod 'Total a receber/a pagar
TBProdutividade!qtdeNC = Valor_Cofins_Serv 'Total a pagar/pago

If TBCarteira!status = "DUPLICATA DESCONTADA EM ABERTO" Then
    valor = valor + TBProdutividade!qtdeOK 'Total decontado
Else
    If TBCarteira!Tipo = "R" Then
        If TBCarteira!Logsit = "N" Then
            Valor1 = Valor1 + Valor_Cofins_Prod 'Total a receber
        Else
            Valor1 = Valor1 + Valor_Cofins_Serv 'Total recebido
        End If
    End If
End If

If TBCarteira!Tipo = "P" Then
    If TBCarteira!Logsit = "N" Then
        Valor2 = Valor2 + Valor_Cofins_Prod 'Total a pagar
    Else
        Valor2 = Valor2 + Valor_Cofins_Serv 'Total pago
    End If
End If

TBProdutividade!Totalhsutil = IIf(IsNull(TBCarteira!txt_ndocumento), "", TBCarteira!txt_ndocumento) 'Número do documento
TBProdutividade!Data2 = IIf(IsNull(TBCarteira!NFiscal), "", TBCarteira!NFiscal) 'NF
TBProdutividade!Data1 = IIf(IsNull(TBCarteira!Parcela), "", TBCarteira!Parcela) 'Parcela
TBProdutividade!DescEvento = IIf(IsNull(TBCarteira!Razao), "", TBCarteira!Razao) 'Cliente/forncedor
TBProdutividade!Data5 = IIf(IsNull(TBCarteira!Data_pagamento), Null, Format(TBCarteira!Data_pagamento, "dd/mm/yy")) 'Data recebimento
TBProdutividade!Data4 = IIf(IsNull(TBCarteira!NDoctoBaixa), "", TBCarteira!NDoctoBaixa) 'Documento baixa
TBProdutividade!Totalhsprev = IIf(IsNull(TBCarteira!Tipo), "", TBCarteira!Tipo) 'R para receber e P para pagar

If optEmisao.Value = True Then
    TBProdutividade!Data7 = IIf(IsNull(TBCarteira!emissao), Null, Format(TBCarteira!emissao, "dd/mm/yy"))
ElseIf optVencimento.Value = True Then
        TBProdutividade!Data7 = IIf(IsNull(TBCarteira!Vencimento), Null, Format(TBCarteira!Vencimento, "dd/mm/yy"))
    Else
        TBProdutividade!Data7 = IIf(IsNull(TBCarteira!Data_pagamento), Null, Format(TBCarteira!Data_pagamento, "dd/mm/yy"))
End If

TBProdutividade.Update
TBProdutividade.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcCriarResumido()
On Error GoTo tratar_erro

Permitido = True
Select Case cmbPor
    Case "Dia":
        qt = 0
        Dataini = msk_fltInicio
        DataFim = msk_fltFim
        Do While Dataini <= DataFim
            qt = qt + 1
            If Cmb_tipo = "A receber" Or Cmb_tipo = "Recebidas" Or Cmb_tipo = "A receber e recebidas" Then
                ProcEnviaDadosResumidoRec
            ElseIf Cmb_tipo = "A receber e a pagar" Or Cmb_tipo = "Recebidas e pagas" Then
                ProcEnviaDadosResumidoRecPag
            Else
                ProcEnviaDadosResumidoPag
            End If
            Dataini = Dataini + 1
        Loop
    Case "Mês":
        qt = MesX
        Qtd = MesX1
        Do While qt <= Qtd
            If Cmb_tipo = "A receber" Or Cmb_tipo = "Recebidas" Or Cmb_tipo = "A receber e recebidas" Then
                ProcEnviaDadosResumidoRec
            ElseIf Cmb_tipo = "A receber e a pagar" Or Cmb_tipo = "Recebidas e pagas" Then
                ProcEnviaDadosResumidoRecPag
            Else
                ProcEnviaDadosResumidoPag
            End If
            qt = qt + 1
        Loop
    Case "Ano":
        qt = Cmb_ano_de1
        Qtd = Cmb_ano_ate1
        Do While qt <= Qtd
            If Cmb_tipo = "A receber" Or Cmb_tipo = "Recebidas" Or Cmb_tipo = "A receber e recebidas" Then
                ProcEnviaDadosResumidoRec
            ElseIf Cmb_tipo = "A receber e a pagar" Or Cmb_tipo = "Recebidas e pagas" Then
                ProcEnviaDadosResumidoRecPag
            Else
                ProcEnviaDadosResumidoPag
            End If
            qt = qt + 1
        Loop
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcCriarResumido1()
On Error GoTo tratar_erro

Permitido = True
Select Case cmbPor
    Case "Dia":
        qt = 0
        Dataini = msk_fltInicio
        DataFim = msk_fltFim
        Do While Dataini <= DataFim
            qt = qt + 1
            If Cmb_tipo = "A receber" Or Cmb_tipo = "Recebidas" Or Cmb_tipo = "A receber e recebidas" Then
                ProcEnviaDadosResumidoRec1
            ElseIf Cmb_tipo = "A receber e a pagar" Or Cmb_tipo = "Recebidas e pagas" Then
                ProcEnviaDadosResumidoRecPag1
            Else
                ProcEnviaDadosResumidoPag1
            End If
            Dataini = Dataini + 1
        Loop
    Case "Mês":
        qt = MesX
        Qtd = MesX1
        Do While qt <= Qtd
            If Cmb_tipo = "A receber" Or Cmb_tipo = "Recebidas" Or Cmb_tipo = "A receber e recebidas" Then
                ProcEnviaDadosResumidoRec1
            ElseIf Cmb_tipo = "A receber e a pagar" Or Cmb_tipo = "Recebidas e pagas" Then
                ProcEnviaDadosResumidoRecPag1
            Else
                ProcEnviaDadosResumidoPag1
            End If
            qt = qt + 1
        Loop
    Case "Ano":
        qt = Cmb_ano_de1
        Qtd = Cmb_ano_ate1
        Do While qt <= Qtd
            If Cmb_tipo = "A receber" Or Cmb_tipo = "Recebidas" Or Cmb_tipo = "A receber e recebidas" Then
                ProcEnviaDadosResumidoRec1
            ElseIf Cmb_tipo = "A receber e a pagar" Or Cmb_tipo = "Recebidas e pagas" Then
                ProcEnviaDadosResumidoRecPag1
            Else
                ProcEnviaDadosResumidoPag1
            End If
            qt = qt + 1
        Loop
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcEnviaDadosResumidoPag()
On Error GoTo tratar_erro

Select Case cmbfiltrarpor
    Case "Conta contábil": Familiatext = TBCarteira!CODIGO & " - " & TBCarteira!Descricao
    Case "Documento": Familiatext = IIf(IsNull(TBCarteira!txt_ndocumento), "", TBCarteira!txt_ndocumento)
    Case "Fornecedor": Familiatext = TBCarteira!Txt_fornecedor
    Case "Tipo do doc.": Familiatext = TBCarteira!Class_conta
    Case "Status": Familiatext = TBCarteira!status
    Case "Instituição": Familiatext = TBCarteira!Banco
    Case "Documento baixa": Familiatext = IIf(IsNull(TBCarteira!NDoctoBaixa), "", TBCarteira!NDoctoBaixa)
    Case "Forma de pagamento": Familiatext = IIf(IsNull(TBCarteira!FormaBaixa), "", TBCarteira!FormaBaixa)
End Select
Select Case cmbPor
    Case "Dia":
        DataFiltro = "Data = '" & Dataini & "'"
        DataTexto = Dataini
    Case "Mês":
        DataFiltro = "Execucaoprev = '" & qt & "/" & Cmb_ano_de & "'"
        DataTexto = "01/" & qt & "/" & Cmb_ano_de
    Case "Ano":
        DataFiltro = "Execucaoprev = '" & qt & "'"
        DataTexto = "01" & "/01/" & qt
End Select

'If cmbfiltrarpor <> "Conta contábil" Then Familiatext = Left(Familiatext, 25)

Set TBProdutividade = CreateObject("adodb.recordset")
TBProdutividade.Open "Select * from Producao_Relatorios where Modulo = '" & Formulario & "' and Responsavel = '" & pubUsuario & "' and Maquina = '" & Familiatext & "' and " & DataFiltro, Conexao, adOpenKeyset, adLockOptimistic
If TBProdutividade.EOF = True Then TBProdutividade.AddNew
TBProdutividade!Responsavel = pubUsuario
TBProdutividade!Modulo = Formulario
TBProdutividade!Data = Format(DataTexto, "dd/mm/yyyy")
Select Case cmbPor
    Case "Dia":
        DiaX = Dataini
        TotalCreditar = IIf(IsNull(TBCarteira(DiaX)), 0, TBCarteira(DiaX))
        TBProdutividade!Execucaoprev = Format(Dataini, "dd/mm/yy")
    Case "Mês":
        Select Case qt
            Case 1: TotalCreditar = IIf(IsNull(TBCarteira![1]), 0, Format(TBCarteira![1], "###,##0.00"))
            Case 2: TotalCreditar = IIf(IsNull(TBCarteira![2]), 0, Format(TBCarteira![2], "###,##0.00"))
            Case 3: TotalCreditar = IIf(IsNull(TBCarteira![3]), 0, Format(TBCarteira![3], "###,##0.00"))
            Case 4: TotalCreditar = IIf(IsNull(TBCarteira![4]), 0, Format(TBCarteira![4], "###,##0.00"))
            Case 5: TotalCreditar = IIf(IsNull(TBCarteira![5]), 0, Format(TBCarteira![5], "###,##0.00"))
            Case 6: TotalCreditar = IIf(IsNull(TBCarteira![6]), 0, Format(TBCarteira![6], "###,##0.00"))
            Case 7: TotalCreditar = IIf(IsNull(TBCarteira![7]), 0, Format(TBCarteira![7], "###,##0.00"))
            Case 8: TotalCreditar = IIf(IsNull(TBCarteira![8]), 0, Format(TBCarteira![8], "###,##0.00"))
            Case 9: TotalCreditar = IIf(IsNull(TBCarteira![9]), 0, Format(TBCarteira![9], "###,##0.00"))
            Case 10: TotalCreditar = IIf(IsNull(TBCarteira![10]), 0, Format(TBCarteira![10], "###,##0.00"))
            Case 11: TotalCreditar = IIf(IsNull(TBCarteira![11]), 0, Format(TBCarteira![11], "###,##0.00"))
            Case 12: TotalCreditar = IIf(IsNull(TBCarteira![12]), 0, Format(TBCarteira![12], "###,##0.00"))
        End Select
        TBProdutividade!Execucaoprev = qt & "/" & Cmb_ano_de
    Case "Ano":
        DiaX = qt
        TotalCreditar = IIf(IsNull(TBCarteira(DiaX)), 0, TBCarteira(DiaX))
        TBProdutividade!Execucaoprev = qt
End Select
If TBCarteira!Logsit = "N" Then
    If Cmb_tipo = "A pagar" Or Cmb_tipo = "A pagar e pagas" Then TBProdutividade!qtdeOK = IIf(IsNull(TotalCreditar), 0, Format(TotalCreditar, "###,##0.00")) 'Total a pagar
Else
    If Cmb_tipo = "Pagas" Or Cmb_tipo = "A pagar e pagas" Then TBProdutividade!qtdeNC = IIf(IsNull(TotalCreditar), 0, Format(TotalCreditar, "###,##0.00")) 'Total pago
End If
TBProdutividade!Ordem = qt

TBProdutividade!maquina = Familiatext
Select Case Cmb_tipo
    Case "A pagar": TBProdutividade!Nota = "P1"
    Case "Pagas": TBProdutividade!Nota = "P2"
    Case "A pagar e pagas": TBProdutividade!Nota = "P3"
End Select

TBProdutividade.Update
TBProdutividade.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcEnviaDadosResumidoPag1()
On Error GoTo tratar_erro

Select Case cmbfiltrarpor
    Case "Conta contábil": Familiatext = TBGravar!CODIGO & " - " & TBGravar!Descricao
    Case "Documento": Familiatext = IIf(IsNull(TBGravar!txt_ndocumento), "", TBGravar!txt_ndocumento)
    Case "Fornecedor": Familiatext = TBGravar!Txt_fornecedor
    Case "Tipo do doc.": Familiatext = TBGravar!Class_conta
    Case "Status": Familiatext = TBGravar!status
    Case "Instituição": Familiatext = TBGravar!Banco
    Case "Documento baixa": Familiatext = IIf(IsNull(TBGravar!NDoctoBaixa), "", TBGravar!NDoctoBaixa)
    Case "Forma de pagamento": Familiatext = IIf(IsNull(TBGravar!FormaBaixa), "", TBGravar!FormaBaixa)
End Select
Select Case cmbPor
    Case "Dia":
        DataFiltro = "Data = '" & Dataini & "'"
        DataTexto = Dataini
    Case "Mês":
        DataFiltro = "Execucaoprev = '" & qt & "/" & Cmb_ano_de & "'"
        DataTexto = "01/" & qt & "/" & Cmb_ano_de
    Case "Ano":
        DataFiltro = "Execucaoprev = '" & qt & "'"
        DataTexto = "01" & "/01/" & qt
End Select

'If cmbfiltrarpor <> "Conta contábil" Then Familiatext = Left(Familiatext, 25)

Set TBProdutividade = CreateObject("adodb.recordset")
TBProdutividade.Open "Select * from Producao_Relatorios where Modulo = '" & Formulario & "' and Responsavel = '" & pubUsuario & "' and Maquina = '" & Familiatext & "' and " & DataFiltro, Conexao, adOpenKeyset, adLockOptimistic
If TBProdutividade.EOF = True Then TBProdutividade.AddNew
TBProdutividade!Responsavel = pubUsuario
TBProdutividade!Modulo = Formulario
TBProdutividade!Data = Format(DataTexto, "dd/mm/yyyy")
Select Case cmbPor
    Case "Dia":
        DiaX = Dataini
        TotalCreditar = IIf(IsNull(TBGravar(DiaX)), 0, TBGravar(DiaX))
        TBProdutividade!Execucaoprev = Format(Dataini, "dd/mm/yy")
    Case "Mês":
        Select Case qt
            Case 1: TotalCreditar = IIf(IsNull(TBGravar![1]), 0, Format(TBGravar![1], "###,##0.00"))
            Case 2: TotalCreditar = IIf(IsNull(TBGravar![2]), 0, Format(TBGravar![2], "###,##0.00"))
            Case 3: TotalCreditar = IIf(IsNull(TBGravar![3]), 0, Format(TBGravar![3], "###,##0.00"))
            Case 4: TotalCreditar = IIf(IsNull(TBGravar![4]), 0, Format(TBGravar![4], "###,##0.00"))
            Case 5: TotalCreditar = IIf(IsNull(TBGravar![5]), 0, Format(TBGravar![5], "###,##0.00"))
            Case 6: TotalCreditar = IIf(IsNull(TBGravar![6]), 0, Format(TBGravar![6], "###,##0.00"))
            Case 7: TotalCreditar = IIf(IsNull(TBGravar![7]), 0, Format(TBGravar![7], "###,##0.00"))
            Case 8: TotalCreditar = IIf(IsNull(TBGravar![8]), 0, Format(TBGravar![8], "###,##0.00"))
            Case 9: TotalCreditar = IIf(IsNull(TBGravar![9]), 0, Format(TBGravar![9], "###,##0.00"))
            Case 10: TotalCreditar = IIf(IsNull(TBGravar![10]), 0, Format(TBGravar![10], "###,##0.00"))
            Case 11: TotalCreditar = IIf(IsNull(TBGravar![11]), 0, Format(TBGravar![11], "###,##0.00"))
            Case 12: TotalCreditar = IIf(IsNull(TBGravar![12]), 0, Format(TBGravar![12], "###,##0.00"))
        End Select
        TBProdutividade!Execucaoprev = qt & "/" & Cmb_ano_de
    Case "Ano":
        DiaX = qt
        TotalCreditar = IIf(IsNull(TBGravar(DiaX)), 0, TBGravar(DiaX))
        TBProdutividade!Execucaoprev = qt
End Select
TBProdutividade!qtdeNC = IIf(IsNull(TotalCreditar), 0, Format(TotalCreditar, "###,##0.00")) 'Total pago
TBProdutividade!Ordem = qt

TBProdutividade!maquina = Familiatext
Select Case Cmb_tipo
    Case "A pagar": TBProdutividade!Nota = "P1"
    Case "Pagas": TBProdutividade!Nota = "P2"
    Case "A pagar e pagas": TBProdutividade!Nota = "P3"
End Select

TBProdutividade.Update
TBProdutividade.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcEnviaDadosResumidoRec()
On Error GoTo tratar_erro

Select Case cmbfiltrarpor
    Case "Conta contábil": Familiatext = TBCarteira!CODIGO & " - " & TBCarteira!Descricao
    Case "Nota fiscal": Familiatext = IIf(IsNull(TBCarteira!NFiscal), "", TBCarteira!NFiscal)
    Case "Documento": Familiatext = IIf(IsNull(TBCarteira!txt_ndocumento), "", TBCarteira!txt_ndocumento)
    Case "Cliente": Familiatext = TBCarteira!Nome_Razao
    Case "Status": Familiatext = TBCarteira!status
    Case "Instituição": Familiatext = TBCarteira!Banco
    Case "Documento baixa": Familiatext = IIf(IsNull(TBCarteira!txt_ndocumento), "", TBCarteira!txt_ndocumento)
End Select
Select Case cmbPor
    Case "Dia":
        DataFiltro = "Data = '" & Dataini & "'"
        DataTexto = Dataini
    Case "Mês":
        DataFiltro = "Execucaoprev = '" & qt & "/" & Cmb_ano_de & "'"
        DataTexto = "01/" & qt & "/" & Cmb_ano_de
    Case "Ano":
        DataFiltro = "Execucaoprev = '" & qt & "'"
        DataTexto = "01" & "/01/" & qt
End Select

'If cmbfiltrarpor <> "Conta contábil" Then Familiatext = Left(Familiatext, 25)

Set TBProdutividade = CreateObject("adodb.recordset")
TBProdutividade.Open "Select * from Producao_Relatorios where Modulo = '" & Formulario & "' and Responsavel = '" & pubUsuario & "' and Maquina = '" & Familiatext & "' and " & DataFiltro, Conexao, adOpenKeyset, adLockOptimistic
If TBProdutividade.EOF = True Then TBProdutividade.AddNew
TBProdutividade!Responsavel = pubUsuario
TBProdutividade!Modulo = Formulario
TBProdutividade!Data = Format(DataTexto, "dd/mm/yyyy")
Select Case cmbPor
    Case "Dia":
        DiaX = Dataini
        TotalCreditar = IIf(IsNull(TBCarteira(DiaX)), 0, TBCarteira(DiaX))
        TBProdutividade!Execucaoprev = Format(Dataini, "dd/mm/yy")
    Case "Mês":
        Select Case qt
            Case 1: TotalCreditar = IIf(IsNull(TBCarteira![1]), 0, Format(TBCarteira![1], "###,##0.00"))
            Case 2: TotalCreditar = IIf(IsNull(TBCarteira![2]), 0, Format(TBCarteira![2], "###,##0.00"))
            Case 3: TotalCreditar = IIf(IsNull(TBCarteira![3]), 0, Format(TBCarteira![3], "###,##0.00"))
            Case 4: TotalCreditar = IIf(IsNull(TBCarteira![4]), 0, Format(TBCarteira![4], "###,##0.00"))
            Case 5: TotalCreditar = IIf(IsNull(TBCarteira![5]), 0, Format(TBCarteira![5], "###,##0.00"))
            Case 6: TotalCreditar = IIf(IsNull(TBCarteira![6]), 0, Format(TBCarteira![6], "###,##0.00"))
            Case 7: TotalCreditar = IIf(IsNull(TBCarteira![7]), 0, Format(TBCarteira![7], "###,##0.00"))
            Case 8: TotalCreditar = IIf(IsNull(TBCarteira![8]), 0, Format(TBCarteira![8], "###,##0.00"))
            Case 9: TotalCreditar = IIf(IsNull(TBCarteira![9]), 0, Format(TBCarteira![9], "###,##0.00"))
            Case 10: TotalCreditar = IIf(IsNull(TBCarteira![10]), 0, Format(TBCarteira![10], "###,##0.00"))
            Case 11: TotalCreditar = IIf(IsNull(TBCarteira![11]), 0, Format(TBCarteira![11], "###,##0.00"))
            Case 12: TotalCreditar = IIf(IsNull(TBCarteira![12]), 0, Format(TBCarteira![12], "###,##0.00"))
        End Select
        TBProdutividade!Execucaoprev = qt & "/" & Cmb_ano_de
    Case "Ano":
        DiaX = qt
        TotalCreditar = IIf(IsNull(TBCarteira(DiaX)), 0, TBCarteira(DiaX))
        TBProdutividade!Execucaoprev = qt
End Select
If TBCarteira!Logsit = "N" Then
    If Cmb_tipo = "A receber" Or Cmb_tipo = "A receber e recebidas" Then TBProdutividade!qtdeOK = IIf(IsNull(TotalCreditar), 0, Format(TotalCreditar, "###,##0.00")) 'Total a receber
Else
    If Cmb_tipo = "Recebidas" Or Cmb_tipo = "A receber e recebidas" Then TBProdutividade!qtdeNC = IIf(IsNull(TotalCreditar), 0, Format(TotalCreditar, "###,##0.00")) 'Total recebido
End If
TBProdutividade!Ordem = qt

TBProdutividade!maquina = Familiatext
Select Case Cmb_tipo
    Case "A receber": TBProdutividade!Nota = "R1"
    Case "Recebidas": TBProdutividade!Nota = "R2"
    Case "A receber e recebidas": TBProdutividade!Nota = "R3"
End Select

TBProdutividade.Update
TBProdutividade.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcEnviaDadosResumidoRecPag()
On Error GoTo tratar_erro

Select Case cmbfiltrarpor
    Case "Conta contábil": Familiatext = TBCarteira!CODIGO & " - " & TBCarteira!Descricao
    Case "Documento": Familiatext = IIf(IsNull(TBCarteira!txt_ndocumento), "", TBCarteira!txt_ndocumento)
    Case "Tipo do doc.": Familiatext = IIf(IsNull(TBCarteira!Tipo_doc), "", TBCarteira!Tipo_doc)
    Case "Cliente/Fornecedor": Familiatext = TBCarteira!Razao
    Case "Status": Familiatext = TBCarteira!status
    Case "Instituição": Familiatext = TBCarteira!Banco
    Case "Documento baixa": Familiatext = IIf(IsNull(TBCarteira!NDoctoBaixa), "", TBCarteira!NDoctoBaixa)
End Select
Select Case cmbPor
    Case "Dia":
        DataFiltro = "Data = '" & Dataini & "'"
        DataTexto = Dataini
    Case "Mês":
        DataFiltro = "Execucaoprev = '" & qt & "/" & Cmb_ano_de & "'"
        DataTexto = "01/" & qt & "/" & Cmb_ano_de
    Case "Ano":
        DataFiltro = "Execucaoprev = '" & qt & "'"
        DataTexto = "01" & "/01/" & qt
End Select

Set TBProdutividade = CreateObject("adodb.recordset")
TBProdutividade.Open "Select * from Producao_Relatorios where Modulo = '" & Formulario & "' and Responsavel = '" & pubUsuario & "' and Maquina = '" & Familiatext & "' and totalhsprev = 'R' and " & DataFiltro, Conexao, adOpenKeyset, adLockOptimistic
If TBProdutividade.EOF = True Then TBProdutividade.AddNew
TBProdutividade!Responsavel = pubUsuario
TBProdutividade!Modulo = Formulario
TBProdutividade!Data = Format(DataTexto, "dd/mm/yyyy")
Select Case cmbPor
    Case "Dia":
        DiaX = Dataini
        TotalCreditar = IIf(IsNull(TBCarteira(DiaX)), 0, TBCarteira(DiaX))
        TBProdutividade!Execucaoprev = Format(Dataini, "dd/mm/yy")
    Case "Mês":
        Select Case qt
            Case 1: TotalCreditar = IIf(IsNull(TBCarteira![1]), 0, Format(TBCarteira![1], "###,##0.00"))
            Case 2: TotalCreditar = IIf(IsNull(TBCarteira![2]), 0, Format(TBCarteira![2], "###,##0.00"))
            Case 3: TotalCreditar = IIf(IsNull(TBCarteira![3]), 0, Format(TBCarteira![3], "###,##0.00"))
            Case 4: TotalCreditar = IIf(IsNull(TBCarteira![4]), 0, Format(TBCarteira![4], "###,##0.00"))
            Case 5: TotalCreditar = IIf(IsNull(TBCarteira![5]), 0, Format(TBCarteira![5], "###,##0.00"))
            Case 6: TotalCreditar = IIf(IsNull(TBCarteira![6]), 0, Format(TBCarteira![6], "###,##0.00"))
            Case 7: TotalCreditar = IIf(IsNull(TBCarteira![7]), 0, Format(TBCarteira![7], "###,##0.00"))
            Case 8: TotalCreditar = IIf(IsNull(TBCarteira![8]), 0, Format(TBCarteira![8], "###,##0.00"))
            Case 9: TotalCreditar = IIf(IsNull(TBCarteira![9]), 0, Format(TBCarteira![9], "###,##0.00"))
            Case 10: TotalCreditar = IIf(IsNull(TBCarteira![10]), 0, Format(TBCarteira![10], "###,##0.00"))
            Case 11: TotalCreditar = IIf(IsNull(TBCarteira![11]), 0, Format(TBCarteira![11], "###,##0.00"))
            Case 12: TotalCreditar = IIf(IsNull(TBCarteira![12]), 0, Format(TBCarteira![12], "###,##0.00"))
        End Select
        TBProdutividade!Execucaoprev = qt & "/" & Cmb_ano_de
    Case "Ano":
        DiaX = qt
        TotalCreditar = IIf(IsNull(TBCarteira(DiaX)), 0, TBCarteira(DiaX))
        TBProdutividade!Execucaoprev = qt
End Select
TBProdutividade!qtdeOK = IIf(IsNull(TotalCreditar), 0, Format(TotalCreditar, "0.00")) 'Total receber/recebida
TBProdutividade!Ordem = qt
TBProdutividade!Totalhsprev = "R"

TBProdutividade!maquina = Familiatext
Select Case Cmb_tipo
    Case "A receber e a pagar": TBProdutividade!Nota = "RP1"
    Case "Recebidas e pagas": TBProdutividade!Nota = "RP2"
End Select

If cmbfiltrarpor = "Conta contábil" Then TBProdutividade!Turno = TBCarteira!Nivel

TBProdutividade.Update
TBProdutividade.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcEnviaDadosResumidoRecPag1()
On Error GoTo tratar_erro

Select Case cmbfiltrarpor
    Case "Conta contábil": Familiatext = TBGravar!CODIGO & " - " & TBGravar!Descricao
    Case "Documento": Familiatext = IIf(IsNull(TBGravar!txt_ndocumento), "", TBGravar!txt_ndocumento)
    Case "Tipo do doc.": Familiatext = IIf(IsNull(TBGravar!Tipo_doc), "", TBGravar!Tipo_doc)
    Case "Cliente/Fornecedor": Familiatext = TBGravar!Razao
    Case "Status": Familiatext = TBGravar!status
    Case "Instituição": Familiatext = TBGravar!Banco
    Case "Documento baixa": Familiatext = IIf(IsNull(TBGravar!NDoctoBaixa), "", TBGravar!NDoctoBaixa)
End Select
Select Case cmbPor
    Case "Dia":
        DataFiltro = "Data = '" & Dataini & "'"
        DataTexto = Dataini
    Case "Mês":
        DataFiltro = "Execucaoprev = '" & qt & "/" & Cmb_ano_de & "'"
        DataTexto = "01/" & qt & "/" & Cmb_ano_de
    Case "Ano":
        DataFiltro = "Execucaoprev = '" & qt & "'"
        DataTexto = "01" & "/01/" & qt
End Select

Set TBProdutividade = CreateObject("adodb.recordset")
TBProdutividade.Open "Select * from Producao_Relatorios where Modulo = '" & Formulario & "' and Responsavel = '" & pubUsuario & "' and Maquina = '" & Familiatext & "' and totalhsprev = 'P' and " & DataFiltro, Conexao, adOpenKeyset, adLockOptimistic
If TBProdutividade.EOF = True Then TBProdutividade.AddNew
TBProdutividade!Responsavel = pubUsuario
TBProdutividade!Modulo = Formulario
TBProdutividade!Data = Format(DataTexto, "dd/mm/yyyy")
Select Case cmbPor
    Case "Dia":
        DiaX = Dataini
        TotalCreditar = IIf(IsNull(TBGravar(DiaX)), 0, TBGravar(DiaX))
        TBProdutividade!Execucaoprev = Format(Dataini, "dd/mm/yy")
    Case "Mês":
        Select Case qt
            Case 1: TotalCreditar = IIf(IsNull(TBGravar![1]), 0, Format(TBGravar![1], "0.00"))
            Case 2: TotalCreditar = IIf(IsNull(TBGravar![2]), 0, Format(TBGravar![2], "0.00"))
            Case 3: TotalCreditar = IIf(IsNull(TBGravar![3]), 0, Format(TBGravar![3], "0.00"))
            Case 4: TotalCreditar = IIf(IsNull(TBGravar![4]), 0, Format(TBGravar![4], "0.00"))
            Case 5: TotalCreditar = IIf(IsNull(TBGravar![5]), 0, Format(TBGravar![5], "0.00"))
            Case 6: TotalCreditar = IIf(IsNull(TBGravar![6]), 0, Format(TBGravar![6], "0.00"))
            Case 7: TotalCreditar = IIf(IsNull(TBGravar![7]), 0, Format(TBGravar![7], "0.00"))
            Case 8: TotalCreditar = IIf(IsNull(TBGravar![8]), 0, Format(TBGravar![8], "0.00"))
            Case 9: TotalCreditar = IIf(IsNull(TBGravar![9]), 0, Format(TBGravar![9], "0.00"))
            Case 10: TotalCreditar = IIf(IsNull(TBGravar![10]), 0, Format(TBGravar![10], "0.00"))
            Case 11: TotalCreditar = IIf(IsNull(TBGravar![11]), 0, Format(TBGravar![11], "0.00"))
            Case 12: TotalCreditar = IIf(IsNull(TBGravar![12]), 0, Format(TBGravar![12], "0.00"))
        End Select
        TBProdutividade!Execucaoprev = qt & "/" & Cmb_ano_de
    Case "Ano":
        DiaX = qt
        TotalCreditar = IIf(IsNull(TBGravar(DiaX)), 0, TBGravar(DiaX))
        TBProdutividade!Execucaoprev = qt
End Select
TBProdutividade!qtdeNC = IIf(IsNull(TotalCreditar), 0, Format(TotalCreditar, "0.00")) 'Total receber/recebida
TBProdutividade!Ordem = qt
TBProdutividade!Totalhsprev = "P"

TBProdutividade!maquina = Familiatext
Select Case Cmb_tipo
    Case "A receber e a pagar": TBProdutividade!Nota = "RP1"
    Case "Recebidas e pagas": TBProdutividade!Nota = "RP2"
End Select

If cmbfiltrarpor = "Conta contábil" Then TBProdutividade!Turno = TBGravar!Nivel

TBProdutividade.Update
TBProdutividade.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcEnviaDadosResumidoRec1()
On Error GoTo tratar_erro

Select Case cmbfiltrarpor
    Case "Conta contábil": Familiatext = TBGravar!CODIGO & " - " & TBGravar!Descricao
    Case "Nota fiscal": Familiatext = IIf(IsNull(TBGravar!NFiscal), "", TBGravar!NFiscal)
    Case "Documento": Familiatext = IIf(IsNull(TBGravar!txt_ndocumento), "", TBGravar!txt_ndocumento)
    Case "Cliente": Familiatext = TBGravar!Nome_Razao
    Case "Status": Familiatext = TBGravar!status
    Case "Instituição": Familiatext = TBGravar!Banco
    Case "Documento baixa": Familiatext = IIf(IsNull(TBGravar!txt_ndocumento), "", TBGravar!txt_ndocumento)
End Select
Select Case cmbPor
    Case "Dia":
        DataFiltro = "Data = '" & Dataini & "'"
        DataTexto = Dataini
    Case "Mês":
        DataFiltro = "Execucaoprev = '" & qt & "/" & Cmb_ano_de & "'"
        DataTexto = "01/" & qt & "/" & Cmb_ano_de
    Case "Ano":
        DataFiltro = "Execucaoprev = '" & qt & "'"
        DataTexto = "01" & "/01/" & qt
End Select

'If cmbfiltrarpor <> "Conta contábil" Then Familiatext = Left(Familiatext, 25)

Set TBProdutividade = CreateObject("adodb.recordset")
TBProdutividade.Open "Select * from Producao_Relatorios where Modulo = '" & Formulario & "' and Responsavel = '" & pubUsuario & "' and Maquina = '" & Familiatext & "' and " & DataFiltro, Conexao, adOpenKeyset, adLockOptimistic
If TBProdutividade.EOF = True Then TBProdutividade.AddNew
TBProdutividade!Responsavel = pubUsuario
TBProdutividade!Modulo = Formulario
TBProdutividade!Data = Format(DataTexto, "dd/mm/yyyy")
Select Case cmbPor
    Case "Dia":
        DiaX = Dataini
        TotalCreditar = IIf(IsNull(TBGravar(DiaX)), 0, TBGravar(DiaX))
        TBProdutividade!Execucaoprev = Format(Dataini, "dd/mm/yy")
    Case "Mês":
        Select Case qt
            Case 1: TotalCreditar = IIf(IsNull(TBGravar![1]), 0, Format(TBGravar![1], "###,##0.00"))
            Case 2: TotalCreditar = IIf(IsNull(TBGravar![2]), 0, Format(TBGravar![2], "###,##0.00"))
            Case 3: TotalCreditar = IIf(IsNull(TBGravar![3]), 0, Format(TBGravar![3], "###,##0.00"))
            Case 4: TotalCreditar = IIf(IsNull(TBGravar![4]), 0, Format(TBGravar![4], "###,##0.00"))
            Case 5: TotalCreditar = IIf(IsNull(TBGravar![5]), 0, Format(TBGravar![5], "###,##0.00"))
            Case 6: TotalCreditar = IIf(IsNull(TBGravar![6]), 0, Format(TBGravar![6], "###,##0.00"))
            Case 7: TotalCreditar = IIf(IsNull(TBGravar![7]), 0, Format(TBGravar![7], "###,##0.00"))
            Case 8: TotalCreditar = IIf(IsNull(TBGravar![8]), 0, Format(TBGravar![8], "###,##0.00"))
            Case 9: TotalCreditar = IIf(IsNull(TBGravar![9]), 0, Format(TBGravar![9], "###,##0.00"))
            Case 10: TotalCreditar = IIf(IsNull(TBGravar![10]), 0, Format(TBGravar![10], "###,##0.00"))
            Case 11: TotalCreditar = IIf(IsNull(TBGravar![11]), 0, Format(TBGravar![11], "###,##0.00"))
            Case 12: TotalCreditar = IIf(IsNull(TBGravar![12]), 0, Format(TBGravar![12], "###,##0.00"))
        End Select
        TBProdutividade!Execucaoprev = qt & "/" & Cmb_ano_de
    Case "Ano":
        DiaX = qt
        TotalCreditar = IIf(IsNull(TBGravar(DiaX)), 0, TBGravar(DiaX))
        TBProdutividade!Execucaoprev = qt
End Select
TBProdutividade!qtdeNC = IIf(IsNull(TotalCreditar), 0, Format(TotalCreditar, "###,##0.00")) 'Total recebido
TBProdutividade!Ordem = qt

TBProdutividade!maquina = Familiatext
Select Case Cmb_tipo
    Case "A receber": TBProdutividade!Nota = "R1"
    Case "Recebidas": TBProdutividade!Nota = "R2"
    Case "A receber e recebidas": TBProdutividade!Nota = "R3"
End Select

TBProdutividade.Update
TBProdutividade.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Sub ProcCriaColunas()
On Error GoTo tratar_erro

Lista1.ColumnHeaders.Clear
contador = 1
TamanhoColTotal = 1600
Select Case Cmb_tipo
    Case "A pagar": TextoColuna = "Vlr. total pagar"
    Case "A receber": TextoColuna = "Vlr. total receber"
    Case "A pagar e pagas":
        TextoColuna = "Vlr. total pagar  |  pago"
        TamanhoColTotal = 2100
    Case "Pagas": TextoColuna = "Vlr. total pago"
    Case "Recebidas": TextoColuna = "Vlr. total recebido"
    Case "A receber e recebidas":
        TextoColuna = "Vlr. total receber  |  receb."
        TamanhoColTotal = 2100
    Case "Recebidas e pagas":
        TextoColuna = "Vlr. total receb.  |  pagas"
        TamanhoColTotal = 2100
    Case "A receber e a pagar":
        TextoColuna = "Vlr. total receber  |  pagar"
        TamanhoColTotal = 2100
End Select

With Lista1.ColumnHeaders
    .Add
    .Item(contador).Text = cmbfiltrarpor.Text
    .Item(contador).Width = 3500
    If cmbPor.Text = "Dia" Then
        Dataini = msk_fltInicio
        DataFim = msk_fltFim
        Do While Dataini <= DataFim
            .Add
            contador = contador + 1
            .Item(contador).Text = Format(Dataini, "dd/mm/yy")
            .Item(contador).Alignment = lvwColumnRight
            Dataini = Dataini + 1
        Loop
    End If
    If cmbPor.Text = "Mês" Then
        qt = FunVerificaMes(Cmb_mes_de)
        Qtd = FunVerificaMes(Cmb_mes_ate)
        Do While qt <= Qtd
            .Add
            contador = contador + 1
            .Item(contador).Text = qt & "/" & Cmb_ano_de
            .Item(contador).Alignment = lvwColumnRight
            qt = qt + 1
        Loop
    End If
    If cmbPor.Text = "Ano" Then
        qt = Cmb_ano_de1
        Do While qt <= Cmb_ano_ate1
            .Add
            contador = contador + 1
            .Item(contador).Text = qt
            .Item(contador).Alignment = lvwColumnRight
            qt = qt + 1
        Loop
    End If
    .Add
    contador = contador + 1
    .Item(contador).Text = TextoColuna
    .Item(contador).Width = TamanhoColTotal
    .Item(contador).Alignment = lvwColumnRight
End With
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcGravarTotalizacoes()
On Error GoTo tratar_erro

If optResumido.Value = True Then
    valor = 0
    Valor1 = 0
    Valor2 = 0
    
    If optEmisao.Value = True Then
        If Cmb_tipo = "A pagar" Or Cmb_tipo = "Pagas" Or Cmb_tipo = "A pagar e pagas" Then
            Data_pesquisa = "dt_Emissao"
        Else
            Data_pesquisa = "Emissao"
        End If
    ElseIf optVencimento.Value = True Then
        If Cmb_tipo = "A pagar" Or Cmb_tipo = "Pagas" Or Cmb_tipo = "A pagar e pagas" Then
            Data_pesquisa = "dt_Pagamento"
        Else
            Data_pesquisa = "Vencimento"
        End If
    Else
        If Cmb_tipo = "A pagar" Or Cmb_tipo = "Pagas" Or Cmb_tipo = "A pagar e pagas" Then
            Data_pesquisa = "DataBaixa"
        Else
            Data_pesquisa = "Data_pagamento"
        End If
    End If
    
    If Cmb_tipo = "A receber" Or Cmb_tipo = "Recebidas" Or Cmb_tipo = "A receber e recebidas" Then
        TabelaFiltro = "Financeiro_relatorios_historico_detalhado_receber"
        TipoConta = " and FF.tipoconta = 'R'"
        filtrodestino = "R"
        filtrodestino1 = "R"
    ElseIf Cmb_tipo = "A receber e a pagar" Or Cmb_tipo = "Recebidas e pagas" Then
        TabelaFiltro = "Financeiro_relatorios_historico_detalhado"
        TipoConta = ""
        filtrodestino = "R"
        filtrodestino1 = "P"
    Else
        TabelaFiltro = "Financeiro_relatorios_historico_detalhado_pagar"
        TipoConta = " and FF.tipoconta = 'P'"
        filtrodestino = "P"
        filtrodestino1 = "P"
    End If
    
    If Cmb_tipo = "A receber e a pagar" Or Cmb_tipo = "Recebidas e pagas" Then
        If Cmb_tipo = "A receber e a pagar" Then
            FiltroResumido = " and Tipo = 'R' and Logsit = 'N'"
            FiltroResumido1 = " and Tipo = 'P' and Logsit = 'N'"
        Else
            FiltroResumido = " and Tipo = 'R' and Logsit = 'S'"
            FiltroResumido1 = " and Tipo = 'P' and Logsit = 'S'"
        End If
    Else
        FiltroResumido = " and LogSit = 'N'"
        FiltroResumido1 = " and LogSit = 'S'"
    End If
    
    Select Case cmbPor
        Case "Dia":
            If Opt_individual.Value = True Then
                If cmbfiltrarpor = "Conta contábil" Then
                    Data_pesquisa1 = FamiliaAntiga & " = " & cmbTexto.ItemData(cmbTexto.ListIndex) & " and " & Data_pesquisa & " Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "'"
                Else
                    Data_pesquisa1 = FamiliaAntiga & " = '" & cmbTexto & "' and " & Data_pesquisa & " Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "'"
                End If
            Else
                Data_pesquisa1 = Data_pesquisa & " Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "'"
            End If
        Case "Mês":
            MesX = FunVerificaMes(Cmb_mes_de)
            MesX1 = FunVerificaMes(Cmb_mes_ate)
            If Opt_individual.Value = True Then
                If cmbfiltrarpor = "Conta contábil" Then
                    Data_pesquisa1 = FamiliaAntiga & " = " & cmbTexto.ItemData(cmbTexto.ListIndex) & " and Month(" & Data_pesquisa & ") >= '" & MesX & "' And Year(" & Data_pesquisa & ") >= '" & Cmb_ano_de & "' and Month(" & Data_pesquisa & ") <= '" & MesX1 & "' And Year(" & Data_pesquisa & ") <= '" & Cmb_ano_ate & "'"
                Else
                    Data_pesquisa1 = FamiliaAntiga & " = '" & cmbTexto & "' and Month(" & Data_pesquisa & ") >= '" & MesX & "' And Year(" & Data_pesquisa & ") >= '" & Cmb_ano_de & "' and Month(" & Data_pesquisa & ") <= '" & MesX1 & "' And Year(" & Data_pesquisa & ") <= '" & Cmb_ano_ate & "'"
                End If
            Else
                Data_pesquisa1 = "Month(" & Data_pesquisa & ") >= '" & MesX & "' And Year(" & Data_pesquisa & ") >= '" & Cmb_ano_de & "' and Month(" & Data_pesquisa & ") <= '" & MesX1 & "' And Year(" & Data_pesquisa & ") <= '" & Cmb_ano_ate & "'"
            End If
        Case "Ano":
            If Opt_individual.Value = True Then
                If cmbfiltrarpor = "Conta contábil" Then
                    Data_pesquisa1 = FamiliaAntiga & " = " & cmbTexto.ItemData(cmbTexto.ListIndex) & " and Year(" & Data_pesquisa & ") >= '" & Cmb_ano_de1 & "' and Year(" & Data_pesquisa & ") <= '" & Cmb_ano_ate1 & "'"
                Else
                    Data_pesquisa1 = FamiliaAntiga & " = '" & cmbTexto & "' and Year(" & Data_pesquisa & ") >= '" & Cmb_ano_de1 & "' and Year(" & Data_pesquisa & ") <= '" & Cmb_ano_ate1 & "'"
                End If
            Else
                Data_pesquisa1 = "Year(" & Data_pesquisa & ") >= '" & Cmb_ano_de1 & "' and Year(" & Data_pesquisa & ") <= '" & Cmb_ano_ate1 & "'"
            End If
    End Select
    
    If cmbfiltrarpor = "Conta contábil" Then
        If Cmb_tipo = "A receber e a pagar" Or Cmb_tipo = "Recebidas e pagas" Then onamais = "and F.Tipo = FF.tipoconta" Else onamais = ""

        If Cmb_tipo = "A receber" Or Cmb_tipo = "A receber e recebidas" Or Cmb_tipo = "A receber e a pagar" Then
            'Total descontado
            Set TBproducao = CreateObject("adodb.recordset")
            TBproducao.Open "Select Sum(FF.Valor) as Valor from (" & TabelaFiltro & " F INNER JOIN familia_financeiro FF on F.IdIntConta = FF.idconta " & onamais & ") INNER JOIN tbl_familia TBL ON TBL.int_codfamilia = FF.ID_PC WHERE " & Data_pesquisa1 & FiltroResumido & " and FF.Deposito_transf <> 'True' and TBL.Destino = 'R' and F.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and F.Status = 'DUPLICATA DESCONTADA EM ABERTO' and F.bloqueado = 'False'", Conexao, adOpenKeyset, adLockReadOnly
            If TBproducao.EOF = False Then
                valor = IIf(IsNull(TBproducao!valor), 0, TBproducao!valor)
            End If
            TBproducao.Close
        End If
        
        If Cmb_tipo <> "Recebidas" And Cmb_tipo <> "Pagas" Then
            
            'Total a receber/ a pagar
            Set TBproducao = CreateObject("adodb.recordset")
            TBproducao.Open "Select Sum(FF.Valor) as Valor1 from (" & TabelaFiltro & " F inner join familia_financeiro FF on F.IdIntConta = FF.idconta " & onamais & ") INNER JOIN tbl_familia TBL ON TBL.int_codfamilia = FF.ID_PC WHERE " & Data_pesquisa1 & FiltroResumido & " and FF.Deposito_transf <> 'True' and TBL.Destino = '" & filtrodestino & "' and F.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and F.Status <> 'DUPLICATA DESCONTADA EM ABERTO' and F.Bloqueado = 'False'", Conexao, adOpenKeyset, adLockReadOnly
            If TBproducao.EOF = False Then
                Valor1 = IIf(IsNull(TBproducao!Valor1), 0, TBproducao!Valor1)
            End If
            TBproducao.Close
        End If
        
        If Cmb_tipo <> "A pagar" And Cmb_tipo <> "A receber" Then
            'Total recebido/pago
            Set TBproducao = CreateObject("adodb.recordset")
            TBproducao.Open "Select Sum(FF.Valor) as Valor2 from (" & TabelaFiltro & " F inner join familia_financeiro FF on F.IdIntConta = FF.idconta " & onamais & ") INNER JOIN tbl_familia TBL ON TBL.int_codfamilia = FF.ID_PC WHERE " & Data_pesquisa1 & FiltroResumido1 & " and FF.Deposito_transf <> 'True' and TBL.Destino = '" & filtrodestino1 & "' and F.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and F.Bloqueado = 'False'", Conexao, adOpenKeyset, adLockReadOnly
            If TBproducao.EOF = False Then
                Valor2 = IIf(IsNull(TBproducao!Valor2), 0, TBproducao!Valor2)
            End If
            TBproducao.Close
        End If
    Else
        If Cmb_tipo = "Pagas" Or Cmb_tipo = "A pagar" Or Cmb_tipo = "A pagar e pagas" Then
            filtrovalor = "dbl_valorpagto"
            filtrovalor2 = "Valorpago"
        Else
            If Cmb_tipo = "Recebidas e pagas" Then
                filtrovalor = "vlrBaixa"
                filtrovalor2 = "vlrBaixa"
            ElseIf Cmb_tipo = "A receber e a pagar" Then
                filtrovalor = "valor"
                filtrovalor2 = "valor"
            Else
                filtrovalor = "valor"
                filtrovalor2 = "valortitulorecebido"
            End If
        End If
    
        If Cmb_tipo = "A receber" Or Cmb_tipo = "A receber e recebidas" Or Cmb_tipo = "A receber e a pagar" Then
            'Total descontado
            Set TBproducao = CreateObject("adodb.recordset")
            TBproducao.Open "Select Sum(" & filtrovalor & ") as Valor from " & TabelaFiltro & " where " & Data_pesquisa1 & " and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and Status = 'DUPLICATA DESCONTADA EM ABERTO' and Bloqueado = 'False' and " & FamiliaAntiga & " IS NOT NULL and " & FamiliaAntiga & " <> ''", Conexao, adOpenKeyset, adLockOptimistic
            If TBproducao.EOF = False Then
                valor = IIf(IsNull(TBproducao!valor), 0, TBproducao!valor)
            End If
            TBproducao.Close
        End If
            
        If Cmb_tipo <> "Recebidas" And Cmb_tipo <> "Pagas" Then
            'Total a receber/a pagar
            'Quando for receber e pagar aqui e só receber
            Set TBproducao = CreateObject("adodb.recordset")
            TBproducao.Open "Select Sum(" & filtrovalor & ") as Valor1 from " & TabelaFiltro & " where " & Data_pesquisa1 & FiltroResumido & " and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and Status <> 'DUPLICATA DESCONTADA EM ABERTO' and Bloqueado = 'False' and " & FamiliaAntiga & " IS NOT NULL and " & FamiliaAntiga & " <> ''", Conexao, adOpenKeyset, adLockOptimistic
            If TBproducao.EOF = False Then
                Valor1 = IIf(IsNull(TBproducao!Valor1), 0, TBproducao!Valor1)
            End If
            TBproducao.Close
        End If
            
        If Cmb_tipo <> "A pagar" And Cmb_tipo <> "A receber" Then
            'Total recebido/pago
            'Quando for receber e pagar aqui e só pagar
            Set TBproducao = CreateObject("adodb.recordset")
            TBproducao.Open "Select Sum(" & filtrovalor2 & ") as Valor2 from " & TabelaFiltro & " where " & Data_pesquisa1 & FiltroResumido1 & " and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and Bloqueado = 'False' and " & FamiliaAntiga & " IS NOT NULL and " & FamiliaAntiga & " <> ''", Conexao, adOpenKeyset, adLockOptimistic
            If TBproducao.EOF = False Then
                Valor2 = IIf(IsNull(TBproducao!Valor2), 0, TBproducao!Valor2)
            End If
            TBproducao.Close
        End If
    End If
End If
    
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Producao_Relatorios_Total where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = True Then TBAbrir.AddNew
Select Case cmbPor
    Case "Dia":
        Tipo = "D"
        TBAbrir!Totalutilizada = Format(msk_fltInicio.Value, "dd/mm/yy")
        TBAbrir!Totalprevista = Format(msk_fltFim.Value, "dd/mm/yy")
    Case "Mês":
        Tipo = "M"
        TBAbrir!Totalutilizada = Cmb_mes_de & "/" & Cmb_ano_de
        TBAbrir!Totalprevista = Cmb_mes_ate & "/" & Cmb_ano_ate
    Case "Ano":
        Tipo = "A"
        TBAbrir!Totalutilizada = Cmb_ano_de1
        TBAbrir!Totalprevista = Cmb_ano_ate1
End Select

If Opt_individual.Value = True Then TBAbrir!Texto = cmbfiltrarpor & " : " & cmbTexto Else TBAbrir!Texto = cmbfiltrarpor

TBAbrir!Texto1 = Tipo
TBAbrir!Responsavel = pubUsuario
TBAbrir!Modulo = Formulario
TBAbrir!QtdeProduzida = valor 'Total descontado
TBAbrir!QtdePrevista = Valor1 'Total a receber / pagar
TBAbrir!qtdeNC = Valor2 'Total recebido / pago
TBAbrir.Update
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub msk_fltInicio_Click()
On Error GoTo tratar_erro

ProcLimpaListaeCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Opt_comparativo_Click()
On Error GoTo tratar_erro

ProcLimpaListaeCampos
If Opt_comparativo.Value = True Then
    optDetalhado.Enabled = False
    optResumido.Value = True
    cmbTexto.ListIndex = -1
    cmbTexto.Enabled = False
    If cmbfiltrarpor = "Conta contábil" Then Chk_mostrar_todasCC.Visible = True
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Opt_individual_Click()
On Error GoTo tratar_erro

If Opt_individual.Value = True Then
    optDetalhado.Value = True
    optDetalhado.Enabled = True
    cmbTexto.Enabled = True
    ProcCarregaComboTexto
    Chk_mostrar_todasCC.Visible = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub optDetalhado_Click()
On Error GoTo tratar_erro

If optDetalhado.Value = True Then
    ProcLimpaListaeCampos
    Lista.Visible = True
    Lista1.Visible = False
    With cmbPor
        .Clear
        .AddItem "Dia"
        .Text = "Dia"
    End With
    ProcMostrarEsconderCombosData
    ProcLimpaCamposTotais
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub optEmisao_Click()
On Error GoTo tratar_erro

ProcLimpaListaeCampos
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub optPgto_receb_Click()
On Error GoTo tratar_erro

ProcLimpaListaeCampos
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub optResumido_Click()
On Error GoTo tratar_erro

If optResumido.Value = True Then
    ProcLimpaListaeCampos
    Lista.Visible = False
    Lista1.Visible = True
    With cmbPor
        .Clear
        .AddItem "Dia"
        .AddItem "Mês"
        .AddItem "Ano"
        .Text = "Dia"
    End With
    ProcMostrarEsconderCombosData
    ProcLimpaCamposTotais
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcMostrarEsconderCombosData()
On Error GoTo tratar_erro

If cmbPor = "Dia" Then
    msk_fltInicio.Visible = True
    msk_fltFim.Visible = True
    Cmb_mes_de.Visible = False
    Cmb_mes_ate.Visible = False
    Cmb_ano_de.Visible = False
    Cmb_ano_ate.Visible = False
    Cmb_ano_de1.Visible = False
    Cmb_ano_ate1.Visible = False
ElseIf cmbPor = "Mês" Then
        msk_fltInicio.Visible = False
        msk_fltFim.Visible = False
        Cmb_mes_de.Visible = True
        Cmb_mes_ate.Visible = True
        Cmb_ano_de.Visible = True
        Cmb_ano_ate.Visible = True
        Cmb_ano_de1.Visible = False
        Cmb_ano_ate1.Visible = False
    Else
        msk_fltInicio.Visible = False
        msk_fltFim.Visible = False
        Cmb_mes_de.Visible = False
        Cmb_mes_ate.Visible = False
        Cmb_ano_de.Visible = False
        Cmb_ano_ate.Visible = False
        Cmb_ano_de1.Visible = True
        Cmb_ano_ate1.Visible = True
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Sub ProcCarregaFiltrarpor()
On Error GoTo tratar_erro

cmbTexto.Clear
With cmbfiltrarpor
    .Clear
    If Cmb_tipo = "A receber" Or Cmb_tipo = "Recebidas" Or Cmb_tipo = "A receber e recebidas" Then
        .AddItem "Cliente"
        .AddItem "Nota fiscal"
        .Text = "Cliente"
    ElseIf Cmb_tipo = "A pagar" Or Cmb_tipo = "Pagas" Or Cmb_tipo = "A pagar e pagas" Then
        .AddItem "Fornecedor"
        If Cmb_tipo = "Pagas" Or Cmb_tipo = "A pagar e pagas" Then .AddItem "Forma de pagamento"
        .Text = "Fornecedor"
    Else
        .AddItem "Cliente/Fornecedor"
        .Text = "Cliente/Fornecedor"
    End If
    .AddItem "Documento"
    .AddItem "Instituição"
    .AddItem "Conta contábil"
    .AddItem "Tipo do doc."
    If Cmb_tipo = "Pagas" Or Cmb_tipo = "A pagar e pagas" Or Cmb_tipo = "Recebidas" Or Cmb_tipo = "A receber e recebidas" Or Cmb_tipo = "Recebidas e pagas" Then .AddItem "Documento baixa"
    If Cmb_tipo <> "A pagar e pagas" And Cmb_tipo <> "A receber e recebidas" And Cmb_tipo <> "Recebidas e pagas" And Cmb_tipo <> "A receber e a pagar" Then .AddItem "Status"
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Sub ProcArrumaLista()
On Error GoTo tratar_erro

With Lista
    .ColumnHeaders(9).Width = 0
    .ColumnHeaders(10).Width = 0
    .ColumnHeaders(11).Width = 0
    .ColumnHeaders(12).Width = 0
    If Cmb_tipo = "A pagar" Or Cmb_tipo = "A receber" Then
        If Cmb_tipo = "A pagar" Then
            .ColumnHeaders(6).Width = 0
            .ColumnHeaders(8).Text = "Fornecedor"
            .ColumnHeaders(8).Width = 8555
        Else
            .ColumnHeaders(6).Width = 1200
            .ColumnHeaders(8).Text = "Cliente"
            .ColumnHeaders(8).Width = 7355
        End If
    ElseIf Cmb_tipo = "A receber e a pagar" Or Cmb_tipo = "Recebidas e pagas" Then
        .ColumnHeaders(6).Width = 1200
        .ColumnHeaders(8).Text = "Cliente/Fornecedor"
        If Cmb_tipo = "Recebidas e pagas" Then
            .ColumnHeaders(8).Width = 2855
            .ColumnHeaders(9).Width = 1200
            .ColumnHeaders(10).Width = 1200
            .ColumnHeaders(11).Width = 1200
        Else
            .ColumnHeaders(8).Width = 6455
        End If
        .ColumnHeaders(12).Width = 900
    Else
        If Cmb_tipo = "Pagas" Or Cmb_tipo = "A pagar e pagas" Then
            .ColumnHeaders(6).Width = 0
            .ColumnHeaders(8).Text = "Fornecedor"
            .ColumnHeaders(8).Width = 4955
        Else
            .ColumnHeaders(6).Width = 1200
            .ColumnHeaders(8).Text = "Cliente"
            .ColumnHeaders(8).Width = 3755
        End If
        .ColumnHeaders(9).Width = 1200
        .ColumnHeaders(10).Width = 1200
        .ColumnHeaders(11).Width = 1200
    End If
End With
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Sub ProcOrdenaTudo()
On Error GoTo tratar_erro

ProcLimpaListaeCampos
Select Case Cmb_tipo
    Case "A pagar":
        optPgto_receb.Enabled = False
        optEmisao.Value = True
        Frm_pagar_pago_recebido.Visible = True
        Frm_pagarxpago.Visible = False
        Frm_receber.Visible = False
        Frm_receberxrecebido.Visible = False
        frm_Recebidoxpago.Visible = False
        frm_Receberxpagar.Visible = False
    Case "A receber":
        optPgto_receb.Enabled = False
        optEmisao.Value = True
        Frm_pagar_pago_recebido.Visible = False
        Frm_pagarxpago.Visible = False
        Frm_receber.Visible = True
        Frm_receberxrecebido.Visible = False
        frm_Recebidoxpago.Visible = False
        frm_Receberxpagar.Visible = False
    Case "A pagar e pagas":
        If optPgto_receb.Value = True Then optEmisao.Value = True
        optPgto_receb.Enabled = False
        Frm_pagar_pago_recebido.Visible = False
        Frm_pagarxpago.Visible = True
        Frm_receber.Visible = False
        Frm_receberxrecebido.Visible = False
        frm_Recebidoxpago.Visible = False
        frm_Receberxpagar.Visible = False
    Case "Pagas":
        optPgto_receb.Enabled = True
        Frm_pagar_pago_recebido.Visible = True
        Frm_pagarxpago.Visible = False
        Frm_receber.Visible = False
        Frm_receberxrecebido.Visible = False
        frm_Recebidoxpago.Visible = False
        frm_Receberxpagar.Visible = False
    Case "Recebidas":
        optPgto_receb.Enabled = True
        Frm_pagar_pago_recebido.Visible = True
        Frm_pagarxpago.Visible = False
        Frm_receber.Visible = False
        Frm_receberxrecebido.Visible = False
        frm_Recebidoxpago.Visible = False
        frm_Receberxpagar.Visible = False
    Case "A receber e recebidas":
        If optPgto_receb.Value = True Then optEmisao.Value = True
        optPgto_receb.Enabled = False
        Frm_pagar_pago_recebido.Visible = False
        Frm_pagarxpago.Visible = False
        Frm_receber.Visible = False
        Frm_receberxrecebido.Visible = True
        frm_Recebidoxpago.Visible = False
        frm_Receberxpagar.Visible = False
    Case "A receber e a pagar":
        If optPgto_receb.Value = True Then optEmisao.Value = True
        optPgto_receb.Enabled = False
        Frm_pagar_pago_recebido.Visible = False
        Frm_pagarxpago.Visible = False
        Frm_receber.Visible = False
        Frm_receberxrecebido.Visible = False
        frm_Recebidoxpago.Visible = False
        frm_Receberxpagar.Visible = True
    Case "Recebidas e pagas":
        If optPgto_receb.Value = True Then optEmisao.Value = True
        optPgto_receb.Enabled = True
        Frm_pagar_pago_recebido.Visible = False
        Frm_pagarxpago.Visible = False
        Frm_receber.Visible = False
        Frm_receberxrecebido.Visible = False
        frm_Recebidoxpago.Visible = True
        frm_Receberxpagar.Visible = False
End Select
ProcCarregaFiltrarpor
ProcLimpaCamposTotais
ProcArrumaLista
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub optVencimento_Click()
On Error GoTo tratar_erro

ProcLimpaListaeCampos
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Sub ProcVerifPC()
On Error GoTo tratar_erro

Do While TBFamilia.EOF = False
    'Verifica o código e o nível do PC
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from tbl_familia where int_codfamilia = " & IIf(IsNull(TBFamilia!ID_PC), 0, TBFamilia!ID_PC) & " and Codigo is not null and nivel is not null", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Select Case TBAbrir!Nivel
            Case 8:
                Set TBNivel8 = CreateObject("adodb.recordset")
                TBNivel8.Open "Select * from tbl_familia where int_codfamilia = " & TBFamilia!ID_PC & " and Codigo is not null and nivel is not null", Conexao, adOpenKeyset, adLockOptimistic
                If TBNivel8.EOF = False Then
                    
                    Set TBGravar = CreateObject("adodb.recordset")
                    TBGravar.Open "Select * from Plano_de_contas_totalizacao where Codigo = '" & TBNivel8!CODIGO & "' and Nivel = " & TBNivel8!Nivel & " and Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
                    ProcEnviaDadosPCRelFinanc TBNivel8!int_codfamilia, TBNivel8!CODIGO, TBNivel8!Txt_descricao, TBNivel8!Nivel
                    
                    If Chk_mostrar_todasCC.Value = 1 Then ProcNivelPC7
                End If
                TBNivel8.Close
            Case 7:
                Set TBNivel7 = CreateObject("adodb.recordset")
                TBNivel7.Open "Select * from tbl_familia where int_codfamilia = " & TBFamilia!ID_PC & " and Codigo is not null and nivel is not null", Conexao, adOpenKeyset, adLockOptimistic
                If TBNivel7.EOF = False Then
                    
                    Set TBGravar = CreateObject("adodb.recordset")
                    TBGravar.Open "Select * from Plano_de_contas_totalizacao where Codigo = '" & TBNivel7!CODIGO & "' and Nivel = " & TBNivel7!Nivel & " and Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
                    ProcEnviaDadosPCRelFinanc TBNivel7!int_codfamilia, TBNivel7!CODIGO, TBNivel7!Txt_descricao, TBNivel7!Nivel
                    
                    If Chk_mostrar_todasCC.Value = 1 Then ProcNivelPC6
                End If
                TBNivel7.Close
            Case 6:
                Set TBNivel6 = CreateObject("adodb.recordset")
                TBNivel6.Open "Select * from tbl_familia where int_codfamilia = " & TBFamilia!ID_PC & " and Codigo is not null and nivel is not null", Conexao, adOpenKeyset, adLockOptimistic
                If TBNivel6.EOF = False Then
                    
                    Set TBGravar = CreateObject("adodb.recordset")
                    TBGravar.Open "Select * from Plano_de_contas_totalizacao where Codigo = '" & TBNivel6!CODIGO & "' and Nivel = " & TBNivel6!Nivel & " and Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
                    ProcEnviaDadosPCRelFinanc TBNivel6!int_codfamilia, TBNivel6!CODIGO, TBNivel6!Txt_descricao, TBNivel6!Nivel
                    
                    If Chk_mostrar_todasCC.Value = 1 Then ProcNivelPC5
                End If
                TBNivel6.Close
            Case 5:
                Set TBNivel5 = CreateObject("adodb.recordset")
                TBNivel5.Open "Select * from tbl_familia where int_codfamilia = " & TBFamilia!ID_PC & " and Codigo is not null and nivel is not null", Conexao, adOpenKeyset, adLockOptimistic
                If TBNivel5.EOF = False Then
                    
                    Set TBGravar = CreateObject("adodb.recordset")
                    TBGravar.Open "Select * from Plano_de_contas_totalizacao where Codigo = '" & TBNivel5!CODIGO & "' and Nivel = " & TBNivel5!Nivel & " and Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
                    ProcEnviaDadosPCRelFinanc TBNivel5!int_codfamilia, TBNivel5!CODIGO, TBNivel5!Txt_descricao, TBNivel5!Nivel
                    
                    If Chk_mostrar_todasCC.Value = 1 Then ProcNivelPC4
                End If
                TBNivel5.Close
            Case 4:
                Set TBNivel4 = CreateObject("adodb.recordset")
                TBNivel4.Open "Select * from tbl_familia where int_codfamilia = " & TBFamilia!ID_PC & " and Codigo is not null and nivel is not null", Conexao, adOpenKeyset, adLockOptimistic
                If TBNivel4.EOF = False Then
                    
                    Set TBGravar = CreateObject("adodb.recordset")
                    TBGravar.Open "Select * from Plano_de_contas_totalizacao where Codigo = '" & TBNivel4!CODIGO & "' and Nivel = " & TBNivel4!Nivel & " and Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
                    ProcEnviaDadosPCRelFinanc TBNivel4!int_codfamilia, TBNivel4!CODIGO, TBNivel4!Txt_descricao, TBNivel4!Nivel
                    
                    If Chk_mostrar_todasCC.Value = 1 Then ProcNivelPC3
                End If
                TBNivel4.Close
            Case 3:
                Set TBNivel3 = CreateObject("adodb.recordset")
                TBNivel3.Open "Select * from tbl_familia where int_codfamilia = " & TBFamilia!ID_PC & " and Codigo is not null and nivel is not null", Conexao, adOpenKeyset, adLockOptimistic
                If TBNivel3.EOF = False Then
                    
                    Set TBGravar = CreateObject("adodb.recordset")
                    TBGravar.Open "Select * from Plano_de_contas_totalizacao where Codigo = '" & TBNivel3!CODIGO & "' and Nivel = " & TBNivel3!Nivel & " and Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
                    ProcEnviaDadosPCRelFinanc TBNivel3!int_codfamilia, TBNivel3!CODIGO, TBNivel3!Txt_descricao, TBNivel3!Nivel
                    
                    If Chk_mostrar_todasCC.Value = 1 Then ProcNivelPC2
                End If
                TBNivel3.Close
            Case 2:
                Set TBNivel2 = CreateObject("adodb.recordset")
                TBNivel2.Open "Select * from tbl_familia where int_codfamilia = " & TBFamilia!ID_PC & " and Codigo is not null and nivel is not null", Conexao, adOpenKeyset, adLockOptimistic
                If TBNivel2.EOF = False Then
                    
                    Set TBGravar = CreateObject("adodb.recordset")
                    TBGravar.Open "Select * from Plano_de_contas_totalizacao where Codigo = '" & TBNivel2!CODIGO & "' and Nivel = " & TBNivel2!Nivel & " and Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
                    ProcEnviaDadosPCRelFinanc TBNivel2!int_codfamilia, TBNivel2!CODIGO, TBNivel2!Txt_descricao, TBNivel2!Nivel
                    
                    If Chk_mostrar_todasCC.Value = 1 Then ProcNivelPC1
                End If
                TBNivel2.Close
            Case 1:
                Set TBNivel1 = CreateObject("adodb.recordset")
                TBNivel1.Open "Select * from tbl_familia where int_codfamilia = " & TBFamilia!ID_PC & " and Codigo is not null and nivel is not null", Conexao, adOpenKeyset, adLockOptimistic
                If TBNivel1.EOF = False Then
                    
                    Set TBGravar = CreateObject("adodb.recordset")
                    TBGravar.Open "Select * from Plano_de_contas_totalizacao where Codigo = '" & TBNivel1!CODIGO & "' and Nivel = " & TBNivel1!Nivel & " and Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
                    ProcEnviaDadosPCRelFinanc TBNivel1!int_codfamilia, TBNivel1!CODIGO, TBNivel1!Txt_descricao, TBNivel1!Nivel
                    
                End If
                TBNivel1.Close
        End Select
    End If
    TBAbrir.Close
    TBFamilia.MoveNext
Loop

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcFiltrar
    Case 2: ProcImprimir
    Case 4: ProcAjuda
    Case 5: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub
