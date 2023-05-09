VERSION 5.00
Object = "{8C1279ED-044C-4258-A3E3-0D5514B899FC}#1.44#0"; "ControlesUteis.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFaturamento_CartaCorrecao_NS 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Faturamento - Carta de correção"
   ClientHeight    =   10035
   ClientLeft      =   75
   ClientTop       =   435
   ClientWidth     =   15360
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10035
   ScaleWidth      =   15360
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Certificado Digital"
      Height          =   1575
      Left            =   11640
      TabIndex        =   47
      Top             =   2400
      Width           =   3705
      Begin VB.TextBox txtval 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
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
         Left            =   3090
         Locked          =   -1  'True
         MaxLength       =   60
         TabIndex        =   50
         TabStop         =   0   'False
         ToolTipText     =   "Serial Certificado"
         Top             =   450
         Width           =   495
      End
      Begin VB.TextBox txtSerialCertificado 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
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
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   60
         TabIndex        =   48
         TabStop         =   0   'False
         ToolTipText     =   "Serial Certificado"
         Top             =   450
         Width           =   2955
      End
      Begin DrawSuite2022.USButton btnCertificado 
         Height          =   615
         Left            =   90
         TabIndex        =   52
         ToolTipText     =   "Verificar certificado digital"
         Top             =   870
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   1085
         DibPicture      =   "frmFaturamento_CartaCorrecao_NS.frx":0000
         Caption         =   "   Verificar certificado digital"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   5263559
         BorderColorDisabled=   13160660
         BorderColorDown =   4013465
         BorderColorOver =   4408288
         GradientColor1  =   5263559
         GradientColor2  =   5263559
         GradientColor3  =   5263559
         GradientColor4  =   5263559
         GradientColorDisabled1=   13160660
         GradientColorDisabled2=   13160660
         GradientColorDisabled3=   13160660
         GradientColorDisabled4=   13160660
         GradientColorOver1=   4408288
         GradientColorOver2=   4408288
         GradientColorOver3=   4408288
         GradientColorOver4=   4408288
         GradientColorDown1=   4013465
         GradientColorDown2=   4013465
         GradientColorDown3=   4013465
         GradientColorDown4=   4013465
         PicSize         =   5
         PicSizeH        =   32
         PicSizeW        =   32
         Theme           =   4
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Val."
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   3210
         TabIndex        =   51
         Top             =   240
         Width           =   285
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Serial certificado"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   990
         TabIndex        =   49
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Retorno SEFAZ"
      ForeColor       =   &H00800000&
      Height          =   1575
      Left            =   5790
      TabIndex        =   41
      Top             =   2400
      Width           =   5835
      Begin ControlesUteis.txtA txtRetorno 
         Height          =   1035
         Left            =   90
         TabIndex        =   42
         ToolTipText     =   "Digite a descrição das correções a serem feitas na nota fiscal."
         Top             =   270
         Width           =   5625
         _ExtentX        =   9922
         _ExtentY        =   1826
         Text            =   ""
         ShowCounter     =   0   'False
         Caption         =   ""
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483640
      End
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   10980
      Top             =   240
   End
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   10080
      Top             =   180
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   26
      ScreenHeight    =   1080
      ScreenWidth     =   1920
      ScreenHeightDT  =   1080
      ScreenWidthDT   =   1920
      AutoResizeOnLoad=   0   'False
      ApplicationName =   "Active Resize Control Professional"
      FormHeightDT    =   10500
      FormWidthDT     =   15480
      FormScaleHeightDT=   10035
      FormScaleWidthDT=   15360
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin VB.Timer Timer_status_CCe 
      Interval        =   10000
      Left            =   7200
      Top             =   6780
   End
   Begin VB.TextBox Txt_ID_nota 
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
      Left            =   8340
      Locked          =   -1  'True
      MouseIcon       =   "frmFaturamento_CartaCorrecao_NS.frx":23FD
      TabIndex        =   39
      TabStop         =   0   'False
      ToolTipText     =   "ID da NF"
      Top             =   6870
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Frame Frame5 
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
      Height          =   1395
      Left            =   9900
      TabIndex        =   36
      Top             =   990
      Width           =   5445
      Begin VB.TextBox txtnsNrec 
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
         Left            =   4080
         MaxLength       =   60
         TabIndex        =   45
         TabStop         =   0   'False
         ToolTipText     =   "Status NFe."
         Top             =   390
         Width           =   915
      End
      Begin VB.TextBox txtStatus 
         Alignment       =   2  'Center
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
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   60
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Status NFe."
         Top             =   390
         Width           =   3615
      End
      Begin VB.TextBox Txt_chave_acesso 
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
         Height          =   310
         Left            =   120
         MaxLength       =   60
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
         ToolTipText     =   "Chave de acesso NFe."
         Top             =   930
         Width           =   4875
      End
      Begin DrawSuite2022.USButton BTNnsnrec 
         Height          =   315
         Left            =   5010
         TabIndex        =   53
         ToolTipText     =   "Consultar recibo nsNRec no SEFAZ com chave de acesso."
         Top             =   390
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         DibPicture      =   "frmFaturamento_CartaCorrecao_NS.frx":2707
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
         BorderColor     =   8421504
         BorderColorDisabled=   13160660
         BorderColorDown =   7907521
         BorderColorOver =   7907521
         GradientColor2  =   14737632
         GradientColor3  =   12632256
         GradientColor4  =   12632256
         GradientColorDisabled1=   14215660
         GradientColorDisabled2=   14215660
         GradientColorDisabled3=   14215660
         GradientColorDisabled4=   14215660
         GradientColorOver1=   14417407
         GradientColorOver2=   12317439
         GradientColorOver3=   4838399
         GradientColorOver4=   9627391
         GradientColorDown1=   10802943
         GradientColorDown2=   7979263
         GradientColorDown3=   4370174
         GradientColorDown4=   7395582
         GradientColors  =   1
         PicAlign        =   0
         Theme           =   1
         ToolTipTitle    =   "CAPRIND v5.0"
      End
      Begin DrawSuite2022.USButton cmdConsultar 
         Height          =   315
         Left            =   5010
         TabIndex        =   54
         ToolTipText     =   "Consultar chave de acesso no SEFAZ."
         Top             =   930
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         DibPicture      =   "frmFaturamento_CartaCorrecao_NS.frx":989A
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
         BorderColor     =   8421504
         BorderColorDisabled=   13160660
         BorderColorDown =   7907521
         BorderColorOver =   7907521
         GradientColor2  =   14737632
         GradientColor3  =   12632256
         GradientColor4  =   12632256
         GradientColorDisabled1=   14215660
         GradientColorDisabled2=   14215660
         GradientColorDisabled3=   14215660
         GradientColorDisabled4=   14215660
         GradientColorOver1=   14417407
         GradientColorOver2=   12317439
         GradientColorOver3=   4838399
         GradientColorOver4=   9627391
         GradientColorDown1=   10802943
         GradientColorDown2=   7979263
         GradientColorDown3=   4370174
         GradientColorDown4=   7395582
         GradientColors  =   1
         PicAlign        =   0
         Theme           =   1
         ToolTipTitle    =   "CAPRIND v5.0"
      End
      Begin DrawSuite2022.USButton cmdStatus 
         Height          =   315
         Left            =   3750
         TabIndex        =   55
         ToolTipText     =   "Consultar Status no SEFAZ."
         Top             =   390
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         DibPicture      =   "frmFaturamento_CartaCorrecao_NS.frx":10A2D
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
         BorderColor     =   8421504
         BorderColorDisabled=   13160660
         BorderColorDown =   7907521
         BorderColorOver =   7907521
         GradientColor2  =   14737632
         GradientColor3  =   12632256
         GradientColor4  =   12632256
         GradientColorDisabled1=   14215660
         GradientColorDisabled2=   14215660
         GradientColorDisabled3=   14215660
         GradientColorDisabled4=   14215660
         GradientColorOver1=   14417407
         GradientColorOver2=   12317439
         GradientColorOver3=   4838399
         GradientColorOver4=   9627391
         GradientColorDown1=   10802943
         GradientColorDown2=   7979263
         GradientColorDown3=   4370174
         GradientColorDown4=   7395582
         GradientColors  =   1
         PicAlign        =   0
         Theme           =   1
         ToolTipTitle    =   "CAPRIND v5.0"
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Recibo (NS)"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   4125
         TabIndex        =   46
         Top             =   180
         Width           =   840
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   1710
         TabIndex        =   38
         Top             =   180
         Width           =   465
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Chave de acesso"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   2010
         TabIndex        =   37
         Top             =   750
         Width           =   1230
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   55
      TabIndex        =   30
      Top             =   9120
      Width           =   15225
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
         Left            =   2730
         TabIndex        =   13
         Text            =   "30"
         ToolTipText     =   "Número de registros por página."
         Top             =   180
         Width           =   555
      End
      Begin VB.TextBox txtPagIr 
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
         Left            =   9540
         TabIndex        =   14
         ToolTipText     =   "Número da página."
         Top             =   180
         Width           =   555
      End
      Begin DrawSuite2022.USButton cmdPagProx 
         Height          =   315
         Left            =   11760
         TabIndex        =   18
         ToolTipText     =   "Próxima página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmFaturamento_CartaCorrecao_NS.frx":17BC0
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
         BorderColor     =   14404026
         BorderColorDown =   11632444
         BorderColorOver =   11632444
         GradientColor2  =   16777215
         GradientColor3  =   16777215
         GradientColorOver1=   16643560
         GradientColorOver2=   16576988
         GradientColorOver3=   16441780
         GradientColorOver4=   16178091
         GradientColorDown2=   16246986
         GradientColorDown3=   15189380
         GradientColorDown4=   14596208
         PicSizeH        =   19
         PicSizeW        =   19
      End
      Begin DrawSuite2022.USButton cmdPagAnt 
         Height          =   315
         Left            =   11220
         TabIndex        =   17
         ToolTipText     =   "Página anterior."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmFaturamento_CartaCorrecao_NS.frx":1B367
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
         BorderColor     =   14404026
         BorderColorDown =   11632444
         BorderColorOver =   11632444
         GradientColor2  =   16777215
         GradientColor3  =   16777215
         GradientColorOver1=   16643560
         GradientColorOver2=   16576988
         GradientColorOver3=   16441780
         GradientColorOver4=   16178091
         GradientColorDown2=   16246986
         GradientColorDown3=   15189380
         GradientColorDown4=   14596208
         PicSizeH        =   19
         PicSizeW        =   19
      End
      Begin DrawSuite2022.USButton cmdPagIr 
         Height          =   315
         Left            =   10110
         TabIndex        =   15
         Top             =   180
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   556
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
         BorderColor     =   14404026
         BorderColorDown =   11632444
         BorderColorOver =   11632444
         GradientColor2  =   16777215
         GradientColor3  =   16777215
         GradientColorOver1=   16643560
         GradientColorOver2=   16576988
         GradientColorOver3=   16441780
         GradientColorOver4=   16178091
         GradientColorDown2=   16246986
         GradientColorDown3=   15189380
         GradientColorDown4=   14596208
         PicSizeH        =   19
         PicSizeW        =   19
      End
      Begin DrawSuite2022.USButton cmdPagPrim 
         Height          =   315
         Left            =   10680
         TabIndex        =   16
         ToolTipText     =   "Primeira página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmFaturamento_CartaCorrecao_NS.frx":1EE75
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
         BorderColor     =   14404026
         BorderColorDown =   11632444
         BorderColorOver =   11632444
         GradientColor2  =   16777215
         GradientColor3  =   16777215
         GradientColorOver1=   16643560
         GradientColorOver2=   16576988
         GradientColorOver3=   16441780
         GradientColorOver4=   16178091
         GradientColorDown2=   16246986
         GradientColorDown3=   15189380
         GradientColorDown4=   14596208
         PicSizeH        =   19
         PicSizeW        =   19
      End
      Begin DrawSuite2022.USButton cmdPagUlt 
         Height          =   315
         Left            =   12300
         TabIndex        =   19
         ToolTipText     =   "Última página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmFaturamento_CartaCorrecao_NS.frx":22F69
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
         BorderColor     =   14404026
         BorderColorDown =   11632444
         BorderColorOver =   11632444
         GradientColor2  =   16777215
         GradientColor3  =   16777215
         GradientColorOver1=   16643560
         GradientColorOver2=   16576988
         GradientColorOver3=   16441780
         GradientColorOver4=   16178091
         GradientColorDown2=   16246986
         GradientColorDown3=   15189380
         GradientColorDown4=   14596208
         PicSizeH        =   19
         PicSizeW        =   19
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "registros por página"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3360
         TabIndex        =   40
         Top             =   240
         Width           =   1440
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Carregar"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2040
         TabIndex        =   33
         Top             =   240
         Width           =   645
      End
      Begin VB.Label lblRegistros 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nº de registros: 0"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   180
         TabIndex        =   32
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label lblPaginas 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Página: 0 de: 0"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   13050
         TabIndex        =   31
         Top             =   240
         Width           =   1095
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
      Height          =   1395
      Left            =   60
      TabIndex        =   20
      Top             =   990
      Width           =   9825
      Begin VB.TextBox txtdatacancelamento 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   5340
         TabIndex        =   56
         Text            =   "01/01/2020"
         Top             =   390
         Width           =   2145
      End
      Begin VB.TextBox txtSeq 
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
         TabIndex        =   43
         TabStop         =   0   'False
         Text            =   "1"
         ToolTipText     =   "Série."
         Top             =   945
         Width           =   495
      End
      Begin VB.TextBox Txt_destinatario 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3330
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Nome do destinatário."
         Top             =   950
         Width           =   6375
      End
      Begin VB.TextBox Txt_serie 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1770
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Série."
         Top             =   950
         Width           =   645
      End
      Begin VB.ComboBox Cmb_empresa 
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
         ItemData        =   "frmFaturamento_CartaCorrecao_NS.frx":267F6
         Left            =   180
         List            =   "frmFaturamento_CartaCorrecao_NS.frx":267F8
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "Empresa."
         Top             =   390
         Width           =   3180
      End
      Begin VB.TextBox txtResponsavel 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   7500
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Responsável."
         Top             =   390
         Width           =   2175
      End
      Begin VB.TextBox txtNF 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   690
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Número da nota fiscal."
         Top             =   950
         Width           =   1065
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
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
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Código."
         Top             =   390
         Width           =   795
      End
      Begin VB.TextBox txtiddestinatario 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2430
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "ID do destinatário."
         Top             =   950
         Width           =   885
      End
      Begin MSComCtl2.DTPicker txtdataemissao 
         Height          =   315
         Left            =   4170
         TabIndex        =   3
         ToolTipText     =   "Data de emissão."
         Top             =   390
         Width           =   1185
         _ExtentX        =   2090
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
         Format          =   197459969
         CurrentDate     =   39057
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data evento"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5962
         TabIndex        =   57
         Top             =   180
         Width           =   900
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Seq"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   315
         TabIndex        =   44
         Top             =   750
         Width           =   270
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Empresa"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1463
         TabIndex        =   35
         Top             =   180
         Width           =   615
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ID"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2790
         TabIndex        =   28
         Top             =   750
         Width           =   165
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Responsável"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   8130
         TabIndex        =   27
         Top             =   180
         Width           =   915
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N° Nota fiscal"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   750
         TabIndex        =   26
         Top             =   750
         Width           =   975
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Destinatário"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   6082
         TabIndex        =   24
         Top             =   750
         Width           =   870
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3510
         TabIndex        =   23
         Top             =   180
         Width           =   495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data emissão"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4275
         TabIndex        =   22
         Top             =   180
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Série"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1912
         TabIndex        =   21
         Top             =   750
         Width           =   360
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Descrição das correções com no mínimo 15 caracteres. "
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   1575
      Left            =   60
      TabIndex        =   25
      Top             =   2400
      Width           =   5715
      Begin ControlesUteis.txtA txtObs 
         Height          =   1005
         Left            =   150
         TabIndex        =   0
         ToolTipText     =   "Digite a descrição das correções a serem feitas na nota fiscal."
         Top             =   300
         Width           =   5445
         _ExtentX        =   9604
         _ExtentY        =   1773
         Text            =   ""
         CaptionColor    =   8388608
         Caption         =   "Observações : Não usar tecla ""Enter"" e nem carateres especiais"
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483640
      End
      Begin VB.CheckBox Chk_desconsiderar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Desconsiderar valor nota no total faturado dos ultimos 12 meses"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   210
         TabIndex        =   11
         Top             =   1290
         Width           =   5385
      End
   End
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   13770
      Top             =   240
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frmFaturamento_CartaCorrecao_NS.frx":267FA
      Count           =   1
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   29
      Top             =   0
      Width           =   15285
      _ExtentX        =   26961
      _ExtentY        =   1720
      ButtonCount     =   12
      GradientColor2  =   14737632
      GradientColorOverRight1=   16315633
      GradientColorOverRight2=   15195350
      GripperColor    =   15195350
      IsStrech        =   -1  'True
      RightColor1     =   0
      RightColor2     =   0
      ShowEndPanel    =   0   'False
      Theme           =   1
      ButtonCaption1  =   "Novo"
      ButtonEnabled1  =   0   'False
      ButtonIconSize1 =   32
      ButtonToolTipText1=   "Novo (Insert)"
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
      ButtonWidth1    =   36
      ButtonHeight1   =   21
      ButtonUseMaskColor1=   0   'False
      ButtonCaption2  =   "Filtrar"
      ButtonEnabled2  =   0   'False
      ButtonIconSize2 =   32
      ButtonToolTipText2=   "Filtrar (F2)"
      ButtonKey2      =   "2"
      ButtonAlignment2=   2
      BeginProperty ButtonFont2 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft2     =   40
      ButtonTop2      =   2
      ButtonWidth2    =   42
      ButtonHeight2   =   21
      ButtonUseMaskColor2=   0   'False
      ButtonCaption3  =   "Salvar"
      ButtonEnabled3  =   0   'False
      ButtonIconSize3 =   32
      ButtonToolTipText3=   "Salvar (F3)"
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
      ButtonLeft3     =   84
      ButtonTop3      =   2
      ButtonWidth3    =   44
      ButtonHeight3   =   21
      ButtonUseMaskColor3=   0   'False
      ButtonCaption4  =   "Excluir"
      ButtonEnabled4  =   0   'False
      ButtonIconSize4 =   32
      ButtonToolTipText4=   "Excluir (F4)"
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
      ButtonLeft4     =   130
      ButtonTop4      =   2
      ButtonWidth4    =   45
      ButtonHeight4   =   21
      ButtonUseMaskColor4=   0   'False
      ButtonCaption5  =   "Relatório"
      ButtonEnabled5  =   0   'False
      ButtonIconSize5 =   32
      ButtonToolTipText5=   "Relatório (F5)"
      ButtonKey5      =   "5"
      ButtonAlignment5=   2
      BeginProperty ButtonFont5 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft5     =   177
      ButtonTop5      =   2
      ButtonWidth5    =   60
      ButtonHeight5   =   21
      ButtonUseMaskColor5=   0   'False
      ButtonCaption6  =   "Anterior"
      ButtonEnabled6  =   0   'False
      ButtonIconSize6 =   32
      ButtonToolTipText6=   "Registro anterior."
      ButtonKey6      =   "6"
      ButtonAlignment6=   2
      BeginProperty ButtonFont6 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft6     =   239
      ButtonTop6      =   2
      ButtonWidth6    =   55
      ButtonHeight6   =   21
      ButtonUseMaskColor6=   0   'False
      ButtonCaption7  =   "Próximo"
      ButtonEnabled7  =   0   'False
      ButtonIconSize7 =   32
      ButtonToolTipText7=   "Próximo registro."
      ButtonKey7      =   "7"
      ButtonAlignment7=   2
      BeginProperty ButtonFont7 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft7     =   296
      ButtonTop7      =   2
      ButtonWidth7    =   55
      ButtonHeight7   =   21
      ButtonUseMaskColor7=   0   'False
      ButtonCaption8  =   "Enviar"
      ButtonEnabled8  =   0   'False
      ButtonIconSize8 =   32
      ButtonToolTipText8=   "Enviar CCe (F7)"
      ButtonKey8      =   "8"
      ButtonAlignment8=   2
      BeginProperty ButtonFont8 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft8     =   353
      ButtonTop8      =   2
      ButtonWidth8    =   43
      ButtonHeight8   =   21
      ButtonUseMaskColor8=   0   'False
      ButtonCaption9  =   "Consultar Status"
      ButtonEnabled9  =   0   'False
      ButtonIconSize9 =   32
      ButtonToolTipText9=   "Consultar Status CC-e"
      ButtonKey9      =   "9"
      ButtonAlignment9=   2
      BeginProperty ButtonFont9 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft9     =   398
      ButtonTop9      =   2
      ButtonWidth9    =   88
      ButtonHeight9   =   21
      ButtonUseMaskColor9=   0   'False
      ButtonEnabled10 =   0   'False
      ButtonIconSize10=   32
      ButtonAlignment10=   2
      ButtonType10    =   1
      ButtonStyle10   =   -1
      BeginProperty ButtonFont10 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState10   =   -1
      ButtonLeft10    =   488
      ButtonTop10     =   4
      ButtonWidth10   =   2
      ButtonHeight10  =   54
      ButtonCaption11 =   "Ajuda"
      ButtonEnabled11 =   0   'False
      ButtonIconSize11=   32
      ButtonToolTipText11=   "Ajuda (F1)"
      ButtonKey11     =   "11"
      ButtonAlignment11=   2
      BeginProperty ButtonFont11 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft11    =   492
      ButtonTop11     =   2
      ButtonWidth11   =   41
      ButtonHeight11  =   21
      ButtonUseMaskColor11=   0   'False
      ButtonCaption12 =   "Sair"
      ButtonEnabled12 =   0   'False
      ButtonIconSize12=   32
      ButtonToolTipText12=   "Sair (Esc)"
      ButtonKey12     =   "12"
      ButtonAlignment12=   2
      BeginProperty ButtonFont12 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft12    =   535
      ButtonTop12     =   2
      ButtonWidth12   =   30
      ButtonHeight12  =   21
      ButtonUseMaskColor12=   0   'False
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   5145
      Left            =   60
      TabIndex        =   12
      Top             =   3960
      Width           =   15285
      _ExtentX        =   26961
      _ExtentY        =   9075
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
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
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Text            =   "Código"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "T"
         Text            =   "Empresa"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Object.Tag             =   "D"
         Text            =   "Dt. emissão"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "Responsável"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Object.Tag             =   "N"
         Text            =   "Nota fiscal"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Sequencial"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Object.Tag             =   "T"
         Text            =   "Série"
         Object.Width           =   1323
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Object.Tag             =   "T"
         Text            =   "Destinatário"
         Object.Width           =   7629
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   8
         Object.Tag             =   "T"
         Text            =   "Status"
         Object.Width           =   5115
      EndProperty
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   55
      TabIndex        =   34
      Top             =   9750
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
      SearchText      =   "Atualizando..."
      Value           =   0
   End
End
Attribute VB_Name = "frmFaturamento_CartaCorrecao_NS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Novo_Carta As Boolean 'OK
Public StrSql_Localizar_Carta As String 'OK
Dim TBLISTA_Carta As ADODB.Recordset 'OK
Dim NomeArquivo As String
Dim CnpjNF As String

Sub ProcAjuda()
On Error GoTo tratar_erro

FunAbrirVideoWeb ("http://www.youtube.com/watch?v=A8dBCFhbghI&list=UUODGiDjQ-BCrxh0YXfJvoqg&index=2&feature=plcp")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Sub ProcCarregaDados()
On Error GoTo tratar_erro

If IsNull(TBLISTA!ID_empresa) = False And TBLISTA!ID_empresa <> "" Then ProcPuxaDadosComboEmpresa Cmb_empresa, TBLISTA!ID_empresa
txtCodigo = TBLISTA!ID
txtDataemissao.Value = IIf(IsNull(TBLISTA!Data_emissao), Date, TBLISTA!Data_emissao)
txtResponsavel = IIf(IsNull(TBLISTA!Responsavel), "", TBLISTA!Responsavel)
Txt_ID_nota = TBLISTA!ID_nota
txtNF = TBLISTA!int_NotaFiscal
Txt_serie = IIf(IsNull(TBLISTA!Serie), "", TBLISTA!Serie)
txtiddestinatario = IIf(IsNull(TBLISTA!Id_Int_Cliente), "", TBLISTA!Id_Int_Cliente)
Txt_destinatario = IIf(IsNull(TBLISTA!txt_Razao_Nome), "", TBLISTA!txt_Razao_Nome)
txtStatus = funStatus(IIf(IsNull(TBLISTA!status), "", TBLISTA!status))
txtSeq.Text = IIf(IsNull(TBLISTA!Seq), "1", TBLISTA!Seq)

If IsNull(TBLISTA!Chave_acesso) = True Then
    Set TBNivel1 = CreateObject("adodb.recordset")
    TBNivel1.Open "Select Chave_acesso,nsNRec from tbl_Dados_Nota_Fiscal_NFe where ID_nota = " & IIf(Txt_ID_nota = "", 0, Txt_ID_nota), Conexao, adOpenKeyset, adLockReadOnly
    If TBNivel1.EOF = False Then
    Txt_chave_acesso = IIf(IsNull(TBNivel1!Chave_acesso), "", TBNivel1!Chave_acesso)
    txtnsNrec.Text = IIf(IsNull(TBNivel1!nsNRec), "", TBNivel1!nsNRec)
    End If
    TBNivel1.Close
Else
    Txt_chave_acesso = IIf(IsNull(TBLISTA!Chave_acesso), "", TBLISTA!Chave_acesso)
End If

txtObs.Text = IIf(IsNull(TBLISTA!Obs), "", TBLISTA!Obs)
If TBLISTA!Desconsiderar_valor = True Then Chk_desconsiderar.Value = 1 Else Chk_desconsiderar.Value = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcAnterior()
On Error GoTo tratar_erro

If txtCodigo.Text = "" Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select CC.*, NF.int_NotaFiscal, NF.Serie, NF.id_int_cliente, NF.txt_Razao_Nome from NF_Carta_Correcao CC INNER JOIN tbl_Dados_Nota_Fiscal NF ON CC.ID_nota = NF.ID where order by CC.id", Conexao, adOpenKeyset, adLockReadOnly
If TBLISTA.BOF = False Then
    TBLISTA.Find ("id = " & txtCodigo.Text)
    TBLISTA.MovePrevious
    If TBLISTA.BOF = False Then
        ProcLimparCampos
        ProcCarregaDados
        procCarregaEmpresa
    Else
        USMsgBox ("Fim dos cadastros de carta de correção."), vbInformation, "CAPRIND v5.0"
    End If
End If
Novo_Carta = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub


Private Sub btnCertificado_Click()
On Error GoTo tratar_erro

Dim Stor As New Store
Dim Cert As Certificate
Dim Certs As New Certificates
Dim CForNext As Integer

'Abrir o store
Stor.Open

Certs.Clear
For CForNext = 1 To Stor.Certificates.Count
Certs.Add Stor.Certificates.Item(CForNext)
Next CForNext

Set Certs = Certs.Select("LaRoche", "Selecione o Certificado Digital.", False)

'Exibir mensagem com data de validade do certificado

For Each Cert In Certs
    USMsgBox "Nome razão: " & Cert.GetInfo(CAPICOM_CERT_INFO_SUBJECT_SIMPLE_NAME) & vbCrLf & "Certificado válido até: " & Cert.ValidToDate, vbInformation, "CAPRIND v5.0" '(CAPICOM_CHECK_TIME_VALIDITY)
    txtSerialCertificado.Text = Cert.SerialNumber
    txtval.Text = IIf(Cert.IsValid = True, "Sim", "Não")
Next

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub BTNnsnrec_Click()
On Error GoTo tratar_erro
Dim RespostaNSNrec As String
Dim p As Object

NomeArquivo = txtNF
nfDocumento = "CNR" & NomeArquivo
txtRetorno.Text = ""

If USMsgBox("Deseja realmente consultar recibo NS dessa nota?", vbYesNo, "CAPRIND v5.0") = vbYes Then
If Txt_chave_acesso <> "" And Len(Txt_chave_acesso.Text) = 44 Then
RespostaNSNrec = listarNSNRecs(Txt_chave_acesso)
txtRetorno.Text = RespostaNSNrec
status = LerDadosJSON(txtRetorno.Text, "status", "", "")
'Debug.print RespostaNSNrec
   If status = "200" Then
      Set p = JSON.parse(RespostaNSNrec)
      txtnsNrec.Text = p.Item("nsNRecs").Item(1).Item("nsNRec")
   Else
      USMsgBox txtRetorno.Text, vbCritical, "CAPRIND v5.0"
   End If

Else
USMsgBox "Para consultar o recibo NS é necessário a chave de acesso da nota com 44 digitos", vbInformation, "CAPRIND v5.0"
Exit Sub
End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_empresa_Click()
On Error GoTo tratar_erro

IDempresa = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)

ProcVerificaTPNFe
txtSerialCertificado = SerialCertificado
txtval = TPCertificado

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcProximo()
On Error GoTo tratar_erro

If txtCodigo.Text = "" Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select CC.*, NF.int_NotaFiscal, NF.Serie, NF.id_int_cliente, NF.txt_Razao_Nome from NF_Carta_Correcao CC INNER JOIN tbl_Dados_Nota_Fiscal NF ON CC.ID_nota = NF.ID where order by CC.id", Conexao, adOpenKeyset, adLockReadOnly
If TBLISTA.BOF = False Then
    TBLISTA.Find (" id =" & txtCodigo.Text)
    TBLISTA.MoveNext
    If TBLISTA.EOF = False Then
        ProcLimparCampos
        ProcCarregaDados
        procCarregaEmpresa
    Else
        USMsgBox ("Fim dos cadastros de carta de correção."), vbInformation, "CAPRIND v5.0"
    End If
End If
Novo = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub cmdConsultar_Click()
On Error GoTo tratar_erro

NomeArquivo = txtNF.Text
nfDocumento = "CSIT" & NomeArquivo

resposta = consultarSituacao(ReturnNumbersOnly(CnpjNF), Txt_chave_acesso, tpAmb, "4.00")
'resposta = consultarCadastroContribuinte(ReturnNumbersOnly(CnpjNF), "SP", "05272563000152", "CNPJ")
txtRetorno.Text = resposta

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub cmdPagAnt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Carta.AbsolutePage <> 2 Then
    If TBLISTA_Carta.AbsolutePage = -3 Then
        ProcExibePagina (TBLISTA_Carta.PageCount - 1)
    Else
        TBLISTA_Carta.AbsolutePage = TBLISTA_Carta.AbsolutePage - 2
        ProcExibePagina (TBLISTA_Carta.AbsolutePage)
    End If
Else
    ProcExibePagina (1)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub cmdPagIr_Click()
On Error GoTo tratar_erro

If txtPagIr = "" Then Exit Sub
Quant = ReturnNumbersOnly(Right(lblPaginas.Caption, 4))
If Quant <= 1 Or txtPagIr > Quant Then Exit Sub
If txtPagIr.Text >= 1 And txtPagIr.Text <= Quant Then
    TBLISTA_Carta.AbsolutePage = txtPagIr.Text
    ProcExibePagina (TBLISTA_Carta.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub cmdPagPrim_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Carta.AbsolutePage = 1
ProcExibePagina (TBLISTA_Carta.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub cmdPagProx_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Carta.AbsolutePage <> -3 Then
    If TBLISTA_Carta.AbsolutePage = 1 Then
        ProcExibePagina (2)
    Else
        ProcExibePagina (TBLISTA_Carta.AbsolutePage)
    End If
Else
    ProcExibePagina (TBLISTA_Carta.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub cmdPagUlt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Carta.AbsolutePage = TBLISTA_Carta.PageCount
ProcExibePagina (TBLISTA_Carta.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub cmdStatus_Click()
On Error GoTo tratar_erro
  
ProcStatusCCe

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro
  
Select Case KeyCode
    Case vbKeyInsert: ProcNovo
    Case vbKeyF2: ProcFiltrar
    Case vbKeyF3: ProcSalvar
    Case vbKeyF4: ProcExcluir
    Case vbKeyF5: ProcImprimir
    Case vbKeyF7: procEnviar
    Case vbKeyF1: ProcAjuda
    Case vbKeyEscape: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15225, 12, True
Formulario = "Faturamento/Carta de correção"
Direitos
ProcLimpaVariaveisPrincipais
ProcCarregaComboEmpresa Cmb_empresa, False
ProcRemoveObjetosResize Me
Contador2 = 0
txtdatacancelamento.Text = Format(Date, "yyyy-mm-dd") & "T" & Left(Time, 8) & FunVerifFusoHorario(True)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro

If IsInternetOnline = False Then
    USMsgBox ("Internet offline no momento, tente mais tarde."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

If txtCodigo = "" Then
    USMsgBox ("Informe a carta de correção antes de visualizar impressão."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

If Txt_chave_acesso = "" Then
    USMsgBox ("Nota fiscal sem chave de acesso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

If txtStatus = "Evento registrado e vinculado a NF-e" Then
Dim retorno As String
  retorno = NFeAPI.downloadEventoNFeAndSave(Txt_chave_acesso, tpAmb, "P", "CCE", txtSeq.Text, DiretorioXMLDanfe, True)
  xMotivo = LerDadosJSON(retorno, "retEvento", "xMotivo", "")
  USMsgBox xMotivo, vbInformation, "CAPRIND v5.0"
Else
  Set TBproducao = CreateObject("adodb.recordset")
  TBproducao.Open "Select CC.status, E.CNPJ from tbl_Dados_Nota_Fiscal NF INNER JOIN Empresa E ON NF.ID_empresa = E.Codigo INNER JOIN NF_Carta_Correcao CC ON NF.ID = CC.ID_nota WHERE NF.ID = " & Txt_ID_nota, Conexao, adOpenKeyset, adLockReadOnly
  If TBproducao.EOF = False Then
      If TBproducao!status = "135" Then
          procAbrirNotaPDF "NFe", TBproducao!CNPJ, txtNF, Txt_serie, DiretorioXMLDanfe, True
      Else
          NomeRel = "Faturamento_carta correcao.rpt"
          ProcImprimirRel "{NF_Carta_Correcao.id} = " & txtCodigo, ""
      End If
  End If
  TBproducao.Close
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcExcluir()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With Lista
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir esta(s) carta(s) de correção?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            Conexao.Execute "DELETE from NF_Carta_Correcao WHERE id = " & .ListItems(InitFor)
            'Conexao_NFe.Execute "DELETE from CBD001EVENTOS WHERE CbdEmpCodigo = " & .ListItems(InitFor).ListSubItems(7) & " and CbdNtfNumero = " & .ListItems(InitFor).ListSubItems(4) & " and CbdNtfSerie = " & .ListItems(InitFor).ListSubItems(5) & " and CbdEveDescCC = 'Carta de Correcao'"
            'Conexao_NFe.Execute "DELETE from NFE012 WHERE CbdEmpCodigo = " & .ListItems(InitFor).ListSubItems(7) & " and CbdNtfNumero = " & .ListItems(InitFor).ListSubItems(4) & " and CbdNtfSerie = " & .ListItems(InitFor).ListSubItems(5) & " and CbdAcao = 'V'"
            'Conexao_NFe.Execute "DELETE from NFE012EVENTOS WHERE CbdEmpCodigo = " & .ListItems(InitFor).ListSubItems(7) & " and CbdNtfNumero = " & .ListItems(InitFor).ListSubItems(4) & " and CbdNtfSerie = " & .ListItems(InitFor).ListSubItems(5) & " and CbdAcao = 'V'"
            'Conexao_NFe.Execute "DELETE from NFE001EVENTOS WHERE EmpCodigo = " & .ListItems(InitFor).ListSubItems(7) & " and NtfNumero = " & .ListItems(InitFor).ListSubItems(4) & " and NtfSerie = " & .ListItems(InitFor).ListSubItems(5) & " and EveDescCC = 'Carta de Correcao'"
            
            '==================================
            Modulo = "Faturamento/Carta de correção"
            Evento = "Excluir"
            ID_documento = .ListItems(InitFor)
            Documento = "Nota fiscal: " & .ListItems(InitFor).ListSubItems(4) & " - Série: " & .ListItems(InitFor).ListSubItems(5) & " - Destinatário: " & .ListItems(InitFor).ListSubItems(6)
            Documento1 = ""
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe a(s) carta(s) de correção antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Carta(s) de correção excluída(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimparCampos
    ProcCarregaLista (1)
    Frame1.Enabled = False
    'Frame3.Enabled = False
    Frame4.Enabled = False
    Novo_Carta = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcCancelar()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso.")
    Exit Sub
End If
Permitido = False
With Lista
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente cancelar a liberação desta(s) carta(s) de correção?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select CC.*, NF.int_NotaFiscal, NF.Serie, NF.id_int_cliente, NF.txt_Razao_Nome from NF_Carta_Correcao CC INNER JOIN tbl_Dados_Nota_Fiscal NF ON CC.ID_nota = NF.ID where CC.ID = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False Then
                '==================================
                Modulo = "Faturamento/Carta de correção"
                Evento = "Cancelar liberação"
                ID_documento = .ListItems(InitFor)
                Documento = "Nota fiscal: " & TBFI!int_NotaFiscal & " - Série: " & TBFI!Serie & " - Destinatário: " & TBFI!txt_Razao_Nome
                Documento1 = ""
                ProcGravaEvento
                '==================================
                
                OF = TBFI!int_NotaFiscal
                Set TBMaquinas = CreateObject("adodb.recordset")
                TBMaquinas.Open "Select * from Empresa where Codigo = " & TBFI!ID_empresa & " and GNFe = 'True'", Conexao, adOpenKeyset, adLockOptimistic
                If TBMaquinas.EOF = False Then
                    caminho = TBMaquinas!Caminho_Nfe & "\Empresa " & TBFI!ID_empresa & " - Serie " & TBFI!Serie & " - Nota " & OF & " - Status CCE.bat"
                    Set GerArqPastas = CreateObject("Scripting.FileSystemObject")
                    If GerArqPastas.FileExists(caminho) = True Then Kill caminho
                End If
                TBMaquinas.Close
                
                'Conexao_NFe.Execute "DELETE from CBD001EVENTOS WHERE CbdEmpCodigo = " & TBFI!ID_empresa & " and CbdNtfNumero = " & OF & " and CbdNtfSerie = " & TBFI!Serie & " and CbdEveDescCC = 'Carta de Correcao'"
                'Conexao_NFe.Execute "DELETE from NFE012 WHERE CbdEmpCodigo = " & TBFI!ID_empresa & " and CbdNtfNumero = " & OF & " and CbdNtfSerie = " & TBFI!Serie & " and CbdAcao = 'V'"
                'Conexao_NFe.Execute "DELETE from NFE012EVENTOS WHERE CbdEmpCodigo = " & TBFI!ID_empresa & " and CbdNtfNumero = " & OF & " and CbdNtfSerie = " & TBFI!Serie & " and CbdAcao = 'V'"
                'Conexao_NFe.Execute "DELETE from NFE001EVENTOS WHERE EmpCodigo = " & TBFI!ID_empresa & " and NtfNumero = " & OF & " and NtfSerie = " & TBFI!Serie & " and EveDescCC = 'Carta de Correcao'"
                
                TBFI!CODIGO = Null
                TBFI!status = Null
                TBFI!Chave_acesso = Null
                TBFI.Update
            End If
            TBFI.Close
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe a(s) carta(s) de correção antes de cancelar a liberação."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Liberação da(s) carta(s) de correção canceladas(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    txtStatus = "Não enviado"
    Txt_chave_acesso = ""
    ProcCarregaLista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

Novo = False
frmFaturamento_CartaCorrecao_Localizar_NS.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcNovo()
On Error GoTo tratar_erro

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
ProcLimparCampos
Frame1.Enabled = True
Frame4.Enabled = True
Novo_Carta = True
Minuta = False
Faturamento = False
frmMinuta_notas.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub



Private Sub ProcLimparCampos()
On Error GoTo tratar_erro

txtCodigo = ""
txtDataemissao.Value = Date
txtResponsavel = pubUsuario
txtNF = ""
Txt_serie = ""
txtiddestinatario = ""
Txt_destinatario = ""
txtStatus = "Não enviado"
Txt_chave_acesso = ""
txtObs.Text = ""
txtRetorno.Text = ""
Chk_desconsiderar.Value = 0
CodigoLista = 0
txtnsNrec.Text = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

If Novo_Carta = True Then
    If USMsgBox("A carta de correção ainda não foi salva, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvar
        If Novo_Carta = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
Novo_Carta = False
Timer_status_CCe.Enabled = False
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcSalvar()
On Error GoTo tratar_erro
Contador2 = 0

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Frame1.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
Acao = "salvar"
If txtNF = "" Then
    NomeCampo = "o número da nota fiscal"
    ProcVerificaAcao
    txtNF.SetFocus
    Exit Sub
End If
If Txt_serie = "" Then
    NomeCampo = "a série"
    ProcVerificaAcao
    Txt_serie.SetFocus
    Exit Sub
End If
If txtiddestinatario = "" Then
    NomeCampo = "o destinatário"
    Txt_destinatario.SetFocus
    ProcVerificaAcao
    Exit Sub
End If
If txtObs.Text = "" Then
    NomeCampo = "as correções"
    ProcVerificaAcao
    txtObs.SetFocus
    Exit Sub
End If
If txtSeq.Text = "" Then
    NomeCampo = "o numero sequencial"
    ProcVerificaAcao
    txtSeq.SetFocus
    Exit Sub
End If

If Len(Trim(txtObs.Text)) < 15 Then
    USMsgBox ("É necessário informar no mínimo 15 caracteres no campo de correção, favor revisar."), vbExclamation, "CAPRIND v5.0"
    txtObs.SetFocus
    Exit Sub
End If

Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from NF_Carta_Correcao where id = " & IIf(txtCodigo = "", 0, txtCodigo), Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then
TBGravar.AddNew
End If

TBGravar!ID_empresa = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
TBGravar!ID_nota = Txt_ID_nota
TBGravar!Data_emissao = txtDataemissao
TBGravar!Responsavel = txtResponsavel
If Chk_desconsiderar.Value = 1 Then TBGravar!Desconsiderar_valor = True Else TBGravar!Desconsiderar_valor = False
TBGravar!Obs = Replace(Trim(txtObs.Text), vbCrLf, " ")
TBGravar!Seq = IIf(txtSeq.Text <> "", txtSeq.Text, 1)
TBGravar.Update
txtCodigo = TBGravar!ID
TBGravar.Close

If Novo_Carta = True Then
    USMsgBox ("Nova carta de correção cadastrada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo"
    StrSql_Localizar_Carta = "Select E.Empresa, CC.*, NF.int_NotaFiscal, NF.Serie, NF.txt_Razao_Nome from (NF_Carta_Correcao CC INNER JOIN tbl_Dados_Nota_Fiscal NF ON CC.ID_nota = NF.ID) INNER JOIN Empresa E ON E.Codigo = CC.ID_empresa where CC.id = " & txtCodigo
    ProcCarregaLista (1)
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar"
    ProcCarregaLista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
    If Lista.ListItems.Count <> 0 And CodigoLista <> 0 Then
        Lista.SelectedItem = Lista.ListItems(CodigoLista)
        Lista.SetFocus
    End If
End If
'==================================
Modulo = "Faturamento/Carta de correção"
ID_documento = txtCodigo
Documento = "Nota fiscal: " & txtNF & " - Série: " & Txt_serie & " - Destinatário: " & Txt_destinatario
Documento1 = ""
ProcGravaEvento
'==================================
Novo_Carta = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcAtualizarStatus()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Data = Date - 120
If USMsgBox("Deseja realmente atualizar o status da(s) carta(s) de correção com data de emissão a patir de " & Format(Data, "dd/mm/yy") & "?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    Set TBGravar = CreateObject("adodb.recordset")
    TBGravar.Open "Select CC.*, NF.int_NotaFiscal, NF.Serie, NF.id_int_cliente, NF.txt_Razao_Nome from NF_Carta_Correcao CC INNER JOIN tbl_Dados_Nota_Fiscal NF ON CC.ID_nota = NF.ID where CC.Data_emissao >= '" & Format(Data, "Short Date") & "' order by CC.ID", Conexao, adOpenKeyset, adLockOptimistic
    If TBGravar.EOF = False Then
        TBGravar.MoveLast
        PBLista.Min = 0
        PBLista.Max = TBGravar.RecordCount
        PBLista.Value = 1
        Contador = 0
        TBGravar.MoveFirst
        Do While TBGravar.EOF = False
            OF = TBGravar!int_NotaFiscal
            
            'Verifica cartas de correção geradas para essa nota
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select * from NF_Carta_Correcao where ID_nota = " & TBGravar!ID_nota & " and ID < " & TBGravar!ID, Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                Contador2 = TBAbrir.RecordCount + 1
            Else
                Contador2 = 1
            End If
            TBAbrir.Close
            
'            Set TBGravar_NFe = CreateObject("adodb.recordset")
'            TBGravar_NFe.Open "Select * from NFE012EVENTOS where CbdNtfNumero = " & OF & " and CbdNtfSerie = '" & TBGravar!Serie & "' and CbdAcao = 'V' and CbdEvenSeq = " & Contador2 & " and CbdEveStsRetCod is not null order by CbdNtfNumero, CbdNtfSerie", Conexao_NFe, adOpenKeyset, adLockOptimistic
'            If TBGravar_NFe.EOF = False Then
'                TBGravar!CODIGO = TBGravar_NFe!CbdEveStsRetCod
'                TBGravar!Status = TBGravar_NFe!CbdEveStsRetNome
'                TBGravar!Chave_acesso = IIf(IsNull(TBGravar_NFe!CbdEveId), "", TBGravar_NFe!CbdEveId)
'                TBGravar.Update
'
'                Set TBMaquinas = CreateObject("adodb.recordset")
'                TBMaquinas.Open "Select * from Empresa where Codigo = " & TBGravar!ID_empresa & " and GNFe = 'True'", Conexao, adOpenKeyset, adLockOptimistic
'                If TBMaquinas.EOF = False Then
'                    Caminho = TBMaquinas!Caminho_Nfe & "\Empresa " & TBGravar!ID_empresa & " - Serie " & TBGravar!Serie & " - Nota " & OF & " - Status CCE.bat"
'                    Set GerArqPastas = CreateObject("Scripting.FileSystemObject")
'                    If GerArqPastas.FileExists(Caminho) = True Then Kill Caminho
'                End If
'                TBMaquinas.Close
'            End If
            TBGravar_NFe.Close
            TBGravar.MoveNext
            Contador = Contador + 1
            PBLista.Value = Contador
        Loop
    End If
    TBGravar.Close
    USMsgBox ("Status das(s) carta(s) de correção atualizado(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    '==================================
    Modulo = "Faturamento/Carta de correção"
    Evento = "Atualizar status"
    ID_documento = 0
    Documento = ""
    Documento1 = ""
    ProcGravaEvento
    '==================================
    ProcCarregaLista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
End If
Contador = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Sub ProcCarregaLista(Pagina As Integer)
On Error GoTo tratar_erro

Lista.ListItems.Clear
lblRegistros.Caption = "Nº de registros: 0"
lblPaginas.Caption = "Página: 0 de: 0"
If StrSql_Localizar_Carta = "" Then Exit Sub
Set TBLISTA_Carta = CreateObject("adodb.recordset")
TBLISTA_Carta.Open StrSql_Localizar_Carta, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA_Carta.EOF = False Then ProcExibePagina (Pagina)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcExibePagina(Pagina)
On Error GoTo tratar_erro

Lista.ListItems.Clear
TBLISTA_Carta.PageSize = IIf(txtNreg = "", 30, txtNreg)
TBLISTA_Carta.AbsolutePage = Pagina
TamanhoPagina = TBLISTA_Carta.PageSize
ContadorReg = 1

PBLista.Min = 0
PBLista.Max = FunVerifMaxPBListaPaginacao(TBLISTA_Carta.RecordCount - IIf(Pagina > 1, (TBLISTA_Carta.PageSize * (Pagina - 1)), 0), TBLISTA_Carta.PageSize)
PBLista.Value = 1
Contador = 0
Do While TBLISTA_Carta.EOF = False And (ContadorReg <= TamanhoPagina)
    With Lista.ListItems
        .Add , , TBLISTA_Carta!ID
        .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA_Carta!Empresa), "", TBLISTA_Carta!Empresa)
        .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA_Carta!Data_emissao), "", Format(TBLISTA_Carta!Data_emissao, "dd/mm/yy"))
        .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA_Carta!Responsavel), "", TBLISTA_Carta!Responsavel)
        .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA_Carta!int_NotaFiscal), "", TBLISTA_Carta!int_NotaFiscal)
        .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA_Carta!Seq), "1", TBLISTA_Carta!Seq)
        .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA_Carta!Serie), "", TBLISTA_Carta!Serie)
        .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA_Carta!txt_Razao_Nome), "", TBLISTA_Carta!txt_Razao_Nome)
        .Item(.Count).SubItems(8) = funStatus(IIf(IsNull(TBLISTA_Carta!status), "", TBLISTA_Carta!status))
    End With
    TBLISTA_Carta.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista.Value = Contador
Loop
lblRegistros.Caption = "Nº de registros: " & TBLISTA_Carta.RecordCount
If TBLISTA_Carta.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "Página: 1 de: " & TBLISTA_Carta.PageCount
ElseIf TBLISTA_Carta.AbsolutePage = adPosEOF Then
        lblPaginas.Caption = "Página: " & TBLISTA_Carta.PageCount & " de: " & TBLISTA_Carta.PageCount
    Else
        lblPaginas.Caption = "Página: " & TBLISTA_Carta.AbsolutePage - 1 & " de: " & TBLISTA_Carta.PageCount
End If
Contador = 0
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

Timer_status_CCe.Enabled = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "Código" Then
    With Lista
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open "Select * from NF_Carta_Correcao where ID = " & .ListItems(InitFor) & " AND CODIGO = 135", Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then GoTo Proximo
                TBAbrir.Close
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView Lista, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Lista_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With Lista
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select * from NF_Carta_Correcao where ID = " & .ListItems(InitFor) & " AND CODIGO = 135", Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                USMsgBox ("Não é permitido excluir carta de correção, pois a mesma está com status: " & TBAbrir!status & "."), vbExclamation, "CAPRIND v5.0"
                .ListItems.Item(InitFor).Checked = False
            End If
            TBAbrir.Close
        End If
    Next InitFor
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Lista_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select CC.*, NF.int_NotaFiscal, NF.Serie, NF.id_int_cliente, NF.txt_Razao_Nome, NF.ID_empresa from NF_Carta_Correcao CC INNER JOIN tbl_Dados_Nota_Fiscal NF ON CC.ID_nota = NF.ID where CC.id = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    ProcLimparCampos
    ProcCarregaDados
    procCarregaEmpresa
    CodigoLista = Lista.SelectedItem.index
End If
TBLISTA.Close
Frame1.Enabled = True
Frame4.Enabled = True
Novo_Carta = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Timer_status_CCe_Timer()
On Error GoTo tratar_erro

'If Timer_status_CCe.Enabled = True Then ProcAtualizaStatusCCe
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Timer_Timer()
On Error GoTo tratar_erro

'PBLista.Value = PBLista.Value + 1
'If Dir(DiretorioEnvio & "/" & NomeArquivo & ".xml") = "" Or PBLista.Value >= 200 Then
'    Timer.Enabled = False
'    procLerRetornoXML
'End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
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
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcNovo
    Case 2: ProcFiltrar
    Case 3: ProcSalvar
    Case 4: ProcExcluir
    Case 5: ProcImprimir
    Case 6: ProcAnterior
    Case 7: ProcProximo
    Case 8: procEnviar
    Case 9: ProcStatusCCe 'procLogErros
    'Case 10: procReenviarEmail
    'Case 11: ProcAjuda
    Case 12: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Public Sub ProcStatusCCe()
On Error GoTo tratar_erro

Dim ResultadoNFe As String
Dim StatusNFe As String
Dim protocolo As String
Dim CodigoStatus As String

If Len(Txt_chave_acesso) < 44 Then
USMsgBox "Chave de acesso inválida, favor corrigir.", vbCritical, "CAPRIND v5.0"
Exit Sub
End If

If USMsgBox("Deseja realmente consultar o status da CCe da nota N° " & txtNF.Text & " na SEFAZ?", vbYesNo, "CAPRIND 5.0") = vbNo Then Exit Sub

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from empresa where Empresa = '" & Cmb_empresa.Text & "'", Conexao, adOpenKeyset, adLockReadOnly
If TBAbrir.EOF = False Then
CnpjNF = ReturnNumbersOnly(TBAbrir!CNPJ)
End If
TBAbrir.Close
   
'ResultadoNFe = consultarStatusProcessamento(CnpjNF, txtnsNrec.Text, tpAmb)
StatusNFe = consultarSituacao(CnpjNF, Txt_chave_acesso.Text, tpAmb, "4.00")
txtRetorno.Text = StatusNFe
'Debug.print ResultadoNFe
'Debug.print StatusNFe
'===============================================================================
   Set p = JSON.parse(txtRetorno.Text)
   CodigoStatus = p.Item("retConsSitNFe").Item("cStat")
   If CodigoStatus = "100" Then
   CodigoStatus = p.Item("retConsSitNFe").Item("procEventoNFe").Item(1).Item("retEvento").Item("infEvento").Item("cStat") 'infEvento
'===============================================================================
' Grava o retorno da consulta
'===============================================================================
   If CodigoStatus = "135" Then
   Var = "200" 'RemoveAccents(StatusCorrecao)
   Conexao.Execute "Update NF_Carta_Correcao Set Status = " & Var & " where id_nota = " & Txt_ID_nota
   ProcCarregaLista (1)
   txtStatus = funStatus(Var)
   USMsgBox txtStatus, vbInformation, "CAPRIND v5.0"
   End If
   
'   txtRetorno.Text = p.Item("retConsSitNFe").Item("procEventoNFe").Item(1).Item("retEvento").Item("infEvento").Item("xMotivo") 'infEvento
'   txtRetorno.Text = txtRetorno.Text & vbCrLf & p.Item("retConsSitNFe").Item("procEventoNFe").Item(1).Item("retEvento").Item("infEvento").Item("chNFe") 'infEvento
'   txtRetorno.Text = txtRetorno.Text & vbCrLf & p.Item("retConsSitNFe").Item("procEventoNFe").Item(1).Item("retEvento").Item("infEvento").Item("tpEvento") 'infEvento
'   txtRetorno.Text = txtRetorno.Text & vbCrLf & p.Item("retConsSitNFe").Item("procEventoNFe").Item(1).Item("retEvento").Item("infEvento").Item("xEvento") 'infEvento
'   txtRetorno.Text = txtRetorno.Text & vbCrLf & p.Item("retConsSitNFe").Item("procEventoNFe").Item(1).Item("retEvento").Item("infEvento").Item("nSeqEvento") 'infEvento
'   txtRetorno.Text = txtRetorno.Text & vbCrLf & p.Item("retConsSitNFe").Item("procEventoNFe").Item(1).Item("retEvento").Item("infEvento").Item("nProt") 'infEvento
'===============================================================================
End If
 'Debug.print StatusNFe

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Public Sub procEnviar()
On Error GoTo tratar_erro
Dim CorrigeNFe As String
Dim StatusCorrecao As String

If tpAmb = "" Then

Exit Sub
End If

If IsInternetOnline = False Then
USMsgBox ("Internet encontra-se offline no momento, tente mais tarde!"), vbCritical, "CAPRIND v5.0"
Exit Sub
End If

'If txtStatus.Text = "Evento registrado e vinculado a NF-e" Then
'    USMsgBox ("Carta de correção já enviada."), vbExclamation, "CAPRIND v5.0"
'    Exit Sub
'End If

'If txtnsNrec.Text = "" Then
'    Usmsgbox ("Recibo não localizado, favor enviar o xml dessa nota para suporte@caprind.com.br."), vbExclamation, "CAPRIND v5.0"
'    Exit Sub
'End If


If Txt_chave_acesso.Text = "" Then
    USMsgBox ("Chave de acesso não identificada."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If txtCodigo = "" Then
    USMsgBox ("Informe a carta de correção antes de enviar."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

Set TBproducao = CreateObject("adodb.recordset")
TBproducao.Open "Select status from tbl_Dados_Nota_Fiscal_NFe WHERE ID_nota = " & Txt_ID_nota & " AND status <> 100", Conexao, adOpenKeyset, adLockReadOnly
If TBproducao.EOF = False Then
    USMsgBox ("Só é possível enviar carta de correção de notas aprovadas."), vbExclamation, "CAPRIND v5.0"
    TBproducao.Close
    Exit Sub
End If

Acao = "cancelar"
If funVerificaCampos = False Then Exit Sub

'=====================================================
' NS TECNOLOGIA
'=====================================================
'If TPns = True Then
nfDocumento = "CC" & txtNF.Text

If USMsgBox("Deseja enviar esta carta de correção com sequencia N°" & txtSeq.Text & " da nota fiscal N° " & txtNF.Text & " para aprovação do Sefaz?", vbYesNo, "CAPRIND 5.0") = vbYes Then
   txtdatacancelamento.Text = Format(Date, "yyyy-mm-dd") & "T" & Left(Time, 8) & FunVerifFusoHorario(True)
   Var = txtSeq.Text
   
   If Len(txtObs.Text) <> 0 Then
   Dim StrTextoCorrecao As String
            StrTextoCorrecao = RemoveAccents(txtObs.Text)
            StrTextoCorrecao = Replace(StrTextoCorrecao, vbCrLf, "")
            StrTextoCorrecao = RemoverCaracter(StrTextoCorrecao)
            'Debug.print StrTextoCorrecao
    End If
Var = txtSeq.Text

   CorrigeNFe = NFeAPI.corrigirNFe(Txt_chave_acesso.Text, tpAmb, txtdatacancelamento.Text, Var, StrTextoCorrecao, "XP", DiretorioRetorno, True)
   
   If CorrigeNFe = "" Then Exit Sub
   StatusCorrecao = LerDadosJSON(CorrigeNFe, "status", "", "")
   
   If StatusCorrecao = "-999" Then
   Dim xMotivo As String
       xMotivo = LerDadosJSON(CorrigeNFe, "erro", "xMotivo", "")
       txtRetorno.Text = xMotivo
   End If
   If StatusCorrecao = "-3" Then
       xMotivo = LerDadosJSON(CorrigeNFe, "erro", "xMotivo", "")
       txtRetorno.Text = xMotivo
       StatusCorrecao = "200"
        StatusCorrecao = RemoveAccents(StatusCorrecao)
        Conexao.Execute "Update NF_Carta_Correcao Set Status = " & StatusCorrecao & " where id_nota = " & Txt_ID_nota
        ProcCarregaLista (1)
        txtStatus = funStatus(StatusCorrecao)
        'USMsgBox txtStatus, vbInformation, "CAPRIND v5.0"
       
   End If
   If StatusCorrecao = "200" Then
       xMotivo = LerDadosJSON(CorrigeNFe, "motivo", "", "")
       txtRetorno.Text = xMotivo
   End If
   
   StatusCorrecao = RemoveAccents(StatusCorrecao)
   Conexao.Execute "Update NF_Carta_Correcao Set Status = " & StatusCorrecao & " where id_nota = " & Txt_ID_nota
   ProcCarregaLista (1)
   txtStatus = funStatus(StatusCorrecao)
   USMsgBox txtStatus, vbInformation, "CAPRIND v5.0"
End If
'End If


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub


Public Sub procCarregaEmpresa()
On Error GoTo tratar_erro

Set TBFI = CreateObject("adodb.recordset")
TBFI.Open "Select E.UF, E.Cidade, E.CNPJ, E.Caminho_Nfe, E.Caminho_XMLDanfe, E.Caminho_RetornoNfe from Empresa E INNER JOIN tbl_Dados_Nota_Fiscal N ON E.Codigo = N.ID_empresa where N.ID = " & Txt_ID_nota, Conexao, adOpenKeyset, adLockReadOnly
If TBFI.EOF = False Then
    CnpjNF = IIf(IsNull(TBFI!CNPJ), "", TBFI!CNPJ)
    DiretorioEnvio = IIf(IsNull(TBFI!Caminho_Nfe), "", TBFI!Caminho_Nfe)
    DiretorioXMLDanfe = IIf(IsNull(TBFI!Caminho_XMLDanfe), "", TBFI!Caminho_XMLDanfe)
    DiretorioRetorno = IIf(IsNull(TBFI!Caminho_RetornoNfe), "", TBFI!Caminho_RetornoNfe)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Function funVerificaCampos() As Boolean
On Error GoTo tratar_erro
funVerificaCampos = False

If DiretorioEnvio = "" Then
    NomeCampo = "o diretório de envio no cadastro da empresa"
    ProcVerificaAcao
    Exit Function
End If

If DiretorioRetorno = "" Then
    NomeCampo = "o diretório de retorno no cadastro da empresa"
    ProcVerificaAcao
    Exit Function
End If

If DiretorioXMLDanfe = "" Then
    NomeCampo = "o diretório de XML e Danfe no cadastro da empresa"
    ProcVerificaAcao
    Exit Function
End If

funVerificaCampos = True

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Function


Function funStatus(statusXML As String) As String
On Error GoTo tratar_erro

If statusXML = "135" Then
    funStatus = "Evento registrado e vinculado a NF-e"
    
ElseIf statusXML = "200" Then
    funStatus = "Evento registrado e vinculado a NF-e"
ElseIf statusXML = "-3" Then
    funStatus = "Não foi possível vincular"
ElseIf statusXML = "" Then
    funStatus = "Não enviado"
ElseIf statusXML = "Evento registrado e vinculado a NF-e" Then
    funStatus = "Evento registrado e vinculado a NF-e"
Else
    funStatus = "Não foi possível vincular"
End If

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Function

Public Sub procLogErros()
On Error GoTo tratar_erro

If Txt_ID_nota = 0 Then
    USMsgBox ("Informe a nota fiscal antes de consultar log de erros."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

Acao = "verificar o log"
If funVerificaCampos = False Then Exit Sub

Sit_REG = 3
frmFaturamento_Prod_Serv_NFSe_Log.Show 1
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Public Sub procAcionaTimer()
On Error GoTo tratar_erro

PBLista.Min = 0
PBLista.Max = 200
PBLista.Value = 0
Timer.Enabled = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub
