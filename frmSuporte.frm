VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{D2B08629-3629-406E-B7BD-0CBED5F2C38F}#63.0#0"; "kmail.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSuporte 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Suporte - Chamado"
   ClientHeight    =   10035
   ClientLeft      =   3555
   ClientTop       =   4590
   ClientWidth     =   15360
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSuporte.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10035
   ScaleWidth      =   15360
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   11130
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   99
      ScreenHeight    =   1080
      ScreenWidth     =   2560
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
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   3840
      Top             =   5700
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   3240
      Top             =   5700
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   60
      TabIndex        =   53
      Top             =   9120
      Width           =   15195
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
         TabIndex        =   25
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
         TabIndex        =   24
         Text            =   "30"
         ToolTipText     =   "Número de registros por página."
         Top             =   180
         Width           =   555
      End
      Begin DrawSuite2022.USButton cmdPagProx 
         Height          =   315
         Left            =   11760
         TabIndex        =   29
         ToolTipText     =   "Próxima página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmSuporte.frx":030A
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
         TabIndex        =   28
         ToolTipText     =   "Página anterior."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmSuporte.frx":3AB1
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
         TabIndex        =   26
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
         TabIndex        =   27
         ToolTipText     =   "Primeira página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmSuporte.frx":75BE
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
         TabIndex        =   30
         ToolTipText     =   "Última página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmSuporte.frx":B6AF
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
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "registros por página"
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
         Left            =   4410
         TabIndex        =   62
         Top             =   240
         Width           =   1440
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
         TabIndex        =   56
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
         TabIndex        =   55
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Carregar"
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
         TabIndex        =   54
         Top             =   240
         Width           =   645
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Retorno técnico"
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
      Height          =   2145
      Left            =   7345
      TabIndex        =   45
      Top             =   3060
      Width           =   7905
      Begin VB.TextBox Txt_horas_utilizada 
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
         Left            =   5460
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade de horas utilizada."
         Top             =   390
         Width           =   1035
      End
      Begin VB.TextBox Txt_horas_prevista 
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
         Left            =   4410
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade de horas prevista."
         Top             =   390
         Width           =   1035
      End
      Begin VB.TextBox Txt_data_conclusao 
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
         Left            =   6510
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   "Data de conclusão."
         Top             =   390
         Width           =   1185
      End
      Begin VB.TextBox Txt_prazo 
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
         Left            =   3420
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "Prazo."
         Top             =   390
         Width           =   975
      End
      Begin VB.TextBox Txt_proposta 
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
         Left            =   2550
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Proposta."
         Top             =   390
         Width           =   855
      End
      Begin VB.TextBox Txt_descricao_procam 
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
         Height          =   1005
         Left            =   180
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   "Descrição."
         Top             =   960
         Width           =   7515
      End
      Begin VB.TextBox Txt_programador 
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
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Programador."
         Top             =   390
         Width           =   2355
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hs. util."
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
         Left            =   5700
         TabIndex        =   59
         Top             =   180
         Width           =   555
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hs. prev."
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
         Left            =   4590
         TabIndex        =   58
         Top             =   180
         Width           =   675
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dt. conclusão"
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
         Left            =   6600
         TabIndex        =   50
         Top             =   180
         Width           =   975
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Prazo"
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
         Left            =   3705
         TabIndex        =   49
         Top             =   180
         Width           =   405
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Proposta"
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
         Left            =   2655
         TabIndex        =   48
         Top             =   180
         Width           =   645
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Programador"
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
         Left            =   892
         TabIndex        =   47
         Top             =   180
         Width           =   930
      End
      Begin VB.Label Label6 
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
         Height          =   195
         Left            =   3592
         TabIndex        =   46
         Top             =   750
         Width           =   690
      End
   End
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   11700
      Top             =   120
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frmSuporte.frx":EF3D
      Count           =   1
   End
   Begin VB.TextBox Txt_ID 
      Height          =   285
      Left            =   720
      TabIndex        =   39
      Text            =   "0"
      Top             =   6030
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Frame Frame2 
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
      Height          =   2055
      Left            =   55
      TabIndex        =   32
      Top             =   990
      Width           =   15195
      Begin VB.TextBox Txt_email 
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
         Left            =   8010
         MaxLength       =   150
         TabIndex        =   9
         ToolTipText     =   "E-mail."
         Top             =   990
         Width           =   7005
      End
      Begin VB.CommandButton Cmd_limpar_caminho 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   14370
         Picture         =   "frmSuporte.frx":14252
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Limpar caminho."
         Top             =   1570
         Width           =   315
      End
      Begin VB.CommandButton Cmd_visualizar_arquivo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   14700
         Picture         =   "frmSuporte.frx":14390
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Visualizar arquivo."
         Top             =   1570
         Width           =   315
      End
      Begin VB.TextBox Txt_prioridade 
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
         Left            =   4665
         TabIndex        =   1
         ToolTipText     =   "Prioridade."
         Top             =   390
         Width           =   935
      End
      Begin VB.CommandButton Cmd_anexo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   14040
         Picture         =   "frmSuporte.frx":14952
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Localizar arquivo."
         Top             =   1570
         Width           =   315
      End
      Begin VB.TextBox Txt_anexo 
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
         MaxLength       =   255
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Anexo."
         Top             =   1570
         Width           =   13845
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
         ItemData        =   "frmSuporte.frx":14A54
         Left            =   180
         List            =   "frmSuporte.frx":14A56
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Empresa."
         Top             =   390
         Width           =   4470
      End
      Begin VB.ComboBox Cmb_tipo 
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
         ItemData        =   "frmSuporte.frx":14A58
         Left            =   12960
         List            =   "frmSuporte.frx":14A74
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         ToolTipText     =   "Tipo."
         Top             =   390
         Width           =   2055
      End
      Begin VB.TextBox Txt_status 
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
         Left            =   11250
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Status."
         Top             =   390
         Width           =   1695
      End
      Begin VB.ComboBox Cmb_solicitante 
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
         ItemData        =   "frmSuporte.frx":14AC4
         Left            =   180
         List            =   "frmSuporte.frx":14AC6
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   7
         ToolTipText     =   "Solicitante."
         Top             =   990
         Width           =   4395
      End
      Begin VB.TextBox Txt_data 
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
         Left            =   6660
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Data do cadastro."
         Top             =   390
         Width           =   1035
      End
      Begin VB.TextBox Txt_setor 
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
         Left            =   4590
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Setor."
         Top             =   990
         Width           =   3405
      End
      Begin VB.TextBox Txt_codigo 
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
         Left            =   5610
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Código."
         Top             =   390
         Width           =   1035
      End
      Begin VB.TextBox Txt_responsavel 
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
         Left            =   7710
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Responsável pelo cadastro."
         Top             =   390
         Width           =   3525
      End
      Begin KmailProject.kmail kmail 
         Height          =   615
         Left            =   600
         TabIndex        =   61
         Top             =   240
         Visible         =   0   'False
         Width           =   3630
         _ExtentX        =   6403
         _ExtentY        =   1085
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Prioridade"
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
         Left            =   4777
         TabIndex        =   60
         Top             =   180
         Width           =   720
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Anexo"
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
         Left            =   6870
         TabIndex        =   57
         Top             =   1380
         Width           =   465
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Empresa"
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
         Left            =   2048
         TabIndex        =   52
         Top             =   180
         Width           =   735
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo*"
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
         Left            =   13792
         TabIndex        =   51
         Top             =   180
         Width           =   390
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "E-mail*"
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
         Index           =   7
         Left            =   11257
         TabIndex        =   44
         Top             =   780
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Solicitante*"
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
         Left            =   1965
         TabIndex        =   43
         Top             =   780
         Width           =   825
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
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
         Left            =   11865
         TabIndex        =   40
         Top             =   180
         Width           =   465
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Setor"
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
         Left            =   6097
         TabIndex        =   38
         Top             =   780
         Width           =   390
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
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
         Left            =   5880
         TabIndex        =   35
         Top             =   180
         Width           =   495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data"
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
         Index           =   0
         Left            =   7005
         TabIndex        =   34
         Top             =   180
         Width           =   345
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Responsável"
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
         Left            =   9015
         TabIndex        =   33
         Top             =   180
         Width           =   915
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
      Height          =   2145
      Left            =   55
      TabIndex        =   31
      Top             =   3060
      Width           =   7275
      Begin VB.ComboBox Cmb_modulo 
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
         ItemData        =   "frmSuporte.frx":14AC8
         Left            =   180
         List            =   "frmSuporte.frx":14ACA
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   14
         ToolTipText     =   "Módulo."
         Top             =   390
         Width           =   6885
      End
      Begin VB.TextBox Txt_descricao 
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
         Height          =   1005
         Left            =   180
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         ToolTipText     =   "Descrição."
         Top             =   960
         Width           =   6885
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição*"
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
         Left            =   3232
         TabIndex        =   37
         Top             =   750
         Width           =   780
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Módulo*"
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
         Left            =   3322
         TabIndex        =   36
         Top             =   180
         Width           =   600
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   41
      Top             =   0
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   1720
      ButtonCount     =   10
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft1     =   2
      ButtonTop1      =   2
      ButtonWidth1    =   33
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft2     =   37
      ButtonTop2      =   2
      ButtonWidth2    =   36
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft3     =   75
      ButtonTop3      =   2
      ButtonWidth3    =   38
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft4     =   115
      ButtonTop4      =   2
      ButtonWidth4    =   39
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft5     =   156
      ButtonTop5      =   2
      ButtonWidth5    =   51
      ButtonHeight5   =   21
      ButtonUseMaskColor5=   0   'False
      ButtonCaption6  =   "Renumerar"
      ButtonEnabled6  =   0   'False
      ButtonIconSize6 =   32
      ButtonToolTipText6=   "Renumerar prioridades (F8)"
      ButtonKey6      =   "6"
      ButtonAlignment6=   2
      BeginProperty ButtonFont6 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft6     =   209
      ButtonTop6      =   2
      ButtonWidth6    =   61
      ButtonHeight6   =   21
      ButtonUseMaskColor6=   0   'False
      ButtonEnabled7  =   0   'False
      ButtonIconSize7 =   32
      ButtonAlignment7=   2
      ButtonType7     =   1
      ButtonStyle7    =   -1
      BeginProperty ButtonFont7 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState7    =   -1
      ButtonLeft7     =   272
      ButtonTop7      =   4
      ButtonWidth7    =   2
      ButtonHeight7   =   54
      ButtonCaption8  =   "Ajuda"
      ButtonEnabled8  =   0   'False
      ButtonIconSize8 =   32
      ButtonToolTipText8=   "Ajuda (F1)"
      ButtonKey8      =   "8"
      ButtonAlignment8=   2
      BeginProperty ButtonFont8 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft8     =   276
      ButtonTop8      =   2
      ButtonWidth8    =   36
      ButtonHeight8   =   21
      ButtonUseMaskColor8=   0   'False
      ButtonCaption9  =   "Sair"
      ButtonEnabled9  =   0   'False
      ButtonIconSize9 =   32
      ButtonToolTipText9=   "Sair (Esc)"
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
      ButtonLeft9     =   314
      ButtonTop9      =   2
      ButtonWidth9    =   26
      ButtonHeight9   =   21
      ButtonUseMaskColor9=   0   'False
      ButtonEnabled10 =   0   'False
      ButtonIconSize10=   32
      ButtonKey10     =   "10"
      ButtonAlignment10=   2
      BeginProperty ButtonFont10 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState10   =   5
      ButtonLeft10    =   342
      ButtonTop10     =   2
      ButtonWidth10   =   24
      ButtonHeight10  =   24
      ButtonUseMaskColor10=   0   'False
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   55
      TabIndex        =   42
      Top             =   9750
      Width           =   15200
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4620
      Top             =   5790
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   3885
      Left            =   60
      TabIndex        =   23
      Top             =   5220
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   6853
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
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
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "N"
         Text            =   "Priorid."
         Object.Width           =   1147
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Código"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Object.Tag             =   "D"
         Text            =   "Data"
         Object.Width           =   1588
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Responsável"
         Object.Width           =   3879
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Tag             =   "T"
         Text            =   "Status"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Object.Tag             =   "T"
         Text            =   "Tipo"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Object.Tag             =   "T"
         Text            =   "Solicitante"
         Object.Width           =   3223
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Object.Tag             =   "T"
         Text            =   "Setor"
         Object.Width           =   3223
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Object.Tag             =   "T"
         Text            =   "Módulo"
         Object.Width           =   5645
      EndProperty
   End
End
Attribute VB_Name = "frmSuporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Novo_Suporte_Tecnico As Boolean 'OK
Public Sql_Atendimento_Localizar As String 'OK
Dim TBLISTA_Chamado As ADODB.Recordset 'OK
Public FormulaRel_Suporte_Tecnico As String 'OK

Sub ProcLimpaCampos()
On Error GoTo tratar_erro

Txt_ID = 0
txt_Codigo = ""
Txt_data = Format(Date, "dd/mm/yy")
Txt_responsavel = pubUsuario
Txt_status = "Aberto"
Cmb_tipo.ListIndex = -1
Cmb_solicitante.ListIndex = -1
Txt_setor = ""
Txt_email.Text = ""
Txt_prioridade = ""
Txt_anexo = ""
Nome_anexo = ""

Cmb_modulo.ListIndex = -1
Txt_descricao = ""

Txt_programador = ""
txt_proposta = ""
Txt_prazo = ""
Txt_horas_prevista = ""
Txt_horas_utilizada = ""
Txt_data_conclusao = ""
Txt_descricao_procam = ""

CodigoLista = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluir()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
FamiliaAntiga = ""
With Lista
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) chamado(s)?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            FunAbreBDSite
            Set TBFI = New ADODB.Recordset
            TBFI.Open "Select * from Chamado where ID = " & .ListItems(InitFor), ConexaoMySql, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False Then
                ConexaoMySql.Execute "Update Chamado Set Prioridade_cliente = Prioridade_cliente - 1 where Prioridade_cliente > " & TBFI!Prioridade_cliente & " and Prioridade_cliente <> 9999 and Prioridade_cliente IS NOT NULL and ID_cliente = " & FunVerifIDEmpresaSite(Cmb_empresa.ItemData(Cmb_empresa.ListIndex))
                
                Set TBMySQL = New ADODB.Recordset
                TBMySQL.Open "Select * From Clientes where ID = " & TBFI!ID_Cliente, ConexaoMySql, adOpenKeyset, adLockOptimistic
                If TBMySQL.EOF = False Then
                    Set TBFIltro = CreateObject("adodb.recordset")
                    TBFIltro.Open "Select Empresa from Empresa where CNPJ = '" & TBMySQL!CNPJ & "'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBFIltro.EOF = False Then
                        Empresa = TBFIltro!Empresa
                    End If
                    TBFIltro.Close
                End If
                
                
                Familiatext = "Empresa: " & Empresa & " - Exclusão do(s) chamado(s)"
                If FamiliaAntiga = "" Then FamiliaAntiga = "Chamado: " & TBFI!Numero & " - Modulo: " & TBFI!Modulo Else FamiliaAntiga = FamiliaAntiga & ";" & "Chamado: " & TBFI!Numero & " - Modulo: " & TBFI!Modulo
                
                ConexaoMySql.Execute "Delete from Chamado where ID = " & .ListItems(InitFor)
                
                '==================================
                Modulo = "Suporte/Chamado"
                Evento = "Excluir"
                ID_documento = .ListItems(InitFor)
                Documento = "Código: " & .ListItems(InitFor).ListSubItems(1)
                Documento1 = ""
                ProcGravaEvento
                '==================================
            End If
            TBFI.Close
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) chamados(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    FunFechaBDSite
    USMsgBox ("Chamado(s) excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCampos
    ProcCarregaLista (1)
    Frame1.Enabled = False
    Frame2.Enabled = False
    Frame3.Enabled = False
    Novo_Suporte_Tecnico = False
    
    With kmail
        .Charset = 0
        .Priority = NORMAL_PRIORITY
        If .sendEmail("smtp.caprind.com.br", Cmb_solicitante, "suporte@caprind.com.br", "", "", "suporte@caprind.com.br", "suporte@caprind.com.br", Familiatext, FamiliaAntiga, "", False, "non HTML text", True, 587, False, "suporte@caprind.com.br", "sup0902loc") Then
            USMsgBox ("E-mail enviado com sucesso."), vbInformation, "CAPRIND v5.0"
        Else
            USMsgBox ("Ocorreu um erro ao enviar o e-mail."), vbExclamation, "CAPRIND v5.0"
        End If
        .abort
    End With
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro
  
If Lista.ListItems.Count = 0 Then Exit Sub
If PermitidoRel = False Then LocalrelNovo = Localrel Else LocalrelNovo = LocalRelPersonalizado
Set Report = crAPP.OpenReport(LocalrelNovo & "\Suporte_lista de chamados.rpt")
frmimprimir.CrystalActiveXReportViewer1.ReportSource = Report
Report.FormulaSyntax = crCrystalSyntaxFormula
Report.RecordSelectionFormula = FormulaRel_Suporte_Tecnico
frmimprimir.CrystalActiveXReportViewer1.ViewReport
frmimprimir.Show 1
2:
    Set Report = Nothing
    Set crAPP = Nothing

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovo()
On Error GoTo tratar_erro
  
If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
ProcLimpaCampos
Procliberacampos
ProcCarregaComboTipo True
Novo_Suporte_Tecnico = True
Frame1.Enabled = True
Frame2.Enabled = True
Cmb_empresa.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCriarNovoNumero()
On Error GoTo tratar_erro

FunAbreBDSite
Set TBAbrir = New ADODB.Recordset
TBAbrir.Open "Select * from Chamado where ID_cliente = " & FunVerifIDEmpresaSite(Cmb_empresa.ItemData(Cmb_empresa.ListIndex)) & " and Year(data) = '" & Year(Date) & "' order by ID", ConexaoMySql, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    TBAbrir.MoveLast
    Numero = Left(TBAbrir!Numero, Len(TBAbrir!Numero) - 3)
    Numero = Right(Numero, 5) + 1
Else
    Numero = 1
End If
TBAbrir.Close

a = Numero
Ano = Right(Year(Date), 2)
Select Case Len(a)
    Case 1: a = "CH-0000" & Numero & "/" & Ano
    Case 2: a = "CH-000" & Numero & "/" & Ano
    Case 3: a = "CH-00" & Numero & "/" & Ano
    Case 4: a = "CH-0" & Numero & "/" & Ano
    Case 5: a = "CH-" & Numero & "/" & Ano
End Select
txt_Codigo = a

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLocalizar()
On Error GoTo tratar_erro

frmSuporte_Abrir.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

If Novo_Suporte_Tecnico = True Then
    If USMsgBox("O chamado ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvar
        If Novo_Suporte_Tecnico = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
Novo_Suporte_Tecnico = False
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_solicitante_Click()
On Error GoTo tratar_erro

Txt_setor = ""
Txt_email.Text = ""
If Cmb_solicitante = "" Then Exit Sub
Set TBUsuarios = CreateObject("adodb.recordset")
TBUsuarios.Open "Select * from Usuarios where usuario = '" & Cmb_solicitante & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBUsuarios.EOF = False Then
    Txt_setor = IIf(IsNull(TBUsuarios!Setor), "", TBUsuarios!Setor)
    Txt_email.Text = IIf(IsNull(TBUsuarios!Email), "", TBUsuarios!Email)
End If
TBUsuarios.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_tipo_Click()
On Error GoTo tratar_erro

ProcCarregaComboModulos Cmb_modulo, False, Cmb_tipo

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_anexo_Click()
On Error GoTo tratar_erro

ProcCarregaCaminhoNomeArquivo CommonDialog1, "*.*", "*.*"
Txt_anexo = caminho

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_limpar_caminho_Click()
On Error GoTo tratar_erro

Txt_anexo = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_visualizar_arquivo_Click()
On Error GoTo tratar_erro

If Txt_anexo <> "" Then ProcAbrirArquivo Txt_anexo

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt_Click()
On Error GoTo tratar_erro

FunAbreBDSite
If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Chamado.AbsolutePage <> 2 Then
    If TBLISTA_Chamado.AbsolutePage = -3 Then
        ProcExibePagina (TBLISTA_Chamado.PageCount - 1)
    Else
        TBLISTA_Chamado.AbsolutePage = TBLISTA_Chamado.AbsolutePage - 2
        ProcExibePagina (TBLISTA_Chamado.AbsolutePage)
    End If
Else
    ProcExibePagina (1)
End If
FunFechaBDSite

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagIr_Click()
On Error GoTo tratar_erro

FunAbreBDSite
If txtPagIr = "" Then Exit Sub
Quant = ReturnNumbersOnly(Right(lblPaginas.Caption, 4))
If Quant <= 1 Or txtPagIr > Quant Then Exit Sub
If txtPagIr.Text >= 1 And txtPagIr.Text <= Quant Then
    TBLISTA_Chamado.AbsolutePage = txtPagIr.Text
    ProcExibePagina (TBLISTA_Chamado.AbsolutePage)
End If
FunFechaBDSite

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim_Click()
On Error GoTo tratar_erro

FunAbreBDSite
If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Chamado.AbsolutePage = 1
ProcExibePagina (TBLISTA_Chamado.AbsolutePage)
FunFechaBDSite

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx_Click()
On Error GoTo tratar_erro

FunAbreBDSite
If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Chamado.AbsolutePage <> -3 Then
    If TBLISTA_Chamado.AbsolutePage = 1 Then
        ProcExibePagina (2)
    Else
        ProcExibePagina (TBLISTA_Chamado.AbsolutePage)
    End If
Else
    ProcExibePagina (TBLISTA_Chamado.PageCount)
End If
FunFechaBDSite

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt_Click()
On Error GoTo tratar_erro

FunAbreBDSite
If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Chamado.AbsolutePage = TBLISTA_Chamado.PageCount
ProcExibePagina (TBLISTA_Chamado.AbsolutePage)
FunFechaBDSite

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyEscape: ProcSair
    Case vbKeyInsert: ProcNovo
    Case vbKeyF2: ProcLocalizar
    Case vbKeyF3: ProcSalvar
    Case vbKeyF4: ProcExcluir
    Case vbKeyF5: ProcImprimir
    Case vbKeyF8: ProcRenumerar
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15195, 10, True
Formulario = "Suporte/Chamado"
Direitos
ProcCarregaComboEmpresa Cmb_empresa, False

With Cmb_solicitante
    .Clear
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from Empresa", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        FunAbreBDSite
        Set TBUsuarios = New ADODB.Recordset
        TBUsuarios.Open "Select U.usuario, U.ID from usuarios U INNER JOIN Clientes C ON U.CNPJ = C.CNPJ where C.ID = " & FunVerifIDEmpresaSite(TBAbrir!CODIGO) & " order by U.Usuario", ConexaoMySql, adOpenKeyset, adLockOptimistic
        If TBUsuarios.EOF = False Then
            Do While TBUsuarios.EOF = False
                .AddItem TBUsuarios!Usuario
                .ItemData(.NewIndex) = TBUsuarios!ID
                TBUsuarios.MoveNext
            Loop
        End If
        TBUsuarios.Close
    End If
    TBAbrir.Close
End With

ProcLimpaVariaveisPrincipais

ProcRemoveObjetosResize Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

Formulario = "Suporte/Chamado"
Direitos
ProcLimpaVariaveisPrincipais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvar()
On Error GoTo tratar_erro
  
If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Frame2.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
Acao = "salvar"
If Cmb_tipo = "" Then
    NomeCampo = "o tipo"
    ProcVerificaAcao
    Cmb_tipo.SetFocus
    Exit Sub
End If
If Cmb_solicitante = "" Then
    NomeCampo = "o solicitante"
    ProcVerificaAcao
    Cmb_solicitante.SetFocus
    Exit Sub
End If
If Txt_email.Text = "" Then
    NomeCampo = "o e-mail"
    ProcVerificaAcao
    Txt_email.SetFocus
    Exit Sub
End If
If Cmb_modulo = "" Then
    NomeCampo = "o módulo"
    ProcVerificaAcao
    Cmb_modulo.SetFocus
    Exit Sub
End If
If Txt_descricao = "" Then
    NomeCampo = "a descrição"
    ProcVerificaAcao
    Txt_descricao.SetFocus
    Exit Sub
End If

Permitido = False

FunAbreBDSite
Set TBGravar = New ADODB.Recordset
TBGravar.Open "Select * from Chamado where ID = " & Txt_ID, ConexaoMySql, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then
    ProcCriarNovoNumero
    
    TBGravar.AddNew
    TBGravar!Prioridade = 9999
    Permitido = True
Else
    TextoFiltroPos = ""
    If Txt_prioridade <> "" Then
        If Txt_prioridade < TBGravar!Prioridade_cliente Then
             TextoFiltroPos = "Prioridade_cliente = Prioridade_cliente + 1 where Prioridade_cliente >= " & Txt_prioridade & " and Prioridade_cliente < " & TBGravar!Prioridade_cliente
        ElseIf Txt_prioridade > TBGravar!Prioridade_cliente Then
                TextoFiltroPos = "Prioridade_cliente = Prioridade_cliente - 1 where Prioridade_cliente > " & TBGravar!Prioridade_cliente & " and Prioridade_cliente <= " & Txt_prioridade
        End If
    End If
    Prioridade = IIf(Txt_prioridade = "", "9999", Txt_prioridade)
    If Prioridade = TBGravar!Prioridade_cliente Then Permitido = True
End If
If TextoFiltroPos <> "" Then ConexaoMySql.Execute "Update Chamado Set " & TextoFiltroPos & " and Prioridade_cliente <> 9999 and Prioridade_cliente IS NOT NULL and ID_cliente = " & FunVerifIDEmpresaSite(Cmb_empresa.ItemData(Cmb_empresa.ListIndex))

TBGravar!Prioridade_cliente = IIf(Txt_prioridade = "", "9999", Txt_prioridade)
TBGravar!Numero = txt_Codigo
TBGravar!Data = Txt_data
TBGravar!Responsavel = Txt_responsavel
TBGravar!status = Txt_status
TBGravar!ID_Cliente = FunVerifIDEmpresaSite(Cmb_empresa.ItemData(Cmb_empresa.ListIndex))
TBGravar!ID_usuario = Cmb_solicitante.ItemData(Cmb_solicitante.ListIndex)
TBGravar!Email = LCase(Txt_email)
TBGravar!Anexo = Txt_anexo
TBGravar!Nome_anexo = Nome_anexo
Select Case Cmb_tipo
    Case "Bug": TBGravar!Tipo = "B"
    Case "Customização": TBGravar!Tipo = "C"
    Case "Configuração": TBGravar!Tipo = "CO"
    Case "Dúvida": TBGravar!Tipo = "D"
    Case "Instalação": TBGravar!Tipo = "I"
    Case "Outros": TBGravar!Tipo = "O"
    Case "Reindexação": TBGravar!Tipo = "R"
End Select
TBGravar!Modulo = Cmb_modulo
TBGravar!Descricao = Txt_descricao
TBGravar.Update
Txt_ID = TBGravar!ID
TBGravar.Close
FunFechaBDSite
If Novo_Suporte_Tecnico = True Then
    USMsgBox ("Novo chamado cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo"
    Familiatext = "Empresa: " & Cmb_empresa & " - Novo chamado: " & txt_Codigo & " - Modulo: " & Cmb_modulo
    
    Sql_Atendimento_Localizar = "Select * from Chamado where ID = " & Txt_ID
    ProcCarregaLista (1)
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar"
    Familiatext = "Empresa: " & Cmb_empresa & " - Alteração do chamado: " & txt_Codigo & " - Modulo: " & Cmb_modulo
    
    ProcCarregaLista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
    If Lista.ListItems.Count <> 0 And CodigoLista <> 0 Then
        Lista.SelectedItem = Lista.ListItems(CodigoLista)
        Lista.SetFocus
    End If
End If
'==================================
Modulo = "Suporte/Chamado"
ID_documento = Txt_ID
Documento = "Chamado: " & txt_Codigo
Documento1 = ""
ProcGravaEvento
'==================================
Novo_Suporte_Tecnico = False

If Permitido = True Then
    If Cmb_tipo = "Customização" Then FamiliaAntiga = "fabio.vendas@caprind.com.br" Else FamiliaAntiga = "suporte@caprind.com.br"
    Txt_descricao = Txt_descricao & vbCrLf & vbCrLf & "E-mail solicitante: " & Txt_email
    
    If FunVerifEnviarEmailOutlook(Cmb_empresa.ItemData(Cmb_empresa.ListIndex)) = True Then
        ProcEnviarEmailAutomatico MAPISession1, MAPIMessages1, FamiliaAntiga, Familiatext, Txt_descricao, Txt_anexo, Nome_anexo
    Else
        With kmail
            .Charset = 0
            .Priority = NORMAL_PRIORITY
            If .sendEmail("smtp.caprind.com.br", Cmb_solicitante, "suporte@caprind.com.br", "", "", FamiliaAntiga, FamiliaAntiga, Familiatext, Txt_descricao, Txt_anexo, False, "non HTML text", False, 587, True, "suporte@caprind.com.br", "sup0802loc") Then
                USMsgBox ("E-mail enviado com sucesso."), vbInformation, "CAPRIND v5.0"
            Else
                USMsgBox ("Ocorreu um erro ao enviar o e-mail."), vbExclamation, "CAPRIND v5.0"
            End If
            .abort
        End With
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcRenumerar()
On Error GoTo tratar_erro
  
If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If USMsgBox("Deseja realmente renumear as prioridades?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    Contador = 1
    FunAbreBDSite
    ConexaoMySql.Execute "Update Chamado Set Prioridade_cliente = 9999 where ID_cliente = " & FunVerifIDEmpresaSite(Cmb_empresa.ItemData(Cmb_empresa.ListIndex)) & " and (Status = 'Concluído' or Status = 'Cancelado')"
    Set TBGravar = New ADODB.Recordset
    TBGravar.Open "Select * from Chamado where ID_cliente = " & FunVerifIDEmpresaSite(Cmb_empresa.ItemData(Cmb_empresa.ListIndex)) & " And Prioridade_cliente <> 9999 and Status <> 'Concluído' and Status <> 'Cancelado' order by Prioridade_cliente", ConexaoMySql, adOpenKeyset, adLockOptimistic
    If TBGravar.EOF = False Then
        Do While TBGravar.EOF = False
            TBGravar!Prioridade_cliente = Contador
            TBGravar.Update
            
            '==================================
            Modulo = "Suporte/Chamado"
            Evento = "Renumerar prioridade"
            ID_documento = TBGravar!ID
            Documento = "Chamado: " & TBGravar!Numero
            Documento1 = ""
            ProcGravaEvento
            '==================================
            Contador = Contador + 1
            TBGravar.MoveNext
        Loop
        USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
        ProcCarregaLista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
    Else
        USMsgBox ("Não existe nenhum chamado com prioridade para ser renumerado."), vbInformation, "CAPRIND v5.0"
    End If
    TBGravar.Close
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub

End Sub

Sub ProcCarregaLista(Pagina As Integer)
On Error GoTo tratar_erro

lblRegistros.Caption = "Nº de registros: 0"
lblPaginas.Caption = "Página: 0 de: 0"
Lista.ListItems.Clear
If Sql_Atendimento_Localizar = "" Then Exit Sub
FunAbreBDSite
Set TBLISTA_Chamado = New ADODB.Recordset
TBLISTA_Chamado.Open Sql_Atendimento_Localizar, ConexaoMySql, adOpenKeyset, adLockOptimistic
If TBLISTA_Chamado.EOF = False Then ProcExibePagina (Pagina)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina(Pagina)
On Error GoTo tratar_erro

Lista.ListItems.Clear
TBLISTA_Chamado.PageSize = IIf(txtNreg = "", 30, txtNreg)
TBLISTA_Chamado.AbsolutePage = Pagina
TamanhoPagina = TBLISTA_Chamado.PageSize
ContadorReg = 1

PBLista.Min = 0
PBLista.Max = FunVerifMaxPBListaPaginacao(TBLISTA_Chamado.RecordCount - IIf(Pagina > 1, (TBLISTA_Chamado.PageSize * (Pagina - 1)), 0), TBLISTA_Chamado.PageSize)
PBLista.Value = 1
Contador = 0
Do While TBLISTA_Chamado.EOF = False And (ContadorReg <= TamanhoPagina)
    With Lista.ListItems
        .Add , , TBLISTA_Chamado!ID
        .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA_Chamado!Prioridade_cliente), "", IIf(TBLISTA_Chamado!Prioridade_cliente = 9999, "", TBLISTA_Chamado!Prioridade_cliente))
        .Item(.Count).SubItems(2) = TBLISTA_Chamado!Numero
        .Item(.Count).SubItems(3) = Format(TBLISTA_Chamado!Data, "dd/mm/yy")
        .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA_Chamado!Responsavel), "", TBLISTA_Chamado!Responsavel)
        .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA_Chamado!status), "", TBLISTA_Chamado!status)
        
        Select Case TBLISTA_Chamado!Tipo
            Case "B": Tipo = "Bug"
            Case "C": Tipo = "Customização"
            Case "CO": Tipo = "Configuração"
            Case "D": Tipo = "Dúvida"
            Case "I": Tipo = "Instalação"
            Case "O": Tipo = "Outros"
            Case "R": Tipo = "Reindexação"
        End Select
        .Item(.Count).SubItems(6) = Tipo
        
        Set TBMySQL = New ADODB.Recordset
        TBMySQL.Open "Select Usuario, Cargo From usuarios where ID = " & TBLISTA_Chamado!ID_usuario, ConexaoMySql, adOpenKeyset, adLockOptimistic
        If TBMySQL.EOF = False Then
            .Item(.Count).SubItems(7) = IIf(IsNull(TBMySQL!Usuario), "", TBMySQL!Usuario)
            .Item(.Count).SubItems(8) = IIf(IsNull(TBMySQL!Cargo), "", TBMySQL!Cargo)
        End If
        TBMySQL.Close
        
        .Item(.Count).SubItems(9) = IIf(IsNull(TBLISTA_Chamado!Modulo), "", TBLISTA_Chamado!Modulo)
    End With
    TBLISTA_Chamado.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista.Value = Contador
Loop
lblRegistros.Caption = "Nº de registros: " & TBLISTA_Chamado.RecordCount
If TBLISTA_Chamado.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "Página: 1 de: " & TBLISTA_Chamado.PageCount
ElseIf TBLISTA_Chamado.AbsolutePage = adPosEOF Then
        lblPaginas.Caption = "Página: " & TBLISTA_Chamado.PageCount & " de: " & TBLISTA_Chamado.PageCount
    Else
        lblPaginas.Caption = "Página: " & TBLISTA_Chamado.AbsolutePage - 1 & " de: " & TBLISTA_Chamado.PageCount
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Lista
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If .ListItems.Item(InitFor).ListSubItems(5) <> "Aberto" Then GoTo Proximo
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
    Exit Sub
End Sub

Private Sub Lista_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With Lista
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If .ListItems.Item(InitFor).ListSubItems(5) <> "Aberto" Then
                USMsgBox ("Não é permitido excluir este chamado, pois o mesmo se encontra " & .ListItems.Item(InitFor).ListSubItems(5) & "."), vbExclamation, "CAPRIND v5.0"
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
        End If
    Next InitFor
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro
  
If Lista.ListItems.Count = 0 Then Exit Sub
FunAbreBDSite
Set TBLISTA = New ADODB.Recordset
TBLISTA.Open "Select * from Chamado where ID = " & Lista.SelectedItem, ConexaoMySql, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    ProcLimpaCampos
    ProcPuxaDados
    CodigoLista = Lista.SelectedItem.index
End If
TBLISTA.Close
FunFechaBDSite

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcPuxaDados()
On Error GoTo tratar_erro

Set TBMySQL = New ADODB.Recordset
TBMySQL.Open "Select * From Clientes where ID = " & TBLISTA!ID_Cliente, ConexaoMySql, adOpenKeyset, adLockOptimistic
If TBMySQL.EOF = False Then
    Set TBFIltro = CreateObject("adodb.recordset")
    TBFIltro.Open "Select Empresa from Empresa where CNPJ = '" & TBMySQL!CNPJ & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBFIltro.EOF = False Then
        Cmb_empresa = TBFIltro!Empresa
    End If
    TBFIltro.Close
End If

Txt_ID = TBLISTA!ID
txt_Codigo = TBLISTA!Numero
Txt_data = Format(TBLISTA!Data, "dd/mm/yy")
Txt_responsavel = IIf(IsNull(TBLISTA!Responsavel), "", TBLISTA!Responsavel)
Txt_status = IIf(IsNull(TBLISTA!status), "", TBLISTA!status)

ProcCarregaComboTipo False
Select Case TBLISTA!Tipo
    Case "B": Cmb_tipo = "Bug"
    Case "C": Cmb_tipo = "Customização"
    Case "CO": Cmb_tipo = "Configuração"
    Case "D": Cmb_tipo = "Dúvida"
    Case "I": Cmb_tipo = "Instalação"
    Case "O": Cmb_tipo = "Outros"
    Case "R": Cmb_tipo = "Reindexação"
End Select

Set TBMySQL = New ADODB.Recordset
TBMySQL.Open "Select * From usuarios where ID = " & TBLISTA!ID_usuario, ConexaoMySql, adOpenKeyset, adLockOptimistic
If TBMySQL.EOF = False Then
    Cmb_solicitante = IIf(IsNull(TBMySQL!Usuario), "", TBMySQL!Usuario)
    Txt_setor = IIf(IsNull(TBMySQL!Cargo), "", TBMySQL!Cargo)
End If
TBMySQL.Close

Txt_email.Text = IIf(IsNull(TBLISTA!Email), "", TBLISTA!Email)
Txt_prioridade = IIf(IsNull(TBLISTA!Prioridade_cliente), "", IIf(TBLISTA!Prioridade_cliente = 9999, "", TBLISTA!Prioridade_cliente))
If TBLISTA!Anexo <> "" Then
    Txt_anexo = IIf(IsNull(TBLISTA!Anexo), "", TBLISTA!Anexo)
    Nome_anexo = IIf(IsNull(TBLISTA!Nome_anexo), "", TBLISTA!Nome_anexo)
End If

If IsNull(TBLISTA!Modulo) = False And TBLISTA!Modulo <> "" Then Cmb_modulo = TBLISTA!Modulo
Txt_descricao = IIf(IsNull(TBLISTA!Descricao), "", TBLISTA!Descricao)

Txt_programador = IIf(IsNull(TBLISTA!Programador), "", TBLISTA!Programador)
txt_proposta = IIf(IsNull(TBLISTA!Proposta), "", TBLISTA!Proposta)
Txt_prazo = IIf(IsNull(TBLISTA!Prazo), "", Format(TBLISTA!Prazo, "dd/mm/yy"))
If IsNull(TBLISTA!Horas_prevista) = False And TBLISTA!Horas_prevista <> "" Then Txt_horas_prevista = FormataTempo(TBLISTA!Horas_prevista * 3600)
If IsNull(TBLISTA!Horas_utilizada) = False And TBLISTA!Horas_utilizada <> "" Then Txt_horas_utilizada = FormataTempo(TBLISTA!Horas_utilizada * 3600)
Txt_data_conclusao = IIf(IsNull(TBLISTA!data_conclusao), "", Format(TBLISTA!data_conclusao, "dd/mm/yy"))
Txt_descricao_procam = IIf(IsNull(TBLISTA!Descricao_procam), "", TBLISTA!Descricao_procam)

Frame1.Enabled = True
Frame2.Enabled = True
Frame3.Enabled = True

Novo_Suporte_Tecnico = False

If Txt_status <> "Aberto" Then ProcBloqueiaCampos Else Procliberacampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcBloqueiaCampos()
On Error GoTo tratar_erro

With Cmb_tipo
    .Locked = True
    .TabStop = False
End With
With Cmb_solicitante
    .Locked = True
    .TabStop = False
End With
With Txt_email
    .Locked = True
    .TabStop = False
End With
Cmd_anexo.Enabled = False
Cmd_limpar_caminho.Enabled = False
With Cmb_modulo
    .Locked = True
    .TabStop = False
End With
With Txt_descricao
    .Locked = True
    .TabStop = False
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Procliberacampos()
On Error GoTo tratar_erro

With Cmb_tipo
    .Locked = False
    .TabStop = True
End With
With Cmb_solicitante
    .Locked = False
    .TabStop = True
End With
With Txt_email
    .Locked = False
    .TabStop = True
End With
Cmd_anexo.Enabled = True
Cmd_limpar_caminho.Enabled = True
With Cmb_modulo
    .Locked = False
    .TabStop = True
End With
With Txt_descricao
    .Locked = False
    .TabStop = True
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_codigo_Change()
On Error GoTo tratar_erro

If Novo_Suporte_Tecnico = True Then
VerifCodigo:
    FunAbreBDSite
    Set TBFI = CreateObject("adodb.recordset")
    TBFI.Open "Select * from Chamado where ID_cliente = " & FunVerifIDEmpresaSite(Cmb_empresa.ItemData(Cmb_empresa.ListIndex)) & " and Numero = '" & txt_Codigo & "' and ID <> " & Txt_ID, ConexaoMySql, adOpenKeyset, adLockOptimistic
    If TBFI.EOF = False Then
        Numero = Left(txt_Codigo, Len(txt_Codigo) - 3)
        Numero = Right(Numero, 5) + 1
        Ano = Right(Year(Date), 2)
        a = Numero
        Select Case Len(a)
            Case 1: a = "CH-0000" & Numero & "/" & Ano
            Case 2: a = "CH-000" & Numero & "/" & Ano
            Case 3: a = "CH-00" & Numero & "/" & Ano
            Case 4: a = "CH-0" & Numero & "/" & Ano
            Case 5: a = "CH-" & Numero & "/" & Ano
        End Select
        txt_Codigo = a
        GoTo VerifCodigo
    End If
    TBFI.Close
End If
        
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_email_LostFocus()
On Error GoTo tratar_erro

Txt_email = LCase(Txt_email)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_prioridade_Change()
On Error GoTo tratar_erro

If Txt_prioridade <> "" Then
    VerifNumero = Txt_prioridade
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_prioridade = ""
        Txt_prioridade.SetFocus
        Exit Sub
    End If
End If
    
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

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcNovo
    Case 2: ProcLocalizar
    Case 3: ProcSalvar
    Case 4: ProcExcluir
    Case 5: ProcImprimir
    Case 6: ProcRenumerar
    'Case 8: ProcAjuda
    Case 9: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaComboTipo(Novo As Boolean)
On Error GoTo tratar_erro

With Cmb_tipo
    .Clear
    .AddItem "Bug"
    .AddItem "Configuração"
    .AddItem "Customização"
    If Novo = False Then .AddItem "Dúvida"
    .AddItem "Instalação"
    .AddItem "Outros"
    .AddItem "Reindexação"
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
