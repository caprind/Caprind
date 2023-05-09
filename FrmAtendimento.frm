VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{8C1279ED-044C-4258-A3E3-0D5514B899FC}#1.44#0"; "ControlesUteis.ocx"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmAtendimento 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Suporte - Solicitação de atendimento"
   ClientHeight    =   10035
   ClientLeft      =   120
   ClientTop       =   4590
   ClientWidth     =   15360
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmAtendimento.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   10035
   ScaleWidth      =   15360
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Cmb_filtrar_por 
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
      ItemData        =   "FrmAtendimento.frx":0CCA
      Left            =   13740
      List            =   "FrmAtendimento.frx":0CD7
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   23
      ToolTipText     =   "Filtrar mensagens por."
      Top             =   3270
      Width           =   1395
   End
   Begin VB.Timer Tmr_mensagem 
      Interval        =   5000
      Left            =   5070
      Top             =   3270
   End
   Begin VB.Frame Frame4 
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
      ForeColor       =   &H00000000&
      Height          =   1185
      Left            =   0
      TabIndex        =   38
      Top             =   8550
      Width           =   3555
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
         Left            =   2040
         TabIndex        =   11
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
         Left            =   180
         TabIndex        =   12
         ToolTipText     =   "Número da página."
         Top             =   510
         Width           =   555
      End
      Begin DrawSuite2022.USButton cmdPagProx 
         Height          =   315
         Left            =   2370
         TabIndex        =   16
         ToolTipText     =   "Próxima página."
         Top             =   510
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "FrmAtendimento.frx":0CF0
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
         Left            =   1830
         TabIndex        =   15
         ToolTipText     =   "Página anterior."
         Top             =   510
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "FrmAtendimento.frx":4494
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
         Left            =   750
         TabIndex        =   13
         Top             =   510
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
         Left            =   1290
         TabIndex        =   14
         ToolTipText     =   "Primeira página."
         Top             =   510
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "FrmAtendimento.frx":7F9D
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
         Left            =   2910
         TabIndex        =   17
         ToolTipText     =   "Última página."
         Top             =   510
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "FrmAtendimento.frx":C08C
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "reg. por pág."
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
         Left            =   2670
         TabIndex        =   44
         Top             =   240
         Width           =   795
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Carregar"
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
         Left            =   1440
         TabIndex        =   41
         Top             =   240
         Width           =   540
      End
      Begin VB.Label lblRegistros 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nº de reg.: 0"
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
         Left            =   180
         TabIndex        =   40
         Top             =   240
         Width           =   810
      End
      Begin VB.Label lblPaginas 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Página: 0 de: 0"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1560
         TabIndex        =   39
         Top             =   900
         Width           =   1875
      End
   End
   Begin VB.TextBox Txt_ID 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   720
      TabIndex        =   29
      Text            =   "0"
      Top             =   6030
      Visible         =   0   'False
      Width           =   675
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
      Height          =   2055
      Left            =   60
      TabIndex        =   24
      Top             =   990
      Width           =   15195
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
         Height          =   915
         Left            =   9900
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         ToolTipText     =   "Descrição."
         Top             =   990
         Width           =   5115
      End
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
         Left            =   180
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   8
         ToolTipText     =   "Módulo."
         Top             =   1570
         Width           =   9615
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
         Left            =   180
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
         ItemData        =   "FrmAtendimento.frx":F918
         Left            =   12960
         List            =   "FrmAtendimento.frx":F92B
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
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
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Status."
         Top             =   390
         Width           =   1695
      End
      Begin VB.TextBox Txt_data 
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
         Left            =   5970
         Locked          =   -1  'True
         TabIndex        =   2
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
         Left            =   180
         TabIndex        =   6
         ToolTipText     =   "Setor."
         Top             =   990
         Width           =   4395
      End
      Begin VB.TextBox Txt_codigo 
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
         Left            =   4650
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Código."
         Top             =   390
         Width           =   1305
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
         Left            =   7020
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Responsável pelo cadastro."
         Top             =   390
         Width           =   4215
      End
      Begin ControlesUteis.txt Txt_email 
         Height          =   360
         Left            =   4590
         TabIndex        =   7
         ToolTipText     =   "E-mail."
         Top             =   990
         Width           =   5205
         _ExtentX        =   9181
         _ExtentY        =   635
         Tamanho         =   5205
         Tipo            =   2
         Text            =   ""
         FocusColor      =   16777215
         CaptionColor    =   0
         ShowCaption     =   0   'False
         Caption         =   ""
         MaxLength       =   60
         BackColor       =   14737632
         Negative        =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição"
         Height          =   195
         Left            =   12112
         TabIndex        =   37
         Top             =   780
         Width           =   690
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Módulo"
         Height          =   195
         Left            =   4732
         TabIndex        =   35
         Top             =   1380
         Width           =   510
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
         TabIndex        =   34
         Top             =   180
         Width           =   735
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo"
         Height          =   195
         Left            =   13837
         TabIndex        =   33
         Top             =   180
         Width           =   300
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "E-mail"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   7
         Left            =   6982
         TabIndex        =   32
         Top             =   780
         Width           =   420
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         Height          =   195
         Left            =   11865
         TabIndex        =   30
         Top             =   180
         Width           =   465
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Setor"
         Height          =   195
         Left            =   2250
         TabIndex        =   28
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5017
         TabIndex        =   27
         Top             =   180
         Width           =   570
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data"
         Height          =   195
         Index           =   0
         Left            =   6315
         TabIndex        =   26
         Top             =   180
         Width           =   345
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Responsável"
         Height          =   195
         Left            =   8670
         TabIndex        =   25
         Top             =   180
         Width           =   915
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   31
      Top             =   0
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   1720
      ButtonCount     =   8
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
      ButtonCaption4  =   "Status"
      ButtonEnabled4  =   0   'False
      ButtonIconSize4 =   32
      ButtonToolTipText4=   "Alterar status (F7)"
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
      ButtonEnabled5  =   0   'False
      ButtonIconSize5 =   32
      ButtonAlignment5=   2
      ButtonType5     =   1
      ButtonStyle5    =   -1
      BeginProperty ButtonFont5 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState5    =   -1
      ButtonLeft5     =   156
      ButtonTop5      =   4
      ButtonWidth5    =   2
      ButtonHeight5   =   54
      ButtonCaption6  =   "Ajuda"
      ButtonEnabled6  =   0   'False
      ButtonIconSize6 =   32
      ButtonToolTipText6=   "Ajuda (F1)"
      ButtonKey6      =   "7"
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
      ButtonLeft6     =   160
      ButtonTop6      =   2
      ButtonWidth6    =   36
      ButtonHeight6   =   21
      ButtonUseMaskColor6=   0   'False
      ButtonCaption7  =   "Sair"
      ButtonEnabled7  =   0   'False
      ButtonIconSize7 =   32
      ButtonToolTipText7=   "Sair (Esc)"
      ButtonKey7      =   "8"
      ButtonAlignment7=   2
      BeginProperty ButtonFont7 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft7     =   198
      ButtonTop7      =   2
      ButtonWidth7    =   26
      ButtonHeight7   =   21
      ButtonUseMaskColor7=   0   'False
      ButtonEnabled8  =   0   'False
      ButtonIconSize8 =   32
      ButtonKey8      =   "9"
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
      ButtonState8    =   5
      ButtonLeft8     =   226
      ButtonTop8      =   2
      ButtonWidth8    =   24
      ButtonHeight8   =   24
      ButtonUseMaskColor8=   0   'False
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   10290
         Top             =   180
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "FrmAtendimento.frx":F95D
         Count           =   1
      End
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   8040
      Top             =   3270
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7530
      Top             =   3330
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   5385
      Left            =   0
      TabIndex        =   10
      Top             =   3150
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   9499
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "T"
         Text            =   "Código"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Object.Tag             =   "D"
         Text            =   "Data"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "Responsável"
         Object.Width           =   1790
      EndProperty
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   0
      TabIndex        =   42
      Top             =   9750
      Width           =   3555
      _ExtentX        =   6271
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
      Height          =   6945
      Left            =   3600
      TabIndex        =   36
      Top             =   3060
      Width           =   11655
      Begin VB.Timer Tmr_alarme_nova_mensagem 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   1920
         Top             =   210
      End
      Begin VB.TextBox Txt_mensagem 
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
         Height          =   960
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   18
         ToolTipText     =   "Mensagem a ser enviada."
         Top             =   5850
         Width           =   10335
      End
      Begin InetCtlsObjects.Inet InetFTP 
         Left            =   810
         Top             =   150
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
      End
      Begin MSComctlLib.ListView Lista_mensagem 
         Height          =   4695
         Left            =   210
         TabIndex        =   20
         Top             =   690
         Width           =   11205
         _ExtentX        =   19764
         _ExtentY        =   8281
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         AllowReorder    =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ForeColor       =   0
         BackColor       =   16777215
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   3
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Data/Hora"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Mensagem"
            Object.Width           =   16960
         EndProperty
      End
      Begin DrawSuite2022.USLabel lblUtilizado 
         Height          =   195
         Left            =   9300
         TabIndex        =   45
         Top             =   5580
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   344
         Caption         =   "0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   255
         NoHTMLCaption   =   "0"
      End
      Begin DrawSuite2022.USLabel LblRestam 
         Height          =   195
         Left            =   10970
         TabIndex        =   46
         Top             =   5580
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   344
         Caption         =   "500"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   12582912
         NoHTMLCaption   =   "500"
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   180
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   702
         ImageHeight     =   25
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmAtendimento.frx":139E4
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmAtendimento.frx":14C19
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmAtendimento.frx":1611B
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin DrawSuite2022.USLabel USLabel12 
         Height          =   195
         Left            =   8490
         TabIndex        =   47
         Top             =   240
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   344
         Caption         =   "Filtrar mensagens por:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         NoHTMLCaption   =   "Filtrar mensagens por:"
      End
      Begin DrawSuite2022.USLabel USLabel1 
         Height          =   195
         Left            =   6060
         TabIndex        =   48
         Top             =   5580
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   344
         Caption         =   "Disponível : 500 caracteres        Utilizado (s):                    Resta (m):"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         NoHTMLCaption   =   "Disponível : 500 caracteres        Utilizado (s):                    Resta (m):"
      End
      Begin DrawSuite2022.USButton Cmd_enviar_mensagem 
         Height          =   960
         Left            =   10470
         TabIndex        =   19
         Top             =   5850
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   1693
         Caption         =   "Enviar"
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
         PicAlign        =   6
         PicSize         =   4
         PicSizeH        =   48
         PicSizeW        =   48
      End
      Begin DrawSuite2022.USButton Cmd_enviar 
         Height          =   255
         Left            =   420
         TabIndex        =   21
         Top             =   5550
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   450
         Caption         =   "Enviar arquivo"
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
         PicAlign        =   6
         PicSize         =   4
         PicSizeH        =   48
         PicSizeW        =   48
      End
      Begin DrawSuite2022.USButton Cmd_receber 
         Height          =   255
         Left            =   2160
         TabIndex        =   22
         Top             =   5550
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   450
         Caption         =   "Receber arquivo"
         Enabled         =   0   'False
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
         PicAlign        =   6
         PicSize         =   4
         PicSizeH        =   48
         PicSizeW        =   48
         State           =   3
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   4905
         Left            =   90
         Top             =   600
         Width           =   11445
      End
   End
   Begin WMPLibCtl.WindowsMediaPlayer WMP 
      Height          =   1155
      Left            =   1920
      TabIndex        =   43
      Top             =   3630
      Visible         =   0   'False
      Width           =   1335
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   51
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   2355
      _cy             =   2037
   End
End
Attribute VB_Name = "FrmAtendimento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Novo_Atendimento As Boolean 'OK
Public Sql_Atendimento_Localizar As String 'OK
Dim TBLISTA_Atendimento As ADODB.Recordset 'OK
Dim TBLISTA_AtendimentoMsg As ADODB.Recordset 'OK
Dim IDMsgemNova As Long 'OK
Dim NTotalMensagem As Long 'OK

Private Sub ProcLimpaCampos()
On Error GoTo tratar_erro

Txt_ID = 0
txt_Codigo = ""
Txt_data = Format(Date, "dd/mm/yy")
Txt_responsavel = pubUsuario
Txt_status = "Aberto"
Cmb_tipo.ListIndex = -1

Set TBUsuarios = CreateObject("adodb.recordset")
TBUsuarios.Open "Select Setor, Email from Usuarios where usuario = '" & Txt_responsavel & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBUsuarios.EOF = False Then
    Txt_setor = IIf(IsNull(TBUsuarios!Setor), "", TBUsuarios!Setor)
    Txt_email.Text = IIf(IsNull(TBUsuarios!Email), "", TBUsuarios!Email)
End If
TBUsuarios.Close

Cmb_modulo.ListIndex = -1
Txt_descricao = ""
Lista_mensagem.ListItems.Clear
Txt_mensagem = ""
lblUtilizado = 0
LblRestam = 500
CodigoLista = 0

IDMsgemNova = 0
NTotalMensagem = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcStatus()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With Lista
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente alterar o status desta(s) solicitação(ões) de atendimento?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            
            FunAbreBDSite
            Set TBFI = New ADODB.Recordset
            TBFI.Open "Select Status from Atendimentos where ID = " & .ListItems(InitFor), ConexaoMySql, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False Then
                If TBFI!status = "Aberto" Then TBFI!status = "Encerrado" Else TBFI!status = "Aberto"
                TBFI.Update
            End If
            TBFI.Close
            FunFechaBDSite
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe a(s) solicitação(ões) de atendimento antes de alterar o status."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Status alterado com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCampos
    ProcCarregaLista (1)
End If

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
Novo_Atendimento = True
Frame1.Enabled = True
Frame2.Enabled = False
Cmb_empresa.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCriarNovoNumero()
On Error GoTo tratar_erro

FunAbreBDSite
Set TBAbrir = New ADODB.Recordset
TBAbrir.Open "Select Numero from Atendimentos where ID_cliente = " & FunVerifIDEmpresaSite(Cmb_empresa.ItemData(Cmb_empresa.ListIndex)) & " and Year(data) = '" & Year(Date) & "' order by ID desc", ConexaoMySql, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    Numero = Left(TBAbrir!Numero, Len(TBAbrir!Numero) - 3)
    Numero = Right(Numero, 5) + 1
Else
    Numero = 1
End If
TBAbrir.Close

a = Numero
Ano = Right(Year(Date), 2)
Select Case Len(a)
    Case 1: a = "AT-0000" & Numero & "/" & Ano
    Case 2: a = "AT-000" & Numero & "/" & Ano
    Case 3: a = "AT-00" & Numero & "/" & Ano
    Case 4: a = "AT-0" & Numero & "/" & Ano
    Case 5: a = "AT-" & Numero & "/" & Ano
End Select
txt_Codigo = a

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLocalizar()
On Error GoTo tratar_erro

FrmAtendimento_Abrir.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

If Novo_Atendimento = True Then
    If USMsgBox("A solicitação de atendimento ainda não foi salva, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvar
        If Novo_Atendimento = True Then Exit Sub Else Unload Me
    End If
End If
Novo_Atendimento = False
Unload Me

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

Private Sub Cmb_filtrar_por_Click()
On Error GoTo tratar_erro

If Txt_ID = 0 Then Exit Sub
ProcVerifFiltroMsgem

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcVerifFiltroMsgem()
On Error GoTo tratar_erro

Select Case Cmb_filtrar_por
    Case "Hoje": ProcCarregaListaMensagens " and data >= '" & Format(Date, "yyyy,mm,dd") & "'"
    Case "Semana": ProcCarregaListaMensagens " and data >= '" & Format(Date - 7, "yyyy,mm,dd") & "'"
    Case "Todos":  ProcCarregaListaMensagens ""
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaListaMensagens(TextoFiltro As String)
On Error GoTo tratar_erro

CompLetra = 90
DataFim = 0
Dataini = 0
Contador = 1
Anterior = ""
Frase = ""
Texto = ""

FunAbreBDSite
Set TBLISTA_AtendimentoMsg = New ADODB.Recordset
TBLISTA_AtendimentoMsg.Open "Select ID, Data, Hora, De, Tipo_de, Mensagem from Atendimentos_chat where ID_atendimento = " & Txt_ID & TextoFiltro & " group by Data, ID, Hora, De, Tipo_de, Mensagem order by Data, ID", ConexaoMySql, adOpenKeyset, adLockReadOnly
If TBLISTA_AtendimentoMsg.EOF = False Then
    If NTotalMensagem <> TBLISTA_AtendimentoMsg.RecordCount Then
        Lista_mensagem.ListItems.Clear
        Do While TBLISTA_AtendimentoMsg.EOF = False
            With Lista_mensagem.ListItems
                'Criar cabeçalho
                If Dataini <> TBLISTA_AtendimentoMsg!Data Then
                    .Add , , TBLISTA_AtendimentoMsg!ID, , 1
                    .Item(.Count).SubItems(2) = Format(TBLISTA_AtendimentoMsg!Data, "dd/mm/yy")
                    .Item(Lista_mensagem.ListItems.Count).ListSubItems(2).Bold = True
                End If
            
                If TBLISTA_AtendimentoMsg!De <> Anterior Then
                    .Add , , TBLISTA_AtendimentoMsg!ID, , IIf(TBLISTA_AtendimentoMsg!Tipo_de = "U", 2, 3)
                    .Item(.Count).SubItems(2) = TBLISTA_AtendimentoMsg!De
                    .Item(.Count).ListSubItems(2).Bold = True
                    If TBLISTA_AtendimentoMsg!Tipo_de = "U" Then .Item(.Count).ListSubItems(2).ForeColor = vbBlack
                End If
                
                'Adicionar mensagem
                Texto = TBLISTA_AtendimentoMsg!Mensagem
                If Len(Texto) <= CompLetra Then
                    Lista_mensagem.Font.Bold = False
                    .Add , , TBLISTA_AtendimentoMsg!ID
                    .Item(.Count).SubItems(1) = Format(TBLISTA_AtendimentoMsg!Hora, "hh:mm:ss")
                    .Item(.Count).SubItems(2) = TBLISTA_AtendimentoMsg!Mensagem 'Adiciona a frase na lista
                Else
                    'Construir a frase a ser carregada na lista
                    Do While Len(Texto) > Contador 'Enquanto não termina de ler todo o texto
                        Contador = Contador + 1 'Soma os caracteres do texto
                        a = Mid(Texto, Contador, 1) 'Pega o caracter do texto para montar a frase
                        Frase = Frase & a 'Monta a frase
                        
                        If Len(Frase) >= CompLetra Then 'And A = " " Or A = "." Then 'Verifica se a frase termina com espaço ou ponto
                            Lista_mensagem.Font.Bold = False
                            .Add , , TBLISTA_AtendimentoMsg!ID
                            .Item(.Count).SubItems(1) = IIf(DataFim <> TBLISTA_AtendimentoMsg!Hora, Format(TBLISTA_AtendimentoMsg!Hora, "hh:mm:ss"), "")
                            .Item(.Count).SubItems(2) = Trim(Frase) & IIf(a = " " Or a = ".", "", "_") 'Adiciona a frase na lista
                            Frase = "" 'Limpa a frase
                            DataFim = TBLISTA_AtendimentoMsg!Hora
                        End If
                     Loop
                     Frase = Trim(Frase)
                     If Len(Texto) >= Contador Then
                        Lista_mensagem.Font.Bold = False
                        .Add , , TBLISTA_AtendimentoMsg!ID
                        .Item(.Count).SubItems(1) = IIf(DataFim <> TBLISTA_AtendimentoMsg!Hora, Format(TBLISTA_AtendimentoMsg!Hora, "hh:mm:ss"), "")
                        .Item(.Count).SubItems(2) = Trim(Frase) 'Adiciona a frase na lista
                        Frase = "" 'Limpa a frase
                    End If
                End If
            End With
            Anterior = TBLISTA_AtendimentoMsg!De
            Dataini = TBLISTA_AtendimentoMsg!Data
            TBLISTA_AtendimentoMsg.MoveNext
        Loop
        Lista_mensagem.Refresh
    End If
    NTotalMensagem = TBLISTA_AtendimentoMsg.RecordCount
End If
FunFechaBDSite

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_enviar_mensagem_Click()
On Error GoTo tratar_erro

If Txt_mensagem = "" Then Exit Sub
ProcEnviaDadosMsg "", False
With Txt_mensagem
    .SelStart = 0
    .SelLength = Len(.Text)
    .Text = ""
    .SetFocus
    lblUtilizado = 0
    LblRestam = 500
End With
Tmr_mensagem.Enabled = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviaDadosMsg(Nome_anexo As String, Enviar As Boolean)
On Error GoTo tratar_erro

FunAbreBDSite
Set TBAbrir = New ADODB.Recordset
TBAbrir.Open "Select * from Atendimentos_chat", ConexaoMySql, adOpenKeyset, adLockOptimistic
TBAbrir.AddNew
TBAbrir!ID_atendimento = Txt_ID
TBAbrir!Data = Date
TBAbrir!Hora = Time
TBAbrir!De = pubUsuario
TBAbrir!Tipo_de = "U"
TBAbrir!Para = IIf(FunVerificaPara = "", Null, FunVerificaPara)
If IsNull(TBAbrir!Para) = False Then TBAbrir!Tipo_para = "A"
If Nome_anexo = "" Then
    TBAbrir!Mensagem = Txt_mensagem
Else
    TBAbrir!Mensagem = "O arquivo " & Nome_anexo & " foi " & IIf(Enviar = True, "enviado", "recebido") & " com sucesso."
    TBAbrir!Arquivo = Nome_anexo
End If
TBAbrir.Update
ConexaoMySql.Execute "Update Atendimentos_chat Set Respondida = 1 where ID_atendimento = " & Txt_ID & " and ID < " & TBAbrir!ID
TBAbrir.Close
FunFechaBDSite

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Function FunVerificaPara() As String
On Error GoTo tratar_erro

FunVerificaPara = ""
'FunAbreBDSite
Set TBFI = New ADODB.Recordset
TBFI.Open "Select Para from Atendimentos_chat where ID_atendimento = " & Txt_ID & " and Tipo_para = 'A' order by ID desc", ConexaoMySql, adOpenKeyset, adLockOptimistic
If TBFI.EOF = False Then FunVerificaPara = TBFI!Para
TBFI.Close

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Private Sub Cmd_Receber_Click()
On Error GoTo tratar_erro

Anexo = ""
With CommonDialog1
    .Filter = "(*.*) | *.*"
    .InitDir = App.Path
    .DefaultExt = "*.*"
    .filename = Nome_anexo
    .ShowSave
     Anexo = .filename
End With

If Nome_anexo = "" Then Exit Sub
Arquivo_local = IIf(Anexo = "", "", Anexo)
Arquivo_Site = "/public_html/Arquivos/Chat/Atendimentos/" & Nome_anexo
url_ftp = "ftp.caprind.com.br"
usuario_ftp = "caprind1"
Senha_ftp = "cap0902loc"
hopen = InternetOpen("ftp", 1, vbNullString, vbNullString, &H10000000)
hconnect = InternetConnect(hopen, url_ftp, 21, usuario_ftp, Senha_ftp, 1, &H8000000, 0)
ftp = FtpGetFile(hconnect, Arquivo_Site, Arquivo_local, False, INTERNET_FLAG_RELOAD, dwType, 0)
ProcEnviaDadosMsg Nome_anexo, False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_enviar_Click()
On Error GoTo tratar_erro

ProcCarregaCaminhoNomeArquivo CommonDialog1, "*.*", "*.*"
If Nome_anexo = "" Then Exit Sub
Arquivo_local = caminho
Arquivo_Site = "/public_html/Arquivos/Chat/Atendimentos/" & Nome_anexo
url_ftp = "ftp.caprind.com.br"
usuario_ftp = "caprind1"
Senha_ftp = "cap0902loc"
hopen = InternetOpen("ftp", 1, vbNullString, vbNullString, &H10000000)
hconnect = InternetConnect(hopen, url_ftp, 21, usuario_ftp, Senha_ftp, 1, &H8000000, 0)
ftp = FtpPutFile(hconnect, Arquivo_local, Arquivo_Site, 0, 0)
ProcEnviaDadosMsg Nome_anexo, True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt_Click()
On Error GoTo tratar_erro

FunAbreBDSite
If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Atendimento.AbsolutePage <> 2 Then
    If TBLISTA_Atendimento.AbsolutePage = -3 Then
        ProcExibePagina (TBLISTA_Atendimento.PageCount - 1)
    Else
        TBLISTA_Atendimento.AbsolutePage = TBLISTA_Atendimento.AbsolutePage - 2
        ProcExibePagina (TBLISTA_Atendimento.AbsolutePage)
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
    TBLISTA_Atendimento.AbsolutePage = txtPagIr.Text
    ProcExibePagina (TBLISTA_Atendimento.AbsolutePage)
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
TBLISTA_Atendimento.AbsolutePage = 1
ProcExibePagina (TBLISTA_Atendimento.AbsolutePage)
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
If TBLISTA_Atendimento.AbsolutePage <> -3 Then
    If TBLISTA_Atendimento.AbsolutePage = 1 Then
        ProcExibePagina (2)
    Else
        ProcExibePagina (TBLISTA_Atendimento.AbsolutePage)
    End If
Else
    ProcExibePagina (TBLISTA_Atendimento.PageCount)
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
TBLISTA_Atendimento.AbsolutePage = TBLISTA_Atendimento.PageCount
ProcExibePagina (TBLISTA_Atendimento.AbsolutePage)
FunFechaBDSite

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyInsert: ProcNovo
    Case vbKeyF2: ProcLocalizar
    Case vbKeyF3: ProcSalvar
    Case vbKeyF7: ProcStatus
    'Case vbKeyF1: ProcAjuda
    Case vbKeyEscape: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15195, 8, True
Formulario = "Suporte/Solicitação de atendimento"
Direitos
ProcLimpaVariaveisPrincipais

ProcCarregaComboEmpresa Cmb_empresa, False
Cmb_filtrar_por = "Hoje"
If PBLista.Value = 0 Then PBLista.Value = 100

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

Formulario = "Suporte/Solicitação de atendimento"
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
If Frame1.Enabled = False Then
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
If Txt_setor = "" Then
    NomeCampo = "o setor"
    ProcVerificaAcao
    TxtSolicitante.SetFocus
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

FunAbreBDSite
Set TBGravar = New ADODB.Recordset
TBGravar.Open "Select * from Atendimentos where ID = " & Txt_ID & "", ConexaoMySql, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then
    ProcCriarNovoNumero
    TBGravar.AddNew
End If
TBGravar!ID_Cliente = FunVerifIDEmpresaSite(Cmb_empresa.ItemData(Cmb_empresa.ListIndex))
TBGravar!Numero = txt_Codigo
TBGravar!Data = Txt_data
TBGravar!Responsavel = Txt_responsavel
TBGravar!status = Txt_status
Select Case Cmb_tipo
    Case "Bug": TBGravar!Tipo = "B"
    Case "Dúvida": TBGravar!Tipo = "D"
    Case "Instalação": TBGravar!Tipo = "I"
    Case "Reindexação": TBGravar!Tipo = "R"
    Case "Outros": TBGravar!Tipo = "O"
End Select
TBGravar!Setor = Txt_setor
TBGravar!Email = Txt_email.Text
TBGravar!Modulo = Cmb_modulo
TBGravar!Descricao = Txt_descricao
TBGravar.Update
Txt_ID = TBGravar!ID
TBGravar.Close
FunFechaBDSite
If Novo_Atendimento = True Then
    USMsgBox ("Nova solicitação de atendimento cadastrada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo"
    Sql_Atendimento_Localizar = "Select * from Atendimentos where ID = " & Txt_ID
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
Modulo = "Suporte/Solicitação de atendimento"
ID_documento = Txt_ID
Documento = "Atendimento: " & txt_Codigo
Documento1 = ""
ProcGravaEvento
'==================================
Novo_Atendimento = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaLista(Pagina As Integer)
On Error GoTo tratar_erro

Lista.ListItems.Clear
If Sql_Atendimento_Localizar = "" Then Exit Sub
FunAbreBDSite
Set TBLISTA_Atendimento = New ADODB.Recordset
TBLISTA_Atendimento.Open Sql_Atendimento_Localizar, ConexaoMySql, adOpenKeyset, adLockOptimistic
If TBLISTA_Atendimento.EOF = False Then ProcExibePagina (Pagina)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina(Pagina)
On Error GoTo tratar_erro

Lista.ListItems.Clear
TBLISTA_Atendimento.PageSize = IIf(txtNreg = "", 30, txtNreg)
TBLISTA_Atendimento.AbsolutePage = Pagina
TamanhoPagina = TBLISTA_Atendimento.PageSize
ContadorReg = 1

PBLista.Min = 0
PBLista.Max = IIf(IIf(TBLISTA_Atendimento.RecordCount - TBLISTA_Atendimento.AbsolutePosition <= 0, 1, TBLISTA_Atendimento.RecordCount - TBLISTA_Atendimento.AbsolutePosition) < TBLISTA_Atendimento.PageSize, IIf(TBLISTA_Atendimento.RecordCount - TBLISTA_Atendimento.AbsolutePosition <= 0, 1, TBLISTA_Atendimento.RecordCount - TBLISTA_Atendimento.AbsolutePosition), TBLISTA_Atendimento.PageSize)
PBLista.Value = 1
Contador = 0
Do While TBLISTA_Atendimento.EOF = False And (ContadorReg <= TamanhoPagina)
    With Lista.ListItems
        .Add , , TBLISTA_Atendimento!ID
        .Item(.Count).SubItems(1) = TBLISTA_Atendimento!Numero
        .Item(.Count).SubItems(2) = Format(TBLISTA_Atendimento!Data, "dd/mm/yy")
        .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA_Atendimento!Responsavel), "", TBLISTA_Atendimento!Responsavel)
    End With
    TBLISTA_Atendimento.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista.Value = Contador
Loop
lblRegistros.Caption = "Nº de reg.: " & TBLISTA_Atendimento.RecordCount
If TBLISTA_Atendimento.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "Página: 1 de: " & TBLISTA_Atendimento.PageCount
ElseIf TBLISTA_Atendimento.AbsolutePage = adPosEOF Then
        lblPaginas.Caption = "Página: " & TBLISTA_Atendimento.PageCount & " de: " & TBLISTA_Atendimento.PageCount
    Else
        lblPaginas.Caption = "Página: " & TBLISTA_Atendimento.AbsolutePage - 1 & " de: " & TBLISTA_Atendimento.PageCount
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro
  
If Lista.ListItems.Count = 0 Then Exit Sub
FunAbreBDSite
Set TBFIltro = New ADODB.Recordset
TBFIltro.Open "Select * from Atendimentos where ID = " & Lista.SelectedItem, ConexaoMySql, adOpenKeyset, adLockOptimistic
If TBFIltro.EOF = False Then
    ProcLimpaCampos
    ProcPuxaDados
    CodigoLista = Lista.SelectedItem.index
End If
TBFIltro.Close
FunFechaBDSite

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcPuxaDados()
On Error GoTo tratar_erro

Set TBMySQL = New ADODB.Recordset
TBMySQL.Open "Select * From Clientes where ID = " & TBFIltro!ID_Cliente & "", ConexaoMySql, adOpenKeyset, adLockOptimistic
If TBMySQL.EOF = False Then
    Set TBFI = CreateObject("adodb.recordset")
    TBFI.Open "Select Empresa from Empresa where CNPJ = '" & TBMySQL!CNPJ & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBFI.EOF = False Then
        Cmb_empresa = TBFI!Empresa
    End If
    TBFI.Close
End If

Txt_ID = TBFIltro!ID
txt_Codigo = TBFIltro!Numero
Txt_data = Format(TBFIltro!Data, "dd/mm/yy")
Txt_responsavel = IIf(IsNull(TBFIltro!Responsavel), "", TBFIltro!Responsavel)
Txt_status = IIf(IsNull(TBFIltro!status), "", TBFIltro!status)
Select Case TBFIltro!Tipo
    Case "B": Cmb_tipo = "Bug"
    Case "D": Cmb_tipo = "Dúvida"
    Case "I": Cmb_tipo = "Instalação"
    Case "R": Cmb_tipo = "Reindexação"
    Case "O": Cmb_tipo = "Outros"
End Select
Txt_setor = IIf(IsNull(TBFIltro!Setor), "", TBFIltro!Setor)
Txt_email.Text = IIf(IsNull(TBFIltro!Email), "", TBFIltro!Email)
If IsNull(TBFIltro!Modulo) = False And TBFIltro!Modulo <> "" Then Cmb_modulo = TBFIltro!Modulo
Txt_descricao = IIf(IsNull(TBFIltro!Descricao), "", TBFIltro!Descricao)

If Txt_status = "Aberto" Then
    Frame1.Enabled = True
    If TBFIltro!Chat_liberado = True Then Frame2.Enabled = True Else Frame2.Enabled = False
Else
    Frame1.Enabled = False
    Frame2.Enabled = False
End If
Novo_Atendimento = False

ProcVerifFiltroMsgem

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Tmr_alarme_nova_mensagem_Timer()
On Error GoTo tratar_erro

With WMP
    .settings.Volume = 100
    .url = Localrel & "\Sons\Nova_mensagem.wav"
End With
Tmr_alarme_nova_mensagem.Enabled = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Tmr_mensagem_Timer()
On Error GoTo tratar_erro

ProcVerifMensagemNova

FunAbreBDSite
Set TBTempo = New ADODB.Recordset
TBTempo.Open "Select ID from Atendimentos_chat where ID_atendimento = " & Txt_ID & " and Respondida IS NULL", ConexaoMySql, adOpenKeyset, adLockReadOnly
If TBTempo.EOF = False Then
    
    'Verifica se o Chat foi liberado e habilita digitação
    Set TBAbrir = New ADODB.Recordset
    TBAbrir.Open "Select ID from Atendimentos where ID = " & Txt_ID & " and Chat_liberado = 1", ConexaoMySql, adOpenKeyset, adLockReadOnly
    If TBAbrir.EOF = False Then Frame2.Enabled = True Else Frame2.Enabled = False
    TBAbrir.Close
    
    ProcVerifFiltroMsgem
End If
TBTempo.Close
FunFechaBDSite

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcVerifMensagemNova()
On Error GoTo tratar_erro

FunAbreBDSite
Set TBTempo = New ADODB.Recordset
TBTempo.Open "Select ID from Atendimentos_chat where ID_atendimento = " & Txt_ID & " and Para = '" & pubUsuario & "' and Respondida IS NULL", ConexaoMySql, adOpenKeyset, adLockReadOnly
If TBTempo.EOF = False Then
    If TBTempo!ID <> IDMsgemNova Then Tmr_alarme_nova_mensagem.Enabled = True
    IDMsgemNova = TBTempo!ID
End If
TBTempo.Close
FunFechaBDSite

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_mensagem_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista_mensagem.ListItems.Count = 0 Then Exit Sub

Nome_anexo = ""
With Cmd_Receber
    FunAbreBDSite
    Set TBClientes = New ADODB.Recordset
    TBClientes.Open "Select Arquivo from Atendimentos_chat where ID = " & Lista_mensagem.SelectedItem & " and Para = '" & pubUsuario & "' and Arquivo IS NOT NULL", ConexaoMySql, adOpenKeyset, adLockOptimistic
    If TBClientes.EOF = False Then
        Nome_anexo = TBClientes!Arquivo
        .Enabled = True
    Else
        .Enabled = False
    End If
    TBClientes.Close
    FunFechaBDSite
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_codigo_Change()
On Error GoTo tratar_erro

If Novo_Atendimento = True Then
VerifCodigo:
    FunAbreBDSite
    Set TBFI = CreateObject("adodb.recordset")
    TBFI.Open "Select * from Chamado where ID_cliente = " & FunVerifIDEmpresaSite(Cmb_empresa.ItemData(Cmb_empresa.ListIndex)) & " and Numero = '" & txt_Codigo & "' and ID <> " & Txt_ID & "", ConexaoMySql, adOpenKeyset, adLockOptimistic
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

Private Sub Txt_mensagem_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro
Dim Utilizado As Integer
Dim Restam As Integer

Utilizado = Len(Txt_mensagem.Text)
Restam = 500 - Utilizado
lblUtilizado.Caption = Utilizado
LblRestam.Caption = Restam

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
    Case 4: ProcStatus
    Case 6: 'ProcAjuda
    Case 7: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
