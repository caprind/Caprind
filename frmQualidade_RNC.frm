VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmQualidade_RNC 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Qualidade - RNC"
   ClientHeight    =   10035
   ClientLeft      =   60
   ClientTop       =   3870
   ClientWidth     =   15240
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10035
   ScaleWidth      =   15240
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   0
      Top             =   0
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
      FormWidthDT     =   15360
      FormScaleHeightDT=   10035
      FormScaleWidthDT=   15240
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   135
      TabIndex        =   118
      Top             =   9705
      Width           =   15015
      _ExtentX        =   26485
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   10035
      Left            =   0
      TabIndex        =   62
      Top             =   0
      Width           =   15360
      _ExtentX        =   27093
      _ExtentY        =   17701
      _Version        =   393216
      Tabs            =   4
      Tab             =   2
      TabsPerRow      =   4
      TabHeight       =   520
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
      TabCaption(0)   =   "Abertura e fechamento"
      TabPicture(0)   =   "frmQualidade_RNC.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtTipo"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtid"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "ListView1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "USToolBar1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Contenção"
      TabPicture(1)   =   "frmQualidade_RNC.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "USToolBar2"
      Tab(1).Control(1)=   "Framelista"
      Tab(1).Control(2)=   "USImageList2"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Causa e ações"
      TabPicture(2)   =   "frmQualidade_RNC.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "USToolBar3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "SSTab2"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "USImageList3"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Documentos e assinaturas"
      TabPicture(3)   =   "frmQualidade_RNC.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Lista_doc"
      Tab(3).Control(1)=   "USToolBar4"
      Tab(3).Control(2)=   "CommonDialog1"
      Tab(3).Control(3)=   "txtID_doc"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Frame14"
      Tab(3).Control(5)=   "Frame3"
      Tab(3).ControlCount=   6
      Begin VB.Frame Frame3 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   -74910
         TabIndex        =   168
         Top             =   8400
         Width           =   15165
         Begin VB.CommandButton cmdSalvarRespRNC 
            BackColor       =   &H00E0E0E0&
            Height          =   315
            Left            =   13860
            Picture         =   "frmQualidade_RNC.frx":0070
            Style           =   1  'Graphical
            TabIndex        =   175
            Top             =   570
            Width           =   405
         End
         Begin VB.TextBox txtAuditor 
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
            Left            =   480
            Locked          =   -1  'True
            TabIndex        =   172
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pelo cadastro."
            Top             =   570
            Width           =   6135
         End
         Begin VB.TextBox txtRespQualidade 
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
            Left            =   7320
            Locked          =   -1  'True
            TabIndex        =   171
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pelo cadastro."
            Top             =   570
            Width           =   6135
         End
         Begin VB.CommandButton cmdAuditorLider 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   6630
            Picture         =   "frmQualidade_RNC.frx":00C3
            Style           =   1  'Graphical
            TabIndex        =   170
            ToolTipText     =   "Visualizar arquivo."
            Top             =   570
            Width           =   345
         End
         Begin VB.CommandButton cmdResponsavelQualidade 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   13470
            Picture         =   "frmQualidade_RNC.frx":0685
            Style           =   1  'Graphical
            TabIndex        =   169
            ToolTipText     =   "Visualizar arquivo."
            Top             =   570
            Width           =   345
         End
         Begin VB.Label Label45 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Auditor Lider"
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
            TabIndex        =   174
            Top             =   360
            Width           =   915
         End
         Begin VB.Label Label46 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Responsável pela qualidade"
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
            Left            =   9390
            TabIndex        =   173
            Top             =   360
            Width           =   1995
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   -74955
         TabIndex        =   120
         Top             =   9110
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
            TabIndex        =   26
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
            TabIndex        =   25
            Text            =   "30"
            ToolTipText     =   "Número de registros por página."
            Top             =   180
            Width           =   555
         End
         Begin DrawSuite2022.USButton cmdPagProx 
            Height          =   315
            Left            =   11760
            TabIndex        =   30
            ToolTipText     =   "Próxima página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmQualidade_RNC.frx":0C47
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
            TabIndex        =   29
            ToolTipText     =   "Página anterior."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmQualidade_RNC.frx":43EE
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
            TabIndex        =   27
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
            TabIndex        =   28
            ToolTipText     =   "Primeira página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmQualidade_RNC.frx":7EFB
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
            TabIndex        =   31
            ToolTipText     =   "Última página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmQualidade_RNC.frx":BFED
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
         Begin VB.Label Label33 
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
            TabIndex        =   132
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
            TabIndex        =   123
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
            TabIndex        =   122
            Top             =   240
            Width           =   1275
         End
         Begin VB.Label Label3 
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
            TabIndex        =   121
            Top             =   240
            Width           =   645
         End
      End
      Begin VB.Frame Frame14 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2445
         Left            =   -74925
         TabIndex        =   115
         Top             =   1320
         Width           =   15195
         Begin VB.CommandButton Cmd_limpar_caminho 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   14370
            Picture         =   "frmQualidade_RNC.frx":F87A
            Style           =   1  'Graphical
            TabIndex        =   127
            ToolTipText     =   "Limpar caminho."
            Top             =   540
            Width           =   315
         End
         Begin VB.CommandButton Cmd_visualizar_arquivo 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   14700
            Picture         =   "frmQualidade_RNC.frx":F9B8
            Style           =   1  'Graphical
            TabIndex        =   126
            ToolTipText     =   "Visualizar arquivo."
            Top             =   540
            Width           =   315
         End
         Begin VB.TextBox txtResponsavel_doc 
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
            Left            =   1390
            Locked          =   -1  'True
            TabIndex        =   57
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pelo cadastro."
            Top             =   540
            Width           =   2745
         End
         Begin VB.CommandButton cmdImportar 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   14040
            Picture         =   "frmQualidade_RNC.frx":FF7A
            Style           =   1  'Graphical
            TabIndex        =   59
            ToolTipText     =   "Localizar arquivo (F2)"
            Top             =   540
            Width           =   315
         End
         Begin VB.TextBox Txt_caminho 
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
            Left            =   4150
            Locked          =   -1  'True
            MaxLength       =   255
            TabIndex        =   58
            TabStop         =   0   'False
            ToolTipText     =   "Caminho."
            Top             =   540
            Width           =   9885
         End
         Begin VB.TextBox Txt_obs_doc 
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
            Left            =   180
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   60
            ToolTipText     =   "Observação."
            Top             =   1350
            Width           =   14835
         End
         Begin MSComCtl2.DTPicker txtData_doc 
            Height          =   315
            Left            =   180
            TabIndex        =   56
            ToolTipText     =   "Data do cadastro."
            Top             =   540
            Width           =   1215
            _ExtentX        =   2143
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
            Format          =   172687363
            CurrentDate     =   39057
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Caminho do arquivo*"
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
            Left            =   8340
            TabIndex        =   131
            Top             =   330
            Width           =   1515
         End
         Begin VB.Label Label31 
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
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   2310
            TabIndex        =   130
            Top             =   330
            Width           =   915
         End
         Begin VB.Label Label42 
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
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   600
            TabIndex        =   119
            Top             =   330
            Width           =   345
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Observações"
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
            Left            =   7155
            TabIndex        =   116
            Top             =   1080
            Width           =   945
         End
      End
      Begin VB.TextBox txtID_doc 
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
         Height          =   315
         Left            =   -70335
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   114
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   5700
         Visible         =   0   'False
         Width           =   675
      End
      Begin DrawSuite2022.USImageList USImageList3 
         Left            =   8610
         Top             =   480
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmQualidade_RNC.frx":1007C
         Count           =   1
      End
      Begin DrawSuite2022.USImageList USImageList2 
         Left            =   -68070
         Top             =   450
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmQualidade_RNC.frx":14B32
         Count           =   1
      End
      Begin VB.TextBox txtTipo 
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
         Height          =   315
         Left            =   -73440
         Locked          =   -1  'True
         MouseIcon       =   "frmQualidade_RNC.frx":195E8
         MousePointer    =   99  'Custom
         TabIndex        =   83
         TabStop         =   0   'False
         Top             =   7560
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.TextBox txtid 
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
         Height          =   315
         Left            =   -72660
         Locked          =   -1  'True
         MouseIcon       =   "frmQualidade_RNC.frx":198F2
         MousePointer    =   99  'Custom
         TabIndex        =   81
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   7560
         Visible         =   0   'False
         Width           =   795
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2040
         Left            =   -74955
         TabIndex        =   24
         Top             =   6990
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   3598
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
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
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Object.Tag             =   "T"
            Text            =   "Nº RNC"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Object.Tag             =   "N"
            Text            =   "Seq."
            Object.Width           =   970
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Object.Tag             =   "D"
            Text            =   "Data"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "Responsável"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Object.Tag             =   "T"
            Text            =   "Cód. interno"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Object.Tag             =   "T"
            Text            =   "Descrição"
            Object.Width           =   12621
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Object.Tag             =   "N"
            Text            =   "Qtde. NC"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Custo da RNC"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Frame Framelista 
         BackColor       =   &H00E0E0E0&
         Height          =   8685
         Left            =   -74925
         TabIndex        =   68
         Top             =   1320
         Width           =   15195
         Begin VB.TextBox txtResponsavel2 
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
            Left            =   1470
            MaxLength       =   50
            TabIndex        =   33
            ToolTipText     =   "Responsável."
            Top             =   375
            Width           =   13515
         End
         Begin VB.TextBox txtTexto2 
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
            Height          =   3780
            Left            =   180
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   35
            ToolTipText     =   "Abrangência na contenção."
            Top             =   4785
            Width           =   14805
         End
         Begin VB.TextBox txtTexto1 
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
            Height          =   3435
            Left            =   180
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   34
            ToolTipText     =   "Ações de contenção (disposição imediata)."
            Top             =   990
            Width           =   14805
         End
         Begin MSComCtl2.DTPicker txtData2 
            Height          =   315
            Left            =   210
            TabIndex        =   32
            ToolTipText     =   "Data."
            Top             =   375
            Width           =   1245
            _ExtentX        =   2196
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
            Format          =   197984259
            CurrentDate     =   39057
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
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
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   630
            TabIndex        =   77
            Top             =   180
            Width           =   345
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
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
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   7770
            TabIndex        =   76
            Top             =   180
            Width           =   915
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Ações de contenção (disposição imediata)"
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
            Index           =   2
            Left            =   6082
            TabIndex        =   74
            Top             =   780
            Width           =   3000
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Abrangência na contenção"
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
            Index           =   1
            Left            =   6622
            TabIndex        =   69
            Top             =   4560
            Width           =   1920
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   6585
         Left            =   -74955
         TabIndex        =   63
         Top             =   1320
         Width           =   15195
         Begin VB.CommandButton cmdValorUnitario 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   6990
            Picture         =   "frmQualidade_RNC.frx":19BFC
            Style           =   1  'Graphical
            TabIndex        =   166
            ToolTipText     =   "Visualizar arquivo."
            Top             =   5130
            Width           =   345
         End
         Begin VB.TextBox txtCustoNC 
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
            Left            =   8760
            Locked          =   -1  'True
            TabIndex        =   164
            ToolTipText     =   "custo da NC"
            Top             =   5130
            Width           =   1335
         End
         Begin VB.TextBox txtFatorNC 
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
            Left            =   7380
            TabIndex        =   162
            Text            =   "1"
            ToolTipText     =   "Fator para calculo do custo da NC"
            Top             =   5130
            Width           =   1335
         End
         Begin VB.TextBox txtValorUnitario 
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
            Left            =   5640
            Locked          =   -1  'True
            TabIndex        =   160
            ToolTipText     =   "Valor unitario do item"
            Top             =   5130
            Width           =   1335
         End
         Begin VB.CommandButton cmdImagemRNC 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   14010
            Picture         =   "frmQualidade_RNC.frx":1A1BE
            Style           =   1  'Graphical
            TabIndex        =   158
            ToolTipText     =   "Localizar arquivo (F2)"
            Top             =   4380
            Width           =   315
         End
         Begin VB.TextBox Txt_imagemRNC 
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
            TabIndex        =   157
            TabStop         =   0   'False
            ToolTipText     =   "Imagem da RNC"
            Top             =   4380
            Width           =   13815
         End
         Begin VB.CommandButton Cmd_visualizar_arquivo1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   14670
            Picture         =   "frmQualidade_RNC.frx":1A2C0
            Style           =   1  'Graphical
            TabIndex        =   156
            ToolTipText     =   "Visualizar arquivo."
            Top             =   4380
            Width           =   315
         End
         Begin VB.CommandButton Cmd_limpar_caminho1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   14340
            Picture         =   "frmQualidade_RNC.frx":1A882
            Style           =   1  'Graphical
            TabIndex        =   155
            ToolTipText     =   "Limpar caminho."
            Top             =   4380
            Width           =   315
         End
         Begin VB.ComboBox cmbMateriaPrima 
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
            ItemData        =   "frmQualidade_RNC.frx":1A9C0
            Left            =   4140
            List            =   "frmQualidade_RNC.frx":1A9CD
            Style           =   2  'Dropdown List
            TabIndex        =   144
            ToolTipText     =   "Status."
            Top             =   5130
            Width           =   1485
         End
         Begin VB.TextBox txtQuaisRNC 
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
            Left            =   7800
            TabIndex        =   140
            ToolTipText     =   "Documento de referência."
            Top             =   3000
            Width           =   7215
         End
         Begin VB.ComboBox cmbSimilaridade 
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
            ItemData        =   "frmQualidade_RNC.frx":1A9E9
            Left            =   6630
            List            =   "frmQualidade_RNC.frx":1A9F6
            Style           =   2  'Dropdown List
            TabIndex        =   138
            ToolTipText     =   "Eficaz."
            Top             =   3000
            Width           =   1155
         End
         Begin VB.TextBox txtQuaisRequisitos 
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
            Left            =   2160
            TabIndex        =   137
            ToolTipText     =   "Documento de referência."
            Top             =   3000
            Width           =   4455
         End
         Begin VB.ComboBox cmbRequisitos 
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
            ItemData        =   "frmQualidade_RNC.frx":1AA06
            Left            =   1170
            List            =   "frmQualidade_RNC.frx":1AA13
            Style           =   2  'Dropdown List
            TabIndex        =   135
            ToolTipText     =   "Eficaz."
            Top             =   3000
            Width           =   975
         End
         Begin VB.ComboBox cmbProcede 
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
            ItemData        =   "frmQualidade_RNC.frx":1AA23
            Left            =   210
            List            =   "frmQualidade_RNC.frx":1AA30
            Style           =   2  'Dropdown List
            TabIndex        =   133
            ToolTipText     =   "Eficaz."
            Top             =   3000
            Width           =   975
         End
         Begin VB.CheckBox Chk_acao_corretiva 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Implementar ação corretiva"
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
            Left            =   12300
            TabIndex        =   128
            Top             =   405
            Value           =   1  'Checked
            Width           =   2835
         End
         Begin VB.TextBox txt_qtdeLote 
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
            Left            =   180
            TabIndex        =   20
            ToolTipText     =   "Quantidade do lote."
            Top             =   5130
            Width           =   1245
         End
         Begin VB.TextBox txt_QtdeAprovada 
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
            Left            =   1440
            TabIndex        =   21
            ToolTipText     =   "Quantidade aprovada."
            Top             =   5130
            Width           =   1335
         End
         Begin VB.CommandButton cmdClassificacao 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   14700
            Picture         =   "frmQualidade_RNC.frx":1AA40
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Localizar funções."
            Top             =   990
            Width           =   315
         End
         Begin VB.TextBox txtClassificacao 
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
            Left            =   11520
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   12
            TabStop         =   0   'False
            ToolTipText     =   "Classificação."
            Top             =   990
            Width           =   3165
         End
         Begin VB.Frame Frame8 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Tipo"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   7890
            TabIndex        =   86
            Top             =   180
            Width           =   2265
            Begin VB.OptionButton optCorretiva 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Corretiva"
               DisabledPicture =   "frmQualidade_RNC.frx":1AB42
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
               Left            =   120
               TabIndex        =   5
               Top             =   240
               Value           =   -1  'True
               Width           =   1005
            End
            Begin VB.OptionButton optPreventiva 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Preventiva"
               DisabledPicture =   "frmQualidade_RNC.frx":264A84
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
               Left            =   1125
               TabIndex        =   6
               Top             =   240
               Width           =   1095
            End
         End
         Begin VB.TextBox txtEquipe 
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
            Left            =   180
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   23
            ToolTipText     =   "Formação da equipe (nome dos envolvidos)"
            Top             =   3720
            Width           =   14835
         End
         Begin VB.CommandButton Cmd_localizar_cliente_fornecedor 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   14700
            Picture         =   "frmQualidade_RNC.frx":4AE9C6
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Localizar cliente/fornecedor."
            Top             =   1590
            Width           =   315
         End
         Begin VB.CommandButton Cmd_localizar_item 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   1920
            Picture         =   "frmQualidade_RNC.frx":4AEAC8
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Localizar produtos."
            Top             =   990
            Width           =   315
         End
         Begin VB.TextBox Txt_doc_referencia 
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
            Left            =   12750
            TabIndex        =   18
            ToolTipText     =   "Documento de referência."
            Top             =   5130
            Width           =   2235
         End
         Begin VB.Frame Frame7 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Origem"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   10200
            TabIndex        =   82
            Top             =   180
            Width           =   1965
            Begin VB.OptionButton Opt_externo 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Externo"
               DisabledPicture =   "frmQualidade_RNC.frx":4AEBCA
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
               Height          =   210
               Left            =   1005
               TabIndex        =   8
               Top             =   240
               Width           =   885
            End
            Begin VB.OptionButton Opt_interno 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Interno"
               DisabledPicture =   "frmQualidade_RNC.frx":6F8B0C
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
               Left            =   120
               TabIndex        =   7
               Top             =   240
               Value           =   -1  'True
               Width           =   855
            End
         End
         Begin VB.ComboBox cmbEficaz 
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
            ItemData        =   "frmQualidade_RNC.frx":942A4E
            Left            =   11670
            List            =   "frmQualidade_RNC.frx":942A5B
            Style           =   2  'Dropdown List
            TabIndex        =   17
            ToolTipText     =   "Eficaz."
            Top             =   5130
            Width           =   1065
         End
         Begin VB.ComboBox cmbStatus 
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
            ItemData        =   "frmQualidade_RNC.frx":942A6B
            Left            =   6435
            List            =   "frmQualidade_RNC.frx":942A7B
            Style           =   2  'Dropdown List
            TabIndex        =   4
            ToolTipText     =   "Status."
            Top             =   390
            Width           =   1365
         End
         Begin VB.TextBox txtID_forn 
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
            Height          =   315
            Left            =   180
            Locked          =   -1  'True
            TabIndex        =   98
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   1590
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.TextBox txtFornecedor 
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
            Left            =   180
            MaxLength       =   255
            TabIndex        =   14
            ToolTipText     =   "Cliente/fornecedor."
            Top             =   1590
            Width           =   14505
         End
         Begin VB.TextBox txtdescricao 
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
            Height          =   315
            Left            =   2330
            MaxLength       =   255
            TabIndex        =   11
            ToolTipText     =   "Descrição."
            Top             =   990
            Width           =   9180
         End
         Begin VB.TextBox txtqtde 
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
            Left            =   2790
            TabIndex        =   22
            ToolTipText     =   "Quantidade não conforme."
            Top             =   5130
            Width           =   1335
         End
         Begin VB.TextBox txtdesenho 
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
            Height          =   315
            Left            =   180
            MaxLength       =   50
            TabIndex        =   9
            ToolTipText     =   "Código interno."
            Top             =   990
            Width           =   1725
         End
         Begin VB.TextBox txtResponsavel 
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
            Left            =   2835
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   3
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pelo cadastro."
            Top             =   390
            Width           =   3585
         End
         Begin VB.TextBox txtNao_conformidade 
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
            Height          =   435
            Left            =   180
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   19
            ToolTipText     =   "Descrição da não conformidade."
            Top             =   2190
            Width           =   14835
         End
         Begin MSComCtl2.DTPicker txtData 
            Height          =   315
            Left            =   1590
            TabIndex        =   2
            ToolTipText     =   "Data do cadastro."
            Top             =   390
            Width           =   1245
            _ExtentX        =   2196
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
            Format          =   143589379
            CurrentDate     =   39057
         End
         Begin MSMask.MaskEdBox txtfim 
            Height          =   315
            Left            =   10140
            TabIndex        =   16
            ToolTipText     =   "Data de fechamento."
            Top             =   5130
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtID_texto 
            Height          =   315
            Left            =   180
            TabIndex        =   0
            ToolTipText     =   "Número da RNC."
            Top             =   390
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            MaxLength       =   7
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "####-##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox Txt_sequencial 
            Height          =   315
            Left            =   1080
            TabIndex        =   1
            ToolTipText     =   "Sequencial do número da RNC."
            Top             =   390
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            MaxLength       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##"
            PromptChar      =   "_"
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Custo da NC"
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
            Left            =   8985
            TabIndex        =   165
            Top             =   4920
            Width           =   900
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Fator da NC"
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
            Left            =   7575
            TabIndex        =   163
            Top             =   4920
            Width           =   870
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Valor unitario"
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
            Left            =   5835
            TabIndex        =   161
            Top             =   4920
            Width           =   945
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Caminho da imagem"
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
            Left            =   6375
            TabIndex        =   159
            Top             =   4170
            Width           =   1425
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Materia prima"
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
            TabIndex        =   143
            Top             =   4920
            Width           =   975
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Quais RNCs?"
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
            Left            =   10830
            TabIndex        =   142
            Top             =   2790
            Width           =   915
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Quais requisitos?"
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
            Left            =   3690
            TabIndex        =   141
            Top             =   2790
            Width           =   1215
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Similaridade"
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
            Left            =   6780
            TabIndex        =   139
            Top             =   2790
            Width           =   840
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Requisitos"
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
            Left            =   1290
            TabIndex        =   136
            Top             =   2790
            Width           =   735
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Procede"
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
            TabIndex        =   134
            Top             =   2790
            Width           =   585
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Seq.*"
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
            Left            =   1095
            TabIndex        =   129
            Top             =   180
            Width           =   465
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Qtde. lote"
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
            Left            =   435
            TabIndex        =   125
            Top             =   4920
            Width           =   735
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Qtde. aprovada"
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
            Left            =   1530
            TabIndex        =   124
            Top             =   4920
            Width           =   1155
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Classificação"
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
            Left            =   12645
            TabIndex        =   110
            Top             =   780
            Width           =   915
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Formação da equipe (nome dos envolvidos)"
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
            Height          =   285
            Left            =   6030
            TabIndex        =   85
            Top             =   3480
            Width           =   3120
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Doc. de referência"
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
            Left            =   13200
            TabIndex        =   84
            Top             =   4920
            Width           =   1335
         End
         Begin VB.Image Imgcalendario1 
            Height          =   360
            Left            =   11355
            Picture         =   "frmQualidade_RNC.frx":942AA5
            Stretch         =   -1  'True
            ToolTipText     =   "Abrir calendário."
            Top             =   5100
            Width           =   330
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Eficaz"
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
            Left            =   11985
            TabIndex        =   80
            Top             =   4920
            Width           =   420
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Fechamento"
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
            Left            =   10305
            TabIndex        =   79
            Top             =   4920
            Width           =   885
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
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
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   6885
            TabIndex        =   78
            Top             =   180
            Width           =   465
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Cliente/fornecedor"
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
            Left            =   6757
            TabIndex        =   75
            Top             =   1380
            Width           =   1350
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Qtde. NC"
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
            Left            =   3120
            TabIndex        =   73
            Top             =   4920
            Width           =   675
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Descrição"
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
            Left            =   6508
            TabIndex        =   72
            Top             =   780
            Width           =   825
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Código interno*"
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
            Left            =   420
            TabIndex        =   71
            Top             =   780
            Width           =   1335
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Nº RNC*"
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
            Left            =   285
            TabIndex        =   70
            Top             =   180
            Width           =   675
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
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
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   4170
            TabIndex        =   67
            Top             =   180
            Width           =   915
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
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
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   2040
            TabIndex        =   66
            Top             =   180
            Width           =   345
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackColor       =   &H80000001&
            Caption         =   "Nº:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   -11580
            TabIndex        =   65
            Top             =   4200
            Width           =   270
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Descrição da não conformidade"
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
            Left            =   6472
            TabIndex        =   64
            Top             =   1980
            Width           =   2250
         End
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   8775
         Left            =   75
         TabIndex        =   87
         Top             =   1320
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   15478
         _Version        =   393216
         Tabs            =   2
         Tab             =   1
         TabsPerRow      =   2
         TabHeight       =   520
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
         TabCaption(0)   =   "Determinação da causa"
         TabPicture(0)   =   "frmQualidade_RNC.frx":942F28
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Frame6"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Ações"
         TabPicture(1)   =   "frmQualidade_RNC.frx":942F44
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Frame9"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         Begin VB.Frame Frame9 
            BackColor       =   &H00E0E0E0&
            Height          =   8385
            Left            =   30
            TabIndex        =   95
            Top             =   330
            Width           =   15105
            Begin VB.TextBox txtRevisao 
               Enabled         =   0   'False
               Height          =   315
               Left            =   9930
               TabIndex        =   154
               Top             =   7530
               Width           =   4965
            End
            Begin VB.CheckBox chkRevisao 
               BackColor       =   &H00E0E0E0&
               Caption         =   "As ações implementadas sugerem uma revisão na planilha de gestão de riscos.     Item que foi promovida a revisão"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   255
               Left            =   180
               TabIndex        =   153
               Top             =   7560
               Width           =   9705
            End
            Begin VB.TextBox txtResponsavel6 
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
               Height          =   345
               Left            =   6450
               MaxLength       =   50
               TabIndex        =   51
               ToolTipText     =   "Responsável."
               Top             =   6870
               Width           =   3555
            End
            Begin VB.TextBox txtResponsavel7 
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
               Left            =   1530
               MaxLength       =   50
               TabIndex        =   48
               ToolTipText     =   "Responsável."
               Top             =   6885
               Width           =   3555
            End
            Begin VB.TextBox txtResponsavel5 
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
               Left            =   11370
               MaxLength       =   50
               TabIndex        =   54
               ToolTipText     =   "Responsável."
               Top             =   6870
               Width           =   3555
            End
            Begin VB.TextBox txtTexto8 
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
               Height          =   525
               Left            =   10020
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   55
               ToolTipText     =   "Abrangência na correção/fechamento."
               Top             =   6090
               Width           =   4905
            End
            Begin VB.TextBox txtTexto7 
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
               TabIndex        =   46
               ToolTipText     =   "Ações corretivas/preventivas"
               Top             =   2745
               Width           =   14685
            End
            Begin VB.TextBox txtAcompanhamento 
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
               Height          =   525
               Left            =   180
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   49
               ToolTipText     =   "Verificação de implementação."
               Top             =   6090
               Width           =   4905
            End
            Begin VB.TextBox txtEficacia 
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
               Height          =   525
               Left            =   5100
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   52
               ToolTipText     =   "Verificação da eficácia."
               Top             =   6090
               Width           =   4905
            End
            Begin VB.OptionButton optNao 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Não"
               DisabledPicture =   "frmQualidade_RNC.frx":942F60
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
               Left            =   2295
               TabIndex        =   45
               Top             =   330
               Width           =   585
            End
            Begin VB.OptionButton optSim 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Sim"
               DisabledPicture =   "frmQualidade_RNC.frx":B8CEA2
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
               Left            =   1440
               TabIndex        =   44
               Top             =   330
               Value           =   -1  'True
               Width           =   585
            End
            Begin VB.TextBox txtResponsavel4 
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
               Left            =   1410
               MaxLength       =   50
               TabIndex        =   43
               ToolTipText     =   "Responsável."
               Top             =   5175
               Width           =   13485
            End
            Begin MSComCtl2.DTPicker txtData4 
               Height          =   315
               Left            =   180
               TabIndex        =   42
               ToolTipText     =   "Prazo."
               Top             =   5175
               Width           =   1215
               _ExtentX        =   2143
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
               Format          =   198574083
               CurrentDate     =   39057
            End
            Begin MSMask.MaskEdBox txtData5 
               Height          =   345
               Left            =   5100
               TabIndex        =   50
               ToolTipText     =   "Data."
               Top             =   6870
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   609
               _Version        =   393216
               BackColor       =   16777215
               MaxLength       =   10
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox txtData6 
               Height          =   345
               Left            =   10020
               TabIndex        =   53
               ToolTipText     =   "Data."
               Top             =   6870
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   609
               _Version        =   393216
               BackColor       =   16777215
               MaxLength       =   10
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox txtdata7 
               Height          =   315
               Left            =   180
               TabIndex        =   47
               ToolTipText     =   "Data."
               Top             =   6885
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
               MaxLength       =   10
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin MSComctlLib.ListView ListView2 
               Height          =   2040
               Left            =   210
               TabIndex        =   167
               Top             =   570
               Width           =   14685
               _ExtentX        =   25903
               _ExtentY        =   3598
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   0   'False
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
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
               NumItems        =   5
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Object.Tag             =   "T"
                  Text            =   "SA"
                  Object.Width           =   2117
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   2
                  SubItemIndex    =   1
                  Object.Tag             =   "N"
                  Text            =   "Tipo"
                  Object.Width           =   1764
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   2
                  SubItemIndex    =   2
                  Object.Tag             =   "D"
                  Text            =   "Prazo"
                  Object.Width           =   2117
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   3
                  Object.Tag             =   "T"
                  Text            =   "Responsável"
                  Object.Width           =   3528
               EndProperty
               BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   4
                  Object.Tag             =   "T"
                  Text            =   "Objetivo"
                  Object.Width           =   52917
               EndProperty
            End
            Begin VB.Line Line1 
               BorderColor     =   &H00C0C0C0&
               X1              =   210
               X2              =   14850
               Y1              =   5670
               Y2              =   5670
            End
            Begin VB.Label Label27 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
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
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   7770
               TabIndex        =   113
               Top             =   6660
               Width           =   915
            End
            Begin VB.Image Imgcalendario2 
               Height          =   360
               Left            =   1215
               MouseIcon       =   "frmQualidade_RNC.frx":DD6DE4
               MousePointer    =   99  'Custom
               Picture         =   "frmQualidade_RNC.frx":DD70EE
               Stretch         =   -1  'True
               ToolTipText     =   "Abrir calendário."
               Top             =   6855
               Width           =   330
            End
            Begin VB.Image Imgcalendario4 
               Height          =   360
               Left            =   11055
               Picture         =   "frmQualidade_RNC.frx":DD7571
               Stretch         =   -1  'True
               ToolTipText     =   "Abrir calendário."
               Top             =   6840
               Width           =   330
            End
            Begin VB.Image Imgcalendario3 
               Height          =   360
               Left            =   6135
               Picture         =   "frmQualidade_RNC.frx":DD79F4
               Stretch         =   -1  'True
               ToolTipText     =   "Abrir calendário."
               Top             =   6840
               Width           =   330
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Data*"
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
               Index           =   13
               Left            =   525
               TabIndex        =   108
               Top             =   6660
               Width           =   435
            End
            Begin VB.Label Label25 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
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
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   2850
               TabIndex        =   107
               Top             =   6660
               Width           =   915
            End
            Begin VB.Label Label24 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
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
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   12690
               TabIndex        =   106
               Top             =   6660
               Width           =   915
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Data*"
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
               Index           =   6
               Left            =   10395
               TabIndex        =   105
               Top             =   6660
               Width           =   435
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Verificação de implementação"
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
               Left            =   1350
               TabIndex        =   104
               Top             =   5880
               Width           =   2550
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Data*"
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
               Height          =   285
               Index           =   11
               Left            =   5445
               TabIndex        =   103
               Top             =   6660
               Width           =   435
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Verificação da eficácia"
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
               Left            =   6615
               TabIndex        =   102
               Top             =   5880
               Width           =   1875
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Abrangência na correção/fechamento"
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
               Left            =   10860
               TabIndex        =   101
               Top             =   5880
               Width           =   3225
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Requer ação?"
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
               Index           =   12
               Left            =   270
               TabIndex        =   100
               Top             =   330
               Width           =   990
            End
            Begin VB.Label Label23 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Ações corretivas e preventivas"
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
               Left            =   6210
               TabIndex        =   99
               Top             =   300
               Width           =   2655
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Prazo*"
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
               Index           =   10
               Left            =   600
               TabIndex        =   97
               Top             =   4980
               Width           =   495
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Responsável pelo acompanhamento das SAs"
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
               Left            =   6555
               TabIndex        =   96
               Top             =   4920
               Width           =   3195
            End
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00E0E0E0&
            Height          =   8385
            Left            =   -74970
            TabIndex        =   88
            Top             =   330
            Width           =   15105
            Begin VB.CheckBox chkOutros 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Outros"
               Height          =   225
               Left            =   12210
               TabIndex        =   152
               Top             =   960
               Width           =   1575
            End
            Begin VB.CheckBox chkSGQ 
               BackColor       =   &H00E0E0E0&
               Caption         =   "SGQ"
               Height          =   225
               Left            =   10920
               TabIndex        =   151
               Top             =   960
               Width           =   1575
            End
            Begin VB.CheckBox chkMaterial 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Material"
               Height          =   225
               Left            =   7500
               TabIndex        =   150
               Top             =   960
               Width           =   1575
            End
            Begin VB.CheckBox chkMeioAmbiente 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Meio ambiente"
               Height          =   225
               Left            =   9090
               TabIndex        =   149
               Top             =   960
               Width           =   1575
            End
            Begin VB.CheckBox chkMedidas 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Medidas"
               Height          =   225
               Left            =   6060
               TabIndex        =   148
               Top             =   960
               Width           =   1575
            End
            Begin VB.CheckBox chkMetodo 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Método"
               Height          =   225
               Left            =   4650
               TabIndex        =   147
               Top             =   960
               Width           =   1575
            End
            Begin VB.CheckBox chkMaquina 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Máquina"
               Height          =   225
               Left            =   3240
               TabIndex        =   146
               Top             =   960
               Width           =   1575
            End
            Begin VB.CheckBox chkMaoObra 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Mão de obra"
               Height          =   225
               Left            =   1620
               TabIndex        =   145
               Top             =   960
               Width           =   1575
            End
            Begin VB.TextBox txtTexto6 
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
               Height          =   2715
               Left            =   9990
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   41
               ToolTipText     =   "Por que passou?."
               Top             =   5550
               Width           =   4905
            End
            Begin VB.TextBox txtTexto5 
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
               Height          =   2715
               Left            =   5115
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   40
               ToolTipText     =   "Por que?."
               Top             =   5550
               Width           =   4845
            End
            Begin VB.TextBox txtTexto3 
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
               Height          =   3165
               Left            =   180
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   38
               ToolTipText     =   "Determinação da causa."
               Top             =   1680
               Width           =   14715
            End
            Begin VB.TextBox txtTexto4 
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
               Height          =   2715
               Left            =   180
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   39
               ToolTipText     =   "Por que?."
               Top             =   5550
               Width           =   4905
            End
            Begin VB.TextBox txtResponsavel3 
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
               Left            =   1470
               MaxLength       =   50
               TabIndex        =   37
               ToolTipText     =   "Responsável."
               Top             =   375
               Width           =   13425
            End
            Begin MSComCtl2.DTPicker txtData3 
               Height          =   315
               Left            =   180
               TabIndex        =   36
               ToolTipText     =   "Data."
               Top             =   375
               Width           =   1245
               _ExtentX        =   2196
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
               Format          =   198639619
               CurrentDate     =   39057
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Por que passou?"
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
               Index           =   4
               Left            =   11850
               TabIndex        =   94
               Top             =   5340
               Width           =   1185
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Por que?"
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
               Index           =   3
               Left            =   2317
               TabIndex        =   93
               Top             =   5340
               Width           =   630
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Determinação da causa"
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
               Index           =   0
               Left            =   6690
               TabIndex        =   92
               Top             =   1440
               Width           =   1995
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
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
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   630
               TabIndex        =   91
               Top             =   180
               Width           =   345
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
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
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   7725
               TabIndex        =   90
               Top             =   180
               Width           =   915
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Por que?"
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
               Index           =   5
               Left            =   7222
               TabIndex        =   89
               Top             =   5340
               Width           =   630
            End
         End
      End
      Begin DrawSuite2022.USToolBar USToolBar1 
         Height          =   975
         Left            =   -74925
         TabIndex        =   109
         Top             =   330
         Width           =   15195
         _ExtentX        =   26802
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
         ButtonCaption8  =   "Atualizar"
         ButtonEnabled8  =   0   'False
         ButtonIconSize8 =   32
         ButtonToolTipText8=   "Utilizado pelo administrador do sistema."
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
         ButtonWidth8    =   59
         ButtonHeight8   =   21
         ButtonUseMaskColor8=   0   'False
         ButtonEnabled9  =   0   'False
         ButtonIconSize9 =   32
         ButtonAlignment9=   2
         ButtonType9     =   1
         ButtonStyle9    =   -1
         BeginProperty ButtonFont9 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState9    =   -1
         ButtonLeft9     =   414
         ButtonTop9      =   4
         ButtonWidth9    =   2
         ButtonHeight9   =   54
         ButtonCaption10 =   "Ajuda"
         ButtonEnabled10 =   0   'False
         ButtonIconSize10=   32
         ButtonToolTipText10=   "Ajuda (F1)"
         ButtonKey10     =   "10"
         ButtonAlignment10=   2
         BeginProperty ButtonFont10 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft10    =   418
         ButtonTop10     =   2
         ButtonWidth10   =   41
         ButtonHeight10  =   21
         ButtonUseMaskColor10=   0   'False
         ButtonCaption11 =   "Sair"
         ButtonEnabled11 =   0   'False
         ButtonIconSize11=   32
         ButtonToolTipText11=   "Sair (Esc)"
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
         ButtonLeft11    =   461
         ButtonTop11     =   2
         ButtonWidth11   =   30
         ButtonHeight11  =   21
         ButtonUseMaskColor11=   0   'False
         ButtonEnabled12 =   0   'False
         ButtonIconSize12=   32
         ButtonKey12     =   "12"
         ButtonAlignment12=   2
         BeginProperty ButtonFont12 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState12   =   5
         ButtonLeft12    =   493
         ButtonTop12     =   2
         ButtonWidth12   =   24
         ButtonHeight12  =   24
         Begin DrawSuite2022.USImageList USImageList1 
            Left            =   10020
            Top             =   240
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmQualidade_RNC.frx":DD7E77
            Count           =   1
         End
      End
      Begin DrawSuite2022.USToolBar USToolBar2 
         Height          =   975
         Left            =   -74925
         TabIndex        =   111
         Top             =   330
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   1720
         ButtonCount     =   9
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft1     =   2
         ButtonTop1      =   2
         ButtonWidth1    =   44
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft2     =   48
         ButtonTop2      =   2
         ButtonWidth2    =   45
         ButtonHeight2   =   21
         ButtonUseMaskColor2=   0   'False
         ButtonCaption3  =   "Relatório"
         ButtonEnabled3  =   0   'False
         ButtonIconSize3 =   32
         ButtonToolTipText3=   "Relatório (F5)"
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
         ButtonLeft3     =   95
         ButtonTop3      =   2
         ButtonWidth3    =   60
         ButtonHeight3   =   21
         ButtonUseMaskColor3=   0   'False
         ButtonCaption4  =   "Anterior"
         ButtonEnabled4  =   0   'False
         ButtonIconSize4 =   32
         ButtonToolTipText4=   "Registro anterior."
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
         ButtonLeft4     =   157
         ButtonTop4      =   2
         ButtonWidth4    =   55
         ButtonHeight4   =   21
         ButtonUseMaskColor4=   0   'False
         ButtonCaption5  =   "Próximo"
         ButtonEnabled5  =   0   'False
         ButtonIconSize5 =   32
         ButtonToolTipText5=   "Próximo registro."
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
         ButtonLeft5     =   214
         ButtonTop5      =   2
         ButtonWidth5    =   55
         ButtonHeight5   =   21
         ButtonUseMaskColor5=   0   'False
         ButtonEnabled6  =   0   'False
         ButtonIconSize6 =   32
         ButtonAlignment6=   2
         ButtonType6     =   1
         ButtonStyle6    =   -1
         BeginProperty ButtonFont6 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState6    =   -1
         ButtonLeft6     =   271
         ButtonTop6      =   4
         ButtonWidth6    =   2
         ButtonHeight6   =   54
         ButtonCaption7  =   "Ajuda"
         ButtonEnabled7  =   0   'False
         ButtonIconSize7 =   32
         ButtonToolTipText7=   "Ajuda (F1)"
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
         ButtonLeft7     =   275
         ButtonTop7      =   2
         ButtonWidth7    =   41
         ButtonHeight7   =   21
         ButtonUseMaskColor7=   0   'False
         ButtonCaption8  =   "Sair"
         ButtonEnabled8  =   0   'False
         ButtonIconSize8 =   32
         ButtonToolTipText8=   "Sair (Esc)"
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
         ButtonLeft8     =   318
         ButtonTop8      =   2
         ButtonWidth8    =   30
         ButtonHeight8   =   21
         ButtonUseMaskColor8=   0   'False
         ButtonEnabled9  =   0   'False
         ButtonIconSize9 =   32
         ButtonKey9      =   "9"
         ButtonAlignment9=   2
         BeginProperty ButtonFont9 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState9    =   5
         ButtonLeft9     =   350
         ButtonTop9      =   2
         ButtonWidth9    =   24
         ButtonHeight9   =   24
         ButtonUseMaskColor9=   0   'False
      End
      Begin DrawSuite2022.USToolBar USToolBar3 
         Height          =   975
         Left            =   75
         TabIndex        =   112
         Top             =   330
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   1720
         ButtonCount     =   9
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft1     =   2
         ButtonTop1      =   2
         ButtonWidth1    =   44
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft2     =   48
         ButtonTop2      =   2
         ButtonWidth2    =   45
         ButtonHeight2   =   21
         ButtonUseMaskColor2=   0   'False
         ButtonCaption3  =   "Relatório"
         ButtonEnabled3  =   0   'False
         ButtonIconSize3 =   32
         ButtonToolTipText3=   "Relatório (F5)"
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
         ButtonLeft3     =   95
         ButtonTop3      =   2
         ButtonWidth3    =   60
         ButtonHeight3   =   21
         ButtonUseMaskColor3=   0   'False
         ButtonCaption4  =   "Anterior"
         ButtonEnabled4  =   0   'False
         ButtonIconSize4 =   32
         ButtonToolTipText4=   "Registro anterior."
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
         ButtonLeft4     =   157
         ButtonTop4      =   2
         ButtonWidth4    =   55
         ButtonHeight4   =   21
         ButtonUseMaskColor4=   0   'False
         ButtonCaption5  =   "Próximo"
         ButtonEnabled5  =   0   'False
         ButtonIconSize5 =   32
         ButtonToolTipText5=   "Próximo registro."
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
         ButtonLeft5     =   214
         ButtonTop5      =   2
         ButtonWidth5    =   55
         ButtonHeight5   =   21
         ButtonUseMaskColor5=   0   'False
         ButtonEnabled6  =   0   'False
         ButtonIconSize6 =   32
         ButtonAlignment6=   2
         ButtonType6     =   1
         ButtonStyle6    =   -1
         BeginProperty ButtonFont6 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState6    =   -1
         ButtonLeft6     =   271
         ButtonTop6      =   4
         ButtonWidth6    =   2
         ButtonHeight6   =   54
         ButtonCaption7  =   "Ajuda"
         ButtonEnabled7  =   0   'False
         ButtonIconSize7 =   32
         ButtonToolTipText7=   "Ajuda (F1)"
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
         ButtonLeft7     =   275
         ButtonTop7      =   2
         ButtonWidth7    =   41
         ButtonHeight7   =   21
         ButtonUseMaskColor7=   0   'False
         ButtonCaption8  =   "Sair"
         ButtonEnabled8  =   0   'False
         ButtonIconSize8 =   32
         ButtonToolTipText8=   "Sair (Esc)"
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
         ButtonLeft8     =   318
         ButtonTop8      =   2
         ButtonWidth8    =   30
         ButtonHeight8   =   21
         ButtonUseMaskColor8=   0   'False
         ButtonEnabled9  =   0   'False
         ButtonIconSize9 =   32
         ButtonKey9      =   "9"
         ButtonAlignment9=   2
         BeginProperty ButtonFont9 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState9    =   5
         ButtonLeft9     =   350
         ButtonTop9      =   2
         ButtonWidth9    =   24
         ButtonHeight9   =   24
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   -69585
         Top             =   3930
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin DrawSuite2022.USToolBar USToolBar4 
         Height          =   975
         Left            =   -74925
         TabIndex        =   117
         Top             =   330
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft2     =   40
         ButtonTop2      =   2
         ButtonWidth2    =   44
         ButtonHeight2   =   21
         ButtonUseMaskColor2=   0   'False
         ButtonCaption3  =   "Excluir"
         ButtonEnabled3  =   0   'False
         ButtonIconSize3 =   32
         ButtonToolTipText3=   "Excluir (F4)"
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
         ButtonLeft3     =   86
         ButtonTop3      =   2
         ButtonWidth3    =   45
         ButtonHeight3   =   21
         ButtonUseMaskColor3=   0   'False
         ButtonCaption4  =   "Relatório"
         ButtonEnabled4  =   0   'False
         ButtonIconSize4 =   32
         ButtonToolTipText4=   "Relatório (F5)"
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
         ButtonLeft4     =   133
         ButtonTop4      =   2
         ButtonWidth4    =   60
         ButtonHeight4   =   21
         ButtonCaption5  =   "Anterior"
         ButtonEnabled5  =   0   'False
         ButtonIconSize5 =   32
         ButtonToolTipText5=   "Registro anterior."
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
         ButtonLeft5     =   195
         ButtonTop5      =   2
         ButtonWidth5    =   55
         ButtonHeight5   =   21
         ButtonUseMaskColor5=   0   'False
         ButtonCaption6  =   "Próximo"
         ButtonEnabled6  =   0   'False
         ButtonIconSize6 =   32
         ButtonToolTipText6=   "Próximo registro."
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
         ButtonLeft6     =   252
         ButtonTop6      =   2
         ButtonWidth6    =   55
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
         ButtonLeft7     =   309
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft8     =   313
         ButtonTop8      =   2
         ButtonWidth8    =   41
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft9     =   356
         ButtonTop9      =   2
         ButtonWidth9    =   30
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
         ButtonLeft10    =   388
         ButtonTop10     =   2
         ButtonWidth10   =   24
         ButtonHeight10  =   24
         Begin DrawSuite2022.USImageList USImageList4 
            Left            =   13200
            Top             =   270
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmQualidade_RNC.frx":DDEAE7
            Count           =   1
         End
      End
      Begin MSComctlLib.ListView Lista_doc 
         Height          =   4635
         Left            =   -74925
         TabIndex        =   61
         Top             =   3750
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   8176
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "N"
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Tag             =   "T"
            Text            =   "Caminho"
            Object.Width           =   25576
         EndProperty
      End
   End
End
Attribute VB_Name = "frmQualidade_RNC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Novo_RNC                   As Boolean 'OK
Dim Novo_RNC1                  As Boolean 'OK
Public Responsavel_RNC         As Boolean 'OK
Public StrSql_CQ_RNC_Localizar As String 'OK
Public FormulaRel_CQ_RNC       As String 'OK
Dim TBLISTA_CQ_RNC As ADODB.Recordset 'OK
Public RespRNC As String

Private Sub Chk_acao_corretiva_Click()
On Error GoTo tratar_erro

With SSTab1
    If Chk_acao_corretiva.Value = 1 Then
        .TabVisible(1) = True
        .TabVisible(2) = True
        .TabsPerRow = 4
    Else
        .TabVisible(1) = False
        .TabVisible(2) = False
        .TabsPerRow = 2
    End If
End With
        
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkRevisao_Click()
On Error GoTo tratar_erro

If chkRevisao.Value = 1 Then
    txtrevisao.Enabled = True
    txtrevisao.SetFocus
Else
    txtrevisao.Enabled = False
    txtrevisao.Text = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub cmbRequisitos_Click()
On Error GoTo tratar_erro

If cmbRequisitos.Text = "Sim" Then
    txtQuaisRequisitos.Enabled = True
Else
    txtQuaisRequisitos.Enabled = False
    txtQuaisRequisitos.Text = ""
End If
    

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub cmbSimilaridade_Click()
On Error GoTo tratar_erro

If cmbSimilaridade.Text = "Sim" Then
    txtQuaisRNC.Enabled = True
Else
    txtQuaisRNC.Enabled = False
    txtQuaisRNC.Text = ""
End If
    

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_limpar_caminho_Click()
On Error GoTo tratar_erro

txt_Caminho = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_limpar_caminho1_Click()
On Error GoTo tratar_erro

Txt_imagemRNC = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_visualizar_arquivo1_Click()
On Error GoTo tratar_erro

If Txt_imagemRNC <> "" Then ProcAbrirArquivo Txt_imagemRNC

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_visualizar_arquivo_Click()
On Error GoTo tratar_erro

If txt_Caminho <> "" Then ProcAbrirArquivo txt_Caminho

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdAuditorLider_Click()
On Error GoTo tratar_erro

RespRNC = "auditor"
frmQualidade_RNC_Usuarios.Show 1


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub

End Sub

Private Sub cmdClassificacao_Click()
On Error GoTo tratar_erro

RespRNC = "qualidade"
frmQualidade_RNC_Classificacao.Show 1


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdImagemRNC_Click()
On Error GoTo tratar_erro

ProcCarregaCaminhoNomeArquivo CommonDialog1, "*.*", "*.*"
Txt_imagemRNC = caminho

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdImportar_Click()
On Error GoTo tratar_erro

ProcCarregaCaminhoNomeArquivo CommonDialog1, "*.*", "*.*"
txt_Caminho = caminho

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_CQ_RNC.AbsolutePage <> 2 Then
    If TBLISTA_CQ_RNC.AbsolutePage = -3 Then
        ProcExibePagina (TBLISTA_CQ_RNC.PageCount - 1)
    Else
        TBLISTA_CQ_RNC.AbsolutePage = TBLISTA_CQ_RNC.AbsolutePage - 2
        ProcExibePagina (TBLISTA_CQ_RNC.AbsolutePage)
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
    TBLISTA_CQ_RNC.AbsolutePage = txtPagIr.Text
    ProcExibePagina (TBLISTA_CQ_RNC.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_CQ_RNC.AbsolutePage = 1
ProcExibePagina (TBLISTA_CQ_RNC.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_CQ_RNC.AbsolutePage <> -3 Then
    If TBLISTA_CQ_RNC.AbsolutePage = 1 Then
        ProcExibePagina (2)
    Else
        ProcExibePagina (TBLISTA_CQ_RNC.AbsolutePage)
    End If
Else
    ProcExibePagina (TBLISTA_CQ_RNC.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_CQ_RNC.AbsolutePage = TBLISTA_CQ_RNC.PageCount
ProcExibePagina (TBLISTA_CQ_RNC.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub



Private Sub cmdResponsavelQualidade_Click()
On Error GoTo tratar_erro

RespRNC = "qualidade"
frmQualidade_RNC_Usuarios.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub

End Sub

Private Sub cmdSalvarRespRNC_Click()
On Error GoTo tratar_erro
  
Set TBUsuarios = CreateObject("adodb.recordset")
TBUsuarios.Open "Select * from CQ_RNC where id = '" & txtId & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBUsuarios.EOF = False Then
    TBUsuarios!txtAuditor = txtAuditor
    TBUsuarios!txtRespQualidade = txtRespQualidade
    TBUsuarios.Update
End If
TBUsuarios.Close

MsgBox ("Os responsáveis da RNC foram incluídos com sucesso!"), vbInformation + vbOKOnly


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub

End Sub

Private Sub cmdValorUnitario_Click()
On Error GoTo tratar_erro

    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select * from ProjProduto where desenho = '" & txtdesenho & "'", Conexao, adOpenKeyset, adLockOptimistic
    txtvalorunitario.Text = ""
    If TBProduto.EOF = False And cmbMateriaPrima = "" Then txtvalorunitario = TBProduto!PConsumo
    If TBProduto.EOF = False And cmbMateriaPrima <> "" Then txtvalorunitario = TBProduto!PCusto
    TBProduto.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub

End Sub

Private Sub Imgcalendario2_Click()
On Error GoTo tratar_erro

Sit_Data = 2
ProcAbrirCalendario

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Imgcalendario3_Click()
On Error GoTo tratar_erro

Sit_Data = 3
ProcAbrirCalendario

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Imgcalendario4_Click()
On Error GoTo tratar_erro

Sit_Data = 4
ProcAbrirCalendario

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub




Private Sub Lista_doc_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Lista_doc
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then .ListItems.Item(InitFor).Checked = False Else .ListItems.Item(InitFor).Checked = True
        Next InitFor
    End With
Else
    ProcOrdenaListView Lista_doc, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_doc_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista_doc.ListItems.Count = 0 Then Exit Sub
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from CQ_RNC_documentos where id = " & Lista_doc.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    Proclimpacampos_doc
    ProcPuxadados_Doc
    CodigoLista1 = Lista_doc.SelectedItem.index
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub




Private Sub optNao_Click()
On Error GoTo tratar_erro

With Label23
    If optNao.Value = True Then
        .Caption = "Justificativa"
        .ToolTipText = "Justificativa"
    Else
        .Caption = "Ações corretivas/preventivas"
        .ToolTipText = "Ações corretivas/preventivas"
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optSim_Click()
On Error GoTo tratar_erro

With Label23
    If optSim.Value = True Then
        .ToolTipText = "Ações corretivas/preventivas"
        .Caption = "Ações corretivas/preventivas"
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_localizar_cliente_fornecedor_Click()
On Error GoTo tratar_erro

frmQualidade_RNC_cliente_forn.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_localizar_item_Click()
On Error GoTo tratar_erro

frmQualidade_RNC_item.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAnterior()
On Error GoTo tratar_erro

If txtId = 0 Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from CQ_RNC order by id", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.BOF = False Then
    TBLISTA.Find ("id = " & txtId)
    TBLISTA.MovePrevious
    If TBLISTA.BOF = False Then
        ProcLimpaCampos
        ProcLimpaCampos2
        ProcLimpaCampos3
        txtId = TBLISTA!ID
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from CQ_RNC where id  = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            ProcPuxaDados
        End If
        ProcPuxadados2
        ProcPuxadados3
        ProcCarregaLista_Doc
    Else
        USMsgBox ("Fim dos cadastros de RNC."), vbInformation, "CAPRIND v5.0"
    End If
End If
Novo_RNC = False
Novo_RNC1 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procExcluir3()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso.")
    Exit Sub
End If
If USMsgBox("Deseja realmente excluir as ações do fornecedor?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    Set TBGravar = CreateObject("adodb.recordset")
    TBGravar.Open "Select * from CQ_RNC where id  = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
    If TBGravar.EOF = False Then
        TBGravar!Texto3 = Null
        TBGravar!texto4 = Null
        TBGravar!Texto5 = Null
        TBGravar!Texto6 = Null
        TBGravar!Texto7 = Null
        TBGravar!Texto8 = Null
        TBGravar!Data3 = Null
        TBGravar!Data4 = Null
        TBGravar!Data5 = Null
        TBGravar!responsavel3 = Null
        TBGravar!responsavel4 = Null
        TBGravar!Acompanhamento = Null
        TBGravar!Eficacia = Null
        TBGravar.Update
    End If
    USMsgBox ("Ações do fornecedor excluídas com sucesso."), vbInformation, "CAPRIND v5.0"
    '==================================
    Modulo = "Qualidade/RNC"
    Evento = "Excluir ações do fornecedor"
    ID_documento = txtId
    Documento = IIf(Txt_sequencial = "__", "Nº RNC: " & txtID_texto, "Nº RNC: " & txtID_texto & " - Sequencial: " & Txt_sequencial)
    Documento1 = ""
    ProcGravaEvento
    '==================================
    ProcLimpaCampos3
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procExcluir_doc()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With Lista_doc
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir esta(s) documento(s)?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            Conexao.Execute "DELETE from CQ_RNC_documentos where ID  = " & .ListItems(InitFor)
            '==================================
            Modulo = "Qualidade/RNC"
            Evento = "Excluir documento"
            ID_documento = .ListItems(InitFor)
            Documento = IIf(Txt_sequencial = "__", "Nº RNC: " & txtID_texto, "Nº RNC: " & txtID_texto & " - Sequencial: " & Txt_sequencial)
            Documento1 = "Caminho: " & .ListItems(InitFor).ListSubItems(1)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) documento(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Documento(s) excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    Proclimpacampos_doc
    ProcCarregaLista_Doc
    Novo_RNC1 = False
    Frame14.Enabled = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcProximo()
On Error GoTo tratar_erro

If txtId = 0 Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from CQ_RNC order by id", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.BOF = False Then
    TBLISTA.Find ("id = " & txtId)
    TBLISTA.MoveNext
    If TBLISTA.EOF = False Then
        ProcLimpaCampos
        ProcLimpaCampos2
        ProcLimpaCampos3
        txtId = TBLISTA!ID
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from CQ_RNC where id = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            ProcPuxaDados
        End If
        ProcPuxadados2
        ProcPuxadados3
        ProcCarregaLista_Doc
    Else
        USMsgBox ("Fim dos cadastros de RNC."), vbInformation, "CAPRIND v5.0"
    End If
End If
Novo_RNC = False
Novo_RNC1 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procSalvar3()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If optNao.Value = True And txtTexto7 = "" Then
    USMsgBox ("Informe a justificativa antes de salvar."), vbExclamation, "CAPRIND v5.0"
    txtTexto7.SetFocus
    Exit Sub
End If
If txtdata7 <> "__/__/____" Then
    If IsDate(txtdata7) = False Then
        NomeCampo = "a data de verificação da implementação"
        ProcVerificaAcao
        txtdata7.SetFocus
        Exit Sub
    End If
End If
If txtData5 <> "__/__/____" Then
    If IsDate(txtData5) = False Then
        NomeCampo = "a data de verificação da eficácia"
        ProcVerificaAcao
        txtData5.SetFocus
        Exit Sub
    End If
End If
If txtData6 <> "__/__/____" Then
    If IsDate(txtData6) = False Then
        NomeCampo = "a data da abrangência"
        ProcVerificaAcao
        txtData6.SetFocus
        Exit Sub
    End If
End If
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from CQ_RNC WHERE ID = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then TBGravar.AddNew
ProcEnviaDados3
TBGravar.Update
TBGravar.Close
USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
'==================================
Modulo = "Qualidade/RNC"
Evento = "Alterar ações do fornecedor"
ID_documento = txtId
Documento = IIf(Txt_sequencial = "__", "Nº RNC: " & txtID_texto, "Nº RNC: " & txtID_texto & " - Sequencial: " & Txt_sequencial)
Documento1 = ""
ProcGravaEvento
'==================================

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procSalvar_doc()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Frame14.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
Acao = "salvar"
If txt_Caminho = "" Then
    NomeCampo = "o caminho"
    ProcVerificaAcao
    cmdImportar.SetFocus
    Exit Sub
End If
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from CQ_RNC_documentos where ID = " & txtID_doc, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then TBGravar.AddNew
ProcEnviadados_doc
TBGravar.Update
txtID_doc = TBGravar!ID
TBGravar.Close
ProcCarregaLista_Doc
If Novo_RNC1 = True Then
    USMsgBox ("Novo documento cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo documento"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar documento"
    If CodigoLista1 <> 0 And Lista_doc.ListItems.Count <> 0 Then
        Lista_doc.SelectedItem = Lista_doc.ListItems(CodigoLista1)
        Lista_doc.SetFocus
    End If
End If
'==================================
Modulo = "Qualidade/RNC"
ID_documento = txtId
Documento = IIf(Txt_sequencial = "__", "Nº RNC: " & txtID_texto, "Nº RNC: " & txtID_texto & " - Sequencial: " & Txt_sequencial)
Documento1 = "Caminho: " & txt_Caminho
ProcGravaEvento
'==================================
Novo_RNC1 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro
  
Select Case SSTab1.Tab
    Case 0:
        Select Case KeyCode
            Case vbKeyInsert: ProcNovo
            Case vbKeyF2: ProcAbrir
            Case vbKeyF3: ProcSalvar
            Case vbKeyF4: ProcExcluir
            Case vbKeyF5: ProcImprimir
            Case vbKeyEscape: ProcSair
        End Select
    Case 1:
        Select Case KeyCode
            Case vbKeyF3: procSalvar2
            Case vbKeyF4: procExcluir2
            Case vbKeyF5: ProcImprimir
            Case vbKeyEscape: ProcSair
        End Select
    Case 2:
        Select Case KeyCode
            Case vbKeyF3: procSalvar3
            Case vbKeyF4: procExcluir3
            Case vbKeyF5: ProcImprimir
            Case vbKeyEscape: ProcSair
        End Select
    Case 3:
        Select Case KeyCode
            Case vbKeyInsert: procNovo_doc
            Case vbKeyF3: procSalvar_doc
            Case vbKeyF4: procExcluir_doc
            Case vbKeyF5: ProcImprimir
            Case vbKeyEscape: ProcSair
        End Select
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcPuxaDados()
On Error GoTo tratar_erro

Caption = "Qualidade - RNC (RNC : " & IIf(IsNull(TBAbrir!id_texto), "", TBAbrir!id_texto) & ")"
txtId = TBAbrir!ID
txtID_texto = IIf(IsNull(TBAbrir!id_texto), "", TBAbrir!id_texto)
If IsNull(TBAbrir!Seq) = False Then If TBAbrir!Seq < 10 Then Txt_sequencial = "0" & TBAbrir!Seq Else Txt_sequencial = TBAbrir!Seq

txtData = IIf(IsNull(TBAbrir!Data), Date, TBAbrir!Data)
txtResponsavel.Text = IIf(IsNull(TBAbrir!Responsavel), "", TBAbrir!Responsavel)
If IsNull(TBAbrir!status) = False And TBAbrir!status <> "" Then cmbStatus = TBAbrir!status
If IsNull(TBAbrir!Origem) = False And TBAbrir!Origem <> "" Then
    If TBAbrir!Origem = "I" Then Opt_interno.Value = True Else Opt_externo.Value = True
End If
If TBAbrir!Acao_corretiva = True Then Chk_acao_corretiva.Value = 1 Else Chk_acao_corretiva.Value = 0
Txt_imagemRNC = IIf(IsNull(TBAbrir!imagem), "", TBAbrir!imagem)
txt_QtdeAprovada.Text = IIf(IsNull(TBAbrir!QtdeAprovada), 0, TBAbrir!QtdeAprovada)
txt_qtdeLote.Text = IIf(IsNull(TBAbrir!QtdeLote), 0, TBAbrir!QtdeLote)
1:
    If TBAbrir!Preventiva = True Then optPreventiva.Value = True Else optCorretiva.Value = True
    txtdesenho = IIf(IsNull(TBAbrir!Desenho), "", TBAbrir!Desenho)
    txtdescricao = IIf(IsNull(TBAbrir!Descricao), "", TBAbrir!Descricao)
    txtQtde = IIf(IsNull(TBAbrir!Qtde), "", Format(TBAbrir!Qtde, "###,##0.0000"))
    If IsNull(TBAbrir!Tipo) = False And TBAbrir!Tipo <> "" Then
        txttipo = TBAbrir!Tipo
        If txttipo = "F" Then
            Set TBFornecedor = CreateObject("adodb.recordset")
            TBFornecedor.Open "Select * from Compras_fornecedores where idcliente = " & TBAbrir!ID_forn, Conexao, adOpenKeyset, adLockOptimistic
            If TBFornecedor.EOF = False Then
                txtID_forn = IIf(IsNull(TBFornecedor!IDCliente), "", TBFornecedor!IDCliente)
                txtFornecedor = IIf(IsNull(TBFornecedor!Nome_Razao), "", TBFornecedor!Nome_Razao)
            End If
            TBFornecedor.Close
        Else
            Set TBFornecedor = CreateObject("adodb.recordset")
            TBFornecedor.Open "Select * from clientes where idcliente = " & TBAbrir!ID_forn, Conexao, adOpenKeyset, adLockOptimistic
            If TBFornecedor.EOF = False Then
                txtID_forn = IIf(IsNull(TBFornecedor!IDCliente), "", TBFornecedor!IDCliente)
                txtFornecedor = IIf(IsNull(TBFornecedor!NomeRazao), "", TBFornecedor!NomeRazao)
            End If
            TBFornecedor.Close
        End If
    Else
        txtID_forn = IIf(IsNull(TBAbrir!ID_forn), "", TBAbrir!ID_forn)
        txtFornecedor = IIf(IsNull(TBAbrir!Cliente_forn), "", TBAbrir!Cliente_forn)
    End If
    txtNao_conformidade = IIf(IsNull(TBAbrir!nao_conformidade), "", TBAbrir!nao_conformidade)
    If IsNull(TBAbrir!eficaz) = False And TBAbrir!eficaz <> "" Then cmbEficaz = TBAbrir!eficaz
    txtfim = IIf(IsNull(TBAbrir!FIM), "__/__/____", Format(TBAbrir!FIM, "dd/mm/yyyy"))
    Txt_doc_referencia = IIf(IsNull(TBAbrir!Documento_ref), "", TBAbrir!Documento_ref)
    txtEquipe = IIf(IsNull(TBAbrir!Equipe), "", TBAbrir!Equipe)
    txtClassificacao = IIf(IsNull(TBAbrir!classificacao), "", TBAbrir!classificacao)
    If IsNull(TBAbrir!Procede) = False And TBAbrir!Procede <> "" Then cmbProcede = TBAbrir!Procede
    If IsNull(TBAbrir!Requisitos) = False And TBAbrir!Requisitos <> "" Then cmbRequisitos = TBAbrir!Requisitos
    If IsNull(TBAbrir!QuaisRequisitos) = False And TBAbrir!QuaisRequisitos <> "" Then txtQuaisRequisitos = TBAbrir!QuaisRequisitos
    If IsNull(TBAbrir!Similaridade) = False And TBAbrir!Similaridade <> "" Then cmbSimilaridade = TBAbrir!Similaridade
    If IsNull(TBAbrir!QuaisRNC) = False And TBAbrir!QuaisRNC <> "" Then txtQuaisRNC = TBAbrir!QuaisRNC
    If IsNull(TBAbrir!MateriaPrima) = False And TBAbrir!MateriaPrima <> "" Then cmbMateriaPrima = TBAbrir!MateriaPrima
    If IsNull(TBAbrir!txtvalorunitario) = False And TBAbrir!txtvalorunitario <> "" Then txtvalorunitario = TBAbrir!txtvalorunitario
    If IsNull(TBAbrir!txtFatorNC) = False And TBAbrir!txtFatorNC <> "" Then txtFatorNC = TBAbrir!txtFatorNC
    If IsNull(TBAbrir!txtCustoNC) = False And TBAbrir!txtCustoNC <> "" Then txtCustoNC = TBAbrir!txtCustoNC
    If IsNull(TBAbrir!txtAuditor) = False And TBAbrir!txtAuditor <> "" Then txtAuditor = TBAbrir!txtAuditor
    If IsNull(TBAbrir!txtRespQualidade) = False And TBAbrir!txtRespQualidade <> "" Then txtRespQualidade = TBAbrir!txtRespQualidade
    
    
    Novo_RNC = False
    Frame1.Enabled = True
    
    'Verifica se a RNC está amarrada a outro módulo e bloqueia os botões necessários
    Permitido = True
    Set TBFI = CreateObject("adodb.recordset")
    TBFI.Open "Select * from compras_recebimento where ID_RNC = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
    If TBFI.EOF = False Then Permitido = False
    Set TBFI = CreateObject("adodb.recordset")
    TBFI.Open "Select * from Medicao where ID_RNC = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
    If TBFI.EOF = False Then Permitido = False
    Set TBFI = CreateObject("adodb.recordset")
    TBFI.Open "Select * from CQ_NC_FABRICA where ID_RNC = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
    If TBFI.EOF = False Then Permitido = False
    Set TBFI = CreateObject("adodb.recordset")
    TBFI.Open "Select * from CQ_SD where ID_RNC = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
    If TBFI.EOF = False Then Permitido = False
    TBFI.Close
    If Permitido = True Then
        With txtdesenho
            .Locked = False
            .TabStop = True
        End With
        With txtdescricao
            .Locked = False
            .TabStop = True
        End With
        With txt_qtdeLote
            .Locked = False
            .TabStop = True
        End With
        With txtQtde
            .Locked = False
            .TabStop = True
        End With
        Cmd_localizar_item.Enabled = True
        Cmd_localizar_cliente_fornecedor.Enabled = True
    Else
        With txtdesenho
            .Locked = True
            .TabStop = False
        End With
        With txtdescricao
            .Locked = True
            .TabStop = False
        End With
        With txt_qtdeLote
            .Locked = True
            .TabStop = False
        End With
        With txtQtde
            .Locked = True
            .TabStop = False
        End With
        Cmd_localizar_item.Enabled = False
        Cmd_localizar_cliente_fornecedor.Enabled = False
    End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCampos2()
On Error GoTo tratar_erro

txtData2.Value = Date
txtResponsavel2 = ""
txtTexto1 = ""
txtTexto2 = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCampos3()
On Error GoTo tratar_erro

txtData3.Value = Date
txtResponsavel3 = ""
txtTexto3 = ""
txtTexto4 = ""
txtTexto5 = ""
txtTexto6 = ""
txtTexto7 = ""
txtData4.Value = Date
txtResponsavel4 = ""
optSim.Value = True
optNao.Value = False
txtAcompanhamento = ""
txtData5 = "__/__/____"
txtEficacia = ""
txtData6 = "__/__/____"
txtResponsavel5 = ""
txtTexto8 = ""
txtdata7 = "__/__/____"
txtResponsavel7 = ""
txtResponsavel6 = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub Proclimpacampos_doc()
On Error GoTo tratar_erro

txtID_doc = 0
txtData_doc = Format(Date, "dd/mm/yy")
txtResponsavel_doc = pubUsuario
txt_Caminho = ""
Txt_obs_doc = ""
CodigoLista1 = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCampos()
On Error GoTo tratar_erro

txtId = 0
txtID_texto = "____-__"
Txt_sequencial = "__"
Txt_imagemRNC.Text = ""
txt_QtdeAprovada.Text = ""
txt_qtdeLote.Text = ""
txtData.Value = Date
txtResponsavel = pubUsuario
cmbStatus = "Ativa"
optCorretiva.Value = True
optPreventiva.Value = False
Opt_interno.Value = True
Opt_externo.Value = False
Chk_acao_corretiva.Value = 1
txtClassificacao = ""
txtfim = "__/__/____"
cmbEficaz.ListIndex = -1
Txt_doc_referencia = ""
txtNao_conformidade = ""
txtEquipe = ""
CodigoLista = 0
Caption = "Qualidade - RNC"
cmbProcede.ListIndex = -1
cmbRequisitos.ListIndex = -1
txtQuaisRequisitos.Text = ""
cmbSimilaridade.ListIndex = -1
txtQuaisRNC.Text = ""
cmbMateriaPrima.ListIndex = -1
txtvalorunitario = ""
txtFatorNC = "1"
txtCustoNC = ""
txtAuditor = ""
txtRespQualidade = ""

If RNC_Inspecao_Recebimento = False And RNC_Controle_Medicao = False And RNC_Nao_Conformidade = False And RNC_Solicitacao_Desvio = False Then
    With txtdesenho
        .Text = ""
        .Locked = False
        .TabStop = True
    End With
    With txtdescricao
        .Text = ""
        .Locked = False
        .TabStop = True
    End With
    With txt_qtdeLote
        .Text = ""
        .Locked = False
        .TabStop = True
    End With
    With txtQtde
        .Text = ""
        .Locked = False
        .TabStop = True
    End With
    txtID_forn = 0
    With txtFornecedor
        .Text = ""
        .Locked = False
        .TabStop = True
    End With
Else
    With txtdesenho
        .Locked = True
        .TabStop = False
    End With
    With txtdescricao
        .Locked = True
        .TabStop = False
    End With
    With txt_qtdeLote
        .Locked = True
        .TabStop = False
    End With
    With txtQtde
        .Locked = True
        .TabStop = False
    End With
    With txtFornecedor
        .Locked = True
        .TabStop = False
    End With
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviaDados2()
On Error GoTo tratar_erro

TBGravar!Data2 = txtData2
TBGravar!responsavel2 = txtResponsavel2
TBGravar!Texto1 = Trim(txtTexto1)
TBGravar!Texto2 = Trim(txtTexto2)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviaDados3()
On Error GoTo tratar_erro

TBGravar!Data3 = txtData3
TBGravar!responsavel3 = txtResponsavel3
TBGravar!Texto3 = Trim(txtTexto3)
TBGravar!texto4 = Trim(txtTexto4)
TBGravar!Texto5 = Trim(txtTexto5)
TBGravar!Texto6 = Trim(txtTexto6)
TBGravar!Texto7 = Trim(txtTexto7)
TBGravar!Data4 = txtData4
TBGravar!responsavel4 = txtResponsavel4
If optSim.Value = True Then TBGravar!Acao = True Else TBGravar!Acao = False
TBGravar!Acompanhamento = Trim(txtAcompanhamento)
TBGravar!Data5 = IIf(txtData5 = "__/__/____", Null, txtData5)
TBGravar!responsavel6 = txtResponsavel6
TBGravar!Eficacia = Trim(txtEficacia)
TBGravar!Data6 = IIf(txtData6 = "__/__/____", Null, txtData6)
TBGravar!responsavel5 = txtResponsavel5
TBGravar!Texto8 = Trim(txtTexto8)
TBGravar!Data7 = IIf(txtdata7 = "__/__/____", Null, txtdata7)
TBGravar!responsavel7 = txtResponsavel7
TBGravar!chkMaoObra = chkMaoObra.Value
TBGravar!chkMaquina = chkMaquina.Value
TBGravar!chkMetodo = chkMetodo.Value
TBGravar!chkMedidas = chkMedidas.Value
TBGravar!chkMaterial = chkMaterial.Value
TBGravar!chkMeioAmbiente = chkMeioAmbiente.Value
TBGravar!chkSGQ = chkSGQ.Value
TBGravar!chkOutros = chkOutros.Value
TBGravar!chkRevisao = chkRevisao.Value
TBGravar!txtrevisao = txtrevisao

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviadados_doc()
On Error GoTo tratar_erro

TBGravar!ID_RNC = txtId
TBGravar!Data = txtData_doc
TBGravar!Responsavel = txtResponsavel_doc
TBGravar!Caminho_documento = txt_Caminho
TBGravar!Observacao = Txt_obs_doc

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procSalvar2()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from CQ_RNC WHERE ID = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then TBGravar.AddNew
ProcEnviaDados2
TBGravar.Update
TBGravar.Close
USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
'==================================
Modulo = "Qualidade/RNC"
Evento = "Alterar ações do fornecedor"
ID_documento = txtId
Documento = IIf(Txt_sequencial = "__", "Nº RNC: " & txtID_texto, "Nº RNC: " & txtID_texto & " - Sequencial: " & Txt_sequencial)
Documento1 = ""
ProcGravaEvento
'==================================

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procExcluir2()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso.")
    Exit Sub
End If
If USMsgBox("Deseja realmente excluir as ações do fornecedor?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    Set TBGravar = CreateObject("adodb.recordset")
    TBGravar.Open "Select * from CQ_RNC where id  = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
    If TBGravar.EOF = False Then
        TBGravar!Texto1 = Null
        TBGravar!Texto2 = Null
        TBGravar!Data2 = Null
        TBGravar!responsavel2 = Null
        TBGravar.Update
    End If
    USMsgBox ("Ações do fornecedor excluídas com sucesso."), vbInformation, "CAPRIND v5.0"
    '==================================
    Modulo = "Qualidade/RNC"
    Evento = "Excluir ações do fornecedor"
    ID_documento = txtId
    Documento = IIf(Txt_sequencial = "__", "Nº RNC: " & txtID_texto, "Nº RNC: " & txtID_texto & " - Sequencial: " & Txt_sequencial)
    Documento1 = ""
    ProcGravaEvento
    '==================================
    ProcLimpaCampos2
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15195, 12, True
ProcCarregaToolBar2 Me, 15195, 9, True
ProcCarregaToolBar3 Me, 15195, 9, True
ProcCarregaToolBar4 Me, 15195, 10, True

Formulario = "Qualidade/RNC"
Direitos
SSTab1.Tab = 0
SSTab2.Tab = 0
ProcLimpaVariaveisPrincipais
txtData = Date
txtData2 = Date
txtData3 = Date
txtData4 = Date
txtData_doc = Date

StrSql_CQ_RNC_Localizar = ""

ProcRemoveObjetosResize Me

If RNC_Inspecao_Recebimento = False And RNC_Controle_Medicao = False And RNC_Nao_Conformidade = False And RNC_Solicitacao_Desvio = False Then
    With USToolBar1
        .ButtonState(1) = 0
        .Refresh
    End With
    Exit Sub
Else
    With USToolBar1
        .ButtonState(1) = 5
        .Refresh
    End With
    'Carrega dados da RNC vinculada com inspecao de recebimento, controle de medicao, nao conformidade e solicitação de desvio
    Set TBAbrir = CreateObject("adodb.recordset")
    If RNC_Inspecao_Recebimento = True Then TextoFiltro = IIf(frmCompras_recebimento.Txt_ID_RNC = "", 0, frmCompras_recebimento.Txt_ID_RNC)
    If RNC_Controle_Medicao = True Then TextoFiltro = IIf(frmPlanomedicao.Txt_ID_RNC = "", 0, frmPlanomedicao.Txt_ID_RNC)
    If RNC_Nao_Conformidade = True Then
        If Sit_REG = 1 Then TextoFiltro = IIf(frmcqnc.Txt_ID_RNC = "", 0, frmcqnc.Txt_ID_RNC) Else TextoFiltro = 0
    End If
    If RNC_Solicitacao_Desvio = True Then TextoFiltro = IIf(frmCQ_SD.Txt_ID_RNC = "", 0, frmCQ_SD.Txt_ID_RNC)
    TBAbrir.Open "Select * from CQ_RNC where ID = " & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        ProcLimpaCampos
        ProcPuxaDados
        StrSql_CQ_RNC_Localizar = "Select * from CQ_RNC where ID = " & txtId
        ProcCarregaLista (1)
    Else
        cmbStatus = "Ativa"
        
        If RNC_Inspecao_Recebimento = True Then
            With frmCompras_recebimento
                txtdesenho = .txtNomenclatura
                txtdescricao = .txtEspecificacoes
                txtQtde = .Txtrejeitado
                txt_qtdeLote = .Txt_qtde_recebida
                txtFornecedor = .Txt_cliente_forn
                
                Set TBTempo = CreateObject("adodb.recordset")
                TBTempo.Open "Select IDestoque from Qualidade_inspecao_recebimento where IDestoque = " & .ListProdReceb.SelectedItem.ListSubItems(5) & " and Consignacao = 'True'", Conexao, adOpenKeyset, adLockOptimistic
                If TBTempo.EOF = False Then
                    NomeTabela = "Clientes"
                    NomeCampo = "NomeRazao"
                    txttipo = "C"
                Else
                    NomeTabela = "Compras_fornecedores"
                    NomeCampo = "Nome_Razao"
                    txttipo = "F"
                End If
                TBTempo.Close
                                    
                Set TBFornecedor = CreateObject("adodb.recordset")
                TBFornecedor.Open "Select IDcliente from " & NomeTabela & " where " & NomeCampo & " = '" & txtFornecedor & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBFornecedor.EOF = False Then
                    txtID_forn = TBFornecedor!IDCliente
                End If
                TBFornecedor.Close
            End With
        End If
        If RNC_Controle_Medicao = True Then
            With frmPlanomedicao
                txtdesenho = .txtdesenho
                txtdescricao = .txtdescricao
                txtQtde = .txtQuant_liber
                txt_qtdeLote.Text = .Txt_qtde_lote.Text
                Set TBproducao = CreateObject("adodb.recordset")
                TBproducao.Open "Select Cliente from producao where Ordem = " & .Txtpeca, Conexao, adOpenKeyset, adLockOptimistic
                If TBproducao.EOF = False Then
                    txtFornecedor = Trim(TBproducao!Cliente)
                End If
                TBproducao.Close
                Set TBFornecedor = CreateObject("adodb.recordset")
                TBFornecedor.Open "Select IDCliente from Clientes where NomeRazao = '" & txtFornecedor & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBFornecedor.EOF = False Then
                    txtID_forn = TBFornecedor!IDCliente
                End If
                TBFornecedor.Close
                txttipo = "C"
            End With
        End If
        If RNC_Nao_Conformidade = True Then
            If Sit_REG = 1 Then
                With frmcqnc
                    txtdesenho = .txtdesenho
                    txtdescricao = .txtdescricao
                    txtQtde = .txtnc
                    txt_qtdeLote.Text = .txtLote.Text
                    Set TBproducao = CreateObject("adodb.recordset")
                    TBproducao.Open "Select Cliente from producao where Ordem = " & .txtof, Conexao, adOpenKeyset, adLockOptimistic
                    If TBproducao.EOF = False Then
                        txtFornecedor = IIf(IsNull(TBproducao!Cliente), "", TBproducao!Cliente)
                    End If
                    TBproducao.Close
                    If txtFornecedor <> "" Then
                        Set TBFornecedor = CreateObject("adodb.recordset")
                        TBFornecedor.Open "Select IDCliente from Clientes where NomeRazao = '" & txtFornecedor & "'", Conexao, adOpenKeyset, adLockOptimistic
                        If TBFornecedor.EOF = False Then
                            txtID_forn = TBFornecedor!IDCliente
                        End If
                        TBFornecedor.Close
                        txttipo = "C"
                    End If
                End With
            Else
                Set TBFI = CreateObject("adodb.recordset")
                TBFI.Open "Select CQNCF.PARECERCQ, CQNCF.Lote, CQNCF.Analizada, CQNCF.Data, CQNCF.HOra, CQNCF.OS, CQNCF.Ordem, CQNCF.Operador, P.Desenho, P.Produto from CQ_NC_FABRICA CQNCF INNER JOIN Producao P ON CQNCF.Ordem = P.Ordem where CQNCF.Codigo = " & quantidade, Conexao, adOpenKeyset, adLockOptimistic
                If TBFI.EOF = False Then
                    txtdesenho = IIf(IsNull(TBFI!Desenho), "", TBFI!Desenho)
                    txtdescricao = IIf(IsNull(TBFI!Produto), "", TBFI!Produto)
                    txtQtde = Qtde 'qtde não conforme
                    txt_qtdeLote.Text = IIf(IsNull(TBFI!LOTE), "", TBFI!LOTE)
                    Set TBproducao = CreateObject("adodb.recordset")
                    TBproducao.Open "Select Cliente from producao where Ordem = " & TBFI!Ordem, Conexao, adOpenKeyset, adLockOptimistic
                    If TBproducao.EOF = False Then
                        txtFornecedor = IIf(IsNull(TBproducao!Cliente), "", TBproducao!Cliente)
                    End If
                    TBproducao.Close
                    If txtFornecedor <> "" Then
                        Set TBFornecedor = CreateObject("adodb.recordset")
                        TBFornecedor.Open "Select IDCliente from Clientes where NomeRazao = '" & txtFornecedor & "'", Conexao, adOpenKeyset, adLockOptimistic
                        If TBFornecedor.EOF = False Then
                            txtID_forn = TBFornecedor!IDCliente
                        End If
                        TBFornecedor.Close
                        txttipo = "C"
                    End If
                End If
                TBFI.Close
            End If
        End If
        If RNC_Solicitacao_Desvio = True Then
            With frmCQ_SD
                txtdesenho = .txtdesenho
                txtdescricao = .txtdescricao
                txtFornecedor = .txtCliente
                txt_qtdeLote.Text = .txtQtde.Text
                Set TBFornecedor = CreateObject("adodb.recordset")
                TBFornecedor.Open "Select IDCliente from Clientes where NomeRazao = '" & txtFornecedor & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBFornecedor.EOF = False Then
                    txtID_forn = TBFornecedor!IDCliente
                End If
                TBFornecedor.Close
                txttipo = "C"
            End With
        End If
        
        ProcNovo
        ProcSalvar
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from CQ_RNC where id = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            ProcPuxaDados
        End If
        TBAbrir.Close
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

Formulario = "Qualidade/RNC"
Direitos
ProcLimpaVariaveisPrincipais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAbrir()
On Error GoTo tratar_erro

frmQualidade_RNC_abrir.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procAtualiza()
On Error GoTo tratar_erro

If InputBox("Informe a senha para liberar.") = "280362R" Then
    If USMsgBox("Deseja realmente atualizar os codigos das RNC?", vbYesNo, "CAPRIND v5.0") = vbYes Then
        Set TBGravar = CreateObject("adodb.recordset")
        TBGravar.Open "Select * from cq_rnc", Conexao, adOpenKeyset, adLockOptimistic
        If TBGravar.EOF = False Then
            Do While TBGravar.EOF = False
                Data_Prog = IIf(IsNull(TBGravar!Data), Format(Date, "yy"), Format(TBGravar!Data, "yy"))
                Cont = TBGravar!ID
                ProcGeraNumero
                TBGravar!id_texto = a
                
                'Atualiza descrição do produto
                Set TBProduto = CreateObject("adodb.recordset")
                TBProduto.Open "Select * from projproduto where desenho = '" & TBGravar!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBProduto.EOF = False Then
                    TBGravar!Descricao = IIf(IsNull(TBProduto!Descricao), "", TBProduto!Descricao)
                End If
                TBProduto.Close
                
                'Atualiza cliente/fornecedor
                Set TBFornecedor = CreateObject("adodb.recordset")
                If TBGravar!Tipo = "F" Then
                    TBFornecedor.Open "Select * from Compras_fornecedores where idcliente = " & TBGravar!ID_forn, Conexao, adOpenKeyset, adLockOptimistic
                    If TBFornecedor.EOF = False Then
                        TBGravar!Cliente_forn = IIf(IsNull(TBFornecedor!Nome_Razao), "", TBFornecedor!Nome_Razao)
                    End If
                    TBFornecedor.Close
                Else
                    TBFornecedor.Open "Select * from clientes where idcliente = " & TBGravar!ID_forn, Conexao, adOpenKeyset, adLockOptimistic
                    If TBFornecedor.EOF = False Then
                        TBGravar!Cliente_forn = IIf(IsNull(TBFornecedor!NomeRazao), "", TBFornecedor!NomeRazao)
                    End If
                    TBFornecedor.Close
                End If
                
                TBGravar.Update
                TBGravar.MoveNext
            Loop
        End If
        TBGravar.Close
        USMsgBox ("Atualização efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
        '==================================
        Modulo = "Qualidade/RNC"
        Evento = "Atualizar"
        ID_documento = 0
        Documento = ""
        Documento1 = ""
        ProcGravaEvento
        '==================================
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ImgCalendario1_Click()
On Error GoTo tratar_erro

Sit_Data = 1
ProcAbrirCalendario

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAbrirCalendario()
On Error GoTo tratar_erro

Faturamento = False
Compras_Pedido = False
Compras_Requisicao = False
Compras_Fallow_up = False
Vendas_Carteira = False
Vendas_Proposta = False
Vendas_PI = False
Manutencao = False
Compras_Cotacao = False
Usuarios = False
Inspecao_recebimento = False
Funcionario = False
RNC = True
SolicitacaoAcao = False
Troca_Duplicata = False
Financeiro_Contas_Recebidas = False
Engenharia_Normas = False
Qualidade_PPAP_PSW = False
Qualidade_PPAP_Plano = False
Qualidade_PPAP_FMEA = False
Qualidade_sistema = False
Engenharia = False
Compras_Fornecedores = False
Vendas_Programacao = False
Outros_solicitacaoPCP = False
Estoque_recebimento = False
FrmCalendario.Show 1

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
With ListView1
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir esta(s) RNC('s)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            
            Conexao.Execute "DELETE from CQ_RNC where id = " & .ListItems(InitFor)
            'Exclui RNC na inspecao de recebimento, controle de medicao, nao conformidade e solicitação de desvio
            If RNC_Inspecao_Recebimento = True Then
                Conexao.Execute "Update Compras_recebimento Set ID_RNC = Null where id = " & frmCompras_recebimento.txtId
                frmCompras_recebimento.Txtsac = ""
            ElseIf RNC_Controle_Medicao = True Then
                    Conexao.Execute "Update Medicao Set ID_RNC = Null where IdPlano = " & frmPlanomedicao.txtPm
                    frmPlanomedicao.txtRNC = ""
                ElseIf RNC_Nao_Conformidade = True Then
                        Conexao.Execute "Update CQ_NC_FABRICA Set ID_RNC = Null where Codigo = " & IIf(frmcqnc.txtidos = "", 0, frmcqnc.txtidos)
                        frmcqnc.txtRNC = ""
                    ElseIf RNC_Solicitacao_Desvio = True Then
                            Conexao.Execute "Update CQ_SD Set ID_RNC = Null where ID = " & IIf(frmCQ_SD.txtId = "", 0, frmCQ_SD.txtId)
                            frmCQ_SD.txtRNC = ""
                        Else
                            Conexao.Execute "Update Compras_recebimento Set ID_RNC = Null where ID_RNC = " & .ListItems(InitFor)
                            Conexao.Execute "Update Medicao Set ID_RNC = Null where ID_RNC = " & .ListItems(InitFor)
                            Conexao.Execute "Update CQ_NC_FABRICA Set ID_RNC = Null where ID_RNC = " & .ListItems(InitFor)
                            Conexao.Execute "Update CQ_SD Set ID_RNC = Null where ID_RNC = " & .ListItems(InitFor)
            End If
            
            '==================================
            Modulo = "Qualidade/RNC"
            Evento = "Excluir"
            ID_documento = .ListItems(InitFor)
            Documento = IIf(.ListItems(InitFor).ListSubItems(2) = "", "Nº RNC: " & .ListItems(InitFor).ListSubItems(1), "Nº RNC: " & .ListItems(InitFor).ListSubItems(1) & " - Sequencial: " & .ListItems(InitFor).ListSubItems(2))
            Documento1 = ""
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe a(s) RNC('s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("RNC('s) excluída(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCampos
    ProcCarregaLista (1)
    Novo_RNC = False
    Frame1.Enabled = False
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro

frmQualidade_RNC_menuimpressao.Show 1

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
If RNC_Inspecao_Recebimento = False And RNC_Controle_Medicao = False And RNC_Nao_Conformidade = False And RNC_Solicitacao_Desvio = False Then
    Frame1.Enabled = True
    ProcLimpaCampos
    Cmd_localizar_item.Enabled = True
    Cmd_localizar_cliente_fornecedor.Enabled = True
End If
Novo_RNC = True
If txtID_texto.Visible = True Then txtID_texto.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procNovo_doc()
On Error GoTo tratar_erro

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Proclimpacampos_doc
Novo_RNC1 = True
Frame14.Enabled = True
txtData_doc.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

If Novo_RNC = True Then
    If USMsgBox("A RNC ainda não foi salva, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvar
        If Novo_RNC = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
If Novo_RNC1 = True Then
    If USMsgBox("O documento ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        procSalvar_doc
        If Novo_RNC1 = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
Novo_RNC = False
Novo_RNC1 = False
ProcExcluirDadosProducaoRelatoriosTotal

Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcSalvar()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Acao = "salvar"
If Frame1.Enabled = False And RNC_Inspecao_Recebimento = False And RNC_Controle_Medicao = False And RNC_Nao_Conformidade = False And RNC_Solicitacao_Desvio = False Then
    ProcVerificaSalvar
    Exit Sub
ElseIf Frame1.Enabled = False And Novo_RNC = False Then
        If RNC_Inspecao_Recebimento = True Then NomeCampo = "Qualidade/Inspeção de recebimento"
        If RNC_Controle_Medicao = True Then NomeCampo = "Qualidade/Controle de medição"
        If RNC_Nao_Conformidade = True Then NomeCampo = "Qualidade/Não conformidade"
        If RNC_Solicitacao_Desvio = True Then NomeCampo = "Qualidade/Solicitação de desvio"
        USMsgBox ("E necessário clicar novamente no botão de criar RNC no módulo de " & NomeCampo & "."), vbExclamation, "CAPRIND v5.0"
        Unload Me
        Exit Sub
End If
If txtID_texto <> "____-__" Then
    Contador = 1
    Do While Contador < 8
        If Mid(txtID_texto, Contador, 1) = "_" Then
            NomeCampo = "o número da RNC"
            ProcVerificaAcao
            txtID_texto.SetFocus
            Exit Sub
        End If
        Contador = Contador + 1
    Loop
End If

OF = 0
TextoFiltro = ""
If Txt_sequencial <> "__" Then
    Contador = 1
    Do While Contador < 3
        If Mid(Txt_sequencial, Contador, 1) = "_" Then
            NomeCampo = "o sequencial do número da RNC"
            ProcVerificaAcao
            Txt_sequencial.SetFocus
            Exit Sub
        End If
        Contador = Contador + 1
    Loop
    OF = Txt_sequencial
    TextoFiltro = " and Seq = " & OF
End If

If txtdesenho = "" Then
    NomeCampo = "o código interno"
    ProcVerificaAcao
    frmQualidade_RNC_item.Show 1
    Exit Sub
End If
If txtdescricao = "" Then
    NomeCampo = "a descrição"
    ProcVerificaAcao
    txtdescricao.SetFocus
    Exit Sub
End If
Qtde = IIf(txtQtde = "", 0, txtQtde)
If Qtde < 0 Then
    NomeCampo = "a quantidade não conforme"
    ProcVerificaAcao
    txtQtde.SetFocus
    Exit Sub
End If
If txtfim <> "__/__/____" Then
    If IsDate(txtfim) = False Then
        NomeCampo = "a data de fechamento"
        ProcVerificaAcao
        txtfim.SetFocus
        Exit Sub
    End If
End If

'Verifica se já exite RNC com este número
If txtID_texto <> "____-__" Then
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from CQ_RNC where id_texto = '" & txtID_texto & "' " & TextoFiltro & " and ID <> " & IIf(txtId = "", 0, txtId), Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        USMsgBox ("Este número de RNC está sendo utilizado, favor alterar."), vbExclamation, "CAPRIND v5.0"
        txtID_texto.SetFocus
        TBAbrir.Close
        Exit Sub
    End If
    TBAbrir.Close
End If

Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from CQ_RNC where ID = " & IIf(txtId = "", 0, txtId), Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then
    If txtID_texto = "____-__" Then
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from CQ_RNC where Year(data) = '" & Year(Date) & "' order by id_texto", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = True Then
            Cont = 1
        Else
            TBAbrir.MoveLast
            Cont = ReturnNumbersOnly(Left(TBAbrir!id_texto, 4))
            Cont = Cont + 1
        End If
        Data_Prog = Format(Date, "yy")
        ProcGeraNumero
        txtID_texto = a
    End If
    
    TBGravar.AddNew
End If
ProcEnviaDados
TBGravar.Update
txtId = TBGravar!ID
TBGravar.Close
If Novo_RNC = True Then
    USMsgBox ("Nova RNC cadastrada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo"
    StrSql_CQ_RNC_Localizar = "Select * from CQ_RNC where ID = " & txtId
    ProcCarregaLista (1)
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar"
    ProcCarregaLista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
    If CodigoLista <> 0 And ListView1.ListItems.Count <> 0 Then
        ListView1.SelectedItem = ListView1.ListItems(CodigoLista)
        ListView1.SetFocus
    End If
End If
'==================================
Modulo = "Qualidade/RNC"
ID_documento = txtId.Text
Documento = IIf(Txt_sequencial = "__", "Nº RNC: " & txtID_texto, "Nº RNC: " & txtID_texto & " - Sequencial: " & Txt_sequencial)
Documento1 = ""
ProcGravaEvento
'==================================
Novo_RNC = False

'Grava RNC na inspecao de recebimento, controle de medicao, nao conformidade e solicitação de desvio
If RNC_Inspecao_Recebimento = True Then
    Conexao.Execute "Update Compras_recebimento Set ID_RNC = " & txtId & " where id = " & frmCompras_recebimento.txtId
    With frmCompras_recebimento
        .Txt_ID_RNC = txtId
        .Txtsac = txtID_texto
    End With
End If
If RNC_Controle_Medicao = True Then
    Conexao.Execute "Update Medicao Set ID_RNC = " & txtId & " where IdPlano = " & frmPlanomedicao.txtPm
    With frmPlanomedicao
        .Txt_ID_RNC = txtId
        .txtRNC = txtID_texto
    End With
End If
If RNC_Nao_Conformidade = True Then
    If Sit_REG = 1 Then
        Conexao.Execute "Update CQ_NC_FABRICA Set ID_RNC = " & txtId & " where Codigo = " & IIf(frmcqnc.txtidos = "", 0, frmcqnc.txtidos)
        With frmcqnc
            .Txt_ID_RNC = txtId
            .txtRNC = txtID_texto
        End With
    Else
        With frmcqnc.ListaFases
            For InitFor = 1 To .ListItems.Count
                If .ListItems.Item(InitFor).Checked = True Then
                    Conexao.Execute "Update CQ_NC_FABRICA Set ID_RNC = " & txtId & " where Codigo = " & .ListItems(InitFor)
                End If
            Next InitFor
        End With
    End If
End If
If RNC_Solicitacao_Desvio = True Then
    Conexao.Execute "Update CQ_SD Set ID_RNC = " & txtId & " where ID = " & IIf(frmCQ_SD.txtId = "", 0, frmCQ_SD.txtId)
    With frmcqnc
        .Txt_ID_RNC = txtId
        .txtRNC = txtID_texto
    End With
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With ListView1
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then .ListItems.Item(InitFor).Checked = False Else .ListItems.Item(InitFor).Checked = True
        Next InitFor
    End With
Else
    ProcOrdenaListView ListView1, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If ListView1.ListItems.Count = 0 Then Exit Sub
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from CQ_RNC where id = " & ListView1.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    ProcLimpaCampos
    ProcPuxaDados
    CodigoLista = ListView1.SelectedItem.index
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error GoTo tratar_erro

If txtId = 0 Then
    SSTab1.Tab = 0
    Exit Sub
End If
PBLista.Visible = True
Select Case SSTab1.Tab
    Case 0:
        If ListView1.Visible = True And ListView1.Enabled = True Then ListView1.SetFocus
    Case 1:
        PBLista.Visible = False
        ListView1.SetFocus
        ProcVerificaProsseguir
        If Permitido = False Then Exit Sub
        ProcLimpaCampos2
        ProcPuxadados2
    Case 2:
        PBLista.Visible = False
        txtData3.SetFocus
        ProcVerificaProsseguir
        If Permitido = False Then Exit Sub
        ProcLimpaCampos3
        ProcPuxadados3
    Case 3:
        Lista_doc.SetFocus
        ProcVerificaProsseguir
        If Permitido = False Then Exit Sub
        ProcCarregaLista_Doc
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcVerificaProsseguir()
On Error GoTo tratar_erro

Permitido = True
If Novo_RNC = True Then
    USMsgBox ("Salve a RNC antes de prosseguir."), vbExclamation, "CAPRIND v5.0"
    SSTab1.Tab = 0
    Permitido = False
    Exit Sub
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaLista_Doc()
On Error GoTo tratar_erro

Lista_doc.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "select * from CQ_RNC_documentos where ID_RNC = " & txtId & " order by ID", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    TBLISTA.MoveLast
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    TBLISTA.MoveFirst
    Do While TBLISTA.EOF = False
        With Lista_doc.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Caminho_documento), "", TBLISTA!Caminho_documento)
        End With
        TBLISTA.MoveNext
        Contador = Contador + 1
        PBLista.Value = Contador
    Loop
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub SSTab2_Click(PreviousTab As Integer)
On Error GoTo tratar_erro

Select Case SSTab2.Tab
    Case 0: If txtData3.Visible = True Then txtData3.SetFocus
    Case 1: If txtData3.Visible = True Then txtData4.SetFocus
End Select


ListView2.ListItems.Clear
Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from CQ_SA where RNC = '" & txtID_texto & "'", Conexao, adOpenKeyset, adLockOptimistic
    Do While TBAbrir.EOF = False
        With ListView2.ListItems
            .Add , , TBAbrir!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBAbrir!Tipo), "", TBAbrir!Tipo)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBAbrir!Previsao), "", TBAbrir!Previsao)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBAbrir!ResponsavelSA), "", TBAbrir!ResponsavelSA)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBAbrir!Objetivo), "", TBAbrir!Objetivo)
            TBAbrir.MoveNext
        End With
    Loop
    TBAbrir.Close



Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub




Private Sub txtData5_LostFocus()
On Error GoTo tratar_erro

If txtData5 <> "__/__/____" Then
    VerifData = txtData5
    ProcVerificaData
    If VerifData = False Then
        txtData5 = "__/__/____"
        txtData5.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtData6_LostFocus()
On Error GoTo tratar_erro

If txtData6 <> "__/__/____" Then
    VerifData = txtData6
    ProcVerificaData
    If VerifData = False Then
        txtData6 = "__/__/____"
        txtData6.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtdata7_LostFocus()
On Error GoTo tratar_erro

If txtdata7 <> "__/__/____" Then
    VerifData = txtdata7
    ProcVerificaData
    If VerifData = False Then
        txtdata7 = "__/__/____"
        txtdata7.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtDesenho_Change()
On Error GoTo tratar_erro

txtdescricao = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtdesenho_LostFocus()
On Error GoTo tratar_erro

With txtdescricao
    .Locked = False
    .TabStop = True
    If txtdesenho <> "" Then
        Set TBProduto = CreateObject("adodb.recordset")
        TBProduto.Open "Select * from projproduto WHERE desenho = '" & txtdesenho & "' and Bloqueado = 'False'", Conexao, adOpenKeyset, adLockOptimistic
        If TBProduto.EOF = False Then
            txtdesenho = IIf(IsNull(TBProduto!Desenho), "", TBProduto!Desenho)
            txtdescricao = IIf(IsNull(TBProduto!Descricao), "", TBProduto!Descricao)
            .Locked = True
            .TabStop = False
        End If
        TBProduto.Close
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub Procbloqueia()
On Error GoTo tratar_erro

Framelista.Enabled = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcDesbloqueia()
On Error GoTo tratar_erro

Framelista.Enabled = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaLista(Pagina As Integer)
On Error GoTo tratar_erro

ListView1.ListItems.Clear
lblRegistros.Caption = "Nº de registros: 0"
lblPaginas.Caption = "Página: 0 de: 0"
If StrSql_CQ_RNC_Localizar = "" Then Exit Sub
Set TBLISTA_CQ_RNC = CreateObject("adodb.recordset")
TBLISTA_CQ_RNC.Open StrSql_CQ_RNC_Localizar, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA_CQ_RNC.EOF = False Then ProcExibePagina (Pagina)
       
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina(Pagina)
On Error GoTo tratar_erro

ListView1.ListItems.Clear
TBLISTA_CQ_RNC.PageSize = IIf(txtNreg = "", 30, txtNreg)
TBLISTA_CQ_RNC.AbsolutePage = Pagina
TamanhoPagina = TBLISTA_CQ_RNC.PageSize
ContadorReg = 1

PBLista.Min = 0
PBLista.Max = FunVerifMaxPBListaPaginacao(TBLISTA_CQ_RNC.RecordCount - IIf(Pagina > 1, (TBLISTA_CQ_RNC.PageSize * (Pagina - 1)), 0), TBLISTA_CQ_RNC.PageSize)
PBLista.Value = 1
Contador = 0
Do While TBLISTA_CQ_RNC.EOF = False And (ContadorReg <= TamanhoPagina)
    With ListView1.ListItems
        .Add , , TBLISTA_CQ_RNC!ID
        .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA_CQ_RNC!id_texto), "", TBLISTA_CQ_RNC!id_texto)
        If IsNull(TBLISTA_CQ_RNC!Seq) = False Then If TBLISTA_CQ_RNC!Seq < 10 Then .Item(.Count).SubItems(2) = "0" & TBLISTA_CQ_RNC!Seq Else .Item(.Count).SubItems(2) = TBLISTA_CQ_RNC!Seq
        .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA_CQ_RNC!Data), "", Format(TBLISTA_CQ_RNC!Data, "dd/mm/yy"))
        .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA_CQ_RNC!Responsavel), "", TBLISTA_CQ_RNC!Responsavel)
        .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA_CQ_RNC!Desenho), "", TBLISTA_CQ_RNC!Desenho)
        .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA_CQ_RNC!Descricao), "", TBLISTA_CQ_RNC!Descricao)
        .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA_CQ_RNC!Qtde), "0,0000", Format(TBLISTA_CQ_RNC!Qtde, "###,##0.0000"))
        .Item(.Count).SubItems(8) = IIf(IsNull(TBLISTA_CQ_RNC!txtCustoNC), "0,00", Format(TBLISTA_CQ_RNC!txtCustoNC, "###,##0.00"))
    End With
    TBLISTA_CQ_RNC.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista.Value = Contador
Loop
lblRegistros.Caption = "Nº de registros: " & TBLISTA_CQ_RNC.RecordCount
If TBLISTA_CQ_RNC.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "Página: 1 de: " & TBLISTA_CQ_RNC.PageCount
ElseIf TBLISTA_CQ_RNC.AbsolutePage = adPosEOF Then
        lblPaginas.Caption = "Página: " & TBLISTA_CQ_RNC.PageCount & " de: " & TBLISTA_CQ_RNC.PageCount
    Else
        lblPaginas.Caption = "Página: " & TBLISTA_CQ_RNC.AbsolutePage - 1 & " de: " & TBLISTA_CQ_RNC.PageCount
End If


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcEnviaDados()
On Error GoTo tratar_erro

TBGravar!id_texto = txtID_texto
TBGravar!Seq = IIf(OF = 0, Null, OF)
TBGravar!Data = txtData
TBGravar!status = IIf(cmbStatus = "", Null, cmbStatus)
If Opt_interno.Value = True Then TBGravar!Origem = "I" Else TBGravar!Origem = "E"
If optPreventiva.Value = True Then TBGravar!Preventiva = True Else TBGravar!Preventiva = False
If Chk_acao_corretiva.Value = 1 Then
    TBGravar!Acao_corretiva = True
Else
    TBGravar!Acao_corretiva = False
    
    TBGravar!Data2 = Null
    TBGravar!responsavel2 = Null
    TBGravar!Texto1 = Null
    TBGravar!Texto2 = Null
    TBGravar!Data3 = Null
    TBGravar!responsavel3 = Null
    TBGravar!Texto3 = Null
    TBGravar!texto4 = Null
    TBGravar!Texto5 = Null
    TBGravar!Texto6 = Null
    TBGravar!Texto7 = Null
    TBGravar!Data4 = Null
    TBGravar!responsavel4 = Null
    TBGravar!Acao = Null
    TBGravar!Acompanhamento = Null
    TBGravar!Data5 = Null
    TBGravar!responsavel6 = Null
    TBGravar!Eficacia = Null
    TBGravar!Data6 = Null
    TBGravar!responsavel5 = Null
    TBGravar!Texto8 = Null
    TBGravar!Data7 = Null
    TBGravar!responsavel7 = Null
End If
    
TBGravar!Desenho = txtdesenho
TBGravar!Descricao = txtdescricao
TBGravar!Qtde = IIf(txtQtde = "", 0, txtQtde)
TBGravar!ID_forn = txtID_forn
TBGravar!Cliente_forn = IIf(txtFornecedor = "", Null, txtFornecedor)
TBGravar!Tipo = txttipo
If cmbEficaz = "Sim" Then TBGravar!FIM = IIf(txtfim = "__/__/____", Null, txtfim) Else TBGravar!FIM = Null
TBGravar!eficaz = IIf(cmbEficaz = "", Null, cmbEficaz)
TBGravar!Documento_ref = Txt_doc_referencia
TBGravar!nao_conformidade = txtNao_conformidade
TBGravar!Equipe = txtEquipe
TBGravar!classificacao = txtClassificacao
TBGravar!imagem = IIf(Txt_imagemRNC = "", Null, Txt_imagemRNC)
TBGravar!QtdeLote = IIf(txt_qtdeLote = "", Null, txt_qtdeLote)
TBGravar!QtdeAprovada = IIf(txt_QtdeAprovada = "", Null, txt_QtdeAprovada)
TBGravar!Procede = IIf(cmbProcede = "", Null, cmbProcede)
TBGravar!Requisitos = IIf(cmbRequisitos = "", Null, cmbRequisitos)
TBGravar!QuaisRequisitos = IIf(txtQuaisRequisitos = "", Null, txtQuaisRequisitos)
TBGravar!Similaridade = IIf(cmbSimilaridade = "", Null, cmbSimilaridade)
TBGravar!QuaisRNC = IIf(txtQuaisRNC = "", Null, txtQuaisRNC)
TBGravar!MateriaPrima = IIf(cmbMateriaPrima = "", Null, cmbMateriaPrima)
TBGravar!txtvalorunitario = IIf(txtvalorunitario = "", Null, txtvalorunitario)
TBGravar!txtFatorNC = IIf(txtFatorNC = "", Null, txtFatorNC)
TBGravar!txtCustoNC = IIf(txtCustoNC = "", Null, txtCustoNC)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub




Private Sub txtFatorNC_LostFocus()
On Error GoTo tratar_erro
Dim Valorunitario As Double
Dim Fator As Double
Dim Custo As Double

Valorunitario = txtvalorunitario
Fator = txtFatorNC
Custo = Valorunitario * Fator
txtCustoNC = Custo

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtfim_LostFocus()
On Error GoTo tratar_erro

If txtfim <> "__/__/____" Then
    VerifData = txtfim
    ProcVerificaData
    If VerifData = False Then
        txtfim = "__/__/____"
        txtfim.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtFornecedor_LostFocus()
On Error GoTo tratar_erro

txtID_forn = 0
txttipo = ""
If txtFornecedor <> "" Then
    Set TBFornecedor = CreateObject("adodb.recordset")
    TBFornecedor.Open "Select * from Compras_fornecedores where Nome_Razao = '" & txtFornecedor & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBFornecedor.EOF = False Then
        txtID_forn = IIf(IsNull(TBFornecedor!IDCliente), "", TBFornecedor!IDCliente)
        txtFornecedor = IIf(IsNull(TBFornecedor!Nome_Razao), "", TBFornecedor!Nome_Razao)
        txttipo = "F"
    Else
        Set TBFornecedor = CreateObject("adodb.recordset")
        TBFornecedor.Open "Select * from clientes where NomeRazao = '" & txtFornecedor & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBFornecedor.EOF = False Then
            txtID_forn = IIf(IsNull(TBFornecedor!IDCliente), "", TBFornecedor!IDCliente)
            txtFornecedor = IIf(IsNull(TBFornecedor!NomeRazao), "", TBFornecedor!NomeRazao)
            txttipo = "C"
        End If
    End If
    TBFornecedor.Close
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

Private Sub txtQtde_Change()
On Error GoTo tratar_erro

If txtQtde.Text <> "" Then
    VerifNumero = txtQtde.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtQtde.Text = ""
        txtQtde.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtQtde_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus txtQtde

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtqtde_LostFocus()
On Error GoTo tratar_erro

txtQtde = Format(txtQtde, "###,##0.0000")
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcPuxadados2()
On Error GoTo tratar_erro

Set TBFornecedor = CreateObject("adodb.recordset")
TBFornecedor.Open "Select * from CQ_RNC where id = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
If TBFornecedor.EOF = False Then
    txtData2 = IIf(IsNull(TBFornecedor!Data2), Date, TBFornecedor!Data2)
    txtResponsavel2 = IIf(IsNull(TBFornecedor!responsavel2), "", TBFornecedor!responsavel2)
    txtTexto1 = IIf(IsNull(TBFornecedor!Texto1), "", TBFornecedor!Texto1)
    txtTexto2 = IIf(IsNull(TBFornecedor!Texto2), "", TBFornecedor!Texto2)
End If
TBFornecedor.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcPuxadados3()
On Error GoTo tratar_erro

Set TBFornecedor = CreateObject("adodb.recordset")
TBFornecedor.Open "Select * from CQ_RNC where id = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
If TBFornecedor.EOF = False Then
    txtData3 = IIf(IsNull(TBFornecedor!Data3), Date, TBFornecedor!Data3)
    txtResponsavel3 = IIf(IsNull(TBFornecedor!responsavel3), "", TBFornecedor!responsavel3)
    txtTexto3 = IIf(IsNull(TBFornecedor!Texto3), "", TBFornecedor!Texto3)
    txtTexto4 = IIf(IsNull(TBFornecedor!texto4), "", TBFornecedor!texto4)
    txtTexto5 = IIf(IsNull(TBFornecedor!Texto5), "", TBFornecedor!Texto5)
    txtTexto6 = IIf(IsNull(TBFornecedor!Texto6), "", TBFornecedor!Texto6)
    txtTexto7 = IIf(IsNull(TBFornecedor!Texto7), "", TBFornecedor!Texto7)
    txtData4 = IIf(IsNull(TBFornecedor!Data4), Date, TBFornecedor!Data4)
    txtResponsavel4 = IIf(IsNull(TBFornecedor!responsavel4), "", TBFornecedor!responsavel4)
    If TBFornecedor!Acao = True Then optSim.Value = True Else optNao.Value = True
    txtAcompanhamento = IIf(IsNull(TBFornecedor!Acompanhamento), "", TBFornecedor!Acompanhamento)
    txtData5 = IIf(IsNull(TBFornecedor!Data5), "__/__/____", Format(TBFornecedor!Data5, "dd/mm/yyyy"))
    txtEficacia = IIf(IsNull(TBFornecedor!Eficacia), "", TBFornecedor!Eficacia)
    txtData6 = IIf(IsNull(TBFornecedor!Data6), "__/__/____", Format(TBFornecedor!Data6, "dd/mm/yyyy"))
    txtResponsavel5 = IIf(IsNull(TBFornecedor!responsavel5), "", TBFornecedor!responsavel5)
    txtTexto8 = IIf(IsNull(TBFornecedor!Texto8), "", TBFornecedor!Texto8)
    txtdata7 = IIf(IsNull(TBFornecedor!Data7), "__/__/____", Format(TBFornecedor!Data7, "dd/mm/yyyy"))
    txtResponsavel7 = IIf(IsNull(TBFornecedor!responsavel7), "", TBFornecedor!responsavel7)
    txtResponsavel6 = IIf(IsNull(TBFornecedor!responsavel6), "", TBFornecedor!responsavel6)
    
    If (IsNull(TBFornecedor!chkMaoObra)) = True Or TBFornecedor!chkMaoObra = False Then chkMaoObra.Value = 0 Else chkMaoObra.Value = 1
    If (IsNull(TBFornecedor!chkMaquina)) = True Or TBFornecedor!chkMaquina = False Then chkMaquina.Value = 0 Else chkMaquina.Value = 1
    If (IsNull(TBFornecedor!chkMetodo)) = True Or TBFornecedor!chkMetodo = False Then chkMetodo.Value = 0 Else chkMetodo.Value = 1
    If (IsNull(TBFornecedor!chkMedidas)) = True Or TBFornecedor!chkMedidas = False Then chkMedidas.Value = 0 Else chkMedidas.Value = 1
    If (IsNull(TBFornecedor!chkMaterial)) = True Or TBFornecedor!chkMaterial = False Then chkMaterial.Value = 0 Else chkMaterial.Value = 1
    If (IsNull(TBFornecedor!chkMeioAmbiente)) = True Or TBFornecedor!chkMeioAmbiente = False Then chkMeioAmbiente.Value = 0 Else chkMeioAmbiente.Value = 1
    If (IsNull(TBFornecedor!chkSGQ)) = True Or TBFornecedor!chkSGQ = False Then chkSGQ.Value = 0 Else chkSGQ.Value = 1
    If (IsNull(TBFornecedor!chkOutros)) = True Or TBFornecedor!chkOutros = False Then chkOutros.Value = 0 Else chkOutros.Value = 1
    If (IsNull(TBFornecedor!chkRevisao)) = True Or TBFornecedor!chkRevisao = False Then chkRevisao.Value = 0 Else chkRevisao.Value = 1
    txtrevisao = IIf(IsNull(TBFornecedor!txtrevisao), "", TBFornecedor!txtrevisao)
    
    
End If
TBFornecedor.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcPuxadados_Doc()
On Error GoTo tratar_erro

txtID_doc = TBAbrir!ID
txtData_doc = IIf(IsNull(TBAbrir!Data), "", Format(TBAbrir!Data, "dd/mm/yy"))
txtResponsavel_doc = IIf(IsNull(TBAbrir!Responsavel), "", TBAbrir!Responsavel)
txt_Caminho = IIf(IsNull(TBAbrir!Caminho_documento), "", TBAbrir!Caminho_documento)
Txt_obs_doc = IIf(IsNull(TBAbrir!Observacao), "", TBAbrir!Observacao)
Novo_RNC1 = False
Frame14.Enabled = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcGeraNumero()
On Error GoTo tratar_erro

a = Cont
Select Case Len(a)
    Case 1: a = "000" & Cont & "-" & Data_Prog
    Case 2: a = "00" & Cont & "-" & Data_Prog
    Case 3: a = "0" & Cont & "-" & Data_Prog
    Case 4: a = Cont & "-" & Data_Prog
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcNovo
    Case 2: ProcAbrir
    Case 3: ProcSalvar
    Case 4: ProcExcluir
    Case 5: ProcImprimir
    Case 6: ProcAnterior
    Case 7: ProcProximo
    Case 8: procAtualiza
    'Case 10: ProcAjuda
    Case 11: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar2_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: procSalvar2
    Case 2: procExcluir2
    Case 3: ProcImprimir
    Case 4: ProcAnterior
    Case 5: ProcProximo
    'Case 7: ProcAjuda
    Case 8: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar3_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: procSalvar3
    Case 2: procExcluir3
    Case 3: ProcImprimir
    Case 4: ProcAnterior
    Case 5: ProcProximo
    'Case 7: ProcAjuda
    Case 8: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar4_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: procNovo_doc
    Case 2: procSalvar_doc
    Case 3: procExcluir_doc
    Case 4: ProcImprimir
    Case 5: ProcAnterior
    Case 6: ProcProximo
    'Case 8: ProcAjuda
    Case 9: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
