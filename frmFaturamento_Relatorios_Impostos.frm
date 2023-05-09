VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFaturamento_Relatorios_Impostos 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Administrativo - Faturamento - Relatórios - Impostos"
   ClientHeight    =   10035
   ClientLeft      =   2145
   ClientTop       =   1860
   ClientWidth     =   15360
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
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   0
      Top             =   0
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
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   13110
      Top             =   510
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frmFaturamento_Relatorios_Impostos.frx":0000
      Count           =   1
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   90
      TabIndex        =   60
      Top             =   330
      Width           =   15225
      _ExtentX        =   26855
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft1     =   2
      ButtonTop1      =   2
      ButtonWidth1    =   42
      ButtonHeight1   =   21
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft2     =   46
      ButtonTop2      =   2
      ButtonWidth2    =   60
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
      ButtonLeft3     =   108
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft4     =   112
      ButtonTop4      =   2
      ButtonWidth4    =   41
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft5     =   155
      ButtonTop5      =   2
      ButtonWidth5    =   30
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
      ButtonLeft6     =   187
      ButtonTop6      =   2
      ButtonWidth6    =   24
      ButtonHeight6   =   24
      ButtonUseMaskColor6=   0   'False
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   90
      TabIndex        =   59
      Top             =   9720
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
   Begin VB.Frame Frame21 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Filtrar notas:                                               "
      Height          =   765
      Left            =   90
      TabIndex        =   46
      Top             =   1320
      Width           =   15195
      Begin VB.OptionButton Opt_periodo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Por período"
         Height          =   195
         Left            =   11460
         TabIndex        =   2
         Top             =   30
         Width           =   1125
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   765
         Left            =   11370
         TabIndex        =   119
         Top             =   0
         Width           =   3795
         Begin MSComCtl2.DTPicker msk_fltFim 
            Height          =   315
            Left            =   2430
            TabIndex        =   120
            ToolTipText     =   "Data final."
            Top             =   300
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
            Format          =   109707265
            CurrentDate     =   39057
         End
         Begin MSComCtl2.DTPicker msk_fltInicio 
            Height          =   315
            Left            =   570
            TabIndex        =   121
            ToolTipText     =   "Data inicio."
            Top             =   300
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
            Format          =   109707265
            CurrentDate     =   39057
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Até :"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   1980
            TabIndex        =   123
            Top             =   300
            Width           =   360
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "De :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   1
            Left            =   180
            TabIndex        =   122
            Top             =   300
            Width           =   300
         End
      End
      Begin VB.OptionButton OptDomes 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Do mês"
         Height          =   195
         Left            =   1200
         TabIndex        =   0
         Top             =   30
         Value           =   -1  'True
         Width           =   825
      End
      Begin VB.OptionButton OptAteomes 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Até o mês"
         Height          =   195
         Left            =   2130
         TabIndex        =   1
         Top             =   30
         Width           =   1035
      End
      Begin VB.ComboBox cmbAno 
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
         Height          =   300
         ItemData        =   "frmFaturamento_Relatorios_Impostos.frx":2DF9
         Left            =   10170
         List            =   "frmFaturamento_Relatorios_Impostos.frx":2DFB
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   450
         Width           =   795
      End
      Begin MSComctlLib.TabStrip TabFiltro 
         Height          =   345
         Left            =   180
         TabIndex        =   3
         Top             =   450
         Width           =   10125
         _ExtentX        =   17859
         _ExtentY        =   609
         MultiRow        =   -1  'True
         TabMinWidth     =   1411
         TabStyle        =   1
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   12
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Janeiro"
               Key             =   "Jan"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Fevereiro"
               Key             =   "Fev"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Março"
               Key             =   "Mar"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Abril"
               Key             =   "Abr"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Maio"
               Key             =   "Maio"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Junho"
               Key             =   "Jun"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Julho"
               Key             =   "Jul"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Agosto"
               Key             =   "Ago"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab9 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Setembro"
               Key             =   "Set"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab10 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Outubro"
               Key             =   "Out"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab11 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Novembro"
               Key             =   "Nov"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab12 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Dezembro"
               Key             =   "Dez"
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   10095
      Left            =   0
      TabIndex        =   37
      Top             =   0
      Width           =   15360
      _ExtentX        =   27093
      _ExtentY        =   17806
      _Version        =   393216
      Tab             =   2
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
      TabCaption(0)   =   "Notas fiscais de saída"
      TabPicture(0)   =   "frmFaturamento_Relatorios_Impostos.frx":2DFD
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "ListaNF"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Notas fiscais de entrada"
      TabPicture(1)   =   "frmFaturamento_Relatorios_Impostos.frx":2E19
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ListaNF1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame3"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Impostos"
      TabPicture(2)   =   "frmFaturamento_Relatorios_Impostos.frx":2E35
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "SSTab2"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   -74910
         TabIndex        =   65
         Top             =   9090
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
            TabIndex        =   7
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
            TabIndex        =   6
            Text            =   "30"
            ToolTipText     =   "Número de registros por página."
            Top             =   180
            Width           =   555
         End
         Begin DrawSuite2022.USButton cmdPagProx 
            Height          =   315
            Left            =   11760
            TabIndex        =   11
            ToolTipText     =   "Próxima página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmFaturamento_Relatorios_Impostos.frx":2E51
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
            TabIndex        =   10
            ToolTipText     =   "Página anterior."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmFaturamento_Relatorios_Impostos.frx":65FB
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
            TabIndex        =   8
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
            TabIndex        =   9
            ToolTipText     =   "Primeira página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmFaturamento_Relatorios_Impostos.frx":A116
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
            TabIndex        =   12
            ToolTipText     =   "Última página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmFaturamento_Relatorios_Impostos.frx":E20F
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
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "registros por página"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   48
            Left            =   4410
            TabIndex        =   69
            Top             =   240
            Width           =   1440
         End
         Begin VB.Label lblPaginas 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Página: 0 de: 0"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   13050
            TabIndex        =   68
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lblRegistros 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nº de registros: 0"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   180
            TabIndex        =   67
            Top             =   240
            Width           =   1275
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Carregar"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   3090
            TabIndex        =   66
            Top             =   240
            Width           =   645
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   -74910
         TabIndex        =   61
         Top             =   9090
         Width           =   15195
         Begin VB.TextBox txtPagIr1 
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
            TabIndex        =   15
            ToolTipText     =   "Número da página."
            Top             =   180
            Width           =   555
         End
         Begin VB.TextBox txtNreg1 
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
            TabIndex        =   14
            Text            =   "30"
            ToolTipText     =   "Número de registros por página."
            Top             =   180
            Width           =   555
         End
         Begin DrawSuite2022.USButton cmdPagProx1 
            Height          =   315
            Left            =   11760
            TabIndex        =   19
            ToolTipText     =   "Próxima página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmFaturamento_Relatorios_Impostos.frx":11ADB
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
         Begin DrawSuite2022.USButton cmdPagAnt1 
            Height          =   315
            Left            =   11220
            TabIndex        =   18
            ToolTipText     =   "Página anterior."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmFaturamento_Relatorios_Impostos.frx":15282
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
         Begin DrawSuite2022.USButton cmdPagIr1 
            Height          =   315
            Left            =   10110
            TabIndex        =   16
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
         Begin DrawSuite2022.USButton cmdPagPrim1 
            Height          =   315
            Left            =   10680
            TabIndex        =   17
            ToolTipText     =   "Primeira página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmFaturamento_Relatorios_Impostos.frx":18D97
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
         Begin DrawSuite2022.USButton cmdPagUlt1 
            Height          =   315
            Left            =   12300
            TabIndex        =   20
            ToolTipText     =   "Última página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmFaturamento_Relatorios_Impostos.frx":1CE8F
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
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "registros por página"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   1
            Left            =   4410
            TabIndex        =   70
            Top             =   240
            Width           =   1440
         End
         Begin VB.Label lblPaginas1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Página: 0 de: 0"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   13050
            TabIndex        =   64
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lblRegistros1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nº de registros: 0"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   180
            TabIndex        =   63
            Top             =   240
            Width           =   1275
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Carregar"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   3090
            TabIndex        =   62
            Top             =   240
            Width           =   645
         End
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   8775
         Left            =   90
         TabIndex        =   38
         Top             =   1320
         Width           =   15255
         _ExtentX        =   26908
         _ExtentY        =   15478
         _Version        =   393216
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
         TabCaption(0)   =   "Produtos"
         TabPicture(0)   =   "frmFaturamento_Relatorios_Impostos.frx":2073A
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame10"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Serviços"
         TabPicture(1)   =   "frmFaturamento_Relatorios_Impostos.frx":20756
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame1"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Faturamento"
         TabPicture(2)   =   "frmFaturamento_Relatorios_Impostos.frx":20772
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Frame2"
         Tab(2).ControlCount=   1
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
            Height          =   8055
            Left            =   -74970
            TabIndex        =   44
            Top             =   330
            Width           =   15165
            Begin VB.TextBox txtValorDAS_entrada 
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
               Height          =   315
               Left            =   9225
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   36
               TabStop         =   0   'False
               Text            =   "0,00"
               ToolTipText     =   "Valor total de DAS entrada."
               Top             =   2400
               Width           =   2205
            End
            Begin VB.TextBox txtValorDAS_saida 
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
               Height          =   315
               Left            =   3405
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   35
               TabStop         =   0   'False
               Text            =   "0,00"
               ToolTipText     =   "Valor total de DAS saída."
               Top             =   2400
               Width           =   2205
            End
            Begin VB.Label Label13 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Valor total de DAS (entrada) :"
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
               Left            =   6645
               TabIndex        =   58
               Top             =   2400
               Width           =   2505
            End
            Begin VB.Label Label5 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Valor total de DAS (saída) :"
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
               Left            =   1050
               TabIndex        =   45
               Top             =   2400
               Width           =   2295
            End
         End
         Begin VB.Frame Frame1 
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
            Height          =   8055
            Left            =   -74970
            TabIndex        =   40
            Top             =   330
            Width           =   15165
            Begin VB.TextBox txtValorIRRF_serv_entrada 
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
               Height          =   315
               Left            =   9225
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   34
               TabStop         =   0   'False
               Text            =   "0,00"
               ToolTipText     =   "Valor total de IRRF entrada."
               Top             =   3660
               Width           =   2205
            End
            Begin VB.TextBox txtValorISS_serv_entrada 
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
               Height          =   315
               Left            =   9225
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   31
               TabStop         =   0   'False
               Text            =   "0,00"
               ToolTipText     =   "Valor total de ISSQN entrada."
               Top             =   2565
               Width           =   2205
            End
            Begin VB.TextBox txtValorINSS_serv_entrada 
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
               Height          =   315
               Left            =   9225
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   32
               TabStop         =   0   'False
               Text            =   "0,00"
               ToolTipText     =   "Valor total de INSS entrada."
               Top             =   2925
               Width           =   2205
            End
            Begin VB.TextBox txtValorIRPJ_serv_entrada 
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
               Height          =   315
               Left            =   9225
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   33
               TabStop         =   0   'False
               Text            =   "0,00"
               ToolTipText     =   "Valor total de IRPJ entrada."
               Top             =   3300
               Width           =   2205
            End
            Begin VB.TextBox txtValorIRRF_serv_saida 
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
               Height          =   315
               Left            =   3405
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   27
               TabStop         =   0   'False
               Text            =   "0,00"
               ToolTipText     =   "Valor total de IRRF saída."
               Top             =   3660
               Width           =   2205
            End
            Begin VB.TextBox txtValorPIS_serv_entrada 
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
               Height          =   315
               Left            =   9225
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   28
               TabStop         =   0   'False
               Text            =   "0,00"
               ToolTipText     =   "Valor total do PIS entrada."
               Top             =   1470
               Width           =   2205
            End
            Begin VB.TextBox txtValorCofins_serv_entrada 
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
               Height          =   315
               Left            =   9225
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   29
               TabStop         =   0   'False
               Text            =   "0,00"
               ToolTipText     =   "Valor total do Cofins entrada."
               Top             =   1830
               Width           =   2205
            End
            Begin VB.TextBox txtValorCSLL_serv_entrada 
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
               Height          =   315
               Left            =   9225
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   30
               TabStop         =   0   'False
               Text            =   "0,00"
               ToolTipText     =   "Valor total do CSLL entrada."
               Top             =   2205
               Width           =   2205
            End
            Begin VB.TextBox txtValorPIS_serv_saida 
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
               Height          =   315
               Left            =   3405
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   21
               TabStop         =   0   'False
               Text            =   "0,00"
               ToolTipText     =   "Valor total do PIS saída."
               Top             =   1470
               Width           =   2205
            End
            Begin VB.TextBox txtValorCofins_serv_saida 
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
               Height          =   315
               Left            =   3405
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   22
               TabStop         =   0   'False
               Text            =   "0,00"
               ToolTipText     =   "Valor total do Cofins saída."
               Top             =   1830
               Width           =   2205
            End
            Begin VB.TextBox txtValorCSLL_serv_saida 
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
               Height          =   315
               Left            =   3405
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   23
               TabStop         =   0   'False
               Text            =   "0,00"
               ToolTipText     =   "Valor total do CSLL saída."
               Top             =   2205
               Width           =   2205
            End
            Begin VB.TextBox txtValorIRPJ_serv_saida 
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
               Height          =   315
               Left            =   3405
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   26
               TabStop         =   0   'False
               Text            =   "0,00"
               ToolTipText     =   "Valor total de IRPJ saída."
               Top             =   3300
               Width           =   2205
            End
            Begin VB.TextBox txtValorINSS_serv_saida 
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
               Height          =   315
               Left            =   3405
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   25
               TabStop         =   0   'False
               Text            =   "0,00"
               ToolTipText     =   "Valor total de INSS saída."
               Top             =   2925
               Width           =   2205
            End
            Begin VB.TextBox txtValorISS_serv_saida 
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
               Height          =   315
               Left            =   3405
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   24
               TabStop         =   0   'False
               Text            =   "0,00"
               ToolTipText     =   "Valor total de ISSQN saída."
               Top             =   2565
               Width           =   2205
            End
            Begin VB.Label Label40 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Valor total de IRRF (entrada) :"
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
               Left            =   6555
               TabIndex        =   57
               Top             =   3660
               Width           =   2565
            End
            Begin VB.Label Label39 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Valor total de ISSQN (entrada) :"
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
               Left            =   6435
               TabIndex        =   56
               Top             =   2565
               Width           =   2685
            End
            Begin VB.Label Label38 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Valor total de INSS (entrada) :"
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
               Left            =   6555
               TabIndex        =   55
               Top             =   2925
               Width           =   2565
            End
            Begin VB.Label Label37 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Valor total de IRPJ (entrada) :"
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
               Left            =   6555
               TabIndex        =   54
               Top             =   3300
               Width           =   2565
            End
            Begin VB.Label Label36 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Valor total de IRRF (saída) :"
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
               Left            =   960
               TabIndex        =   53
               Top             =   3660
               Width           =   2355
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Valor total PIS (entrada) :"
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
               Left            =   6915
               TabIndex        =   52
               Top             =   1470
               Width           =   2205
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Valor total Cofins (entrada) :"
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
               Left            =   6705
               TabIndex        =   51
               Top             =   1830
               Width           =   2415
            End
            Begin VB.Label Label35 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Valor total CSLL (entrada) :"
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
               Left            =   6825
               TabIndex        =   50
               Top             =   2205
               Width           =   2295
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Valor total PIS (saída) :"
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
               Left            =   1350
               TabIndex        =   49
               Top             =   1470
               Width           =   1965
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Valor total Cofins (saída) :"
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
               Left            =   1110
               TabIndex        =   48
               Top             =   1830
               Width           =   2205
            End
            Begin VB.Label Label34 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Valor total CSLL (saída) :"
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
               Left            =   1230
               TabIndex        =   47
               Top             =   2205
               Width           =   2085
            End
            Begin VB.Label Label16 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Valor total de IRPJ (saída) :"
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
               Left            =   990
               TabIndex        =   43
               Top             =   3300
               Width           =   2325
            End
            Begin VB.Label txtValorINSS 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Valor total de INSS (saída) :"
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
               Left            =   990
               TabIndex        =   42
               Top             =   2925
               Width           =   2325
            End
            Begin VB.Label Label18 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Valor total de ISSQN (saída) :"
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
               Left            =   870
               TabIndex        =   41
               Top             =   2565
               Width           =   2445
            End
         End
         Begin VB.Frame Frame10 
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
            Height          =   8055
            Left            =   30
            TabIndex        =   39
            Top             =   330
            Width           =   15165
            Begin VB.Frame Frame7 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Totalização impostos"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   7485
               Left            =   90
               TabIndex        =   71
               Top             =   510
               Width           =   15015
               Begin VB.Frame Frame14 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Impostos retidos (PIS e COFINS)"
                  Height          =   1395
                  Left            =   60
                  TabIndex        =   110
                  Top             =   1860
                  Width           =   14835
                  Begin VB.TextBox txtValor_retencao_Cofins_prod_saida 
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
                     Height          =   315
                     Left            =   3405
                     Locked          =   -1  'True
                     MaxLength       =   50
                     TabIndex        =   114
                     TabStop         =   0   'False
                     Text            =   "0,00"
                     ToolTipText     =   "Valor total de retenção do Cofins saída."
                     Top             =   810
                     Width           =   915
                  End
                  Begin VB.TextBox txtValor_retencao_PIS_prod_saida 
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
                     Height          =   315
                     Left            =   3405
                     Locked          =   -1  'True
                     MaxLength       =   50
                     TabIndex        =   113
                     TabStop         =   0   'False
                     Text            =   "0,00"
                     ToolTipText     =   "Valor total de retenção do PIS saída."
                     Top             =   450
                     Width           =   915
                  End
                  Begin VB.TextBox txtValor_retencao_Cofins_prod_entrada 
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
                     Height          =   315
                     Left            =   7845
                     Locked          =   -1  'True
                     MaxLength       =   50
                     TabIndex        =   112
                     TabStop         =   0   'False
                     Text            =   "0,00"
                     ToolTipText     =   "Valor total de retenção do Cofins entrada."
                     Top             =   840
                     Width           =   915
                  End
                  Begin VB.TextBox txtValor_retencao_PIS_prod_entrada 
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
                     Height          =   315
                     Left            =   7845
                     Locked          =   -1  'True
                     MaxLength       =   50
                     TabIndex        =   111
                     TabStop         =   0   'False
                     Text            =   "0,00"
                     ToolTipText     =   "Valor total de retenção do PIS entrada."
                     Top             =   480
                     Width           =   915
                  End
                  Begin VB.Label Label25 
                     Alignment       =   2  'Center
                     Appearance      =   0  'Flat
                     AutoSize        =   -1  'True
                     BackColor       =   &H80000005&
                     BackStyle       =   0  'Transparent
                     Caption         =   "Valor total retenção Cofins (saída) :"
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
                     Left            =   330
                     TabIndex        =   118
                     Top             =   810
                     Width           =   3015
                  End
                  Begin VB.Label Label26 
                     Alignment       =   2  'Center
                     Appearance      =   0  'Flat
                     AutoSize        =   -1  'True
                     BackColor       =   &H80000005&
                     BackStyle       =   0  'Transparent
                     Caption         =   "Valor total retenção PIS (saída) :"
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
                     Left            =   570
                     TabIndex        =   117
                     Top             =   450
                     Width           =   2775
                  End
                  Begin VB.Label Label30 
                     Alignment       =   2  'Center
                     Appearance      =   0  'Flat
                     AutoSize        =   -1  'True
                     BackColor       =   &H80000005&
                     BackStyle       =   0  'Transparent
                     Caption         =   "Valor total retenção Cofins (entrada) :"
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
                     Left            =   4545
                     TabIndex        =   116
                     Top             =   840
                     Width           =   3225
                  End
                  Begin VB.Label Label31 
                     Alignment       =   2  'Center
                     Appearance      =   0  'Flat
                     AutoSize        =   -1  'True
                     BackColor       =   &H80000005&
                     BackStyle       =   0  'Transparent
                     Caption         =   "Valor total retenção PIS (entrada) :"
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
                     Left            =   4755
                     TabIndex        =   115
                     Top             =   480
                     Width           =   3015
                  End
               End
               Begin VB.Frame Frame13 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "IRPJ"
                  Height          =   1545
                  Left            =   12630
                  TabIndex        =   86
                  Top             =   270
                  Width           =   2265
                  Begin VB.TextBox Text5 
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
                     Height          =   315
                     Left            =   1230
                     Locked          =   -1  'True
                     MaxLength       =   50
                     TabIndex        =   127
                     TabStop         =   0   'False
                     Text            =   "0,00"
                     ToolTipText     =   "Saldo de ICMS."
                     Top             =   1050
                     Width           =   915
                  End
                  Begin VB.TextBox txtValorIRPJ_prod_entrada 
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
                     Height          =   315
                     Left            =   1230
                     Locked          =   -1  'True
                     MaxLength       =   50
                     TabIndex        =   109
                     TabStop         =   0   'False
                     Text            =   "0,00"
                     ToolTipText     =   "Valor total do IRPJ entrada."
                     Top             =   720
                     Width           =   915
                  End
                  Begin VB.TextBox txtValorIRPJ_prod_saida 
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
                     Height          =   315
                     Left            =   1230
                     Locked          =   -1  'True
                     MaxLength       =   50
                     TabIndex        =   108
                     TabStop         =   0   'False
                     Text            =   "0,00"
                     ToolTipText     =   "Valor total do IRPJ saída."
                     Top             =   360
                     Width           =   915
                  End
                  Begin VB.Label Label22 
                     Alignment       =   2  'Center
                     Appearance      =   0  'Flat
                     AutoSize        =   -1  'True
                     BackColor       =   &H80000005&
                     BackStyle       =   0  'Transparent
                     Caption         =   "(entrada) :"
                     ForeColor       =   &H00000000&
                     Height          =   195
                     Left            =   240
                     TabIndex        =   104
                     Top             =   750
                     Width           =   825
                  End
                  Begin VB.Label Label1 
                     Alignment       =   2  'Center
                     Appearance      =   0  'Flat
                     AutoSize        =   -1  'True
                     BackColor       =   &H80000005&
                     BackStyle       =   0  'Transparent
                     Caption         =   "(saída) :"
                     ForeColor       =   &H00000000&
                     Height          =   195
                     Index           =   6
                     Left            =   450
                     TabIndex        =   103
                     Top             =   390
                     Width           =   615
                  End
                  Begin VB.Label Label21 
                     Alignment       =   2  'Center
                     Appearance      =   0  'Flat
                     AutoSize        =   -1  'True
                     BackColor       =   &H80000005&
                     BackStyle       =   0  'Transparent
                     Caption         =   "Saldo :"
                     ForeColor       =   &H00000000&
                     Height          =   195
                     Left            =   555
                     TabIndex        =   102
                     Top             =   1110
                     Width           =   525
                  End
               End
               Begin VB.Frame Frame12 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "CSLL"
                  Height          =   1545
                  Left            =   10116
                  TabIndex        =   82
                  Top             =   270
                  Width           =   2475
                  Begin VB.TextBox Text4 
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
                     Height          =   315
                     Left            =   1230
                     Locked          =   -1  'True
                     MaxLength       =   50
                     TabIndex        =   126
                     TabStop         =   0   'False
                     Text            =   "0,00"
                     ToolTipText     =   "Saldo de ICMS."
                     Top             =   1050
                     Width           =   915
                  End
                  Begin VB.TextBox txtValorCSLL_prod_entrada 
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
                     Height          =   315
                     Left            =   1230
                     Locked          =   -1  'True
                     MaxLength       =   50
                     TabIndex        =   107
                     TabStop         =   0   'False
                     Text            =   "0,00"
                     ToolTipText     =   "Valor total do CSLL entrada."
                     Top             =   690
                     Width           =   915
                  End
                  Begin VB.TextBox txtValorCSLL_prod_saida 
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
                     Height          =   315
                     Left            =   1230
                     Locked          =   -1  'True
                     MaxLength       =   50
                     TabIndex        =   106
                     TabStop         =   0   'False
                     Text            =   "0,00"
                     ToolTipText     =   "Valor total do CSLL saída."
                     Top             =   330
                     Width           =   915
                  End
                  Begin VB.Label Label14 
                     Alignment       =   2  'Center
                     Appearance      =   0  'Flat
                     AutoSize        =   -1  'True
                     BackColor       =   &H80000005&
                     BackStyle       =   0  'Transparent
                     Caption         =   "(entrada) :"
                     ForeColor       =   &H00000000&
                     Height          =   195
                     Left            =   240
                     TabIndex        =   98
                     Top             =   750
                     Width           =   825
                  End
                  Begin VB.Label Label1 
                     Alignment       =   2  'Center
                     Appearance      =   0  'Flat
                     AutoSize        =   -1  'True
                     BackColor       =   &H80000005&
                     BackStyle       =   0  'Transparent
                     Caption         =   "(saída) :"
                     ForeColor       =   &H00000000&
                     Height          =   195
                     Index           =   4
                     Left            =   450
                     TabIndex        =   97
                     Top             =   390
                     Width           =   615
                  End
                  Begin VB.Label Label12 
                     Alignment       =   2  'Center
                     Appearance      =   0  'Flat
                     AutoSize        =   -1  'True
                     BackColor       =   &H80000005&
                     BackStyle       =   0  'Transparent
                     Caption         =   "Saldo :"
                     ForeColor       =   &H00000000&
                     Height          =   195
                     Left            =   555
                     TabIndex        =   96
                     Top             =   1110
                     Width           =   525
                  End
               End
               Begin VB.Frame Frame11 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "COFINS"
                  Height          =   1545
                  Left            =   7602
                  TabIndex        =   81
                  Top             =   270
                  Width           =   2445
                  Begin VB.TextBox Text3 
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
                     Height          =   315
                     Left            =   1230
                     Locked          =   -1  'True
                     MaxLength       =   50
                     TabIndex        =   125
                     TabStop         =   0   'False
                     Text            =   "0,00"
                     ToolTipText     =   "Saldo de ICMS."
                     Top             =   1050
                     Width           =   915
                  End
                  Begin VB.TextBox txtValorCofins_prod_entrada 
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
                     Height          =   315
                     Left            =   1230
                     Locked          =   -1  'True
                     MaxLength       =   50
                     TabIndex        =   105
                     TabStop         =   0   'False
                     Text            =   "0,00"
                     ToolTipText     =   "Valor total do Cofins entrada."
                     Top             =   690
                     Width           =   915
                  End
                  Begin VB.TextBox txtValorCofins_prod_saida 
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
                     Height          =   315
                     Left            =   1230
                     Locked          =   -1  'True
                     MaxLength       =   50
                     TabIndex        =   95
                     TabStop         =   0   'False
                     Text            =   "0,00"
                     ToolTipText     =   "Valor total do Cofins saída."
                     Top             =   330
                     Width           =   915
                  End
                  Begin VB.Label Label20 
                     Alignment       =   2  'Center
                     Appearance      =   0  'Flat
                     AutoSize        =   -1  'True
                     BackColor       =   &H80000005&
                     BackStyle       =   0  'Transparent
                     Caption         =   "(entrada) :"
                     ForeColor       =   &H00000000&
                     Height          =   195
                     Left            =   240
                     TabIndex        =   101
                     Top             =   720
                     Width           =   825
                  End
                  Begin VB.Label Label1 
                     Alignment       =   2  'Center
                     Appearance      =   0  'Flat
                     AutoSize        =   -1  'True
                     BackColor       =   &H80000005&
                     BackStyle       =   0  'Transparent
                     Caption         =   "(saída) :"
                     ForeColor       =   &H00000000&
                     Height          =   195
                     Index           =   5
                     Left            =   450
                     TabIndex        =   100
                     Top             =   360
                     Width           =   615
                  End
                  Begin VB.Label Label15 
                     Alignment       =   2  'Center
                     Appearance      =   0  'Flat
                     AutoSize        =   -1  'True
                     BackColor       =   &H80000005&
                     BackStyle       =   0  'Transparent
                     Caption         =   "Saldo :"
                     ForeColor       =   &H00000000&
                     Height          =   195
                     Left            =   555
                     TabIndex        =   99
                     Top             =   1080
                     Width           =   525
                  End
               End
               Begin VB.Frame Frame9 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "PIS"
                  Height          =   1545
                  Left            =   5088
                  TabIndex        =   80
                  Top             =   270
                  Width           =   2475
                  Begin VB.TextBox Text2 
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
                     Height          =   315
                     Left            =   1230
                     Locked          =   -1  'True
                     MaxLength       =   50
                     TabIndex        =   124
                     TabStop         =   0   'False
                     Text            =   "0,00"
                     ToolTipText     =   "Saldo de ICMS."
                     Top             =   1080
                     Width           =   915
                  End
                  Begin VB.TextBox txtValorPIS_prod_entrada 
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
                     Height          =   315
                     Left            =   1230
                     Locked          =   -1  'True
                     MaxLength       =   50
                     TabIndex        =   94
                     TabStop         =   0   'False
                     Text            =   "0,00"
                     ToolTipText     =   "Valor total do PIS entrada."
                     Top             =   720
                     Width           =   915
                  End
                  Begin VB.TextBox txtValorPIS_prod_saida 
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
                     Height          =   315
                     Left            =   1230
                     Locked          =   -1  'True
                     MaxLength       =   50
                     TabIndex        =   90
                     TabStop         =   0   'False
                     Text            =   "0,00"
                     ToolTipText     =   "Valor total do PIS saída."
                     Top             =   360
                     Width           =   915
                  End
                  Begin VB.Label Label11 
                     Alignment       =   2  'Center
                     Appearance      =   0  'Flat
                     AutoSize        =   -1  'True
                     BackColor       =   &H80000005&
                     BackStyle       =   0  'Transparent
                     Caption         =   "(entrada) :"
                     ForeColor       =   &H00000000&
                     Height          =   195
                     Left            =   270
                     TabIndex        =   93
                     Top             =   720
                     Width           =   825
                  End
                  Begin VB.Label Label1 
                     Alignment       =   2  'Center
                     Appearance      =   0  'Flat
                     AutoSize        =   -1  'True
                     BackColor       =   &H80000005&
                     BackStyle       =   0  'Transparent
                     Caption         =   "(saída) :"
                     ForeColor       =   &H00000000&
                     Height          =   195
                     Index           =   3
                     Left            =   480
                     TabIndex        =   92
                     Top             =   390
                     Width           =   615
                  End
                  Begin VB.Label Label6 
                     Alignment       =   2  'Center
                     Appearance      =   0  'Flat
                     AutoSize        =   -1  'True
                     BackColor       =   &H80000005&
                     BackStyle       =   0  'Transparent
                     Caption         =   "Saldo :"
                     ForeColor       =   &H00000000&
                     Height          =   195
                     Left            =   585
                     TabIndex        =   91
                     Top             =   1080
                     Width           =   525
                  End
               End
               Begin VB.Frame Frame8 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "IPI"
                  Height          =   1545
                  Left            =   2574
                  TabIndex        =   79
                  Top             =   270
                  Width           =   2445
                  Begin VB.TextBox Text1 
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
                     Height          =   315
                     Left            =   1200
                     Locked          =   -1  'True
                     MaxLength       =   50
                     TabIndex        =   89
                     TabStop         =   0   'False
                     Text            =   "0,00"
                     ToolTipText     =   "Saldo de ICMS."
                     Top             =   1080
                     Width           =   915
                  End
                  Begin VB.TextBox txtValorIPI_entrada 
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
                     Height          =   315
                     Left            =   1200
                     Locked          =   -1  'True
                     MaxLength       =   50
                     TabIndex        =   88
                     TabStop         =   0   'False
                     Text            =   "0,00"
                     ToolTipText     =   "Valor total de IPI entrada."
                     Top             =   720
                     Width           =   915
                  End
                  Begin VB.TextBox txtValorIPI_saida 
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
                     Height          =   315
                     Left            =   1200
                     Locked          =   -1  'True
                     MaxLength       =   50
                     TabIndex        =   87
                     TabStop         =   0   'False
                     Text            =   "0,00"
                     ToolTipText     =   "Valor total de IPI saída."
                     Top             =   360
                     Width           =   915
                  End
                  Begin VB.Label Label10 
                     Alignment       =   2  'Center
                     Appearance      =   0  'Flat
                     AutoSize        =   -1  'True
                     BackColor       =   &H80000005&
                     BackStyle       =   0  'Transparent
                     Caption         =   "(entrada) :"
                     ForeColor       =   &H00000000&
                     Height          =   195
                     Left            =   240
                     TabIndex        =   85
                     Top             =   750
                     Width           =   825
                  End
                  Begin VB.Label Label1 
                     Alignment       =   2  'Center
                     Appearance      =   0  'Flat
                     AutoSize        =   -1  'True
                     BackColor       =   &H80000005&
                     BackStyle       =   0  'Transparent
                     Caption         =   "(saída) :"
                     ForeColor       =   &H00000000&
                     Height          =   195
                     Index           =   2
                     Left            =   450
                     TabIndex        =   84
                     Top             =   390
                     Width           =   615
                  End
                  Begin VB.Label Label9 
                     Alignment       =   2  'Center
                     Appearance      =   0  'Flat
                     AutoSize        =   -1  'True
                     BackColor       =   &H80000005&
                     BackStyle       =   0  'Transparent
                     Caption         =   "Saldo :"
                     ForeColor       =   &H00000000&
                     Height          =   195
                     Left            =   555
                     TabIndex        =   83
                     Top             =   1110
                     Width           =   525
                  End
               End
               Begin VB.Frame Frame6 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "ICMS"
                  Height          =   1545
                  Left            =   60
                  TabIndex        =   72
                  Top             =   270
                  Width           =   2475
                  Begin VB.TextBox txtValorICMS_Saída 
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
                     Height          =   315
                     Left            =   1215
                     Locked          =   -1  'True
                     MaxLength       =   50
                     TabIndex        =   75
                     TabStop         =   0   'False
                     Text            =   "0,00"
                     ToolTipText     =   "Valor total de ICMS saída."
                     Top             =   360
                     Width           =   915
                  End
                  Begin VB.TextBox txtValorICMS_Entrada 
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
                     Height          =   315
                     Left            =   1215
                     Locked          =   -1  'True
                     MaxLength       =   50
                     TabIndex        =   74
                     TabStop         =   0   'False
                     Text            =   "0,00"
                     ToolTipText     =   "Valor total de ICMS entrada."
                     Top             =   720
                     Width           =   915
                  End
                  Begin VB.TextBox txtSaldo_ICMS 
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
                     Height          =   315
                     Left            =   1215
                     Locked          =   -1  'True
                     MaxLength       =   50
                     TabIndex        =   73
                     TabStop         =   0   'False
                     Text            =   "0,00"
                     ToolTipText     =   "Saldo de ICMS."
                     Top             =   1050
                     Width           =   915
                  End
                  Begin VB.Label Label19 
                     Alignment       =   2  'Center
                     Appearance      =   0  'Flat
                     AutoSize        =   -1  'True
                     BackColor       =   &H80000005&
                     BackStyle       =   0  'Transparent
                     Caption         =   "(entrada) :"
                     ForeColor       =   &H00000000&
                     Height          =   195
                     Left            =   255
                     TabIndex        =   78
                     Top             =   720
                     Width           =   825
                  End
                  Begin VB.Label Label1 
                     Alignment       =   2  'Center
                     Appearance      =   0  'Flat
                     AutoSize        =   -1  'True
                     BackColor       =   &H80000005&
                     BackStyle       =   0  'Transparent
                     Caption         =   "(saída) :"
                     ForeColor       =   &H00000000&
                     Height          =   195
                     Index           =   0
                     Left            =   465
                     TabIndex        =   77
                     Top             =   360
                     Width           =   615
                  End
                  Begin VB.Label Label2 
                     Alignment       =   2  'Center
                     Appearance      =   0  'Flat
                     AutoSize        =   -1  'True
                     BackColor       =   &H80000005&
                     BackStyle       =   0  'Transparent
                     Caption         =   "Saldo :"
                     ForeColor       =   &H00000000&
                     Height          =   195
                     Left            =   570
                     TabIndex        =   76
                     Top             =   1080
                     Width           =   525
                  End
               End
            End
         End
      End
      Begin MSComctlLib.ListView ListaNF 
         Height          =   6960
         Left            =   -74910
         TabIndex        =   5
         Top             =   2100
         Width           =   15225
         _ExtentX        =   26855
         _ExtentY        =   12277
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
         NumItems        =   31
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
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "N"
            Text            =   "Nota fiscal"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Object.Tag             =   "T"
            Text            =   "Tipo"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "Destinatário"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Object.Tag             =   "N"
            Text            =   "Vlr. total produtos"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Object.Tag             =   "N"
            Text            =   "Base de cálculo ICMS"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Object.Tag             =   "N"
            Text            =   "Vlr. ICMS"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Object.Tag             =   "N"
            Text            =   "Base de calculo ICMS subst."
            Object.Width           =   3969
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   9
            Object.Tag             =   "N"
            Text            =   "Vlr. ICMS subst."
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   10
            Object.Tag             =   "N"
            Text            =   "Vlr. total IPI"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   11
            Object.Tag             =   "N"
            Text            =   "Vlr. total PIS"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   12
            Object.Tag             =   "N"
            Text            =   "Vlr. total Cofins"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   13
            Object.Tag             =   "N"
            Text            =   "Vlr. total CSLL"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   14
            Object.Tag             =   "N"
            Text            =   "Vlr. total IRPJ"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   15
            Object.Tag             =   "N"
            Text            =   "Vlr. total ret. PIS"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   16
            Object.Tag             =   "N"
            Text            =   "Vlr. total ret. Cofins"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   17
            Object.Tag             =   "N"
            Text            =   "Vlr. total serviços"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   18
            Object.Tag             =   "N"
            Text            =   "Vlr. total PIS"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   19
            Object.Tag             =   "N"
            Text            =   "Vlr. total Cofins"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   20
            Object.Tag             =   "N"
            Text            =   "Vlr. total CSLL"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   21
            Object.Tag             =   "N"
            Text            =   "Vlr. total ISSQN"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   22
            Object.Tag             =   "N"
            Text            =   "Vlr. total INSS"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   23
            Object.Tag             =   "N"
            Text            =   "Vlr. total IRPJ"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   24
            Object.Tag             =   "N"
            Text            =   "Vlr. total IRRF"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   25
            Object.Tag             =   "N"
            Text            =   "Vlr. frete"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   26
            Object.Tag             =   "N"
            Text            =   "Vlr. seguro"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   27
            Object.Tag             =   "N"
            Text            =   "Outras desp. acessórias"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   28
            Object.Tag             =   "N"
            Text            =   "Vlr. total DAS"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(30) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   29
            Object.Tag             =   "N"
            Text            =   "Vlr. total nota"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(31) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   30
            Object.Tag             =   "N"
            Text            =   "Vlr. total duplicata"
            Object.Width           =   2646
         EndProperty
      End
      Begin MSComctlLib.ListView ListaNF1 
         Height          =   6960
         Left            =   -74910
         TabIndex        =   13
         Top             =   2100
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   12277
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
         NumItems        =   31
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "N"
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Object.Tag             =   "D"
            Text            =   "Dt. entrada"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "N"
            Text            =   "Nota fiscal"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Object.Tag             =   "T"
            Text            =   "Tipo"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "Emitente"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Object.Tag             =   "N"
            Text            =   "Vlr. total produtos"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Object.Tag             =   "N"
            Text            =   "Base de cálculo ICMS"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Object.Tag             =   "N"
            Text            =   "Vlr. ICMS"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Object.Tag             =   "N"
            Text            =   "Base de calculo ICMS subst."
            Object.Width           =   3969
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   9
            Object.Tag             =   "N"
            Text            =   "Vlr. ICMS subst."
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   10
            Object.Tag             =   "N"
            Text            =   "Vlr. total IPI"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   11
            Object.Tag             =   "N"
            Text            =   "Vlr. total PIS"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   12
            Object.Tag             =   "N"
            Text            =   "Vlr. total Cofins"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   13
            Object.Tag             =   "N"
            Text            =   "Vlr. total CSLL"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   14
            Object.Tag             =   "N"
            Text            =   "Vlr. total IRPJ"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   15
            Object.Tag             =   "N"
            Text            =   "Vlr. total ret. PIS"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   16
            Object.Tag             =   "N"
            Text            =   "Vlr. total ret. Cofins"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   17
            Object.Tag             =   "N"
            Text            =   "Vlr. total serviços"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   18
            Object.Tag             =   "N"
            Text            =   "Vlr. total PIS"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   19
            Object.Tag             =   "N"
            Text            =   "Vlr. total Cofins"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   20
            Object.Tag             =   "N"
            Text            =   "Vlr. total CSLL"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   21
            Object.Tag             =   "N"
            Text            =   "Vlr. total ISSQN"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   22
            Object.Tag             =   "N"
            Text            =   "Vlr. total INSS"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   23
            Object.Tag             =   "N"
            Text            =   "Vlr. total IRPJ"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   24
            Object.Tag             =   "N"
            Text            =   "Vlr. total IRRF"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   25
            Object.Tag             =   "N"
            Text            =   "Vlr. frete"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   26
            Object.Tag             =   "N"
            Text            =   "Vlr. seguro"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   27
            Object.Tag             =   "N"
            Text            =   "Outras desp. acessórias"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   28
            Object.Tag             =   "N"
            Text            =   "Vlr. total DAS"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(30) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   29
            Object.Tag             =   "N"
            Text            =   "Vlr. total nota"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(31) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   30
            Object.Tag             =   "N"
            Text            =   "Vlr. total duplicata"
            Object.Width           =   2646
         EndProperty
      End
   End
End
Attribute VB_Name = "frmFaturamento_Relatorios_Impostos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StrSql_Faturamento_Relatorios_Impostos As String 'OK
Dim StrSql_Faturamento_Relatorios_Impostos1 As String 'OK
Dim FormulaRel_Faturamento_Relatorios_Impostos As String 'OK
Dim TBLISTA_Faturamento_Impostos As ADODB.Recordset 'OK
Dim TBLISTA_Faturamento_Impostos1 As ADODB.Recordset 'OK

Sub ProcAjuda()
On Error GoTo tratar_erro

FunAbrirVideoWeb ("http://www.youtube.com/watch?v=gaGLvvro5T4&list=UUODGiDjQ-BCrxh0YXfJvoqg&index=13&feature=plcp")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

M = FunVerificaMes(TabFiltro.SelectedItem.key)
If OptDomes.Value = True Then
    StrSql_Faturamento_Relatorios_Impostos = "Select * FROM tbl_Dados_Nota_Fiscal where DtValidacao IS NOT NULL and int_status = 1 and month(dt_DataEmissao)= '" & M & "' and Year(dt_DataEmissao) = '" & cmbAno.Text & "' and int_TipoNota = 1 ORDER BY dt_Saida_Entrada, dt_DataEmissao, int_notafiscal"
    StrSql_Faturamento_Relatorios_Impostos1 = "Select * FROM tbl_Dados_Nota_Fiscal where DtValidacao IS NOT NULL and int_status = 1 and month(dt_Saida_Entrada)= '" & M & "' and Year(dt_Saida_Entrada) = '" & cmbAno.Text & "' and int_TipoNota = 2 ORDER BY dt_Saida_Entrada, dt_DataEmissao, int_notafiscal"
    FormulaRel_Faturamento_Relatorios_Impostos = "{tbl_Dados_Nota_Fiscal.RespValidacao} <> 'Null' and {tbl_Dados_Nota_Fiscal.int_status} = 1 and Month ({tbl_Dados_Nota_Fiscal.dt_DataEmissao}) = " & M & " and Year ({tbl_Dados_Nota_Fiscal.dt_DataEmissao}) = " & cmbAno & " and {tbl_Dados_Nota_Fiscal.int_TipoNota} = 1 or {tbl_Dados_Nota_Fiscal.int_status} = 1 and Month ({tbl_Dados_Nota_Fiscal.dt_Saida_Entrada}) = " & M & " and Year ({tbl_Dados_Nota_Fiscal.dt_Saida_Entrada}) = " & cmbAno & " and {tbl_Dados_Nota_Fiscal.int_TipoNota} = 2"
ElseIf OptAteomes.Value = True Then
        StrSql_Faturamento_Relatorios_Impostos = "Select * FROM tbl_Dados_Nota_Fiscal where DtValidacao IS NOT NULL and int_status = 1 and month(dt_DataEmissao)<= '" & M & "' and Year(dt_DataEmissao) = '" & cmbAno.Text & "' and int_TipoNota = 1 ORDER BY dt_Saida_Entrada, dt_DataEmissao, int_notafiscal"
        StrSql_Faturamento_Relatorios_Impostos1 = "Select * FROM tbl_Dados_Nota_Fiscal where DtValidacao IS NOT NULL and int_status = 1 and month(dt_Saida_Entrada)<= '" & M & "' and Year(dt_Saida_Entrada) = '" & cmbAno.Text & "' and int_TipoNota = 2 ORDER BY dt_Saida_Entrada, dt_DataEmissao, int_notafiscal"
        FormulaRel_Faturamento_Relatorios_Impostos = "{tbl_Dados_Nota_Fiscal.RespValidacao} <> 'Null' and {tbl_Dados_Nota_Fiscal.int_status} = 1 and Month ({tbl_Dados_Nota_Fiscal.dt_DataEmissao}) <= " & M & " and Year ({tbl_Dados_Nota_Fiscal.dt_DataEmissao}) = " & cmbAno & " and {tbl_Dados_Nota_Fiscal.int_TipoNota} = 1 or {tbl_Dados_Nota_Fiscal.int_status} = 1 and Month ({tbl_Dados_Nota_Fiscal.dt_Saida_Entrada}) <= " & M & " and Year ({tbl_Dados_Nota_Fiscal.dt_Saida_Entrada}) = " & cmbAno & " and {tbl_Dados_Nota_Fiscal.int_TipoNota} = 2"
    Else
        StrSql_Faturamento_Relatorios_Impostos = "Select * FROM tbl_Dados_Nota_Fiscal where DtValidacao IS NOT NULL and int_status = 1 and (dt_DataEmissao) Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "' and int_TipoNota = 1 ORDER BY dt_Saida_Entrada, dt_DataEmissao, int_notafiscal"
        StrSql_Faturamento_Relatorios_Impostos1 = "Select * FROM tbl_Dados_Nota_Fiscal where DtValidacao IS NOT NULL and int_status = 1 and (dt_Saida_Entrada) Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "' and int_TipoNota = 2 ORDER BY dt_Saida_Entrada, dt_DataEmissao, int_notafiscal"
        FormulaRel_Faturamento_Relatorios_Impostos = "{tbl_Dados_Nota_Fiscal.RespValidacao} <> 'Null' and {tbl_Dados_Nota_Fiscal.int_status} = 1 and {tbl_Dados_Nota_Fiscal.dt_DataEmissao} >= Date(" & Year(msk_fltInicio.Value) & "," & Month(msk_fltInicio.Value) & "," & Day(msk_fltInicio.Value) & ") and {tbl_Dados_Nota_Fiscal.dt_DataEmissao} <= Date(" & Year(msk_fltFim.Value) & "," & Month(msk_fltFim.Value) & "," & Day(msk_fltFim.Value) & ") and {tbl_Dados_Nota_Fiscal.int_TipoNota} = 1 or {tbl_Dados_Nota_Fiscal.int_status} = 1 and {tbl_Dados_Nota_Fiscal.dt_Saida_Entrada} >= Date(" & Year(msk_fltInicio.Value) & "," & Month(msk_fltInicio.Value) & "," & Day(msk_fltInicio.Value) & ") and {tbl_Dados_Nota_Fiscal.dt_Saida_Entrada} <= Date(" & _
                                Year(msk_fltFim.Value) & "," & Month(msk_fltFim.Value) & "," & Day(msk_fltFim.Value) & ") and {tbl_Dados_Nota_Fiscal.int_TipoNota} = 2"
End If
ListaNF.ListItems.Clear
ListaNF1.ListItems.Clear
ProcCarregaListaNF
ProcCarregaListaNF1
ProcSomaImpostos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbAno_Click()
On Error GoTo tratar_erro

ListaNF.ListItems.Clear
ListaNF1.ListItems.Clear
ProcLimparTotais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro

NomeRel = "Faturamento_impostos.rpt"
ProcImprimirRel FormulaRel_Faturamento_Relatorios_Impostos, ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaListaNF()
On Error GoTo tratar_erro

lblRegistros.Caption = "Nº de registros: 0"
lblPaginas.Caption = "Página: 0 de: 0"
ListaNF.ListItems.Clear
Set TBLISTA_Faturamento_Impostos = CreateObject("adodb.recordset")
TBLISTA_Faturamento_Impostos.Open StrSql_Faturamento_Relatorios_Impostos, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA_Faturamento_Impostos.EOF = False Then ProcExibePagina (1)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaListaNF1()
On Error GoTo tratar_erro

lblRegistros1.Caption = "Nº de registros: 0"
lblPaginas1.Caption = "Página: 0 de: 0"
ListaNF1.ListItems.Clear
Set TBLISTA_Faturamento_Impostos1 = CreateObject("adodb.recordset")
TBLISTA_Faturamento_Impostos1.Open StrSql_Faturamento_Relatorios_Impostos1, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA_Faturamento_Impostos1.EOF = False Then ProcExibePagina1 (1)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Faturamento_Impostos.AbsolutePage <> 2 Then
    If TBLISTA_Faturamento_Impostos.AbsolutePage = -3 Then
        ProcExibePagina (TBLISTA_Faturamento_Impostos.PageCount - 1)
    Else
        TBLISTA_Faturamento_Impostos.AbsolutePage = TBLISTA_Faturamento_Impostos.AbsolutePage - 2
        ProcExibePagina (TBLISTA_Faturamento_Impostos.AbsolutePage)
    End If
Else
    ProcExibePagina (1)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt1_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas1.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Faturamento_Impostos1.AbsolutePage <> 2 Then
    If TBLISTA_Faturamento_Impostos1.AbsolutePage = -3 Then
        ProcExibePagina1 (TBLISTA_Faturamento_Impostos1.PageCount - 1)
    Else
        TBLISTA_Faturamento_Impostos1.AbsolutePage = TBLISTA_Faturamento_Impostos1.AbsolutePage - 2
        ProcExibePagina1 (TBLISTA_Faturamento_Impostos1.AbsolutePage)
    End If
Else
    ProcExibePagina1 (1)
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
    TBLISTA_Faturamento_Impostos.AbsolutePage = txtPagIr.Text
    ProcExibePagina (TBLISTA_Faturamento_Impostos.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagIr1_Click()
On Error GoTo tratar_erro

If txtPagIr1 = "" Then Exit Sub
Quant = ReturnNumbersOnly(Right(lblPaginas1.Caption, 4))
If Quant <= 1 Or txtPagIr1 > Quant Then Exit Sub
If txtPagIr1.Text >= 1 And txtPagIr1.Text <= Quant Then
    TBLISTA_Faturamento_Impostos1.AbsolutePage = txtPagIr1.Text
    ProcExibePagina1 (TBLISTA_Faturamento_Impostos1.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Faturamento_Impostos.AbsolutePage = 1
ProcExibePagina (TBLISTA_Faturamento_Impostos.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim1_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas1.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Faturamento_Impostos1.AbsolutePage = 1
ProcExibePagina1 (TBLISTA_Faturamento_Impostos1.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Faturamento_Impostos.AbsolutePage <> -3 Then
    If TBLISTA_Faturamento_Impostos.AbsolutePage = 1 Then
        ProcExibePagina (2)
    Else
        ProcExibePagina (TBLISTA_Faturamento_Impostos.AbsolutePage)
    End If
Else
    ProcExibePagina (TBLISTA_Faturamento_Impostos.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx1_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas1.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Faturamento_Impostos1.AbsolutePage <> -3 Then
    If TBLISTA_Faturamento_Impostos1.AbsolutePage = 1 Then
        ProcExibePagina1 (2)
    Else
        ProcExibePagina1 (TBLISTA_Faturamento_Impostos1.AbsolutePage)
    End If
Else
    ProcExibePagina1 (TBLISTA_Faturamento_Impostos1.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Faturamento_Impostos.AbsolutePage = TBLISTA_Faturamento_Impostos.PageCount
ProcExibePagina (TBLISTA_Faturamento_Impostos.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt1_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas1.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Faturamento_Impostos1.AbsolutePage = TBLISTA_Faturamento_Impostos1.PageCount
ProcExibePagina1 (TBLISTA_Faturamento_Impostos1.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyF2: ProcFiltrar
    Case vbKeyF5: ProcImprimir
    Case vbKeyF1: ProcAjuda
    Case vbKeyEscape: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15195, 6, True
ProcLimpaVariaveisPrincipais
SSTab1.Tab = 0
If PBLista.Value = 0 Then PBLista.Value = 100
ProcCarregaComboAno cmbAno, "2005", 1
TabFiltro.Tabs(Month(Date)).Selected = True
msk_fltInicio.Value = Date
msk_fltFim.Value = Date

ProcRemoveObjetosResize Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaNF_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView ListaNF, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaNF1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView ListaNF1, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub msk_fltFim_Change()
On Error GoTo tratar_erro

ListaNF.ListItems.Clear
ListaNF1.ListItems.Clear
ProcLimparTotais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub msk_fltInicio_Change()
On Error GoTo tratar_erro

ListaNF.ListItems.Clear
ListaNF1.ListItems.Clear
ProcLimparTotais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt_periodo_Click()
On Error GoTo tratar_erro

ListaNF.ListItems.Clear
ListaNF1.ListItems.Clear
ProcLimparTotais
If Opt_periodo.Value = True Then
    Frame5.Enabled = True
    msk_fltInicio.SetFocus
Else
    Frame5.Enabled = False
    msk_fltInicio.Value = Date
    msk_fltFim.Value = Date
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub OptAteomes_Click()
On Error GoTo tratar_erro

ListaNF.ListItems.Clear
ListaNF1.ListItems.Clear
ProcLimparTotais
If OptAteomes.Value = True Then
    Frame5.Enabled = False
    msk_fltInicio.Value = Date
    msk_fltFim.Value = Date
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub OptDomes_Click()
On Error GoTo tratar_erro

ListaNF.ListItems.Clear
ListaNF1.ListItems.Clear
ProcLimparTotais
If OptDomes.Value = True Then
    Frame5.Enabled = False
    msk_fltInicio.Value = Date
    msk_fltFim.Value = Date
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimparTotais()
On Error GoTo tratar_erro

'Produtos saída
txtValorICMS_Saída = "0,00"
txtValorICMS_Entrada = "0,00"
txtSaldo_ICMS = "0,00"
txtValorIPI_saida = "0,00"
txtValorPIS_prod_saida = "0,00"
txtValorCofins_prod_saida = "0,00"
txtValorCSLL_prod_saida = "0,00"
txtValorIRPJ_prod_saida = "0,00"
txtValor_retencao_PIS_prod_saida = "0,00"
txtValor_retencao_Cofins_prod_saida = "0,00"
'Produtos entrada
txtValorIPI_entrada = "0,00"
txtValorPIS_prod_entrada = "0,00"
txtValorCofins_prod_entrada = "0,00"
txtValorCSLL_prod_entrada = "0,00"
txtValorIRPJ_prod_entrada = "0,00"
txtValor_retencao_PIS_prod_entrada = "0,00"
txtValor_retencao_Cofins_prod_entrada = "0,00"
'Serviços saída
txtValorPIS_serv_saida = "0,00"
txtValorCofins_serv_saida = "0,00"
txtValorCSLL_serv_saida = "0,00"
txtValorISS_serv_saida = "0,00"
txtValorINSS_serv_saida = "0,00"
txtValorIRPJ_serv_saida = "0,00"
txtValorIRRF_serv_saida = "0,00"
'Serviços entrada
txtValorPIS_serv_entrada = "0,00"
txtValorCofins_serv_entrada = "0,00"
txtValorCSLL_serv_entrada = "0,00"
txtValorISS_serv_entrada = "0,00"
txtValorINSS_serv_entrada = "0,00"
txtValorIRPJ_serv_entrada = "0,00"
txtValorIRRF_serv_entrada = "0,00"
'Faturamento saída
txtValorDAS_saida = "0,00"
'Faturamento entrada
txtValorDAS_entrada = "0,00"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error GoTo tratar_erro

Select Case SSTab1.Tab
    Case 0: Frame21.Visible = True
    Case 1: Frame21.Visible = True
    Case 2:
        Frame21.Visible = False
        SSTab2.Tab = 0
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub TabFiltro_Click()
On Error GoTo tratar_erro

ListaNF.ListItems.Clear
ListaNF1.ListItems.Clear
ProcLimparTotais

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

Private Sub txtNreg1_Change()
On Error GoTo tratar_erro

If txtNreg1 <> "" Then
    VerifNumero = txtNreg1
    ProcVerificaNumero
    If VerifNumero = False Then
        txtNreg1 = ""
        txtNreg1.SetFocus
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

Private Sub txtPagIr1_Change()
On Error GoTo tratar_erro

If txtPagIr1 <> "" Then
    VerifNumero = txtPagIr1
    ProcVerificaNumero
    If VerifNumero = False Then
        txtPagIr1 = ""
        txtPagIr1.SetFocus
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
    Case 1: ProcFiltrar
    Case 2: ProcImprimir
    Case 4: ProcAjuda
    Case 5: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina(Pagina)
On Error GoTo tratar_erro

ListaNF.ListItems.Clear
TBLISTA_Faturamento_Impostos.PageSize = IIf(txtNreg = "", 30, txtNreg)
TBLISTA_Faturamento_Impostos.AbsolutePage = Pagina
TamanhoPagina = TBLISTA_Faturamento_Impostos.PageSize
ContadorReg = 1

PBLista.Min = 0
PBLista.Max = IIf(IIf(TBLISTA_Faturamento_Impostos.RecordCount - TBLISTA_Faturamento_Impostos.AbsolutePosition <= 0, 1, TBLISTA_Faturamento_Impostos.RecordCount - TBLISTA_Faturamento_Impostos.AbsolutePosition) < TBLISTA_Faturamento_Impostos.PageSize, IIf(TBLISTA_Faturamento_Impostos.RecordCount - TBLISTA_Faturamento_Impostos.AbsolutePosition <= 0, 1, TBLISTA_Faturamento_Impostos.RecordCount - TBLISTA_Faturamento_Impostos.AbsolutePosition), TBLISTA_Faturamento_Impostos.PageSize)
PBLista.Value = 1
Contador = 0
Do While TBLISTA_Faturamento_Impostos.EOF = False And (ContadorReg <= TamanhoPagina)
    If TBLISTA_Faturamento_Impostos!int_TipoNota = 1 Then
        With ListaNF.ListItems
            .Add , , TBLISTA_Faturamento_Impostos!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA_Faturamento_Impostos!dt_DataEmissao), "", Format(TBLISTA_Faturamento_Impostos!dt_DataEmissao, "dd/mm/yy"))
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA_Faturamento_Impostos!int_NotaFiscal), "", TBLISTA_Faturamento_Impostos!int_NotaFiscal)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA_Faturamento_Impostos!TipoNF), "", TBLISTA_Faturamento_Impostos!TipoNF)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA_Faturamento_Impostos!txt_Razao_Nome), "", Trim(TBLISTA_Faturamento_Impostos!txt_Razao_Nome))
            
            Set TBTotaisnota = CreateObject("adodb.recordset")
            TBTotaisnota.Open "Select * from tbl_Totais_Nota where id_nota = " & TBLISTA_Faturamento_Impostos!ID, Conexao, adOpenKeyset, adLockOptimistic
            If TBTotaisnota.EOF = False Then
                'Produtos
                Quant = IIf(IsNull(TBTotaisnota!Qtde_total_prod), 0, TBTotaisnota!Qtde_total_prod)
                .Item(.Count).SubItems(5) = IIf(IsNull(TBTotaisnota!dbl_Valor_Total_Produtos), "", Format(TBTotaisnota!dbl_Valor_Total_Produtos, "###,##0.00"))
                
                If IsNull(TBTotaisnota!Valor_total_ICMS_SN) = False And TBTotaisnota!Valor_total_ICMS_SN > 0 Then
                    .Item(.Count).SubItems(6) = IIf(IsNull(TBTotaisnota!dbl_Valor_Total_Nota), "", Format(TBTotaisnota!dbl_Valor_Total_Nota, "###,##0.00"))
                    .Item(.Count).SubItems(7) = IIf(IsNull(TBTotaisnota!Valor_total_ICMS_SN), "", Format(TBTotaisnota!Valor_total_ICMS_SN, "###,##0.00"))
                Else
                    .Item(.Count).SubItems(6) = IIf(IsNull(TBTotaisnota!dbl_Base_ICMS), "", Format(TBTotaisnota!dbl_Base_ICMS, "###,##0.00"))
                    .Item(.Count).SubItems(7) = IIf(IsNull(TBTotaisnota!dbl_Valor_ICMS), "", Format(TBTotaisnota!dbl_Valor_ICMS, "###,##0.00"))
                End If
                
                .Item(.Count).SubItems(8) = IIf(IsNull(TBTotaisnota!dbl_Base_ICMS_Subst), "", Format(TBTotaisnota!dbl_Base_ICMS_Subst, "###,##0.00"))
                .Item(.Count).SubItems(9) = IIf(IsNull(TBTotaisnota!dbl_Valor_ICMS_Subst), "", Format(TBTotaisnota!dbl_Valor_ICMS_Subst, "###,##0.00"))
                .Item(.Count).SubItems(10) = IIf(IsNull(TBTotaisnota!dbl_Valor_Total_IPI), "", Format(TBTotaisnota!dbl_Valor_Total_IPI, "###,##0.00"))
                .Item(.Count).SubItems(11) = IIf(IsNull(TBTotaisnota!Total_PIS_prod), "", Format(TBTotaisnota!Total_PIS_prod, "###,##0.00"))
                .Item(.Count).SubItems(12) = IIf(IsNull(TBTotaisnota!Total_Cofins_prod), "", Format(TBTotaisnota!Total_Cofins_prod, "###,##0.00"))
                .Item(.Count).SubItems(13) = IIf(IsNull(TBTotaisnota!Total_CSLL_prod), "", Format(TBTotaisnota!Total_CSLL_prod, "###,##0.00"))
                .Item(.Count).SubItems(14) = IIf(IsNull(TBTotaisnota!Total_IRPJ_prod), "", Format(TBTotaisnota!Total_IRPJ_prod, "###,##0.00"))
                .Item(.Count).SubItems(15) = IIf(IsNull(TBTotaisnota!Total_retencao_PIS), "", Format(TBTotaisnota!Total_retencao_PIS, "###,##0.00"))
                .Item(.Count).SubItems(16) = IIf(IsNull(TBTotaisnota!Total_retencao_Cofins), "", Format(TBTotaisnota!Total_retencao_Cofins, "###,##0.00"))
                'Serviços
                quantidade = IIf(IsNull(TBTotaisnota!Qtde_total_serv), 0, TBTotaisnota!Qtde_total_serv)
                .Item(.Count).SubItems(17) = IIf(IsNull(TBTotaisnota!dbl_Valor_Total_Nota_Serv), "", Format(TBTotaisnota!dbl_Valor_Total_Nota_Serv, "###,##0.00"))
                .Item(.Count).SubItems(18) = IIf(IsNull(TBTotaisnota!Total_PIS_serv), "", Format(TBTotaisnota!Total_PIS_serv, "###,##0.00"))
                .Item(.Count).SubItems(19) = IIf(IsNull(TBTotaisnota!Total_Cofins_serv), "", Format(TBTotaisnota!Total_Cofins_serv, "###,##0.00"))
                .Item(.Count).SubItems(20) = IIf(IsNull(TBTotaisnota!Total_CSLL_serv), "", Format(TBTotaisnota!Total_CSLL_serv, "###,##0.00"))
                .Item(.Count).SubItems(21) = IIf(IsNull(TBTotaisnota!dbl_valor_total_iss), "", Format(TBTotaisnota!dbl_valor_total_iss, "###,##0.00"))
                .Item(.Count).SubItems(22) = IIf(IsNull(TBTotaisnota!Total_INSS_serv), "", Format(TBTotaisnota!Total_INSS_serv, "###,##0.00"))
                .Item(.Count).SubItems(23) = IIf(IsNull(TBTotaisnota!Total_IRPJ_serv), "", Format(TBTotaisnota!Total_IRPJ_serv, "###,##0.00"))
                .Item(.Count).SubItems(24) = IIf(IsNull(TBTotaisnota!Total_IRRF_serv), "", Format(TBTotaisnota!Total_IRRF_serv, "###,##0.00"))
                'Outros
                .Item(.Count).SubItems(25) = IIf(IsNull(TBTotaisnota!dbl_Valor_Frete), "", Format(TBTotaisnota!dbl_Valor_Frete, "###,##0.00"))
                .Item(.Count).SubItems(26) = IIf(IsNull(TBTotaisnota!dbl_Valor_Seguro), "", Format(TBTotaisnota!dbl_Valor_Seguro, "###,##0.00"))
                .Item(.Count).SubItems(27) = IIf(IsNull(TBTotaisnota!dbl_Desp_Adicionais), "", Format(TBTotaisnota!dbl_Desp_Adicionais, "###,##0.00"))
                .Item(.Count).SubItems(28) = IIf(IsNull(TBTotaisnota!Total_DAS), "", Format(TBTotaisnota!Total_DAS, "###,##0.00"))
                'Total
                Valor1 = IIf(IsNull(TBTotaisnota!dbl_Valor_Total_Nota), 0, TBTotaisnota!dbl_Valor_Total_Nota)
                Valor2 = IIf(IsNull(TBTotaisnota!Total_retencao_PIS), 0, TBTotaisnota!Total_retencao_PIS)
                Valor3 = IIf(IsNull(TBTotaisnota!Total_retencao_Cofins), 0, TBTotaisnota!Total_retencao_Cofins)
                .Item(.Count).SubItems(29) = Format(Valor1 - Valor2 - Valor3, "###,##0.00")
                .Item(.Count).SubItems(30) = IIf(IsNull(TBTotaisnota!Valor_total_receber_pagar), "0,00", Format(TBTotaisnota!Valor_total_receber_pagar, "###,##0.00"))
            End If
            TBTotaisnota.Close
        End With
    End If
    TBLISTA_Faturamento_Impostos.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista.Value = Contador
Loop
lblRegistros.Caption = "Nº de registros: " & TBLISTA_Faturamento_Impostos.RecordCount
If TBLISTA_Faturamento_Impostos.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "Página: 1 de: " & TBLISTA_Faturamento_Impostos.PageCount
ElseIf TBLISTA_Faturamento_Impostos.AbsolutePage = adPosEOF Then
        lblPaginas.Caption = "Página: " & TBLISTA_Faturamento_Impostos.PageCount & " de: " & TBLISTA_Faturamento_Impostos.PageCount
    Else
        lblPaginas.Caption = "Página: " & TBLISTA_Faturamento_Impostos.AbsolutePage - 1 & " de: " & TBLISTA_Faturamento_Impostos.PageCount
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina1(Pagina)
On Error GoTo tratar_erro

ListaNF1.ListItems.Clear
TBLISTA_Faturamento_Impostos1.PageSize = IIf(txtNreg1 = "", 30, txtNreg1)
TBLISTA_Faturamento_Impostos1.AbsolutePage = Pagina
TamanhoPagina = TBLISTA_Faturamento_Impostos1.PageSize
ContadorReg = 1

PBLista.Min = 0
PBLista.Max = IIf(IIf(TBLISTA_Faturamento_Impostos1.RecordCount - TBLISTA_Faturamento_Impostos1.AbsolutePosition <= 0, 1, TBLISTA_Faturamento_Impostos1.RecordCount - TBLISTA_Faturamento_Impostos1.AbsolutePosition) < TBLISTA_Faturamento_Impostos1.PageSize, IIf(TBLISTA_Faturamento_Impostos1.RecordCount - TBLISTA_Faturamento_Impostos1.AbsolutePosition <= 0, 1, TBLISTA_Faturamento_Impostos1.RecordCount - TBLISTA_Faturamento_Impostos1.AbsolutePosition), TBLISTA_Faturamento_Impostos1.PageSize)
PBLista.Value = 1
Contador = 0
Do While TBLISTA_Faturamento_Impostos1.EOF = False And (ContadorReg <= TamanhoPagina)
    If TBLISTA_Faturamento_Impostos1!int_TipoNota = 2 Then
        With ListaNF1.ListItems
            .Add , , TBLISTA_Faturamento_Impostos1!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA_Faturamento_Impostos1!dt_Saida_Entrada), "", Format(TBLISTA_Faturamento_Impostos1!dt_Saida_Entrada, "dd/mm/yy"))
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA_Faturamento_Impostos1!int_NotaFiscal), "", TBLISTA_Faturamento_Impostos1!int_NotaFiscal)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA_Faturamento_Impostos1!TipoNF), "", TBLISTA_Faturamento_Impostos1!TipoNF)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA_Faturamento_Impostos1!txt_Razao_Nome), "", Trim(TBLISTA_Faturamento_Impostos1!txt_Razao_Nome))
            
            Set TBTotaisnota = CreateObject("adodb.recordset")
            TBTotaisnota.Open "Select * from tbl_Totais_Nota where id_nota = " & TBLISTA_Faturamento_Impostos1!ID, Conexao, adOpenKeyset, adLockOptimistic
            If TBTotaisnota.EOF = False Then
                'Produtos
                Quant = IIf(IsNull(TBTotaisnota!Qtde_total_prod), 0, TBTotaisnota!Qtde_total_prod)
                .Item(.Count).SubItems(5) = IIf(IsNull(TBTotaisnota!dbl_Valor_Total_Produtos), "", Format(TBTotaisnota!dbl_Valor_Total_Produtos, "###,##0.00"))
    'Debug.print TBTotaisnota!int_NotaFiscal
    
                If IsNull(TBTotaisnota!Valor_total_ICMS_SN) = False And TBTotaisnota!Valor_total_ICMS_SN > 0 Then
                    .Item(.Count).SubItems(6) = IIf(IsNull(TBTotaisnota!dbl_Valor_Total_Nota), "", Format(TBTotaisnota!dbl_Valor_Total_Nota, "###,##0.00"))
                Else
                    .Item(.Count).SubItems(6) = IIf(IsNull(TBTotaisnota!dbl_Base_ICMS), "", Format(TBTotaisnota!dbl_Base_ICMS, "###,##0.00"))
                End If
                '.Item(.Count).SubItems(7) = IIf(IsNull(TBTotaisnota!Total_Credito_ICMS), "", Format(TBTotaisnota!Total_Credito_ICMS, "###,##0.00"))
                .Item(.Count).SubItems(7) = IIf(IsNull(TBTotaisnota!dbl_Valor_ICMS), "", Format(TBTotaisnota!dbl_Valor_ICMS, "###,##0.00"))
                
                .Item(.Count).SubItems(8) = IIf(IsNull(TBTotaisnota!dbl_Base_ICMS_Subst), "", Format(TBTotaisnota!dbl_Base_ICMS_Subst, "###,##0.00"))
                .Item(.Count).SubItems(9) = IIf(IsNull(TBTotaisnota!dbl_Valor_ICMS_Subst), "", Format(TBTotaisnota!dbl_Valor_ICMS_Subst, "###,##0.00"))
                .Item(.Count).SubItems(10) = IIf(IsNull(TBTotaisnota!dbl_Valor_Total_IPI), "", Format(TBTotaisnota!dbl_Valor_Total_IPI, "###,##0.00"))
                .Item(.Count).SubItems(11) = IIf(IsNull(TBTotaisnota!Total_PIS_prod), "", Format(TBTotaisnota!Total_PIS_prod, "###,##0.00"))
                .Item(.Count).SubItems(12) = IIf(IsNull(TBTotaisnota!Total_Cofins_prod), "", Format(TBTotaisnota!Total_Cofins_prod, "###,##0.00"))
                .Item(.Count).SubItems(13) = IIf(IsNull(TBTotaisnota!Total_CSLL_prod), "", Format(TBTotaisnota!Total_CSLL_prod, "###,##0.00"))
                .Item(.Count).SubItems(14) = IIf(IsNull(TBTotaisnota!Total_IRPJ_prod), "", Format(TBTotaisnota!Total_IRPJ_prod, "###,##0.00"))
                .Item(.Count).SubItems(15) = IIf(IsNull(TBTotaisnota!Total_retencao_PIS), "", Format(TBTotaisnota!Total_retencao_PIS, "###,##0.00"))
                .Item(.Count).SubItems(16) = IIf(IsNull(TBTotaisnota!Total_retencao_Cofins), "", Format(TBTotaisnota!Total_retencao_Cofins, "###,##0.00"))
                'Serviços
                quantidade = IIf(IsNull(TBTotaisnota!Qtde_total_serv), 0, TBTotaisnota!Qtde_total_serv)
                .Item(.Count).SubItems(17) = IIf(IsNull(TBTotaisnota!dbl_Valor_Total_Nota_Serv), "", Format(TBTotaisnota!dbl_Valor_Total_Nota_Serv, "###,##0.00"))
                .Item(.Count).SubItems(18) = IIf(IsNull(TBTotaisnota!Total_PIS_serv), "", Format(TBTotaisnota!Total_PIS_serv, "###,##0.00"))
                .Item(.Count).SubItems(19) = IIf(IsNull(TBTotaisnota!Total_Cofins_serv), "", Format(TBTotaisnota!Total_Cofins_serv, "###,##0.00"))
                .Item(.Count).SubItems(20) = IIf(IsNull(TBTotaisnota!Total_CSLL_serv), "", Format(TBTotaisnota!Total_CSLL_serv, "###,##0.00"))
                .Item(.Count).SubItems(21) = IIf(IsNull(TBTotaisnota!dbl_valor_total_iss), "", Format(TBTotaisnota!dbl_valor_total_iss, "###,##0.00"))
                .Item(.Count).SubItems(22) = IIf(IsNull(TBTotaisnota!Total_INSS_serv), "", Format(TBTotaisnota!Total_INSS_serv, "###,##0.00"))
                .Item(.Count).SubItems(23) = IIf(IsNull(TBTotaisnota!Total_IRPJ_serv), "", Format(TBTotaisnota!Total_IRPJ_serv, "###,##0.00"))
                .Item(.Count).SubItems(24) = IIf(IsNull(TBTotaisnota!Total_IRRF_serv), "", Format(TBTotaisnota!Total_IRRF_serv, "###,##0.00"))
                'Outros
                .Item(.Count).SubItems(25) = IIf(IsNull(TBTotaisnota!dbl_Valor_Frete), "", Format(TBTotaisnota!dbl_Valor_Frete, "###,##0.00"))
                .Item(.Count).SubItems(26) = IIf(IsNull(TBTotaisnota!dbl_Valor_Seguro), "", Format(TBTotaisnota!dbl_Valor_Seguro, "###,##0.00"))
                .Item(.Count).SubItems(27) = IIf(IsNull(TBTotaisnota!dbl_Desp_Adicionais), "", Format(TBTotaisnota!dbl_Desp_Adicionais, "###,##0.00"))
                .Item(.Count).SubItems(28) = IIf(IsNull(TBTotaisnota!Total_DAS), "", Format(TBTotaisnota!Total_DAS, "###,##0.00"))
                'Total
                Valor1 = IIf(IsNull(TBTotaisnota!dbl_Valor_Total_Nota), 0, TBTotaisnota!dbl_Valor_Total_Nota)
                Valor2 = IIf(IsNull(TBTotaisnota!Total_retencao_PIS), 0, TBTotaisnota!Total_retencao_PIS)
                Valor3 = IIf(IsNull(TBTotaisnota!Total_retencao_Cofins), 0, TBTotaisnota!Total_retencao_Cofins)
                .Item(.Count).SubItems(29) = Format(Valor1 - Valor2 - Valor3, "###,##0.00")
                .Item(.Count).SubItems(30) = IIf(IsNull(TBTotaisnota!Valor_total_receber_pagar), "0,00", Format(TBTotaisnota!Valor_total_receber_pagar, "###,##0.00"))
            End If
            TBTotaisnota.Close
        End With
    End If
    TBLISTA_Faturamento_Impostos1.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista.Value = Contador
Loop
lblRegistros1.Caption = "Nº de registros: " & TBLISTA_Faturamento_Impostos1.RecordCount
If TBLISTA_Faturamento_Impostos1.AbsolutePage = adPosBOF Then
   lblPaginas1.Caption = "Página: 1 de: " & TBLISTA_Faturamento_Impostos1.PageCount
ElseIf TBLISTA_Faturamento_Impostos1.AbsolutePage = adPosEOF Then
        lblPaginas1.Caption = "Página: " & TBLISTA_Faturamento_Impostos1.PageCount & " de: " & TBLISTA_Faturamento_Impostos1.PageCount
    Else
        lblPaginas1.Caption = "Página: " & TBLISTA_Faturamento_Impostos1.AbsolutePage - 1 & " de: " & TBLISTA_Faturamento_Impostos1.PageCount
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcSomaImpostos()
On Error GoTo tratar_erro

Posicao = 0
Quant = 0
quantidade = 0

'Produtos saída
ValorICMS = 0 'ICMS saída
ValorPagar = 0
VlrIPI = 0
Valor_PIS_Prod = 0
Valor_Cofins_Prod = 0
Valor_CSLL_Prod = 0
Valor_IRPJ_Prod = 0
Valor_Retencao_PIS = 0
Valor_Retencao_Cofins = 0
'Produtos entrada
VLFRETE = 0 'ICMS entrada
vlrTotalProd = 0 'IPI
CMO = 0 'PIS
CMP = 0 'Cofins
Porcentagem = 0 'CSLL
Vlrmateriallu = 0 'IRPJ
Vlrmaodeobralu = 0 'Retenção PIS
Vlrterceiroslu = 0 'Retenção Cofins
'Serviços saída
Valor_PIS_Serv = 0
Valor_Cofins_Serv = 0
Valor_CSLL_Serv = 0
Valor_ISS_Serv = 0
Valor_INSS_Serv = 0
Valor_IRPJ_Serv = 0
Valor_IRRF_Serv = 0
'Serviços entrada
VlrTotalServ = 0 'PIS
Total = 0 'Cofins
VLSEGURO = 0 'CSLL
VLOUTROS = 0 'ISSQN
VlrSubTotal = 0 'INSS
ValorTotal = 0 'IRPJ
Vlrtotallu = 0 'IRRF
'Faturamento saída
DAS = 0
'Faturamento entrada
Vlrcomvend = 0 'DAS

'Saida
Set TBTotaisnota = CreateObject("adodb.recordset")
Totais = "Sum(TN.dbl_Valor_ICMS) as ValorICMS, Sum(TN.Valor_total_ICMS_SN) as ValorPagar, Sum(TN.dbl_Valor_Total_IPI) as VlrIPI, Sum(TN.Total_PIS_prod) as Valor_PIS_Prod, Sum(TN.Total_Cofins_prod) as Valor_Cofins_Prod, Sum(TN.Total_CSLL_prod) as Valor_CSLL_Prod, Sum(TN.Total_IRPJ_prod) as Valor_IRPJ_Prod, Sum(TN.Total_retencao_PIS) as Valor_Retencao_PIS, Sum(TN.Total_retencao_Cofins) as Valor_Retencao_Cofins, Sum(TN.Total_PIS_serv) as Valor_PIS_Serv, Sum(TN.Total_Cofins_serv) as Valor_Cofins_Serv, Sum(TN.Total_CSLL_serv) as Valor_CSLL_Serv, Sum(TN.dbl_valor_total_iss) as vlriss, Sum(TN.Total_INSS_serv) as Valor_INSS_Serv, Sum(TN.Total_IRPJ_serv) as Valor_IRPJ_Serv, Sum(TN.Total_IRRF_serv) as Valor_IRRF_Serv, Sum(TN.Total_DAS) as DAS"
If OptDomes.Value = True Then
    TBTotaisnota.Open "Select " & Totais & " FROM tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_Totais_Nota TN on NF.ID = TN.ID_nota where DtValidacao IS NOT NULL and NF.int_status = 1 and month((NF.dt_DataEmissao))= '" & M & "' and Year((NF.dt_DataEmissao)) = '" & cmbAno.Text & "' and NF.int_TipoNota = 1", Conexao, adOpenKeyset, adLockOptimistic
ElseIf OptAteomes.Value = True Then
        TBTotaisnota.Open "Select " & Totais & " FROM tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_Totais_Nota TN on NF.ID = TN.ID_nota where DtValidacao IS NOT NULL and NF.int_status = 1 and month((NF.dt_DataEmissao))<= '" & M & "' and Year((NF.dt_DataEmissao)) = '" & cmbAno.Text & "' and NF.int_TipoNota = 1", Conexao, adOpenKeyset, adLockOptimistic
    Else
        TBTotaisnota.Open "Select " & Totais & " FROM tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_Totais_Nota TN on NF.ID = TN.ID_nota where DtValidacao IS NOT NULL and NF.int_status = 1 and NF.dt_DataEmissao Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "' and NF.int_TipoNota = 1", Conexao, adOpenKeyset, adLockOptimistic
End If
If TBTotaisnota.EOF = False Then
    'Produtos saída
    ValorICMS = IIf(IsNull(TBTotaisnota!ValorICMS), 0, TBTotaisnota!ValorICMS) + IIf(IsNull(TBTotaisnota!ValorPagar), 0, TBTotaisnota!ValorPagar)
    VlrIPI = IIf(IsNull(TBTotaisnota!VlrIPI), 0, TBTotaisnota!VlrIPI)
    Valor_PIS_Prod = IIf(IsNull(TBTotaisnota!Valor_PIS_Prod), 0, TBTotaisnota!Valor_PIS_Prod)
    Valor_Cofins_Prod = IIf(IsNull(TBTotaisnota!Valor_Cofins_Prod), 0, TBTotaisnota!Valor_Cofins_Prod)
    Valor_CSLL_Prod = IIf(IsNull(TBTotaisnota!Valor_CSLL_Prod), 0, TBTotaisnota!Valor_CSLL_Prod)
    Valor_IRPJ_Prod = IIf(IsNull(TBTotaisnota!Valor_IRPJ_Prod), 0, TBTotaisnota!Valor_IRPJ_Prod)
    Valor_Retencao_PIS = IIf(IsNull(TBTotaisnota!Valor_Retencao_PIS), 0, TBTotaisnota!Valor_Retencao_PIS)
    Valor_Retencao_Cofins = IIf(IsNull(TBTotaisnota!Valor_Retencao_Cofins), 0, TBTotaisnota!Valor_Retencao_Cofins)
    'Serviços saída
    Valor_PIS_Serv = IIf(IsNull(TBTotaisnota!Valor_PIS_Serv), 0, TBTotaisnota!Valor_PIS_Serv)
    Valor_Cofins_Serv = IIf(IsNull(TBTotaisnota!Valor_Cofins_Serv), 0, TBTotaisnota!Valor_Cofins_Serv)
    Valor_CSLL_Serv = IIf(IsNull(TBTotaisnota!Valor_CSLL_Serv), 0, TBTotaisnota!Valor_CSLL_Serv)
    Valor_ISS_Serv = IIf(IsNull(TBTotaisnota!VlrISS), 0, TBTotaisnota!VlrISS)
    Valor_INSS_Serv = IIf(IsNull(TBTotaisnota!Valor_INSS_Serv), 0, TBTotaisnota!Valor_INSS_Serv)
    Valor_IRPJ_Serv = IIf(IsNull(TBTotaisnota!Valor_IRPJ_Serv), 0, TBTotaisnota!Valor_IRPJ_Serv)
    Valor_IRRF_Serv = IIf(IsNull(TBTotaisnota!Valor_IRRF_Serv), 0, TBTotaisnota!Valor_IRRF_Serv)
    'Faturamento saída
    DAS = IIf(IsNull(TBTotaisnota!DAS), 0, TBTotaisnota!DAS)
End If

'Entrada
Set TBTotaisnota = CreateObject("adodb.recordset")
'Totais = "Sum(TN.Total_Credito_ICMS) as VLFRETE, Sum(TN.dbl_Valor_Total_IPI) as vlrTotalProd, Sum(TN.Total_PIS_prod) as CMO, Sum(TN.Total_Cofins_prod) as CMP, Sum(TN.Total_CSLL_prod) as Porcentagem, Sum(TN.Total_IRPJ_prod) as Vlrmateriallu, Sum(TN.Total_retencao_PIS) as Vlrmaodeobralu, Sum(TN.Total_retencao_Cofins) as Vlrterceiroslu, Sum(TN.Total_PIS_serv) as VlrTotalServ, Sum(TN.Total_Cofins_serv) as Total, Sum(TN.Total_CSLL_serv) as VLSEGURO, Sum(TN.dbl_valor_total_iss) as VLOUTROS, Sum(TN.Total_INSS_serv) as VlrSubTotal, Sum(TN.Total_IRPJ_serv) as Valortotal, Sum(TN.Total_IRRF_serv) as Vlrtotallu, Sum(TN.Total_DAS) as Vlrcomvend"
Totais = "Sum(TN.dbl_Valor_ICMS) as VLFRETE, Sum(TN.dbl_Valor_Total_IPI) as vlrTotalProd, Sum(TN.Total_PIS_prod) as CMO, Sum(TN.Total_Cofins_prod) as CMP, Sum(TN.Total_CSLL_prod) as Porcentagem, Sum(TN.Total_IRPJ_prod) as Vlrmateriallu, Sum(TN.Total_retencao_PIS) as Vlrmaodeobralu, Sum(TN.Total_retencao_Cofins) as Vlrterceiroslu, Sum(TN.Total_PIS_serv) as VlrTotalServ, Sum(TN.Total_Cofins_serv) as Total, Sum(TN.Total_CSLL_serv) as VLSEGURO, Sum(TN.dbl_valor_total_iss) as VLOUTROS, Sum(TN.Total_INSS_serv) as VlrSubTotal, Sum(TN.Total_IRPJ_serv) as Valortotal, Sum(TN.Total_IRRF_serv) as Vlrtotallu, Sum(TN.Total_DAS) as Vlrcomvend"

If OptDomes.Value = True Then
    TBTotaisnota.Open "Select " & Totais & " FROM tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_Totais_Nota TN on NF.ID = TN.ID_nota where DtValidacao IS NOT NULL and NF.int_status = 1 and month(NF.dt_Saida_Entrada)= '" & M & "' and Year(NF.dt_Saida_Entrada) = '" & cmbAno.Text & "' and NF.int_TipoNota = 2", Conexao, adOpenKeyset, adLockOptimistic
ElseIf OptAteomes.Value = True Then
        TBTotaisnota.Open "Select " & Totais & " FROM tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_Totais_Nota TN on NF.ID = TN.ID_nota where DtValidacao IS NOT NULL and NF.int_status = 1 and month(NF.dt_Saida_Entrada)<= '" & M & "' and Year(NF.dt_Saida_Entrada) = '" & cmbAno.Text & "' and NF.int_TipoNota = 2", Conexao, adOpenKeyset, adLockOptimistic
    Else
        TBTotaisnota.Open "Select " & Totais & " FROM tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_Totais_Nota TN on NF.ID = TN.ID_nota where DtValidacao IS NOT NULL and NF.int_status = 1 and NF.dt_Saida_Entrada Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "' and NF.int_TipoNota = 2", Conexao, adOpenKeyset, adLockOptimistic
End If
If TBTotaisnota.EOF = False Then
    VLFRETE = IIf(IsNull(TBTotaisnota!VLFRETE), 0, TBTotaisnota!VLFRETE)
    vlrTotalProd = IIf(IsNull(TBTotaisnota!vlrTotalProd), 0, TBTotaisnota!vlrTotalProd)
    CMO = IIf(IsNull(TBTotaisnota!CMO), 0, TBTotaisnota!CMO)
    CMP = IIf(IsNull(TBTotaisnota!CMP), 0, TBTotaisnota!CMP)
    Porcentagem = IIf(IsNull(TBTotaisnota!Porcentagem), 0, TBTotaisnota!Porcentagem)
    Vlrmateriallu = IIf(IsNull(TBTotaisnota!Vlrmateriallu), 0, TBTotaisnota!Vlrmateriallu)
    Vlrmaodeobralu = IIf(IsNull(TBTotaisnota!Vlrmaodeobralu), 0, TBTotaisnota!Vlrmaodeobralu)
    Vlrterceiroslu = IIf(IsNull(TBTotaisnota!Vlrterceiroslu), 0, TBTotaisnota!Vlrterceiroslu)
    'Serviços entrada
    VlrTotalServ = IIf(IsNull(TBTotaisnota!VlrTotalServ), 0, TBTotaisnota!VlrTotalServ)
    Total = IIf(IsNull(TBTotaisnota!Total), 0, TBTotaisnota!Total)
    VLSEGURO = IIf(IsNull(TBTotaisnota!VLSEGURO), 0, TBTotaisnota!VLSEGURO)
    VLOUTROS = IIf(IsNull(TBTotaisnota!VLOUTROS), 0, TBTotaisnota!VLOUTROS)
    VlrSubTotal = IIf(IsNull(TBTotaisnota!VlrSubTotal), 0, TBTotaisnota!VlrSubTotal)
    ValorTotal = IIf(IsNull(TBTotaisnota!ValorTotal), 0, TBTotaisnota!ValorTotal)
    Vlrtotallu = IIf(IsNull(TBTotaisnota!Vlrtotallu), 0, TBTotaisnota!Vlrtotallu)
    'Faturamento entrada
    Vlrcomvend = IIf(IsNull(TBTotaisnota!Vlrcomvend), 0, TBTotaisnota!Vlrcomvend)
End If

'Produtos saída
txtValorICMS_Saída = Format(ValorICMS, "###,##0.00")
txtValorICMS_Entrada = Format(VLFRETE, "###,##0.00")
txtSaldo_ICMS = Format(VLFRETE - ValorICMS, "###,##0.00")
txtValorIPI_saida = Format(VlrIPI, "###,##0.00")
txtValorPIS_prod_saida = Format(Valor_PIS_Prod, "###,##0.00")
txtValorCofins_prod_saida = Format(Valor_Cofins_Prod, "###,##0.00")
txtValorCSLL_prod_saida = Format(Valor_CSLL_Prod, "###,##0.00")
txtValorIRPJ_prod_saida = Format(Valor_IRPJ_Prod, "###,##0.00")
txtValor_retencao_PIS_prod_saida = Format(Valor_Retencao_PIS, "###,##0.00")
txtValor_retencao_Cofins_prod_saida = Format(Valor_Retencao_Cofins, "###,##0.00")
'Produtos entrada
txtValorIPI_entrada = Format(vlrTotalProd, "###,##0.00")
txtValorPIS_prod_entrada = Format(CMO, "###,##0.00")
txtValorCofins_prod_entrada = Format(CMP, "###,##0.00")
txtValorCSLL_prod_entrada = Format(Porcentagem, "###,##0.00")
txtValorIRPJ_prod_entrada = Format(Vlrmateriallu, "###,##0.00")
txtValor_retencao_PIS_prod_entrada = Format(Vlrmaodeobralu, "###,##0.00")
txtValor_retencao_Cofins_prod_entrada = Format(Vlrterceiroslu, "###,##0.00")
'Serviços saída
txtValorPIS_serv_saida = Format(Valor_PIS_Serv, "###,##0.00")
txtValorCofins_serv_saida = Format(Valor_Cofins_Serv, "###,##0.00")
txtValorCSLL_serv_saida = Format(Valor_CSLL_Serv, "###,##0.00")
txtValorISS_serv_saida = Format(VlrISS, "###,##0.00")
txtValorINSS_serv_saida = Format(Valor_INSS_Serv, "###,##0.00")
txtValorIRPJ_serv_saida = Format(Valor_IRPJ_Serv, "###,##0.00")
txtValorIRRF_serv_saida = Format(Valor_IRRF_Serv, "###,##0.00")
'Serviços entrada
txtValorPIS_serv_entrada = Format(VlrTotalServ, "###,##0.00")
txtValorCofins_serv_entrada = Format(Total, "###,##0.00")
txtValorCSLL_serv_entrada = Format(VLSEGURO, "###,##0.00")
txtValorISS_serv_entrada = Format(VLOUTROS, "###,##0.00")
txtValorINSS_serv_entrada = Format(VlrSubTotal, "###,##0.00")
txtValorIRPJ_serv_entrada = Format(ValorTotal, "###,##0.00")
txtValorIRRF_serv_entrada = Format(Vlrtotallu, "###,##0.00")
'Faturamento saída
txtValorDAS_saida = Format(DAS, "###,##0.00")
'Faturamento entrada
txtValorDAS_entrada = Format(Vlrcomvend, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
