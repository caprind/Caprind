VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProd_programacao 
   BackColor       =   &H00E0E0E0&
   Caption         =   "PCP - Programação da produção"
   ClientHeight    =   10035
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15360
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   7.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProd_programacao.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10035
   ScaleWidth      =   15360
   WindowState     =   2  'Maximized
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
      FormWidthDT     =   15480
      FormScaleHeightDT=   10035
      FormScaleWidthDT=   15360
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   60
      TabIndex        =   41
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
      SearchText      =   "Atualizando..."
      Value           =   0
   End
   Begin TabDlg.SSTab SStab1 
      Height          =   10065
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   15360
      _ExtentX        =   27093
      _ExtentY        =   17754
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
      TabCaption(0)   =   "Posto de trabalho"
      TabPicture(0)   =   "frmProd_programacao.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtIDmaquina"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtID"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "USToolBar1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Lista"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Programação"
      TabPicture(1)   =   "frmProd_programacao.frx":045E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Lista_prioridade"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "USToolBar2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "txtIDprioridade"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      Begin VB.TextBox txtIDprioridade 
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
         Left            =   1995
         Locked          =   -1  'True
         MaxLength       =   10
         MousePointer    =   99  'Custom
         TabIndex        =   38
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   4380
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.TextBox txtIDmaquina 
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
         Left            =   -73215
         Locked          =   -1  'True
         MaxLength       =   10
         MousePointer    =   99  'Custom
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   3180
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.TextBox txtID 
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
         Left            =   -74235
         Locked          =   -1  'True
         MaxLength       =   10
         MousePointer    =   99  'Custom
         TabIndex        =   36
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   3120
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   825
         Left            =   -74925
         TabIndex        =   19
         Top             =   1185
         Width           =   15195
         Begin VB.ComboBox cmbMaquina 
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
            MouseIcon       =   "frmProd_programacao.frx":047A
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   0
            ToolTipText     =   "Posto de trabalho."
            Top             =   390
            Width           =   2670
         End
         Begin VB.CommandButton cmdMaquina 
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   2865
            MouseIcon       =   "frmProd_programacao.frx":0784
            MousePointer    =   99  'Custom
            Picture         =   "frmProd_programacao.frx":0A8E
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Localizar posto de trabalho."
            Top             =   390
            Width           =   315
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
            Left            =   3225
            Locked          =   -1  'True
            MouseIcon       =   "frmProd_programacao.frx":0B90
            MousePointer    =   99  'Custom
            TabIndex        =   2
            TabStop         =   0   'False
            ToolTipText     =   "Descrição do posto de trabalho."
            Top             =   390
            Width           =   11835
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descrição do posto de trabalho"
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
            Left            =   6255
            TabIndex        =   21
            Top             =   180
            Width           =   2235
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Posto de trabalho"
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
            Left            =   878
            TabIndex        =   20
            Top             =   180
            Width           =   1275
         End
      End
      Begin VB.Frame Frame2 
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
         Height          =   2745
         Left            =   75
         TabIndex        =   22
         Top             =   1185
         Width           =   15195
         Begin VB.TextBox txtInstrucoes 
            Height          =   675
            Left            =   180
            Locked          =   -1  'True
            TabIndex        =   52
            Top             =   1980
            Width           =   14865
         End
         Begin VB.CommandButton cmdFiltrarData 
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   14700
            MouseIcon       =   "frmProd_programacao.frx":0E9A
            MousePointer    =   99  'Custom
            Picture         =   "frmProd_programacao.frx":11A4
            Style           =   1  'Graphical
            TabIndex        =   51
            ToolTipText     =   "Localizar ordem."
            Top             =   1560
            Width           =   345
         End
         Begin MSComCtl2.DTPicker txtDataFinal 
            Height          =   315
            Left            =   12420
            TabIndex        =   40
            ToolTipText     =   "Data de início."
            Top             =   1560
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
            Format          =   198377475
            CurrentDate     =   39057
         End
         Begin VB.TextBox txtCliente 
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
            TabIndex        =   9
            TabStop         =   0   'False
            ToolTipText     =   "Cliente."
            Top             =   975
            Width           =   14895
         End
         Begin VB.TextBox txtReferencia 
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
            Left            =   5265
            Locked          =   -1  'True
            TabIndex        =   7
            TabStop         =   0   'False
            ToolTipText     =   "Código de referência."
            Top             =   390
            Width           =   1755
         End
         Begin VB.TextBox txtDescricao_item 
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
            Left            =   7035
            Locked          =   -1  'True
            TabIndex        =   8
            TabStop         =   0   'False
            ToolTipText     =   "Descrição."
            Top             =   390
            Width           =   8025
         End
         Begin VB.TextBox txtRevitem 
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
            Left            =   4815
            Locked          =   -1  'True
            TabIndex        =   6
            TabStop         =   0   'False
            ToolTipText     =   "Revisão."
            Top             =   390
            Width           =   435
         End
         Begin VB.TextBox txtDesenho 
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
            Left            =   3060
            Locked          =   -1  'True
            TabIndex        =   5
            TabStop         =   0   'False
            ToolTipText     =   "Código interno."
            Top             =   390
            Width           =   1740
         End
         Begin VB.TextBox txtTempo 
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
            Left            =   8100
            Locked          =   -1  'True
            MaxLength       =   10
            MousePointer    =   99  'Custom
            TabIndex        =   14
            TabStop         =   0   'False
            ToolTipText     =   "Tempo total."
            Top             =   1560
            Width           =   2040
         End
         Begin VB.TextBox txtFase 
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
            Left            =   1980
            MaxLength       =   10
            MousePointer    =   99  'Custom
            TabIndex        =   11
            ToolTipText     =   "Fase."
            Top             =   1560
            Width           =   1095
         End
         Begin VB.TextBox txtGrupo 
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
            Left            =   3090
            Locked          =   -1  'True
            TabIndex        =   12
            TabStop         =   0   'False
            ToolTipText     =   "Grupo/operação"
            Top             =   1560
            Width           =   1095
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
            Left            =   4200
            MaxLength       =   10
            MousePointer    =   99  'Custom
            TabIndex        =   13
            TabStop         =   0   'False
            ToolTipText     =   "Quantidade."
            Top             =   1560
            Width           =   1725
         End
         Begin VB.TextBox txtOS 
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
            TabIndex        =   10
            ToolTipText     =   "Número da OS."
            Top             =   1560
            Width           =   1770
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
            Left            =   1395
            MaxLength       =   10
            MousePointer    =   99  'Custom
            TabIndex        =   4
            ToolTipText     =   "Número da ordem."
            Top             =   390
            Width           =   1665
         End
         Begin VB.TextBox txtPrioridade 
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
            MaxLength       =   10
            MousePointer    =   99  'Custom
            TabIndex        =   3
            ToolTipText     =   "Prioridade."
            Top             =   390
            Width           =   1200
         End
         Begin MSComCtl2.DTPicker txtHoraInicio 
            Height          =   315
            Left            =   11370
            TabIndex        =   16
            ToolTipText     =   "Hora início."
            Top             =   1560
            Width           =   1035
            _ExtentX        =   1826
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
            Format          =   198377474
            CurrentDate     =   39055
         End
         Begin MSComCtl2.DTPicker txtDataInicio 
            Height          =   315
            Left            =   10155
            TabIndex        =   15
            ToolTipText     =   "Data de início."
            Top             =   1560
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
            Format          =   198377475
            CurrentDate     =   39057
         End
         Begin MSComCtl2.DTPicker txtHoraFinal 
            Height          =   315
            Left            =   13650
            TabIndex        =   42
            ToolTipText     =   "Hora final."
            Top             =   1560
            Width           =   1035
            _ExtentX        =   1826
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
            Format          =   198377474
            CurrentDate     =   39055
         End
         Begin MSMask.MaskEdBox txtpreparacao 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "H:mm:ss"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   4
            EndProperty
            Height          =   315
            Left            =   5925
            TabIndex        =   44
            ToolTipText     =   "Tempo de preparação previsto."
            Top             =   1560
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   0
            AutoTab         =   -1  'True
            MaxLength       =   9
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "###:##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtexecucao 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "H:mm:ss"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   4
            EndProperty
            Height          =   315
            Left            =   7020
            TabIndex        =   45
            ToolTipText     =   "Tempo de execução previsto."
            Top             =   1560
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   0
            AutoTab         =   -1  'True
            MaxLength       =   9
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "###:##:##"
            PromptChar      =   "_"
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Execução"
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
            Index           =   27
            Left            =   7200
            TabIndex        =   47
            Top             =   1350
            Width           =   705
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Preparação"
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
            Index           =   26
            Left            =   6030
            TabIndex        =   46
            Top             =   1350
            Width           =   855
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Final"
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
            Left            =   13410
            TabIndex        =   43
            Top             =   1320
            Width           =   330
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cliente"
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
            Left            =   5595
            TabIndex        =   39
            Top             =   765
            Width           =   495
         End
         Begin VB.Label Label16 
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
            Left            =   10702
            TabIndex        =   35
            Top             =   180
            Width           =   690
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cod. de referência"
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
            Left            =   5467
            TabIndex        =   34
            Top             =   180
            Width           =   1350
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rev."
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
            Left            =   4860
            TabIndex        =   33
            Top             =   180
            Width           =   345
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
            Left            =   3480
            TabIndex        =   32
            Top             =   180
            Width           =   900
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Início"
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
            Left            =   11070
            TabIndex        =   31
            Top             =   1350
            Width           =   375
         End
         Begin VB.Label Label14 
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
            Left            =   4642
            TabIndex        =   30
            Top             =   1350
            Width           =   840
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sequência"
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
            Left            =   413
            TabIndex        =   29
            Top             =   180
            Width           =   735
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tempo total"
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
            Left            =   8693
            TabIndex        =   28
            Top             =   1350
            Width           =   855
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Grupo/op."
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
            Left            =   3270
            TabIndex        =   27
            Top             =   1350
            Width           =   735
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fase"
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
            Left            =   2355
            TabIndex        =   26
            Top             =   1350
            Width           =   345
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "OS"
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
            Left            =   953
            TabIndex        =   25
            Top             =   1350
            Width           =   225
         End
         Begin VB.Label Label5 
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
            Left            =   1942
            TabIndex        =   24
            Top             =   180
            Width           =   570
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Instruções de trabalho"
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
            Left            =   5025
            TabIndex        =   23
            Top             =   1950
            Width           =   1635
         End
      End
      Begin DrawSuite2022.USToolBar USToolBar1 
         Height          =   855
         Left            =   -74910
         TabIndex        =   48
         Top             =   330
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   1508
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
         ButtonLeft6     =   209
         ButtonTop6      =   4
         ButtonWidth6    =   2
         ButtonHeight6   =   46
         ButtonCaption7  =   "Ajuda"
         ButtonEnabled7  =   0   'False
         ButtonIconSize7 =   32
         ButtonToolTipText7=   "Ajuda (F1)"
         ButtonKey7      =   "14"
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
         ButtonLeft7     =   213
         ButtonTop7      =   2
         ButtonWidth7    =   36
         ButtonHeight7   =   21
         ButtonUseMaskColor7=   0   'False
         ButtonCaption8  =   "Sair"
         ButtonEnabled8  =   0   'False
         ButtonIconSize8 =   32
         ButtonToolTipText8=   "Sair (Esc)"
         ButtonKey8      =   "15"
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
         ButtonLeft8     =   251
         ButtonTop8      =   2
         ButtonWidth8    =   26
         ButtonHeight8   =   21
         ButtonUseMaskColor8=   0   'False
         ButtonEnabled9  =   0   'False
         ButtonIconSize9 =   32
         ButtonKey9      =   "16"
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
         ButtonState9    =   5
         ButtonLeft9     =   279
         ButtonTop9      =   2
         ButtonWidth9    =   24
         ButtonHeight9   =   24
         ButtonUseMaskColor9=   0   'False
         Begin DrawSuite2022.USImageList USImageList1 
            Left            =   11640
            Top             =   195
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmProd_programacao.frx":1766
            Count           =   1
         End
      End
      Begin DrawSuite2022.USToolBar USToolBar2 
         Height          =   855
         Left            =   60
         TabIndex        =   49
         Top             =   330
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   1508
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
         ButtonCaption2  =   "Salvar"
         ButtonEnabled2  =   0   'False
         ButtonIconSize2 =   32
         ButtonToolTipText2=   "Salvar (F3)"
         ButtonKey2      =   "3"
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
         ButtonWidth2    =   38
         ButtonHeight2   =   21
         ButtonUseMaskColor2=   0   'False
         ButtonCaption3  =   "Excluir"
         ButtonEnabled3  =   0   'False
         ButtonIconSize3 =   32
         ButtonToolTipText3=   "Excluir (F4)"
         ButtonKey3      =   "4"
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
         ButtonLeft3     =   77
         ButtonTop3      =   2
         ButtonWidth3    =   39
         ButtonHeight3   =   21
         ButtonUseMaskColor3=   0   'False
         ButtonCaption4  =   "Relatório"
         ButtonEnabled4  =   0   'False
         ButtonIconSize4 =   32
         ButtonToolTipText4=   "Relatório (F5)"
         ButtonKey4      =   "5"
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
         ButtonLeft4     =   118
         ButtonTop4      =   2
         ButtonWidth4    =   51
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
         ButtonLeft5     =   171
         ButtonTop5      =   4
         ButtonWidth5    =   2
         ButtonHeight5   =   46
         ButtonCaption6  =   "Ajuda"
         ButtonEnabled6  =   0   'False
         ButtonIconSize6 =   32
         ButtonToolTipText6=   "Ajuda (F1)"
         ButtonKey6      =   "14"
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
         ButtonLeft6     =   175
         ButtonTop6      =   2
         ButtonWidth6    =   36
         ButtonHeight6   =   21
         ButtonUseMaskColor6=   0   'False
         ButtonCaption7  =   "Sair"
         ButtonEnabled7  =   0   'False
         ButtonIconSize7 =   32
         ButtonToolTipText7=   "Sair (Esc)"
         ButtonKey7      =   "15"
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
         ButtonLeft7     =   213
         ButtonTop7      =   2
         ButtonWidth7    =   26
         ButtonHeight7   =   21
         ButtonUseMaskColor7=   0   'False
         ButtonEnabled8  =   0   'False
         ButtonIconSize8 =   32
         ButtonKey8      =   "16"
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
         ButtonLeft8     =   241
         ButtonTop8      =   2
         ButtonWidth8    =   24
         ButtonHeight8   =   24
         ButtonUseMaskColor8=   0   'False
         Begin DrawSuite2022.USImageList USImageList2 
            Left            =   11640
            Top             =   195
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmProd_programacao.frx":62EC
            Count           =   1
         End
      End
      Begin MSComctlLib.ListView Lista_prioridade 
         Height          =   5775
         Left            =   75
         TabIndex        =   17
         Top             =   3945
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   10186
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
         MousePointer    =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmProd_programacao.frx":A2EF
         NumItems        =   11
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "D"
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Inicio"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Fim"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Prioridade"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Ordem"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Cód. interno"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Rev."
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Descrição"
            Object.Width           =   15875
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Object.Tag             =   "T"
            Text            =   "OS"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   9
            Text            =   "Fase"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Grupo/op."
            Object.Width           =   1764
         EndProperty
      End
      Begin MSComctlLib.ListView Lista 
         Height          =   7710
         Left            =   -74940
         TabIndex        =   50
         Top             =   2010
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   13600
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
         MousePointer    =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmProd_programacao.frx":A609
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "N"
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Tag             =   "T"
            Text            =   "Posto de trabalho"
            Object.Width           =   3705
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Descrição"
            Object.Width           =   31750
         EndProperty
      End
   End
End
Attribute VB_Name = "frmProd_programacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public StrSql_Ordem_programacao     As String 'OK
Dim Novo_ordem_programacao          As Boolean 'OK
Dim Novo_ordem_programacao1         As Boolean 'OK
Public FormulaRel_Ordem_Programacao    As String 'OK
Dim PrioridadeAntiga                As Integer 'OK
Public FiltroImpressao              As Boolean 'OK
Dim InicioIntervalo                 As Date 'OK
Dim TotalIntervalo                  As Date 'OK
Dim DataTotal                       As Date 'OK
Dim Prioridade                      As Integer 'OK
Dim DatainicioINTER                 As Date
Dim DatafimINTER                    As Date
Dim DataTempoINTER                  As Date

Private Sub cmbMaquina_Click()
On Error GoTo tratar_erro

If cmbMaquina <> "" Then txtIDmaquina = cmbMaquina.ItemData(cmbMaquina.ListIndex)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAbrir()
On Error GoTo tratar_erro

frmProd_programacao_abrir.Show 1
        
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
If txtId.Text = "0" Then
   USMsgBox ("Informe o posto de trabalho antes de excluir."), vbExclamation, "CAPRIND v5.0"
   Exit Sub
End If
If USMsgBox("Deseja realmente excluir a programação do posto de trabalho " & cmbMaquina.Text & "?", vbYesNo) = vbYes Then
    Conexao.Execute "DELETE from PCP_programacao where ID = " & txtId
    Conexao.Execute "DELETE from PCP_programacao_ordem where IDprogramacao = " & txtId
    USMsgBox ("Programação do posto de trabalho excluído com sucesso."), vbInformation, "CAPRIND v5.0"
    '==================================
    Modulo = "PCP/Programação da produção"
    Evento = "Excluir"
    ID_documento = txtId
    Documento = "Posto de trabalho: " & cmbMaquina
    Documento1 = ""
    ProcGravaEvento
    '==================================
    Frame3.Enabled = False
    ProcLimpaCampos
    ProcCarregaLista
End If
        
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluir_ordem()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If txtIDprioridade.Text = "0" Then
   USMsgBox ("Informe a ordem antes de excluir."), vbExclamation, "CAPRIND v5.0"
   Exit Sub
End If
If USMsgBox("Deseja realmente excluir a ordem da programação?", vbYesNo) = vbYes Then
    Conexao.Execute "DELETE from PCP_programacao_ordem where ID = " & txtIDprioridade
    USMsgBox ("Ordem excluída da programação com sucesso."), vbInformation, "CAPRIND v5.0"
    '==================================
    Modulo = "PCP/Programação da produção"
    Evento = "Excluir ordem"
    ID_documento = txtId
    Documento = "Posto de trabalho: " & cmbMaquina.Text
    Documento1 = "Ordem: " & txtOrdem & " OS: " & txtOS
    ProcGravaEvento
    '==================================
    ProcCorrigirPrioridadeExcluir
    Frame2.Enabled = False
    ProcLimpaCamposOrdem
    ProcCarregaListaOrdem
End If
        
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub ProcSair()
On Error GoTo tratar_erro

If Novo_ordem_programacao = True Then
    If USMsgBox("O posto de trabalho ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvar
        If Novo_ordem_programacao = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
If Novo_ordem_programacao1 = True Then
    If USMsgBox("A programação ainda não foi salva, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvar_ordem
        If Novo_ordem_programacao1 = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
Novo_ordem_programacao = False
Novo_ordem_programacao1 = False
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdImprimir_ordem_Click()
On Error GoTo tratar_erro

ProcImprimir

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdImprimir_ordem2_Click()
On Error GoTo tratar_erro

ProcImprimir2

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub cmdFiltrarData_Click()
On Error GoTo tratar_erro

ProcCarregaListaOrdem

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub

End Sub

Private Sub cmdMaquina_Click()
On Error GoTo tratar_erro

frmProd_programacao_maquina.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovo_ordem()
On Error GoTo tratar_erro

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
ProcLimpaCamposOrdem
Frame2.Enabled = True
Novo_ordem_programacao1 = True
frmProd_programacao_OS.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub cmdordem_Click()
On Error GoTo tratar_erro

frmProd_programacao_Ordem.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdOS_Click()
On Error GoTo tratar_erro

frmProd_programacao_OS.Show 1

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
If Frame3.Enabled = False Then
    ProcVerificaSalvar
    cmdNovo.SetFocus
    Exit Sub
End If
Acao = "salvar"
If cmbMaquina.Text = "" Then
    NomeCampo = "o posto de trabalho"
    ProcVerificaAcao
    frmProd_programacao_maquina.Show 1
    Exit Sub
End If
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from PCP_programacao where idmaquina = " & txtIDmaquina & " and ID <> " & txtId, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    USMsgBox ("Posto de trabalho já cadastrado, favor alterar."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
TBAbrir.Close

Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "select * from PCP_programacao where id = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then TBGravar.AddNew
TBGravar!IDMaquina = txtIDmaquina
TBGravar.Update
txtId = TBGravar!ID

ProcCarregaLista
If Novo_ordem_programacao = True Then
    USMsgBox ("Nova programação cadastrada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar"
    If Lista.ListItems.Count <> 0 And CodigoLista <> 0 Then
        Lista.SelectedItem = Lista.ListItems(CodigoLista)
        Lista.SetFocus
    End If
End If

'==================================
Modulo = "PCP/Programação da produção"
ID_documento = txtId
Documento = "Posto de trabalho: " & cmbMaquina.Text
Documento1 = ""
ProcGravaEvento
'==================================
Novo_ordem_programacao = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvar_ordem()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Frame2.Enabled = False Then
    ProcVerificaSalvar
    cmdNovo.SetFocus
    Exit Sub
End If
Acao = "salvar"
If txtOrdem.Text = "" Then
    NomeCampo = "a ordem"
    ProcVerificaAcao
    txtOrdem.SetFocus
    Exit Sub
End If
If txtOS.Text = "" Then
    NomeCampo = "a OS"
    ProcVerificaAcao
    txtOS.SetFocus
    Exit Sub
End If
'verificar se existe a ordem
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "select * from producao where ordem = " & txtOrdem, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = True Then
    USMsgBox "Não foi encontrada nenhuma ordem com o numero " & txtOrdem & ", favor alterar antes de salvar", vbExclamation, "CAPRIND v5.0"
    TBAbrir.Close
    Exit Sub
End If
TBAbrir.Close
'Verificar se exesite a OS
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "select * from ordemservico where idproducao = " & txtOS, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = True Then
    USMsgBox "Não foi encontrada nenhuma OS com o numero " & txtOS & ", favor alterar antes de salvar", vbExclamation, "CAPRIND v5.0"
    TBAbrir.Close
    Exit Sub
End If
TBAbrir.Close

Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "select * from PCP_programacao_ordem where id = " & txtIDprioridade, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then
    TBGravar.AddNew
Else
    PrioridadeAntiga = IIf(IsNull(TBGravar!Prioridade), 0, TBGravar!Prioridade)
End If
ProcEnviaDados
TBGravar.Update
txtIDprioridade = TBGravar!ID
ProcCorrigirPrioridade
ProcCarregaListaOrdem

If Novo_ordem_programacao1 = True Then
    USMsgBox ("Nova ordem agregada a programação com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Nova ordem"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar ordem"
    If Lista_prioridade.ListItems.Count <> 0 And CodigoLista1 <> 0 Then
        Lista_prioridade.SelectedItem = Lista_prioridade.ListItems(CodigoLista1)
        Lista_prioridade.SetFocus
    End If
End If

'==================================
Modulo = "PCP/Programação da produção"
ID_documento = txtId
Documento = "Posto de trabalho: " & cmbMaquina
Documento1 = "Ordem: " & txtOrdem & " OS: " & txtOS
ProcGravaEvento
'==================================
Novo_ordem_programacao1 = False

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
            Case vbKeyF5: ProcImprimir2
            Case vbKeyEscape: ProcSair
        End Select
    Case 1:
        Select Case KeyCode
            Case vbKeyInsert: ProcNovo_ordem
            Case vbKeyF3: ProcSalvar_ordem
            Case vbKeyF4: ProcExcluir_ordem
            Case vbKeyF5: ProcImprimir2
            Case vbKeyEscape: ProcSair
        End Select
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaLista()
On Error GoTo tratar_erro

Lista.ListItems.Clear
If StrSql_Ordem_programacao = "" Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open StrSql_Ordem_programacao, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    TBLISTA.MoveLast
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    TBLISTA.MoveFirst
    Do While TBLISTA.EOF = False
        With Lista.ListItems
            .Add , , IIf(IsNull(TBLISTA!ID), 0, TBLISTA!ID)
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "select * from CadMaquinas where IDMaquina = " & TBLISTA!IDMaquina, Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                .Item(.Count).SubItems(1) = IIf(IsNull(TBAbrir!maquina), "", TBAbrir!maquina)
                .Item(.Count).SubItems(2) = IIf(IsNull(TBAbrir!Descricao), "", TBAbrir!Descricao)
            End If
            TBAbrir.Close
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

Sub ProcLimpaCampos()
On Error GoTo tratar_erro

txtId = "0"
txtIDmaquina = ""
cmbMaquina.ListIndex = -1
txtdescricao = ""
CodigoLista = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCamposOrdem()
On Error GoTo tratar_erro

txtIDprioridade = "0"
txtPrioridade = "0"
txtOrdem = ""
txtOS = ""
txtFase = ""
txtGrupo = ""
txtQtde = ""
txtTempo = ""
txtDescricao_item = ""
txtdesenho = ""
txtRevitem = ""
txtreferencia = ""
txtDataInicio = Date
txtHoraInicio = "00:00:00"
txtDataFinal = Date
txtHoraFinal = "00:00:00"
txtInstrucoes = ""
txtCliente = ""
CodigoLista1 = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15195, 8, True
ProcCarregaToolBar2 Me, 15195, 7, True
Formulario = "PCP/Programação da produção"
Direitos
ProcLimpaVariaveisPrincipais
ProcCarregaMaquina
SSTab1.Tab = 0
StrSql_Ordem_programacao = "Select * from PCP_programacao where idmaquina <> 0 order by ID"
Lista.ListItems.Clear
ProcCarregaLista

ProcRemoveObjetosResize Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

Formulario = "PCP/Prioridades da produção"
Direitos
ProcLimpaVariaveisPrincipais
ProcCarregaLista


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
Novo_ordem_programacao = True
Frame3.Enabled = True
frmProd_programacao_maquina.Show 1

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
  
If Lista.ListItems.Count = 0 Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from PCP_programacao where id = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    ProcLimpaCampos
    txtId = TBLISTA!ID
    txtIDmaquina = TBLISTA!IDMaquina
    Frame3.Enabled = True
    Novo_ordem_programacao = False
    CodigoLista = Lista.SelectedItem.index
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviaDados()
On Error GoTo tratar_erro

TBGravar!idprogramacao = txtId
TBGravar!Prioridade = IIf(txtPrioridade = "", "0", txtPrioridade)
TBGravar!Ordem = txtOrdem
TBGravar!OS = txtOS
DataInicio = txtDataInicio & " " & txtHoraInicio
DataFinal = txtDataFinal & " " & txtHoraFinal
TBGravar!Inicio = DataInicio
TBGravar!Final = DataFinal
TBGravar!TempoTotal = txtTempo
TBGravar!Qtd = txtQtde

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCorrigirPrioridade()
On Error GoTo tratar_erro

Prioridade = txtPrioridade
Set TBproducao = CreateObject("adodb.recordset")
TBproducao.Open "select * from PCP_programacao_ordem where prioridade = " & txtPrioridade & " and id <> " & txtIDprioridade & " and IDprogramacao = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
If TBproducao.EOF = False Then
    Set TBAbrir = CreateObject("adodb.recordset")
    If txtPrioridade < PrioridadeAntiga Or PrioridadeAntiga = "0" Then
        TBAbrir.Open "select * from PCP_programacao_ordem where prioridade <> 0 and IDprogramacao = " & txtId & " order by prioridade", Conexao, adOpenKeyset, adLockOptimistic
        Do While TBAbrir.EOF = False
            If TBAbrir!ID <> txtIDprioridade Then
                If TBAbrir!Prioridade = Prioridade Then
                    TBAbrir!Prioridade = TBAbrir!Prioridade + 1
                    TBAbrir.Update
                End If
                If TBAbrir!Prioridade >= Prioridade Then Prioridade = Prioridade + 1
            End If
            TBAbrir.MoveNext
        Loop
    Else
        TBAbrir.Open "select * from PCP_programacao_ordem where prioridade <> 0 and IDprogramacao = " & txtId & " order by prioridade desc", Conexao, adOpenKeyset, adLockOptimistic
        Do While TBAbrir.EOF = False
            If TBAbrir!ID <> txtIDprioridade Then
                If TBAbrir!Prioridade = Prioridade Then
                    TBAbrir!Prioridade = TBAbrir!Prioridade - 1
                    TBAbrir.Update
                End If
                If TBAbrir!Prioridade <= Prioridade Then Prioridade = Prioridade - 1
            End If
            TBAbrir.MoveNext
        Loop
    End If
End If
TBproducao.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCorrigirPrioridadeExcluir()
On Error GoTo tratar_erro

If txtPrioridade = "0" Then Exit Sub
Prioridade = txtPrioridade
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "select * from ordemservico where prioridade > " & txtPrioridade & " order by prioridade", Conexao, adOpenKeyset, adLockOptimistic
Do While TBAbrir.EOF = False
    TBAbrir!Prioridade = TBAbrir!Prioridade - 1
    TBAbrir.Update
    TBAbrir.MoveNext
Loop
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_prioridade_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView Lista_prioridade, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_prioridade_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro
  
If Lista_prioridade.ListItems.Count = 0 Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from PCP_programacao_ordem where id = " & Lista_prioridade.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    ProcLimpaCamposOrdem
    ProcPuxaDados
    CodigoLista1 = Lista_prioridade.SelectedItem.index
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error GoTo tratar_erro

If txtId = "0" Then
    SSTab1.Tab = 0
    Exit Sub
End If
Select Case SSTab1.Tab
    
    Case 1:
        ProcVerificaProsseguir
        If Permitido = False Then Exit Sub
        ProcLimpaCamposOrdem
        ProcCarregaListaOrdem
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtDataInicio_Change()
On Error GoTo tratar_erro

ProcCalculaDataFinal

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtFase_LostFocus()
On Error GoTo tratar_erro

txtOS = ""
txtGrupo = ""
txtQtde = ""
txtTempo = ""
txtInstrucoes = ""
If txtFase.Text <> "" Then
    VerifNumero = txtFase.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtFase.Text = ""
        txtFase.SetFocus
        Exit Sub
    End If
    ProcCarregaOS_Fase
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtHoraInicio_Change()
On Error GoTo tratar_erro

'ProcCalculaDataFinal

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtIDmaquina_Change()
On Error GoTo tratar_erro

If txtIDmaquina = "" Then Exit Sub
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from CadMaquinas where idmaquina = " & txtIDmaquina, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    txtdescricao = IIf(IsNull(TBAbrir!Descricao), "", TBAbrir!Descricao)
    If IsNull(TBAbrir!maquina) = False Then cmbMaquina = TBAbrir!maquina
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    If Err.Number = "383" Then
        USMsgBox ("Não foi encontrado o posto de trabalho."), vbExclamation, "CAPRIND v5.0"
        Exit Sub
    End If
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtOrdem_LostFocus()
On Error GoTo tratar_erro

ProcLimpaCamposOrdem2
If txtOrdem.Text <> "" Then
    VerifNumero = txtOrdem.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtOrdem.Text = ""
        txtOrdem.SetFocus
        Exit Sub
    End If
    ProcCarregaOrdem
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtOS_LostFocus()
On Error GoTo tratar_erro

txtFase = ""
txtGrupo = ""
txtQtde = ""
txtTempo = ""
txtInstrucoes = ""
If txtOS.Text <> "" Then
    VerifNumero = txtOS.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtOS.Text = ""
        txtOS.SetFocus
        Exit Sub
    End If
    ProcCarregaOS_Fase
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtPrioridade_LostFocus()
On Error GoTo tratar_erro

If txtPrioridade.Text <> "" Then
    VerifNumero = txtPrioridade.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtPrioridade.Text = ""
        txtPrioridade.SetFocus
        Exit Sub
    End If
Else
    txtPrioridade = "0"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCalculaDataFinal()
On Error GoTo tratar_erro

Qtde = 0 'Data e hora final
Qtd = 0 'Tempo total da OS
valor = 0
Dia = 0
Contador = 0
Contador2 = 0
Contador3 = 0
DataTotal = "00:00:00"
qt = 0
DatainicioINTER = "00:00:00"
DatafimINTER = "00:00:00"
DataTempoINTER = "00:00:00"
If txtTempo = "" Or txtTempo = "00:00:00" Or txtHoraInicio = "00:00:00" Then Exit Sub
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from CadMaquinas where idmaquina = " & txtIDmaquina, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    'Formata tempo total utilizar OS
    ProcFormataHora (txtTempo)
    Qtd = s
    
    Dataini = txtDataInicio
    ProcVerificaDia (Dataini)
    Contador = 0
    
    Do While Qtd > 0
        Set TBTempo = CreateObject("adodb.recordset")
        TBTempo.Open "Select * from CadmaqTurnos where maquina = '" & cmbMaquina & "' and Diasemana = '" & Diasemana & "' and Bloqueado = 'false'", Conexao, adOpenKeyset, adLockOptimistic
        If TBTempo.EOF = False Then
            'Verifica tempo total disponivel da maquina
            TempoExecucao = Right(TBTempo!TotalDia, 8)
            ProcFormataHora (TempoExecucao)
            valor = s
            
            ProcVerificaTurno
            
            'Verifica o tempo disponivel da máquina conforme a hora início
            ValorTotal = 0
            If Contador = 0 Then
                If Outros <> 0 Then
                    DataTotal = txtHoraInicio.Value
                    Set TBMaquinas = CreateObject("adodb.recordset")
                    TBMaquinas.Open "Select * from CadMaqturnos where maquina = '" & cmbMaquina & "' and diasemana = '" & Diasemana & "' and Turno = " & Outros & " and Bloqueado = 'false'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBMaquinas.EOF = False Then
                        TempoExecucao = IIf(IsNull(TBMaquinas!Inicioturno), 0, Left(TBMaquinas!Inicioturno, 8))
                        If DataTotal > TempoExecucao Then
                            ProcFormataHora (DataTotal)
                            ValorTotal = s 'Tempo Inicio
                            
                            TempoExecucao = Left(TBMaquinas!Inicioturno, 8)
                            ProcFormataHora (TempoExecucao)
                            ValorPagar = s 'Inicio do turno
                            If DataTotal <= IIf(IsNull(TBMaquinas!Inicio_intervalo), 0, Left(TBMaquinas!Inicio_intervalo, 8)) Then
                                ValorTotal = ValorTotal - ValorPagar
                            Else
                                ProcFormataHora (IIf(IsNull(TBMaquinas!intervalo), 0, TBMaquinas!intervalo))
                                ValorTotal = (ValorTotal - ValorPagar) - s
                            End If
                        End If
                    End If
                    TBMaquinas.Close
                End If
                'Tira diferença do tempo disponivel x hora inicio do tempo disponivel padrao
                valor = valor - ValorTotal
                Contador = 1
            End If
            
            If Qtd > valor Then
                'Soma horas do inicio preenchido até final do turno
                If Contador3 = 0 Then
                    ProcFormataHora (txtHoraInicio)
                    qt = s 'Tempo Inicio
                    ProcFormataHora (IIf(IsNull(TBTempo!finalturno), "00:00:00", Left(TBTempo!finalturno, 8)))
                    qt = s - qt  ' soma tempo inicio com Tempo final
                    
                    'Soma final do turno com inicio de turno do outro dia
                    Data_Prog = IIf(IsNull(TBTempo!finalturno), "00:00:00", Left(TBTempo!finalturno, 8))
                    
                    Dataini2 = Dataini
                    Dataini2 = Dataini2 + 1
                    ProcVerificaDia (Dataini2)
                    Contador2 = 0
                    Do While Contador2 = 0
                        ProcVerificaDia (Dataini2)
                        Set TBTempo = CreateObject("adodb.recordset")
                        TBTempo.Open "Select * from CadmaqTurnos where maquina = '" & cmbMaquina & "' and Diasemana = '" & Diasemana & "' and Bloqueado = 'false'", Conexao, adOpenKeyset, adLockOptimistic
                        If TBTempo.EOF = False Then
                            ProcCalculaTempos
                            Contador2 = 1
                        Else
                            Dataini2 = Dataini2 + 1
                            ProcVerificaDia (Dataini2)
                        End If
                    Loop
                    ProcFormataHora (TempoTotal)
                    qt = s + qt
                    Contador3 = 1
                Else
                    qt = 86400
                End If
                
                Qtde = Qtde + qt
                Qtd = Qtd - valor
            Else
                DatainicioINTER = IIf(IsNull(TBTempo!Inicio_intervalo), "00:00:00", Left(TBTempo!Inicio_intervalo, 8))
                DatafimINTER = IIf(IsNull(TBTempo!finalturno), "00:00:00", Left(TBTempo!finalturno, 8))
                DataTempoINTER = IIf(IsNull(TBTempo!intervalo), "00:00:00", Left(TBTempo!intervalo, 8))
                Qtde = Qtde + Qtd
                Qtd = Qtd - Qtd
            End If
        Else
            Qtde = Qtde + 86400
        End If
        TBTempo.Close
        Dataini = Dataini + 1
        ProcVerificaDia (Dataini)
    Loop
End If
TBAbrir.Close

Dataini = txtDataInicio & " " & txtHoraInicio
DataFim = DateAdd("S", Qtde, Dataini)
txtDataFinal = Format(DataFim, "dd/mm/yyyy")

If Format(DataFim, "hh:mm:ss") > DatainicioINTER Then
    ProcFormataHora (DataTempoINTER)
    Qtde = Qtde + s
    DataFim = DateAdd("S", Qtde, Dataini)
    If Format(DataFim, "hh:mm:ss") > DatafimINTER Then
        txtHoraFinal = Format(DatafimINTER, "hh:mm:ss")
    Else
        txtHoraFinal = Format(DataFim, "hh:mm:ss")
    End If
Else
    txtHoraFinal = Format(DataFim, "hh:mm:ss")
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCalculaTempos()
On Error GoTo tratar_erro
Dim Final       As Date 'OK
Dim TotalTurno  As Date 'OK

If Len(Data_Prog) = 8 And Len(Data_Prog) = 8 Then
    If Data_Prog <> "00:00:00" And Data_Prog <> "00:00:00" Then
        Inicio = Data_Prog
        Final = IIf(IsNull(TBTempo!Inicioturno), "00:00:00", Left(TBTempo!Inicioturno, 8))
        If Final > Inicio Then
            TempoTotal = Final - Inicio
        Else
            Final = Final + 1
            TempoTotal = Final - Inicio
        End If
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcVerificaDia(Data1 As Date)
On Error GoTo tratar_erro

Diasemana = Weekday(Data)
Select Case Diasemana
    Case 1: Diasemana = "Domingo"
    Case 2: Diasemana = "Segunda"
    Case 3: Diasemana = "Terça"
    Case 4: Diasemana = "Quarta"
    Case 5: Diasemana = "Quinta"
    Case 6: Diasemana = "Sexta"
    Case 7: Diasemana = "Sabado"
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcVerificaTurno()
On Error GoTo tratar_erro

Outros = 0
TempoInicio = 0
TempoFinal = 0
ProcVerificaDia (txtDataInicio.Value)
DataTotal = txtHoraInicio.Value

Set TBMaquinas = CreateObject("adodb.recordset")
TBMaquinas.Open "Select * from CadMaqturnos where maquina = '" & cmbMaquina & "' and diasemana = '" & Diasemana & "' and Bloqueado = 'false' order by diasemana,turno", Conexao, adOpenKeyset, adLockOptimistic
If TBMaquinas.EOF = False Then
    Do While TBMaquinas.EOF = False
        If IsNull(TBMaquinas!Inicioturno) = False Then
            TempoInicio = Left(TBMaquinas!Inicioturno, 8)
            TempoFinal = Left(TBMaquinas!finalturno, 8)
            If TempoInicio > TempoFinal Then
                DataTotal = txtDataInicio.Value & " " & DataTotal
                TempoInicio = txtDataInicio.Value & " " & TempoInicio
                TempoFinal = txtDataInicio.Value & "  " & TempoFinal
                TempoInicio = TempoInicio - 1
                TempoFinal = TempoFinal + 1
            End If
            Select Case TBMaquinas!Turno
                Case 1:
                    If DataTotal >= TempoInicio And DataTotal <= TempoFinal Then
                        Outros = 1
                        GoTo Sair
                    End If
                Case 2:
                    If DataTotal >= TempoInicio And DataTotal <= TempoFinal Then
                        Outros = 2
                        GoTo Sair
                    End If
                Case 3:
                    If DataTotal >= TempoInicio And DataTotal <= TempoFinal Then
                        Outros = 3
                        GoTo Sair
                    End If
                Case 4:
                    If DataTotal >= TempoInicio And DataTotal <= TempoFinal Then
                        Outros = 4
                        GoTo Sair
                    End If
            End Select
        End If
        TBMaquinas.MoveNext
    Loop
End If
Sair:
    TBMaquinas.Close
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaMaquina()
On Error GoTo tratar_erro

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from CadMaquinas", Conexao, adOpenKeyset, adLockOptimistic
Do While TBAbrir.EOF = False
    cmbMaquina.AddItem Trim(TBAbrir!maquina)
    cmbMaquina.ItemData(cmbMaquina.NewIndex) = TBAbrir!IDMaquina
    TBAbrir.MoveNext
Loop
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaListaOrdem()
On Error GoTo tratar_erro

Dim DataInicio As String
Dim DataFim As String

DataInicio = txtDataInicio & " 00:00:01.000"
DataFinal = txtDataFinal & " 23:59:59.000"

Lista_prioridade.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from PCP_programacao_ordem where IDprogramacao = " & txtId & " and Inicio >= '" & Format(DataInicio, "Long Date") & "' and Final <= '" & Format(DataFinal, "Long Date") & "' order by prioridade, ordem, OS", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    TBLISTA.MoveLast
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    TBLISTA.MoveFirst
    Do While TBLISTA.EOF = False
        With Lista_prioridade.ListItems
            .Add , , IIf(IsNull(TBLISTA!ID), 0, TBLISTA!ID)
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Inicio), "", Format(TBLISTA!Inicio, "dd/mm/yy hh:mm:ss"))
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Final), "", Format(TBLISTA!Final, "dd/mm/yy hh:mm:ss"))
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Prioridade), "", TBLISTA!Prioridade)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!Ordem), "", TBLISTA!Ordem)
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "select * from producao where ordem = " & TBLISTA!Ordem, Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                .Item(.Count).SubItems(5) = IIf(IsNull(TBAbrir!Desenho), "", TBAbrir!Desenho)
                .Item(.Count).SubItems(6) = IIf(IsNull(TBAbrir!Revitem), "", TBAbrir!Revitem)
                .Item(.Count).SubItems(7) = IIf(IsNull(TBAbrir!Produto), "", TBAbrir!Produto)
            End If
            TBAbrir.Close
            .Item(.Count).SubItems(8) = IIf(IsNull(TBLISTA!OS), "", TBLISTA!OS)
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "select * from ordemservico where IDProducao = " & TBLISTA!OS, Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                .Item(.Count).SubItems(9) = IIf(IsNull(TBAbrir!Fase), "", TBAbrir!Fase)
                .Item(.Count).SubItems(10) = IIf(IsNull(TBAbrir!Grupo_op), "", TBAbrir!Grupo_op)
            End If
            TBAbrir.Close
            
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

Sub ProcPuxaDados()
On Error GoTo tratar_erro

txtIDprioridade = TBLISTA!ID
txtPrioridade = IIf(IsNull(TBLISTA!Prioridade), "", TBLISTA!Prioridade)
txtOrdem = IIf(IsNull(TBLISTA!Ordem), "", TBLISTA!Ordem)
ProcCarregaOrdem
txtOS = IIf(IsNull(TBLISTA!OS), "", TBLISTA!OS)
ProcCarregaOS_Fase
txtQtde.Text = IIf(IsNull(TBLISTA!Qtd), "", Format(TBLISTA!Qtd, "###,##0.0000"))
txtTempo.Text = IIf(IsNull(TBLISTA!TempoTotal), "00:00:00", Format(TBLISTA!TempoTotal, "hh:mm:ss"))
txtDataInicio = IIf(IsNull(TBLISTA!Inicio), Date, Format(TBLISTA!Inicio, "dd/mm/yyyy"))
txtHoraInicio = IIf(IsNull(TBLISTA!Inicio), Date, Format(TBLISTA!Inicio, "hh:mm:ss"))
txtDataFinal = IIf(IsNull(TBLISTA!Final), Date, Format(TBLISTA!Final, "dd/mm/yyyy"))
txtHoraFinal = IIf(IsNull(TBLISTA!Final), Date, Format(TBLISTA!Final, "hh:mm:ss"))
Frame2.Enabled = True
Novo_ordem_programacao1 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCamposOrdem2()
On Error GoTo tratar_erro

txtOS = ""
txtFase = ""
txtGrupo = ""
txtQtde = ""
txtTempo = ""
txtDescricao_item = ""
txtdesenho = ""
txtRevitem = ""
txtreferencia = ""
txtInstrucoes = ""
txtCliente = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaOS_Fase()
On Error GoTo tratar_erro

If txtOS = "" And txtFase = "" Then Exit Sub
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Ordemservico where IDProducao = " & txtOS, Conexao, adOpenKeyset, adLockOptimistic

If TBAbrir.EOF = False Then
    txtFase = IIf(IsNull(TBAbrir!Fase), "", TBAbrir!Fase)
    txtGrupo = IIf(IsNull(TBAbrir!Grupo_op), "", TBAbrir!Grupo_op)
    Qtde = 0
    VlttTotal = 0
    Qtde = IIf(IsNull(TBAbrir!Totalprod), 0, TBAbrir!Totalprod)
    VlttTotal = IIf(IsNull(TBAbrir!TETTUTILSEG), 0, TBAbrir!TETTUTILSEG)

    txtQtde = IIf(IsNull(TBAbrir!quantidade), "", Format(TBAbrir!quantidade - Qtde, "###,##0.0000"))
    ProcFormataHora (IIf(IsNull(TBAbrir!TempoTotalLote), "", TBAbrir!TempoTotalLote))
    VlttTotal = s + VlttTotal
    
    'Rotina de tranformacao de preparacao
    If IsNull(TBAbrir!Preparacao) = False And TBAbrir!Preparacao <> "__:__:__" Then
        ElapsedTime (TBAbrir!Preparacao)
        TempoPreparacao = HoraTotal
        If Len(TempoPreparacao) = 8 Then
            TempoPreparacao = "0" & TempoPreparacao
        End If
    End If

    'Rotina de tranformacao de execucao
    If IsNull(TBAbrir!Execucao) = False And TBAbrir!Execucao <> "__:__:__" Then
        ElapsedTime (TBAbrir!Execucao)
        TempoExecucao = HoraTotal
        If Len(TempoExecucao) = 8 Then
            TempoExecucao = "0" & TempoExecucao
        End If
    End If
    
    txtpreparacao = TempoPreparacao
    txtexecucao = TempoExecucao
    txtTempo = FormataTempo(VlttTotal)
    txtInstrucoes = IIf(IsNull(TBAbrir!descfase), "", TBAbrir!descfase)
    
End If

Set TBOrdem = CreateObject("adodb.recordset")
TBOrdem.Open "Select * from producao where ordem = " & TBAbrir!Ordem & "", Conexao, adOpenKeyset, adLockOptimistic
If TBOrdem.EOF = False Then
    txtOrdem = IIf(IsNull(TBOrdem!Ordem), "", TBOrdem!Ordem)
    txtdesenho = IIf(IsNull(TBOrdem!Desenho), "", TBOrdem!Desenho)
    txtRevitem = IIf(IsNull(TBOrdem!Revitem), "", TBOrdem!Revitem)
    txtreferencia = IIf(IsNull(TBOrdem!N_referencia), "", TBOrdem!N_referencia)
    txtDescricao_item = IIf(IsNull(TBOrdem!Produto), "", TBOrdem!Produto)
    txtCliente = IIf(IsNull(TBOrdem!Cliente), "", TBOrdem!Cliente)
End If
TBAbrir.Close
TBOrdem.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaOrdem()
On Error GoTo tratar_erro


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcVerificaProsseguir()
On Error GoTo tratar_erro

Permitido = True
If Novo_Proposta = True Then
    USMsgBox ("Salve o posto de trabalho antes de prosseguir."), vbExclamation, "CAPRIND v5.0"
    SSTab1.Tab = 0
    CmdSalvar.SetFocus
    Permitido = False
    Exit Sub
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcImprimir()
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 Then Exit Sub
FiltroImpressao = True
frmProd_programacao_imprimir.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcImprimir2()
On Error GoTo tratar_erro

If txtId = 0 Then
    USMsgBox ("Informe o posto de trabalho antes de visualizar impressão"), vbExclamation, "CAPRIND v5.0"
    cmdAbrir.SetFocus
    Exit Sub
Else
    FiltroImpressao = False
    frmProd_programacao_imprimir.Show 1
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
    Case 2: ProcAbrir
    Case 3: ProcSalvar
    Case 4: ProcExcluir
    Case 5: ProcImprimir2
    Case 7:
    Case 8: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar2_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcNovo_ordem
    Case 2: ProcSalvar_ordem
    Case 3: ProcExcluir_ordem
    Case 4: ProcImprimir
    Case 7: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub

End Sub
