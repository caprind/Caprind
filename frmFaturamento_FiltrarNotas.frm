VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Begin VB.Form frmFaturamento_FiltrarNotas 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Faturamento | Filtrar notas emitidas no período"
   ClientHeight    =   5850
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7485
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5850
   ScaleWidth      =   7485
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Opções para pesquisa"
      Height          =   4515
      Index           =   0
      Left            =   240
      TabIndex        =   11
      Top             =   630
      Width           =   7095
      Begin VB.Frame FramePeriodo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Escolha o período"
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
         Height          =   765
         Left            =   4320
         TabIndex        =   15
         Top             =   1740
         Width           =   2655
         Begin MSComCtl2.DTPicker msk_fltFim 
            Height          =   315
            Left            =   1500
            TabIndex        =   6
            ToolTipText     =   "Data final."
            Top             =   270
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
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
            Format          =   198639617
            CurrentDate     =   39057
         End
         Begin MSComCtl2.DTPicker msk_fltInicio 
            Height          =   315
            Left            =   120
            TabIndex        =   5
            ToolTipText     =   "Data inicio."
            Top             =   270
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
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
            Format          =   198639617
            CurrentDate     =   39057
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "até"
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
            Left            =   1215
            TabIndex        =   16
            Top             =   330
            Width           =   240
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Escolha o Status"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   1140
         TabIndex        =   37
         Top             =   1740
         Width           =   1845
         Begin VB.ComboBox cmbStatus 
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
            ItemData        =   "frmFaturamento_FiltrarNotas.frx":0000
            Left            =   540
            List            =   "frmFaturamento_FiltrarNotas.frx":000A
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   38
            ToolTipText     =   "Opções para filtro."
            Top             =   300
            Width           =   1245
         End
         Begin DrawSuite2022.USLabel USLabel3 
            Height          =   195
            Left            =   60
            TabIndex        =   40
            Top             =   390
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   344
            Caption         =   "Status"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483630
            NoHTMLCaption   =   "Status"
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Escolha a operação"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   150
         TabIndex        =   31
         Top             =   1080
         Width           =   6825
         Begin DrawSuite2022.USOptionButton optVendas 
            Height          =   285
            Left            =   1080
            TabIndex        =   32
            Top             =   300
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   503
            Caption         =   "Vendas"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   128
            ShowFocusRect   =   0   'False
         End
         Begin DrawSuite2022.USOptionButton opttodas 
            Height          =   285
            Left            =   300
            TabIndex        =   33
            Top             =   300
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   503
            Caption         =   "Todas"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   128
            ShowFocusRect   =   0   'False
         End
         Begin DrawSuite2022.USOptionButton optIndustrializacao 
            Height          =   285
            Left            =   1950
            TabIndex        =   34
            Top             =   300
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            Caption         =   "Industrialização"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   128
            ShowFocusRect   =   0   'False
         End
         Begin DrawSuite2022.USOptionButton optRemessa 
            Height          =   285
            Left            =   3360
            TabIndex        =   35
            Top             =   300
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   503
            Caption         =   "Remessa"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   128
            ShowFocusRect   =   0   'False
         End
         Begin DrawSuite2022.USOptionButton optRetorno 
            Height          =   285
            Left            =   4530
            TabIndex        =   36
            Top             =   300
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   503
            Caption         =   "Retorno"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   128
            ShowFocusRect   =   0   'False
         End
         Begin DrawSuite2022.USOptionButton optDevolucao 
            Height          =   285
            Left            =   5610
            TabIndex        =   39
            Top             =   300
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   503
            Caption         =   "Devolução"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   128
            ShowFocusRect   =   0   'False
         End
      End
      Begin VB.Frame FrameAno 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Escolha o ano"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   4320
         TabIndex        =   30
         Top             =   1740
         Visible         =   0   'False
         Width           =   2655
         Begin DrawSuite2022.USLabel USLabel1 
            Height          =   195
            Left            =   120
            TabIndex        =   41
            Top             =   390
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   344
            Caption         =   "Ano"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483630
            NoHTMLCaption   =   "Ano"
         End
         Begin VB.ComboBox cmbAno 
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
            ItemData        =   "frmFaturamento_FiltrarNotas.frx":0023
            Left            =   510
            List            =   "frmFaturamento_FiltrarNotas.frx":0057
            Sorted          =   -1  'True
            TabIndex        =   8
            ToolTipText     =   "Opções para filtro."
            Top             =   300
            Width           =   1125
         End
      End
      Begin DrawSuite2022.USOptionButton optMes 
         Height          =   285
         Left            =   4650
         TabIndex        =   7
         Top             =   690
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   503
         Caption         =   "Por mês"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocusRect   =   0   'False
      End
      Begin DrawSuite2022.USOptionButton optPeriodo 
         Height          =   285
         Left            =   4650
         TabIndex        =   4
         Top             =   420
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         Caption         =   "Por período      "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocusRect   =   0   'False
      End
      Begin VB.Frame FrameMes 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Escolha o mês"
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
         Height          =   945
         Left            =   150
         TabIndex        =   17
         Top             =   2580
         Width           =   6825
         Begin DrawSuite2022.USOptionButton opt1 
            Height          =   285
            Left            =   330
            TabIndex        =   18
            Top             =   330
            Width           =   825
            _ExtentX        =   1455
            _ExtentY        =   503
            Caption         =   "Janeiro"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   128
            ShowFocusRect   =   0   'False
         End
         Begin DrawSuite2022.USOptionButton opt2 
            Height          =   285
            Left            =   1260
            TabIndex        =   19
            Top             =   330
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            Caption         =   "Fevereiro"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   128
            ShowFocusRect   =   0   'False
         End
         Begin DrawSuite2022.USOptionButton opt3 
            Height          =   285
            Left            =   2340
            TabIndex        =   20
            Top             =   330
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            Caption         =   "Março"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   128
            ShowFocusRect   =   0   'False
         End
         Begin DrawSuite2022.USOptionButton opt4 
            Height          =   285
            Left            =   3420
            TabIndex        =   21
            Top             =   330
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   503
            Caption         =   "Abril"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   128
            ShowFocusRect   =   0   'False
         End
         Begin DrawSuite2022.USOptionButton opt5 
            Height          =   285
            Left            =   4470
            TabIndex        =   22
            Top             =   330
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   503
            Caption         =   "Maio"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   128
            ShowFocusRect   =   0   'False
         End
         Begin DrawSuite2022.USOptionButton opt6 
            Height          =   285
            Left            =   5610
            TabIndex        =   23
            Top             =   330
            Width           =   825
            _ExtentX        =   1455
            _ExtentY        =   503
            Caption         =   "Junho"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   128
            ShowFocusRect   =   0   'False
         End
         Begin DrawSuite2022.USOptionButton opt7 
            Height          =   285
            Left            =   330
            TabIndex        =   24
            Top             =   570
            Width           =   825
            _ExtentX        =   1455
            _ExtentY        =   503
            Caption         =   "Julho"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   128
            ShowFocusRect   =   0   'False
         End
         Begin DrawSuite2022.USOptionButton opt8 
            Height          =   285
            Left            =   1260
            TabIndex        =   25
            Top             =   570
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            Caption         =   "Agosto"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   128
            ShowFocusRect   =   0   'False
         End
         Begin DrawSuite2022.USOptionButton opt9 
            Height          =   285
            Left            =   2340
            TabIndex        =   26
            Top             =   570
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            Caption         =   "Setembro"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   128
            ShowFocusRect   =   0   'False
         End
         Begin DrawSuite2022.USOptionButton opt10 
            Height          =   285
            Left            =   3420
            TabIndex        =   27
            Top             =   570
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   503
            Caption         =   "Outubro"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   128
            ShowFocusRect   =   0   'False
         End
         Begin DrawSuite2022.USOptionButton opt11 
            Height          =   285
            Left            =   4470
            TabIndex        =   28
            Top             =   570
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   503
            Caption         =   "Novembro"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   128
            ShowFocusRect   =   0   'False
         End
         Begin DrawSuite2022.USOptionButton opt12 
            Height          =   285
            Left            =   5610
            TabIndex        =   29
            Top             =   570
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   503
            Caption         =   "Dezembro"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   128
            ShowFocusRect   =   0   'False
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Escolha a série"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   3000
         TabIndex        =   14
         Top             =   1740
         Width           =   1305
         Begin DrawSuite2022.USLabel USLabel2 
            Height          =   195
            Left            =   120
            TabIndex        =   42
            Top             =   390
            Width           =   435
            _ExtentX        =   767
            _ExtentY        =   344
            Caption         =   "Série"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483630
            NoHTMLCaption   =   "Série"
         End
         Begin VB.ComboBox cmbSerie 
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
            ItemData        =   "frmFaturamento_FiltrarNotas.frx":00BB
            Left            =   510
            List            =   "frmFaturamento_FiltrarNotas.frx":00CB
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   3
            ToolTipText     =   "Opções para filtro."
            Top             =   300
            Width           =   735
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Tipo nota"
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
         Height          =   765
         Left            =   180
         TabIndex        =   13
         Top             =   1740
         Width           =   945
         Begin DrawSuite2022.USCheckBox chkSaida 
            Height          =   255
            Left            =   150
            TabIndex        =   2
            Top             =   330
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   450
            Caption         =   "Saida"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   8388608
            ShowFocusRect   =   0   'False
            Value           =   1
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Empresa emitente"
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
         Height          =   720
         Index           =   1
         Left            =   150
         TabIndex        =   12
         Top             =   300
         Width           =   4410
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
            ItemData        =   "frmFaturamento_FiltrarNotas.frx":00DB
            Left            =   120
            List            =   "frmFaturamento_FiltrarNotas.frx":00DD
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   0
            ToolTipText     =   "Empresa."
            Top             =   270
            Width           =   4125
         End
      End
      Begin DrawSuite2022.USButton btnFiltrar 
         Height          =   615
         Left            =   4050
         TabIndex        =   9
         Top             =   3720
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   1085
         DibPicture      =   "frmFaturamento_FiltrarNotas.frx":00DF
         Caption         =   "   Filtrar notas fiscais"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
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
         ShowFocusRect   =   0   'False
         Theme           =   4
      End
   End
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   10
      Top             =   5445
      Width           =   7485
      _ExtentX        =   13203
      _ExtentY        =   714
   End
   Begin DrawSuite2022.USForm frmFaturamento_Fltrar 
      Align           =   1  'Align Top
      Height          =   435
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7485
      _ExtentX        =   13203
      _ExtentY        =   767
      DibPicture      =   "frmFaturamento_FiltrarNotas.frx":372F
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
      Icon            =   "frmFaturamento_FiltrarNotas.frx":6D7F
   End
End
Attribute VB_Name = "frmFaturamento_FiltrarNotas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnFiltrar_Click()
On Error GoTo tratar_erro
Dim FormulaRelatorio As String

ID_empresa = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)

SerieNF = cmbSerie.Text


If optVendas.Value = True Then
TipoNF = 1
NomeRel = "Faturamento2.rpt"
End If

If optIndustrializacao.Value = True Then
TipoNF = 1
NomeRel = "Faturamento2.rpt"
End If


If opttodas.Value = True Then
NomeRel = "Faturamento_Notas.rpt"
End If

If optDevolucao.Value = True Then
TipoNF = 1
NomeRel = "Faturamento2.rpt"
End If

If optRetorno.Value = True Then
TipoNF = 1
NomeRel = "Faturamento2.rpt"
End If


If cmbStatus.Text = "Cancelada" Then
 IntStatus = 2
Else
 IntStatus = 1
End If


DiaINI = Day(msk_fltInicio.Value)
MesINI = Month(msk_fltInicio.Value)
AnoINI = Year(msk_fltInicio.Value)

DiaFim = Day(msk_fltFim.Value)
MesFim = Month(msk_fltFim.Value)
AnoFim = Year(msk_fltFim.Value)

varinicio = "Date(" & AnoINI & "," & MesINI & "," & DiaINI & ")"
varfim = "Date(" & AnoFim & "," & MesFim & "," & DiaFim & ")"

If optPeriodo.Value = False And optMes.Value = True Then
If Opt1.Value = True Then
Mes = 1
End If

If Opt2.Value = True Then
Mes = 2
End If

If opt3.Value = True Then
Mes = 3
End If

If Opt4.Value = True Then
Mes = 4
End If

If Opt5.Value = True Then
Mes = 5
End If

If opt6.Value = True Then
Mes = 6
End If

If opt7.Value = True Then
Mes = 7
End If

If opt8.Value = True Then
Mes = 8
End If

If opt9.Value = True Then
Mes = 9
End If

If opt10.Value = True Then
Mes = 10
End If

If opt11.Value = True Then
Mes = 11
End If

If opt12.Value = True Then
Mes = 12
End If

Ano = cmbAno.Text
End If

'=======================================================
' Todas da Notas por série e período
'=======================================================
If opttodas.Value = True Then
   'Filtrar por periodo com serie
   If optPeriodo.Value = True And cmbSerie.Text <> "" Then
   FormulaRelatorio = "{Faturamento_NotasFiscais.serie} = '" & SerieNF & "' and {Faturamento_NotasFiscais.tiponf} = 'M1' and {Faturamento_NotasFiscais.Aplicacao} = 'P' and {Faturamento_NotasFiscais.ID_empresa} = " & ID_empresa & " and {Faturamento_NotasFiscais.dt_DataEmissao} >= " & varinicio & " And {Faturamento_NotasFiscais.dt_DataEmissao}<= " & varfim & " and {Faturamento_NotasFiscais.int_status} = " & IntStatus & ""
   End If
   
   'Filtrar por mês com série
   If optMes.Value = True And cmbSerie.Text <> "" Then
   FormulaRelatorio = "{Faturamento_NotasFiscais.serie} = '" & SerieNF & "' and {Faturamento_NotasFiscais.tiponf} = 'M1' and {Faturamento_NotasFiscais.Aplicacao} = 'P' and {Faturamento_NotasFiscais.ID_empresa} = " & ID_empresa & " and Month({Faturamento_NotasFiscais.dt_DataEmissao}) = " & Mes & " And Year({Faturamento_NotasFiscais.dt_DataEmissao}) = " & Ano & " and {Faturamento_NotasFiscais.int_status} = " & IntStatus & ""
   End If
   
   'Filtrar por período todas as Series
   If optPeriodo.Value = True And cmbSerie.Text = "" Then
   FormulaRelatorio = "{Faturamento_NotasFiscais.tiponf} = 'M1' and {Faturamento_NotasFiscais.Aplicacao} = 'P' and {Faturamento_NotasFiscais.ID_empresa} = " & ID_empresa & " and {Faturamento_NotasFiscais.dt_DataEmissao} >= " & varinicio & " And {Faturamento_NotasFiscais.dt_DataEmissao}<= " & varfim & " and {Faturamento_NotasFiscais.int_status} = " & IntStatus & ""
   End If
   
   'Filtrar por mês todas as Series
   If optMes.Value = True And cmbSerie.Text = "" Then
   FormulaRelatorio = "{Faturamento_NotasFiscais.tiponf} = 'M1' and {Faturamento_NotasFiscais.Aplicacao} = 'P' and {Faturamento_NotasFiscais.ID_empresa} = " & ID_empresa & " and Month({Faturamento_NotasFiscais.dt_DataEmissao}) = " & Mes & " And Year({Faturamento_NotasFiscais.dt_DataEmissao}) = " & Ano & " and {Faturamento_NotasFiscais.int_status} = " & IntStatus & ""
   End If
   
End If

'=======================================================
' Notas de vendas
'=======================================================
If optVendas.Value = True Then
   'Filtrar por periodo com serie
   If optPeriodo.Value = True And cmbSerie.Text <> "" Then
   FormulaRelatorio = "{Faturamento_Serie.Vendas}=True and {Faturamento_Serie.serie} = '" & SerieNF & "' and {Faturamento_Serie.tiponf} = 'M1' and {Faturamento_Serie.Aplicacao} = 'P' and {Faturamento_Serie.ID_empresa} = " & ID_empresa & " and {Faturamento_Serie.dt_DataEmissao} >= " & varinicio & " And {Faturamento_Serie.dt_DataEmissao}<= " & varfim & " And {Faturamento_Serie.Int_TipoNota} = " & TipoNF & ""
   End If
   
   'Filtrar por mês com série
   If optMes.Value = True And cmbSerie.Text <> "" Then
   FormulaRelatorio = "{Faturamento_Serie.Vendas}=True and {Faturamento_Serie.serie} = '" & SerieNF & "' and {Faturamento_Serie.tiponf} = 'M1' and {Faturamento_Serie.Aplicacao} = 'P' and {Faturamento_Serie.ID_empresa} = " & ID_empresa & " and Month({Faturamento_Serie.dt_DataEmissao}) = " & Mes & " And Year({Faturamento_Serie.dt_DataEmissao}) = " & Ano & " And {Faturamento_Serie.Int_TipoNota} = " & TipoNF & ""
   End If
   
   'Filtrar por período todas as Series
   If optPeriodo.Value = True And cmbSerie.Text = "" Then
   FormulaRelatorio = "{Faturamento_Serie.Vendas}=True and {Faturamento_Serie.tiponf} = 'M1' and {Faturamento_Serie.Aplicacao} = 'P' and {Faturamento_Serie.ID_empresa} = " & ID_empresa & " and {Faturamento_Serie.dt_DataEmissao} >= " & varinicio & " And {Faturamento_Serie.dt_DataEmissao}<= " & varfim & " And {Faturamento_Serie.Int_TipoNota} = " & TipoNF & ""
   End If
   
   'Filtrar por mês todas as Series
   If optMes.Value = True And cmbSerie.Text = "" Then
   FormulaRelatorio = "{Faturamento_Serie.Vendas}=True and {Faturamento_Serie.tiponf} = 'M1' and {Faturamento_Serie.Aplicacao} = 'P' and {Faturamento_Serie.ID_empresa} = " & ID_empresa & " and Month({Faturamento_Serie.dt_DataEmissao}) = " & Mes & " And Year({Faturamento_Serie.dt_DataEmissao}) = " & Ano & " And {Faturamento_Serie.Int_TipoNota} = " & TipoNF & ""
   End If
End If

'=======================================================
' Notas de industrialização
'=======================================================
If optIndustrializacao.Value = True Then
   'Filtrar por periodo com serie
   If optPeriodo.Value = True And cmbSerie.Text <> "" Then
   FormulaRelatorio = "{Faturamento_Serie.MaoObra}=True and {Faturamento_Serie.serie} = '" & SerieNF & "' and {Faturamento_Serie.tiponf} = 'M1' and {Faturamento_Serie.Aplicacao} = 'P' and {Faturamento_Serie.ID_empresa} = " & ID_empresa & " and {Faturamento_Serie.dt_DataEmissao} >= " & varinicio & " And {Faturamento_Serie.dt_DataEmissao}<= " & varfim & " And {Faturamento_Serie.Int_TipoNota} = " & TipoNF & ""
   End If
   
   'Filtrar por mês com série
   If optMes.Value = True And cmbSerie.Text <> "" Then
   FormulaRelatorio = "{Faturamento_Serie.MaoObra}=True and {Faturamento_Serie.serie} = '" & SerieNF & "' and {Faturamento_Serie.tiponf} = 'M1' and {Faturamento_Serie.Aplicacao} = 'P' and {Faturamento_Serie.ID_empresa} = " & ID_empresa & " and Month({Faturamento_Serie.dt_DataEmissao}) = " & Mes & " And Year({Faturamento_Serie.dt_DataEmissao}) = " & Ano & " And {Faturamento_Serie.Int_TipoNota} = " & TipoNF & ""
   End If
   
   'Filtrar por período todas as Series
   If optPeriodo.Value = True And cmbSerie.Text = "" Then
   FormulaRelatorio = "{Faturamento_Serie.MaoObra}=True and {Faturamento_Serie.tiponf} = 'M1' and {Faturamento_Serie.Aplicacao} = 'P' and {Faturamento_Serie.ID_empresa} = " & ID_empresa & " and {Faturamento_Serie.dt_DataEmissao} >= " & varinicio & " And {Faturamento_Serie.dt_DataEmissao}<= " & varfim & " And {Faturamento_Serie.Int_TipoNota} = " & TipoNF & ""
   End If
   
   'Filtrar por mês todas as Series
   If optMes.Value = True And cmbSerie.Text = "" Then
   FormulaRelatorio = "{Faturamento_Serie.MaoObra}=True and {Faturamento_Serie.tiponf} = 'M1' and {Faturamento_Serie.Aplicacao} = 'P' and {Faturamento_Serie.ID_empresa} = " & ID_empresa & " and Month({Faturamento_Serie.dt_DataEmissao}) = " & Mes & " And Year({Faturamento_Serie.dt_DataEmissao}) = " & Ano & " And {Faturamento_Serie.Int_TipoNota} = " & TipoNF & ""
   End If
End If

'=======================================================
' Notas de devolucao
'=======================================================
If optDevolucao.Value = True Then
   'Filtrar por periodo com serie
   If optPeriodo.Value = True And cmbSerie.Text <> "" Then
   FormulaRelatorio = "{Faturamento_Serie.devolucao}=True and {Faturamento_Serie.serie} = '" & SerieNF & "' and {Faturamento_Serie.tiponf} = 'M1' and {Faturamento_Serie.Aplicacao} = 'P' and {Faturamento_Serie.ID_empresa} = " & ID_empresa & " and {Faturamento_Serie.dt_DataEmissao} >= " & varinicio & " And {Faturamento_Serie.dt_DataEmissao}<= " & varfim & " And {Faturamento_Serie.Int_TipoNota} = " & TipoNF & ""
   End If
   
   'Filtrar por mês com série
   If optMes.Value = True And cmbSerie.Text <> "" Then
   FormulaRelatorio = "{Faturamento_Serie.devolucao}=True and {Faturamento_Serie.serie} = '" & SerieNF & "' and {Faturamento_Serie.tiponf} = 'M1' and {Faturamento_Serie.Aplicacao} = 'P' and {Faturamento_Serie.ID_empresa} = " & ID_empresa & " and Month({Faturamento_Serie.dt_DataEmissao}) = " & Mes & " And Year({Faturamento_Serie.dt_DataEmissao}) = " & Ano & " And {Faturamento_Serie.Int_TipoNota} = " & TipoNF & ""
   End If
   
   'Filtrar por período todas as Series
   If optPeriodo.Value = True And cmbSerie.Text = "" Then
   FormulaRelatorio = "{Faturamento_Serie.devolucao}=True and {Faturamento_Serie.tiponf} = 'M1' and {Faturamento_Serie.Aplicacao} = 'P' and {Faturamento_Serie.ID_empresa} = " & ID_empresa & " and {Faturamento_Serie.dt_DataEmissao} >= " & varinicio & " And {Faturamento_Serie.dt_DataEmissao}<= " & varfim & " And {Faturamento_Serie.Int_TipoNota} = " & TipoNF & ""
   End If
   
   'Filtrar por mês todas as Series
   If optMes.Value = True And cmbSerie.Text = "" Then
   FormulaRelatorio = "{Faturamento_Serie.devolucao}=True and {Faturamento_Serie.tiponf} = 'M1' and {Faturamento_Serie.Aplicacao} = 'P' and {Faturamento_Serie.ID_empresa} = " & ID_empresa & " and Month({Faturamento_Serie.dt_DataEmissao}) = " & Mes & " And Year({Faturamento_Serie.dt_DataEmissao}) = " & Ano & " And {Faturamento_Serie.Int_TipoNota} = " & TipoNF & ""
   End If
End If

'=======================================================
' Notas de Retorno
'=======================================================
If optRetorno.Value = True Then
   'Filtrar por periodo com serie
   If optPeriodo.Value = True And cmbSerie.Text <> "" Then
   FormulaRelatorio = "{Faturamento_Serie.Retorno}=True and {Faturamento_Serie.serie} = '" & SerieNF & "' and {Faturamento_Serie.tiponf} = 'M1' and {Faturamento_Serie.Aplicacao} = 'P' and {Faturamento_Serie.ID_empresa} = " & ID_empresa & " and {Faturamento_Serie.dt_DataEmissao} >= " & varinicio & " And {Faturamento_Serie.dt_DataEmissao}<= " & varfim & " And {Faturamento_Serie.Int_TipoNota} = " & TipoNF & ""
   End If
   
   'Filtrar por mês com série
   If optMes.Value = True And cmbSerie.Text <> "" Then
   FormulaRelatorio = "{Faturamento_Serie.Retorno}=True and {Faturamento_Serie.serie} = '" & SerieNF & "' and {Faturamento_Serie.tiponf} = 'M1' and {Faturamento_Serie.Aplicacao} = 'P' and {Faturamento_Serie.ID_empresa} = " & ID_empresa & " and Month({Faturamento_Serie.dt_DataEmissao}) = " & Mes & " And Year({Faturamento_Serie.dt_DataEmissao}) = " & Ano & " And {Faturamento_Serie.Int_TipoNota} = " & TipoNF & ""
   End If
   
   'Filtrar por período todas as Series
   If optPeriodo.Value = True And cmbSerie.Text = "" Then
   FormulaRelatorio = "{Faturamento_Serie.Retorno}=True and {Faturamento_Serie.tiponf} = 'M1' and {Faturamento_Serie.Aplicacao} = 'P' and {Faturamento_Serie.ID_empresa} = " & ID_empresa & " and {Faturamento_Serie.dt_DataEmissao} >= " & varinicio & " And {Faturamento_Serie.dt_DataEmissao}<= " & varfim & " And {Faturamento_Serie.Int_TipoNota} = " & TipoNF & ""
   End If
   
   'Filtrar por mês todas as Series
   If optMes.Value = True And cmbSerie.Text = "" Then
   FormulaRelatorio = "{Faturamento_Serie.Retorno}=True and {Faturamento_Serie.tiponf} = 'M1' and {Faturamento_Serie.Aplicacao} = 'P' and {Faturamento_Serie.ID_empresa} = " & ID_empresa & " and Month({Faturamento_Serie.dt_DataEmissao}) = " & Mes & " And Year({Faturamento_Serie.dt_DataEmissao}) = " & Ano & " And {Faturamento_Serie.Int_TipoNota} = " & TipoNF & ""
   End If
End If

'=======================================================
' Notas de Remessa
'=======================================================
If optRemessa.Value = True Then
   'Filtrar por periodo com serie
   If optPeriodo.Value = True And cmbSerie.Text <> "" Then
   FormulaRelatorio = "{Faturamento_Serie.Remessa}=True and {Faturamento_Serie.serie} = '" & SerieNF & "' and {Faturamento_Serie.tiponf} = 'M1' and {Faturamento_Serie.Aplicacao} = 'P' and {Faturamento_Serie.ID_empresa} = " & ID_empresa & " and {Faturamento_Serie.dt_DataEmissao} >= " & varinicio & " And {Faturamento_Serie.dt_DataEmissao}<= " & varfim & " And {Faturamento_Serie.Int_TipoNota} = " & TipoNF & ""
   End If
   
   'Filtrar por mês com série
   If optMes.Value = True And cmbSerie.Text <> "" Then
   FormulaRelatorio = "{Faturamento_Serie.Remessa}=True and {Faturamento_Serie.serie} = '" & SerieNF & "' and {Faturamento_Serie.tiponf} = 'M1' and {Faturamento_Serie.Aplicacao} = 'P' and {Faturamento_Serie.ID_empresa} = " & ID_empresa & " and Month({Faturamento_Serie.dt_DataEmissao}) = " & Mes & " And Year({Faturamento_Serie.dt_DataEmissao}) = " & Ano & " And {Faturamento_Serie.Int_TipoNota} = " & TipoNF & ""
   End If
   
   'Filtrar por período todas as Series
   If optPeriodo.Value = True And cmbSerie.Text = "" Then
   FormulaRelatorio = "{Faturamento_Serie.Remessa}=True and {Faturamento_Serie.tiponf} = 'M1' and {Faturamento_Serie.Aplicacao} = 'P' and {Faturamento_Serie.ID_empresa} = " & ID_empresa & " and {Faturamento_Serie.dt_DataEmissao} >= " & varinicio & " And {Faturamento_Serie.dt_DataEmissao}<= " & varfim & " And {Faturamento_Serie.Int_TipoNota} = " & TipoNF & ""
   End If
   
   'Filtrar por mês todas as Series
   If optMes.Value = True And cmbSerie.Text = "" Then
   FormulaRelatorio = "{Faturamento_Serie.Remessa}=True and {Faturamento_Serie.tiponf} = 'M1' and {Faturamento_Serie.Aplicacao} = 'P' and {Faturamento_Serie.ID_empresa} = " & ID_empresa & " and Month({Faturamento_Serie.dt_DataEmissao}) = " & Mes & " And Year({Faturamento_Serie.dt_DataEmissao}) = " & Ano & " And {Faturamento_Serie.Int_TipoNota} = " & TipoNF & ""
   End If
End If

'Debug.print FormulaRelatorio

ProcImprimirRel FormulaRelatorio, ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcImprimir_ListaNotas(FormulaRel As String, FormulaRelSubReport As String)
On Error GoTo tratar_erro

'Exemplo de como colocar variavel no relatorio

ProcVerifRelPersonalizado

If PermitidoRel = False Then LocalrelNovo = Localrel Else LocalrelNovo = LocalRelPersonalizado
Set Report = crAPP.OpenReport(LocalrelNovo & "\" & NomeRel)
'Login SQL
Contador = Report.Database.Tables.Count
Do While Contador > 0
    Set DBTable = Report.Database.Tables(Contador)
    ProcLogonBDSQL
    Contador = Contador - 1
Loop
ProcVerifSubReport FormulaRelSubReport

frmimprimir.CrystalActiveXReportViewer1.ReportSource = Report
Report.FormulaSyntax = crCrystalSyntaxFormula
Report.RecordSelectionFormula = FormulaRel
Report.ParameterFields(1).AddCurrentValue (Qtd)
Report.ParameterFields(2).AddCurrentValue (Quant)
Report.ParameterFields(3).AddCurrentValue (quantidade)

frmimprimir.CrystalActiveXReportViewer1.ViewReport
frmimprimir.CrystalActiveXReportViewer1.DisplayGroupTree = True
frmimprimir.WindowState = 0
frmimprimir.Show 1
2:
    Set Report = Nothing
    Set crAPP = Nothing

Exit Sub
tratar_erro:
    If Err.Number = "-2147206461" Then
        USMsgBox ("Não foi encontrado o relatório " & NomeRel & " na pasta " & LocalrelNovo), vbExclamation, "CAPRIND v5.0"
        GoTo 2
    End If
    If Err.Number = "-2147483638" Then
        USMsgBox ("Não foi possível visualizar o relatório, favor reiniciar o sistema."), vbExclamation, "CAPRIND v5.0"
        GoTo 2
    End If
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub Form_Load()
On Error GoTo tratar_erro

msk_fltInicio.Value = Date
msk_fltFim.Value = Date
TipoNF = "M1"

cmbAno.Text = Year(Date)

ProcCarregaComboEmpresa Cmb_empresa, False
ProcCarregaSerie

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optMes_Click()
On Error GoTo tratar_erro

FramePeriodo.Enabled = optPeriodo.Value
FrameMes.Enabled = optMes.Value
FrameAno.Visible = True
FramePeriodo.Visible = False
cmbAno.Visible = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optPeriodo_Click()
On Error GoTo tratar_erro

FramePeriodo.Enabled = optPeriodo.Value
FrameMes.Enabled = optMes.Value
FrameAno.Visible = False
cmbAno.Visible = False
FramePeriodo.Visible = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub ProcCarregaSerie()
On Error GoTo tratar_erro

With cmbSerie
    .Clear
    .AddItem "0"
    .AddItem "1"
    .AddItem "2"
    .AddItem "3"
    .AddItem "4"
    .AddItem "5"
    .AddItem "6"
    .AddItem "7"
    .AddItem "8"
    .AddItem "9"
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

