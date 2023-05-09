VERSION 5.00
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Begin VB.Form frmFaturamento_Impostos_Filtrar 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Faturamento | Filtrar Notas x CST"
   ClientHeight    =   6750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8655
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
   ScaleHeight     =   6750
   ScaleWidth      =   8655
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Opções pra filtro"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5325
      Index           =   1
      Left            =   180
      TabIndex        =   2
      Top             =   690
      Width           =   8175
      Begin DrawSuite2022.USButton btnFiltrar 
         Height          =   675
         Left            =   5520
         TabIndex        =   29
         Top             =   4500
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   1191
         DibPicture      =   "frmFaturamento_Impostos_Filtrar.frx":0000
         Caption         =   "Filtrar ICMS IPI"
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
      Begin VB.Frame FrameMes 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Escolha o mês"
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
         Left            =   1200
         TabIndex        =   16
         Top             =   1200
         Width           =   6885
         Begin DrawSuite2022.USOptionButton opt1 
            Height          =   285
            Left            =   120
            TabIndex        =   17
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
            Left            =   1167
            TabIndex        =   18
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
            Left            =   2274
            TabIndex        =   19
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
            Left            =   3381
            TabIndex        =   20
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
            Left            =   4518
            TabIndex        =   21
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
            Left            =   5745
            TabIndex        =   22
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
            Left            =   120
            TabIndex        =   23
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
            Left            =   1116
            TabIndex        =   24
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
            Left            =   2172
            TabIndex        =   25
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
            Left            =   3228
            TabIndex        =   26
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
            Left            =   4314
            TabIndex        =   27
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
            Left            =   5490
            TabIndex        =   28
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
            Value           =   -1  'True
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Empresa responsável pelas notas"
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
         Height          =   945
         Index           =   2
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   5910
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
            ItemData        =   "frmFaturamento_Impostos_Filtrar.frx":3650
            Left            =   120
            List            =   "frmFaturamento_Impostos_Filtrar.frx":3652
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   15
            ToolTipText     =   "Empresa."
            Top             =   390
            Width           =   5610
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Escolha o tipo de notas"
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
         Left            =   6060
         TabIndex        =   11
         Top             =   240
         Width           =   2025
         Begin DrawSuite2022.USOptionButton optEntrada 
            Height          =   195
            Left            =   150
            TabIndex        =   12
            Top             =   330
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   344
            Caption         =   "Notas de entrada"
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
            Value           =   -1  'True
         End
         Begin DrawSuite2022.USOptionButton optSaida 
            Height          =   195
            Left            =   150
            TabIndex        =   13
            Top             =   600
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   344
            Caption         =   "Notas de saida"
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
      End
      Begin VB.Frame FrameAno 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Escolha ano"
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
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   1065
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
            ItemData        =   "frmFaturamento_Impostos_Filtrar.frx":3654
            Left            =   90
            List            =   "frmFaturamento_Impostos_Filtrar.frx":3688
            Sorted          =   -1  'True
            TabIndex        =   10
            ToolTipText     =   "Opções para filtro."
            Top             =   480
            Width           =   795
         End
         Begin DrawSuite2022.USLabel USLabel1 
            Height          =   195
            Left            =   240
            TabIndex        =   35
            Top             =   300
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
      End
      Begin VB.Frame Frame3 
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
         Height          =   735
         Left            =   120
         TabIndex        =   7
         Top             =   2130
         Width           =   7965
         Begin VB.TextBox txtOrigem 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   150
            TabIndex        =   32
            Top             =   270
            Width           =   705
         End
         Begin VB.ComboBox cmbOrigem 
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
            ItemData        =   "frmFaturamento_Impostos_Filtrar.frx":36EC
            Left            =   870
            List            =   "frmFaturamento_Impostos_Filtrar.frx":36EE
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   270
            Width           =   6915
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "CST IPI"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   5
         Top             =   3570
         Width           =   7965
         Begin VB.TextBox txtCSTIPI 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   180
            TabIndex        =   31
            Top             =   270
            Width           =   705
         End
         Begin VB.ComboBox cmbCSTIPI 
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
            ItemData        =   "frmFaturamento_Impostos_Filtrar.frx":36F0
            Left            =   900
            List            =   "frmFaturamento_Impostos_Filtrar.frx":36F2
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   270
            Width           =   6915
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "CST ICMS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   2850
         Width           =   7965
         Begin VB.TextBox txtCSTICMS 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   150
            TabIndex        =   30
            Top             =   270
            Width           =   705
         End
         Begin VB.ComboBox cmbCSTIcms 
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
            ItemData        =   "frmFaturamento_Impostos_Filtrar.frx":36F4
            Left            =   870
            List            =   "frmFaturamento_Impostos_Filtrar.frx":36F6
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   270
            Width           =   6915
         End
      End
      Begin DrawSuite2022.USButton btnCreditoICMS 
         Height          =   675
         Left            =   150
         TabIndex        =   33
         Top             =   4500
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   1191
         DibPicture      =   "frmFaturamento_Impostos_Filtrar.frx":36F8
         Caption         =   "Créditos do ICMS"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   4960354
         BorderColorDisabled=   13160660
         BorderColorDown =   4210752
         BorderColorOver =   49152
         GradientColor1  =   4960354
         GradientColor2  =   4960354
         GradientColor3  =   4960354
         GradientColor4  =   4960354
         GradientColorDisabled1=   14215660
         GradientColorDisabled2=   14215660
         GradientColorDisabled3=   14215660
         GradientColorDisabled4=   14215660
         GradientColorOver1=   49152
         GradientColorOver2=   49152
         GradientColorOver3=   49152
         GradientColorOver4=   49152
         GradientColorDown1=   32768
         GradientColorDown2=   32768
         GradientColorDown3=   32768
         GradientColorDown4=   32768
         ShowFocusRect   =   0   'False
         Theme           =   3
      End
      Begin DrawSuite2022.USButton btnCreditoIPI 
         Height          =   675
         Left            =   2820
         TabIndex        =   34
         Top             =   4500
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   1191
         DibPicture      =   "frmFaturamento_Impostos_Filtrar.frx":6D48
         Caption         =   "Créditos do IPI"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   1154291
         BorderColorDisabled=   13160660
         BorderColorDown =   16576
         BorderColorOver =   8438015
         GradientColor1  =   1154291
         GradientColor2  =   1154291
         GradientColor3  =   1154291
         GradientColor4  =   1154291
         GradientColorDisabled1=   14215660
         GradientColorDisabled2=   14215660
         GradientColorDisabled3=   14215660
         GradientColorDisabled4=   14215660
         GradientColorOver1=   8438015
         GradientColorOver2=   8438015
         GradientColorOver3=   8438015
         GradientColorOver4=   8438015
         GradientColorDown1=   16576
         GradientColorDown2=   16576
         GradientColorDown3=   16576
         GradientColorDown4=   16576
         ShowFocusRect   =   0   'False
         Theme           =   5
      End
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   435
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   767
      DibPicture      =   "frmFaturamento_Impostos_Filtrar.frx":A398
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
      Icon            =   "frmFaturamento_Impostos_Filtrar.frx":D9E8
   End
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   6345
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   714
   End
End
Attribute VB_Name = "frmFaturamento_Impostos_Filtrar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ProcFiltrar()
On Error GoTo tratar_erro


FormulaRelatorio = ""
StrSql = ""
StrSQLTotais = ""


ID_empresa = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)

If optEntrada.Value = True Then
Tipo_Nota = "2"
End If

If optSaida.Value = True Then
Tipo_Nota = "1"
End If

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

'=======================================================================
' Filtrar CST com origem com CST Icms com CST IPI (Origem + icms + ipi)
'=======================================================================
If txtOrigem.Text <> "" And txtCSTICMS.Text <> "" And txtCSTIPI.Text <> "" Then
StrSql = "Select * from Faturamento_Detalhes_Fiscal_CreditoICMS Where ID_empresa = '" & ID_empresa & "' and int_TipoNota = '" & Tipo_Nota & "' And CSTICMS = '" & txtOrigem.Text & txtCSTICMS.Text & "' And CSTIPI = '" & txtCSTIPI.Text & "' and Year(Data) = '" & cmbAno.Text & "' and Month(Data) = '" & Mes & "'"
StrSQLTotais = "Select Sum(ValorICMS) as TotalICMS, Sum(ValorIPI) as TotalIPI from Faturamento_Detalhes_Fiscal_CreditoICMS Where ID_empresa = '" & ID_empresa & "' and int_TipoNota = '" & Tipo_Nota & "' And CSTICMS = '" & txtOrigem.Text & txtCSTICMS.Text & "' And CSTIPI = '" & txtCSTIPI.Text & "' and Year(Data) = '" & cmbAno.Text & "' and Month(Data) = '" & Mes & "'"
FormulaRelatorio = "{Faturamento_Detalhes_Fiscal_CreditoICMS.ID_empresa} = " & ID_empresa & " and {Faturamento_Detalhes_Fiscal_CreditoICMS.int_TipoNota} = " & Tipo_Nota & " And {Faturamento_Detalhes_Fiscal_CreditoICMS.CSTICMS} = '" & txtOrigem.Text & txtCSTICMS.Text & "' And {Faturamento_Detalhes_Fiscal_CreditoICMS.CSTIPI} = '" & txtCSTIPI.Text & "' and Year({Faturamento_Detalhes_Fiscal_CreditoICMS.Data}) = " & cmbAno.Text & " and Month({Faturamento_Detalhes_Fiscal_CreditoICMS.Data}) = " & Mes & ""
End If

'=======================================================================
' Filtrar CST com origem com CST Icms sem CST IPI (Origem + icms)
'=======================================================================
If txtOrigem.Text <> "" And txtCSTICMS.Text <> "" And txtCSTIPI.Text = "" Then
StrSql = "Select * from Faturamento_Detalhes_Fiscal_CreditoICMS Where ID_empresa = '" & ID_empresa & "' and int_TipoNota = '" & Tipo_Nota & "' And CSTICMS = '" & txtOrigem.Text & txtCSTICMS.Text & "' and Year(Data) = '" & cmbAno.Text & "' and Month(Data) = '" & Mes & "'"
StrSQLTotais = "Select Sum(ValorICMS) as TotalICMS, Sum(ValorIPI) as TotalIPI from Faturamento_Detalhes_Fiscal_CreditoICMS Where ID_empresa = '" & ID_empresa & "' and int_TipoNota = '" & Tipo_Nota & "' And CSTICMS = '" & txtOrigem.Text & txtCSTICMS.Text & "' and Year(Data) = '" & cmbAno.Text & "' and Month(Data) = '" & Mes & "'"
FormulaRelatorio = "{Faturamento_Detalhes_Fiscal_CreditoICMS.ID_empresa} = " & ID_empresa & " and {Faturamento_Detalhes_Fiscal_CreditoICMS.int_TipoNota} = " & Tipo_Nota & " And {Faturamento_Detalhes_Fiscal_CreditoICMS.CSTICMS} = '" & txtOrigem.Text & txtCSTICMS.Text & "' and Year({Faturamento_Detalhes_Fiscal_CreditoICMS.Data}) = " & cmbAno.Text & " and Month({Faturamento_Detalhes_Fiscal_CreditoICMS.Data}) = " & Mes & ""
End If

'=======================================================================
' Filtrar CST com origem sem CST Icms sem CST IPI (Só origem)
'=======================================================================
If txtOrigem.Text <> "" And txtCSTICMS.Text = "" And txtCSTIPI.Text = "" Then
StrSql = "Select * from Faturamento_Detalhes_Fiscal_CreditoICMS where ID_empresa = '" & ID_empresa & "' and int_TipoNota = '" & Tipo_Nota & "' And CSTICMS Like '" & txtOrigem.Text & "%' and Year(Data) = '" & cmbAno.Text & "' and Month(Data) = '" & Mes & "'"
StrSQLTotais = "Select Sum(ValorICMS) as TotalICMS, Sum(ValorIPI) as TotalIPI from Faturamento_Detalhes_Fiscal_CreditoICMS where ID_empresa = '" & ID_empresa & "' and int_TipoNota = '" & Tipo_Nota & "' And CSTICMS Like '" & txtOrigem.Text & "%' and Year(Data) = '" & cmbAno.Text & "' and Month(Data) = '" & Mes & "'"
FormulaRelatorio = "{Faturamento_Detalhes_Fiscal_CreditoICMS.ID_empresa} = " & ID_empresa & " and {Faturamento_Detalhes_Fiscal_CreditoICMS.int_TipoNota} = " & Tipo_Nota & " And {Faturamento_Detalhes_Fiscal_CreditoICMS.CSTICMS} Like '" & txtOrigem.Text & "%' and Year({Faturamento_Detalhes_Fiscal_CreditoICMS.Data}) = " & cmbAno.Text & " and Month({Faturamento_Detalhes_Fiscal_CreditoICMS.Data}) = " & Mes & ""
End If

'=======================================================================
' Filtrar CST com origem sem CST Icms Com CST IPI (Origem + IPI)
'=======================================================================
If txtOrigem.Text <> "" And txtCSTICMS.Text = "" And txtCSTIPI.Text <> "" Then
StrSql = "Select * from Faturamento_Detalhes_Fiscal_CreditoICMS Where ID_empresa = '" & ID_empresa & "' and int_TipoNota = '" & Tipo_Nota & "' And CSTICMS Like '" & txtOrigem.Text & "%' and CSTIPI = '" & txtCSTIPI.Text & "' and Year(Data) = '" & cmbAno.Text & "' and Month(Data) = '" & Mes & "'"
StrSQLTotais = "Select Sum(ValorICMS) as TotalICMS, Sum(ValorIPI) as TotalIPI from Faturamento_Detalhes_Fiscal_CreditoICMS Where ID_empresa = '" & ID_empresa & "' and int_TipoNota = '" & Tipo_Nota & "' And CSTICMS Like '" & txtOrigem.Text & "%' and CSTIPI = '" & txtCSTIPI.Text & "' and Year(Data) = '" & cmbAno.Text & "' and Month(Data) = '" & Mes & "'"
FormulaRelatorio = "{Faturamento_Detalhes_Fiscal_CreditoICMS.ID_empresa} = " & ID_empresa & " and {Faturamento_Detalhes_Fiscal_CreditoICMS.int_TipoNota} = " & Tipo_Nota & " And {Faturamento_Detalhes_Fiscal_CreditoICMS.CSTICMS} = '" & txtOrigem.Text & "' And {Faturamento_Detalhes_Fiscal_CreditoICMS.CSTIPI} = '" & txtCSTIPI.Text & "' and Year({Faturamento_Detalhes_Fiscal_CreditoICMS.Data}) = " & cmbAno.Text & " and Month({Faturamento_Detalhes_Fiscal_CreditoICMS.Data}) = " & Mes & ""
End If

'=======================================================================
' Filtrar CST sem origem com CST Icms Com CST IPI (Só Icms + ipi)
'=======================================================================
If txtOrigem.Text = "" And txtCSTICMS.Text <> "" And txtCSTIPI.Text <> "" Then
StrSql = "Select * from Faturamento_Detalhes_Fiscal_CreditoICMS Where ID_empresa = '" & ID_empresa & "' and int_TipoNota = '" & Tipo_Nota & "' And CSTICMS Like '%" & txtCSTICMS.Text & "' and CSTIPI = '" & txtCSTIPI.Text & "' and Year(Data) = '" & cmbAno.Text & "' and Month(Data) = '" & Mes & "'"
StrSQLTotais = "Select Sum(ValorICMS) as TotalICMS, Sum(ValorIPI) as TotalIPI from Faturamento_Detalhes_Fiscal_CreditoICMS Where ID_empresa = '" & ID_empresa & "' and int_TipoNota = '" & Tipo_Nota & "' And CSTICMS Like '%" & txtCSTICMS.Text & "' and CSTIPI = '" & txtCSTIPI.Text & "' and Year(Data) = '" & cmbAno.Text & "' and Month(Data) = '" & Mes & "'"
FormulaRelatorio = "{Faturamento_Detalhes_Fiscal_CreditoICMS.ID_empresa} = " & ID_empresa & " and {Faturamento_Detalhes_Fiscal_CreditoICMS.int_TipoNota} = " & Tipo_Nota & " And {Faturamento_Detalhes_Fiscal_CreditoICMS.CSTICMS} = '" & txtCSTICMS.Text & "' And {Faturamento_Detalhes_Fiscal_CreditoICMS.CSTIPI} = '" & txtCSTIPI.Text & "' and Year({Faturamento_Detalhes_Fiscal_CreditoICMS.Data}) = " & cmbAno.Text & " and Month({Faturamento_Detalhes_Fiscal_CreditoICMS.Data}) = " & Mes & ""
End If

'=======================================================================
' Filtrar CST sem origem sem CST Icms Com CST IPI (Só ipi)
'=======================================================================
If txtOrigem.Text = "" And txtCSTICMS.Text = "" And txtCSTIPI.Text <> "" Then
StrSql = "Select * from Faturamento_Detalhes_Fiscal_CreditoICMS Where ID_empresa = '" & ID_empresa & "' and int_TipoNota = '" & Tipo_Nota & "' and CSTIPI = '" & txtCSTIPI.Text & "' and Year(Data) = '" & cmbAno.Text & "' and Month(Data) = '" & Mes & "'"
StrSQLTotais = "Select Sum(ValorICMS) as TotalICMS, Sum(ValorIPI) as TotalIPI from Faturamento_Detalhes_Fiscal_CreditoICMS Where ID_empresa = '" & ID_empresa & "' and int_TipoNota = '" & Tipo_Nota & "' and CSTIPI = '" & txtCSTIPI.Text & "' and Year(Data) = '" & cmbAno.Text & "' and Month(Data) = '" & Mes & "'"
FormulaRelatorio = "{Faturamento_Detalhes_Fiscal_CreditoICMS.ID_empresa} = " & ID_empresa & " and {Faturamento_Detalhes_Fiscal_CreditoICMS.int_TipoNota} = " & Tipo_Nota & " And {Faturamento_Detalhes_Fiscal_CreditoICMS.CSTIPI} = '" & txtCSTIPI.Text & "' and Year({Faturamento_Detalhes_Fiscal_CreditoICMS.Data}) = " & cmbAno.Text & " and Month({Faturamento_Detalhes_Fiscal_CreditoICMS.Data}) = " & Mes & ""
End If

'=======================================================================
' Filtrar CST sem origem com CST Icms sem CST IPI (Só Icms)
'=======================================================================
If txtOrigem.Text = "" And txtCSTICMS.Text <> "" And txtCSTIPI.Text = "" Then
StrSql = "Select * from Faturamento_Detalhes_Fiscal_CreditoICMS Where ID_empresa = '" & ID_empresa & "' and int_TipoNota = '" & Tipo_Nota & "' And CSTICMS Like '%" & txtCSTICMS.Text & "' and Year(Data) = '" & cmbAno.Text & "' and Month(Data) = '" & Mes & "'"
StrSQLTotais = "Select Sum(ValorICMS) as TotalICMS, Sum(ValorIPI) as TotalIPI from Faturamento_Detalhes_Fiscal_CreditoICMS Where ID_empresa = '" & ID_empresa & "' and int_TipoNota = '" & Tipo_Nota & "' And CSTICMS Like '%" & txtCSTICMS.Text & "' and Year(Data) = '" & cmbAno.Text & "' and Month(Data) = '" & Mes & "'"
FormulaRelatorio = "{Faturamento_Detalhes_Fiscal_CreditoICMS.ID_empresa} = " & ID_empresa & " and {Faturamento_Detalhes_Fiscal_CreditoICMS.int_TipoNota} = " & Tipo_Nota & " And {Faturamento_Detalhes_Fiscal_CreditoICMS.CSTICMS} = '" & txtCSTICMS.Text & "' And {Faturamento_Detalhes_Fiscal_CreditoICMS.CSTIPI} = '" & txtCSTIPI.Text & "' and Year({Faturamento_Detalhes_Fiscal_CreditoICMS.Data}) = " & cmbAno.Text & " and Month({Faturamento_Detalhes_Fiscal_CreditoICMS.Data}) = " & Mes & ""
End If

'=======================================================================
' Filtrar CST sem origem sem CST Icms sem CST IPI (Tudo)
'=======================================================================
If txtOrigem.Text = "" And txtCSTICMS.Text = "" And txtCSTIPI.Text = "" Then
StrSql = "Select * from Faturamento_Detalhes_Fiscal_CreditoICMS Where ID_empresa = '" & ID_empresa & "' and int_TipoNota = '" & Tipo_Nota & "' and Year(Data) = '" & cmbAno.Text & "' and Month(Data) = '" & Mes & "'"
StrSQLTotais = "Select Sum(ValorICMS) as TotalICMS, Sum(ValorIPI) as TotalIPI from Faturamento_Detalhes_Fiscal_CreditoICMS Where ID_empresa = '" & ID_empresa & "' and int_TipoNota = '" & Tipo_Nota & "' and Year(Data) = '" & cmbAno.Text & "' and Month(Data) = '" & Mes & "'"
FormulaRelatorio = "{Faturamento_Detalhes_Fiscal_CreditoICMS.ID_empresa} = " & ID_empresa & " and {Faturamento_Detalhes_Fiscal_CreditoICMS.int_TipoNota} = " & Tipo_Nota & " And Year({Faturamento_Detalhes_Fiscal_CreditoICMS.Data}) = " & cmbAno.Text & " and Month({Faturamento_Detalhes_Fiscal_CreditoICMS.Data}) = " & Mes & ""
End If


'Debug.print StrSql
'Debug.print FormulaRelatorio
'Debug.print StrSQLTotais

frmFaturamento_Impostos.ProcCarregaLista

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub btnCreditoICMS_Click()
On Error GoTo tratar_erro
Dim Tipo_Nota As String

ProcBloqueiaCampos
FormulaRelatorio = ""
StrSql = ""
StrSQLTotais = ""


ID_empresa = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)

If optEntrada.Value = True Then
Tipo_Nota = "2"
End If

If optSaida.Value = True Then
Tipo_Nota = "1"
End If

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

'=======================================================================
' Filtrar todas com crédito do ICMS
'=======================================================================

StrSql = "Select * from Faturamento_Detalhes_Fiscal_CreditoICMS Where ID_empresa = '" & ID_empresa & "' and int_TipoNota = '" & Tipo_Nota & "' And (Tributacao_ICMS = '101' or Tributacao_ICMS = '00') and Year(Data) = '" & cmbAno.Text & "' and Month(Data) = '" & Mes & "'"
StrSQLTotais = "Select Sum(ValorICMS) as TotalICMS, Sum(ValorIPI) as TotalIPI from Faturamento_Detalhes_Fiscal_CreditoICMS Where ID_empresa = '" & ID_empresa & "' and int_TipoNota = '" & Tipo_Nota & "' And (Tributacao_ICMS = '101' or Tributacao_ICMS = '00') and Year(Data) = '" & cmbAno.Text & "' and Month(Data) = '" & Mes & "'"
FormulaRelatorio = "{Faturamento_Detalhes_Fiscal_CreditoICMS.ID_empresa} = " & ID_empresa & " and {Faturamento_Detalhes_Fiscal_CreditoICMS.int_TipoNota} = " & Tipo_Nota & " And ({Faturamento_Detalhes_Fiscal_CreditoICMS.Tributacao_ICMS} = 101 or  {Faturamento_Detalhes_Fiscal_CreditoICMS.Tributacao_ICMS} = 00) and Year({Faturamento_Detalhes_Fiscal_CreditoICMS.Data}) = " & cmbAno.Text & " and Month({Faturamento_Detalhes_Fiscal_CreditoICMS.Data}) = " & Mes & ""


'Debug.print StrSql
'Debug.print FormulaRelatorio
'Debug.print StrSQLTotais

frmFaturamento_Impostos.ProcCarregaLista

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub btnCreditoIPI_Click()
On Error GoTo tratar_erro
Dim Tipo_Nota As String

ProcBloqueiaCampos
FormulaRelatorio = ""
StrSql = ""
StrSQLTotais = ""


ID_empresa = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)

If optEntrada.Value = True Then
Tipo_Nota = "2"
End If

If optSaida.Value = True Then
Tipo_Nota = "1"
End If

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

'=======================================================================
' Filtrar todas com crédito do IPI
'=======================================================================

StrSql = "Select * from Faturamento_Detalhes_Fiscal_CreditoICMS Where ID_empresa = '" & ID_empresa & "' and int_TipoNota = '" & Tipo_Nota & "' And CSTIPI = '00' and Year(Data) = '" & cmbAno.Text & "' and Month(Data) = '" & Mes & "'"
StrSQLTotais = "Select Sum(ValorICMS) as TotalICMS, Sum(ValorIPI) as TotalIPI from Faturamento_Detalhes_Fiscal_CreditoICMS Where ID_empresa = '" & ID_empresa & "' and int_TipoNota = '" & Tipo_Nota & "' And CSTIPI = '00' and Year(Data) = '" & cmbAno.Text & "' and Month(Data) = '" & Mes & "'"
FormulaRelatorio = "{Faturamento_Detalhes_Fiscal_CreditoICMS.ID_empresa} = " & ID_empresa & " and {Faturamento_Detalhes_Fiscal_CreditoICMS.int_TipoNota} = " & Tipo_Nota & " And {Faturamento_Detalhes_Fiscal_CreditoICMS.CSTIPI} = '00' and Year({Faturamento_Detalhes_Fiscal_CreditoICMS.Data}) = " & cmbAno.Text & " and Month({Faturamento_Detalhes_Fiscal_CreditoICMS.Data}) = " & Mes & ""


'Debug.print StrSql
'Debug.print FormulaRelatorio
'Debug.print StrSQLTotais

frmFaturamento_Impostos.ProcCarregaLista

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub btnFiltrar_Click()
On Error GoTo tratar_erro

ProcFiltrar
ProcBloqueiaCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub cmbCSTIcms_Click()
On Error GoTo tratar_erro

If cmbCSTIcms.Text <> "" Then
   CSTICMS = ReturnNumbersOnly(cmbCSTIcms)
   txtCSTICMS.Text = ReturnNumbersOnly(cmbCSTIcms)
End If

'If cmbCSTIcms <> "" And txtOrigem.Text <> "" Then
'txtCSTICMS.Text = txtOrigem.Text & ReturnNumbersOnly(cmbCSTIcms)
'End If

If cmbCSTIcms.Text = "" Then
txtCSTICMS.Text = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbCSTIPI_Change()
On Error GoTo tratar_erro

txtCSTIPI = Left(cmbCSTIPI.Text, 2)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbCSTIPI_Click()
On Error GoTo tratar_erro

txtCSTIPI.Text = Left(cmbCSTIPI.Text, 2)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub cmbOrigem_Click()
On Error GoTo tratar_erro

If cmbOrigem <> "" Then
    Origem = Left(cmbOrigem, 1)
    txtOrigem.Text = Origem
End If

'If cmbCSTIcms <> "" And txtOrigem.Text <> "" Then
'txtCSTICMS.Text = txtOrigem.Text & ReturnNumbersOnly(cmbCSTIcms)
'End If

If cmbOrigem.Text = "" Then
txtOrigem = ""
txtCSTICMS.Text = ""
cmbCSTIcms.ListIndex = -1
End If



Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcBloqueiaCampos()
On Error GoTo tratar_erro

With frmFaturamento_Impostos


If optEntrada.Value = False Then
    .USToolBar1.ButtonState(2) = 5
    .GridItens.Column(9).Locked = True
    .GridItens.Column(10).Locked = True
    .GridItens.Column(11).Locked = True
    .GridItens.Column(12).Locked = True
    .GridItens.Column(13).Locked = True
    .GridItens.Column(14).Locked = True
    .GridItens.Column(15).Locked = True
Else
    .USToolBar1.ButtonState(2) = 0
    .GridItens.Column(9).Locked = True
    .GridItens.Column(10).Locked = True
    .GridItens.Column(11).Locked = True
    .GridItens.Column(12).Locked = True
    .GridItens.Column(13).Locked = True
    .GridItens.Column(14).Locked = True
    .GridItens.Column(15).Locked = True
End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaComboEmpresa Cmb_empresa, False
ProcCarregaComboOrigem
ProcCarregaComboCSTICMS
ProcCarregaComboCSTIPI


ProcCarregaComboAno cmbAno, "2005", 1
cmbAno.Text = Year(Date)

Select Case Month(Date)
Case 1: Opt1.Value = True
Case 2: Opt2.Value = True
Case 3: opt3.Value = True
Case 4: Opt4.Value = True
Case 5: Opt5.Value = True
Case 6: opt6.Value = True
Case 7: opt7.Value = True
Case 8: opt8.Value = True
Case 9: opt9.Value = True
Case 10: opt10.Value = True
Case 11: opt11.Value = True
Case 12: opt12.Value = True
End Select


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaComboOrigem()
On Error GoTo tratar_erro

With cmbOrigem

.AddItem ""
.AddItem "0 - Nacional, exceto as indicadas nos códigos 3, 4, 5 e 8"
.AddItem "1 - Estrangeira – Importação direta, exceto a indicada no código 6"
.AddItem "2 - Estrangeira – Adquirida no mercado interno, exceto a indicada no código 7"
.AddItem "3 - Nacional, mercadoria ou bem com Conteúdo de Importação superior a 40% (quarenta por cento) e inferior ou igual a 70% (setenta por cento)"
.AddItem "4 - Nacional, cuja produção tenha sido feita em conformidade com os processos produtivos básicos (PPB) de que tratam o Decreto-Lei nº 288/1967, e as Leis nºs 8.248/1991, 8.387/1991, 10.176/2001 e 11.484/2007."
.AddItem "5 - Nacional, mercadoria ou bem com Conteúdo de Importação inferior ou igual a 40% (quarenta por cento)."
.AddItem "6 - Estrangeira – Importação direta, sem similar nacional, constante em lista de Resolução CAMEX e gás natural."
.AddItem "7 - Estrangeira – Adquirida no mercado interno, sem similar nacional, constante em lista de Resolução CAMEX e gás natural."
.AddItem "8 - Nacional, mercadoria ou bem com Conteúdo de Importação superior a 70% (setenta por cento)"

End With


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaComboCSTICMS()
On Error GoTo tratar_erro
cmbCSTIcms.Clear

With cmbCSTIcms
    .AddItem ""
    .AddItem "00 - Tributada integralmente"
    .AddItem "10 - Tributada e com cobrança do ICMS por substituição"
    .AddItem "20 - Com redução de base de cálculo"
    .AddItem "40 - Isenta"
    .AddItem "41 - Não tributada"
    .AddItem "50 - Suspensão"
    .AddItem "51 - Diferimento"
    .AddItem "60 - ICMS cobrado anteriormente por substituição tributária"
    .AddItem "70 - Com redução de base de cálculo e cobrança do ICMS por substituição tributária"
    .AddItem "90 - Outras"
    .AddItem "101 - Tributada pelo Simples Nacional com permissão de crédito"
    .AddItem "102 - Tributada pelo Simples Nacional sem permissão de crédito"
    .AddItem "103 - Isenção do ICMS no Simples Nacional para faixa de receita bruta"
    .AddItem "201 - Tributada pelo Simples Nacional com permissão de crédito e com cobrança do ICMS por Substituição Tributária"
    .AddItem "202 - Tributada pelo Simples Nacional sem permissão de crédito e com cobrança do ICMS por Substituição Tributária"
    .AddItem "203 - Isenção do ICMS nos Simples Nacional para faixa de receita bruta e com cobrança do ICMS por Substituição Tributária"
    .AddItem "300 - Imune"
    .AddItem "400 - Não tributada pelo Simples Nacional"
    .AddItem "500 - ICMS cobrado anteriormente por substituição tributária (substituído) ou por antecipação"
    .AddItem "900 - Outros"
End With


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaComboCSTIPI()
On Error GoTo tratar_erro
cmbCSTIPI.Clear

With cmbCSTIPI
    .AddItem ""
    .AddItem "00 - Entrada com recuperação de crédito"
    .AddItem "01 - Entrada tributada com alíquota zero"
    .AddItem "02 - Entrada isenta"
    .AddItem "03 - Entrada não-tributada"
    .AddItem "04 - Entrada imune"
    .AddItem "05 - Entrada com suspensão"
    .AddItem "49 - Outras entradas"
    .AddItem "50 - Saída tributada"
    .AddItem "51 - Saída tributada com alíquota zero"
    .AddItem "52 - Saída isenta"
    .AddItem "53 - Saída não-tributada"
    .AddItem "54 - Saída imune"
    .AddItem "55 - Saída com suspensão"
    .AddItem "99 - Outras saídas"
End With


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optEntrada_Click()
On Error GoTo tratar_erro

ProcBloqueiaCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optSaida_Click()
On Error GoTo tratar_erro

ProcBloqueiaCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
