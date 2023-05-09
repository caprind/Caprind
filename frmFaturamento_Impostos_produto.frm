VERSION 5.00
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Begin VB.Form frmFaturamento_Impostos_produto 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "CAPRIND v5.0 | Impostos do produto"
   ClientHeight    =   9870
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11790
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
   ScaleHeight     =   9870
   ScaleWidth      =   11790
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Impostos do PIS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   3135
      Left            =   3960
      TabIndex        =   45
      Top             =   6240
      Width           =   3765
      Begin VB.TextBox txtPISValor 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2340
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   1575
         Width           =   1005
      End
      Begin VB.TextBox txtPISBaseCalc 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2340
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   1170
         Width           =   1005
      End
      Begin VB.TextBox txtPISPercentual 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2340
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   765
         Width           =   1005
      End
      Begin VB.TextBox txtPISCST 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2340
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   360
         Width           =   1005
      End
      Begin DrawSuite2022.USButton btnCSTPIS 
         Height          =   285
         Left            =   3360
         TabIndex        =   88
         ToolTipText     =   "Informações gerais sobre a tributação do PIS"
         Top             =   360
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   503
         DibPicture      =   "frmFaturamento_Impostos_produto.frx":0000
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
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
         ShowFocusRect   =   0   'False
         Theme           =   1
      End
      Begin DrawSuite2022.USButton btnSalvarPIS 
         Height          =   285
         Left            =   3360
         TabIndex        =   97
         ToolTipText     =   "Gravar valores"
         Top             =   1580
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   503
         DibPicture      =   "frmFaturamento_Impostos_produto.frx":7193
         Caption         =   ""
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
         Theme           =   4
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Valor  do PIS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   180
         TabIndex        =   57
         Top             =   1590
         Width           =   2265
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "CST"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   180
         TabIndex        =   56
         Top             =   390
         Width           =   585
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "% PIS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   180
         TabIndex        =   55
         Top             =   765
         Width           =   2175
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Valor base de cálculo do PIS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   180
         TabIndex        =   54
         Top             =   1170
         Width           =   2265
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "IPI Devolução"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   615
      Left            =   180
      TabIndex        =   100
      Top             =   8760
      Width           =   3765
      Begin VB.TextBox txtvIPIdevolv 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2340
         TabIndex        =   104
         TabStop         =   0   'False
         Top             =   270
         Width           =   975
      End
      Begin VB.TextBox txtpIPIdevolv 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   630
         TabIndex        =   101
         TabStop         =   0   'False
         Top             =   270
         Width           =   405
      End
      Begin DrawSuite2022.USButton btnIPIdevolucao 
         Height          =   285
         Left            =   3330
         TabIndex        =   102
         ToolTipText     =   "Gravar valores"
         Top             =   270
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   503
         DibPicture      =   "frmFaturamento_Impostos_produto.frx":FB98
         Caption         =   ""
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
         Theme           =   4
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Valor"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   3
         Left            =   1920
         TabIndex        =   105
         Top             =   300
         Width           =   975
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "% IPI"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   2
         Left            =   120
         TabIndex        =   103
         Top             =   300
         Width           =   975
      End
   End
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   92
      Top             =   9465
      Width           =   11790
      _ExtentX        =   20796
      _ExtentY        =   714
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Impostos do ICMS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   5625
      Left            =   180
      TabIndex        =   1
      Top             =   630
      Width           =   11325
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Diferencial de aliquota (DIFAL)"
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
         Height          =   2835
         Index           =   4
         Left            =   5610
         TabIndex        =   15
         Top             =   270
         Width           =   5535
         Begin VB.TextBox txtICMSDIFValorFCP 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3960
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   2460
            Width           =   1005
         End
         Begin VB.TextBox txtICMSDIFPercentualFCP 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3960
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   2142
            Width           =   1005
         End
         Begin VB.TextBox txtICMSDIFPercentualProvisorio 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3960
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   1825
            Width           =   1005
         End
         Begin VB.TextBox txtICMSDIFValorUFRemetente 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3960
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   1508
            Width           =   1005
         End
         Begin VB.TextBox txtICMSDIFValorUFDest 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3960
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   1191
            Width           =   1005
         End
         Begin VB.TextBox txtICMSDIFBaseUFDestino 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3960
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   874
            Width           =   1005
         End
         Begin VB.TextBox txtICMSDIFValor 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3960
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   557
            Width           =   1005
         End
         Begin VB.TextBox txtICMSDIFPercentual 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3960
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   240
            Width           =   1005
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Valor do fundo de combate a pobreza"
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
            Index           =   8
            Left            =   270
            TabIndex        =   23
            Top             =   2490
            Width           =   3945
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "% Percentual fundo de combate a pobreza (FCP)"
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
            Index           =   7
            Left            =   270
            TabIndex        =   22
            Top             =   2175
            Width           =   3945
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "% Percentual provisório"
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
            Index           =   6
            Left            =   270
            TabIndex        =   21
            Top             =   1860
            Width           =   3945
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Valor do ICMS interno UF remetente"
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
            Index           =   5
            Left            =   270
            TabIndex        =   20
            Top             =   1545
            Width           =   3945
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Valor do ICMS interno UF destino"
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
            Index           =   4
            Left            =   300
            TabIndex        =   19
            Top             =   1230
            Width           =   3945
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Valor base de calculo ICMS UF destino"
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
            Index           =   3
            Left            =   270
            TabIndex        =   18
            Top             =   930
            Width           =   3945
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Valor do ICMS diferencial"
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
            Index           =   2
            Left            =   270
            TabIndex        =   17
            Top             =   615
            Width           =   3945
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "% Percentual do ICMS diferencial"
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
            Index           =   1
            Left            =   270
            TabIndex        =   16
            Top             =   300
            Width           =   3945
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Desoneração do ICMS (Suframa)"
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
         Height          =   2355
         Index           =   3
         Left            =   5610
         TabIndex        =   12
         Top             =   3120
         Width           =   5535
         Begin VB.TextBox txtDescontoSuframa 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3960
            TabIndex        =   83
            TabStop         =   0   'False
            Top             =   1920
            Width           =   1005
         End
         Begin DrawSuite2022.USCheckBox CFOPDesonerada 
            Height          =   225
            Left            =   240
            TabIndex        =   81
            Top             =   480
            Width           =   4695
            _ExtentX        =   8281
            _ExtentY        =   397
            Alignment       =   1
            Caption         =   "Natureza de operação (CFOP) com desoneração do ICMS"
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
         End
         Begin VB.TextBox txtICMSDesonValor 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3960
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   1605
            Width           =   1005
         End
         Begin VB.TextBox txtICMSDesonMotivo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3960
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   1290
            Width           =   1005
         End
         Begin DrawSuite2022.USCheckBox Suframa 
            Height          =   225
            Left            =   210
            TabIndex        =   82
            Top             =   810
            Width           =   4725
            _ExtentX        =   8334
            _ExtentY        =   397
            Alignment       =   1
            Caption         =   "Cliente com regime especial (Suframa)"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   16384
            ShowFocusRect   =   0   'False
         End
         Begin DrawSuite2022.USButton USButton1 
            Height          =   285
            Left            =   5010
            TabIndex        =   90
            ToolTipText     =   "Informações gerais sobre a origem da mercadoria"
            Top             =   480
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   503
            DibPicture      =   "frmFaturamento_Impostos_produto.frx":1859D
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
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
            ShowFocusRect   =   0   'False
            Theme           =   1
         End
         Begin DrawSuite2022.USButton USButton2 
            Height          =   285
            Left            =   5010
            TabIndex        =   91
            ToolTipText     =   "Informações gerais sobre a tributação do ICMS"
            Top             =   825
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   503
            DibPicture      =   "frmFaturamento_Impostos_produto.frx":1F730
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
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
            ShowFocusRect   =   0   'False
            Theme           =   1
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Valor Desconto Suframa"
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
            Index           =   2
            Left            =   240
            TabIndex        =   84
            Top             =   1995
            Width           =   4065
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Valor ICMS Desonerado"
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
            Index           =   1
            Left            =   240
            TabIndex        =   14
            Top             =   1680
            Width           =   4065
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Motivo ICMS desonerado"
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
            Left            =   240
            TabIndex        =   13
            Top             =   1380
            Width           =   4065
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Substituição tributária"
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
         Height          =   2355
         Index           =   2
         Left            =   150
         TabIndex        =   5
         Top             =   3120
         Width           =   5445
         Begin VB.TextBox txtICMSSTAliquota 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3960
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   1950
            Width           =   1005
         End
         Begin VB.TextBox txtICMSSTValor 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3960
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   1620
            Width           =   1005
         End
         Begin VB.TextBox txtICMSSTValorBaseCal 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3960
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   1290
            Width           =   1005
         End
         Begin VB.TextBox txtICMSSTPercentualRed 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3960
            TabIndex        =   32
            TabStop         =   0   'False
            Top             =   960
            Width           =   1005
         End
         Begin VB.TextBox txtICMSSTMargem 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3960
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   630
            Width           =   1005
         End
         Begin VB.TextBox txtICMSSTModalidade 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3960
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   300
            Width           =   1005
         End
         Begin DrawSuite2022.USButton btnSubstituicaotributaria 
            Height          =   285
            Left            =   4980
            TabIndex        =   93
            ToolTipText     =   "Informações gerais sobre a substituição tributária"
            Top             =   300
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   503
            DibPicture      =   "frmFaturamento_Impostos_produto.frx":268C3
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
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
            ShowFocusRect   =   0   'False
            Theme           =   1
         End
         Begin DrawSuite2022.USButton btnSalvarST 
            Height          =   285
            Left            =   4980
            TabIndex        =   99
            ToolTipText     =   "Gravar valores"
            Top             =   1950
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   503
            DibPicture      =   "frmFaturamento_Impostos_produto.frx":2DA56
            Caption         =   ""
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
            Theme           =   4
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Aliquota do ICMS ST"
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
            Index           =   1
            Left            =   270
            TabIndex        =   11
            Top             =   2040
            Width           =   1815
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Valor do ICMS ST"
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
            Left            =   270
            TabIndex        =   10
            Top             =   1692
            Width           =   1815
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Valor Base de cálculo ICMS ST"
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
            Left            =   270
            TabIndex        =   9
            Top             =   1344
            Width           =   3045
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "% Redução Base de cálculo ICMS ST"
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
            Left            =   270
            TabIndex        =   8
            Top             =   996
            Width           =   4245
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "% Margem substituição tributária"
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
            Left            =   270
            TabIndex        =   7
            Top             =   648
            Width           =   3945
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Modalidade determinação Substituição tributária"
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
            Left            =   270
            TabIndex        =   6
            Top             =   300
            Width           =   4065
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "CST Icms"
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
         Height          =   2835
         Index           =   1
         Left            =   150
         TabIndex        =   3
         Top             =   270
         Width           =   5445
         Begin DrawSuite2022.USButton btnOrigem 
            Height          =   285
            Left            =   4980
            TabIndex        =   85
            ToolTipText     =   "Informações gerais sobre a origem da mercadoria"
            Top             =   240
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   503
            DibPicture      =   "frmFaturamento_Impostos_produto.frx":3645B
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
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
            ShowFocusRect   =   0   'False
            Theme           =   1
         End
         Begin VB.TextBox txtICMSValor 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3960
            TabIndex        =   51
            TabStop         =   0   'False
            Top             =   2340
            Width           =   1005
         End
         Begin VB.TextBox txtICMSPercentual 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3960
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   1990
            Width           =   1005
         End
         Begin VB.TextBox txtICMSVlrBaseCal 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3960
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   1640
            Width           =   1005
         End
         Begin VB.TextBox txtICMSPercentualRedBaseCalc 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3960
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   1290
            Width           =   1005
         End
         Begin VB.TextBox txtICMSModalidade 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3960
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   940
            Width           =   1005
         End
         Begin VB.TextBox txtICMSCST 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3960
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   590
            Width           =   1005
         End
         Begin VB.TextBox txtICMSOrigem 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3960
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   240
            Width           =   1005
         End
         Begin DrawSuite2022.USButton btnCSTICMS 
            Height          =   285
            Left            =   4980
            TabIndex        =   86
            ToolTipText     =   "Informações gerais sobre a tributação do ICMS"
            Top             =   585
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   503
            DibPicture      =   "frmFaturamento_Impostos_produto.frx":3D5EE
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
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
            ShowFocusRect   =   0   'False
            Theme           =   1
         End
         Begin DrawSuite2022.USButton btnSalvar 
            Height          =   285
            Left            =   4980
            TabIndex        =   95
            ToolTipText     =   "Gravar valores"
            Top             =   2340
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   503
            DibPicture      =   "frmFaturamento_Impostos_produto.frx":44781
            Caption         =   ""
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
            Theme           =   4
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "Valor do ICMS"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   240
            TabIndex        =   80
            Top             =   2370
            Width           =   1815
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
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
            Height          =   225
            Left            =   240
            TabIndex        =   79
            Top             =   300
            Width           =   585
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "CST"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   240
            TabIndex        =   78
            Top             =   645
            Width           =   585
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Modalidade"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   240
            TabIndex        =   77
            Top             =   990
            Width           =   855
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "% Redução Base de cálculo"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   240
            TabIndex        =   76
            Top             =   1335
            Width           =   2175
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "% do ICMS"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   240
            TabIndex        =   75
            Top             =   2025
            Width           =   1815
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Valor Base de cáculo"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   240
            TabIndex        =   4
            Top             =   1680
            Width           =   1815
         End
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Impostos do COFINS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   3135
      Left            =   7740
      TabIndex        =   44
      Top             =   6240
      Width           =   3765
      Begin VB.TextBox txtCONFINSValor 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2340
         TabIndex        =   74
         TabStop         =   0   'False
         Top             =   1575
         Width           =   1005
      End
      Begin VB.TextBox txtCONFINSBaseCalc 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2340
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   1170
         Width           =   1005
      End
      Begin VB.TextBox txtCONFINSPercentual 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2340
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   765
         Width           =   1005
      End
      Begin VB.TextBox txtCONFINSCST 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2340
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   360
         Width           =   1005
      End
      Begin DrawSuite2022.USButton btnCSTCIFINS 
         Height          =   285
         Left            =   3360
         TabIndex        =   89
         ToolTipText     =   "Informações gerais sobre a tributação do COFINS"
         Top             =   360
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   503
         DibPicture      =   "frmFaturamento_Impostos_produto.frx":4D186
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
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
         ShowFocusRect   =   0   'False
         Theme           =   1
      End
      Begin DrawSuite2022.USButton btnSalvarCofins 
         Height          =   285
         Left            =   3360
         TabIndex        =   98
         ToolTipText     =   "Gravar valores"
         Top             =   1580
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   503
         DibPicture      =   "frmFaturamento_Impostos_produto.frx":54319
         Caption         =   ""
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
         Theme           =   4
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Valor do COFINS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   1
         Left            =   180
         TabIndex        =   61
         Top             =   1620
         Width           =   2265
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "CST"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   1
         Left            =   180
         TabIndex        =   60
         Top             =   420
         Width           =   585
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "% COFINS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   1
         Left            =   180
         TabIndex        =   59
         Top             =   795
         Width           =   2175
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Valor base de cálculo COFINS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   1
         Left            =   180
         TabIndex        =   58
         Top             =   1200
         Width           =   2775
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Impostos do IPI"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   2535
      Index           =   0
      Left            =   180
      TabIndex        =   2
      Top             =   6240
      Width           =   3765
      Begin VB.TextBox txtIPICondigoEnq 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2340
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   1980
         Width           =   1005
      End
      Begin VB.TextBox txtIPIValor 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2340
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   1575
         Width           =   1005
      End
      Begin VB.TextBox txtIPIBaseCalc 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2340
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   1170
         Width           =   1005
      End
      Begin VB.TextBox txtIPIPercentual 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2340
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   765
         Width           =   1005
      End
      Begin VB.TextBox txtIPICST 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2340
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   360
         Width           =   1005
      End
      Begin DrawSuite2022.USButton btnCSTIPI 
         Height          =   285
         Left            =   3360
         TabIndex        =   87
         ToolTipText     =   "Informações gerais sobre a tributação do IPI"
         Top             =   360
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   503
         DibPicture      =   "frmFaturamento_Impostos_produto.frx":5CD1E
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
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
         ShowFocusRect   =   0   'False
         Theme           =   1
      End
      Begin DrawSuite2022.USButton btnCodEnqIPI 
         Height          =   285
         Left            =   3360
         TabIndex        =   94
         ToolTipText     =   "Informações gerais sobre a tributação do IPI"
         Top             =   1980
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   503
         DibPicture      =   "frmFaturamento_Impostos_produto.frx":63EB1
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
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
         ShowFocusRect   =   0   'False
         Theme           =   1
      End
      Begin DrawSuite2022.USButton btnSalvarIPI 
         Height          =   285
         Left            =   3360
         TabIndex        =   96
         ToolTipText     =   "Gravar valores"
         Top             =   1570
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   503
         DibPicture      =   "frmFaturamento_Impostos_produto.frx":6B044
         Caption         =   ""
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
         Theme           =   4
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Código enquadramento IPI"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   150
         TabIndex        =   53
         Top             =   2010
         Width           =   2265
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Valor  do IPI"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   0
         Left            =   150
         TabIndex        =   52
         Top             =   1605
         Width           =   2265
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "CST"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   0
         Left            =   150
         TabIndex        =   50
         Top             =   390
         Width           =   585
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "% IPI"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   0
         Left            =   150
         TabIndex        =   49
         Top             =   795
         Width           =   2175
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Valor base de cálculo do IPI"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   0
         Left            =   150
         TabIndex        =   48
         Top             =   1200
         Width           =   2265
      End
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11790
      _ExtentX        =   20796
      _ExtentY        =   688
      DibPicture      =   "frmFaturamento_Impostos_produto.frx":73A49
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
      Icon            =   "frmFaturamento_Impostos_produto.frx":7DB6C
   End
End
Attribute VB_Name = "frmFaturamento_Impostos_produto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnCSTCIFINS_Click()
On Error GoTo tratar_erro

Select Case txtCONFINSCST.Text
        Case "01": Texto = "01 - Operação Tributável (base de cálculo = valor da operação alíquota normal (cumulativo/não cumulativo))"
        Case "02": Texto = "02 - Operação Tributável (base de cálculo = valor da operação (alíquota diferenciada))"
        Case "03": Texto = "03 - Operação Tributável (base de cálculo = quantidade vendida x alíquota por unidade de produto)"
        Case "04": Texto = "04 - Operação Tributável (tributação monofásica (alíquota zero))"
        Case "06": Texto = "06 - Operação Tributável (alíquota zero)"
        Case "07": Texto = "07 - Operação Isenta da Contribuição"
        Case "08": Texto = "08 - Operação Sem Incidência da Contribuição"
        Case "09": Texto = "09 - Operação com Suspensão da Contribuição"
        Case "49": Texto = "49 - Outras Operações de Saída"
        Case "50": Texto = "50 - Operação com Direito a Crédito - Vinculada Exclusivamente a Receita Tributada no Mercado Interno"
        Case "51": Texto = "51 - Operação com Direito a Crédito - Vinculada Exclusivamente a Receita Não Tributada no Mercado Interno"
        Case "52": Texto = "52 - Operação com Direito a Crédito - Vinculada Exclusivamente a Receita de Exportação"
        Case "53": Texto = "53 - Operação com Direito a Crédito - Vinculada a Receitas Tributadas e Não-Tributadas no Mercado Interno"
        Case "54": Texto = "54 - Operação com Direito a Crédito - Vinculada a Receitas Tributadas no Mercado Interno e de Exportação"
        Case "55": Texto = "55 - Operação com Direito a Crédito - Vinculada a Receitas Não-Tributadas no Mercado Interno e de Exportação"
        Case "56": Texto = "56 - Operação com Direito a Crédito - Vinculada a Receitas Tributadas e Não-Tributadas no Mercado Interno, e de Exportação"
        Case "60": Texto = "60 - Crédito Presumido - Operação de Aquisição Vinculada Exclusivamente a Receita Tributada no Mercado Interno"
        Case "61": Texto = "61 - Crédito Presumido - Operação de Aquisição Vinculada Exclusivamente a Receita Não Tributada no Mercado Interno"
        Case "62": Texto = "62 - Crédito Presumido - Operação de Aquisição Vinculada Exclusivamente a Receita de Exportação"
        Case "63": Texto = "63 - Crédito Presumido - Operação de Aquisição Vinculada a Receitas Tributadas e Não-Tributadas no Mercado Interno"
        Case "64": Texto = "64 - Crédito Presumido - Operação de Aquisição Vinculada a Receitas Tributadas no Mercado Interno e de Exportação"
        Case "65": Texto = "65 - Crédito Presumido - Operação de Aquisição Vinculada a Receitas Não-Tributadas no Mercado Interno e de Exportação"
        Case "66": Texto = "66 - Crédito Presumido - Operação de Aquisição Vinculada a Receitas Tributadas e Não-Tributadas no Mercado Interno, e de Exportação"
        Case "67": Texto = "67 - Crédito Presumido - Outras Operações"
        Case "70": Texto = "70 - Operação de Aquisição sem Direito a Crédito"
        Case "71": Texto = "71 - Operação de Aquisição com Isenção"
        Case "72": Texto = "72 - Operação de Aquisição com Suspensão"
        Case "73": Texto = "73 - Operação de Aquisição a Alíquota Zero"
        Case "74": Texto = "74 - Operação de Aquisição sem Incidência da Contribuição"
        Case "75": Texto = "75 - Operação de Aquisição por Substituição Tributária"
        Case "98": Texto = "98 - Outras Operações de Entrada"
        Case "99": Texto = "99 - Outras Operações"
End Select

USMsgBox "CST DO COFINS " & Texto, vbInformation, "CAPRIND v5.0"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub btnCSTICMS_Click()
On Error GoTo tratar_erro

Select Case txtICMSCST.Text
Case "00": Texto = "00 - Tributada integralmente"
Case "10": Texto = "10 - Tributada e com cobrança do ICMS por substituição"
Case "20": Texto = "20 - Com redução de base de cálculo"
Case "40": Texto = "40 - Isenta"
Case "41": Texto = "41 - Não tributada"
Case "50": Texto = "50 - Suspensão"
Case "51": Texto = "51 - Diferimento"
Case "60": Texto = "60 - ICMS cobrado anteriormente por substituição tributária"
Case "70": Texto = "70 - Com redução de base de cálculo e cobrança do ICMS por substituição tributária"
Case "90": Texto = "90 - Outras"
Case "101": Texto = "101 - Tributada pelo Simples Nacional com permissão de crédito"
Case "102": Texto = "102 - Tributada pelo Simples Nacional sem permissão de crédito"
Case "103": Texto = "103 - Isenção do ICMS no Simples Nacional para faixa de receita bruta"
Case "201": Texto = "201 - Tributada pelo Simples Nacional com permissão de crédito e com cobrança do ICMS por Substituição Tributária"
Case "202": Texto = "202 - Tributada pelo Simples Nacional sem permissão de crédito e com cobrança do ICMS por Substituição Tributária"
Case "203": Texto = "203 - Isenção do ICMS nos Simples Nacional para faixa de receita bruta e com cobrança do ICMS por Substituição Tributária"
Case "300": Texto = "300 - Imune"
Case "400": Texto = "400 - Não tributada pelo Simples Nacional"
Case "500": Texto = "500 - ICMS cobrado anteriormente por substituição tributária (substituído) ou por antecipação"
Case "900": Texto = "900 - Outros"
End Select

USMsgBox "CST DO ICMS " & Texto, vbInformation, "CAPRIND v5.0"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub btnCSTIPI_Click()
On Error GoTo tratar_erro

Select Case txtIPICST.Text
Case "00": Texto = "00 - Entrada com recuperação de crédito"
Case "01": Texto = "01 - Entrada tributada com alíquota zero"
Case "02": Texto = "02 - Entrada isenta"
Case "03": Texto = "03 - Entrada não-tributada"
Case "04": Texto = "04 - Entrada imune"
Case "05": Texto = "05 - Entrada com suspensão"
Case "49": Texto = "49 - Outras entradas"
Case "50": Texto = "50 - Saída tributada"
Case "51": Texto = "51 - Saída tributada com alíquota zero"
Case "52": Texto = "52 - Saída isenta"
Case "53": Texto = "53 - Saída não-tributada"
Case "54": Texto = "54 - Saída imune"
Case "55": Texto = "55 - Saída com suspensão"
Case "99": Texto = "99 - Outras saídas"
End Select

USMsgBox "CST DO IPI " & Texto, vbInformation, "CAPRIND v5.0"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub btnCSTPIS_Click()
On Error GoTo tratar_erro

Select Case txtPISCST.Text
        Case "01": Texto = "01 - Operação Tributável (base de cálculo = valor da operação alíquota normal (cumulativo/não cumulativo))"
        Case "02": Texto = "02 - Operação Tributável (base de cálculo = valor da operação (alíquota diferenciada))"
        Case "03": Texto = "03 - Operação Tributável (base de cálculo = quantidade vendida x alíquota por unidade de produto)"
        Case "04": Texto = "04 - Operação Tributável (tributação monofásica (alíquota zero))"
        Case "05": Texto = "05 - Operação Tributável (Substituição Tributária)"
        Case "06": Texto = "06 - Operação Tributável (alíquota zero)"
        Case "07": Texto = "07 - Operação Isenta da Contribuição"
        Case "08": Texto = "08 - Operação Sem Incidência da Contribuição"
        Case "09": Texto = "09 - Operação com Suspensão da Contribuição"
        Case "49": Texto = "49 - Outras Operações de Saída"
        Case "50": Texto = "50 - Operação com Direito a Crédito - Vinculada Exclusivamente a Receita Tributada no Mercado Interno"
        Case "51": Texto = "51 - Operação com Direito a Crédito - Vinculada Exclusivamente a Receita Não Tributada no Mercado Interno"
        Case "52": Texto = "52 - Operação com Direito a Crédito - Vinculada Exclusivamente a Receita de Exportação"
        Case "53": Texto = "53 - Operação com Direito a Crédito - Vinculada a Receitas Tributadas e Não-Tributadas no Mercado Interno"
        Case "54": Texto = "54 - Operação com Direito a Crédito - Vinculada a Receitas Tributadas no Mercado Interno e de Exportação"
        Case "55": Texto = "55 - Operação com Direito a Crédito - Vinculada a Receitas Não-Tributadas no Mercado Interno e de Exportação"
        Case "56": Texto = "56 - Operação com Direito a Crédito - Vinculada a Receitas Tributadas e Não-Tributadas no Mercado Interno, e de Exportação"
        Case "60": Texto = "60 - Crédito Presumido - Operação de Aquisição Vinculada Exclusivamente a Receita Tributada no Mercado Interno"
        Case "61": Texto = "61 - Crédito Presumido - Operação de Aquisição Vinculada Exclusivamente a Receita Não Tributada no Mercado Interno"
        Case "62": Texto = "62 - Crédito Presumido - Operação de Aquisição Vinculada Exclusivamente a Receita de Exportação"
        Case "63": Texto = "63 - Crédito Presumido - Operação de Aquisição Vinculada a Receitas Tributadas e Não-Tributadas no Mercado Interno"
        Case "64": Texto = "64 - Crédito Presumido - Operação de Aquisição Vinculada a Receitas Tributadas no Mercado Interno e de Exportação"
        Case "65": Texto = "65 - Crédito Presumido - Operação de Aquisição Vinculada a Receitas Não-Tributadas no Mercado Interno e de Exportação"
        Case "66": Texto = "66 - Crédito Presumido - Operação de Aquisição Vinculada a Receitas Tributadas e Não-Tributadas no Mercado Interno, e de Exportação"
        Case "67": Texto = "67 - Crédito Presumido - Outras Operações"
        Case "70": Texto = "70 - Operação de Aquisição sem Direito a Crédito"
        Case "71": Texto = "71 - Operação de Aquisição com Isenção"
        Case "72": Texto = "72 - Operação de Aquisição com Suspensão"
        Case "73": Texto = "73 - Operação de Aquisição a Alíquota Zero"
        Case "74": Texto = "74 - Operação de Aquisição sem Incidência da Contribuição"
        Case "75": Texto = "75 - Operação de Aquisição por Substituição Tributária"
        Case "98": Texto = "98 - Outras Operações de Entrada"
        Case "99": Texto = "99 - Outras Operações"
End Select

USMsgBox "CST DO PIS " & Texto, vbInformation, "CAPRIND v5.0"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub btnIPIdevolucao_Click()
On Error GoTo tratar_erro


Set TBAliquota = CreateObject("adodb.recordset")
TBAliquota.Open "Select * from tbl_Detalhes_Nota_CST_IPI where ID_item = " & IDAntigo, Conexao, adOpenKeyset, adLockOptimistic
If TBAliquota.EOF = False Then
TBAliquota!pIPIdevolv = txtpIPIdevolv.Text
TBAliquota!vIPIdevolv = txtvIPIdevolv.Text
TBAliquota.Update
End If

TBAliquota.Close

Set TBAliquota = CreateObject("adodb.recordset")
StrSql = "Select sum(IPI.vIPIdevolv) as vTotalIPIdevolv from tbl_Detalhes_Nota DTN inner join tbl_Detalhes_Nota_CST_IPI IPI on IPI.ID_Item = DTN.int_Codigo where DTN.ID_Nota  = " & ID_nota
'Debug.print StrSql

TBAliquota.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
If TBAliquota.EOF = False Then
vTotalIPIdevolv = TBAliquota!vTotalIPIdevolv
End If

TBAliquota.Close

Conexao.Execute "Update tbl_totais_Nota set Total_IPI_devolv = " & Replace(vTotalIPIdevolv, ",", ".") & " Where ID_Nota = '" & ID_nota & "'"
'TBAliquota.Close

USMsgBox "Dados gravados com sucesso!", vbInformation, "CAPRIND v5.0"

frmFaturamento_Prod_Serv.ProcCarregaLista


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"

End Sub

Private Sub btnOrigem_Click()
On Error GoTo tratar_erro

Select Case txtICMSOrigem.Text
Case 0: Texto = "0 - Nacional"
Case 1: Texto = "1 - Estrangeira - Importação direta"
Case 2: Texto = "2 - Estrangeira - Adquirida no mercado interno"
Case 3: Texto = "3 - Nacional - Mercadoria ou bem com Conteúdo de Importação superior a 40% (quarenta por cento)"
Case 4: Texto = "4 - Nacional - Cuja produção tenha sido feita em conformidade com os processos produtivos básicos"
Case 5: Texto = "5 - Nacional - Mercadoria ou bem com Conteúdo de Importação inferior ou igual a 40% (quarenta por cento)"
Case 6: Texto = "6 - Estrangeira - Importação direta, sem similar nacional, constante em lista de Resolução CAMEX"
Case 7: Texto = "7 - Estrangeira - Adquirida no mercado interno, sem similar nacional, constante em lista de Resolução CAMEX"
Case 8: Texto = "8 - Nacional - Mercadoria ou bem com Conteúdo de Importação superior a 70%"
End Select
USMsgBox "Origem da mercadoria " & Texto, vbInformation, "CAPRIND v5.0"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub btnSalvar_Click()
On Error GoTo tratar_erro


Set TBAliquota = CreateObject("adodb.recordset")
TBAliquota.Open "Select * from tbl_Detalhes_Nota_CST_ICMS where ID_item = " & IDAntigo, Conexao, adOpenKeyset, adLockOptimistic
If TBAliquota.EOF = False Then

TBAliquota!Origem_mercadoria = txtICMSOrigem.Text
TBAliquota!Modalidade_determinacao = txtICMSModalidade.Text
TBAliquota!ICMS_SN = txtICMSPercentual.Text
TBAliquota!Valor_ICMS_SN = txtICMSValor.Text
TBAliquota!Tributacao_ICMS = txtICMSCST.Text
TBAliquota!Valor_ICMS = txtICMSValor.Text
TBAliquota!Percentual_reducao_BC = txtICMSPercentualRedBaseCalc.Text
TBAliquota!Valor_BC = txtICMSVlrBaseCal.Text
TBAliquota.Update
End If

Conexao.Execute "Update tbl_detalhes_Nota set int_ICMS = " & Replace(txtICMSPercentual.Text, ",", ".") & " Where int_codigo = '" & IDAntigo & "'"
TBAliquota.Close

USMsgBox "Dados gravados com sucesso!", vbInformation, "CAPRIND v5.0"

NotaFiscalPronta = False

If frmFaturamento_Prod_Serv.cmbFinalidade_emissao <> "2 - Complementar" Then
    frmFaturamento_Prod_Serv.ProcCarregaLista
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub btnSalvarCofins_Click()
On Error GoTo tratar_erro


Set TBAliquota = CreateObject("adodb.recordset")
TBAliquota.Open "Select * from tbl_Detalhes_Nota_CST_COFINS where ID_item = " & IDAntigo, Conexao, adOpenKeyset, adLockOptimistic
If TBAliquota.EOF = False Then
TBAliquota!Valor_BC = txtCONFINSBaseCalc.Text
TBAliquota.Update
End If

TBAliquota.Close

Set TBAliquota = CreateObject("adodb.recordset")
TBAliquota.Open "Select * from tbl_Detalhes_Nota where Int_codigo = " & IDAntigo, Conexao, adOpenKeyset, adLockOptimistic
If TBAliquota.EOF = False Then
TBAliquota!Total_Cofins_prod = txtCONFINSValor.Text
TBAliquota.Update
End If

TBAliquota.Close

USMsgBox "Dados gravados com sucesso!", vbInformation, "CAPRIND v5.0"


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub btnSalvarIPI_Click()
On Error GoTo tratar_erro


Set TBAliquota = CreateObject("adodb.recordset")
TBAliquota.Open "Select * from tbl_Detalhes_Nota_CST_IPI where ID_item = " & IDAntigo, Conexao, adOpenKeyset, adLockOptimistic
If TBAliquota.EOF = False Then
TBAliquota!Valor_BC = txtIPIBaseCalc.Text
TBAliquota.Update
End If

TBAliquota.Close

Set TBAliquota = CreateObject("adodb.recordset")
TBAliquota.Open "Select * from tbl_Detalhes_Nota where Int_codigo = " & IDAntigo, Conexao, adOpenKeyset, adLockOptimistic
If TBAliquota.EOF = False Then
TBAliquota!dbl_valoripi = txtIPIValor.Text
TBAliquota.Update
End If

Conexao.Execute "Update tbl_detalhes_Nota set int_IPI = " & Replace(txtIPIPercentual.Text, ",", ".") & ", dbl_valorIPI = " & Replace(txtIPIValor.Text, ",", ".") & ",  CST_IPI = " & Replace(txtIPICST.Text, ",", ".") & "  Where int_codigo = '" & IDAntigo & "'"
TBAliquota.Close

USMsgBox "Dados gravados com sucesso!", vbInformation, "CAPRIND v5.0"

frmFaturamento_Prod_Serv.ProcCarregaLista


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub btnSalvarPIS_Click()
On Error GoTo tratar_erro


Set TBAliquota = CreateObject("adodb.recordset")
TBAliquota.Open "Select * from tbl_Detalhes_Nota_CST_PIS where ID_item = " & IDAntigo, Conexao, adOpenKeyset, adLockOptimistic
If TBAliquota.EOF = False Then
TBAliquota!Valor_BC = txtPISBaseCalc.Text
TBAliquota.Update
End If

TBAliquota.Close

Set TBAliquota = CreateObject("adodb.recordset")
TBAliquota.Open "Select * from tbl_Detalhes_Nota where Int_codigo = " & IDAntigo, Conexao, adOpenKeyset, adLockOptimistic
If TBAliquota.EOF = False Then
TBAliquota!Total_PIS_prod = txtPISValor.Text
TBAliquota.Update
End If

TBAliquota.Close

USMsgBox "Dados gravados com sucesso!", vbInformation, "CAPRIND v5.0"


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub btnSalvarST_Click()
On Error GoTo tratar_erro


Set TBAliquota = CreateObject("adodb.recordset")
TBAliquota.Open "Select * from tbl_Detalhes_Nota_CST_ICMS where ID_item = " & IDAntigo, Conexao, adOpenKeyset, adLockOptimistic
If TBAliquota.EOF = False Then

TBAliquota!Origem_mercadoria = txtICMSOrigem.Text
TBAliquota!Modalidade_determinacao = txtICMSModalidade.Text
TBAliquota!ICMS_SN = txtICMSPercentual.Text
TBAliquota!Valor_ICMS_SN = txtICMSValor.Text
TBAliquota!Tributacao_ICMS = txtICMSCST.Text
TBAliquota!Valor_ICMS = txtICMSValor.Text
TBAliquota!Percentual_reducao_BC = txtICMSPercentualRedBaseCalc.Text
TBAliquota!Valor_BC = txtICMSVlrBaseCal.Text
TBAliquota.Update
End If

Conexao.Execute "Update tbl_detalhes_Nota set int_ICMS = " & Replace(txtICMSPercentual.Text, ",", ".") & " Where int_codigo = '" & IDAntigo & "'"
TBAliquota.Close

USMsgBox "Dados gravados com sucesso!", vbInformation, "CAPRIND v5.0"

NotaFiscalPronta = False
frmFaturamento_Prod_Serv.ProcCarregaLista

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub btnSubstituicaotributaria_Click()
On Error GoTo tratar_erro

frmFaturamento_Impostos_Substituicao.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

StrSql = "SELECT DN.Int_codigo AS ID_ITEM,DN.txt_CST AS ICMSCST, DN.int_ICMS AS PICMS, DN.int_IPI AS PIPI, DN.dbl_ValorIPI AS VIPI, DN.PIS_prod AS PPIS, DN.Total_PIS_prod AS VPIS, DN.Cofins_prod AS PCOFINS, DN.Total_Cofins_prod AS VCOFINS," _
                         & "DN.ICMS_suframa AS VICMSSUFRAMA, DN.CST_IPI AS IPICST, DN.CST_PIS AS PISCST, DN.CST_Cofins AS COFINSCST, DN.ICMS_SN AS ICMSSN, DN.Valor_desconto_SUFRAMA AS DESCSUFFRAMA, DN.Codigo_enquadramento_IPI," _
                         & "ICMS.Origem_mercadoria, ICMS.Tributacao_ICMS, ICMS.Modalidade_determinacao, ICMS.Percentual_reducao_BC, ICMS.Valor_BC, ICMS.Valor_ICMS, ICMS.Modalidade_determinacao_ST," _
                         & "ICMS.Percentual_margem_ICMS_ST, ICMS.Percentual_reducao_BC_ST, ICMS.Valor_BC_ST, ICMS.Aliquota_imposto_ST, ICMS.Valor_ICMS_ST, ICMS.ICMS_SN, ICMS.Valor_ICMS_SN," _
                         & "ICMS.Valor_ICMS_desonerado, ICMS.Motivo_ICMS_desonerado, ICMS.Percentual_ICMS_DIF, ICMS.Valor_ICMS_DIF, ICMS.Valor_BC_ICMS_UF_dest, ICMS.Percentual_provisorio, ICMS.Valor_ICMS_INT_UF_dest," _
                         & "ICMS.Valor_ICMS_INT_UF_rem, ICMS.Percentual_FCP, ICMS.Valor_ICMS_FCP, IPI.Codigo_situacaoTributaria, IPI.Valor_BC AS IPIBC, IPI.pIPIdevolv as pIPIdevolv, IPI.vIPIdevolv as vIPIdevolv, PIS.Valor_BC AS PISValor_BC , COFINS.Valor_BC AS COFINSValor_BC, DN.ID_CFOP" _
& " FROM dbo.tbl_Detalhes_Nota AS DN INNER JOIN " _
                         & "dbo.tbl_Detalhes_Nota_CST_ICMS AS ICMS ON DN.Int_codigo = ICMS.ID_item INNER JOIN " _
                         & "dbo.tbl_Detalhes_Nota_CST_IPI AS IPI ON ICMS.ID_item = IPI.Id_item INNER JOIN " _
                         & "dbo.tbl_Detalhes_Nota_CST_PIS AS PIS ON ICMS.ID_item = PIS.ID_item INNER JOIN " _
                         & "dbo.tbl_Detalhes_Nota_CST_Cofins AS COFINS ON ICMS.ID_item = COFINS.ID_item"
'Debug.print StrSql

Set TBAliquota = CreateObject("adodb.recordset")
TBAliquota.Open StrSql & " where DN.Int_codigo = " & IDAntigo, Conexao, adOpenKeyset, adLockOptimistic
If TBAliquota.EOF = False Then

'=================================================================================
' Verifica se a CFOP tem Suframa
'=================================================================================
Set TBCFOP = CreateObject("adodb.recordset")
TBCFOP.Open "Select Suframa from tbl_NaturezaOperacao where IDCountCfop = " & TBAliquota!ID_CFOP, Conexao, adOpenKeyset, adLockOptimistic
If TBCFOP.EOF = False Then
CFOPDesonerada.Value = IIf(TBCFOP!Suframa = False, 0, 1)
End If
TBCFOP.Close

'=================================================================================
' ICMS tributação
'=================================================================================
txtICMSOrigem.Text = TBAliquota!Origem_mercadoria
txtICMSModalidade.Text = IIf(IsNull(TBAliquota!Modalidade_determinacao), 0, TBAliquota!Modalidade_determinacao)
If Len(TBAliquota!ICMSCST) = 4 Then
Frame2(1).Caption = "CSOSN"
Label2.Caption = "CSOSN"
txtICMSCST.Text = Right(TBAliquota!ICMSCST, 3)
txtICMSPercentual.Text = IIf(IsNull(TBAliquota!ICMS_SN), 0, TBAliquota!ICMS_SN)
txtICMSValor.Text = IIf(IsNull(TBAliquota!Valor_ICMS_SN), 0, Format(TBAliquota!Valor_ICMS_SN, "###,##0.00"))
Else
Frame2(1).Caption = "CST"
Label2.Caption = "CST"
txtICMSCST.Text = Right(TBAliquota!ICMSCST, 2)
txtICMSPercentual.Text = IIf(IsNull(TBAliquota!pICMS), 0, TBAliquota!pICMS)
txtICMSValor.Text = IIf(IsNull(TBAliquota!Valor_ICMS), 0, Format(TBAliquota!Valor_ICMS, "###,##0.00"))
End If
txtICMSPercentualRedBaseCalc.Text = IIf(IsNull(TBAliquota!Percentual_reducao_BC), 0, TBAliquota!Percentual_reducao_BC)
txtICMSVlrBaseCal.Text = IIf(IsNull(TBAliquota!Valor_BC), 0, Format(TBAliquota!Valor_BC, "###,##0.00")) '

'=================================================================================
' ICMS substituição tributária
'=================================================================================
txtICMSSTModalidade.Text = IIf(IsNull(TBAliquota!Modalidade_determinacao_ST), 0, TBAliquota!Modalidade_determinacao_ST)
txtICMSSTMargem.Text = IIf(IsNull(TBAliquota!Percentual_margem_ICMS_ST), 0, TBAliquota!Percentual_margem_ICMS_ST)
txtICMSSTPercentualRed.Text = IIf(IsNull(TBAliquota!Percentual_reducao_BC_ST), 0, TBAliquota!Percentual_reducao_BC_ST)
txtICMSSTValorBaseCal.Text = IIf(IsNull(TBAliquota!Valor_BC_ST), 0, Format(TBAliquota!Valor_BC_ST, "###,##0.00"))
txtICMSSTValor.Text = IIf(IsNull(TBAliquota!Valor_ICMS_ST), 0, Format(TBAliquota!Valor_ICMS_ST, "###,##0.00"))
txtICMSSTAliquota.Text = IIf(IsNull(TBAliquota!Aliquota_imposto_ST), 0, TBAliquota!Aliquota_imposto_ST)
'=================================================================================
' ICMS DIFAL
'=================================================================================
txtICMSDIFPercentual.Text = IIf(IsNull(TBAliquota!Percentual_ICMS_DIF), 0, TBAliquota!Percentual_ICMS_DIF)
txtICMSDIFValor.Text = IIf(IsNull(TBAliquota!Valor_ICMS_DIF), 0, Format(TBAliquota!Valor_ICMS_DIF, "###,##0.00"))
txtICMSDIFBaseUFDestino.Text = IIf(IsNull(TBAliquota!Valor_BC_ICMS_UF_dest), 0, Format(TBAliquota!Valor_BC_ICMS_UF_dest, "###,##0.00"))
txtICMSDIFValorUFDest.Text = IIf(IsNull(TBAliquota!Valor_ICMS_INT_UF_dest), 0, Format(TBAliquota!Valor_ICMS_INT_UF_dest, "###,##0.00"))
txtICMSDIFValorUFRemetente.Text = IIf(IsNull(TBAliquota!Valor_ICMS_INT_UF_rem), 0, Format(TBAliquota!Valor_ICMS_INT_UF_rem, "###,##0.00"))
txtICMSDIFPercentualProvisorio.Text = IIf(IsNull(TBAliquota!Percentual_provisorio), 0, TBAliquota!Percentual_provisorio)
txtICMSDIFPercentualFCP.Text = IIf(IsNull(TBAliquota!Percentual_FCP), 0, TBAliquota!Percentual_FCP)
txtICMSDIFValorFCP.Text = IIf(IsNull(TBAliquota!Valor_ICMS_FCP), 0, Format(TBAliquota!Valor_ICMS_FCP, "###,##0.00"))
'=================================================================================
' ICMS SUFRAMA
'=================================================================================
txtICMSDesonMotivo.Text = IIf(IsNull(TBAliquota!Motivo_ICMS_desonerado), 0, TBAliquota!Motivo_ICMS_desonerado)
txtICMSDesonValor.Text = IIf(IsNull(TBAliquota!Valor_ICMS_desonerado), 0, Format(TBAliquota!Valor_ICMS_desonerado, "###,##0.00"))
'=================================================================================
' IPI IMPOSTOS
'=================================================================================
txtIPICST.Text = IIf(IsNull(TBAliquota!IPICST), 0, TBAliquota!IPICST)
txtIPIPercentual.Text = IIf(IsNull(TBAliquota!pIPI), 0, TBAliquota!pIPI)
txtIPIBaseCalc.Text = IIf(IsNull(TBAliquota!IPIBC), 0, Format(TBAliquota!IPIBC, "###,##0.00"))
txtIPIValor.Text = IIf(IsNull(TBAliquota!vIPI), 0, Format(TBAliquota!vIPI, "###,##0.00"))
txtIPICondigoEnq.Text = IIf(IsNull(TBAliquota!Codigo_enquadramento_IPI), 0, TBAliquota!Codigo_enquadramento_IPI)

txtpIPIdevolv.Text = IIf(IsNull(TBAliquota!pIPIdevolv), 0, TBAliquota!pIPIdevolv)
txtvIPIdevolv.Text = IIf(IsNull(TBAliquota!vIPIdevolv), 0, TBAliquota!vIPIdevolv)

'=================================================================================
' PIS IMPOSTOS
'=================================================================================
txtPISCST.Text = IIf(IsNull(TBAliquota!PISCST), 0, TBAliquota!PISCST)
txtPISPercentual.Text = IIf(IsNull(TBAliquota!pPIS), 0, TBAliquota!pPIS)
txtPISBaseCalc.Text = IIf(IsNull(TBAliquota!PISValor_BC), 0, Format(TBAliquota!PISValor_BC, "###,##0.00"))
txtPISValor.Text = IIf(IsNull(TBAliquota!vPIS), 0, Format(TBAliquota!vPIS, "###,##0.00"))
'=================================================================================
' COFINS IMPOSTOS
'=================================================================================
txtCONFINSCST.Text = IIf(IsNull(TBAliquota!COFINSCST), 0, TBAliquota!COFINSCST)
txtCONFINSPercentual.Text = IIf(IsNull(TBAliquota!pCofins), 0, TBAliquota!pCofins)
txtCONFINSBaseCalc.Text = IIf(IsNull(TBAliquota!COFINSValor_BC), 0, Format(TBAliquota!COFINSValor_BC, "###,##0.00"))
txtCONFINSValor.Text = IIf(IsNull(TBAliquota!vCOFINS), 0, Format(TBAliquota!vCOFINS, "###,##0.00"))
End If
TBAliquota.Close

Set TBClientes = CreateObject("adodb.recordset")
TBClientes.Open "Select chkSuframa from Clientes where IDCliente = " & frmFaturamento_Prod_Serv.txtIDcliente, Conexao, adOpenKeyset, adLockOptimistic
If TBClientes.EOF = False Then
Suframa.Value = IIf(TBClientes!chkSuframa = False, 0, 1)
End If
TBClientes.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub USButton3_Click()

End Sub
