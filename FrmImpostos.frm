VERSION 5.00
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Begin VB.Form FrmImpostos 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Administrativo - Vendas - Proposta comercial  - Impostos"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11730
   Icon            =   "FrmImpostos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   11730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USButton Cmd_salvar_tabelaSN 
      Height          =   765
      Left            =   10710
      TabIndex        =   56
      Top             =   6030
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1349
      Caption         =   "Salvar (F3)"
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
      BorderColorDisabled=   0
      BorderColorDown =   15048022
      BorderColorOver =   15381630
      GradientColor2  =   14737632
      GradientColor3  =   12632256
      GradientColor4  =   12632256
      PicSizeH        =   48
      PicSizeW        =   48
      Theme           =   1
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Valor aproximado de tributos"
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
      Height          =   765
      Left            =   5910
      TabIndex        =   117
      Top             =   5190
      Width           =   5775
      Begin VB.TextBox Txt_valor_total_aprox_tributos 
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
         Left            =   1605
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   54
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor total dos impostos retidos."
         Top             =   330
         Width           =   1245
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Optante pelo regime tributário"
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
      ForeColor       =   &H00000000&
      Height          =   675
      Left            =   55
      TabIndex        =   74
      Top             =   0
      Width           =   11625
      Begin VB.OptionButton Opt_simples1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "4 - Simples nacional (excesso de sublimite de receita bruta)"
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
         Left            =   6900
         TabIndex        =   3
         Top             =   300
         Width           =   4545
      End
      Begin VB.OptionButton Opt_simples 
         BackColor       =   &H00E0E0E0&
         Caption         =   "1 - Simples nacional"
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
         TabIndex        =   0
         Top             =   300
         Width           =   1725
      End
      Begin VB.OptionButton Opt_presumido 
         BackColor       =   &H00E0E0E0&
         Caption         =   "2 - Lucro presumido"
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
         Left            =   2580
         TabIndex        =   1
         Top             =   300
         Width           =   1725
      End
      Begin VB.OptionButton Opt_real 
         BackColor       =   &H00E0E0E0&
         Caption         =   "3 - Lucro real"
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
         Left            =   4980
         TabIndex        =   2
         Top             =   300
         Width           =   1245
      End
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Destacados"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5295
      Left            =   55
      TabIndex        =   57
      Top             =   660
      Width           =   5775
      Begin VB.Frame Frame10 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Sobre produtos"
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
         Height          =   3285
         Left            =   120
         TabIndex        =   58
         Top             =   240
         Width           =   2730
         Begin VB.TextBox Txt_aliquota_total_prod 
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
            Left            =   810
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   14
            TabStop         =   0   'False
            Text            =   "0,00"
            ToolTipText     =   "Porcentagem total dos impostos destacados sobre produtos."
            Top             =   2850
            Width           =   495
         End
         Begin VB.TextBox Txt_valor_total_prod 
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
            Left            =   1605
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   15
            TabStop         =   0   'False
            Text            =   "0,00"
            ToolTipText     =   "Valor total dos impostos destacados sobre produtos."
            Top             =   2850
            Width           =   945
         End
         Begin VB.TextBox Txt_valor_IRPJ_prod 
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
            Left            =   1605
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   13
            TabStop         =   0   'False
            Text            =   "0,00"
            ToolTipText     =   "Valor do imposto."
            Top             =   2430
            Width           =   945
         End
         Begin VB.TextBox Txt_aliquota_IRPJ_prod 
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
            Left            =   810
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   12
            TabStop         =   0   'False
            Text            =   "0,00"
            ToolTipText     =   "Porcentagem de imposto."
            Top             =   2430
            Width           =   495
         End
         Begin VB.TextBox Txt_valor_CSLL_prod 
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
            Left            =   1605
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   11
            TabStop         =   0   'False
            Text            =   "0,00"
            ToolTipText     =   "Valor do imposto."
            Top             =   2010
            Width           =   945
         End
         Begin VB.TextBox Txt_aliquota_CSLL_prod 
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
            Left            =   810
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   10
            TabStop         =   0   'False
            Text            =   "0,00"
            ToolTipText     =   "Porcentagem de imposto."
            Top             =   2010
            Width           =   495
         End
         Begin VB.TextBox Txt_valor_Cofins_prod 
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
            Left            =   1605
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   9
            TabStop         =   0   'False
            Text            =   "0,00"
            ToolTipText     =   "Valor do imposto."
            Top             =   1590
            Width           =   945
         End
         Begin VB.TextBox Txt_aliquota_Cofins_prod 
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
            Left            =   810
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   8
            TabStop         =   0   'False
            Text            =   "0,00"
            ToolTipText     =   "Porcentagem de imposto."
            Top             =   1590
            Width           =   495
         End
         Begin VB.TextBox Txt_aliquota_PIS_prod 
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
            Left            =   810
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   6
            TabStop         =   0   'False
            Text            =   "0,00"
            ToolTipText     =   "Porcentagem de imposto."
            Top             =   1170
            Width           =   495
         End
         Begin VB.TextBox Txt_valor_PIS_prod 
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
            Left            =   1605
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   7
            TabStop         =   0   'False
            Text            =   "0,00"
            ToolTipText     =   "Valor do imposto."
            Top             =   1170
            Width           =   945
         End
         Begin VB.TextBox Txt_valor_IPI_prod 
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
            Left            =   810
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   5
            TabStop         =   0   'False
            Text            =   "0,00"
            ToolTipText     =   "Valor do imposto."
            Top             =   750
            Width           =   1740
         End
         Begin VB.TextBox Txt_valor_ICMS_prod 
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
            Left            =   810
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   4
            TabStop         =   0   'False
            Text            =   "0,00"
            ToolTipText     =   "Valor do imposto."
            Top             =   330
            Width           =   1740
         End
         Begin VB.Label Label35 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "%"
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
            Left            =   1350
            TabIndex        =   84
            Top             =   2910
            Width           =   225
         End
         Begin VB.Label Label34 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Total :"
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
            Left            =   210
            TabIndex        =   83
            Top             =   2850
            Width           =   525
         End
         Begin VB.Label Label19 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "%"
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
            Left            =   1380
            TabIndex        =   68
            Top             =   2490
            Width           =   165
         End
         Begin VB.Label Label18 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "%"
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
            Left            =   1380
            TabIndex        =   67
            Top             =   2070
            Width           =   165
         End
         Begin VB.Label Label17 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "%"
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
            Left            =   1380
            TabIndex        =   66
            Top             =   1650
            Width           =   165
         End
         Begin VB.Label Label16 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "%"
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
            Left            =   1380
            TabIndex        =   65
            Top             =   1230
            Width           =   165
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "IRPJ :"
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
            Left            =   300
            TabIndex        =   64
            Top             =   2430
            Width           =   435
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "CSLL :"
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
            TabIndex        =   63
            Top             =   2010
            Width           =   555
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Cofins :"
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
            TabIndex        =   62
            Top             =   1590
            Width           =   555
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "PIS :"
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
            TabIndex        =   61
            Top             =   1170
            Width           =   555
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "IPI :"
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
            TabIndex        =   60
            Top             =   750
            Width           =   555
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "ICMS :"
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
            Left            =   180
            TabIndex        =   59
            Top             =   330
            Width           =   555
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Sobre produtos e serviços"
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
         Height          =   1635
         Left            =   120
         TabIndex        =   101
         Top             =   3540
         Width           =   2730
         Begin VB.TextBox Txt_valor_total_prod_serv 
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
            Left            =   1605
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   34
            TabStop         =   0   'False
            Text            =   "0,00"
            ToolTipText     =   "Valor total dos impostos destacados sobre produtos."
            Top             =   1170
            Width           =   945
         End
         Begin VB.TextBox Txt_aliquota_total_prod_serv 
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
            Left            =   810
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   33
            TabStop         =   0   'False
            Text            =   "0,00"
            ToolTipText     =   "Porcentagem total dos impostos destacados sobre produtos."
            Top             =   1170
            Width           =   495
         End
         Begin VB.TextBox Txt_aliquota_ICMS_SN 
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
            Left            =   810
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   31
            TabStop         =   0   'False
            Text            =   "0,00"
            ToolTipText     =   "Porcentagem de imposto."
            Top             =   750
            Width           =   495
         End
         Begin VB.TextBox Txt_valor_ICMS_SN 
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
            Left            =   1605
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   32
            TabStop         =   0   'False
            Text            =   "0,00"
            ToolTipText     =   "Valor do imposto."
            Top             =   750
            Width           =   945
         End
         Begin VB.TextBox Txt_valor_DAS 
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
            Left            =   1605
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   30
            TabStop         =   0   'False
            Text            =   "0,00"
            ToolTipText     =   "Valor do imposto."
            Top             =   330
            Width           =   945
         End
         Begin VB.TextBox Txt_aliquota_DAS 
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
            Left            =   810
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   29
            TabStop         =   0   'False
            Text            =   "0,00"
            ToolTipText     =   "Porcentagem de imposto."
            Top             =   330
            Width           =   495
         End
         Begin VB.Label Label49 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Total :"
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
            Left            =   210
            TabIndex        =   116
            Top             =   1170
            Width           =   525
         End
         Begin VB.Label Label43 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "%"
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
            Left            =   1350
            TabIndex        =   115
            Top             =   1230
            Width           =   225
         End
         Begin VB.Label Label22 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "%"
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
            Left            =   1380
            TabIndex        =   114
            Top             =   810
            Width           =   165
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "ICMS :"
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
            Left            =   225
            TabIndex        =   113
            Top             =   750
            Width           =   480
         End
         Begin VB.Label Label58 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "DAS :"
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
            Left            =   300
            TabIndex        =   103
            Top             =   330
            Width           =   405
         End
         Begin VB.Label Label48 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "%"
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
            Left            =   1380
            TabIndex        =   102
            Top             =   390
            Width           =   165
         End
      End
      Begin VB.Frame Frame12 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Sobre serviços"
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
         Height          =   3285
         Left            =   2910
         TabIndex        =   69
         Top             =   240
         Width           =   2730
         Begin VB.TextBox Txt_aliquota_INSS_serv 
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
            Left            =   810
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   23
            TabStop         =   0   'False
            Text            =   "0,00"
            ToolTipText     =   "Porcentagem de imposto."
            Top             =   2010
            Width           =   495
         End
         Begin VB.TextBox Txt_aliquota_total_serv 
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
            Left            =   810
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   27
            TabStop         =   0   'False
            Text            =   "0,00"
            ToolTipText     =   "Porcentagem total dos impostos destacados sobre serviços."
            Top             =   2850
            Width           =   495
         End
         Begin VB.TextBox Txt_valor_total_serv 
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
            Left            =   1605
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   28
            TabStop         =   0   'False
            Text            =   "0,00"
            ToolTipText     =   "Valor total dos impostos destacados sobre serviços."
            Top             =   2850
            Width           =   945
         End
         Begin VB.TextBox Txt_valor_PIS_serv 
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
            Left            =   1605
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   17
            TabStop         =   0   'False
            Text            =   "0,00"
            ToolTipText     =   "Valor do imposto."
            Top             =   330
            Width           =   945
         End
         Begin VB.TextBox Txt_valor_Cofins_serv 
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
            Left            =   1605
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   19
            TabStop         =   0   'False
            Text            =   "0,00"
            ToolTipText     =   "Valor do imposto."
            Top             =   750
            Width           =   945
         End
         Begin VB.TextBox Txt_aliquota_PIS_serv 
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
            Left            =   810
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   16
            TabStop         =   0   'False
            Text            =   "0,00"
            ToolTipText     =   "Porcentagem de imposto."
            Top             =   330
            Width           =   495
         End
         Begin VB.TextBox Txt_aliquota_Cofins_serv 
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
            Left            =   810
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   18
            TabStop         =   0   'False
            Text            =   "0,00"
            ToolTipText     =   "Porcentagem de imposto."
            Top             =   750
            Width           =   495
         End
         Begin VB.TextBox Txt_valor_CSLL_serv 
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
            Left            =   1605
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   21
            TabStop         =   0   'False
            Text            =   "0,00"
            ToolTipText     =   "Valor do imposto."
            Top             =   1170
            Width           =   945
         End
         Begin VB.TextBox Txt_aliquota_CSLL_serv 
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
            Left            =   810
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   20
            TabStop         =   0   'False
            Text            =   "0,00"
            ToolTipText     =   "Porcentagem de imposto."
            Top             =   1170
            Width           =   495
         End
         Begin VB.TextBox Txt_valor_ISSQN_serv 
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
            Left            =   810
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   22
            TabStop         =   0   'False
            Text            =   "0,00"
            ToolTipText     =   "Valor do imposto."
            Top             =   1590
            Width           =   1740
         End
         Begin VB.TextBox Txt_valor_INSS_serv 
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
            Left            =   1605
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   24
            TabStop         =   0   'False
            Text            =   "0,00"
            ToolTipText     =   "Valor do imposto."
            Top             =   2010
            Width           =   945
         End
         Begin VB.TextBox Txt_aliquota_IRRF_serv 
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
            Left            =   810
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   25
            TabStop         =   0   'False
            Text            =   "0,00"
            ToolTipText     =   "Porcentagem de imposto."
            Top             =   2430
            Width           =   495
         End
         Begin VB.TextBox Txt_valor_IRRF_serv 
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
            Left            =   1605
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   26
            TabStop         =   0   'False
            Text            =   "0,00"
            ToolTipText     =   "Valor do imposto."
            Top             =   2430
            Width           =   945
         End
         Begin VB.Label Label21 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "%"
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
            Left            =   1380
            TabIndex        =   110
            Top             =   2070
            Width           =   165
         End
         Begin VB.Label Label33 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "%"
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
            Left            =   1350
            TabIndex        =   82
            Top             =   2910
            Width           =   225
         End
         Begin VB.Label Label32 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Total :"
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
            Left            =   210
            TabIndex        =   81
            Top             =   2850
            Width           =   525
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "PIS :"
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
            Left            =   390
            TabIndex        =   80
            Top             =   330
            Width           =   345
         End
         Begin VB.Label Label30 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Cofins :"
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
            TabIndex        =   79
            Top             =   750
            Width           =   555
         End
         Begin VB.Label Label29 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "CSLL :"
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
            Left            =   285
            TabIndex        =   78
            Top             =   1170
            Width           =   450
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "ISSQN :"
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
            Left            =   165
            TabIndex        =   77
            Top             =   1590
            Width           =   570
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "INSS :"
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
            Left            =   285
            TabIndex        =   76
            Top             =   2010
            Width           =   450
         End
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "IRRF :"
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
            Left            =   270
            TabIndex        =   75
            Top             =   2430
            Width           =   465
         End
         Begin VB.Label Label25 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "%"
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
            Left            =   1380
            TabIndex        =   73
            Top             =   390
            Width           =   165
         End
         Begin VB.Label Label24 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "%"
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
            Left            =   1380
            TabIndex        =   72
            Top             =   810
            Width           =   165
         End
         Begin VB.Label Label23 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "%"
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
            Left            =   1380
            TabIndex        =   71
            Top             =   1230
            Width           =   165
         End
         Begin VB.Label Label20 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "%"
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
            Left            =   1380
            TabIndex        =   70
            Top             =   2490
            Width           =   165
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Total geral destacado"
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
         Height          =   765
         Left            =   2910
         TabIndex        =   104
         Top             =   3990
         Width           =   2730
         Begin VB.TextBox Txt_aliquota_total_geral 
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
            Left            =   810
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   35
            TabStop         =   0   'False
            Text            =   "0,00"
            ToolTipText     =   "Porcentagem total dos impostos destacados."
            Top             =   330
            Width           =   495
         End
         Begin VB.TextBox Txt_valor_total_geral 
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
            Left            =   1605
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   36
            TabStop         =   0   'False
            Text            =   "0,00"
            ToolTipText     =   "Valor total dos impostos destacados."
            Top             =   330
            Width           =   945
         End
         Begin VB.Label Label42 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "%"
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
            Left            =   1350
            TabIndex        =   106
            Top             =   390
            Width           =   225
         End
         Begin VB.Label Label41 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Total :"
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
            Left            =   180
            TabIndex        =   105
            Top             =   330
            Width           =   525
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Retidos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4515
      Left            =   5910
      TabIndex        =   85
      Top             =   660
      Width           =   5775
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Sobre produtos"
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
         Height          =   3285
         Left            =   120
         TabIndex        =   97
         Top             =   240
         Width           =   2100
         Begin VB.TextBox Txt_valor_PIS_prod1 
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
            Left            =   810
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   37
            TabStop         =   0   'False
            Text            =   "0,00"
            ToolTipText     =   "Valor do imposto."
            Top             =   330
            Width           =   1095
         End
         Begin VB.TextBox Txt_valor_Cofins_prod1 
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
            Left            =   810
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   38
            TabStop         =   0   'False
            Text            =   "0,00"
            ToolTipText     =   "Valor do imposto."
            Top             =   1590
            Width           =   1095
         End
         Begin VB.TextBox Txt_valor_total_prod1 
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
            Left            =   810
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   39
            TabStop         =   0   'False
            Text            =   "0,00"
            ToolTipText     =   "Valor total dos impostos retidos sobre produtos."
            Top             =   2850
            Width           =   1095
         End
         Begin VB.Label Label56 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "PIS :"
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
            Left            =   390
            TabIndex        =   100
            Top             =   330
            Width           =   345
         End
         Begin VB.Label Label55 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Cofins :"
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
            TabIndex        =   99
            Top             =   1590
            Width           =   555
         End
         Begin VB.Label Label44 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Total :"
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
            Left            =   210
            TabIndex        =   98
            Top             =   2850
            Width           =   525
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Sobre serviços"
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
         Height          =   3285
         Left            =   2280
         TabIndex        =   86
         Top             =   240
         Width           =   3390
         Begin VB.TextBox Txt_valor_total_serv1 
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
            Left            =   1605
            MaxLength       =   50
            TabIndex        =   51
            Text            =   "0,00"
            ToolTipText     =   "Valor total dos impostos retidos sobre serviços."
            Top             =   2850
            Width           =   1575
         End
         Begin VB.TextBox Txt_aliquota_IRRF_serv1 
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
            Left            =   810
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   48
            TabStop         =   0   'False
            Text            =   "0,00"
            ToolTipText     =   "Porcentagem de imposto."
            Top             =   2346
            Width           =   495
         End
         Begin VB.TextBox Txt_valor_IRRF_serv1 
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
            Left            =   1605
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   49
            TabStop         =   0   'False
            Text            =   "0,00"
            ToolTipText     =   "Valor do imposto."
            Top             =   2346
            Width           =   1575
         End
         Begin VB.TextBox Txt_valor_INSS_serv1 
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
            Left            =   1605
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   47
            TabStop         =   0   'False
            Text            =   "0,00"
            ToolTipText     =   "Valor do imposto."
            Top             =   1842
            Width           =   1575
         End
         Begin VB.TextBox Txt_aliquota_INSS_serv1 
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
            Left            =   810
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   46
            TabStop         =   0   'False
            Text            =   "0,00"
            ToolTipText     =   "Porcentagem de imposto."
            Top             =   1842
            Width           =   495
         End
         Begin VB.TextBox Txt_aliquota_CSLL_serv1 
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
            Left            =   810
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   44
            TabStop         =   0   'False
            Text            =   "0,00"
            ToolTipText     =   "Porcentagem de imposto."
            Top             =   1338
            Width           =   495
         End
         Begin VB.TextBox Txt_valor_CSLL_serv1 
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
            Left            =   1605
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   45
            TabStop         =   0   'False
            Text            =   "0,00"
            ToolTipText     =   "Valor do imposto."
            Top             =   1338
            Width           =   1575
         End
         Begin VB.TextBox Txt_aliquota_Cofins_serv1 
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
            Left            =   810
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   42
            TabStop         =   0   'False
            Text            =   "0,00"
            ToolTipText     =   "Porcentagem de imposto."
            Top             =   834
            Width           =   495
         End
         Begin VB.TextBox Txt_aliquota_PIS_serv1 
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
            Left            =   810
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   40
            TabStop         =   0   'False
            Text            =   "0,00"
            ToolTipText     =   "Porcentagem de imposto."
            Top             =   330
            Width           =   495
         End
         Begin VB.TextBox Txt_valor_Cofins_serv1 
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
            Left            =   1605
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   43
            TabStop         =   0   'False
            Text            =   "0,00"
            ToolTipText     =   "Valor do imposto."
            Top             =   834
            Width           =   1575
         End
         Begin VB.TextBox Txt_valor_PIS_serv1 
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
            Left            =   1605
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   41
            TabStop         =   0   'False
            Text            =   "0,00"
            ToolTipText     =   "Valor do imposto."
            Top             =   330
            Width           =   1575
         End
         Begin VB.TextBox Txt_aliquota_total_serv1 
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
            Left            =   810
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   50
            TabStop         =   0   'False
            Text            =   "0,00"
            ToolTipText     =   "Porcentagem total dos impostos retidos sobre serviços."
            Top             =   2850
            Width           =   495
         End
         Begin VB.Label Label47 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "IRRF :"
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
            Left            =   285
            TabIndex        =   112
            Top             =   2346
            Width           =   465
         End
         Begin VB.Label Label36 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "%"
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
            Left            =   1380
            TabIndex        =   111
            Top             =   2406
            Width           =   165
         End
         Begin VB.Label Label40 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "%"
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
            Left            =   1380
            TabIndex        =   96
            Top             =   1902
            Width           =   165
         End
         Begin VB.Label Label39 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "%"
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
            Left            =   1380
            TabIndex        =   95
            Top             =   1398
            Width           =   165
         End
         Begin VB.Label Label38 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "%"
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
            Left            =   1380
            TabIndex        =   94
            Top             =   894
            Width           =   165
         End
         Begin VB.Label Label37 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "%"
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
            Left            =   1380
            TabIndex        =   93
            Top             =   390
            Width           =   165
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "INSS :"
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
            Left            =   285
            TabIndex        =   92
            Top             =   1842
            Width           =   450
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "CSLL :"
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
            Left            =   285
            TabIndex        =   91
            Top             =   1338
            Width           =   450
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Cofins :"
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
            TabIndex        =   90
            Top             =   834
            Width           =   555
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "PIS :"
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
            Left            =   390
            TabIndex        =   89
            Top             =   330
            Width           =   345
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Total :"
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
            Left            =   210
            TabIndex        =   88
            Top             =   2850
            Width           =   525
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "%"
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
            Left            =   1350
            TabIndex        =   87
            Top             =   2910
            Width           =   225
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Total geral retido"
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
         Height          =   765
         Left            =   1350
         TabIndex        =   107
         Top             =   3600
         Width           =   3030
         Begin VB.TextBox Txt_valor_total_geral1 
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
            Left            =   1605
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   53
            TabStop         =   0   'False
            Text            =   "0,00"
            ToolTipText     =   "Valor total dos impostos retidos."
            Top             =   330
            Width           =   1245
         End
         Begin VB.TextBox Txt_aliquota_total_geral1 
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
            Left            =   810
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   52
            TabStop         =   0   'False
            Text            =   "0,00"
            ToolTipText     =   "Porcentagem total dos impostos retidos."
            Top             =   330
            Width           =   495
         End
         Begin VB.Label Label46 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Total :"
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
            Left            =   180
            TabIndex        =   109
            Top             =   330
            Width           =   525
         End
         Begin VB.Label Label45 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "%"
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
            Left            =   1350
            TabIndex        =   108
            Top             =   390
            Width           =   225
         End
      End
   End
   Begin VB.Frame Frame13 
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
      ForeColor       =   &H80000007&
      Height          =   855
      Left            =   55
      TabIndex        =   118
      Top             =   5940
      Width           =   10635
      Begin VB.ComboBox Cmb_tipo_TBSN 
         Appearance      =   0  'Flat
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
         ItemData        =   "FrmImpostos.frx":000C
         Left            =   180
         List            =   "FrmImpostos.frx":000E
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   55
         ToolTipText     =   "Tabela do simples nacional."
         Top             =   390
         Width           =   10275
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Tabela do simples nacional"
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
         Left            =   4372
         TabIndex        =   119
         Top             =   180
         Width           =   1890
      End
   End
End
Attribute VB_Name = "FrmImpostos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cmd_salvar_tabelaSN_Click()
On Error GoTo tratar_erro

Acao = "salvar"
If Cmb_tipo_TBSN = "" Then
    NomeCampo = "a tabela"
    ProcVerificaAcao
    Cmb_tipo_TBSN.SetFocus
    Exit Sub
End If
If USMsgBox("Deseja realmente alterar a tabela do simples nacional.", vbYesNo, "CAPRIND v5.0") = vbYes Then
    Select Case Mid(Cmb_tipo_TBSN, 8, 3)
        Case "I -": TabelaSN = 1
        Case "II ": TabelaSN = 2
        Case "III": TabelaSN = 3
        Case "IV ": TabelaSN = 4
        Case "V ": TabelaSN = 4
    End Select
       If Vendas_Proposta = True Or Vendas_PI = True Then
        If Vendas_Proposta = True Then
            With frmVendas_proposta
                If FunVerificaRegistroValidado("Vendas_proposta", "Cotacao = " & .txtId, "proposta", "a tabela do simples nacional", "alterar", False, True) = False Then Exit Sub
                .TabelaSN_Proposta = TabelaSN
                ID_documento = .txtId
                Documento = "Nº proposta: " & .txtCotacao & " - Rev.: " & .txtrevisao
            End With
        Else
            With frmVendas_PI
                If FunVerificaRegistroValidado("Vendas_proposta", "Cotacao = " & .txtId, "pedido", "a tabela do simples nacional", "alterar", True, True) = False Then Exit Sub
                .TabelaSN_PI = TabelaSN
                ID_documento = .txtId
                Documento = "Nº pedido: " & .txtCotacao & " - Rev.: " & .txtrevisao
            End With
        End If
        Conexao.Execute "UPDATE vendas_proposta Set TabelaSN = " & TabelaSN & " where Cotacao = " & ID_documento
        
        'Corrige os valores
        Set TBCotacao = CreateObject("adodb.recordset")
        TBCotacao.Open "Select VP.ID_empresa, VP.IDCliente, VP.Cliente, VP.TabelaSN, VC.* from Vendas_proposta VP INNER JOIN vendas_carteira VC ON VC.Cotacao = VP.Cotacao where VP.cotacao = " & ID_documento, Conexao, adOpenKeyset, adLockOptimistic
        If TBCotacao.EOF = False Then
            Do While TBCotacao.EOF = False
                Valor_total = Format(TBCotacao!preco_unitario_desconto * TBCotacao!quantidade, "###,##0.00")
                ProcControleImposto IIf(IsNull(TBCotacao!ID_CFOP), 0, TBCotacao!ID_CFOP), IIf(IsNull(TBCotacao!IDCliente), 0, TBCotacao!IDCliente)
                If TBCotacao!Tipo = "P" Then
                    ProcVerifImpostosEmpresa TBCotacao!ID_empresa, TBCotacao!retorno, "", False, 0, False, TBCotacao!TabelaSN, 0
                    'Novo cálculo simples nacional 2018
                    TBCotacao!DAS = DAS
                    If DAS <> 0 Then TBCotacao!Total_DAS = Format((Valor_total * DAS) / 100, "###,##0.00") Else TBCotacao!Total_DAS = 0
                    TBCotacao!PIS_Prod = PIS_Prod
                    If PIS_Prod <> 0 Then TBCotacao!Total_PIS_prod = Format((Valor_total * PIS_Prod) / 100, "###,##0.00") Else TBCotacao!Total_PIS_prod = 0
                    TBCotacao!Cofins_Prod = Cofins_Prod
                    If Cofins_Prod <> 0 Then TBCotacao!Total_Cofins_prod = Format((Valor_total * Cofins_Prod) / 100, "###,##0.00") Else TBCotacao!Total_Cofins_prod = 0
                    TBCotacao!CSLL_Prod = CSLL_Prod
                    If CSLL_Prod <> 0 Then TBCotacao!Total_CSLL_prod = Format((Valor_total * CSLL_Prod) / 100, "###,##0.00") Else TBCotacao!Total_CSLL_prod = 0
                    TBCotacao!IRPJ_Prod = IRPJ_Prod
                    If IRPJ_Prod <> 0 Then TBCotacao!Total_IRPJ_prod = Format((Valor_total * IRPJ_Prod) / 100, "###,##0.00") Else TBCotacao!Total_IRPJ_prod = 0
                    TBCotacao!cpp = CPP_Prod
                    If CPP_Prod <> 0 Then TBCotacao!Total_CPP = Format((Valor_total * CPP_Prod) / 100, "###,##0.00") Else TBCotacao!Total_CPP = 0
                    Valor_total = valor
                Else
                    'Novo cálculo simples nacional 2018
                    ProcVerifImpostosEmpresa TBCotacao!ID_empresa, False, TBCotacao!Desenho, IIf(IsNull(TBCotacao!Servico_cliente), False, TBCotacao!Servico_cliente), Valor_total, True, TBCotacao!TabelaSN, 0
                    TBCotacao!DAS = DAS
                    If DAS <> 0 Then TBCotacao!Total_DAS = Format((Valor_total * DAS) / 100, "###,##0.00") Else TBCotacao!Total_DAS = 0
                    TBCotacao!PIS_Serv = PIS_Serv
                    If PIS_Serv <> 0 Then TBCotacao!Total_PIS_serv = Format((Valor_total * PIS_Serv) / 100, "###,##0.00") Else TBCotacao!Total_PIS_serv = 0
                    TBCotacao!Cofins_Serv = Cofins_Serv
                    If Cofins_Serv <> 0 Then TBCotacao!Total_Cofins_serv = Format((Valor_total * Cofins_Serv) / 100, "###,##0.00") Else TBCotacao!Total_Cofins_serv = 0
                    TBCotacao!CSLL_Serv = CSLL_Serv
                    If CSLL_Serv <> 0 Then TBCotacao!Total_CSLL_serv = Format((Valor_total * CSLL_Serv) / 100, "###,##0.00") Else TBCotacao!Total_CSLL_serv = 0
                    TBCotacao!ISS = ISS_Serv
                    If ISS_Serv <> 0 Then TBCotacao!VlrISS = Format((Valor_total * ISS_Serv) / 100, "###,##0.00") Else TBCotacao!VlrISS = 0
                    TBCotacao!INSS_Serv = INSS_Serv
                    If INSS_Serv <> 0 Then TBCotacao!Total_INSS_serv = Format((Valor_total * INSS_Serv) / 100, "###,##0.00") Else TBCotacao!Total_INSS_serv = 0
                    TBCotacao!IRPJ_Serv = IRPJ_Serv
                    If IRPJ_Serv <> 0 Then TBCotacao!Total_IRPJ_serv = Format((Valor_total * IRPJ_Serv) / 100, "###,##0.00") Else TBCotacao!Total_IRPJ_serv = 0
                    TBCotacao!IRRF_Serv = IRRF_Serv
                    If IRRF_Serv <> 0 Then TBCotacao!Total_IRRF_serv = Format((Valor_total * IRRF_Serv) / 100, "###,##0.00") Else TBCotacao!Total_IRRF_serv = 0
                    TBCotacao!cpp = CPP_Serv
                    If CPP_Serv <> 0 Then TBCotacao!Total_CPP = Format((Valor_total * CPP_Serv) / 100, "###,##0.00") Else TBCotacao!Total_CPP = 0
                End If
                TBCotacao.Update
                TBCotacao.MoveNext
            Loop
        End If
        TBCotacao.Close
        If Vendas_Proposta = True Then frmVendas_proposta.ProcGravarTotais (ID_documento) Else frmVendas_PI.ProcGravarTotais (ID_documento)
        ProcCarregaDados
    Else
        With frmFaturamento_Prod_Serv
            If FunVerificaRegistroValidado("tbl_Dados_Nota_Fiscal", "ID = " & .txtId, IIf(.txtNFiscal = "", "ordem de faturamento", "nota fiscal"), "a tabela do simples nacional", "alterar", False, True) = False Then Exit Sub
            '.TabelaSN = TabelaSN
            Conexao.Execute "UPDATE tbl_Dados_Nota_Fiscal Set TabelaSN = " & TabelaSN & " where ID = " & .txtId
            .ProcCorrigeValorImpostosSN .txtId
            .ProcVerificaTipoNF False
            If .txtNFiscal = "" Then NomeCampo = "N° ordem: " & .txtId Else NomeCampo = "N° nota: " & .txtNFiscal
            Documento = NomeCampo & " - Tipo: " & TipoNF & " - Série: " & .txtSerie
        End With
        ProcCarregaDados
    End If
    USMsgBox ("Tabela do simples nacional alterada com sucesso."), vbInformation, "CAPRIND v5.0"
    '==================================
    Modulo = Formulario
    Evento = "Alterar tabela do simples nacional"
    Documento1 = ""
    ProcGravaEvento
    '==================================
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyF3: If Cmd_salvar_tabelaSN.Enabled = True Then Cmd_salvar_tabelaSN_Click
    Case vbKeyEscape: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaDados

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcCarregaDados()
On Error GoTo tratar_erro

DAS = 0
If Vendas_Proposta = True Or Vendas_PI = True Then
    Caption = "Administrativo - Vendas - " & IIf(Vendas_Proposta = True, "Proposta comercial", "Pedido interno") & " - Impostos"
    With IIf(Vendas_PI = True, frmVendas_PI, frmVendas_proposta)
        ID = .txtId
        ID_empresa = .Cmb_empresa.ItemData(.Cmb_empresa.ListIndex)
    End With
    NomeTabela = "Vendas_proposta"
    CampoFiltro = "Cotacao"
    Frame8.Visible = False
Else
    If Formulario = "Faturamento/Nota fiscal/Própria" Then
        Caption = "Administrativo - Faturamento - Nota fiscal - Própria - Impostos"
    ElseIf Formulario = "Faturamento/Nota fiscal/Terceiros" Then
            Caption = "Administrativo - Faturamento - Nota fiscal - Terceiros - Impostos"
        ElseIf Formulario = "Estoque/Ordem de faturamento" Then
                Caption = "Estoque - Ordem de faturamento - Impostos"
            Else
                Caption = "Estoque - Nota fiscal - Impostos"
    End If
    With frmFaturamento_Prod_Serv
        ID = .txtId
        ID_empresa = .txtIDEmpresa
    End With
    NomeTabela = "tbl_Dados_Nota_Fiscal"
    CampoFiltro = "ID"
End If

'Verifica regime vinculado a proposta/pedido ou NF
Set TBFI = CreateObject("adodb.recordset")
TBFI.Open "Select Regime from " & NomeTabela & " where " & CampoFiltro & " = " & ID & " and Regime IS NOT NULL", Conexao, adOpenKeyset, adLockOptimistic
If TBFI.EOF = False Then
    If TBFI!Regime = 1 Then
        Opt_simples.Value = True
        Regime = 1
    End If
    If TBFI!Regime = 2 Then
        Opt_presumido.Value = True
        Regime = 2
    End If
    If TBFI!Regime = 3 Then
        Opt_real.Value = True
        Regime = 3
    End If
    If TBFI!Regime = 4 Then
        Opt_simples1.Value = True
        Regime = 4
    End If
    
    If Regime <> 1 Then
        Frame13.Visible = False
        Cmd_salvar_tabelaSN.Visible = False
        Height = 6435
    Else
        With Cmb_tipo_TBSN
            .Clear
            'Carrega tabelas do simples cadastradas
            Contador = 0
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select Tabela FROM Impostos_TabelaDAS where ID_empresa = " & ID_empresa & " and Ativado = 1 group by Tabela", Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False Then
                Do While TBFI.EOF = False
                    Select Case TBFI!Tabela
                        Case 1: .AddItem "Tabela I - Partilha do Simples Nacional – Comércio"
                        Case 2: .AddItem "Tabela II - Partilha do Simples Nacional - Indústria"
                        Case 3: .AddItem "Tabela III - Partilha do Simples Nacional - Serviços e Locação de Bens Móveis"
                        Case 4: .AddItem "Tabela IV - Partilha do Simples Nacional - Serviços"
                        Case 5: .AddItem "Tabela V - Partilha do Simples Nacional - Partilha do Simples Nacional - Receitas decorrentes da prestação de serviços relacionados no § 5º-I do art. 18 da LC 123/2016"
                    End Select
                    
                    TabelaSN = TBFI!Tabela
                    Contador = Contador + 1
                    TBFI.MoveNext
                Loop
                If Contador = 1 Then
                    Select Case TabelaSN
                        Case 1: .Text = "Tabela I - Partilha do Simples Nacional – Comércio"
                        Case 2: .Text = "Tabela II - Partilha do Simples Nacional - Indústria"
                        Case 3: .Text = "Tabela III - Partilha do Simples Nacional - Serviços e Locação de Bens Móveis"
                        Case 4: .Text = "Tabela IV - Partilha do Simples Nacional - Serviços"
                        Case 5: .Text = "Tabela V - Partilha do Simples Nacional - Partilha do Simples Nacional - Receitas decorrentes da prestação de serviços relacionados no § 5º-I do art. 18 da LC 123/2016"
                    End Select
                    .Locked = True
                    .TabStop = False
                Else
                    .Locked = False
                    .TabStop = True
                End If
            Else
                USMsgBox ("Não existe nenhuma tabela do simples nacional ativa, favor verificar."), vbExclamation, "CAPRIND v5.0"
                TBFI.Close
                Exit Sub
            End If
            TBFI.Close
        End With
    End If
    
    Set TBAbrir = CreateObject("adodb.recordset")
    If Vendas_Proposta = True Or Vendas_PI = True Then
        TBAbrir.Open "Select * from vendas_proposta where Cotacao = " & ID, Conexao, adOpenKeyset, adLockOptimistic
    Else
        TBAbrir.Open "Select TN.*, NF.TabelaSN from tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_Totais_Nota TN ON TN.ID_nota = NF.ID where NF.ID = " & ID, Conexao, adOpenKeyset, adLockOptimistic
    End If
    
    If TBAbrir.EOF = False Then
        
        If IsNull(TBAbrir!TabelaSN) = False And Regime = 1 Then
            With Cmb_tipo_TBSN
                Select Case TBAbrir!TabelaSN
                    Case 1: .Text = "Tabela I - Partilha do Simples Nacional – Comércio"
                    Case 2: .Text = "Tabela II - Partilha do Simples Nacional - Indústria"
                    Case 3: .Text = "Tabela III - Partilha do Simples Nacional - Serviços e Locação de Bens Móveis"
                    Case 4: .Text = "Tabela IV - Partilha do Simples Nacional - Serviços"
                    Case 5: .Text = "Tabela V - Partilha do Simples Nacional - Partilha do Simples Nacional - Receitas decorrentes da prestação de serviços relacionados no § 5º-I do art. 18 da LC 123/2016"
                End Select
            End With
        End If
        
        If Vendas_Proposta = True Or Vendas_PI = True Then
                   
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select * from vendas_carteira where cotacao = " & ID & " and Tipo = 'P'", Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False Then
                
                'Destaque
                Txt_valor_ICMS_prod = IIf(IsNull(TBAbrir!dbl_Valor_ICMS), "0,00", Format(TBAbrir!dbl_Valor_ICMS, "###,##0.00"))
                ValorICMS = Txt_valor_ICMS_prod
                
                Txt_valor_IPI_prod = IIf(IsNull(TBAbrir!dbl_Valor_Total_IPI), "0,00", Format(TBAbrir!dbl_Valor_Total_IPI, "###,##0.00"))
                Valor_IPI = Txt_valor_IPI_prod
                
                Txt_aliquota_PIS_prod = IIf(IsNull(TBFI!PIS_Prod), "0,00", Format(TBFI!PIS_Prod, "###,##0.00"))
                PIS_Prod = Txt_aliquota_PIS_prod
                Txt_valor_PIS_prod = IIf(IsNull(TBAbrir!Total_PIS_prod), "0,00", Format(TBAbrir!Total_PIS_prod, "###,##0.00"))
                Valor_PIS_Prod = Txt_valor_PIS_prod
                
                Txt_aliquota_Cofins_prod = IIf(IsNull(TBFI!Cofins_Prod), "0,00", Format(TBFI!Cofins_Prod, "###,##0.00"))
                Cofins_Prod = Txt_aliquota_Cofins_prod
                Txt_valor_Cofins_prod = IIf(IsNull(TBAbrir!Total_Cofins_prod), "0,00", Format(TBAbrir!Total_Cofins_prod, "###,##0.00"))
                Valor_Cofins_Prod = Txt_valor_Cofins_prod
                
                Txt_aliquota_CSLL_prod = IIf(IsNull(TBFI!CSLL_Prod), "0,00", Format(TBFI!CSLL_Prod, "###,##0.00"))
                CSLL_Prod = Txt_aliquota_CSLL_prod
                Txt_valor_CSLL_prod = IIf(IsNull(TBAbrir!Total_CSLL_prod), "0,00", Format(TBAbrir!Total_CSLL_prod, "###,##0.00"))
                Valor_CSLL_Prod = Txt_valor_CSLL_prod
                
                Txt_aliquota_IRPJ_prod = IIf(IsNull(TBFI!IRPJ_Prod), "0,00", Format(TBFI!IRPJ_Prod, "###,##0.00"))
                IRPJ_Prod = Txt_aliquota_IRPJ_prod
                Txt_valor_IRPJ_prod = IIf(IsNull(TBAbrir!Total_IRPJ_prod), "0,00", Format(TBAbrir!Total_IRPJ_prod, "###,##0.00"))
                Valor_IRPJ_Prod = Txt_valor_IRPJ_prod
                
                Txt_aliquota_total_prod = Format(PIS_Prod + Cofins_Prod + CSLL_Prod + IRPJ_Prod, "###,##0.00")
                Txt_valor_total_prod = Format(ValorICMS + Valor_IPI + Valor_PIS_Prod + Valor_Cofins_Prod + Valor_CSLL_Prod + Valor_IRPJ_Prod, "###,##0.00")
                
                If Permitido = True Then
                    'Retenção
                    Txt_valor_PIS_prod1 = IIf(IsNull(TBAbrir!Total_retencao_PIS), "0,00", Format(TBAbrir!Total_retencao_PIS, "###,##0.00"))
                    Valor_PIS_Prod = Txt_valor_PIS_prod1
                    
                    Txt_valor_Cofins_prod1 = IIf(IsNull(TBAbrir!Total_retencao_Cofins), "0,00", Format(TBAbrir!Total_retencao_Cofins, "###,##0.00"))
                    Valor_Cofins_Prod = Txt_valor_Cofins_prod1
                                
                    Txt_valor_total_prod1 = Format(Valor_PIS_Prod + Valor_Cofins_Prod, "###,##0.00")
                End If
                
                DAS = IIf(IsNull(TBFI!DAS), 0, TBFI!DAS)
            End If
            'Serviços
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select * from vendas_carteira where cotacao = " & ID & " and Tipo = 'S'", Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False Then
                
                'Destaque
                ProcControleImposto IIf(IsNull(TBFI!ID_CFOP), 0, TBFI!ID_CFOP), TBAbrir!IDCliente
                If DestacaImpostos = "SIM" Then
                    Txt_aliquota_PIS_serv = IIf(IsNull(TBFI!PIS_Serv), "0,00", Format(TBFI!PIS_Serv, "###,##0.00"))
                    PIS_Serv = Txt_aliquota_PIS_serv
                    Txt_valor_PIS_serv = IIf(IsNull(TBAbrir!Total_PIS_serv), "0,00", Format(TBAbrir!Total_PIS_serv, "###,##0.00"))
                    Valor_PIS_Serv = Txt_valor_PIS_serv
                    
                    Txt_aliquota_Cofins_serv = IIf(IsNull(TBFI!Cofins_Serv), "0,00", Format(TBFI!Cofins_Serv, "###,##0.00"))
                    Cofins_Serv = Txt_aliquota_Cofins_serv
                    Txt_valor_Cofins_serv = IIf(IsNull(TBAbrir!Total_Cofins_serv), "0,00", Format(TBAbrir!Total_Cofins_serv, "###,##0.00"))
                    Valor_Cofins_Serv = Txt_valor_Cofins_serv
                    
                    Txt_aliquota_CSLL_serv = IIf(IsNull(TBFI!CSLL_Serv), "0,00", Format(TBFI!CSLL_Serv, "###,##0.00"))
                    CSLL_Serv = Txt_aliquota_CSLL_serv
                    Txt_valor_CSLL_serv = IIf(IsNull(TBAbrir!Total_CSLL_serv), "0,00", Format(TBAbrir!Total_CSLL_serv, "###,##0.00"))
                    Valor_CSLL_Serv = Txt_valor_CSLL_serv
                    
                    Txt_valor_ISSQN_serv = IIf(IsNull(TBAbrir!VlrTotaliss), "0,00", Format(TBAbrir!VlrTotaliss, "###,##0.00"))
                    Valor_ISS_Serv = Txt_valor_ISSQN_serv
                                    
                    'INSS
                    Set TBFIltro = CreateObject("adodb.recordset")
                    TBFIltro.Open "Select INSS_Serv, Total_INSS_serv from vendas_carteira where cotacao = " & ID & " and Tipo = 'S' and INSS_serv is not null and INSS_serv <> 0", Conexao, adOpenKeyset, adLockOptimistic
                    If TBFIltro.EOF = False Then
                        Txt_aliquota_INSS_serv = IIf(IsNull(TBFIltro!INSS_Serv), "0,00", Format(TBFIltro!INSS_Serv, "###,##0.00"))
                        INSS_Serv = Txt_aliquota_INSS_serv
                        Txt_valor_INSS_serv = IIf(IsNull(TBAbrir!Total_INSS_serv), "0,00", Format(TBAbrir!Total_INSS_serv, "###,##0.00"))
                        Valor_INSS_Serv = Txt_valor_INSS_serv
                    End If
                    TBFIltro.Close
                                    
                    Txt_aliquota_IRRF_serv = IIf(IsNull(TBFI!IRRF_Serv), "0,00", Format(TBFI!IRRF_Serv, "###,##0.00"))
                    IRRF_Serv = Txt_aliquota_IRRF_serv
                    Txt_valor_IRRF_serv = IIf(IsNull(TBAbrir!Total_IRRF_serv), "0,00", Format(TBAbrir!Total_IRRF_serv, "###,##0.00"))
                    Valor_IRRF_Serv = Txt_valor_IRRF_serv
                    
                    If TBAbrir!dbl_valor_total_servicos > 0 Then ISS_Serv = Format((Valor_ISS_Serv / IIf(IsNull(TBAbrir!dbl_valor_total_servicos), 0, TBAbrir!dbl_valor_total_servicos)) * 100, "0.00") Else ISS_Serv = 0
                    Txt_aliquota_total_serv = Format(PIS_Serv + Cofins_Serv + CSLL_Serv + ISS_Serv + INSS_Serv + IRRF_Serv, "###,##0.00")
                    Txt_valor_total_serv = Format(Valor_PIS_Serv + Valor_Cofins_Serv + Valor_CSLL_Serv + Valor_ISS_Serv + Valor_INSS_Serv + Valor_IRRF_Serv, "###,##0.00")
                End If
                If Permitido = True Then
                    
                    'Retenção
                    Set TBTotaisnota = CreateObject("adodb.recordset")
                    TBTotaisnota.Open "Select * from Impostos where Regime = " & Regime, Conexao, adOpenKeyset, adLockOptimistic
                    If TBTotaisnota.EOF = False Then
                        valor = IIf(IsNull(TBAbrir!dbl_valor_total_servicos), 0, TBAbrir!dbl_valor_total_servicos)
                        If valor > TBTotaisnota!Acima Then
                            Txt_aliquota_PIS_serv1 = IIf(IsNull(TBFI!PIS_Serv), "0,00", Format(TBFI!PIS_Serv, "###,##0.00"))
                            PIS_Serv = Txt_aliquota_PIS_serv1
                            Txt_valor_PIS_serv1 = IIf(IsNull(TBAbrir!Total_PIS_serv), "0,00", Format(TBAbrir!Total_PIS_serv, "###,##0.00"))
                            Valor_PIS_Serv = Txt_valor_PIS_serv1
                            
                            Txt_aliquota_Cofins_serv1 = IIf(IsNull(TBFI!Cofins_Serv), "0,00", Format(TBFI!Cofins_Serv, "###,##0.00"))
                            Cofins_Serv = Txt_aliquota_Cofins_serv1
                            Txt_valor_Cofins_serv1 = IIf(IsNull(TBAbrir!Total_Cofins_serv), "0,00", Format(TBAbrir!Total_Cofins_serv, "###,##0.00"))
                            Valor_Cofins_Serv = Txt_valor_Cofins_serv1
                            
                            Txt_aliquota_CSLL_serv1 = IIf(IsNull(TBFI!CSLL_Serv), "0,00", Format(TBFI!CSLL_Serv, "###,##0.00"))
                            CSLL_Serv = Txt_aliquota_CSLL_serv1
                            Txt_valor_CSLL_serv1 = IIf(IsNull(TBAbrir!Total_CSLL_serv), "0,00", Format(TBAbrir!Total_CSLL_serv, "###,##0.00"))
                            Valor_CSLL_Serv = Txt_valor_CSLL_serv1
                            
                            'INSS
                            Set TBFIltro = CreateObject("adodb.recordset")
                            TBFIltro.Open "Select * from vendas_carteira where cotacao = " & ID & " and Tipo = 'S' and INSS_serv is not null and INSS_serv <> 0", Conexao, adOpenKeyset, adLockOptimistic
                            If TBFIltro.EOF = False Then
                                Txt_aliquota_INSS_serv1 = IIf(IsNull(TBFIltro!INSS_Serv), "0,00", Format(TBFIltro!INSS_Serv, "###,##0.00"))
                                INSS_Serv = Txt_aliquota_INSS_serv1
                                Txt_valor_INSS_serv1 = IIf(IsNull(TBAbrir!Total_INSS_serv), "0,00", Format(TBAbrir!Total_INSS_serv, "###,##0.00"))
                                Valor_INSS_Serv = Txt_valor_INSS_serv1
                            End If
                            TBFIltro.Close
                            
                            Txt_aliquota_IRRF_serv1 = IIf(IsNull(TBFI!IRRF_Serv), "0,00", Format(TBFI!IRRF_Serv, "###,##0.00"))
                            IRRF_Serv = Txt_aliquota_IRRF_serv1
                            Txt_valor_IRRF_serv1 = IIf(IsNull(TBAbrir!Total_IRRF_serv), "0,00", Format(TBAbrir!Total_IRRF_serv, "###,##0.00"))
                            Valor_IRRF_Serv = Txt_valor_IRRF_serv1
                            
                            Txt_aliquota_total_serv1 = Format(PIS_Serv + Cofins_Serv + CSLL_Serv + INSS_Serv + IRRF_Serv, "###,##0.00")
                            Txt_valor_total_serv1 = Format(Valor_PIS_Serv + Valor_Cofins_Serv + Valor_CSLL_Serv + Valor_INSS_Serv + Valor_IRRF_Serv, "###,##0.00")
                        ElseIf valor >= 667 And valor <= 5000 Then
                                Txt_aliquota_IRRF_serv1 = IIf(IsNull(TBFI!IRRF_Serv), "0,00", Format(TBFI!IRRF_Serv, "###,##0.00"))
                                IRRF_Serv = Txt_aliquota_IRRF_serv1
                                Txt_valor_IRRF_serv1 = IIf(IsNull(TBAbrir!Total_IRRF_serv), "0,00", Format(TBAbrir!Total_IRRF_serv, "###,##0.00"))
                                Valor_IRRF_Serv = Txt_valor_IRRF_serv1
                                
                                Txt_aliquota_total_serv1 = Format(IRRF_Serv, "###,##0.00")
                                Txt_valor_total_serv1 = Format(Valor_IRRF_Serv, "###,##0.00")
                        End If
                    End If
                    TBTotaisnota.Close
                End If
                
                DAS = IIf(IsNull(TBFI!DAS), 0, TBFI!DAS)
            End If
            'Produtos e serviços
            Txt_aliquota_DAS = Format(DAS, "###,##0.00")
            DAS = Txt_aliquota_DAS
            Txt_valor_DAS = IIf(IsNull(TBAbrir!Total_DAS), "0,00", Format(TBAbrir!Total_DAS, "###,##0.00"))
            Valor_DAS = Txt_valor_DAS
            Txt_aliquota_total_prod_serv = Format(ICMS_SN + DAS, "###,##0.00")
            Txt_valor_total_prod_serv = Format(Valor_ICMS_SN + Valor_DAS, "###,##0.00")
        Else
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select * from tbl_Detalhes_Nota where ID_nota = " & ID, Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False Then
                
                'Destaque
                Txt_valor_ICMS_prod = IIf(IsNull(TBAbrir!dbl_Valor_ICMS), "0,00", Format(TBAbrir!dbl_Valor_ICMS, "###,##0.00"))
                ValorICMS = Txt_valor_ICMS_prod
                
                Txt_valor_IPI_prod = IIf(IsNull(TBAbrir!dbl_Valor_Total_IPI), "0,00", Format(TBAbrir!dbl_Valor_Total_IPI, "###,##0.00"))
                Valor_IPI = Txt_valor_IPI_prod
                
                Txt_aliquota_PIS_prod = IIf(IsNull(TBFI!PIS_Prod), "0,00", Format(TBFI!PIS_Prod, "###,##0.00"))
                PIS_Prod = Txt_aliquota_PIS_prod
                Txt_valor_PIS_prod = IIf(IsNull(TBAbrir!Total_PIS_prod), "0,00", Format(TBAbrir!Total_PIS_prod, "###,##0.00"))
                Valor_PIS_Prod = Txt_valor_PIS_prod
                
                Txt_aliquota_Cofins_prod = IIf(IsNull(TBFI!Cofins_Prod), "0,00", Format(TBFI!Cofins_Prod, "###,##0.00"))
                Cofins_Prod = Txt_aliquota_Cofins_prod
                Txt_valor_Cofins_prod = IIf(IsNull(TBAbrir!Total_Cofins_prod), "0,00", Format(TBAbrir!Total_Cofins_prod, "###,##0.00"))
                Valor_Cofins_Prod = Txt_valor_Cofins_prod
                
                Txt_aliquota_CSLL_prod = IIf(IsNull(TBFI!CSLL_Prod), "0,00", Format(TBFI!CSLL_Prod, "###,##0.00"))
                CSLL_Prod = Txt_aliquota_CSLL_prod
                Txt_valor_CSLL_prod = IIf(IsNull(TBAbrir!Total_CSLL_prod), "0,00", Format(TBAbrir!Total_CSLL_prod, "###,##0.00"))
                Valor_CSLL_Prod = Txt_valor_CSLL_prod
                
                Txt_aliquota_IRPJ_prod = IIf(IsNull(TBFI!IRPJ_Prod), "0,00", Format(TBFI!IRPJ_Prod, "###,##0.00"))
                IRPJ_Prod = Txt_aliquota_IRPJ_prod
                Txt_valor_IRPJ_prod = IIf(IsNull(TBAbrir!Total_IRPJ_prod), "0,00", Format(TBAbrir!Total_IRPJ_prod, "###,##0.00"))
                Valor_IRPJ_Prod = Txt_valor_IRPJ_prod
                
                Txt_aliquota_total_prod = Format(PIS_Prod + Cofins_Prod + CSLL_Prod + IRPJ_Prod, "###,##0.00")
                Txt_valor_total_prod = Format(ValorICMS + Valor_IPI + Valor_PIS_Prod + Valor_Cofins_Prod + Valor_CSLL_Prod + Valor_IRPJ_Prod, "###,##0.00")
                
                Txt_aliquota_ICMS_SN = IIf(IsNull(TBFI!ICMS_SN), "0,00", Format(TBFI!ICMS_SN, "###,##0.00"))
                ICMS_SN = Txt_aliquota_ICMS_SN
                Txt_valor_ICMS_SN = IIf(IsNull(TBAbrir!Valor_total_ICMS_SN), "0,00", Format(TBAbrir!Valor_total_ICMS_SN, "###,##0.00"))
                Valor_ICMS_SN = Txt_valor_ICMS_SN
                
                If Permitido = True Then
                    'Retenção
                    Txt_valor_PIS_prod1 = IIf(IsNull(TBAbrir!Total_retencao_PIS), "0,00", Format(TBAbrir!Total_retencao_PIS, "###,##0.00"))
                    Valor_PIS_Prod = Txt_valor_PIS_prod1
                    
                    Txt_valor_Cofins_prod1 = IIf(IsNull(TBAbrir!Total_retencao_Cofins), "0,00", Format(TBAbrir!Total_retencao_Cofins, "###,##0.00"))
                    Valor_Cofins_Prod = Txt_valor_Cofins_prod1
                    
                    Txt_valor_total_prod1 = Format(Valor_PIS_Prod + Valor_Cofins_Prod, "###,##0.00")
                End If
            
                'Serviços
                'Destaque
                ProcControleImposto IIf(IsNull(TBFI!ID_CFOP), 0, TBFI!ID_CFOP), IIf(frmFaturamento_Prod_Serv.txtIDcliente = "", 0, frmFaturamento_Prod_Serv.txtIDcliente)
                If DestacaImpostos = "SIM" Then
                    Txt_aliquota_PIS_serv = IIf(IsNull(TBFI!PIS_Serv), "0,00", Format(TBFI!PIS_Serv, "###,##0.00"))
                    PIS_Serv = Txt_aliquota_PIS_serv
                    Txt_valor_PIS_serv = IIf(IsNull(TBAbrir!Total_PIS_serv), "0,00", Format(TBAbrir!Total_PIS_serv, "###,##0.00"))
                    Valor_PIS_Serv = Txt_valor_PIS_serv
                    
                    Txt_aliquota_Cofins_serv = IIf(IsNull(TBFI!Cofins_Serv), "0,00", Format(TBFI!Cofins_Serv, "###,##0.00"))
                    Cofins_Serv = Txt_aliquota_Cofins_serv
                    Txt_valor_Cofins_serv = IIf(IsNull(TBAbrir!Total_Cofins_serv), "0,00", Format(TBAbrir!Total_Cofins_serv, "###,##0.00"))
                    Valor_Cofins_Serv = Txt_valor_Cofins_serv
                    
                    Txt_aliquota_CSLL_serv = IIf(IsNull(TBFI!CSLL_Serv), "0,00", Format(TBFI!CSLL_Serv, "###,##0.00"))
                    CSLL_Serv = Txt_aliquota_CSLL_serv
                    Txt_valor_CSLL_serv = IIf(IsNull(TBAbrir!Total_CSLL_serv), "0,00", Format(TBAbrir!Total_CSLL_serv, "###,##0.00"))
                    Valor_CSLL_Serv = Txt_valor_CSLL_serv
                    
                    Txt_valor_ISSQN_serv = IIf(IsNull(TBAbrir!dbl_valor_total_iss), "0,00", Format(TBAbrir!dbl_valor_total_iss, "###,##0.00"))
                    Valor_ISS_Serv = Txt_valor_ISSQN_serv
                    
                    'INSS
                    Set TBFIltro = CreateObject("adodb.recordset")
                    TBFIltro.Open "Select INSS_Serv, Total_INSS_serv from tbl_Detalhes_Nota where ID_nota = " & ID & " and INSS_serv is not null and INSS_serv <> 0", Conexao, adOpenKeyset, adLockOptimistic
                    If TBFIltro.EOF = False Then
                        Txt_aliquota_INSS_serv = IIf(IsNull(TBFIltro!INSS_Serv), "0,00", Format(TBFIltro!INSS_Serv, "###,##0.00"))
                        INSS_Serv = Txt_aliquota_INSS_serv
                        Txt_valor_INSS_serv = IIf(IsNull(TBAbrir!Total_INSS_serv), "0,00", Format(TBAbrir!Total_INSS_serv, "###,##0.00"))
                        Valor_INSS_Serv = Txt_valor_INSS_serv
                    End If
                    TBFIltro.Close
                                    
                    Txt_aliquota_IRRF_serv = IIf(IsNull(TBFI!IRRF_Serv), "0,00", Format(TBFI!IRRF_Serv, "###,##0.00"))
                    IRRF_Serv = Txt_aliquota_IRRF_serv
                    Txt_valor_IRRF_serv = IIf(IsNull(TBAbrir!Total_IRRF_serv), "0,00", Format(TBAbrir!Total_IRRF_serv, "###,##0.00"))
                    Valor_IRRF_Serv = Txt_valor_IRRF_serv
                    
                    If TBAbrir!dbl_Valor_Total_Nota_Serv > 0 Then ISS_Serv = Format((Valor_ISS_Serv / IIf(IsNull(TBAbrir!dbl_Valor_Total_Nota_Serv), 0, TBAbrir!dbl_Valor_Total_Nota_Serv)) * 100, "0.00") Else ISS_Serv = 0
                    Txt_aliquota_total_serv = Format(PIS_Serv + Cofins_Serv + CSLL_Serv + ISS_Serv + INSS_Serv + IRRF_Serv, "###,##0.00")
                    Txt_valor_total_serv = Format(Valor_PIS_Serv + Valor_Cofins_Serv + Valor_CSLL_Serv + Valor_ISS_Serv + Valor_INSS_Serv + Valor_IRRF_Serv, "###,##0.00")
                End If
                
                If Permitido = True Then
                    'Retenção
                    'Cálculo com valores manuais
                    If IsNull(TBAbrir!Valor_Total_Retencao_Serv) = False Then
                        If TBFI!Retencao_PIS = True Then
                            Txt_aliquota_PIS_serv1 = IIf(IsNull(TBFI!PIS_Serv), "0,00", Format(TBFI!PIS_Serv, "###,##0.00"))
                            PIS_Serv = Txt_aliquota_PIS_serv1
                            Txt_valor_PIS_serv1 = IIf(IsNull(TBAbrir!Total_PIS_serv), "0,00", Format(TBAbrir!Total_PIS_serv, "###,##0.00"))
                            Valor_PIS_Serv = Txt_valor_PIS_serv1
                        End If
                        
                        If TBFI!Retencao_Cofins = True Then
                            Txt_aliquota_Cofins_serv1 = IIf(IsNull(TBFI!Cofins_Serv), "0,00", Format(TBFI!Cofins_Serv, "###,##0.00"))
                            Cofins_Serv = Txt_aliquota_Cofins_serv1
                            Txt_valor_Cofins_serv1 = IIf(IsNull(TBAbrir!Total_Cofins_serv), "0,00", Format(TBAbrir!Total_Cofins_serv, "###,##0.00"))
                            Valor_Cofins_Serv = Txt_valor_Cofins_serv1
                        End If
                        
                        If TBFI!Retencao_CSLL = True Then
                            Txt_aliquota_CSLL_serv1 = IIf(IsNull(TBFI!CSLL_Serv), "0,00", Format(TBFI!CSLL_Serv, "###,##0.00"))
                            CSLL_Serv = Txt_aliquota_CSLL_serv1
                            Txt_valor_CSLL_serv1 = IIf(IsNull(TBAbrir!Total_CSLL_serv), "0,00", Format(TBAbrir!Total_CSLL_serv, "###,##0.00"))
                            Valor_CSLL_Serv = Txt_valor_CSLL_serv1
                        End If
                                        
                        If TBFI!Retencao_INSS = True Then
                            Txt_aliquota_INSS_serv1 = IIf(IsNull(TBFI!INSS_Serv), "0,00", Format(TBFI!INSS_Serv, "###,##0.00"))
                            INSS_Serv = Txt_aliquota_INSS_serv1
                            Txt_valor_INSS_serv1 = IIf(IsNull(TBAbrir!Total_INSS_serv), "0,00", Format(TBAbrir!Total_INSS_serv, "###,##0.00"))
                            Valor_INSS_Serv = Txt_valor_INSS_serv1
                        End If
                        
                        If TBFI!Retencao_IRRF = True Then
                            Txt_aliquota_IRRF_serv1 = IIf(IsNull(TBFI!IRRF_Serv), "0,00", Format(TBFI!IRRF_Serv, "###,##0.00"))
                            IRRF_Serv = Txt_aliquota_IRRF_serv1
                            Txt_valor_IRRF_serv1 = IIf(IsNull(TBAbrir!Total_IRRF_serv), "0,00", Format(TBAbrir!Total_IRRF_serv, "###,##0.00"))
                            Valor_IRRF_Serv = Txt_valor_IRRF_serv1
                        End If
                        
                        Txt_valor_total_serv1 = Format(TBAbrir!Valor_Total_Retencao_Serv, "###,##0.00")
                    Else
                        Set TBTotaisnota = CreateObject("adodb.recordset")
                        TBTotaisnota.Open "Select * from Impostos where Regime = " & Regime, Conexao, adOpenKeyset, adLockOptimistic
                        If TBTotaisnota.EOF = False Then
                            valor = IIf(IsNull(TBAbrir!dbl_Valor_Total_Nota_Serv), 0, TBAbrir!dbl_Valor_Total_Nota_Serv)
                            If valor > TBTotaisnota!Acima Then
                                Txt_aliquota_PIS_serv1 = IIf(IsNull(TBFI!PIS_Serv), "0,00", Format(TBFI!PIS_Serv, "###,##0.00"))
                                PIS_Serv = Txt_aliquota_PIS_serv1
                                Txt_valor_PIS_serv1 = IIf(IsNull(TBAbrir!Total_PIS_serv), "0,00", Format(TBAbrir!Total_PIS_serv, "###,##0.00"))
                                Valor_PIS_Serv = Txt_valor_PIS_serv1
                                
                                Txt_aliquota_Cofins_serv1 = IIf(IsNull(TBFI!Cofins_Serv), "0,00", Format(TBFI!Cofins_Serv, "###,##0.00"))
                                Cofins_Serv = Txt_aliquota_Cofins_serv1
                                Txt_valor_Cofins_serv1 = IIf(IsNull(TBAbrir!Total_Cofins_serv), "0,00", Format(TBAbrir!Total_Cofins_serv, "###,##0.00"))
                                Valor_Cofins_Serv = Txt_valor_Cofins_serv1
                                
                                Txt_aliquota_CSLL_serv1 = IIf(IsNull(TBFI!CSLL_Serv), "0,00", Format(TBFI!CSLL_Serv, "###,##0.00"))
                                CSLL_Serv = Txt_aliquota_CSLL_serv1
                                Txt_valor_CSLL_serv1 = IIf(IsNull(TBAbrir!Total_CSLL_serv), "0,00", Format(TBAbrir!Total_CSLL_serv, "###,##0.00"))
                                Valor_CSLL_Serv = Txt_valor_CSLL_serv1
        
                                'INSS
                                Set TBFIltro = CreateObject("adodb.recordset")
                                TBFIltro.Open "Select * from tbl_Detalhes_Nota where ID_nota = " & ID & " and INSS_serv is not null and INSS_serv <> 0", Conexao, adOpenKeyset, adLockOptimistic
                                If TBFIltro.EOF = False Then
                                    Txt_aliquota_INSS_serv1 = IIf(IsNull(TBFIltro!INSS_Serv), "0,00", Format(TBFIltro!INSS_Serv, "###,##0.00"))
                                    INSS_Serv = Txt_aliquota_INSS_serv1
                                    Txt_valor_INSS_serv1 = IIf(IsNull(TBAbrir!Total_INSS_serv), "0,00", Format(TBAbrir!Total_INSS_serv, "###,##0.00"))
                                    Valor_INSS_Serv = Txt_valor_INSS_serv1
                                End If
                                TBFIltro.Close
                        
                                Txt_aliquota_IRRF_serv1 = IIf(IsNull(TBFI!IRRF_Serv), "0,00", Format(TBFI!IRRF_Serv, "###,##0.00"))
                                IRRF_Serv = Txt_aliquota_IRRF_serv1
                                Txt_valor_IRRF_serv1 = IIf(IsNull(TBAbrir!Total_IRRF_serv), "0,00", Format(TBAbrir!Total_IRRF_serv, "###,##0.00"))
                                Valor_IRRF_Serv = Txt_valor_IRRF_serv1
                                                        
                                Txt_aliquota_total_serv1 = Format(PIS_Serv + Cofins_Serv + CSLL_Serv + INSS_Serv + IRRF_Serv, "###,##0.00")
                                Txt_valor_total_serv1 = Format(Valor_PIS_Serv + Valor_Cofins_Serv + Valor_CSLL_Serv + Valor_INSS_Serv + Valor_IRRF_Serv, "###,##0.00")
                            ElseIf valor >= 667 And valor <= 5000 Then
                                    Txt_aliquota_IRRF_serv1 = IIf(IsNull(TBFI!IRRF_Serv), "0,00", Format(TBFI!IRRF_Serv, "###,##0.00"))
                                    IRRF_Serv = Txt_aliquota_IRRF_serv1
                                    Txt_valor_IRRF_serv1 = IIf(IsNull(TBAbrir!Total_IRRF_serv), "0,00", Format(TBAbrir!Total_IRRF_serv, "###,##0.00"))
                                    Valor_IRRF_Serv = Txt_valor_IRRF_serv1
                                    
                                    Txt_aliquota_total_serv1 = Format(IRRF_Serv, "###,##0.00")
                                    Txt_valor_total_serv1 = Format(Valor_IRRF_Serv, "###,##0.00")
                            End If
                        End If
                        TBTotaisnota.Close
                    End If
                End If
            End If
            'Produtos e serviços
            Txt_aliquota_DAS = IIf(IsNull(TBAbrir!DAS), "0,00", Format(TBAbrir!DAS, "###,##0.00"))
            DAS = Txt_aliquota_DAS
            Txt_valor_DAS = IIf(IsNull(TBAbrir!Total_DAS), "0,00", Format(TBAbrir!Total_DAS, "###,##0.00"))
            Valor_DAS = Txt_valor_DAS
            Txt_aliquota_total_prod_serv = Format(ICMS_SN + DAS, "###,##0.00")
            Txt_valor_total_prod_serv = Format(Valor_ICMS_SN + Valor_DAS, "###,##0.00")
            
            'Valor total aproximado de tributos
            Txt_valor_total_aprox_tributos = IIf(IsNull(TBAbrir!Valor_total_aprox_tributos), "0,00", Format(TBAbrir!Valor_total_aprox_tributos, "###,##0.00"))
        End If
    End If
    'Total geral destacado
    Valor_Produto = Txt_aliquota_total_prod
    valor = Txt_aliquota_total_serv
    Valor1 = Txt_aliquota_total_prod_serv
    Txt_aliquota_total_geral = Format(Valor_Produto + valor + Valor1, "###,##0.00")
    
    Valor_Produto = Txt_valor_total_prod
    valor = Txt_valor_total_serv
    Valor1 = Txt_valor_total_prod_serv
    Txt_valor_total_geral = Format(Valor_Produto + valor + Valor1, "###,##0.00")
    
    'Total geral retido
    Txt_aliquota_total_geral1 = Txt_aliquota_total_serv1
    
    Valor_Produto = Txt_valor_total_prod1
    valor = Txt_valor_total_serv1
    Txt_valor_total_geral1 = Format(Valor_Produto + valor, "###,##0.00")
End If
TBFI.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Txt_valor_total_serv1_Change()
On Error GoTo tratar_erro

If Txt_valor_total_serv1.Text <> "" Then
    VerifNumero = Txt_valor_total_serv1.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_valor_total_serv1.Text = ""
        Txt_valor_total_serv1.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_valor_total_serv1_LostFocus()
On Error GoTo tratar_erro

Txt_valor_total_serv1 = Format(Txt_valor_total_serv1, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

