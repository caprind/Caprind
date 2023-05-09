VERSION 5.00
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmQualidadePPAP 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Qualidade - PPAP - PSW"
   ClientHeight    =   10035
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   15360
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10035
   ScaleWidth      =   15360
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
      ForeColor       =   &H00000000&
      Height          =   2265
      Left            =   55
      TabIndex        =   61
      Top             =   990
      Width           =   15225
      Begin VB.CommandButton cmdLocalizarContatoCliente 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   11190
         Picture         =   "frmQualidadePPAP.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Localizar contato do cliente."
         Top             =   1080
         Width           =   315
      End
      Begin VB.TextBox txtEmail 
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
         Left            =   11610
         MaxLength       =   50
         TabIndex        =   14
         ToolTipText     =   "E-mail do responsável."
         Top             =   1080
         Width           =   3405
      End
      Begin VB.TextBox txtContato 
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
         Left            =   7770
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Contato cliente."
         Top             =   1080
         Width           =   3405
      End
      Begin VB.TextBox txtData_Validacao 
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
         Left            =   9360
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Data do cadastro."
         Top             =   390
         Width           =   2025
      End
      Begin VB.TextBox txtResp_validacao 
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
         Left            =   11400
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Responsável."
         Top             =   390
         Width           =   3615
      End
      Begin VB.CommandButton cmdFiltrar_codigo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   2100
         Picture         =   "frmQualidadePPAP.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Filtrar por código interno."
         Top             =   1800
         Width           =   315
      End
      Begin VB.TextBox txtDataStatus 
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
         Left            =   2415
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Data da revisão."
         Top             =   390
         Width           =   915
      End
      Begin VB.TextBox txtIDCliente 
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
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Código do cliente."
         Top             =   1080
         Width           =   585
      End
      Begin VB.TextBox txtFamiliaProduto 
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
         Left            =   6150
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   "Família."
         Top             =   1800
         Width           =   2715
      End
      Begin VB.TextBox txtDescricaoProduto 
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
         Left            =   8880
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   22
         ToolTipText     =   "Descrição."
         Top             =   1800
         Width           =   6135
      End
      Begin VB.TextBox txtUnidade 
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
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   20
         ToolTipText     =   "Unidade."
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox txtRevProduto 
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
         Left            =   2850
         Locked          =   -1  'True
         TabIndex        =   18
         ToolTipText     =   "Revisão."
         Top             =   1800
         Width           =   495
      End
      Begin VB.TextBox txtCodInterno 
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
         MaxLength       =   50
         TabIndex        =   15
         ToolTipText     =   "Código interno."
         Top             =   1800
         Width           =   1905
      End
      Begin VB.CommandButton cmdLocalizarProduto 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   2430
         Picture         =   "frmQualidadePPAP.frx":051D
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Localizar código interno."
         Top             =   1800
         Width           =   315
      End
      Begin VB.ComboBox cmbReferencia_prod 
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
         Left            =   3360
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   19
         ToolTipText     =   "Código de referência."
         Top             =   1800
         Width           =   2385
      End
      Begin VB.TextBox txtDataEmissao 
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
         Left            =   4740
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Data do cadastro."
         Top             =   390
         Width           =   915
      End
      Begin VB.TextBox txtRespPPAP 
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
         Left            =   5670
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Responsável pelo cadastro."
         Top             =   390
         Width           =   3675
      End
      Begin VB.TextBox txtPPAP 
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
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   0
         ToolTipText     =   "Número do PPAP."
         Top             =   390
         Width           =   1635
      End
      Begin VB.TextBox txtRevPPAP 
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
         Left            =   1830
         Locked          =   -1  'True
         TabIndex        =   1
         ToolTipText     =   "Revisão."
         Top             =   390
         Width           =   570
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
         Left            =   780
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   10
         ToolTipText     =   "Cliente."
         Top             =   1080
         Width           =   6555
      End
      Begin VB.CommandButton cmdLocalizarCliente 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   7350
         Picture         =   "frmQualidadePPAP.frx":061F
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Localizar cliente."
         Top             =   1080
         Width           =   315
      End
      Begin VB.TextBox txtStatus 
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
         Left            =   3360
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   3
         TabStop         =   0   'False
         Text            =   "Revisado"
         ToolTipText     =   "Status."
         Top             =   390
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.ComboBox cmbStatus 
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
         ItemData        =   "frmQualidadePPAP.frx":0721
         Left            =   3360
         List            =   "frmQualidadePPAP.frx":072B
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         ToolTipText     =   "Status."
         Top             =   390
         Width           =   1365
      End
      Begin VB.Label Label29 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Responsável validação"
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
         Left            =   12375
         TabIndex        =   103
         Top             =   180
         Width           =   1665
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "E-mail"
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
         Left            =   13102
         TabIndex        =   102
         Top             =   870
         Width           =   420
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contato"
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
         Left            =   9180
         TabIndex        =   101
         Top             =   870
         Width           =   585
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data/hora da validação"
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
         Left            =   9532
         TabIndex        =   100
         Top             =   180
         Width           =   1680
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Dt. revisão"
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
         Left            =   2475
         TabIndex        =   74
         Top             =   180
         Width           =   795
      End
      Begin VB.Label Label17 
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
         Left            =   11610
         TabIndex        =   73
         Top             =   1590
         Width           =   690
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Família"
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
         Left            =   7260
         TabIndex        =   72
         Top             =   1590
         Width           =   480
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
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
         Index           =   0
         Left            =   2910
         TabIndex        =   71
         Top             =   1590
         Width           =   375
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código de referência"
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
         Left            =   3795
         TabIndex        =   70
         Top             =   1590
         Width           =   1500
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Un."
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
         Left            =   5820
         TabIndex        =   69
         Top             =   1590
         Width           =   255
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código interno"
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
         TabIndex        =   68
         Top             =   1590
         Width           =   1050
      End
      Begin VB.Label Label21 
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
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   3810
         TabIndex        =   67
         Top             =   180
         Width           =   465
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
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
         Left            =   5010
         TabIndex        =   66
         Top             =   180
         Width           =   375
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
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
         Left            =   7035
         TabIndex        =   65
         Top             =   180
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Número PPAP"
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
         Left            =   510
         TabIndex        =   64
         Top             =   180
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
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
         Index           =   1
         Left            =   1928
         TabIndex        =   63
         Top             =   180
         Width           =   375
      End
      Begin VB.Label Label2 
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
         Left            =   3825
         TabIndex        =   62
         Top             =   870
         Width           =   495
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Certificado de submissão"
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
      Height          =   3165
      Left            =   55
      TabIndex        =   82
      Top             =   3270
      Width           =   15225
      Begin VB.TextBox txtNivel2 
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
         Left            =   10560
         MaxLength       =   50
         TabIndex        =   28
         ToolTipText     =   "Nivel de alteração de engenharia."
         Top             =   540
         Width           =   975
      End
      Begin VB.TextBox txtData3 
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
         Left            =   11550
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   29
         ToolTipText     =   "Data do nivel de alteração de engenharia."
         Top             =   540
         Width           =   915
      End
      Begin VB.TextBox txtData2 
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
         Left            =   6750
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   26
         ToolTipText     =   "Data das alterações adicionais de engenharia."
         Top             =   540
         Width           =   915
      End
      Begin VB.TextBox txtData 
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
         Left            =   2610
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   24
         ToolTipText     =   "Data do nivel de alteração de engenharia."
         Top             =   540
         Width           =   915
      End
      Begin VB.TextBox txtCodFornecedor 
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
         Left            =   5520
         MaxLength       =   50
         TabIndex        =   34
         ToolTipText     =   "Código do fornecedor."
         Top             =   1230
         Width           =   1215
      End
      Begin VB.TextBox txtAlteracoes 
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
         Left            =   3960
         MaxLength       =   50
         TabIndex        =   25
         ToolTipText     =   "Alterações adicionais de engenharia"
         Top             =   540
         Width           =   2775
      End
      Begin VB.CheckBox chkSIM 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Sim"
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
         Left            =   4350
         TabIndex        =   32
         Top             =   1290
         Width           =   555
      End
      Begin VB.CheckBox chkNao 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Não"
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
         Left            =   4920
         TabIndex        =   33
         Top             =   1290
         Width           =   585
      End
      Begin VB.CheckBox chkNA3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "N/A"
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
         Left            =   14460
         TabIndex        =   38
         Top             =   1290
         Width           =   585
      End
      Begin VB.CheckBox chkNao4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Não"
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
         Left            =   13830
         TabIndex        =   37
         Top             =   1290
         Width           =   585
      End
      Begin VB.CheckBox chkSim4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Sim"
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
         Left            =   13260
         TabIndex        =   36
         Top             =   1290
         Width           =   555
      End
      Begin VB.TextBox txtAplicacao 
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
         Left            =   6750
         MaxLength       =   50
         TabIndex        =   35
         ToolTipText     =   "Aplicação."
         Top             =   1230
         Width           =   855
      End
      Begin VB.TextBox txtAuxilio 
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
         MaxLength       =   50
         TabIndex        =   27
         ToolTipText     =   "Auxilio para verificação número."
         Top             =   540
         Width           =   2445
      End
      Begin VB.TextBox txtPeso 
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
         Left            =   14250
         MaxLength       =   50
         TabIndex        =   31
         ToolTipText     =   "Peso."
         Top             =   540
         Width           =   765
      End
      Begin VB.TextBox txtPedido 
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
         Left            =   12870
         MaxLength       =   50
         TabIndex        =   30
         ToolTipText     =   "Pedido de compra."
         Top             =   540
         Width           =   1365
      End
      Begin VB.TextBox txtNivel 
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
         MaxLength       =   50
         TabIndex        =   23
         ToolTipText     =   "Nivel de alteração de engenharia."
         Top             =   540
         Width           =   2415
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Informações de material"
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
         Height          =   1455
         Left            =   120
         TabIndex        =   83
         Top             =   1620
         Width           =   8955
         Begin VB.CheckBox chkNA2 
            BackColor       =   &H00E0E0E0&
            Caption         =   "N/A"
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
            Left            =   8160
            TabIndex        =   46
            Top             =   990
            Width           =   585
         End
         Begin VB.CheckBox chkNao3 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Não"
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
            Left            =   7470
            TabIndex        =   45
            Top             =   990
            Width           =   585
         End
         Begin VB.CheckBox chkSim3 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Sim"
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
            Left            =   6810
            TabIndex        =   44
            Top             =   990
            Width           =   555
         End
         Begin VB.TextBox txtIMDS 
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
            Left            =   4020
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   43
            TabStop         =   0   'False
            ToolTipText     =   "Submetido por IMDS ou outro formato do cliente."
            Top             =   630
            Width           =   4755
         End
         Begin VB.CheckBox chkIMDS 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Submetido por IMDS ou outro formato do cliente"
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
            TabIndex        =   42
            Top             =   690
            Width           =   3795
         End
         Begin VB.CheckBox chkNA 
            BackColor       =   &H00E0E0E0&
            Caption         =   "N/A"
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
            Left            =   8160
            TabIndex        =   41
            Top             =   360
            Width           =   585
         End
         Begin VB.CheckBox chkNao2 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Não"
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
            Left            =   7470
            TabIndex        =   40
            Top             =   360
            Width           =   585
         End
         Begin VB.CheckBox chkSim2 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Sim"
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
            Left            =   6810
            TabIndex        =   39
            Top             =   360
            Width           =   555
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "As peças poliméricas estão identificadas com os códigos de marcação ISO apropriados?"
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
            TabIndex        =   86
            Top             =   990
            Width           =   6240
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   3270
            TabIndex        =   85
            Top             =   750
            Width           =   45
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Foram reportadas informações requeridas pelo cliente de substâncias restritas?"
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
            TabIndex        =   84
            Top             =   360
            Width           =   5700
         End
      End
      Begin DrawSuite2022.USButton cmdSubmissaoRazao 
         Height          =   1305
         Left            =   9150
         TabIndex        =   47
         Top             =   1740
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   2302
         Caption         =   "Razão"
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
      Begin DrawSuite2022.USButton cmdSubmissaoNivel 
         Height          =   1305
         Left            =   10620
         TabIndex        =   48
         Top             =   1740
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   2302
         Caption         =   "Nível"
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
      Begin DrawSuite2022.USButton cmdSubmissaoResultados 
         Height          =   1305
         Left            =   12090
         TabIndex        =   49
         Top             =   1740
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   2302
         Caption         =   "Resultados"
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
      Begin DrawSuite2022.USButton cmdSubmissaoDeclaracao 
         Height          =   1305
         Left            =   13560
         TabIndex        =   50
         Top             =   1740
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   2302
         Caption         =   "Declaração"
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
      Begin VB.Image imgCalendario3 
         Height          =   360
         Left            =   12465
         Picture         =   "frmQualidadePPAP.frx":0741
         Stretch         =   -1  'True
         ToolTipText     =   "Abrir calendário."
         Top             =   510
         Width           =   330
      End
      Begin VB.Image imgCalendario2 
         Height          =   360
         Left            =   7680
         Picture         =   "frmQualidadePPAP.frx":0BC4
         Stretch         =   -1  'True
         ToolTipText     =   "Abrir calendário."
         Top             =   510
         Width           =   330
      End
      Begin VB.Image imgCalendario 
         Height          =   360
         Left            =   3540
         Picture         =   "frmQualidadePPAP.frx":1047
         Stretch         =   -1  'True
         ToolTipText     =   "Abrir calendário."
         Top             =   510
         Width           =   330
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cód. fornecedor"
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
         Left            =   5535
         TabIndex        =   99
         Top             =   1020
         Width           =   1185
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alterações adicionais de engenharia"
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
         Left            =   3967
         TabIndex        =   98
         Top             =   330
         Width           =   2580
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Item de segurança e/ou regulamentação governamental:"
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
         TabIndex        =   97
         Top             =   1290
         Width           =   4110
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cada ferramental do cliente está adequadamente etiquetado e identificado?"
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
         Left            =   7680
         TabIndex        =   96
         Top             =   1290
         Width           =   5475
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Aplicação"
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
         Left            =   6840
         TabIndex        =   95
         Top             =   1020
         Width           =   675
      End
      Begin VB.Label Label20 
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
         Left            =   11835
         TabIndex        =   94
         Top             =   330
         Width           =   345
      End
      Begin VB.Label Label19 
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
         Left            =   7035
         TabIndex        =   93
         Top             =   330
         Width           =   345
      End
      Begin VB.Label Label16 
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
         Left            =   2835
         TabIndex        =   92
         Top             =   330
         Width           =   345
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel"
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
         Left            =   10875
         TabIndex        =   91
         Top             =   330
         Width           =   345
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Auxilio para verificação número"
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
         Left            =   8197
         TabIndex        =   90
         Top             =   330
         Width           =   2250
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Peso"
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
         Left            =   14460
         TabIndex        =   89
         Top             =   330
         Width           =   345
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pedido de compra"
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
         Left            =   12915
         TabIndex        =   88
         Top             =   330
         Width           =   1275
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel de alteração de engenharia"
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
         Left            =   202
         TabIndex        =   87
         Top             =   330
         Width           =   2370
      End
   End
   Begin VB.TextBox txtIDProduto 
      Height          =   345
      Left            =   3360
      TabIndex        =   81
      Text            =   "0"
      Top             =   8190
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.TextBox txtIDPPAP 
      Height          =   345
      Left            =   2880
      TabIndex        =   80
      Text            =   "0"
      Top             =   8220
      Visible         =   0   'False
      Width           =   345
   End
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   12630
      Top             =   240
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frmQualidadePPAP.frx":14CA
      Count           =   1
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   55
      TabIndex        =   75
      Top             =   9120
      Width           =   15195
      Begin VB.ComboBox Cmb_opcao_lista 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         ItemData        =   "frmQualidadePPAP.frx":97E1
         Left            =   6990
         List            =   "frmQualidadePPAP.frx":97EB
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   210
         Width           =   1965
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
         TabIndex        =   54
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
         Left            =   3000
         TabIndex        =   52
         Text            =   "30"
         ToolTipText     =   "Número de registros por página."
         Top             =   180
         Width           =   555
      End
      Begin DrawSuite2022.USButton cmdPagProx 
         Height          =   315
         Left            =   11760
         TabIndex        =   58
         ToolTipText     =   "Próxima página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmQualidadePPAP.frx":9803
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
         TabIndex        =   57
         ToolTipText     =   "Página anterior."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmQualidadePPAP.frx":CFA7
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
         TabIndex        =   55
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
         TabIndex        =   56
         ToolTipText     =   "Primeira página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmQualidadePPAP.frx":10AB0
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
         TabIndex        =   59
         ToolTipText     =   "Última página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmQualidadePPAP.frx":14B9F
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
      Begin VB.Label Label31 
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
         Left            =   3630
         TabIndex        =   105
         Top             =   240
         Width           =   1440
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Operação da lista"
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
         Left            =   5640
         TabIndex        =   104
         Top             =   210
         Width           =   1260
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
         TabIndex        =   78
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
         TabIndex        =   77
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label Label34 
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
         Left            =   2310
         TabIndex        =   76
         Top             =   240
         Width           =   645
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   60
      Top             =   0
      Width           =   15225
      _ExtentX        =   26855
      _ExtentY        =   1720
      ButtonCount     =   14
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft2     =   40
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft3     =   78
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
      ButtonLeft4     =   124
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
      ButtonLeft5     =   171
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
      ButtonLeft6     =   233
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
      ButtonLeft7     =   290
      ButtonTop7      =   2
      ButtonWidth7    =   55
      ButtonHeight7   =   21
      ButtonUseMaskColor7=   0   'False
      ButtonCaption8  =   "Copiar"
      ButtonEnabled8  =   0   'False
      ButtonIconSize8 =   32
      ButtonToolTipText8=   "Copiar (F7)"
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
      ButtonLeft8     =   347
      ButtonTop8      =   2
      ButtonWidth8    =   39
      ButtonHeight8   =   21
      ButtonUseMaskColor8=   0   'False
      ButtonCaption9  =   "Revisar"
      ButtonEnabled9  =   0   'False
      ButtonIconSize9 =   32
      ButtonToolTipText9=   "Revisar (F8)"
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
      ButtonLeft9     =   388
      ButtonTop9      =   2
      ButtonWidth9    =   44
      ButtonHeight9   =   21
      ButtonUseMaskColor9=   0   'False
      ButtonCaption10 =   "Validação"
      ButtonEnabled10 =   0   'False
      ButtonIconSize10=   32
      ButtonToolTipText10=   "Validação (F9)"
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
      ButtonLeft10    =   434
      ButtonTop10     =   2
      ButtonWidth10   =   53
      ButtonHeight10  =   21
      ButtonUseMaskColor10=   0   'False
      ButtonEnabled11 =   0   'False
      ButtonIconSize11=   32
      ButtonAlignment11=   2
      ButtonType11    =   1
      ButtonStyle11   =   -1
      BeginProperty ButtonFont11 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState11   =   -1
      ButtonLeft11    =   489
      ButtonTop11     =   4
      ButtonWidth11   =   2
      ButtonHeight11  =   54
      ButtonCaption12 =   "Ajuda"
      ButtonEnabled12 =   0   'False
      ButtonIconSize12=   32
      ButtonToolTipText12=   "Ajuda (F1)"
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
      ButtonLeft12    =   493
      ButtonTop12     =   2
      ButtonWidth12   =   41
      ButtonHeight12  =   21
      ButtonUseMaskColor12=   0   'False
      ButtonCaption13 =   "Sair"
      ButtonEnabled13 =   0   'False
      ButtonIconSize13=   32
      ButtonToolTipText13=   "Sair (Esc)"
      ButtonKey13     =   "13"
      ButtonAlignment13=   2
      BeginProperty ButtonFont13 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft13    =   536
      ButtonTop13     =   2
      ButtonWidth13   =   30
      ButtonHeight13  =   21
      ButtonUseMaskColor13=   0   'False
      ButtonEnabled14 =   0   'False
      ButtonIconSize14=   32
      ButtonKey14     =   "14"
      ButtonAlignment14=   2
      BeginProperty ButtonFont14 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState14   =   5
      ButtonLeft14    =   568
      ButtonTop14     =   2
      ButtonWidth14   =   24
      ButtonHeight14  =   24
      ButtonUseMaskColor14=   0   'False
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   2655
      Left            =   60
      TabIndex        =   51
      Top             =   6450
      Width           =   15225
      _ExtentX        =   26855
      _ExtentY        =   4683
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
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "D"
         Text            =   "Data"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Responsável"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "Nº PPAP"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Object.Tag             =   "N"
         Text            =   "Rev."
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Tag             =   "T"
         Text            =   "Cliente"
         Object.Width           =   6145
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Object.Tag             =   "T"
         Text            =   "Cód. interno"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Object.Tag             =   "T"
         Text            =   "Descrição"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Object.Tag             =   "T"
         Text            =   "Status"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Object.Tag             =   "T"
         Text            =   "Validada"
         Object.Width           =   1499
      EndProperty
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   55
      TabIndex        =   79
      Top             =   9750
      Width           =   15225
      _ExtentX        =   26855
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
End
Attribute VB_Name = "frmQualidadePPAP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Novo_qualidadePPAP   As Boolean 'OK
Public SQL_qualidade_PPAP   As String 'OK
Public Data_PPAP            As String 'OK
Public PPAP                 As String 'OK
Public Sql_PPAP_Localizar   As String 'OK
Dim TBLISTA_Compras_PPAP    As ADODB.Recordset 'OK

Sub ProcLimpaCamposPPAP()
On Error GoTo tratar_erro

txtIDPPAP = 0
txtidproduto = 0
txtPPAP.Text = ""
txtRevPPAP.Text = ""
txtCodinterno.Text = ""
txtRevProduto.Text = ""
cmbReferencia_prod.Clear
txtunidade.Text = ""
txtDataemissao.Text = ""
txtDescricaoProduto.Text = ""
txtFamiliaProduto.Text = ""
cmbStatus.ListIndex = -1
txtRespPPAP.Text = ""
txtIDcliente = ""
txtCliente.Text = ""
txtStatus.Visible = False
cmbStatus.Visible = True
Caption = "Qualidade - PPAP - PSW"
ProcLimpaCamposSubmissao

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcEnviaDadosPAPP()
On Error GoTo tratar_erro

TBGravar!revPPAP = txtRevPPAP.Text
TBGravar!IDProduto = txtidproduto.Text
TBGravar!Codinterno = txtCodinterno.Text
TBGravar!Revproduto = IIf(txtRevProduto.Text = "", 0, txtRevProduto.Text)
TBGravar!N_referencia = cmbReferencia_prod
TBGravar!Descricao = txtDescricaoProduto.Text
TBGravar!IDCliente = IIf(txtIDcliente.Text = "", 0, txtIDcliente.Text)
TBGravar!DtEmissao = txtDataemissao.Text
TBGravar!Cliente = txtCliente.Text
TBGravar!status = cmbStatus.Text
ProcEnviaDadosSubmissao

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcEnviaDadosSubmissao()
On Error GoTo tratar_erro

TBGravar!Nivel = txtNivel.Text
If txtData.Text <> "" Then TBGravar!Data = Format(txtData.Text, "dd/mm/yy") Else TBGravar!Data = Null
TBGravar!Alteracoes = txtAlteracoes.Text
If txtData2.Text <> "" Then TBGravar!Data2 = Format(txtData2.Text, "dd/mm/yy") Else TBGravar!Data2 = Null
TBGravar!Auxilio = txtAuxilio.Text
TBGravar!Nivel2 = txtNivel2.Text
If txtData3.Text <> "" Then TBGravar!Data3 = Format(txtData3.Text, "dd/mm/yy") Else TBGravar!Data3 = Null
TBGravar!contato = txtContato.Text
TBGravar!Pedido = txtPedido.Text
TBGravar!Peso = IIf(txtpeso = "", Null, txtpeso)
If chkSim.Value = 1 Then TBGravar!Sim = True Else TBGravar!Sim = False
If chkNao.Value = 1 Then TBGravar!NAO = True Else TBGravar!NAO = False
TBGravar!CodFornecedor = txtCodFornecedor
TBGravar!Aplicacao = txtAplicacao.Text
If chkSim2.Value = 1 Then TBGravar!Sim2 = True Else TBGravar!Sim2 = False
If chkNao2.Value = 1 Then TBGravar!Nao2 = True Else TBGravar!Nao2 = False
If chkNA.Value = 1 Then TBGravar!NA = True Else TBGravar!NA = False
If chkIMDS.Value = 1 Then
    TBGravar!chkIMDS = True
    TBGravar!IMDS = txtIMDS
Else
    TBGravar!chkIMDS = False
    TBGravar!IMDS = Null
End If
If chkSim3.Value = 1 Then TBGravar!Sim3 = True Else TBGravar!Sim3 = False
If chkNao3.Value = 1 Then TBGravar!Nao3 = True Else TBGravar!Nao3 = False
If chkNA2.Value = 1 Then TBGravar!na2 = True Else TBGravar!na2 = False
If chkSim4.Value = 1 Then TBGravar!Sim4 = True Else TBGravar!Sim4 = False
If chkNao4.Value = 1 Then TBGravar!Nao4 = True Else TBGravar!Nao4 = False
If chkNA3.Value = 1 Then TBGravar!na3 = True Else TBGravar!na3 = False
TBGravar!Email = txtEmail.Text

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCamposSubmissao()
On Error GoTo tratar_erro

txtNivel.Text = ""
txtData = ""
txtAlteracoes.Text = ""
txtData2 = ""
txtAuxilio.Text = ""
txtData3 = ""
txtNivel2.Text = ""
txtContato.Text = ""
txtPedido.Text = ""
txtpeso.Text = ""
chkSim.Value = 0
chkNao.Value = 0
txtCodFornecedor.Text = ""
txtAplicacao.Text = ""
chkSim2.Value = 0
chkNao2.Value = 0
chkNA.Value = 0
chkIMDS.Value = 0
txtIMDS.Text = ""
chkSim3.Value = 0
chkNao3.Value = 0
chkNA2.Value = 0
chkSim4.Value = 0
chkNao4.Value = 0
chkNA3.Value = 0
txtData_Validacao.Text = ""
txtResp_validacao.Text = ""
txtEmail.Text = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcPuxadadosSubmissao()
On Error GoTo tratar_erro

Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "SELECT * from QualidadePPAP where IDPPAP = " & txtIDPPAP.Text, Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    txtNivel.Text = IIf(IsNull(TBProduto!Nivel), "", TBProduto!Nivel)
    txtData.Text = IIf(IsNull(TBProduto!Data), "", Format(TBProduto!Data, "dd/mm/yy"))
    txtAlteracoes.Text = IIf(IsNull(TBProduto!Alteracoes), "", TBProduto!Alteracoes)
    txtData2.Text = IIf(IsNull(TBProduto!Data2), "", Format(TBProduto!Data2, "dd/mm/yy"))
    txtAuxilio.Text = IIf(IsNull(TBProduto!Auxilio), "", TBProduto!Auxilio)
    txtNivel2.Text = IIf(IsNull(TBProduto!Nivel2), "", TBProduto!Nivel2)
    txtData3.Text = IIf(IsNull(TBProduto!Data3), "", Format(TBProduto!Data3, "dd/mm/yy"))
    txtContato.Text = IIf(IsNull(TBProduto!contato), "", TBProduto!contato)
    txtPedido.Text = IIf(IsNull(TBProduto!Pedido), "", TBProduto!Pedido)
    txtpeso.Text = IIf(IsNull(TBProduto!Peso), "", TBProduto!Peso)
    If TBProduto!Sim = True Then chkSim.Value = 1 Else chkSim.Value = 0
    If TBProduto!NAO = True Then chkNao.Value = 1 Else chkNao.Value = 0
    txtCodFornecedor.Text = IIf(IsNull(TBProduto!CodFornecedor), "", TBProduto!CodFornecedor)
    txtAplicacao.Text = IIf(IsNull(TBProduto!Aplicacao), "", TBProduto!Aplicacao)
    If TBProduto!Sim2 = True Then chkSim2.Value = 1 Else chkSim2.Value = 0
    If TBProduto!Nao2 = True Then chkNao2.Value = 1 Else chkNao2.Value = 0
    If TBProduto!NA = True Then chkNA.Value = 1 Else chkNA.Value = 0
    If TBProduto!chkIMDS = True Then
        chkIMDS.Value = 1
        txtIMDS = IIf(IsNull(TBProduto!IMDS), "", TBProduto!IMDS)
    Else
        chkIMDS.Value = 0
        txtIMDS = ""
    End If
    If TBProduto!Sim3 = True Then chkSim3.Value = 1 Else chkSim3.Value = 0
    If TBProduto!Nao3 = True Then chkNao3.Value = 1 Else chkNao3.Value = 0
    If TBProduto!na2 = True Then chkNA2.Value = 1 Else chkNA2.Value = 0
    If TBProduto!Sim4 = True Then chkSim4.Value = 1 Else chkSim4.Value = 0
    If TBProduto!Nao4 = True Then chkNao4.Value = 1 Else chkNao4.Value = 0
    If TBProduto!na3 = True Then chkNA3.Value = 1 Else chkNA3.Value = 0
    txtData_Validacao = IIf(IsNull(TBProduto!DtValidacao), "", TBProduto!DtValidacao)
    txtResp_validacao = IIf(IsNull(TBProduto!RespValidacao), "", TBProduto!RespValidacao)
    txtEmail = IIf(IsNull(TBProduto!Email), "", TBProduto!Email)
End If
TBProduto.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkIMDS_Click()
On Error GoTo tratar_erro

If chkIMDS.Value = 1 Then
    txtIMDS.Locked = False
    txtIMDS.TabStop = True
Else
    txtIMDS.Locked = True
    txtIMDS.TabStop = False
    txtIMDS = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkNA_Click()
On Error GoTo tratar_erro

If chkNA.Value = 1 Then
    chkNao2.Value = 0
    chkSim2.Value = 0
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkna2_Click()
On Error GoTo tratar_erro

If chkNA2.Value = 1 Then
    chkNao3.Value = 0
    chkSim3.Value = 0
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkna3_Click()
On Error GoTo tratar_erro

If chkNA3.Value = 1 Then
    chkNao4.Value = 0
    chkSim4.Value = 0
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkNao_Click()
On Error GoTo tratar_erro

If chkNao.Value = 1 Then
    chkSim.Value = 0
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkNao2_Click()
On Error GoTo tratar_erro

If chkNao2.Value = 1 Then
    chkSim2.Value = 0
    chkNA.Value = 0
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkNao3_Click()
On Error GoTo tratar_erro

If chkNao3.Value = 1 Then
    chkSim3.Value = 0
    chkNA2.Value = 0
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkNao4_Click()
On Error GoTo tratar_erro

If chkNao4.Value = 1 Then
    chkSim4.Value = 0
    chkNA3.Value = 0
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkSim_Click()
On Error GoTo tratar_erro

If chkSim.Value = 1 Then
    chkNao.Value = 0
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkSim2_Click()
On Error GoTo tratar_erro

If chkSim2.Value = 1 Then
    chkNao2.Value = 0
    chkNA.Value = 0
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkSim3_Click()
On Error GoTo tratar_erro

If chkSim3.Value = 1 Then
    chkNao3.Value = 0
    chkNA2.Value = 0
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkSim4_Click()
On Error GoTo tratar_erro

If chkSim4.Value = 1 Then
    chkNao4.Value = 0
    chkNA3.Value = 0
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCopiar()
On Error GoTo tratar_erro

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If txtIDPPAP.Text = 0 Then
    USMsgBox ("Informe o PPAP antes de copiar."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Engenharia = False
Sit_REG = 2
frmQualidadePPAP_LocalizarProduto.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_opcao_lista_Click()
On Error GoTo tratar_erro

With Lista
    For InitFor = 1 To .ListItems.Count
        .ListItems.Item(InitFor).Checked = False
    Next InitFor
End With

With USToolBar1
    If Cmb_opcao_lista = "Excluir" Then
        .ButtonState(4) = 0
        .ButtonState(10) = 5
    Else
        .ButtonState(4) = 5
        .ButtonState(10) = 0
    End If
    .Refresh
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdFiltrar_codigo_Click()
On Error GoTo tratar_erro

If txtCodinterno <> "" Then
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select * from projproduto where desenho = '" & txtCodinterno & "' and Bloqueado = 'False' and (tipo = 'P' or tipo = 'I' or tipo = 'PI')", Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
        txtidproduto = 0
        txtRevProduto = ""
        cmbReferencia_prod.Clear
        txtunidade = ""
        txtFamiliaProduto = ""
        txtDescricaoProduto = ""
        txtCodinterno = IIf(IsNull(TBProduto!Desenho), "", TBProduto!Desenho)
        txtidproduto = TBProduto!Codproduto
        txtRevProduto = IIf(IsNull(TBProduto!RevDesenho), "", TBProduto!RevDesenho)
        txtunidade = IIf(IsNull(TBProduto!Unidade), "", TBProduto!Unidade)
        txtFamiliaProduto = IIf(IsNull(TBProduto!Classe), "", TBProduto!Classe)
        txtDescricaoProduto = IIf(IsNull(TBProduto!Descricao), "", TBProduto!Descricao)
        
        Set TBItem = CreateObject("adodb.recordset")
        TBItem.Open "Select N_Referencia from item_aplicacoes where codproduto = " & TBProduto!Codproduto, Conexao, adOpenKeyset, adLockOptimistic
        If TBItem.EOF = False Then
            cmbReferencia_prod.AddItem ""
            Do While TBItem.EOF = False
                cmbReferencia_prod.AddItem TBItem!N_referencia
                TBItem.MoveNext
            Loop
            TBItem.MoveFirst
            cmbReferencia_prod = TBItem!N_referencia
        End If
        TBItem.Close
    Else
        USMsgBox ("Não foi encontrado nenhum produto com este código interno."), vbExclamation, "CAPRIND v5.0"
        txtidproduto = 0
        txtRevProduto = ""
        cmbReferencia_prod.Clear
        txtunidade = ""
        txtFamiliaProduto = ""
        txtDescricaoProduto = ""
    End If
    TBProduto.Close
Else
    txtidproduto = 0
    txtRevProduto = ""
    cmbReferencia_prod.Clear
    txtunidade = ""
    txtFamiliaProduto = ""
    txtDescricaoProduto = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdLocalizarCliente_Click()
On Error GoTo tratar_erro

ProcConfVariaveisLocCliente False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, True, False, False, False, False, False, False, False, False, False, False, False, False, False
frmVendas_LocalizarCliente.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdLocalizarContatoCliente_Click()
On Error GoTo tratar_erro

If txtIDcliente.Text <> "" And txtIDcliente.Text <> "0" Then
    Analise_critica = False
    Vendas_Proposta = False
    Vendas_PI = False
    Telemarketing = False
    Qualidade_PPAP_PSW = True
    Financeiro_Contas_Pagar = False
    Financeiro_Contas_Pagas = False
    Financeiro_Contas_Receber = False
    Financeiro_Contas_Recebidas = False
    frmVendas_propostaII_contato.Show 1
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdLocalizarProduto_Click()
On Error GoTo tratar_erro

Engenharia = False
Sit_REG = 1
frmQualidadePPAP_LocalizarProduto.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcRevisao()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If txtIDPPAP.Text = 0 Then
    USMsgBox ("Informe o PPAP antes de criar a revisão."), vbExclamation, "CAPRIND v5.0"
    ProcLocalizar
    Exit Sub
End If
If USMsgBox("Deseja realmente criar uma revisão do " & txtPPAP.Text & "?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    If txtStatus.Visible = True Then
        USMsgBox ("Não é permitido revisar o PPAP revisado."), vbExclamation, "CAPRIND v5.0"
        Exit Sub
    End If
    '==================================
    Modulo = "Qualidade/PPAP/PSW"
    Evento = "Revisar"
    ID_documento = txtIDPPAP
    Documento = "Número PPAP: " & txtPPAP.Text & " - Cód. interno: " & txtCodinterno
    Documento1 = ""
    ProcGravaEvento
    '==================================
    Contador = 0
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from QualidadePPAP where IDPPAP = " & txtIDPPAP, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Contador = 0
        Set TBGravar = CreateObject("adodb.recordset")
        TBGravar.Open "Select * from QualidadePPAP where PPAP = '" & txtPPAP & "' order by RevPPAP", Conexao, adOpenKeyset, adLockOptimistic
        TBGravar.MoveLast
        Contador = TBAbrir!revPPAP
        Contador = Contador + 1
        
        TBGravar.AddNew
        TBGravar!PPAP = TBAbrir!PPAP
        TBGravar!revPPAP = Contador
        TBGravar!DtEmissao = Date
        TBGravar!status = "Aberto"
        ProcEnviaDadosRevisar
        
        TBAbrir!DataRevisao = Date
        TBAbrir!status = "Revisado"
        TBGravar.Update
        txtIDPPAP = TBGravar!idPPAP
        TBAbrir!IDRevisao = TBGravar!idPPAP
        TBAbrir.Update
        TBGravar.Close
    End If
    TBAbrir.Close
    USMsgBox ("PPAP revisado com sucesso."), vbInformation, "CAPRIND v5.0"
    
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from QualidadePPAP where IDPPAP = " & txtIDPPAP, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        ProcLimpaCamposPPAP
        ProcPuxaDadosPPAP
    End If
    TBAbrir.Close
    
    Lista.ListItems.Clear
    ProcCarregaLista (1)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Compras_PPAP.AbsolutePage <> 2 Then
    If TBLISTA_Compras_PPAP.AbsolutePage = -3 Then
        ProcExibePagina (TBLISTA_Compras_PPAP.PageCount - 1)
    Else
        TBLISTA_Compras_PPAP.AbsolutePage = TBLISTA_Compras_PPAP.AbsolutePage - 2
        ProcExibePagina (TBLISTA_Compras_PPAP.AbsolutePage)
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
    TBLISTA_Compras_PPAP.AbsolutePage = txtPagIr.Text
    ProcExibePagina (TBLISTA_Compras_PPAP.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Compras_PPAP.AbsolutePage = 1
ProcExibePagina (TBLISTA_Compras_PPAP.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Compras_PPAP.AbsolutePage <> -3 Then
    If TBLISTA_Compras_PPAP.AbsolutePage = 1 Then
        ProcExibePagina (2)
    Else
        ProcExibePagina (TBLISTA_Compras_PPAP.AbsolutePage)
    End If
Else
    ProcExibePagina (TBLISTA_Compras_PPAP.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Compras_PPAP.AbsolutePage = TBLISTA_Compras_PPAP.PageCount
ProcExibePagina (TBLISTA_Compras_PPAP.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdSubmissaoDeclaracao_Click()
On Error GoTo tratar_erro

frmQualidadePPAP_SubmissaoDeclaracao.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdSubmissaoNivel_Click()
On Error GoTo tratar_erro

frmQualidadePPAP_SubmissaoNivel.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdSubmissaoRazao_Click()
On Error GoTo tratar_erro

frmQualidadePPAP_SubmissaoRazao.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdSubmissaoResultados_Click()
On Error GoTo tratar_erro

frmQualidadePPAP_SubmissaoResultados.Show 1

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
    Case vbKeyF4: If Cmb_opcao_lista = "Excluir" Then ProcExcluir
    Case vbKeyF5: ProcImprimir
    Case vbKeyF7: ProcCopiar
    Case vbKeyF8: ProcRevisao
    Case vbKeyF9: If Cmb_opcao_lista = "Validação" Then ProcValidarRegistros Lista, "Qualidade/PPAP/PSW"
    Case vbKeyEscape: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15195, 14, True
Cmb_opcao_lista = "Validação"
Formulario = "Qualidade/PPAP/PSW"
Direitos
ProcLimpaVariaveisPrincipais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLocalizar()
On Error GoTo tratar_erro

frmQualidadePPAP_Localizar.Show 1

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
ProcLimpaCamposPPAP
Frame2.Enabled = True
Frame3.Enabled = True
txtDataemissao.Text = Format(Date, "dd/mm/yy")
txtRevPPAP = 0
txtRespPPAP.Text = pubUsuario
cmbStatus.Text = "Aberto"
Novo_qualidadePPAP = True
cmdLocalizarCliente_Click

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluir()
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
                If USMsgBox("Deseja realmente excluir este(s) PPAP?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            
            Conexao.Execute "Update QualidadePPAP Set DataRevisao = NULL, IDRevisao = 0, Status = 'Aberto' where idRevisao = " & .ListItems(InitFor)
            Conexao.Execute "DELETE from QualidadePPAP WHERE IDPPAP = " & .ListItems(InitFor)
            '==================================
            Modulo = "Qualidade/PPAP/PSW"
            Evento = "Excluir"
            ID_documento = .ListItems(InitFor)
            Documento = "Número PPAP: " & .ListItems(InitFor).SubItems(1) & " - Cód. interno: " & .ListItems(InitFor).SubItems(4)
            Documento1 = ""
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) PPAP(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("PPAP(s) excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    Novo_Nota = False
    ProcLimpaCamposPPAP
    Lista.ListItems.Clear
    ProcCarregaLista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
    Frame2.Enabled = False
    Frame3.Enabled = True
    Novo_qualidadePPAP = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

Formulario = "Qualidade/PPAP/PSW"
Direitos
ProcLimpaVariaveisPrincipais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

If Novo_qualidadePPAP = True Then
    If USMsgBox("O PPAP ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvar
        If Novo_qualidadePPAP = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
Novo_qualidadePPAP = False
Unload Me

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
If txtStatus.Visible = True Then
    USMsgBox ("Não é permitida a alteração do PPAP, pois o mesmo esta revisado."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Acao = "salvar"
If FunVerifValidacaoRegistro("alterar", txtData_Validacao, "mesmo", "o PPAP", True) = False Then Exit Sub
If txtIDcliente = "" Then
    NomeCampo = "o cliente"
    ProcVerificaAcao
    cmdLocalizarCliente_Click
    Exit Sub
End If
If txtidproduto = 0 Then
    NomeCampo = "o produto"
    ProcVerificaAcao
    txtCodinterno.SetFocus
    Exit Sub
End If
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "SELECT * from QualidadePPAP where IDPPAP = " & txtIDPPAP.Text, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then
    TBGravar.AddNew
    PPAP = ""
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from QualidadePPAP where year(DtEmissao) = " & Year(Date) & " order by idPPAP", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = True Then
        Cont = 1
        Data_PPAP = Format(Date, "yy")
    Else
        TBAbrir.MoveLast
        PPAP = Left(TBAbrir!PPAP, 9)
        Cont = ReturnNumbersOnly(PPAP)
        Cont = Cont + 1
        Data_PPAP = Format(TBAbrir!DtEmissao, "YY")
    End If
    TBAbrir.Close
    ProcGeraNumero
    txtPPAP = a
    
    TBGravar!PPAP = txtPPAP
    TBGravar!DtEmissao = Format(Date, "dd/mm/yy")
    TBGravar!revPPAP = 0
    TBGravar!Responsavel = pubUsuario
End If
ProcEnviaDadosPAPP
TBGravar.Update
txtIDPPAP = TBGravar!idPPAP
TBGravar.Close

Lista.ListItems.Clear
If Novo_qualidadePPAP = True Then
    USMsgBox ("Novo PPAP cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo"
    Sql_PPAP_Localizar = "Select * from QualidadePPAP where IDPPAP = " & txtIDPPAP.Text
    ProcCarregaLista (1)
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar"
    ProcCarregaLista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
    If CodigoLista <> 0 And Lista.ListItems.Count <> 0 Then
        Lista.SelectedItem = Lista.ListItems(CodigoLista)
        Lista.SetFocus
    End If
End If
'==================================
Modulo = "Qualidade/PPAP/PSW"
ID_documento = txtIDPPAP
Documento = "Número PPAP: " & txtPPAP.Text & " - Cód. interno: " & txtCodinterno
Documento1 = ""
ProcGravaEvento
'==================================
Novo_qualidadePPAP = False
Caption = "Qualidade - PPAP - PSW (PPAP : " & txtPPAP & " - Rev. : " & txtRevPPAP & ")"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub imgCalendario_Click()
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
RNC = False
SolicitacaoAcao = False
Troca_Duplicata = False
Financeiro_Contas_Recebidas = False
Engenharia_Normas = False
Qualidade_PPAP_PSW = True
Qualidade_PPAP_Plano = False
Qualidade_PPAP_FMEA = False
Qualidade_sistema = False
Engenharia = False
Compras_Fornecedores = False
Vendas_Programacao = False
Outros_solicitacaoPCP = False
Estoque_recebimento = False
Sit_Data = 1
FrmCalendario.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Imgcalendario2_Click()
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
RNC = False
SolicitacaoAcao = False
Troca_Duplicata = False
Financeiro_Contas_Recebidas = False
Engenharia_Normas = False
Qualidade_PPAP_PSW = True
Qualidade_PPAP_Plano = False
Qualidade_PPAP_FMEA = False
Qualidade_sistema = False
Engenharia = False
Compras_Fornecedores = False
Vendas_Programacao = False
Outros_solicitacaoPCP = False
Estoque_recebimento = False
Sit_Data = 2
FrmCalendario.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Imgcalendario3_Click()
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
RNC = False
SolicitacaoAcao = False
Troca_Duplicata = False
Financeiro_Contas_Recebidas = False
Engenharia_Normas = False
Qualidade_PPAP_PSW = True
Qualidade_PPAP_Plano = False
Qualidade_PPAP_FMEA = False
Qualidade_sistema = False
Engenharia = False
Compras_Fornecedores = False
Vendas_Programacao = False
Outros_solicitacaoPCP = False
Estoque_recebimento = False
Sit_Data = 3
FrmCalendario.Show 1

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
                If Cmb_opcao_lista = "Excluir" Then
                    Set TBAbrir = CreateObject("adodb.recordset")
                    TBAbrir.Open "Select * from QualidadePPAP where IDPPAP = " & .ListItems(InitFor) & " and idRevisao <> 0", Conexao, adOpenKeyset, adLockOptimistic
                    If TBAbrir.EOF = False Then
                        .ListItems(InitFor).Checked = False
                        GoTo Proximo
                    End If
                    TBAbrir.Close
                End If
                If .ListItems.Item(InitFor).ListSubItems(9) = "SIM" Then
                    .ListItems(InitFor).Checked = False
                    GoTo Proximo
                End If
                
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
            If Cmb_opcao_lista = "Excluir" Then
                Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open "Select * from QualidadePPAP where IDPPAP = " & .ListItems(InitFor) & " and idRevisao <> 0", Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    USMsgBox ("Não é permitido excluir este PPAP, pois o mesmo já foi revisado."), vbExclamation, "CAPRIND v5.0"
                    .ListItems(InitFor).Checked = False
                    TBAbrir.Close
                    Exit Sub
                End If
                TBAbrir.Close
            End If
            If .ListItems.Item(InitFor).ListSubItems(9) = "SIM" Then
                USMsgBox ("Não é permitido excluir este PPAP, pois o mesmo já foi validado."), vbExclamation, "CAPRIND v5.0"
                .ListItems(InitFor).Checked = False
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
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from QualidadePPAP where IDPPAP = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    ProcLimpaCamposPPAP
    ProcPuxaDadosPPAP
    CodigoLista = Lista.SelectedItem.index
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtCodInterno_Change()
On Error GoTo tratar_erro

txtidproduto = 0
txtRevProduto = ""
cmbReferencia_prod.Clear
txtunidade = ""
txtFamiliaProduto = ""
txtDescricaoProduto = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtIDCliente_LostFocus()
On Error GoTo tratar_erro

Set TBClientes = CreateObject("adodb.recordset")
TBClientes.Open "SELECT * from Clientes where IDCliente = " & txtIDcliente.Text, Conexao, adOpenKeyset, adLockOptimistic
If TBClientes.EOF = True Then
    USMsgBox ("Nenhum registro de cliente com o código informado!"), vbCritical, "CAPRIND v5.0" + vbOKOnly
    txtCliente.Text = ""
Else
    txtCliente.Text = TBClientes!NomeRazao
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcPuxaDadosPPAP()
On Error GoTo tratar_erro

txtIDPPAP.Text = IIf(IsNull(TBAbrir!idPPAP), "", TBAbrir!idPPAP)
txtPPAP.Text = IIf(IsNull(TBAbrir!PPAP), "", TBAbrir!PPAP)
txtRevPPAP.Text = IIf(IsNull(TBAbrir!revPPAP), "", TBAbrir!revPPAP)
txtCodinterno.Text = IIf(IsNull(TBAbrir!Codinterno), "", TBAbrir!Codinterno)
txtidproduto.Text = IIf(IsNull(TBAbrir!IDProduto), "0", TBAbrir!IDProduto)
txtIDcliente.Text = IIf(IsNull(TBAbrir!IDCliente), "", TBAbrir!IDCliente)
txtCliente.Text = IIf(IsNull(TBAbrir!Cliente), "", TBAbrir!Cliente)
If txtCodinterno <> "" Then
    Set TBItem = CreateObject("adodb.recordset")
    TBItem.Open "Select * from projproduto where desenho = '" & txtCodinterno & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBItem.EOF = False Then
        If txtIDcliente <> "" Then
            Set TBCiclo = CreateObject("adodb.recordset")
            TBCiclo.Open "Select * from item_aplicacoes where codproduto = " & TBItem!Codproduto & " and id_Cliente_Forn = " & txtIDcliente.Text & " order by n_referencia", Conexao, adOpenKeyset, adLockOptimistic
            If TBCiclo.EOF = True Then
                Set TBCiclo = CreateObject("adodb.recordset")
                TBCiclo.Open "Select * from item_aplicacoes where codproduto = " & TBItem!Codproduto & " order by n_referencia", Conexao, adOpenKeyset, adLockOptimistic
                If TBCiclo.EOF = False Then
                    Do While TBCiclo.EOF = False
                        cmbReferencia_prod.AddItem IIf(IsNull(TBCiclo!N_referencia), "", TBCiclo!N_referencia)
                        TBCiclo.MoveNext
                    Loop
                End If
            Else
                cmbReferencia_prod.AddItem IIf(IsNull(TBCiclo!N_referencia), "", TBCiclo!N_referencia)
            End If
            TBCiclo.Close
        End If
        txtRevProduto = IIf(IsNull(TBItem!RevDesenho), "", TBItem!RevDesenho)
        If IsNull(TBAbrir!N_referencia) = False And TBAbrir!N_referencia <> "" Then cmbReferencia_prod = TBAbrir!N_referencia
        txtDescricaoProduto = IIf(IsNull(TBItem!Descricao), "", TBItem!Descricao)
        txtFamiliaProduto = IIf(IsNull(TBItem!Classe), "", TBItem!Classe)
        txtunidade = IIf(IsNull(TBItem!Unidade), "", TBItem!Unidade)
    End If
    TBItem.Close
End If
    
txtDescricaoProduto.Text = IIf(IsNull(TBAbrir!Descricao), "", TBAbrir!Descricao)
txtDataemissao.Text = IIf(IsNull(TBAbrir!DtEmissao), "", Format(TBAbrir!DtEmissao, "dd/mm/yy"))
txtDataStatus.Text = IIf(IsNull(TBAbrir!DataRevisao), "", Format(TBAbrir!DataRevisao, "dd/mm/yy"))
If IsNull(TBAbrir!status) = False Then
    If TBAbrir!status = "Revisado" Then
        txtStatus.Visible = True
        cmbStatus.Visible = False
    Else
        cmbStatus = TBAbrir!status
        txtStatus.Visible = False
        cmbStatus.Visible = True
    End If
End If

txtRespPPAP.Text = IIf(IsNull(TBAbrir!Responsavel), "", TBAbrir!Responsavel)
Caption = "Qualidade - PPAP - PSW - (PPAP : " & TBAbrir!PPAP & " - Rev. : " & TBAbrir!revPPAP & ")"
Frame2.Enabled = True
Frame3.Enabled = True
Novo_qualidadePPAP = False
ProcPuxadadosSubmissao

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaLista(Pagina As Integer)
On Error GoTo tratar_erro

If Sql_PPAP_Localizar = "" Then Exit Sub
lblRegistros.Caption = "Nº de registros: 0"
lblPaginas.Caption = "Página: 0 de: 0"
Lista.ListItems.Clear
Set TBLISTA_Compras_PPAP = CreateObject("adodb.recordset")
TBLISTA_Compras_PPAP.Open Sql_PPAP_Localizar, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA_Compras_PPAP.EOF = False Then ProcExibePagina (Pagina)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina(Pagina)
On Error GoTo tratar_erro

Lista.ListItems.Clear
TBLISTA_Compras_PPAP.PageSize = IIf(txtNreg = "", 30, txtNreg)
TBLISTA_Compras_PPAP.AbsolutePage = Pagina
TamanhoPagina = TBLISTA_Compras_PPAP.PageSize
ContadorReg = 1

PBLista.Min = 0
PBLista.Max = FunVerifMaxPBListaPaginacao(TBLISTA_Compras_PPAP.RecordCount - IIf(Pagina > 1, (TBLISTA_Compras_PPAP.PageSize * (Pagina - 1)), 0), TBLISTA_Compras_PPAP.PageSize)
PBLista.Value = 1
Contador = 0
Do While TBLISTA_Compras_PPAP.EOF = False And (ContadorReg <= TamanhoPagina)
    With Lista.ListItems
        .Add , , TBLISTA_Compras_PPAP!idPPAP
        .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA_Compras_PPAP!DtEmissao), "", Format(TBLISTA_Compras_PPAP!DtEmissao, "dd/mm/yy"))
        .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA_Compras_PPAP!Responsavel), "", TBLISTA_Compras_PPAP!Responsavel)
        .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA_Compras_PPAP!PPAP), "", TBLISTA_Compras_PPAP!PPAP)
        .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA_Compras_PPAP!revPPAP), "", TBLISTA_Compras_PPAP!revPPAP)
        .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA_Compras_PPAP!Cliente), "", TBLISTA_Compras_PPAP!Cliente)
        .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA_Compras_PPAP!Codinterno), "", TBLISTA_Compras_PPAP!Codinterno)
        .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA_Compras_PPAP!Descricao), "", TBLISTA_Compras_PPAP!Descricao)
        .Item(.Count).SubItems(8) = IIf(IsNull(TBLISTA_Compras_PPAP!status), "", TBLISTA_Compras_PPAP!status)
        .Item(.Count).SubItems(9) = IIf(IsNull(TBLISTA_Compras_PPAP!DtValidacao) = False, "SIM", "NÃO")
    End With
    TBLISTA_Compras_PPAP.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista.Value = Contador
Loop
lblRegistros.Caption = "Nº de registros: " & TBLISTA_Compras_PPAP.RecordCount
If TBLISTA_Compras_PPAP.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "Página: 1 de: " & TBLISTA_Compras_PPAP.PageCount
ElseIf TBLISTA_Compras_PPAP.AbsolutePage = adPosEOF Then
        lblPaginas.Caption = "Página: " & TBLISTA_Compras_PPAP.PageCount & " de: " & TBLISTA_Compras_PPAP.PageCount
    Else
        lblPaginas.Caption = "Página: " & TBLISTA_Compras_PPAP.AbsolutePage - 1 & " de: " & TBLISTA_Compras_PPAP.PageCount
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAnterior()
On Error GoTo tratar_erro

If txtIDPPAP = 0 Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from QualidadePPAP order by IdPPAP", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.BOF = False Then
    TBLISTA.Find ("idPPAP = " & txtIDPPAP)
    TBLISTA.MovePrevious
    If TBLISTA.BOF = False Then
        txtIDPPAP.Text = TBLISTA!idPPAP
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from QualidadePPAP where idPPAP = " & txtIDPPAP, Conexao, adOpenKeyset, adLockOptimistic
        ProcLimpaCamposPPAP
        ProcPuxaDadosPPAP
    Else
        USMsgBox ("Fim dos cadastros de PPAP."), vbInformation, "CAPRIND v5.0"
    End If
End If
Novo_qualidadePPAP = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcProximo()
On Error GoTo tratar_erro

If txtIDPPAP = 0 Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from QualidadePPAP order by idPPAP", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.BOF = False Then
    TBLISTA.Find ("idPPAP = " & txtIDPPAP)
    TBLISTA.MoveNext
    If TBLISTA.EOF = False Then
        txtIDPPAP.Text = TBLISTA!idPPAP
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from QualidadePPAP where idPPAP = " & txtIDPPAP, Conexao, adOpenKeyset, adLockOptimistic
        ProcLimpaCamposPPAP
        ProcPuxaDadosPPAP
    Else
        USMsgBox ("Fim dos cadastros de PPAP."), vbInformation, "CAPRIND v5.0"
    End If
End If
Novo_qualidadePPAP = False

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

Private Sub txtpeso_Change()
On Error GoTo tratar_erro

If txtpeso.Text <> "" Then
    VerifNumero = txtpeso.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtpeso.Text = ""
        txtpeso.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviaDadosRevisar()
On Error GoTo tratar_erro

TBGravar!IDProduto = TBAbrir!IDProduto
TBGravar!Codinterno = TBAbrir!Codinterno
TBGravar!N_referencia = TBAbrir!N_referencia
TBGravar!Descricao = TBAbrir!Descricao
TBGravar!Revproduto = TBAbrir!Revproduto
TBGravar!IDCliente = TBAbrir!IDCliente
TBGravar!Cliente = TBAbrir!Cliente
TBGravar!DtEmissao = TBAbrir!DtEmissao
TBGravar!Responsavel = pubUsuario
TBGravar!status = "Aberto"
TBGravar!Nivel = TBAbrir!Nivel
TBGravar!Data = TBAbrir!Data
TBGravar!Alteracoes = TBAbrir!Alteracoes
TBGravar!Data2 = TBAbrir!Data2
TBGravar!Auxilio = TBAbrir!Auxilio
TBGravar!Nivel2 = TBAbrir!Nivel2
TBGravar!Data3 = TBAbrir!Data3
TBGravar!contato = TBAbrir!contato
TBGravar!Pedido = TBAbrir!Pedido
TBGravar!Peso = TBAbrir!Peso
TBGravar!Sim = TBAbrir!Sim
TBGravar!NAO = TBAbrir!NAO
TBGravar!Aplicacao = TBAbrir!Aplicacao
TBGravar!Sim2 = TBAbrir!Sim2
TBGravar!Nao2 = TBAbrir!Nao2
TBGravar!NA = TBAbrir!NA
TBGravar!chkIMDS = TBAbrir!chkIMDS
TBGravar!IMDS = TBAbrir!IMDS
TBGravar!Sim3 = TBAbrir!Sim3
TBGravar!Nao3 = TBAbrir!Nao3
TBGravar!na2 = TBAbrir!na2
TBGravar!Sim4 = TBAbrir!Sim4
TBGravar!Nao4 = TBAbrir!Nao4
TBGravar!na3 = TBAbrir!na3
TBGravar!Email = TBAbrir!Email
TBGravar!CodFornecedor = TBAbrir!CodFornecedor
TBGravar!Texto_Declaracao = TBAbrir!Texto_Declaracao
TBGravar!OBS_Declaracao = TBAbrir!OBS_Declaracao
TBGravar!opt1_Nivel = TBAbrir!opt1_Nivel
TBGravar!opt2_Nivel = TBAbrir!opt2_Nivel
TBGravar!opt3_Nivel = TBAbrir!opt3_Nivel
TBGravar!opt4_Nivel = TBAbrir!opt4_Nivel
TBGravar!opt5_Nivel = TBAbrir!opt5_Nivel
TBGravar!opt1_Razao = TBAbrir!opt1_Razao
TBGravar!opt2_Razao = TBAbrir!opt2_Razao
TBGravar!opt3_Razao = TBAbrir!opt3_Razao
TBGravar!opt4_Razao = TBAbrir!opt4_Razao
TBGravar!opt5_Razao = TBAbrir!opt5_Razao
TBGravar!opt6_Razao = TBAbrir!opt6_Razao
TBGravar!opt7_Razao = TBAbrir!opt7_Razao
TBGravar!opt8_Razao = TBAbrir!opt8_Razao
TBGravar!opt9_Razao = TBAbrir!opt9_Razao
TBGravar!opt10_Razao = TBAbrir!opt10_Razao
TBGravar!Outras_Razao = TBAbrir!Outras_Razao
TBGravar!Chk1_Resultados = TBAbrir!Chk1_Resultados
TBGravar!Chk2_Resultados = TBAbrir!Chk2_Resultados
TBGravar!Chk3_Resultados = TBAbrir!Chk3_Resultados
TBGravar!Chk4_Resultados = TBAbrir!Chk4_Resultados
TBGravar!Sim_Resultados = TBAbrir!Sim_Resultados
TBGravar!Nao_Resultados = TBAbrir!Nao_Resultados
TBGravar!Obs_Resultados = TBAbrir!Obs_Resultados

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcGeraNumero()
On Error GoTo tratar_erro

a = Cont
Select Case Len(a)
    Case 1: a = "PPAP-" & "000" & Cont & "/" & Data_PPAP
    Case 2: a = "PPAP-" & "00" & Cont & "/" & Data_PPAP
    Case 3: a = "PPAP-" & "0" & Cont & "/" & Data_PPAP
    Case 4: a = "PPAP-" & Cont & "/" & Data_PPAP
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcImprimir()
On Error GoTo tratar_erro

If txtIDPPAP = 0 Then
    USMsgBox ("Informe o PPAP antes de visualizar impressão."), vbExclamation, "CAPRIND v5.0"
    ProcLocalizar
    Exit Sub
End If
NomeRel = "CQ_PPAP_PSW.rpt"
ProcImprimirRel "{QualidadePPAP.IDPPAP} = " & txtIDPPAP, ""

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
    Case 6: ProcAnterior
    Case 7: ProcProximo
    Case 8: ProcCopiar
    Case 9: ProcRevisao
    Case 10: ProcValidarRegistros Lista, "Qualidade/PPAP/PSW"
    'Case 12: ProcAjuda
    Case 13: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
