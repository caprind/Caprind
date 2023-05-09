VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.Controls.v19.3.0.ocx"
Begin VB.Form frmFaturamento_Prod_serv_boleto 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Administrativo - Emissor de boleto bancário"
   ClientHeight    =   10035
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   15360
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFaturamento_Prod_serv_boleto.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   10035
   ScaleWidth      =   15360
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
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
      Height          =   825
      Left            =   30
      TabIndex        =   36
      Top             =   7290
      Width           =   10305
      Begin VB.TextBox Txt_data_envio 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   8430
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         ToolTipText     =   "Data do envio."
         Top             =   345
         Width           =   1635
      End
      Begin VB.CheckBox Chk_email_enviado 
         BackColor       =   &H00E0E0E0&
         Caption         =   "E-mail enviado"
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
         Height          =   195
         Left            =   150
         TabIndex        =   19
         Top             =   30
         Width           =   1545
      End
      Begin VB.TextBox Txt_assunto 
         Alignment       =   2  'Center
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
         Height          =   315
         Left            =   180
         TabIndex        =   18
         ToolTipText     =   "Assunto."
         Top             =   345
         Width           =   8235
      End
      Begin VB.Label Label34 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Dt. do envio"
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
         Left            =   8790
         TabIndex        =   41
         Top             =   150
         Width           =   915
      End
      Begin VB.Label Label32 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Assunto"
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
         Left            =   3990
         TabIndex        =   37
         Top             =   150
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Informações do boleto bancário"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2445
      Left            =   30
      TabIndex        =   21
      Top             =   2490
      Width           =   10305
      Begin VB.ComboBox Cmb_especie_doc 
         BackColor       =   &H00E0E0E0&
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
         ItemData        =   "frmFaturamento_Prod_serv_boleto.frx":000C
         Left            =   3270
         List            =   "frmFaturamento_Prod_serv_boleto.frx":003A
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   7
         ToolTipText     =   "Espécie do documento."
         Top             =   1230
         Width           =   1275
      End
      Begin VB.ComboBox Cmb_carteira1 
         BackColor       =   &H00E0E0E0&
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
         ItemData        =   "frmFaturamento_Prod_serv_boleto.frx":0077
         Left            =   9450
         List            =   "frmFaturamento_Prod_serv_boleto.frx":0079
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   11
         ToolTipText     =   "Carteira."
         Top             =   1230
         Width           =   645
      End
      Begin VB.TextBox Txt_valor_doc 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   8460
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Valor do documento."
         Top             =   1905
         Width           =   1635
      End
      Begin VB.TextBox Txt_valor 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   7170
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Valor."
         Top             =   1905
         Width           =   1275
      End
      Begin VB.TextBox Txt_quantidade_moeda 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   5670
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade moeda."
         Top             =   1905
         Width           =   1485
      End
      Begin VB.TextBox Txt_especie_moeda 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   4530
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   14
         TabStop         =   0   'False
         Text            =   "R$"
         ToolTipText     =   "Espécie da moeda."
         Top             =   1905
         Width           =   1125
      End
      Begin VB.ComboBox Cmb_carteira 
         BackColor       =   &H00E0E0E0&
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
         ItemData        =   "frmFaturamento_Prod_serv_boleto.frx":007B
         Left            =   6810
         List            =   "frmFaturamento_Prod_serv_boleto.frx":007D
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   10
         ToolTipText     =   "Carteira."
         Top             =   1230
         Width           =   3255
      End
      Begin VB.TextBox Txt_uso_banco 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   3270
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Uso do banco."
         Top             =   1905
         Width           =   1245
      End
      Begin VB.TextBox Txt_nosso_numero 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Nosso número/código do documento."
         Top             =   1905
         Width           =   3075
      End
      Begin VB.TextBox Txt_data_processamento 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Data de processamento."
         Top             =   1230
         Width           =   1545
      End
      Begin VB.TextBox Txt_aceite 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   4560
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   "N"
         ToolTipText     =   "Aceite."
         Top             =   1230
         Width           =   705
      End
      Begin VB.TextBox Txt_numero_doc 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         TabIndex        =   6
         ToolTipText     =   "Número do documento."
         Top             =   1230
         Width           =   1785
      End
      Begin VB.TextBox Txt_data_doc 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Data do documento."
         Top             =   1230
         Width           =   1275
      End
      Begin VB.TextBox Txt_vencimento 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   8550
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Vencimento."
         Top             =   480
         Width           =   1545
      End
      Begin VB.TextBox Txt_local_de_pagamento 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Text            =   "Até o vencimento pagável em qualquer banco do sistema de compensação"
         ToolTipText     =   "Local de pagamento."
         Top             =   480
         Width           =   8355
      End
      Begin MSComCtl2.DTPicker Cmb_vencimento 
         Height          =   315
         Left            =   8550
         TabIndex        =   4
         ToolTipText     =   "Vencimento."
         Top             =   480
         Visible         =   0   'False
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
         Format          =   172818435
         CurrentDate     =   39057
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(=) Valor do doc."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   8655
         TabIndex        =   35
         Top             =   1710
         Width           =   1245
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   7620
         TabIndex        =   34
         Top             =   1710
         Width           =   375
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Quantidade moeda"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5715
         TabIndex        =   33
         Top             =   1710
         Width           =   1395
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Espécie moeda"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4530
         TabIndex        =   32
         Top             =   1710
         Width           =   1095
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Carteira"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   8160
         TabIndex        =   31
         Top             =   1050
         Width           =   615
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Uso do banco"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3360
         TabIndex        =   30
         Top             =   1710
         Width           =   1005
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Nosso número/Cód. doc."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   795
         TabIndex        =   29
         Top             =   1710
         Width           =   1815
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Dt. processamento"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5355
         TabIndex        =   28
         Top             =   1050
         Width           =   1395
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Aceite"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4680
         TabIndex        =   27
         Top             =   1050
         Width           =   465
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Espécie doc."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3480
         TabIndex        =   26
         Top             =   1050
         Width           =   915
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Número do documento"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1545
         TabIndex        =   25
         Top             =   1050
         Width           =   1635
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Data documento"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   210
         TabIndex        =   24
         Top             =   1050
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Vencimento"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   8895
         TabIndex        =   23
         Top             =   270
         Width           =   855
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Local para pagamento"
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
         Left            =   3750
         TabIndex        =   22
         Top             =   270
         Width           =   1605
      End
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Juros | Multa | Desconto"
      DragMode        =   1  'Automatic
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
      Height          =   5625
      Left            =   10350
      TabIndex        =   62
      Top             =   2490
      Width           =   4965
      Begin VB.TextBox Txt_outras_deducoes 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   3150
         Locked          =   -1  'True
         TabIndex        =   76
         TabStop         =   0   'False
         ToolTipText     =   "Outras deduções."
         Top             =   3501
         Width           =   1635
      End
      Begin VB.TextBox Txt_mora 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   3150
         Locked          =   -1  'True
         TabIndex        =   75
         TabStop         =   0   'False
         ToolTipText     =   "Mora/multa/juros."
         Top             =   3982
         Width           =   1635
      End
      Begin VB.TextBox Txt_outros_acrescimos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   3150
         TabIndex        =   74
         ToolTipText     =   "Outros acréscimos."
         Top             =   4463
         Width           =   1635
      End
      Begin VB.TextBox Txt_valor_cobrado 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   3150
         Locked          =   -1  'True
         TabIndex        =   73
         TabStop         =   0   'False
         ToolTipText     =   "Valor cobrado."
         Top             =   4950
         Width           =   1635
      End
      Begin VB.TextBox Txt_desconto 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   360
         Left            =   3150
         TabIndex        =   71
         Top             =   2464
         Width           =   1635
      End
      Begin VB.TextBox Txt_percentual_desconto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   3150
         MaxLength       =   30
         TabIndex        =   66
         ToolTipText     =   "Percentual de desconto."
         Top             =   1021
         Width           =   1635
      End
      Begin VB.TextBox Txt_percentual_multa 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   3150
         MaxLength       =   30
         TabIndex        =   65
         ToolTipText     =   "Percentual de multa."
         Top             =   1502
         Width           =   1635
      End
      Begin VB.TextBox Txt_percentual_juros 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   3150
         MaxLength       =   30
         TabIndex        =   64
         ToolTipText     =   "Percentual de juros diário."
         Top             =   540
         Width           =   1635
      End
      Begin VB.TextBox Txt_dias_protesto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   3150
         MaxLength       =   30
         TabIndex        =   63
         ToolTipText     =   "Dias protesto."
         Top             =   1983
         Width           =   1635
      End
      Begin MSComCtl2.DTPicker txtDatalimiteDesc 
         Height          =   345
         Left            =   3150
         TabIndex        =   72
         ToolTipText     =   "Data limite concessão desconto"
         Top             =   2990
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   609
         _Version        =   393216
         CalendarBackColor=   14737632
         CalendarTitleBackColor=   14737632
         Format          =   172359681
         CurrentDate     =   44319
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(-) Desconto | Abatim."
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
         Left            =   1410
         TabIndex        =   82
         Top             =   2040
         Width           =   1635
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(-) Outras deduções"
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
         Left            =   1590
         TabIndex        =   81
         Top             =   3495
         Width           =   1455
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(+) Mora|Multa|Juros"
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
         Left            =   1500
         TabIndex        =   80
         Top             =   3990
         Width           =   1545
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(+) Outros acréscimos"
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
         TabIndex        =   79
         Top             =   4470
         Width           =   1605
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(=) Valor cobrado"
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
         Left            =   1770
         TabIndex        =   78
         Top             =   4965
         Width           =   1275
      End
      Begin VB.Label Label35 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Data limite desconto"
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
         Left            =   1590
         TabIndex        =   77
         Top             =   3015
         Width           =   1455
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Juros diário (%)"
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
         Left            =   1890
         TabIndex        =   70
         Top             =   585
         Width           =   1155
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Multa (%)"
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
         Left            =   2325
         TabIndex        =   69
         Top             =   1560
         Width           =   720
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Desconto (%)"
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
         TabIndex        =   68
         Top             =   1065
         Width           =   1005
      End
      Begin VB.Label Label30 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Dias p/ protesto"
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
         Left            =   1890
         TabIndex        =   67
         Top             =   2535
         Width           =   1155
      End
   End
   Begin VB.Frame FrameAtualizacao 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Dados para vencimento | Juros | Multa | Atraso"
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
      ForeColor       =   &H00800000&
      Height          =   915
      Left            =   4860
      TabIndex        =   38
      Top             =   8100
      Width           =   5475
      Begin VB.TextBox Txt_dias_atraso 
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Text            =   "0"
         ToolTipText     =   "Dias em atraso."
         Top             =   510
         Width           =   1125
      End
      Begin MSComCtl2.DTPicker Cmb_novo_vencimento 
         Height          =   315
         Left            =   270
         TabIndex        =   0
         ToolTipText     =   "Novo vencimento."
         Top             =   510
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
         Format          =   172359683
         CurrentDate     =   39057
      End
      Begin DrawSuite2022.USCheckBox Chk_calcular_juros_multa 
         Height          =   255
         Left            =   3360
         TabIndex        =   46
         Top             =   540
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   450
         Caption         =   "Calcular juros e multa?"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   128
         ShowFocusRect   =   0   'False
      End
      Begin VB.Label Label33 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Atraso (dias)"
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
         Left            =   1935
         TabIndex        =   40
         Top             =   300
         Width           =   930
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Novo vencimento"
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
         TabIndex        =   39
         Top             =   300
         Width           =   1275
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Opções para atualizar boleto vencido"
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
      Height          =   915
      Left            =   30
      TabIndex        =   43
      Top             =   8100
      Width           =   4815
      Begin DrawSuite2022.USCheckBox Chk_novo 
         Height          =   255
         Left            =   330
         TabIndex        =   44
         Top             =   330
         Width           =   3315
         _ExtentX        =   5847
         _ExtentY        =   450
         Caption         =   "Novo vencimento sem juros"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocusRect   =   0   'False
      End
      Begin DrawSuite2022.USCheckBox Chk_atualizar 
         Height          =   255
         Left            =   330
         TabIndex        =   45
         Top             =   600
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   450
         Caption         =   "Atualizar boleto com juros e multa"
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocusRect   =   0   'False
      End
   End
   Begin VB.Frame Frame16 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Status retorno"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   30
      TabIndex        =   104
      Top             =   9000
      Width           =   10305
      Begin XtremeSuiteControls.FlatEdit txtRetorno 
         Height          =   675
         Left            =   90
         TabIndex        =   105
         ToolTipText     =   "Mensagem de retorno da API banco"
         Top             =   210
         Width           =   9375
         _Version        =   1245187
         _ExtentX        =   16536
         _ExtentY        =   1191
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   12640511
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   12640511
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2
         Appearance      =   12
         UseVisualStyle  =   0   'False
      End
      Begin DrawSuite2022.USButton btnRetorno 
         Height          =   675
         Left            =   9480
         TabIndex        =   106
         TabStop         =   0   'False
         ToolTipText     =   "Consultar boleto"
         Top             =   210
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   1191
         DibPicture      =   "frmFaturamento_Prod_serv_boleto.frx":007F
         Caption         =   "Consultar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   8421504
         BorderColorDown =   7907521
         BorderColorOver =   7907521
         ForeColor       =   0
         ForeColorOver   =   0
         ForeColorDown   =   0
         GradientColor1  =   16777215
         GradientColor2  =   14737632
         GradientColor3  =   12632256
         GradientColor4  =   12632256
         GradientColorOver1=   14417407
         GradientColorOver2=   12317439
         GradientColorOver3=   4838399
         GradientColorOver4=   9627391
         GradientColorDown1=   10802943
         GradientColorDown2=   7979263
         GradientColorDown3=   4370174
         GradientColorDown4=   7395582
         PicAlign        =   8
         PicSize         =   3
         PicSizeH        =   32
         PicSizeW        =   32
         ShowFocusRect   =   0   'False
         Theme           =   1
      End
   End
   Begin VB.Frame Frame15 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   30
      TabIndex        =   101
      Top             =   -30
      Width           =   15285
      Begin DrawSuite2022.USButton btnEmitir 
         Height          =   735
         Left            =   90
         TabIndex        =   102
         ToolTipText     =   "Emitir boleto"
         Top             =   210
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   1296
         DibPicture      =   "frmFaturamento_Prod_serv_boleto.frx":18D3
         Caption         =   "Emitir boleto"
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
         BorderColorDown =   7907521
         BorderColorOver =   7907521
         ForeColor       =   0
         ForeColorOver   =   0
         ForeColorDown   =   0
         GradientColor1  =   16777215
         GradientColor2  =   14737632
         GradientColor3  =   12632256
         GradientColor4  =   12632256
         GradientColorOver1=   14417407
         GradientColorOver2=   12317439
         GradientColorOver3=   4838399
         GradientColorOver4=   9627391
         GradientColorDown1=   10802943
         GradientColorDown2=   7979263
         GradientColorDown3=   4370174
         GradientColorDown4=   7395582
         PicAlign        =   7
         PicSize         =   2
         PicSizeH        =   24
         PicSizeW        =   24
         ShowFocusRect   =   0   'False
         Theme           =   1
      End
      Begin DrawSuite2022.USButton btnLimpar 
         Height          =   735
         Left            =   8010
         TabIndex        =   103
         ToolTipText     =   "Cancelar boleto"
         Top             =   210
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   1296
         DibPicture      =   "frmFaturamento_Prod_serv_boleto.frx":8A53
         Caption         =   "Excluir boleto"
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
         BorderColorDown =   7907521
         BorderColorOver =   7907521
         ForeColor       =   0
         ForeColorOver   =   0
         ForeColorDown   =   0
         GradientColor1  =   16777215
         GradientColor2  =   14737632
         GradientColor3  =   12632256
         GradientColor4  =   12632256
         GradientColorOver1=   14417407
         GradientColorOver2=   12317439
         GradientColorOver3=   4838399
         GradientColorOver4=   9627391
         GradientColorDown1=   10802943
         GradientColorDown2=   7979263
         GradientColorDown3=   4370174
         GradientColorDown4=   7395582
         PicAlign        =   7
         PicSize         =   2
         PicSizeH        =   24
         PicSizeW        =   24
         ShowFocusRect   =   0   'False
         Theme           =   1
      End
      Begin DrawSuite2022.USButton btnPdf 
         Height          =   735
         Left            =   2070
         TabIndex        =   107
         ToolTipText     =   "Gerar Pdf do boleto"
         Top             =   210
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   1296
         DibPicture      =   "frmFaturamento_Prod_serv_boleto.frx":EB1B
         Caption         =   "Gerar PDF"
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
         BorderColorDown =   7907521
         BorderColorOver =   7907521
         ForeColor       =   0
         ForeColorOver   =   0
         ForeColorDown   =   0
         GradientColor1  =   16777215
         GradientColor2  =   14737632
         GradientColor3  =   12632256
         GradientColor4  =   12632256
         GradientColorOver1=   14417407
         GradientColorOver2=   12317439
         GradientColorOver3=   4838399
         GradientColorOver4=   9627391
         GradientColorDown1=   10802943
         GradientColorDown2=   7979263
         GradientColorDown3=   4370174
         GradientColorDown4=   7395582
         PicAlign        =   7
         PicSize         =   2
         PicSizeH        =   24
         PicSizeW        =   24
         ShowFocusRect   =   0   'False
         Theme           =   1
      End
      Begin DrawSuite2022.USButton btnRemessa 
         Height          =   735
         Left            =   4050
         TabIndex        =   108
         ToolTipText     =   "Gerar arquivo remessa do boleto"
         Top             =   210
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   1296
         DibPicture      =   "frmFaturamento_Prod_serv_boleto.frx":27E03
         Caption         =   "Gerar remessa"
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
         BorderColorDown =   7907521
         BorderColorOver =   7907521
         ForeColor       =   0
         ForeColorOver   =   0
         ForeColorDown   =   0
         GradientColor1  =   16777215
         GradientColor2  =   14737632
         GradientColor3  =   12632256
         GradientColor4  =   12632256
         GradientColorOver1=   14417407
         GradientColorOver2=   12317439
         GradientColorOver3=   4838399
         GradientColorOver4=   9627391
         GradientColorDown1=   10802943
         GradientColorDown2=   7979263
         GradientColorDown3=   4370174
         GradientColorDown4=   7395582
         PicAlign        =   7
         PicSize         =   2
         PicSizeH        =   24
         PicSizeW        =   24
         ShowFocusRect   =   0   'False
         Theme           =   1
      End
      Begin DrawSuite2022.USButton btnEmail 
         Height          =   735
         Left            =   6030
         TabIndex        =   109
         ToolTipText     =   "Enviar boleto por emailo para o cliente."
         Top             =   210
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   1296
         DibPicture      =   "frmFaturamento_Prod_serv_boleto.frx":2C468
         Caption         =   "Enviar por email"
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
         BorderColorDown =   7907521
         BorderColorOver =   7907521
         ForeColor       =   0
         ForeColorOver   =   0
         ForeColorDown   =   0
         GradientColor1  =   16777215
         GradientColor2  =   14737632
         GradientColor3  =   12632256
         GradientColor4  =   12632256
         GradientColorOver1=   14417407
         GradientColorOver2=   12317439
         GradientColorOver3=   4838399
         GradientColorOver4=   9627391
         GradientColorDown1=   10802943
         GradientColorDown2=   7979263
         GradientColorDown3=   4370174
         GradientColorDown4=   7395582
         PicAlign        =   7
         PicSize         =   2
         PicSizeH        =   24
         PicSizeW        =   24
         ShowFocusRect   =   0   'False
         Theme           =   1
      End
      Begin DrawSuite2022.USButton btnSair 
         Height          =   735
         Left            =   14040
         TabIndex        =   110
         ToolTipText     =   "Fechar formulário"
         Top             =   210
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   1296
         DibPicture      =   "frmFaturamento_Prod_serv_boleto.frx":2E289
         Caption         =   "Sair"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   5263559
         BorderColorDown =   2039646
         BorderColorOver =   3026574
         GradientColor1  =   5263559
         GradientColor2  =   5263559
         GradientColor3  =   5263559
         GradientColor4  =   5263559
         GradientColorOver1=   3026574
         GradientColorOver2=   3026574
         GradientColorOver3=   3026574
         GradientColorOver4=   3026574
         GradientColorDown1=   2039646
         GradientColorDown2=   2039646
         GradientColorDown3=   2039646
         GradientColorDown4=   2039646
         PicAlign        =   7
         PicSize         =   2
         PicSizeH        =   24
         PicSizeW        =   24
         ShowFocusRect   =   0   'False
         Theme           =   4
      End
   End
   Begin VB.Frame Frame14 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Instruções para o banco"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1905
      Left            =   10350
      TabIndex        =   98
      Top             =   8100
      Width           =   4965
      Begin VB.TextBox Txt_instrucoes 
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
         Height          =   1545
         Left            =   150
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   99
         ToolTipText     =   "Instruções."
         Top             =   240
         Width           =   4125
      End
      Begin DrawSuite2022.USButton Cmd_instrucoes 
         Height          =   1545
         Left            =   4290
         TabIndex        =   100
         ToolTipText     =   "Buscar item cadastrado ou vendido"
         Top             =   240
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   2725
         DibPicture      =   "frmFaturamento_Prod_serv_boleto.frx":5FA72
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
         BorderColor     =   8421504
         BorderColorDown =   7907521
         BorderColorOver =   7907521
         ForeColor       =   0
         ForeColorOver   =   0
         ForeColorDown   =   0
         GradientColor1  =   16777215
         GradientColor2  =   14737632
         GradientColor3  =   12632256
         GradientColor4  =   12632256
         GradientColorOver1=   14417407
         GradientColorOver2=   12317439
         GradientColorOver3=   4838399
         GradientColorOver4=   9627391
         GradientColorDown1=   10802943
         GradientColorDown2=   7979263
         GradientColorDown3=   4370174
         GradientColorDown4=   7395582
         PicAlign        =   0
         ShowFocusRect   =   0   'False
         Theme           =   1
         ToolTipTitle    =   "CAPRIND v5.0"
      End
   End
   Begin VB.Frame Frame13 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Informações do sacado"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1515
      Left            =   30
      TabIndex        =   84
      Top             =   5790
      Width           =   10305
      Begin VB.TextBox Txt_sacado 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Locked          =   -1  'True
         TabIndex        =   87
         TabStop         =   0   'False
         ToolTipText     =   "Sacado."
         Top             =   450
         Width           =   8235
      End
      Begin VB.TextBox txt_CNPJ 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   8430
         Locked          =   -1  'True
         TabIndex        =   86
         TabStop         =   0   'False
         ToolTipText     =   "CNPJ."
         Top             =   450
         Width           =   1635
      End
      Begin VB.ComboBox Cmb_endereco 
         BackColor       =   &H00E0E0E0&
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
         ItemData        =   "frmFaturamento_Prod_serv_boleto.frx":7DB77
         Left            =   180
         List            =   "frmFaturamento_Prod_serv_boleto.frx":7DB79
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   85
         ToolTipText     =   "Endereço de cobrança."
         Top             =   1050
         Width           =   9915
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Endereço de cobrança"
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
         Left            =   4290
         TabIndex        =   90
         Top             =   840
         Width           =   1635
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Sacado"
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
         Left            =   4005
         TabIndex        =   89
         Top             =   240
         Width           =   555
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CNPJ"
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
         Left            =   9060
         TabIndex        =   88
         Top             =   240
         Width           =   405
      End
   End
   Begin VB.Frame Frame12 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Informações do cedente"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   30
      TabIndex        =   83
      Top             =   4920
      Width           =   10305
      Begin VB.TextBox Txt_codigo_cedente 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Locked          =   -1  'True
         TabIndex        =   93
         TabStop         =   0   'False
         ToolTipText     =   "Código do cedente/convênio."
         Top             =   450
         Width           =   1305
      End
      Begin VB.TextBox Txt_cedente 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   1500
         Locked          =   -1  'True
         TabIndex        =   92
         TabStop         =   0   'False
         ToolTipText     =   "Cedente."
         Top             =   450
         Width           =   6945
      End
      Begin VB.TextBox Txt_agencia_codigo_cedente 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   8460
         Locked          =   -1  'True
         TabIndex        =   91
         TabStop         =   0   'False
         ToolTipText     =   "Agência/código do cedente."
         Top             =   450
         Width           =   1605
      End
      Begin VB.TextBox Txt_IDempresa 
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
         Left            =   2070
         Locked          =   -1  'True
         TabIndex        =   97
         TabStop         =   0   'False
         Top             =   450
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Cedente"
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
         Left            =   4530
         TabIndex        =   96
         Top             =   255
         Width           =   645
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Agencia|Cód. cedente"
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
         Left            =   8430
         TabIndex        =   95
         Top             =   225
         Width           =   1635
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Código cedente"
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
         TabIndex        =   94
         Top             =   240
         Width           =   1155
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Emissor boleto"
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
      Left            =   11070
      TabIndex        =   56
      Top             =   990
      Width           =   4245
      Begin VB.TextBox txtEmissor 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   57
         TabStop         =   0   'False
         ToolTipText     =   "Data do envio."
         Top             =   270
         Width           =   4005
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Status boleto"
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
      Left            =   8250
      TabIndex        =   54
      Top             =   990
      Width           =   2805
      Begin VB.TextBox txtStatus 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         ForeColor       =   &H00000040&
         Height          =   360
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   270
         Width           =   2550
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Protocolo"
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
      Left            =   5430
      TabIndex        =   52
      Top             =   990
      Width           =   2805
      Begin VB.TextBox txtProtocolo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   360
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   270
         Width           =   2550
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ID Integração"
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
      Left            =   2610
      TabIndex        =   50
      Top             =   990
      Width           =   2805
      Begin VB.TextBox txtIDIntegracao 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   360
         Left            =   150
         Locked          =   -1  'True
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   270
         Width           =   2520
      End
   End
   Begin VB.TextBox TxtHTLM 
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
      Height          =   1515
      Left            =   225
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   42
      TabStop         =   0   'False
      Text            =   "frmFaturamento_Prod_serv_boleto.frx":7DB7B
      ToolTipText     =   "HTML para boleto personalizado"
      Top             =   11610
      Width           =   10215
   End
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   99
      ScreenHeight    =   768
      ScreenWidth     =   1366
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
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Local para armazenamento do arquivo remessa"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   2610
      TabIndex        =   47
      Top             =   1770
      Width           =   6405
      Begin VB.TextBox TxtlocalArmazenamento 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   360
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   240
         Width           =   5670
      End
      Begin DrawSuite2022.USButton cmdLocal 
         Height          =   375
         Left            =   5805
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   240
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   661
         DibPicture      =   "frmFaturamento_Prod_serv_boleto.frx":7EBE2
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   8421504
         BorderColorDown =   7907521
         BorderColorOver =   7907521
         ForeColor       =   0
         ForeColorOver   =   0
         ForeColorDown   =   0
         GradientColor1  =   16777215
         GradientColor2  =   14737632
         GradientColor3  =   12632256
         GradientColor4  =   12632256
         GradientColorOver1=   14417407
         GradientColorOver2=   12317439
         GradientColorOver3=   4838399
         GradientColorOver4=   9627391
         GradientColorDown1=   10802943
         GradientColorDown2=   7979263
         GradientColorDown3=   4370174
         GradientColorDown4=   7395582
         PicAlign        =   2
         ShowFocusRect   =   0   'False
         Theme           =   1
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Local para armazenamento do boleto PDF"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   9030
      TabIndex        =   58
      Top             =   1770
      Width           =   6285
      Begin VB.TextBox txtDiretorioBoleto 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   360
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   240
         Width           =   5580
      End
      Begin DrawSuite2022.USButton btnDiretorioBoleto 
         Height          =   375
         Left            =   5715
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   240
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   661
         DibPicture      =   "frmFaturamento_Prod_serv_boleto.frx":8325D
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   8421504
         BorderColorDown =   7907521
         BorderColorOver =   7907521
         ForeColor       =   0
         ForeColorOver   =   0
         ForeColorDown   =   0
         GradientColor1  =   16777215
         GradientColor2  =   14737632
         GradientColor3  =   12632256
         GradientColor4  =   12632256
         GradientColorOver1=   14417407
         GradientColorOver2=   12317439
         GradientColorOver3=   4838399
         GradientColorOver4=   9627391
         GradientColorDown1=   10802943
         GradientColorDown2=   7979263
         GradientColorDown3=   4370174
         GradientColorDown4=   7395582
         ShowFocusRect   =   0   'False
         Theme           =   1
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Banco"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1485
      Left            =   30
      TabIndex        =   61
      Top             =   990
      Width           =   2565
      Begin VB.Image Img_logo_banco 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Left            =   120
         Stretch         =   -1  'True
         ToolTipText     =   "Imagem."
         Top             =   330
         Width           =   2265
      End
   End
End
Attribute VB_Name = "frmFaturamento_Prod_serv_boleto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Boleto da Tecnospeed
Public FBoletoX As spdBoletoX

Dim EnderecoBoleto As String 'OK
Dim BairroBoleto As String 'OK
Dim CidadeBoleto As String 'OK
Dim EstadoBoleto As String 'OK
Dim CEPBoleto As String 'OK
Dim VarMulta As String
Dim VarDesconto As String
Dim VarJuros As String
Dim VarProtesto As String
Dim VarInstrucoes As String

Private Sub btnDiretorioBoleto_Click()
On Error GoTo tratar_erro

    DS.OpenFolderWithExplorer txtDiretorioBoleto.Text

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub btnEmail_Click()
On Error GoTo tratar_erro


If ProcVerifCampos(True, False) = False Then
    Exit Sub
End If

ProcEnviarEmail

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub btnEmitir_Click()
On Error GoTo tratar_erro

If Emissor = "Tecnospeed" And txtStatus = "" Then
    Sit_REG = 0
    ProcStatusBoleto
End If
    

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub btnLimpar_Click()
On Error GoTo tratar_erro

If txtIDIntegracao <> "" Then
    If USMsgBox("Deseja realmente excluir a emissão desse boleto?", vbYesNo, "CAPRIND v5.0") = vbYes Then
        PlugDescartarBoleto (txtIDIntegracao.Text)
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub btnPDF_Click()
On Error GoTo tratar_erro

 ProcImprimir

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub btnRemessa_Click()
On Error GoTo tratar_erro

 ProcRemessa

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub btnRetorno_Click()
On Error GoTo tratar_erro

If txtIDIntegracao <> "" Then
txtRetorno.Text = ""
    PlugConsultarBoleto (txtIDIntegracao.Text)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub btnSair_Click()
On Error GoTo tratar_erro

Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_atualizar_GotFocus()
On Error GoTo tratar_erro

If Financeiro_Contas_Receber = True And (Sit_REG = 1 Or Sit_REG = 3) Then Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_calcular_juros_multa_Click()
On Error GoTo tratar_erro

ProcCalculaJurosMulta IIf(Txt_dias_atraso = "", 0, Txt_dias_atraso)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_novo_GotFocus()
On Error GoTo tratar_erro

If Financeiro_Contas_Receber = True And (Sit_REG = 1 Or Sit_REG = 3) Then Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_carteira_Click()
On Error GoTo tratar_erro

Txt_local_de_pagamento = "Até o vencimento pagável em qualquer banco do sistema de compensação"

'ProcGeraNossoNumero

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcGeraNossoNumero()
On Error GoTo tratar_erro

'Verif. último nosso numero
If Cmb_carteira = "" Then Exit Sub

With Txt_nosso_numero
    .Text = 1
    .Locked = False
    .TabStop = True
    Texto = .Text
End With

If Financeiro_Contas_Receber = False Then
    TextoFiltro = "txt_Agencia = '" & frmFaturamento_Prod_Serv.txt_Agencia & "' and txt_Conta = '" & frmFaturamento_Prod_Serv.txt_Conta & "'"
Else
    TextoFiltro = "txt_Portador_Banco = '" & frmContas_Receber.cmbBanco & "' and txt_Agencia = '" & Agencia & "' and txt_Conta = '" & ContaCorrente & "'"
End If

StrSql = "Select max(Nosso_numero) as Nosso_numero from tbl_Detalhes_Recebimento where " & TextoFiltro & " and Nosso_numero is not null "

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    If IsNull(TBAbrir!Nosso_Numero) = False And TBAbrir!Nosso_Numero <> "" Then
        Texto = TBAbrir!Nosso_Numero + 1
        With Txt_nosso_numero
            .Locked = True
            .TabStop = False
        End With
    Else
        Texto = 1
    End If
Else
    Texto = 1
End If

Select Case Familiatext
    Case "001": 'Banco do brasil
        Select Case Cmb_carteira
            Case "11 - Simples - Com Registro":
                Texto = FunTamanhoTextoZeroEsq(Texto, 5)
                Especie = "DM"
            Case "11 - Vinculada - Com Registro":
                Texto = FunTamanhoTextoZeroEsq(Texto, 5)
                Especie = "DM"
            Case "17 - Direta Especial - Com Registro":
                Texto = FunTamanhoTextoZeroEsq(Texto, 5)
                Especie = "DM"
            Case "17Simples - Direta Especial Simples - Com Registro":
                Texto = FunTamanhoTextoZeroEsq(Texto, 5)
                Especie = "DM"
            Case "17-7 - Direta Especial - Com Registro Convênio 7 dígitos":
                Texto = FunTamanhoTextoZeroEsq(Texto, 10)
                Especie = "DM"
            Case "18 - Simples - Sem Registro":
                Texto = FunTamanhoTextoZeroEsq(Texto, 5)
                Especie = "RC"
            Case "18-7 - Simples - Sem Registro - Convênio 7 dígitos":
                Texto = FunTamanhoTextoZeroEsq(Texto, 10)
                Especie = "RC"
        End Select
    Case "033": 'Santander
        Select Case Cmb_carteira
            Case "COB - Cobrança Simples":
                Texto = FunTamanhoTextoZeroEsq(Texto, 7)
                Especie = "RC"
            Case "COBR - Cobrança Simples - Rápida Com Registro":
                Texto = FunTamanhoTextoZeroEsq(Texto, 7)
                Especie = "DM"
            Case "COBR-Nova - Cobrança Simples - Rápida Com Registro"
                Texto = FunTamanhoTextoZeroEsq(Texto, 12)
                Especie = "DM"
            Case "CSR - Cobrança Simples Sem Registro":
                Texto = FunTamanhoTextoZeroEsq(Texto, 12)
                Especie = "RC"
            Case "ECR - Cobrança Simples Com Registro":
                Texto = FunTamanhoTextoZeroEsq(Texto, 12)
                Especie = "DM"
        End Select
    Case "104": 'Caixa
        Select Case Cmb_carteira
            Case "CR - Cobrança Rápida":
                Texto = FunTamanhoTextoZeroEsq(Texto, 9)
                Especie = "RC"
            Case "CS - Cobrança Simples"
                Texto = FunTamanhoTextoZeroEsq(Texto, 9)
                Especie = "RC"
            Case "SR - Cobrança Sem Registro":
                Texto = FunTamanhoTextoZeroEsq(Texto, 8)
                Especie = "DM"
            Case "SR5 - SINCO - Sem Registro":
                Texto = FunTamanhoTextoZeroEsq(Texto, 17)
                Especie = "DM"
            Case "SIG14 - SIG Com Registro - Emissão pelo Cedente":
                Txt_local_de_pagamento = "Preferencialmente nas Casas Lotéricas até o valor limite"
                Texto = FunTamanhoTextoZeroEsq(Texto, 15)
                Especie = "DM"
        End Select
    Case "237": 'Bradesco
        Texto = FunTamanhoTextoZeroEsq(Texto, 11)
        Especie = "DM"
    Case "341": 'Itaú
        Select Case Cmb_carteira
            Case "109 - Direta Eletrônica Sem Emissão - Simples":
                Texto = FunTamanhoTextoZeroEsq(Texto, 8)
                Especie = "DM"
            Case "112 - Escritual Eletrônica - simples / contratual":
                Texto = FunTamanhoTextoZeroEsq(Texto, 8)
                Especie = "RC"
            Case "175 - Sem Registro Sem Emissão":
                Texto = FunTamanhoTextoZeroEsq(Texto, 8)
                Especie = "RC"
        End Select
    Case "356": 'ABN e Real
        Select Case Cmb_carteira
            Case "20 - Cobrança Simples":
                Texto = FunTamanhoTextoZeroEsq(Texto, 7)
                Especie = "NB"
        End Select
    Case "399": 'HSBC
        Select Case Cmb_carteira
            Case "CNR - Sem Registro":
                Texto = FunTamanhoTextoZeroEsq(Texto, 13)
                Especie = ""
        End Select
    Case "409": 'Unibanco
        Select Case Cmb_carteira
            Case "Especial":
                Texto = FunTamanhoTextoZeroEsq(Texto, 14)
                Especie = "RC"
        End Select
End Select
Txt_nosso_numero = Texto
Cmb_especie_doc.Text = Especie

If Financeiro_Contas_Receber = False Then TextoFiltro = frmFaturamento_Prod_Serv.cbo_PortBanco.ItemData(frmFaturamento_Prod_Serv.cbo_PortBanco.ListIndex) Else TextoFiltro = frmContas_Receber.cmbBanco.ItemData(frmContas_Receber.cmbBanco.ListIndex)
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select txt_Agencia, Codigo_cedente, Codigo_cedente_registrado from tbl_Instituicoes where ID = " & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    Select Case Familiatext
        Case "001": 'Banco do brasil
            Select Case Cmb_carteira
                Case "11 - Simples - Com Registro": Txt_codigo_cedente = TBAbrir!Codigo_cedente_registrado
                Case "11 - Vinculada - Com Registro": Txt_codigo_cedente = TBAbrir!Codigo_cedente_registrado
                Case "17 - Direta Especial - Com Registro": Txt_codigo_cedente = TBAbrir!Codigo_cedente_registrado
                Case "17Simples - Direta Especial Simples - Com Registro": Txt_codigo_cedente = TBAbrir!Codigo_cedente_registrado
                Case "17-7 - Direta Especial - Com Registro Convênio 7 dígitos": Txt_codigo_cedente = TBAbrir!Codigo_cedente_registrado
                Case "18 - Simples - Sem Registro": Txt_codigo_cedente = TBAbrir!codigo_cedente
                Case "18-7 - Simples - Sem Registro - Convênio 7 dígitos": Txt_codigo_cedente = TBAbrir!codigo_cedente
            End Select
        Case "033": 'Santander
            Select Case Cmb_carteira
                Case "COB - Cobrança Simples": Txt_codigo_cedente = TBAbrir!codigo_cedente
                Case "COBR - Cobrança Simples - Rápida Com Registro": Txt_codigo_cedente = TBAbrir!Codigo_cedente_registrado
                Case "COBR-Nova - Cobrança Simples - Rápida Com Registro": Txt_codigo_cedente = TBAbrir!Codigo_cedente_registrado
                Case "ECR - Cobrança Simples Com Registro": Txt_codigo_cedente = TBAbrir!Codigo_cedente_registrado
            End Select
        Case "104": 'Caixa
            Select Case Cmb_carteira
                Case "CR - Cobrança Rápida": Txt_codigo_cedente = TBAbrir!Codigo_cedente_registrado
                Case "CS - Cobrança Simples": Txt_codigo_cedente = TBAbrir!Codigo_cedente_registrado
                Case "SR - Cobrança Sem Registro": Txt_codigo_cedente = TBAbrir!codigo_cedente
                Case "SR5 - SINCO - Sem Registro": Txt_codigo_cedente = TBAbrir!codigo_cedente
                Case "SIG14 - SIG Com Registro - Emissão pelo Cedente": Txt_codigo_cedente = TBAbrir!Codigo_cedente_registrado
            End Select
        Case "237": 'Bradesco
                Select Case Cmb_carteira
                    Case "06 - Sem Registro": Txt_codigo_cedente = TBAbrir!codigo_cedente
                    Case "09 - Com Registro": Txt_codigo_cedente = TBAbrir!Codigo_cedente_registrado
                End Select
        Case "341": 'Itaú
            Select Case Cmb_carteira
                Case "109 - Direta Eletrônica Sem Emissão - Simples": Txt_codigo_cedente = TBAbrir!Codigo_cedente_registrado
                Case "112 - Escritual Eletrônica - simples / contratual": Txt_codigo_cedente = TBAbrir!Codigo_cedente_registrado
                Case "175 - Sem Registro Sem Emissão": Txt_codigo_cedente = TBAbrir!codigo_cedente
            End Select
        Case "356": 'ABN e Real
                Select Case Cmb_carteira
                    Case "20 - Cobrança Simples": Txt_codigo_cedente = TBAbrir!codigo_cedente
                End Select
        Case "399": 'HSBC
            Select Case Cmb_carteira
                Case "CNR - Sem Registro": Txt_codigo_cedente = TBAbrir!codigo_cedente
            End Select
        Case "409": 'Unibanco
            Select Case Cmb_carteira
                Case "Especial": Txt_codigo_cedente = TBAbrir!codigo_cedente
            End Select
    End Select
    If Financeiro_Contas_Receber = False Then
        Txt_agencia_codigo_cedente = frmFaturamento_Prod_Serv.txt_Agencia & "/" & Txt_codigo_cedente
    Else
        Txt_agencia_codigo_cedente = TBAbrir!txt_Agencia & "/" & Txt_codigo_cedente
    End If
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro
BoletoPDF = ""

If ProcVerifCampos(False, False) = False Then Exit Sub
Remessa = False
Enviar_Email = False
Sit_REG = 0

If Emissor = "Cobrebemx" Then
    ProcPassaDadosContaCorrenteParaCobreBemX Cmb_carteira, Cmb_carteira1, Txt_codigo_cedente, IDempresa, IIf(Chk_novo.Value = 1, True, IIf(Chk_atualizar.Value = 1, True, False)), Txt_assunto
    If Permitido1 = False Then Exit Sub
    ProcPassaDadosBoletosParaCobreBemX1
    CobreBemX1.ImprimeBoletos
    ProcGravarDadosBoleto
Else 'Tecnospeed

If Emissor = "Tecnospeed" And txtStatus = "" Then
    Sit_REG = 0
    ProcStatusBoleto
End If

caminho = ""

With CommonDialog1
    arq = "Arqs. PDF(*.pdf)|*.pdf|Todos " & "Arqs. (*.*)|*.*"
    .filename = ""
    .Filter = arq
    .FilterIndex = 1
    .InitDir = txtDiretorioBoleto
    .DefaultExt = "*.pdf"
    .ShowOpen
    caminho = .filename
End With

BoletoPDF = caminho
Sit_REG = 0
    If txtStatus <> "EMITIDO" Or txtStatus = "" Then
        Do While txtStatus <> "EMITIDO" Or txtStatus = ""
        If Sit_REG = 1 Then Exit Sub
            PlugEmitirBoleto
        Loop
    End If
    
    If txtProtocolo.Text = "" And txtIDIntegracao <> "" Then
        Do While txtProtocolo = "" And Sit_REG = 0
            txtProtocolo.Text = PlugGerarProtocoloBoleto(txtIDIntegracao)
        Loop
    End If
    
    If txtStatus = "SALVO" Then
        Do While txtStatus = "SALVO"
          txtStatus = PlugConsultarBoleto(txtIDIntegracao)
        Loop
    End If
     
    If txtIDIntegracao <> "" And txtProtocolo <> "" And txtStatus = "EMITIDO" Then
        PlugGerarPDFBoleto (txtProtocolo)
        Conexao.Execute ("Update tbl_Instituicoes set NossoNumero = " & TituloNossoNumero & " Where id = " & IDBanco)
        ProcGravarDadosBoleto
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Function ProcVerifCampos(Enviar_Email As Boolean, GerarRemessa As Boolean) As Boolean
On Error GoTo tratar_erro

If Enviar_Email = True Then
    Acao = "enviar e-mail"
ElseIf GerarRemessa = True Then
        Acao = "gerar arquivo remessa"
    Else
        Acao = "visualizar impressão"
End If
If Financeiro_Contas_Receber = True Then IDempresa = frmContas_Receber.Cmb_empresa.ItemData(frmContas_Receber.Cmb_empresa.ListIndex) Else IDempresa = frmFaturamento_Prod_Serv.txtIDEmpresa

ProcVerifCampos = True
If Txt_codigo_cedente = "" Then
    NomeCampo = "o código do cedente no cadastro do banco"
    ProcVerificaAcao
    ProcVerifCampos = False
    Exit Function
End If
If Cmb_carteira = "" Then
    NomeCampo = "a carteira"
    ProcVerificaAcao
    Cmb_carteira.SetFocus
    ProcVerifCampos = False
    Exit Function
End If
If Cmb_carteira1.Visible = True And Cmb_carteira1 = "" Then
    NomeCampo = "a carteira"
    ProcVerificaAcao
    Cmb_carteira1.SetFocus
    ProcVerifCampos = False
    Exit Function
End If
If Txt_nosso_numero = "" Then
    NomeCampo = "o nosso número"
    ProcVerificaAcao
    Txt_nosso_numero.SetFocus
    ProcVerifCampos = False
    Exit Function
End If
If Cmb_endereco = "" Then
    NomeCampo = "o endereço de cobrança"
    ProcVerificaAcao
    If Frame2.Enabled = True Then Cmb_endereco.SetFocus
    ProcVerifCampos = False
    Exit Function
End If
valor = IIf(Txt_valor_doc = "", 0, Txt_valor_doc)
If valor <= 0 Then
    NomeCampo = "valor do documento"
    ProcVerificaAcao
    Txt_valor_doc.SetFocus
    ProcVerifCampos = False
    Exit Function
End If
If Enviar_Email = True Then
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select ID from Empresa_email where ID_empresa = " & IDempresa & " and Aplicacao = 'F'", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = True Then
        USMsgBox ("Não será possível enviar o e-mail, pois não existe e-mail do financeiro configurado na empresa."), vbExclamation, "CAPRIND v5.0"
        ProcVerifCampos = False
        TBAbrir.Close
        Exit Function
    End If
    TBAbrir.Close
    
    If Txt_assunto = "" Then
        NomeCampo = "o assunto"
        ProcVerificaAcao
        Txt_assunto.SetFocus
        ProcVerifCampos = False
        Exit Function
    End If
End If

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Private Sub ProcGravarDadosBoleto()
On Error GoTo tratar_erro

If Remessa = True Then
    Evento = "Gerar arquivo remessa"
ElseIf Enviar_Email = True Then
        Evento = "Enviar e-mail"
    Else
        If Chk_atualizar.Value = 1 Then Evento = "Atualizar boleto" Else Evento = "Emitir boleto"
End If
    
Set TBGravar = CreateObject("adodb.recordset")
If Financeiro_Contas_Receber = False Then
    TBGravar.Open "Select * from tbl_Detalhes_Recebimento where Id = " & IDlista, Conexao, adOpenKeyset, adLockOptimistic
Else
    TBGravar.Open "Select * from tbl_Detalhes_Recebimento where IDContaReceber = " & frmContas_Receber.txtidintconta, Conexao, adOpenKeyset, adLockOptimistic
End If
If TBGravar.EOF = False Then
    If Chk_novo.Value = 1 And Chk_atualizar.Enabled = True Then
        TBGravar!Vencimento_boleto = Cmb_vencimento
    ElseIf Chk_atualizar.Value = 1 Then
            TBGravar!Vencimento_boleto = Cmb_novo_vencimento
        Else
            TBGravar!Vencimento_boleto = txt_Vencimento
    End If
    TBGravar!Valor_boleto = Txt_valor_doc
    TBGravar!Numero_documento = Txt_numero_doc
    TBGravar!Carteira = Cmb_carteira
    If Cmb_carteira1.Visible = True Then TBGravar!Carteira1 = Cmb_carteira1
    TBGravar!Nosso_Numero = Txt_nosso_numero
    TBGravar!Juros = IIf(Txt_percentual_juros = "", Null, Txt_percentual_juros)
    TBGravar!Multa = IIf(Txt_percentual_multa = "", Null, Txt_percentual_multa)
    TBGravar!Desconto = IIf(Txt_percentual_desconto = "", Null, Txt_percentual_desconto)
    TBGravar!Dias_Protesto = IIf(Txt_dias_protesto = "", Null, Txt_dias_protesto)
    TBGravar!Acrescimos = IIf(Txt_outros_acrescimos = "", Null, Txt_outros_acrescimos)
'===================================================================================================
    TBGravar!Valor_desconto = IIf(Txt_desconto.Text = "", Null, Txt_desconto.Text)
    If Txt_desconto <> "" Then
    TBGravar!DataLimiteDesconto = txtDatalimiteDesc.Value
    Else
    TBGravar!DataLimiteDesconto = ""
    End If
    TBGravar!Valor_Outras = IIf(Txt_outros_acrescimos = "", Null, Txt_outros_acrescimos)
    TBGravar!Valor_mora_multa_juros = IIf(Txt_mora = "", Null, Txt_mora)
    TBGravar!Valor_cobrado = IIf(Txt_valor_cobrado = "", Null, Txt_valor_cobrado)
'===================================================================================================
    
    TBGravar!Nosso_Numero = Txt_nosso_numero.Text
    TBGravar!protocolo = txtProtocolo.Text
    TBGravar!status = txtStatus
    
    TBGravar!Instrucoes = Txt_instrucoes
    
    Dataini = Txt_data_doc
    If IsNull(TBGravar!Data_emissao) = True Or TBGravar!Data_emissao = "" Or TBGravar!Data_emissao <> Dataini Then TBGravar!Data_emissao = Date
    
    If Enviar_Email = True Then
        TBGravar!Assunto = Txt_assunto
        TBGravar!Enviado = True
        TBGravar!data_envio = Date
        Chk_email_enviado.Value = 1
    End If
    TBGravar!ID_Cobranca = Cmb_endereco.ItemData(Cmb_endereco.ListIndex)
    
    If Financeiro_Contas_Receber = False Then
        TBGravar!txt_Portador_Banco = frmFaturamento_Prod_Serv.cbo_PortBanco
        
        Set TBFI = CreateObject("adodb.recordset")
        TBFI.Open "Select IDIntconta, ID_nota, txt_ndocumento, Valor, Vencimento, Observacoes from tbl_Contas_receber where id_nota = " & frmFaturamento_Prod_Serv.txtId & " and parcela = '" & frmFaturamento_Prod_Serv.lst_Duplicata.SelectedItem.ListSubItems(2) & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBFI.EOF = False Then
            TBFI!txt_ndocumento = Txt_numero_doc
            TBFI!valor = Txt_valor_cobrado
            TBFI!Vencimento = TBGravar!Vencimento_boleto
            If Remessa = False Then
                If Chk_novo.Value = 1 And Chk_atualizar.Enabled = True Then
                    If IsNull(TBFI!Observacoes) = False And TBFI!Observacoes <> "" Then TBFI!Observacoes = TBFI!Observacoes & " | Gerado novo boleto dia " & Txt_data_doc & " por " & pubUsuario Else TBFI!Observacoes = "Gerado novo boleto dia " & Txt_data_doc & " por " & pubUsuario
                ElseIf Chk_atualizar.Value = 1 Then
                        If IsNull(TBFI!Observacoes) = False And TBFI!Observacoes <> "" Then TBFI!Observacoes = TBFI!Observacoes & " | Boleto atualizado dia " & Txt_data_doc & " por " & pubUsuario Else TBFI!Observacoes = "Boleto atualizado dia " & Txt_data_doc & " por " & pubUsuario
                End If
            End If
            TBFI.Update
            
            ProcGravarNumeroBoleto TBFI!IDintconta, TBFI!ID_nota
            ProcGavarPCJurosMulta TBFI!IDintconta, TBFI!ID_nota, Valor_IPI, ValorTotal, "R", False
        End If
        TBFI.Close
        
        '==================================
        Modulo = Formulario
        With frmFaturamento_Prod_Serv
            ID_documento = IDlista
            .ProcVerificaTipoNF False
            If .txtNFiscal = "" Then NomeCampo = "N° ordem" Else NomeCampo = "N° nota"
            Documento = NomeCampo & ": " & .txtNFiscal & " - Tipo: " & TipoNF & " - Série: " & .txtSerie
            Documento1 = "Data vencimento: " & Format(TBGravar!Vencimento_boleto, "dd/mm/yy") & " - Valor: " & Format(TBGravar!Valor_boleto, "###,##0.00") & " - Parcela: " & .lst_Duplicata.SelectedItem.ListSubItems(2)
        End With
        ProcGravaEvento
        '==================================
    Else
        ProcEnviaDadosContaReceber
        '==================================
        Modulo = Formulario
        ID_documento = TBGravar!ID
        Documento = "Número da conta : " & frmContas_Receber.txtidintconta
        Documento1 = ""
        ProcGravaEvento
        '==================================
    End If
    TBGravar.Update
Else
    TBGravar.AddNew
    TBGravar!int_NotaFiscal = "000000"
    TBGravar!ID_nota = 0
    TBGravar!IDIntegracao = txtIDIntegracao
    TBGravar!protocolo = txtProtocolo
    With frmContas_Receber
        TBGravar!txt_Parcela = .txtparcela
        TBGravar!dt_Vencimento = .mskVencimento
        TBGravar!dbl_Valor = .txtValor
        TBGravar!Valor_Extenso = FunValorExtenso(.txtValor)
        TBGravar!txt_tipopagto = .cmb_forma

    End With
    ProcEnviaDadosContaReceber
    TBGravar.Update
    '==================================
    Modulo = Formulario
    ID_documento = TBGravar!ID
    Documento = "Número da conta : " & frmContas_Receber.txtidintconta
    Documento1 = ""
    ProcGravaEvento
    '==================================
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviaDadosContaReceber()
On Error GoTo tratar_erro

With frmContas_Receber
    TBGravar!txt_Portador_Banco = .cmbBanco
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from tbl_Instituicoes where ID = " & .cmbBanco.ItemData(.cmbBanco.ListIndex), Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        TBGravar!txt_Agencia = TBAbrir!txt_Agencia
        TBGravar!txt_Conta = TBAbrir!txt_Conta
    End If
    TBAbrir.Close
    If Chk_novo.Value = 1 And Chk_atualizar.Enabled = True Then
        TBGravar!Vencimento_boleto = Cmb_vencimento
    ElseIf Chk_atualizar.Value = 1 Then
            TBGravar!Vencimento_boleto = Cmb_novo_vencimento
        Else
            TBGravar!Vencimento_boleto = txt_Vencimento
    End If
    TBGravar!Valor_boleto = Txt_valor_cobrado
    TBGravar!Numero_documento = Txt_numero_doc
    TBGravar!Carteira = Cmb_carteira
    TBGravar!Nosso_Numero = Txt_nosso_numero
    TBGravar!Juros = IIf(Txt_percentual_juros = "", Null, Txt_percentual_juros)
    TBGravar!Multa = IIf(Txt_percentual_multa = "", Null, Txt_percentual_multa)
    TBGravar!Desconto = IIf(Txt_percentual_desconto = "", Null, Txt_percentual_desconto)
    TBGravar!Dias_Protesto = IIf(Txt_dias_protesto = "", Null, Txt_dias_protesto)
    TBGravar!Acrescimos = IIf(Txt_outros_acrescimos = "", Null, Txt_outros_acrescimos)
    TBGravar!Instrucoes = Txt_instrucoes
    If IsNull(TBGravar!Data_emissao) = True Or TBGravar!Data_emissao = "" Then TBGravar!Data_emissao = Date
    TBGravar!IdContaReceber = .txtidintconta
    
    Set TBFI = CreateObject("adodb.recordset")
    TBFI.Open "Select IDIntconta, txt_ndocumento, Valor, Vencimento, Observacoes from tbl_Contas_receber where IDIntconta = " & .txtidintconta, Conexao, adOpenKeyset, adLockOptimistic
    If TBFI.EOF = False Then
        TBFI!txt_ndocumento = Txt_numero_doc
        TBFI!valor = Txt_valor_cobrado
        TBFI!Vencimento = TBGravar!Vencimento_boleto
        If Remessa = False Then
            If Chk_novo.Value = 1 And Chk_atualizar.Enabled = True Then
                If IsNull(TBFI!Observacoes) = False And TBFI!Observacoes <> "" Then TBFI!Observacoes = TBFI!Observacoes & " | Gerado novo boleto dia " & Txt_data_doc & " por " & pubUsuario Else TBFI!Observacoes = "Gerado novo boleto dia " & Txt_data_doc & " por " & pubUsuario
            ElseIf Chk_atualizar.Value = 1 Then
                    If IsNull(TBFI!Observacoes) = False And TBFI!Observacoes <> "" Then TBFI!Observacoes = TBFI!Observacoes & " | Boleto atualizado dia " & Txt_data_doc & " por " & pubUsuario Else TBFI!Observacoes = "Boleto atualizado dia " & Txt_data_doc & " por " & pubUsuario
            End If
        End If
        TBFI.Update
    End If
    TBFI.Close
    
    If Enviar_Email = True Then
        TBGravar!Assunto = Txt_assunto
        TBGravar!Enviado = True
        TBGravar!data_envio = Date
        Chk_email_enviado.Value = 1
    End If
    TBGravar!ID_Cobranca = Cmb_endereco.ItemData(Cmb_endereco.ListIndex)
    
    ProcGravarNumeroBoleto .txtidintconta, TBGravar!ID_nota
    ProcGavarPCJurosMulta .txtidintconta, TBGravar!ID_nota, Valor_IPI, ValorTotal, "R", False
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcGravarNumeroBoleto(IDConta As Long, IDnota As Long)
On Error GoTo tratar_erro

If Chk_novo.Value = 1 Or Chk_atualizar.Value = 1 Then
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from tbl_Detalhes_Recebimento_Nboletos where IDContaReceber = " & IDConta & " and Nosso_numero = '" & Txt_nosso_numero & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = True Then TBAbrir.AddNew
    TBAbrir!Data = Date
    TBAbrir!Responsavel = pubUsuario
    TBAbrir!IdContaReceber = IDConta
    TBAbrir!Nosso_Numero = Txt_nosso_numero
    TBAbrir!ID_nota = IDnota
    TBAbrir.Update
    TBAbrir.Close
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcPassaDadosBoletosParaCobreBemX1()
On Error GoTo tratar_erro

CobreBemX1.DocumentosCobranca.Clear

Set Boleto = CobreBemX1.DocumentosCobranca.Add
Boleto.NumeroDocumento = Txt_numero_doc
Boleto.NomeSacado = Txt_sacado

If Len(txt_CNPJ) > 11 Then
    Boleto.CNPJSacado = txt_CNPJ
Else
    Boleto.CPFSacado = txt_CNPJ
End If


If Familiatext = "104" And Cmb_carteira = "SIG14 - SIG Com Registro - Emissão pelo Cedente" Then

    Set TBAliquota = CreateObject("adodb.recordset")
    TBAliquota.Open "Select Endereco_cobranca from Empresa where Codigo = " & Txt_IDempresa & " and Endereco_cobranca IS NOT NULL and Endereco_cobranca <> N''", Conexao, adOpenKeyset, adLockOptimistic
    If TBAliquota.EOF = False Then
        CobreBemX1.PadroesBoleto.IdentificacaoCedente = TBAliquota!endereco_Cobranca
    End If
    CobreBemX1.PadroesBoleto.PadroesBoletoImpresso.LayoutBoleto = "Padrao"
    
    'Utilizar para sair o endereço em outro campo
    CobreBemX1.PadroesBoleto.PadroesBoletoImpresso.LayoutBoleto = "PadraoReciboPersonalizado"
    CobreBemX1.PadroesBoleto.PadroesBoletoImpresso.HTMLReciboPersonalizado = TxtHTLM
    CobreBemX1.PadroesBoleto.PadroesBoletoImpresso.MargemSuperior = 0
    
    CobreBemX1.LocalPagamento = "Preferencialmente nas Casas Lotéricas até o valor limite"
    
    'Alterar o segmento P, posição 59 para 1, pois esse é o informativo de cobrança registrada do boleto da caixa
    Dim MDados1 As Object
    Set MDados1 = Boleto.MeusDados.Add
    MDados1.Nome = "FormaCadastramento"
    MDados1.valor = 1
    
    'Alterar SEGMENTO P, posição 224 para 2, pois o boleto tem ordem de protesto
    If Txt_dias_protesto <> "" Then Boleto.InstrucaoCobranca3 = 2
    
End If

Boleto.TipoDocumentoCobranca = Cmb_especie_doc
Boleto.EnderecoSacado = EnderecoBoleto
Boleto.BairroSacado = BairroBoleto
Boleto.CidadeSacado = CidadeBoleto
Boleto.EstadoSacado = EstadoBoleto
Boleto.CepSacado = CEPBoleto
Boleto.DataDocumento = Format(Txt_data_doc, "dd/mm/yyyy") 'A formatação de data tem que ser dd/mm/yyyy

If Chk_novo.Value = 1 And Chk_atualizar.Enabled = True Then
    Boleto.DataVencimento = Format(Cmb_vencimento, "dd/mm/yyyy")
ElseIf Chk_atualizar.Value = 1 Then
        Boleto.DataVencimento = Format(Cmb_novo_vencimento, "dd/mm/yyyy")
    Else
        Boleto.DataVencimento = Format(txt_Vencimento, "dd/mm/yyyy")
End If

Boleto.DataProcessamento = Format(Txt_data_processamento, "dd/mm/yyyy")

'======================================================================
' Valor do documento
'======================================================================
Boleto.ValorDocumento = Txt_valor_doc
'CobreBemX1.DocumentosCobranca.Item.
'Boleto.ValorDocumento = Txt_valor_cobrado
Boleto.ValorDesconto = IIf(Txt_desconto.Text <> "", Txt_desconto.Text, 0)
Boleto.DataLimiteDesconto = IIf(Txt_desconto.Text <> "", Format(txtDatalimiteDesc.Value, "dd/mm/yyyy"), "000000") 'Format(txtDatalimiteDesc.Value, "dd/mm/yyyy")
'Boleto.ValorAbatimento = Txt_desconto.Text
'======================================================================

Boleto.PercentualJurosDiaAtraso = IIf(Txt_percentual_juros = "", 0, Txt_percentual_juros)
Boleto.PercentualMultaAtraso = IIf(Txt_percentual_multa = "", 0, Txt_percentual_multa)
Boleto.PercentualDesconto = IIf(Txt_percentual_desconto = "", 0, Txt_percentual_desconto)
If Txt_dias_protesto <> "" Then Boleto.DiasProtesto = Txt_dias_protesto
Boleto.ValorOutrosAcrescimos = IIf(Txt_outros_acrescimos = "", 0, Txt_outros_acrescimos)
'Boleto.ValorOutrosAcrescimos = 0
If Txt_instrucoes <> "" Then
    Boleto.PadroesBoleto.Demonstrativo = Txt_instrucoes
    Boleto.PadroesBoleto.InstrucoesCaixa = Txt_instrucoes
End If

Boleto.ControleProcessamentoDocumento.Imprime = scpExecutar
Boleto.ControleProcessamentoDocumento.GravaRemessa = scpExecutar

If Enviar_Email = True Then
    Set Email = Boleto.EnderecosEmailSacado.Add
    Email.Nome = Boleto.NomeSacado
    Email.Endereco = TBCFOP!Email
    Boleto.ControleProcessamentoDocumento.EnviaEmail = scpExecutar
End If

Boleto.NossoNumero = Txt_nosso_numero
'If Len(Txt_nosso_numero) < (CobreBemX1.MascaraNossoNumero + 1) Then
    'Boleto.CalculaDacNossoNumero = 'True'
'End If

If Familiatext = "001" And Cmb_carteira = "11 - Simples - Com Registro" Then
    Boleto.BancoEmiteBoleto = True
    Boleto.InstrucaoCobranca3 = 2
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_novo_vencimento_Change()
On Error GoTo tratar_erro

ProcVerifDiasAtraso
'Txt_vencimento.Text = Cmb_novo_vencimento.Value

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcVerifDiasAtraso()
On Error GoTo tratar_erro

Data = Cmb_novo_vencimento
DataFim = txt_Vencimento

With Txt_dias_atraso
    If Data > DataFim Then
        .Text = Data - DataFim
        .Locked = False
        .TabStop = True
        Chk_calcular_juros_multa.Value = 1
        Chk_calcular_juros_multa.Enabled = True
    Else
        .Text = 0
        .Locked = True
        .TabStop = Fase
        Chk_calcular_juros_multa.Value = 0
        Chk_calcular_juros_multa.Enabled = False
    End If
    ProcCalculaJurosMulta .Text
End With
            
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCalculaJurosMulta(DiasAtraso As Integer)
On Error GoTo tratar_erro

valor = Txt_valor_doc
If Chk_calcular_juros_multa.Value = 1 And DiasAtraso > 0 Then
    'Juros
    Valor_IPI = (valor * IIf(Txt_percentual_juros = "", 0, Txt_percentual_juros)) / 100
    Valor_IPI = Valor_IPI * IIf(Txt_dias_atraso = "", 0, Txt_dias_atraso)
    
    'Multa
    ValorTotal = (valor * IIf(Txt_percentual_multa = "", 0, Txt_percentual_multa)) / 100
    
    Txt_mora = Format(Valor_IPI + ValorTotal, "###,##0.00")
    Txt_valor_cobrado = Format(valor + Valor_IPI + ValorTotal, "###,##0.00")
Else
    Valor_IPI = 0
    ValorTotal = 0
    Txt_mora = ""
    Txt_valor_cobrado = Format(valor, "###,##0.00")
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_instrucoes_Click()
On Error GoTo tratar_erro

frmFaturamento_Prod_serv_boleto_instrucoes.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcRemessa()
On Error GoTo tratar_erro

Acao = "gerar o arquivo remessa"
If ProcVerifCampos(False, True) = False Then Exit Sub


Set TBFI = CreateObject("adodb.recordset")
If Financeiro_Contas_Receber = False Then
    With frmFaturamento_Prod_Serv
        IDempresa = .txtIDEmpresa.Text
        TBFI.Open "Select * from tbl_Detalhes_Recebimento where ID = " & IDlista, Conexao, adOpenKeyset, adLockOptimistic
        If TBFI.EOF = False Then
            If IsNull(TBFI!Seq_remessa) = False And TBFI!Seq_remessa <> "" Then
                USMsgBox ("Já existe arquivo remessa gerado para esta duplicata."), vbExclamation, "CAPRIND v5.0"
                If USMsgBox("Deseja prosseguir assim mesmo?", vbYesNo, "CAPRIND v5.0") = vbNo Then
                    TBFI.Close
                    Exit Sub
                End If
            End If
        End If
        TBFI.Close
    End With
Else
    With frmContas_Receber
        Aplic = .Cmb_empresa.ItemData(.Cmb_empresa.ListIndex)
        TBFI.Open "Select * from tbl_Detalhes_Recebimento where IDContaReceber = " & .txtidintconta, Conexao, adOpenKeyset, adLockOptimistic
        If TBFI.EOF = False Then
            If IsNull(TBFI!Seq_remessa) = False And TBFI!Seq_remessa <> "" Then
                If USMsgBox("Já existe arquivo remessa gerado para esta duplicata, deseja gerar um novo arquivo?", vbYesNo, "CAPRIND v5.0") = vbNo Then
                    TBFI.Close
                    Exit Sub
                End If
            End If
        End If
        TBFI.Close
    End With
End If
Permitido = True
Select Case Familiatext
    Case "001": 'Banco do brasil
        Select Case Cmb_carteira
            Case "18 - Simples - Sem Registro": Permitido = False
            Case "18-7 - Simples - Sem Registro - Convênio 7 dígitos": Permitido = False
        End Select
    Case "033": 'Santander
        Select Case Cmb_carteira
            Case "COB - Cobrança Simples": Permitido = False
            Case "CSR - Cobrança Simples Sem Registro": Permitido = False
        End Select
    Case "104": 'Caixa
        Select Case Cmb_carteira
            Case "SR - Cobrança Sem Registro": Permitido = False
            Case "SR5 - SINCO - Sem Registro": Permitido = False
        End Select
    Case "237": 'Bradesco
        Select Case Cmb_carteira
            Case "06 - Sem Registro": Permitido = False
        End Select
    Case "341": 'Itaú
        Select Case Cmb_carteira
            Case "175 - Sem Registro Sem Emissão": Permitido = False
        End Select
    Case "399": 'HSBC
        Select Case Cmb_carteira
            Case "CNR - Sem Registro": Permitido = False
        End Select
    Case "409": 'Unibanco
        Select Case Cmb_carteira
            Case "Especial": Permitido = False
        End Select
End Select
If Permitido = False Then
    USMsgBox ("Não é permitido gerar aquivo remessa para esta carteira, pois a mesma não é registrada."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Remessa = True
Enviar_Email = False
ProcGravarDadosBoleto

If Emissor = "CobreBemX" Then
    ProcPassaDadosContaCorrenteParaCobreBemX Cmb_carteira, Cmb_carteira1, Txt_codigo_cedente, IDempresa, True, Txt_assunto
    If Permitido1 = False Then Exit Sub
    ProcPassaDadosBoletosParaCobreBemX1
    Diretorio = txtLocalArmazenamento.Text
    CobreBemX1.ArquivoRemessa.Diretorio = Diretorio
    CobreBemX1.ArquivoRemessa.Arquivo = Arquivo
    CobreBemX1.ArquivoRemessa.Layout = Layout
    CobreBemX1.GravaArquivoRemessa
    If Financeiro_Contas_Receber = True And Sit_REG = 2 Or Financeiro_Contas_Receber = False Then USMsgBox ("Arquivo remessa " & Arquivo & " gerado com sucesso."), vbInformation, "CAPRIND v5.0"
    
Else
Sit_REG = 0
    DiretorioRemessa = txtLocalArmazenamento.Text
    If txtIDIntegracao = "" Then
        Do While txtStatus <> "EMITIDO"
            PlugEmitirBoleto
        Loop
    End If
    
    If txtIDIntegracao <> "" And Sit_REG = 0 Then
        PlugGerarRemessa (txtIDIntegracao)
        If Financeiro_Contas_Receber = True Then frmContas_Receber.ProcCarregaLista (1)
    Else
        USMsgBox "Não foi possivel gerar a remessa, tente de novo!", vbInformation, "CAPRIND v5.0"
    End If
End If


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviarEmail()
On Error GoTo tratar_erro

Remessa = False
Enviar_Email = True
If Financeiro_Contas_Receber = True Then
    With frmContas_Receber
        Set TBContas = CreateObject("adodb.recordset")
        TBContas.Open "Select Tipo, IDCliente from tbl_contas_receber where IDIntconta = " & IIf(.txtidintconta = "", 0, .txtidintconta), Conexao, adOpenKeyset, adLockOptimistic
        If TBContas.EOF = False Then
            If TBContas!Tipo = "CL" Then
                NomeTabela1 = "Clientes_Contatos"
                NomeCampo1 = "IDCliente = " & TBContas!IDCliente
            Else
                NomeTabela1 = "Contatos_fornecedor"
                NomeCampo1 = "IdFornecedor = " & TBContas!IDCliente
            End If
        End If
    End With
Else
    With frmFaturamento_Prod_Serv
        Set TBFI = CreateObject("adodb.recordset")
        TBFI.Open "Select NF.ID_int_Cliente from tbl_Detalhes_Recebimento DR INNER JOIN tbl_Dados_Nota_Fiscal NF on DR.ID_nota = NF.ID where DR.Id = " & .lst_Duplicata.SelectedItem.ListSubItems(3), Conexao, adOpenKeyset, adLockOptimistic
        If TBFI.EOF = False Then
            If Len(.txttipocliente) = 2 Then
                NomeTabela1 = "Clientes_Contatos"
                NomeCampo1 = "IDCliente = " & TBFI!Id_Int_Cliente
            Else
                NomeTabela1 = "Contatos_fornecedor"
                NomeCampo1 = "IdFornecedor = " & TBFI!Id_Int_Cliente
            End If
        End If
        TBFI.Close
    End With
End If

Set TBCFOP = CreateObject("adodb.recordset")
TBCFOP.Open "Select IDContato, Email from " & NomeTabela1 & " where " & NomeCampo1 & " and Enviar_boleto = 'True' and EMail IS NOT NULL and EMail <> N''", Conexao, adOpenKeyset, adLockOptimistic
If TBCFOP.EOF = False Then
    ProcGravarDadosBoleto
    
    Do While TBCFOP.EOF = False
        If Emissor = "Cobrebemx" Then
            ProcPassaDadosContaCorrenteParaCobreBemX Cmb_carteira, Cmb_carteira1, Txt_codigo_cedente, IDempresa, IIf(Chk_novo.Value = 1, True, IIf(Chk_atualizar.Value = 1, True, False)), Txt_assunto
            If Permitido1 = False Then Exit Sub
            ProcPassaDadosBoletosParaCobreBemX1
            CobreBemX1.EnviaBoletosPorEmail
        Else
            USMsgBox "Aqui Tecnospeed"
        End If
        TBCFOP.MoveNext
    Loop
    
    If Financeiro_Contas_Receber = True And Sit_REG = 2 Or Financeiro_Contas_Receber = False Then USMsgBox ("E-mail enviado com sucesso."), vbInformation, "CAPRIND v5.0"
    Chk_email_enviado.Value = 1
    Txt_data_envio = Format(Date, "dd/mm/yy")
Else
    USMsgBox ("Não foi possível enviar o e-mail, pois não existe contato configurado para envio de boleto."), vbExclamation, "CAPRIND v5.0"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

If Financeiro_Contas_Receber = False Then
    With frmFaturamento_Prod_Serv
        .ProcCarregaListaDuplicatas IIf(.txtId = "", 0, .txtId)
    End With
ElseIf Sit_REG = 2 Then
    With frmContas_Receber
        .ProcCarregaLista (IIf(ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5))))
    End With
End If
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdLocal_Click()
On Error GoTo tratar_erro

    DS.OpenFolderWithExplorer txtLocalArmazenamento.Text

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Initialize()
On Error GoTo tratar_erro
'Boleto Tecnospeed

Set FBoletoX = New BoletoX.spdBoletoX

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyF5: ProcImprimir
    Case vbKeyF7: ProcRemessa
    Case vbKeyF8:
        If ProcVerifCampos(True, False) = False Then Exit Sub
        ProcEnviarEmail
    Case vbKeyEscape: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro
Dim DiretorioRemessa As String

ProcCarregaDadosBoleto

Escritural = False

    If Financeiro_Contas_Receber = True Then
        DiretorioRemessa = Localrel & "\Boletos\Arquivos remessa\" & frmContas_Receber.cmbBanco.Text
        DiretorioBoleto = Localrel & "\Boletos\ArquivosPDF"
        
        IDempresa = frmContas_Receber.Cmb_empresa.ItemData(frmContas_Receber.Cmb_empresa.ListIndex)
        IDBanco = frmContas_Receber.cmbBanco.ItemData(frmContas_Receber.cmbBanco.ListIndex)
        IDCliente = frmContas_Receber.txtIDcliente
        TipoSacado = frmContas_Receber.Cmb_tipo
        IDConta = frmContas_Receber.txtidintconta
        TituloNossoNumero = Txt_nosso_numero
    Else
        DiretorioRemessa = Localrel & "\Boletos\Arquivos remessa\" & frmFaturamento_Prod_Serv.cbo_PortBanco.Text
        DiretorioBoleto = Localrel & "\Boletos\ArquivosPDF"
        
        IDempresa = frmFaturamento_Prod_Serv.txtIDEmpresa.Text
        IDBanco = frmFaturamento_Prod_Serv.cbo_PortBanco.ItemData(frmFaturamento_Prod_Serv.cbo_PortBanco.ListIndex)
        IDCliente = frmFaturamento_Prod_Serv.txtIDcliente
        TipoSacado = "Cliente" 'frmFaturamento_Prod_Serv.Cmb_tipo
        IDDuplicata = frmFaturamento_Prod_Serv.Txt_ID_duplicata
        TituloNossoNumero = Txt_nosso_numero
    End If

    If DS.FileOrDirExists(DiretorioRemessa) = False Then
        MkDir DiretorioRemessa
    End If

    If DS.FileOrDirExists(DiretorioBoleto) = False Then
        MkDir DiretorioBoleto
    End If
    
    '================================================================================
    'Busca emissor de boletos
    '================================================================================
    StrSql = "select Id, EmissorBoleto,DiretorioRemessa, Escritural, Diretorioboleto From tbl_Instituicoes where Id = " & IDBanco
    
    TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Emissor = IIf(IsNull(TBAbrir!EmissorBoleto), "Cobrebemx", TBAbrir!EmissorBoleto)
        txtEmissor.Text = Emissor
        If TBAbrir!DiretorioRemessa = "" Or IsNull(TBAbrir!DiretorioRemessa) Then
            txtLocalArmazenamento = DiretorioRemessa
        Else
            txtLocalArmazenamento = TBAbrir!DiretorioRemessa
        End If
        
        If TBAbrir!DiretorioBoleto = "" Or IsNull(TBAbrir!DiretorioBoleto) Then
            txtDiretorioBoleto = DiretorioBoleto
        Else
            txtDiretorioBoleto = TBAbrir!DiretorioBoleto
        End If
        
        If TBAbrir!Escritural <> "" Or TBAbrir!Escritural <> Null Then
            Escritural = TBAbrir!Escritural
        Else
            Escritural = False
        End If
        
    End If
    TBAbrir.Close

    '================================================================================
    'Dados do cedente nossonumero e conta corrente
    '================================================================================
        StrSql = "select CNPJ, EmissorBoleto, int_NBanco,txt_Agencia,txt_Conta,DV,NossoNumero, txt_descricao from Empresa EMP inner join tbl_Instituicoes INST on EMP.codigo = INST.ID_empresa where codigo = " & IDempresa & " and id = " & IDBanco & ""
        TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            FBoletoX.Config.CedenteCpfCnpj = DS.ReturnNumbersOnly(TBAbrir!CNPJ)
            CedenteCpfCnpj = DS.ReturnNumbersOnly(TBAbrir!CNPJ)
            CedenteContaNumero = TBAbrir!txt_Conta
            CedenteContaNumeroDV = IIf(IsNull(TBAbrir!DV), "", TBAbrir!DV)
            CedenteConvenioNumero = TBAbrir!txt_Conta
            CedenteContaCodigoBanco = TBAbrir!int_NBanco
            
            'Se for Bradesco busca o ultimo numero do dia
    
            If CedenteContaCodigoBanco = 237 Then 'Bradesco
                StrSql = "SELECT Seq_remessa from tbl_Detalhes_Recebimento WHERE Data_emissao = '" & Date & "' and txt_Portador_Banco = '" & TBAbrir!Txt_descricao & "' AND txt_Conta = '" & TBAbrir!txt_Conta & "' order by Seq_Remessa"
            End If
            
            'Se for Itaú busca o ultimo numero emitido
     
            If CedenteContaCodigoBanco = 341 Then 'Itaú
                StrSql = "SELECT Seq_remessa from tbl_Detalhes_Recebimento WHERE txt_Portador_Banco = '" & TBAbrir!Txt_descricao & "' AND txt_Conta = '" & TBAbrir!txt_Conta & "' order by Seq_Remessa"
            End If
    
            TBContas.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
            If TBContas.EOF = False Then
                TBContas.MoveLast
                NumeroRemessa = IIf(IsNull(TBContas!Seq_remessa), 1, TBContas!Seq_remessa)
            End If
            TBContas.Close
        End If
        TBAbrir.Close
        
        CedenteContaNumeroDV = Right(CedenteContaNumero, 1) '"9"
        
'=================================================================
' Bradesco
'=================================================================
        If CedenteContaCodigoBanco = 237 Then
     
                Select Case Len(CedenteContaNumero)
                    Case 8:
                    Conta = Left(CedenteContaNumero, 7)
                    Numero = Len(Conta) - 7

                    Case 7:
                    Conta = Left(CedenteContaNumero, 6)
                    Numero = Len(Conta) - 6
                    
                    Case 6:
                    Conta = Left(CedenteContaNumero, 5)
                    Numero = Len(Conta) - 5
                    
                    Case 5:
                    Conta = Left(CedenteContaNumero, 4)
                    Numero = Len(Conta) - 4
                    
                    Case 4:
                    Conta = Left(CedenteContaNumero, 3)
                    Numero = 3
                End Select
            

                    If Len(Conta) < 7 Then
                        CedenteContaNumero = DS.FormatWithZeros(Left(Conta, Numero), 7) '"0000130"
                    Else
                        CedenteContaNumero = Conta
                    End If

            CedenteConvenioNumero = CedenteContaNumero '"0000130"
        End If
'=================================================================
' Itaú
'=================================================================
        If CedenteContaCodigoBanco = 341 Then
     
                Select Case Len(CedenteContaNumero)
                    Case 8:
                    Conta = Left(CedenteContaNumero, 7)
                    Numero = Len(Conta) - 7

                    Case 7:
                    Conta = Left(CedenteContaNumero, 6)
                    Numero = Len(Conta) - 6
                    
                    Case 6:
                    Conta = Left(CedenteContaNumero, 5)
                    Numero = Len(Conta) - 5
                    
                    Case 5:
                    Conta = Left(CedenteContaNumero, 4)
                    Numero = Len(Conta) - 4
                    
                    Case 4:
                    Conta = Left(CedenteContaNumero, 3)
                    Numero = 3
                End Select
            

                  '  If Len(Conta) < 7 Then
                    'teste = DS.FormatWithZeros(Left(Conta, Len(Conta)), 7) '"0000130"
                   '     CedenteContaNumero = DS.FormatWithZeros(Left(Conta, Len(Conta)), 7) '"0000130"
                   ' Else
                        CedenteContaNumero = Conta
                  '  End If

            CedenteConvenioNumero = CedenteContaNumero '"0000130"
        End If
    '=================================================================================
    'Dados do sacado
    '=================================================================================
    If Financeiro_Contas_Receber = True Then 'Contas receber busca por tipo e ID
        TipoSacado = frmContas_Receber.Cmb_tipo.Text
        IDCliente = frmContas_Receber.txtIDcliente
        
        Select Case TipoSacado
            Case "Cliente":
            StrSql = "select CPF_CNPJ, NomeRazao, EMAIL,Tel01 from Clientes where idCliente = " & IDCliente
            
            Case "Fornecedor":
            StrSql = "select CPF_CNPJ, Nome_Razao as NomeRazao, EMAIL,Telefones as Tel01 from Compras_fornecedores where idCliente = " & IDCliente
            
            Case "Funcionário":
            
            Case "Instituição":
            
        End Select
        
            TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                SacadoEmail = IIf(IsNull(TBAbrir!Email), "", TBAbrir!Email)
                SacadoNome = TBAbrir!NomeRazao
                SacadoCPFCNPJ = DS.ReturnNumbersOnly(TBAbrir!CPF_CNPJ)
                SacadoCelular = IIf(IsNull(TBAbrir!Tel01), "", TBAbrir!Tel01)
                TBAbrir.Close
            End If
            
    Else 'Faturamento nota fiscal busca por cnpj
        StrSql = "select CPF_CNPJ, NomeRazao, EMAIL,Tel01 from Clientes where CPF_CNPJ = '" & txt_CNPJ & "'"
        TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                SacadoEmail = IIf(IsNull(TBAbrir!Email), "", TBAbrir!Email)
                SacadoNome = TBAbrir!NomeRazao
                SacadoCPFCNPJ = DS.ReturnNumbersOnly(TBAbrir!CPF_CNPJ)
                SacadoCelular = IIf(IsNull(TBAbrir!Tel01), "", TBAbrir!Tel01)
                TBAbrir.Close
            Else
                TBAbrir.Close
                StrSql = "select CPF_CNPJ, Nome_Razao as NomeRazao, EMAIL,Telefones as Tel01 from Compras_fornecedores where CPF_CNPJ = '" & txt_CNPJ & "'"
                TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    SacadoEmail = IIf(IsNull(TBAbrir!Email), "", TBAbrir!Email)
                    SacadoNome = TBAbrir!NomeRazao
                    SacadoCPFCNPJ = DS.ReturnNumbersOnly(TBAbrir!CPF_CNPJ)
                    SacadoCelular = IIf(IsNull(TBAbrir!Tel01), "", TBAbrir!Tel01)
                    TBAbrir.Close
                End If
            End If
            

    End If

    '==================================================================================
    'Dados titulo
    '==================================================================================
    If Financeiro_Contas_Receber = True Then ' Se for emitir pelo módulo contas a receber
    StrSql = "Select Vencimento,Emissao,Valor,NFiscal,Parcela from tbl_contas_receber where IDIntconta =" & IDConta

        TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
        Parcela = Left(TBAbrir!Parcela, 3)
        Parcela = Right(Parcela, 2)
            TituloNumeroDocumento = Right(TBAbrir!NFiscal, 5) & "/" & Parcela
            TituloDataVencimento = TBAbrir!Vencimento
            TituloDataEmissao = TBAbrir!emissao
            TituloDataJuros = TBAbrir!Vencimento + 1
            TituloValor = Format(TBAbrir!valor, "0.#0")
            TituloMensagem01 = Txt_instrucoes
            TituloMensagem02 = ""
            TituloMensagem03 = ""
            TituloInformacoesAdicionais = ""
            TituloInstrucoes = ""
        End If
        TBAbrir.Close
    Else ' Se for emitir pelo emissor de nota fiscal
    StrSql = "Select dt_Vencimento,Data_emissao,dbl_Valor,int_NotaFiscal as NFiscal, txt_Parcela as Parcela,Status from tbl_detalhes_recebimento where ID = " & IDDuplicata
        
        TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
        Parcela = Left(TBAbrir!Parcela, 3)
        Parcela = Right(Parcela, 2)
            TituloNumeroDocumento = Right(TBAbrir!NFiscal, 5) & "/" & Parcela
            TituloDataVencimento = TBAbrir!dt_Vencimento
            TituloDataJuros = TBAbrir!dt_Vencimento + 1
            TituloDataEmissao = IIf(IsNull(TBAbrir!Data_emissao), Format(Now, "dd/mm/yyyy"), TBAbrir!Data_emissao)
            TituloValor = Format(TBAbrir!dbl_Valor, "0.#0")
            TituloMensagem01 = Txt_instrucoes
            TituloMensagem02 = ""
            TituloMensagem03 = ""
            TituloInformacoesAdicionais = ""
            TituloInstrucoes = ""
            txtStatus.Text = IIf(TBAbrir!status <> Null, TBAbrir!status, "")
        End If
        TBAbrir.Close
    End If
    
    'Juros e multa
    'código 1 valor em reais por dia, 2 valor em percentual, 3 isento de juros
    TituloCodigoJuros = 1
    TituloValorJuros = Txt_percentual_juros
    TituloValorMultaTaxa = Txt_percentual_multa
    
    If TituloValorJuros <> "" Or TituloValorJuros > 0 Then
        ValorJuros = TituloValor * (TituloValorJuros) / 100
    End If
    
    If TituloValorMultaTaxa <> "" Or TituloValorMultaTaxa > 0 Then
        ValorMulta = (TituloValor * TituloValorMultaTaxa) / 100
    End If
    
    If ValorMulta <> "" And ValorJuros <> "" And Emissor <> "CobreBemX" Then
        Txt_instrucoes = "APOS O VENCIMENTO COBRAR MULTA DE R$ " & Format(ValorMulta, "0.#0") & " E JUROS DE R$ " & Format(ValorJuros, "0.#0") & " AO DIA"
    End If
    
    TituloMensagem01 = Txt_instrucoes
    TituloCodigoMulta = 2
    TituloDataMulta = TituloDataJuros


    
    Txt_numero_doc = TituloNumeroDocumento
    
'    If Emissor = "Tecnospeed" And txtStatus = "" Then
'        Sit_REG = 0
'        ProcStatusBoleto
'    End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaDadosBoleto()
On Error GoTo tratar_erro

'Cmb_novo_vencimento = IIf(txt_Vencimento.Text <> "", txt_Vencimento.Text, Date)
'Cmb_vencimento = IIf(txt_Vencimento.Text <> "", txt_Vencimento.Text, Date)
Cmb_carteira.ListIndex = -1
Txt_nosso_numero = ""

Valor_IPI = 0 'Juros
ValorTotal = 0 'Multa

If Financeiro_Contas_Receber = True Then
    Caption = "Administrativo - Financeiro - Contas a receber - Boleto"
    With frmContas_Receber
        Set TBContas = CreateObject("adodb.recordset")
        TBContas.Open "Select * from tbl_contas_receber where IDIntconta = " & IIf(.txtidintconta = "", 0, .txtidintconta), Conexao, adOpenKeyset, adLockOptimistic
        IDDuplicata = .txtidintconta.Text
        If TBContas.EOF = False Then
            If TBContas!Tipo = "CL" Then
                NomeTabela = "Clientes"
                TipoFiltro = "C"
            Else
                NomeTabela = "Compras_fornecedores"
                TipoFiltro = "F"
            End If
            
            txtIDIntegracao = IIf(IsNull(TBContas!IDIntegracao), "", TBContas!IDIntegracao)
            txtProtocolo = IIf(IsNull(TBContas!protocolo), "", TBContas!protocolo)
            
            Set TBAbrir = CreateObject("adodb.recordset")
            
            StrSql = "Select I.*, IB.AssuntoEmail, IB.Desconto, IB.Dias_Protesto, IB.ID_Instrucoes, IB.Instrucoes_protesto, IB.Juros, IB.Multa from tbl_Instituicoes I inner Join tbl_Instituicoes_Instrucoes_Boleto IB on IB.ID_Instituicao = I.Id where I.txt_Descricao = '" & TBContas!Banco & "' and ID_empresa = " & IIf(IsNull(TBContas!ID_empresa), 0, TBContas!ID_empresa) & ""
            'Debug.print StrSql
            
            TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                Familiatext = TBAbrir!int_NBanco
                Img_logo_banco.Picture = LoadPicture(Localrel & "\Imagens\Bancos\" & TBAbrir!int_NBanco & ".jpg")
                Agencia = TBAbrir!txt_Agencia
                ContaCorrente = TBAbrir!txt_Conta
                NomeAgencia = IIf(IsNull(TBAbrir!Nome_agencia), "", TBAbrir!Nome_agencia)
                Txt_codigo_cedente.Text = ContaCorrente
                Txt_agencia_codigo_cedente = TBAbrir!txt_Agencia & "/" & Txt_codigo_cedente
                Txt_assunto = TBAbrir!AssuntoEmail
                Txt_percentual_juros = Format(TBAbrir!Juros, "###,##0.00")
                Txt_percentual_multa = Format(TBAbrir!Multa, "###,##0.00")
                Txt_percentual_desconto = Format(TBAbrir!Desconto, "###,##0.00")
                Txt_dias_protesto = TBAbrir!Dias_Protesto
                Txt_instrucoes = TBAbrir!Instrucoes_protesto
                ProcCarregaCarteira
            End If
            
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select Codigo, Razao from Empresa where Codigo = " & IIf(IsNull(TBContas!ID_empresa), 0, TBContas!ID_empresa), Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                Txt_IDempresa = TBAbrir!CODIGO
                Txt_cedente = TBAbrir!Razao
            End If
            
            'Carrega a carteira utilizada por último
            Set TBCFOP = CreateObject("adodb.recordset")
            TBCFOP.Open "Select Carteira, Carteira1, Assunto, Enviado FROM tbl_Detalhes_Recebimento where txt_Portador_Banco = '" & TBContas!Banco & "' and txt_Agencia = '" & Agencia & "' and txt_Conta = '" & ContaCorrente & "' and Carteira is not null order by Id desc", Conexao, adOpenKeyset, adLockOptimistic
            If TBCFOP.EOF = False Then
                If TBCFOP!Carteira <> "" Then
                Cmb_carteira = TBCFOP!Carteira
                End If
                
                If IsNull(TBCFOP!Carteira1) = False And TBCFOP!Carteira1 <> "" Then
                Cmb_carteira1 = TBCFOP!Carteira1
                End If
            End If
            
1:
            Txt_sacado = TBContas!Nome_Razao
            
            Set TBClientes = CreateObject("adodb.recordset")
            TBClientes.Open "Select idTipoEmpresa, CPF_CNPJ, IDCliente from " & NomeTabela & " where idcliente = " & TBContas!IDCliente, Conexao, adOpenKeyset, adLockOptimistic
            If TBClientes.EOF = False Then
                If TBClientes!idTipoEmpresa = 1 Then txt_CNPJ = TBClientes!CPF_CNPJ
                ProcCarregaComboEndCob "idcliente = " & TBClientes!IDCliente & " and Tipo = '" & TipoFiltro & "'", True
            End If
            
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select * from tbl_Detalhes_Recebimento where IDContaReceber = " & .txtidintconta & " and txt_Portador_Banco = '" & TBContas!Banco & "' and Nosso_numero IS NOT NULL and Nosso_numero <> N''", Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False Then
                If Sit_REG = 2 Then USMsgBox ("Já existe boleto emitido para esta parcela."), vbInformation, "CAPRIND v5.0"
                Chk_novo.Enabled = True
                Chk_atualizar.Enabled = True
                
                txt_Vencimento = Format(TBFI!Vencimento_boleto, "dd/mm/yy")
                Txt_data_doc = IIf(IsNull(TBFI!Data_emissao), "", Format(TBFI!Data_emissao, "dd/mm/yy"))
                Txt_numero_doc = TBFI!Numero_documento
                Txt_data_processamento = Txt_data_doc
                Txt_valor_doc = Format(TBFI!Valor_boleto, "###,##0.00")
                Txt_valor_cobrado = Format(TBFI!Valor_boleto, "###,##0.00")
                If IsNull(TBFI!Carteira) = False And TBFI!Carteira <> "" Then Cmb_carteira = TBFI!Carteira
                If IsNull(TBFI!Carteira1) = False And TBFI!Carteira1 <> "" And Familiatext = "001" Then Cmb_carteira1 = TBFI!Carteira1
                If IsNull(TBFI!Acrescimos) = False And TBFI!Acrescimos <> "" Then Txt_outros_acrescimos = Format(TBFI!Acrescimos, "###,##0.00")
                If TBFI!Enviado = True Then Chk_email_enviado.Value = 1 Else Chk_email_enviado.Value = 0
                Txt_data_envio = IIf(IsNull(TBFI!data_envio), "", Format(TBFI!data_envio, "dd/mm/yy"))
                If IsNull(TBFI!ID_Cobranca) = False And TBFI!ID_Cobranca <> "" Then ProcCarregaComboEndCob "idcobranca = " & TBFI!ID_Cobranca, False
                txtIDIntegracao = IIf(IsNull(TBFI!IDIntegracao), "", TBFI!IDIntegracao)
                txtProtocolo = IIf(IsNull(TBFI!protocolo), "", TBFI!protocolo)
                txtStatus = IIf(IsNull(TBFI!status), "", TBFI!status)
                
                If IsNull(TBFI!Nosso_Numero) = False And TBFI!Nosso_Numero <> "" Then
                    Txt_nosso_numero = TBFI!Nosso_Numero
                End If
                
            Else
                With Chk_novo
                    .Value = 1
                    .Enabled = False
                End With
                txt_Vencimento = Format(TBContas!Vencimento, "dd/mm/yy")
                Txt_data_doc = Format(Date, "dd/mm/yy")
                Txt_numero_doc = Right(TBContas!NFiscal, 6) & "/" & Left(TBContas!Parcela, 3)
                Txt_data_processamento = Date
                Txt_valor_doc = Format(TBContas!valor, "###,##0.00")
                Txt_valor_cobrado = Format(TBContas!valor, "###,##0.00")
                Set TBAbrir = CreateObject("adodb.recordset")
                
                If Financeiro_Contas_Receber = False Then
                    TextoFiltro = "txt_Agencia = '" & frmFaturamento_Prod_Serv.txt_Agencia & "' and txt_Conta = '" & frmFaturamento_Prod_Serv.txt_Conta & "'"
                Else
                    TextoFiltro = "txt_Descricao = '" & frmContas_Receber.cmbBanco & "' and txt_Agencia = '" & Agencia & "' and txt_Conta = '" & ContaCorrente & "'"
                End If
                
                StrSql = "Select * from tbl_Instituicoes where " & TextoFiltro & " and ID_empresa = " & IDempresa
                
                'Debug.print StrSql
                
                TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    ProcCarregaCarteira
                    Cmb_carteira.ListIndex = 0
                End If
                    ProcGeraNossoNumero
                End If
            TBFI.Close
        End If
        TBContas.Close
    End With
    If Sit_REG = 1 Or Sit_REG = 3 Then
        If ProcVerifCampos(IIf(Sit_REG = 1, True, False), IIf(Sit_REG = 1, False, True)) = False Then Exit Sub
        If Sit_REG = 1 Then ProcEnviarEmail Else ProcRemessa
    End If
Else
    If Formulario = "Faturamento/Nota fiscal/Própria" Then
        Caption = "Administrativo - Faturamento - Nota fiscal - Própria - Boleto"
    ElseIf Formulario = "Faturamento/Nota fiscal/Terceiros" Then
            Caption = "Administrativo - Faturamento - Nota fiscal - Terceiros - Boleto"
        ElseIf Formulario = "Estoque/Ordem de faturamento" Then
                Caption = "Estoque - Ordem de faturamento - Boleto"
            Else
                Caption = "Estoque - Nota fiscal - Boleto"
    End If
    With frmFaturamento_Prod_Serv
        Set TBFI = CreateObject("adodb.recordset")
         StrSql = "Select DR.*, NF.ID_empresa, NF.txt_tipocliente, NF.txt_Razao_Nome, NF.txt_CNPJ_CPF, NF.ID_int_Cliente from tbl_Detalhes_Recebimento DR INNER JOIN tbl_Dados_Nota_Fiscal NF on DR.ID_nota = NF.ID where DR.Id = " & .lst_Duplicata.SelectedItem.ListSubItems(3)
         'Debug.print StrSql
         
        TBFI.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
        If TBFI.EOF = False Then
            IDlista = TBFI!ID
            Set TBAbrir = CreateObject("adodb.recordset")
            StrSql = "Select I.*, IB.AssuntoEmail, IB.Desconto, IB.Dias_Protesto, IB.ID_Instrucoes, IB.Instrucoes_protesto, IB.Juros, IB.Multa from tbl_Instituicoes I inner Join tbl_Instituicoes_Instrucoes_Boleto IB on IB.ID_Instituicao = I.Id where I.txt_Descricao = '" & TBFI!txt_Portador_Banco & "' and ID_empresa = " & IIf(IsNull(TBFI!ID_empresa), 0, TBFI!ID_empresa) & ""
            TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                Familiatext = TBAbrir!int_NBanco
                Img_logo_banco.Picture = LoadPicture(Localrel & "\Imagens\Bancos\" & TBAbrir!int_NBanco & ".jpg")
                Agencia = TBAbrir!txt_Agencia
                ContaCorrente = TBAbrir!txt_Conta
                NomeAgencia = IIf(IsNull(TBAbrir!Nome_agencia), "", TBAbrir!Nome_agencia)
                Txt_assunto = TBAbrir!AssuntoEmail
                Txt_percentual_juros = Format(TBAbrir!Juros, "###,##0.00")
                Txt_percentual_multa = Format(TBAbrir!Multa, "###,##0.00")
                Txt_percentual_desconto = Format(TBAbrir!Desconto, "###,##0.00")
                Txt_dias_protesto = TBAbrir!Dias_Protesto
                Txt_instrucoes = TBAbrir!Instrucoes_protesto
                txtIDIntegracao = IIf(IsNull(TBFI!IDIntegracao), "", TBFI!IDIntegracao)
                txtProtocolo = IIf(IsNull(TBFI!protocolo), "", TBFI!protocolo)
                txtStatus = IIf(IsNull(TBFI!status), "", TBFI!status)
                txt_Vencimento = TBFI!dt_Vencimento
                Txt_codigo_cedente.Text = ContaCorrente
                Txt_agencia_codigo_cedente = TBAbrir!txt_Agencia & "/" & Txt_codigo_cedente
                
                ProcCarregaCarteira
            End If
            
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select Codigo, Razao from Empresa where Codigo = " & IIf(IsNull(TBFI!ID_empresa), 0, TBFI!ID_empresa), Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                Txt_IDempresa = TBAbrir!CODIGO
                Txt_cedente = TBAbrir!Razao
            End If
            
            
            
            'Carrega a carteira utilizada por último
            Set TBCFOP = CreateObject("adodb.recordset")
            TBCFOP.Open "Select Carteira, Carteira1, Assunto, Enviado FROM tbl_Detalhes_Recebimento where txt_Portador_Banco = '" & TBFI!txt_Portador_Banco & "' and txt_Agencia = '" & Agencia & "' and txt_Conta = '" & ContaCorrente & "' and Carteira is not null order by Id desc", Conexao, adOpenKeyset, adLockOptimistic
            If TBCFOP.EOF = False Then
                'Txt_assunto = IIf(IsNull(TBCFOP!Assunto), "", TBCFOP!Assunto)
                If IsNull(TBCFOP!Carteira1) = False And TBCFOP!Carteira1 <> "" Then Cmb_carteira1 = TBCFOP!Carteira1
                If TBCFOP!Carteira <> "" Then Cmb_carteira = TBCFOP!Carteira
                If TBCFOP!Enviado = True Then
                    'Txt_assunto = IIf(IsNull(TBCFOP!Assunto), "", TBCFOP!Assunto)
                Else
                    Set TBCFOP = CreateObject("adodb.recordset")
                    TBCFOP.Open "Select Assunto FROM tbl_Detalhes_Recebimento where txt_Portador_Banco = '" & TBFI!txt_Portador_Banco & "' and txt_Agencia = '" & Agencia & "' and txt_Conta = '" & ContaCorrente & "' and Enviado = 'True' and Assunto IS NOT NULL and Assunto <> N'' order by Nosso_numero desc", Conexao, adOpenKeyset, adLockOptimistic
                    If TBCFOP.EOF = False Then
                        'Txt_assunto = IIf(IsNull(TBCFOP!Assunto), "", TBCFOP!Assunto)
                    End If
                End If
            End If
            
2:
            Txt_sacado = TBFI!txt_Razao_Nome
            txt_CNPJ = TBFI!txt_CNPJ_CPF
            
            If Len(.txttipocliente) = 1 Then TipoFiltro = "F" Else TipoFiltro = "C"
            ProcCarregaComboEndCob "idcliente = " & TBFI!Id_Int_Cliente & " and Tipo = '" & TipoFiltro & "'", True
            
            If IsNull(TBFI!Nosso_Numero) = False And TBFI!Nosso_Numero <> "" Then
                USMsgBox ("Já existe boleto emitido para esta parcela."), vbInformation, "CAPRIND v5.0"
         '       frmFaturamento_Prod_serv_boleto.Height = 10450
                
                Chk_novo.Enabled = True
                Chk_atualizar.Enabled = True
                
                'Verifica se a conta já foi recebida
                Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open "Select IDIntconta from tbl_contas_receber where IDIntconta = " & TBFI!IdContaReceber & " and LogSit = 'S'", Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    USMsgBox ("Esse boleto já foi recebido e não poderá sofrer alterações."), vbExclamation, "CAPRIND v5.0"
                    Chk_novo.Enabled = False
                    Chk_atualizar.Enabled = False
                    Frame2.Enabled = False
                End If
                TBAbrir.Close
                
                If txt_Vencimento = "" Then
                    txt_Vencimento = Format(TBFI!Vencimento_boleto, "dd/mm/yy")
                End If
                
                Txt_data_doc = IIf(IsNull(TBFI!Data_emissao), "", Format(TBFI!Data_emissao, "dd/mm/yy"))
                Txt_numero_doc = IIf(IsNull(TBFI!Numero_documento), "", TBFI!Numero_documento)
                Txt_data_processamento = Txt_data_doc
                Txt_valor_doc = Format(TBFI!Valor_boleto, "###,##0.00")
                Txt_valor_cobrado = Format(TBFI!Valor_cobrado, "###,##0.00")
                Txt_desconto.Text = Format(TBFI!Valor_desconto, "###,##0.00")
                txtDatalimiteDesc.Value = Format(TBFI!DataLimiteDesconto, "dd/mm/yy")
                
                If IsNull(TBFI!Carteira) = False And TBFI!Carteira <> "" Then Cmb_carteira = TBFI!Carteira
                If IsNull(TBFI!Carteira1) = False And TBFI!Carteira1 <> "" And Familiatext = "001" Then Cmb_carteira1 = TBFI!Carteira1
                
                If IsNull(TBFI!Nosso_Numero) = False And TBFI!Nosso_Numero <> "" Then
                Txt_nosso_numero = TBFI!Nosso_Numero
                End If
                
                If IsNull(TBFI!Acrescimos) = False And TBFI!Acrescimos <> "" Then Txt_outros_acrescimos = Format(TBFI!Acrescimos, "###,##0.00")
                If TBFI!Enviado = True Then Chk_email_enviado.Value = 1 Else Chk_email_enviado.Value = 0
                Txt_data_envio = IIf(IsNull(TBFI!data_envio), "", Format(TBFI!data_envio, "dd/mm/yy"))
                If IsNull(TBFI!ID_Cobranca) = False And TBFI!ID_Cobranca <> "" Then ProcCarregaComboEndCob "idcobranca = " & TBFI!ID_Cobranca, False
                txtIDIntegracao = IIf(IsNull(TBFI!IDIntegracao), "", TBFI!IDIntegracao)
                txtProtocolo = IIf(IsNull(TBFI!protocolo), "", TBFI!protocolo)
                txtStatus = IIf(IsNull(TBFI!status), "", TBFI!status)
            Else

                With Chk_novo
                    .Value = 1
                    .Enabled = False
                End With
                txt_Vencimento = Format(TBFI!dt_Vencimento, "dd/mm/yy")
                Txt_data_doc = Format(Date, "dd/mm/yy")
                Txt_numero_doc = Right(TBFI!int_NotaFiscal, 6) & "/" & Left(TBFI!txt_Parcela, 3)
                Txt_data_processamento = Date
                Txt_valor_doc = Format(TBFI!dbl_Valor, "###,##0.00")
                Txt_valor_cobrado = Format(TBFI!dbl_Valor, "###,##0.00")
                
                ProcGeraNossoNumero
                

                StrSql = "Update tbl_Detalhes_Recebimento set Nosso_numero = '" & Txt_nosso_numero.Text & "', txt_agencia = '" & Agencia & "', txt_conta = '" & ContaCorrente & "' where id = " & TBFI!ID
                Conexao.Execute StrSql

            
                txtIDIntegracao = IIf(IsNull(TBFI!IDIntegracao), "", TBFI!IDIntegracao)
                txtProtocolo = IIf(IsNull(TBFI!protocolo), "", TBFI!protocolo)
                txtStatus = IIf(IsNull(TBFI!status), "", TBFI!status)
            End If
        End If
        TBFI.Close
    End With
End If

Cmb_novo_vencimento = IIf(txt_Vencimento.Text <> "", txt_Vencimento.Text, Date)
Cmb_vencimento = IIf(txt_Vencimento.Text <> "", txt_Vencimento.Text, Date)

Exit Sub
tratar_erro:
    If Err.Number = "53" Then
        USMsgBox ("Verifique se o número do banco está cadastrado corretamente."), vbExclamation, "CAPRIND v5.0"
        With USToolBar1
            .ButtonState(1) = 5
            .ButtonState(2) = 5
            .ButtonState(3) = 5
            .ButtonState(4) = 5
            .Refresh
        End With
        Frame2.Enabled = False
        Exit Sub
    End If
    If Err.Number = 383 Then
        If Financeiro_Contas_Receber = True Then GoTo 1 Else GoTo 2
    End If
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaCarteira()
On Error GoTo tratar_erro

Cmb_carteira.Clear
Select Case TBAbrir!int_NBanco
    Case "001": 'Banco do brasil
        With Cmb_carteira
            .AddItem "11 - Simples - Com Registro"
            .AddItem "11 - Vinculada - Com Registro"
            .AddItem "17 - Direta Especial - Com Registro"
            .AddItem "17Simples - Direta Especial Simples - Com Registro"
            .AddItem "17-7 - Direta Especial - Com Registro Convênio 7 dígitos"
            .AddItem "18 - Simples - Sem Registro"
            .AddItem "18-7 - Simples - Sem Registro - Convênio 7 dígitos"
        End With
    Case "033": 'Santander
        With Cmb_carteira
            .AddItem "COB - Cobrança Simples"
            .AddItem "COBR - Cobrança Simples - Rápida Com Registro"
            .AddItem "COBR-Nova - Cobrança Simples - Rápida Com Registro"
            .AddItem "CSR - Cobrança Simples Sem Registro"
            .AddItem "ECR - Cobrança Simples Com Registro"
        End With
    Case "104": 'Caixa
        With Cmb_carteira
            .AddItem "CR - Cobrança Rápida"
            .AddItem "CS - Cobrança Simples"
            .AddItem "SR - Cobrança Sem Registro"
            .AddItem "SR5 - SINCO - Sem Registro"
            .AddItem "SIG14 - SIG Com Registro - Emissão pelo Cedente"
        End With
    Case "237": 'Bradesco
        With Cmb_carteira
'            .AddItem "06 - Sem Registro"
            .AddItem "09 - Com Registro"
        End With
    Case "341": 'Itaú
        With Cmb_carteira
            .AddItem "109 - Direta Eletrônica Sem Emissão - Simples"
            .AddItem "112 - Escritual Eletrônica - simples / contratual"
            .AddItem "175 - Sem Registro Sem Emissão"
        End With
    Case "356": 'ABN e Real
        With Cmb_carteira
            .AddItem "20 - Cobrança Simples"
        End With
    Case "399": 'HSBC
        With Cmb_carteira
            .AddItem "CNR - Sem Registro"
        End With
    Case "409": 'Unibanco
        With Cmb_carteira
            .AddItem "Carnet"
            .AddItem "Caução"
            .AddItem "Desconto"
            .AddItem "DescontoEletronico"
            .AddItem "Direta"
            .AddItem "Escritural"
            .AddItem "Especial"
            .AddItem "Seguro"
            .AddItem "Simples"
            .AddItem "Vendor"
            .AddItem "Vinculada"
        End With
End Select

With Cmb_carteira1
    .Clear
    If TBAbrir!int_NBanco = "001" Then 'Banco do brasil
        .Visible = True
        .AddItem "019"
        .AddItem "027"
        
        Cmb_carteira.Width = 2595
    Else
        .Visible = False
        Cmb_carteira.Width = 3255
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_atualizar_Click()
On Error GoTo tratar_erro

If Chk_atualizar.Value = 1 Then
    Txt_data_doc = Format(Date, "dd/mm/yy")
    Txt_data_processamento = Txt_data_doc
    FrameAtualizacao.Enabled = True
    txt_Vencimento.Visible = True
    Cmb_vencimento.Visible = False
    With Txt_valor_doc
        .Locked = True
        .TabStop = False
    End With
    ProcCarregaNNVctoVlr False
    Frame2.Enabled = True
    Chk_novo.Value = 0
    ProcVerifDiasAtraso
ElseIf Chk_novo.Value = 0 Then
        Cmb_novo_vencimento.Value = IIf(txt_Vencimento.Text <> "", txt_Vencimento.Text, Date)
        Txt_dias_atraso = 0
        Chk_calcular_juros_multa.Value = 0
        FrameAtualizacao.Enabled = False
        
        txt_Vencimento.Visible = True
        Cmb_vencimento.Visible = False
        With Txt_valor_doc
            .Locked = True
            .TabStop = False
        End With
        ProcCarregaNNVctoVlr True
        Frame2.Enabled = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_novo_Click()
On Error GoTo tratar_erro

If Chk_novo.Value = 1 Then
    Txt_data_doc = Format(Date, "dd/mm/yy")
    Txt_data_processamento = Txt_data_doc
    Cmb_novo_vencimento.Value = IIf(txt_Vencimento.Text <> "", txt_Vencimento.Text, Date)
    Txt_dias_atraso = 0
    Chk_calcular_juros_multa.Value = 0
    FrameAtualizacao.Enabled = False
    If Chk_atualizar.Enabled = True Then
        txt_Vencimento.Visible = False
        Cmb_vencimento.Visible = True
        With Txt_valor_doc
            .Locked = False
            .TabStop = True
            If Financeiro_Contas_Receber = True Then .Text = frmContas_Receber.txtValor
        End With
    End If
    ProcGeraNossoNumero
    Frame2.Enabled = True
    Chk_atualizar.Value = 0
ElseIf Chk_atualizar.Value = 0 Then
        Cmb_novo_vencimento.Value = IIf(txt_Vencimento.Text <> "", txt_Vencimento.Text, Date)
        Txt_dias_atraso = 0
        Chk_calcular_juros_multa.Value = 0
        FrameAtualizacao.Enabled = False
        
        txt_Vencimento.Visible = True
        Cmb_vencimento.Visible = False
        With Txt_valor_doc
            .Locked = True
            .TabStop = False
        End With
        ProcCarregaNNVctoVlr True
        Frame2.Enabled = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo tratar_erro

    If Financeiro_Contas_Receber = True Then frmContas_Receber.ProcCarregaLista (1)
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub Txt_desconto_LostFocus()
Dim ValorBoleto As Double
Dim ValorDesconto As Double

If Txt_desconto <> "" Then
    VerifNumero = Txt_desconto
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_desconto.SetFocus
        Exit Sub
    End If

ValorBoleto = Txt_valor_doc
ValorDesconto = Txt_desconto
Txt_desconto = Format(Txt_desconto, "###,##0.00")

txtDatalimiteDesc.Value = txt_Vencimento
Txt_percentual_desconto = (ValorDesconto / ValorBoleto) * 100
Txt_percentual_desconto = Format(Txt_percentual_desconto, "###,##0.00")
Txt_valor_cobrado = ValorBoleto - ValorDesconto
Txt_valor_cobrado = Format(Txt_valor_cobrado, "###,##0.00")
Else
Txt_percentual_desconto = 0
Txt_valor_cobrado = Format(Txt_valor_doc, "###,##0.00")
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub

End Sub

Private Sub Txt_dias_atraso_Change()
On Error GoTo tratar_erro

If Txt_dias_atraso <> "" Then
    VerifNumero = Txt_dias_atraso
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_dias_atraso.SetFocus
        Exit Sub
    End If
End If
ProcCalculaJurosMulta IIf(Txt_dias_atraso = "", 0, Txt_dias_atraso)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_dias_protesto_Change()
On Error GoTo tratar_erro

If Txt_dias_protesto <> "" Then
    VerifNumero = Txt_dias_protesto
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_dias_protesto = ""
        Txt_dias_protesto.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_nosso_numero_Change()
On Error GoTo tratar_erro

If Txt_nosso_numero <> "" Then
    VerifNumero = Txt_nosso_numero
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_nosso_numero = ""
        Txt_nosso_numero.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_numero_doc_GotFocus()
On Error GoTo tratar_erro

If Financeiro_Contas_Receber = True And (Sit_REG = 1 Or Sit_REG = 3) Then Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_outros_acrescimos_Change()
On Error GoTo tratar_erro

If Txt_outros_acrescimos <> "" Then
    VerifNumero = Txt_outros_acrescimos
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_outros_acrescimos = ""
        Txt_outros_acrescimos.SetFocus
        Exit Sub
    End If
End If
valor = IIf(Txt_valor_doc = "", 0, Txt_valor_doc)
Valor1 = IIf(Txt_mora = "", 0, Txt_mora)
Valor2 = IIf(Txt_outros_acrescimos = "", 0, Txt_outros_acrescimos)
Txt_valor_cobrado = Format(valor + Valor1 + Valor2, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_outros_acrescimos_LostFocus()
On Error GoTo tratar_erro

Txt_outros_acrescimos = Format(Txt_outros_acrescimos, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_percentual_juros_Change()
On Error GoTo tratar_erro

If Txt_percentual_juros <> "" Then
    VerifNumero = Txt_percentual_juros
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_percentual_juros = ""
        Txt_percentual_juros.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_percentual_juros_LostFocus()
On Error GoTo tratar_erro

Txt_percentual_juros = Format(Txt_percentual_juros, "###,##0.0000000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_percentual_desconto_Change()
On Error GoTo tratar_erro

If Txt_percentual_desconto <> "" Then
    VerifNumero = Txt_percentual_desconto
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_percentual_desconto = ""
        Txt_percentual_desconto.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_percentual_desconto_LostFocus()
On Error GoTo tratar_erro

Txt_percentual_desconto = Format(Txt_percentual_desconto, "###,##0.0000000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_percentual_multa_Change()
On Error GoTo tratar_erro

If Txt_percentual_multa <> "" Then
    VerifNumero = Txt_percentual_multa
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_percentual_multa = ""
        Txt_percentual_multa.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_percentual_multa_LostFocus()
On Error GoTo tratar_erro

Txt_percentual_multa = Format(Txt_percentual_multa, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_valor_doc_Change()
On Error GoTo tratar_erro

Txt_valor_cobrado = Txt_valor_doc
If Txt_valor_doc <> "" Then
    VerifNumero = Txt_valor_doc
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_valor_doc.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_valor_doc_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus Txt_valor_doc

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_valor_doc_LostFocus()
On Error GoTo tratar_erro

Txt_valor_doc = Format(Txt_valor_doc, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcStatusBoleto()
On Error GoTo tratar_erro

If txtEmissor.Text = "Tecnospeed" Then
   
   If txtStatus <> "EMITIDO" Or txtStatus = "" Then
        Do While txtStatus <> "EMITIDO" Or txtStatus = ""
        If Sit_REG = 1 Or txtStatus = "FALHA" Then Exit Sub
            PlugEmitirBoleto
        Loop
    End If
    
    If txtIDIntegracao <> "" Then
        Do While txtStatus <> "EMITIDO"
          txtStatus = PlugConsultarBoleto(txtIDIntegracao)
          txtStatus = Mensagem
        Loop
    End If
    
    If txtProtocolo.Text = "" And txtIDIntegracao <> "" Then
        Do While txtProtocolo = ""
            txtProtocolo.Text = PlugGerarProtocoloBoleto(txtIDIntegracao)
        Loop
    End If
    

End If

If txtIDIntegracao <> "" Then

    ProcGravarDadosBoleto

End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtIDIntegracao_Change()
On Error GoTo tratar_erro

If txtIDIntegracao <> "" Then
    StrSql = "Update tbl_Detalhes_Recebimento set Protocolo = '" & txtProtocolo.Text & "' where IDIntegracao = '" & txtIDIntegracao.Text & "'"
    Conexao.Execute StrSql
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtProtocolo_Change()
On Error GoTo tratar_erro

If txtIDIntegracao <> "" Then
    StrSql = "Update tbl_Detalhes_Recebimento set Protocolo = '" & txtProtocolo.Text & "' where IDIntegracao = '" & txtIDIntegracao.Text & "'"
    Conexao.Execute StrSql
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcImprimir
    Case 2: ProcRemessa
    Case 3:
        If ProcVerifCampos(True, False) = False Then Exit Sub
        ProcEnviarEmail
    'Case 5: ProcAjuda
    Case 6: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaNNVctoVlr(CarregarDatas As Boolean)
On Error GoTo tratar_erro

Set TBFI = CreateObject("adodb.recordset")
If Financeiro_Contas_Receber = True Then
    TBFI.Open "Select DR.Data_emissao, DR.Nosso_numero, DR.dt_Vencimento, DR.dbl_Valor, DR.Vencimento_boleto, DR.Valor_boleto, DR.Acrescimos from tbl_Detalhes_Recebimento DR INNER JOIN tbl_contas_receber CR ON CR.IDIntconta = DR.IDContaReceber AND CR.Banco = DR.txt_Portador_Banco where DR.IDContaReceber = " & frmContas_Receber.txtidintconta & " and DR.Nosso_numero IS NOT NULL and DR.Nosso_numero <> N''", Conexao, adOpenKeyset, adLockOptimistic
Else
    TBFI.Open "Select Data_emissao, Nosso_numero, dt_Vencimento, dbl_Valor, Vencimento_boleto, Valor_boleto, Acrescimos from tbl_Detalhes_Recebimento where Id = " & frmFaturamento_Prod_Serv.lst_Duplicata.SelectedItem.ListSubItems(3) & " and Nosso_numero IS NOT NULL and Nosso_numero <> N''", Conexao, adOpenKeyset, adLockOptimistic
End If
If TBFI.EOF = False Then
    Txt_nosso_numero = TBFI!Nosso_Numero
    txt_Vencimento = Format(IIf(Chk_atualizar.Value = 1, TBFI!dt_Vencimento, TBFI!Vencimento_boleto), "dd/mm/yy")
    Txt_valor_doc = Format(IIf(Chk_atualizar.Value = 1, TBFI!dbl_Valor, TBFI!Valor_boleto), "###,##0.00")
    Txt_valor_cobrado = Format(IIf(Chk_atualizar.Value = 1, TBFI!dbl_Valor, TBFI!Valor_boleto), "###,##0.00")
    txt_Txt_outros_acrescimos = Format(IIf(Chk_atualizar.Value = 1, TBFI!Acrescimos, 0), "###,##0.00")
    
    If CarregarDatas = True Then
        Txt_data_doc = IIf(IsNull(TBFI!Data_emissao), "", Format(TBFI!Data_emissao, "dd/mm/yy"))
        Txt_data_processamento = Txt_data_doc
    End If
End If
TBFI.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaComboEndCob(TextoFiltro As String, CarregaTodos As Boolean)
On Error GoTo tratar_erro

With Cmb_endereco
    If CarregaTodos = True Then .Clear
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from clientes_cobranca where " & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Do While TBAbrir.EOF = False
            If IsNull(TBAbrir!Tipo_endereco) = False And TBAbrir!Tipo_endereco <> "" Then
                EnderecoBoleto = TBAbrir!Tipo_endereco & ": " & IIf(IsNull(TBAbrir!endereco_Cobranca), "", Trim(Left(TBAbrir!endereco_Cobranca, 28))) & ", " & IIf(IsNull(TBAbrir!Numero), "", TBAbrir!Numero)
                SacadoEnderecoLogradouro = EnderecoBoleto
                SacadoEnderecoNumero = TBAbrir!Numero
                SacadoEnderecoBairro = TBAbrir!bairro_Cobranca
                SacadoEnderecoCEP = TBAbrir!cep_Cobranca
            Else
                EnderecoBoleto = IIf(IsNull(TBAbrir!endereco_Cobranca), "", Trim(Left(TBAbrir!endereco_Cobranca, 28))) & ", " & IIf(IsNull(TBAbrir!Numero), "", TBAbrir!Numero)
                SacadoEnderecoLogradouro = EnderecoBoleto
                SacadoEnderecoNumero = TBAbrir!Numero
                SacadoEnderecoBairro = TBAbrir!bairro_Cobranca
                SacadoEnderecoCEP = TBAbrir!cep_Cobranca
            End If
            If IsNull(TBAbrir!Tipo_bairro) = False And TBAbrir!Tipo_bairro <> "" Then
                BairroBoleto = TBAbrir!Tipo_bairro & ": " & IIf(IsNull(TBAbrir!bairro_Cobranca), "", Trim(TBAbrir!bairro_Cobranca))
            Else
                BairroBoleto = IIf(IsNull(TBAbrir!bairro_Cobranca), "", TBAbrir!bairro_Cobranca)
            End If
            CidadeBoleto = IIf(IsNull(TBAbrir!cidade_Cobranca), "", TBAbrir!cidade_Cobranca)
            EstadoBoleto = IIf(IsNull(TBAbrir!uf_Cobranca), "", TBAbrir!uf_Cobranca)
            CEPBoleto = ReturnNumbersOnly(IIf(IsNull(TBAbrir!cep_Cobranca), "", TBAbrir!cep_Cobranca))
            
            If CarregaTodos = True Then
                .AddItem EnderecoBoleto & " - " & BairroBoleto & " - " & CidadeBoleto & " - " & EstadoBoleto & " - " & CEPBoleto
                .ItemData(.NewIndex) = TBAbrir!idCobranca
            Else
                .Text = EnderecoBoleto & " - " & BairroBoleto & " - " & CidadeBoleto & " - " & EstadoBoleto & " - " & CEPBoleto
            End If
            TBAbrir.MoveNext
        Loop
        
        If CarregaTodos = True Then
            TBAbrir.MoveFirst
            .Text = EnderecoBoleto & " - " & BairroBoleto & " - " & CidadeBoleto & " - " & EstadoBoleto & " - " & CEPBoleto
        End If
    End If
    TBAbrir.Close
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

