VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2014.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmproj_EstruturaLocaliza_itemOLD 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Engenharia - Estrutura - Novo/alterar registro"
   ClientHeight    =   9510
   ClientLeft      =   1890
   ClientTop       =   1365
   ClientWidth     =   10215
   ClipControls    =   0   'False
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
   Icon            =   "frmproj_estrutura_localizaItemOLD.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9510
   ScaleWidth      =   10215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   885
      Left            =   55
      TabIndex        =   40
      Top             =   8580
      Width           =   10125
      Begin VB.TextBox Txt_percenual_perda 
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   8820
         MaxLength       =   50
         TabIndex        =   33
         Text            =   "0,0000"
         ToolTipText     =   "Percentual de perda."
         Top             =   420
         Width           =   1035
      End
      Begin VB.CommandButton cmdPesoBruto 
         BackColor       =   &H00C0C0C0&
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
         Left            =   1710
         Picture         =   "frmproj_estrutura_localizaItemOLD.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Carregar peso bruto do produto principal."
         Top             =   420
         Width           =   315
      End
      Begin VB.ComboBox cmbunkg 
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
         ItemData        =   "frmproj_estrutura_localizaItemOLD.frx":03EC
         Left            =   2370
         List            =   "frmproj_estrutura_localizaItemOLD.frx":03FC
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   28
         TabStop         =   0   'False
         ToolTipText     =   "Unidade por kilograma."
         Top             =   420
         Width           =   1095
      End
      Begin VB.TextBox txtpesototal 
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   7440
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   32
         TabStop         =   0   'False
         Text            =   "0,00000"
         ToolTipText     =   "Peso total."
         Top             =   420
         Width           =   1365
      End
      Begin VB.TextBox txtkgpc 
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
         Left            =   4800
         Locked          =   -1  'True
         MaxLength       =   50
         MousePointer    =   99  'Custom
         TabIndex        =   30
         TabStop         =   0   'False
         Text            =   "0,00000"
         ToolTipText     =   "Peso por peça."
         Top             =   420
         Width           =   1335
      End
      Begin VB.TextBox txtdimensao 
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
         Left            =   3480
         MaxLength       =   30
         TabIndex        =   29
         Text            =   "0,00000"
         ToolTipText     =   "Dimensão a ser utilizada por peça."
         Top             =   420
         Width           =   1305
      End
      Begin VB.TextBox txtpeso 
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   270
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   26
         TabStop         =   0   'False
         Text            =   "0,00000"
         ToolTipText     =   "Kilograma por unidade."
         Top             =   420
         Width           =   1410
      End
      Begin VB.TextBox txtquantidade 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   6150
         MaxLength       =   50
         TabIndex        =   31
         Text            =   "0,00000"
         ToolTipText     =   "Quantidade."
         Top             =   420
         Width           =   1275
      End
      Begin VB.TextBox txtVT 
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
         Left            =   7440
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   35
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor total."
         Top             =   420
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.TextBox cmbVU 
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   4950
         MaxLength       =   50
         TabIndex        =   34
         Text            =   "0,00000"
         ToolTipText     =   "Valor unitário."
         Top             =   420
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "% perda"
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
         Left            =   8970
         TabIndex        =   66
         Top             =   210
         Width           =   735
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Un/Kg*"
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
         Left            =   2602
         TabIndex        =   47
         Top             =   210
         Width           =   630
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Peso total"
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
         Left            =   7695
         TabIndex        =   46
         Top             =   210
         Width           =   855
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kg/pç"
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
         Left            =   5220
         TabIndex        =   45
         Top             =   210
         Width           =   495
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Dim. / mm*"
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
         Left            =   3480
         TabIndex        =   44
         Top             =   210
         Width           =   1305
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Left            =   2130
         TabIndex        =   43
         Top             =   540
         Width           =   105
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kg/unidade*"
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
         Left            =   435
         TabIndex        =   42
         Top             =   210
         Width           =   1080
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Quant.*"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   6412
         TabIndex        =   41
         Top             =   210
         Width           =   750
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   2340
      Left            =   55
      TabIndex        =   39
      Top             =   6240
      Width           =   10125
      Begin VB.TextBox Txt_obs 
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
         Height          =   615
         Left            =   180
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   25
         ToolTipText     =   "Observações."
         Top             =   1575
         Width           =   9750
      End
      Begin VB.TextBox Txt_posicao 
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
         Left            =   880
         MaxLength       =   3
         TabIndex        =   18
         ToolTipText     =   "Posição."
         Top             =   370
         Width           =   600
      End
      Begin VB.ComboBox cmbVersao 
         Appearance      =   0  'Flat
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
         ItemData        =   "frmproj_estrutura_localizaItemOLD.frx":0414
         Left            =   180
         List            =   "frmproj_estrutura_localizaItemOLD.frx":0466
         Style           =   2  'Dropdown List
         TabIndex        =   17
         ToolTipText     =   "Versão."
         Top             =   370
         Width           =   705
      End
      Begin VB.TextBox txtcodigo 
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
         MouseIcon       =   "frmproj_estrutura_localizaItemOLD.frx":04B8
         MousePointer    =   99  'Custom
         TabIndex        =   54
         Text            =   "0"
         ToolTipText     =   "Código interno."
         Top             =   990
         Visible         =   0   'False
         Width           =   720
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2070
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   23
         TabStop         =   0   'False
         ToolTipText     =   "Descrição."
         Top             =   990
         Width           =   7415
      End
      Begin VB.TextBox txtun 
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
         Left            =   9495
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   24
         TabStop         =   0   'False
         ToolTipText     =   "Unidade."
         Top             =   990
         Width           =   435
      End
      Begin VB.ComboBox cmbcodref 
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
         ItemData        =   "frmproj_estrutura_localizaItemOLD.frx":07C2
         Left            =   3525
         List            =   "frmproj_estrutura_localizaItemOLD.frx":07C4
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   20
         ToolTipText     =   "Código de referência."
         Top             =   375
         Width           =   2010
      End
      Begin VB.TextBox Txt_familia 
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
         Left            =   5535
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   "Família."
         Top             =   375
         Width           =   4395
      End
      Begin VB.TextBox txtDesenho 
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
         Left            =   1500
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "Código interno."
         Top             =   370
         Width           =   2010
      End
      Begin VB.ComboBox Cmb_part_number_fabricante 
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
         ItemData        =   "frmproj_estrutura_localizaItemOLD.frx":07C6
         Left            =   180
         List            =   "frmproj_estrutura_localizaItemOLD.frx":07C8
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   22
         ToolTipText     =   "Part number do fabricante."
         Top             =   990
         Width           =   1890
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Observações"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   41
         Left            =   4583
         TabIndex        =   68
         Top             =   1380
         Width           =   945
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Part number"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   683
         TabIndex        =   67
         Top             =   780
         Width           =   885
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Pos.*"
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
         Left            =   948
         TabIndex        =   65
         Top             =   180
         Width           =   465
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Versão*"
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
         TabIndex        =   53
         Top             =   180
         Width           =   705
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
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
         Left            =   5365
         TabIndex        =   52
         Top             =   780
         Width           =   825
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
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
         Left            =   1838
         TabIndex        =   51
         Top             =   180
         Width           =   1335
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Un."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9585
         TabIndex        =   50
         Top             =   780
         Width           =   255
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Família"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   7492
         TabIndex        =   49
         Top             =   180
         Width           =   480
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Cód. de referência"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3855
         TabIndex        =   48
         Top             =   180
         Width           =   1350
      End
   End
   Begin VB.CheckBox chkFiltrarEstoque 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Filtrar do estoque"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   1515
      Left            =   60
      TabIndex        =   60
      Top             =   1290
      Width           =   10125
      Begin VB.Frame Frame11 
         BackColor       =   &H00E0E0E0&
         Height          =   510
         Left            =   5160
         TabIndex        =   69
         Top             =   210
         WhatsThisHelpID =   210
         Width           =   4785
         Begin VB.OptionButton optIgual 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Igual"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3930
            TabIndex        =   16
            Top             =   180
            Width           =   705
         End
         Begin VB.OptionButton Optmeio 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Meio frase"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1470
            TabIndex        =   14
            Top             =   180
            Width           =   1275
         End
         Begin VB.OptionButton Optinicio 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Início frase"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   180
            TabIndex        =   13
            Top             =   180
            Value           =   -1  'True
            Width           =   1275
         End
         Begin VB.OptionButton Optfim 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Fim frase"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2760
            TabIndex        =   15
            Top             =   180
            Width           =   1155
         End
      End
      Begin VB.CommandButton Cmd_salvar 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   6930
         Picture         =   "frmproj_estrutura_localizaItemOLD.frx":07CA
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Salvar filtro para pesquisa (F7)."
         Top             =   1050
         Width           =   315
      End
      Begin VB.CommandButton Cmd_excluir 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   7260
         Picture         =   "frmproj_estrutura_localizaItemOLD.frx":081D
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Excluir filtro para pesquisa (F4)."
         Top             =   1050
         Width           =   315
      End
      Begin VB.TextBox txtTexto 
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
         TabIndex        =   2
         ToolTipText     =   "Texto para pesquisa."
         Top             =   1050
         Width           =   6735
      End
      Begin VB.ComboBox cmbfiltrarpor 
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
         ItemData        =   "frmproj_estrutura_localizaItemOLD.frx":095B
         Left            =   180
         List            =   "frmproj_estrutura_localizaItemOLD.frx":0986
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "Opções para filtro."
         Top             =   390
         Width           =   4875
      End
      Begin VB.ComboBox Cmb_ordenar 
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
         ItemData        =   "frmproj_estrutura_localizaItemOLD.frx":0A37
         Left            =   7680
         List            =   "frmproj_estrutura_localizaItemOLD.frx":0A41
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         ToolTipText     =   "Ordenar por."
         Top             =   1050
         Width           =   2265
      End
      Begin VB.ComboBox cmbfamilia 
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
         TabIndex        =   3
         ToolTipText     =   "Texto para pesquisa."
         Top             =   1050
         Visible         =   0   'False
         Width           =   6735
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filtrar por"
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
         Left            =   2197
         TabIndex        =   63
         Top             =   180
         Width           =   840
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Texto para pesquisa"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2812
         TabIndex        =   62
         Top             =   840
         Width           =   1470
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ordenar por"
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
         Left            =   8302
         TabIndex        =   61
         Top             =   840
         Width           =   1020
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   60
      TabIndex        =   56
      Top             =   5340
      Width           =   10125
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
         Left            =   4980
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
         Left            =   2670
         TabIndex        =   6
         Text            =   "30"
         ToolTipText     =   "Número de registros por página."
         Top             =   180
         Width           =   555
      End
      Begin DrawSuite2014.USButton cmdPagProx 
         Height          =   315
         Left            =   7200
         TabIndex        =   11
         ToolTipText     =   "Próxima página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmproj_estrutura_localizaItemOLD.frx":0A60
         BorderColor     =   14404026
         BorderColorDown =   11632444
         BorderColorOver =   11632444
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
         GradientColor2  =   16777215
         GradientColor3  =   16777215
         GradientColorDown2=   16246986
         GradientColorDown3=   15189380
         GradientColorDown4=   14596208
         GradientColorOver1=   16643560
         GradientColorOver2=   16576988
         GradientColorOver3=   16441780
         GradientColorOver4=   16178091
         PicSizeH        =   19
         PicSizeW        =   19
      End
      Begin DrawSuite2014.USButton cmdPagAnt 
         Height          =   315
         Left            =   6660
         TabIndex        =   10
         ToolTipText     =   "Página anterior."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmproj_estrutura_localizaItemOLD.frx":4207
         BorderColor     =   14404026
         BorderColorDown =   11632444
         BorderColorOver =   11632444
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
         GradientColor2  =   16777215
         GradientColor3  =   16777215
         GradientColorDown2=   16246986
         GradientColorDown3=   15189380
         GradientColorDown4=   14596208
         GradientColorOver1=   16643560
         GradientColorOver2=   16576988
         GradientColorOver3=   16441780
         GradientColorOver4=   16178091
         PicSizeH        =   19
         PicSizeW        =   19
      End
      Begin DrawSuite2014.USButton cmdPagIr 
         Height          =   315
         Left            =   5550
         TabIndex        =   8
         Top             =   180
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   556
         BorderColor     =   14404026
         BorderColorDown =   11632444
         BorderColorOver =   11632444
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
         GradientColor2  =   16777215
         GradientColor3  =   16777215
         GradientColorDown2=   16246986
         GradientColorDown3=   15189380
         GradientColorDown4=   14596208
         GradientColorOver1=   16643560
         GradientColorOver2=   16576988
         GradientColorOver3=   16441780
         GradientColorOver4=   16178091
         PicSizeH        =   19
         PicSizeW        =   19
      End
      Begin DrawSuite2014.USButton cmdPagPrim 
         Height          =   315
         Left            =   6120
         TabIndex        =   9
         ToolTipText     =   "Primeira página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmproj_estrutura_localizaItemOLD.frx":7D11
         BorderColor     =   14404026
         BorderColorDown =   11632444
         BorderColorOver =   11632444
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
         GradientColor2  =   16777215
         GradientColor3  =   16777215
         GradientColorDown2=   16246986
         GradientColorDown3=   15189380
         GradientColorDown4=   14596208
         GradientColorOver1=   16643560
         GradientColorOver2=   16576988
         GradientColorOver3=   16441780
         GradientColorOver4=   16178091
         PicSizeH        =   19
         PicSizeW        =   19
      End
      Begin DrawSuite2014.USButton cmdPagUlt 
         Height          =   315
         Left            =   7740
         TabIndex        =   12
         ToolTipText     =   "Última página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmproj_estrutura_localizaItemOLD.frx":BE02
         BorderColor     =   14404026
         BorderColorDown =   11632444
         BorderColorOver =   11632444
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
         GradientColor2  =   16777215
         GradientColor3  =   16777215
         GradientColorDown2=   16246986
         GradientColorDown3=   15189380
         GradientColorDown4=   14596208
         GradientColorOver1=   16643560
         GradientColorOver2=   16576988
         GradientColorOver3=   16441780
         GradientColorOver4=   16178091
         PicSizeH        =   19
         PicSizeW        =   19
      End
      Begin VB.Label lblPaginas 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Página: 0 de: 0"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   8520
         TabIndex        =   59
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
         TabIndex        =   58
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Carregar               registros por página"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   1980
         TabIndex        =   57
         Top             =   240
         Width           =   2760
      End
   End
   Begin DrawSuite2014.USImageList USImageList1 
      Left            =   8850
      Top             =   240
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frmproj_estrutura_localizaItemOLD.frx":F690
      Count           =   1
   End
   Begin DrawSuite2014.USToolBar USToolBar1 
      Height          =   990
      Left            =   60
      TabIndex        =   55
      Top             =   0
      Width           =   10125
      _ExtentX        =   17859
      _ExtentY        =   1746
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
         Weight          =   400
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft2     =   40
      ButtonTop2      =   2
      ButtonWidth2    =   38
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
      ButtonLeft3     =   80
      ButtonTop3      =   4
      ButtonWidth3    =   2
      ButtonHeight3   =   55
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft4     =   84
      ButtonTop4      =   2
      ButtonWidth4    =   36
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft5     =   122
      ButtonTop5      =   2
      ButtonWidth5    =   26
      ButtonHeight5   =   21
      ButtonUseMaskColor5=   0   'False
      ButtonEnabled6  =   0   'False
      ButtonIconSize6 =   32
      ButtonKey6      =   "6"
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
      ButtonState6    =   5
      ButtonLeft6     =   150
      ButtonTop6      =   2
      ButtonWidth6    =   24
      ButtonHeight6   =   24
      ButtonUseMaskColor6=   0   'False
   End
   Begin DrawSuite2014.USProgressBar PBLista 
      Height          =   255
      Left            =   60
      TabIndex        =   64
      Top             =   5970
      Width           =   10125
      _ExtentX        =   17859
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
   Begin MSComctlLib.ListView ListView1 
      Height          =   2475
      Left            =   60
      TabIndex        =   5
      Top             =   2820
      Width           =   10125
      _ExtentX        =   17859
      _ExtentY        =   4366
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Text            =   "Cód."
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "N"
         Text            =   "RE"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Cód. interno"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "Descrição"
         Object.Width           =   9673
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Un."
         Object.Width           =   970
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Tag             =   "T"
         Text            =   "Família"
         Object.Width           =   4410
      EndProperty
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   9465
      Left            =   10200
      TabIndex        =   38
      Top             =   0
      Visible         =   0   'False
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   16695
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
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "T"
         Text            =   "Filtrar por"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Local da frase"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "Texto para pesquisa"
         Object.Width           =   2866
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Tag             =   "N"
         Text            =   "IDTexto"
         Object.Width           =   0
      EndProperty
   End
End
Attribute VB_Name = "frmproj_EstruturaLocaliza_itemOLD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StrSqlLocProdPadrao As String 'OK

Sub ProcPuxaDados()
On Error GoTo tratar_erro

VersaoEstrutura = TBAbrir!Versao
If IsNull(TBAbrir!Versao_desenho) = False And TBAbrir!Versao <> "" Then cmbVersao = TBAbrir!Versao_desenho
Txt_posicao = IIf(IsNull(TBAbrir!Posicao), "", TBAbrir!Posicao)
txtDesenho.Text = IIf(IsNull(TBAbrir!Desenho), "", TBAbrir!Desenho)
If IsNull(TBAbrir!ID_partnumber_fabricante) = False Then
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select Part_number from Projproduto_fabricante where ID = " & TBAbrir!ID_partnumber_fabricante, Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then Cmb_part_number_fabricante = TBProduto!Part_number
    TBProduto.Close
End If
txtUN.Text = IIf(IsNull(TBAbrir!Unidade), "", TBAbrir!Unidade)

Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select classe from Projproduto where desenho = '" & TBAbrir!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    Txt_familia = IIf(IsNull(TBProduto!classe), "", TBProduto!classe)
End If
TBProduto.Close

txtdescricao.Text = IIf(IsNull(TBAbrir!Descricao), "", TBAbrir!Descricao)
Txt_obs = IIf(IsNull(TBAbrir!Obs), "", TBAbrir!Obs)
txtdimensao.Text = IIf(IsNull(TBAbrir!Dimensoes), "0,00000", Format(TBAbrir!Dimensoes, "###,##0.0000000000"))
txtquantidade.Text = IIf(IsNull(TBAbrir!Quantidade), "0,00000", Format(TBAbrir!Quantidade, "###,##0.0000000000"))
txtkgpc.Text = IIf(IsNull(TBAbrir!Peso), "0,00000", Format(TBAbrir!Peso, "###,##0.0000000000"))
txtpeso.Text = IIf(IsNull(TBAbrir!PesoMetro), "0,00000", Format(TBAbrir!PesoMetro, "###,##0.0000000000"))
txtpesototal.Text = IIf(IsNull(TBAbrir!PesoTotal), "0,00000", Format(TBAbrir!PesoTotal, "###,##0.0000000000"))
Txt_percenual_perda = IIf(IsNull(TBAbrir!Percentual_perda), "0,0000", Format(TBAbrir!Percentual_perda, "###,##0.0000"))
cmbVU.Text = IIf(IsNull(TBAbrir!Valor), "0,00000", Format(TBAbrir!Valor, "###,##0.0000000000"))
txtVT.Text = IIf(IsNull(TBAbrir!ValorTotal), "0,00", Format(TBAbrir!ValorTotal, "###,##0.00"))
If IsNull(TBAbrir!Un_Kg) = False Then cmbunkg.Text = TBAbrir!Un_Kg

IDpedido = TBAbrir!CODIGO

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Sub ProcPuxaDados_PCP()
On Error GoTo tratar_erro

ID_Familia = TBAbrir!IdMateriaPrima
VersaoEstrutura = TBAbrir!Versao
With cmbVersao
    If IsNull(TBAbrir!Versao) = False And TBAbrir!Versao <> "" Then .Text = TBAbrir!Versao
    .Locked = True
    .TabStop = False
End With
Txt_posicao = IIf(IsNull(TBAbrir!Posicao), "", TBAbrir!Posicao)
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select Desenho, classe, un_kg, Unidade, classe, Descricao, peso_metro, PCusto from Projproduto where desenho = '" & TBAbrir!CODIGO & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    txtDesenho = IIf(IsNull(TBProduto!Desenho), "", TBProduto!Desenho)
    cmbcodref.Refresh
    txtUN.Text = IIf(IsNull(TBProduto!Unidade), "", TBProduto!Unidade)
    Txt_familia = IIf(IsNull(TBProduto!classe), "", TBProduto!classe)
    txtdescricao.Text = IIf(IsNull(TBProduto!Descricao), "", TBProduto!Descricao)
    txtpeso.Text = IIf(IsNull(TBProduto!peso_metro), "", Format(TBProduto!peso_metro, "###,##0.0000"))
    If IsNull(TBProduto!Un_Kg) = False Then cmbunkg = TBProduto!Un_Kg
    If TBProduto!PCusto <> "" And TBProduto!PCusto <> 0 Then cmbVU = Format(TBProduto!PCusto, "###,##0.0000000000") Else cmbVU = 0
End If
TBProduto.Close

txtdimensao.Text = IIf(IsNull(TBAbrir!Dimensao), "0,00000", Format(TBAbrir!Dimensao, "###,##0.0000000000"))
Qtde = frmprod.txtquantidade
Valor1 = IIf(IsNull(TBAbrir!Quantidade), 0, TBAbrir!Quantidade)
txtquantidade = Format(Valor1 / Qtde, "###,##0.0000000000")

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Sub ProcCarregaDados()
On Error GoTo tratar_erro

txtDesenho = IIf(IsNull(TBProduto!Desenho), "", TBProduto!Desenho)
cmbcodref.Refresh
txtUN.Text = IIf(IsNull(TBProduto!Unidade), "", TBProduto!Unidade)
Txt_familia = IIf(IsNull(TBProduto!classe), "", TBProduto!classe)
txtdescricao.Text = IIf(IsNull(TBProduto!Descricao), "", TBProduto!Descricao)
txtpeso.Text = IIf(IsNull(TBProduto!peso_metro), "", Format(TBProduto!peso_metro, "###,##0.0000"))
If IsNull(TBProduto!Un_Kg) = False Then cmbunkg = TBProduto!Un_Kg
If TBProduto!PCusto <> "" And TBProduto!PCusto <> 0 Then cmbVU = Format(TBProduto!PCusto, "###,##0.0000000000") Else cmbVU = 0
txtquantidade.Text = "1,00000"
   
Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub chkFiltrarEstoque_Click()
On Error GoTo tratar_erro

cmbfiltrarpor.Clear
ListView1.ListItems.Clear
With cmbfiltrarpor
    .AddItem "Cliente"
    .AddItem "Código de referência"
    .AddItem "Código interno"
    .AddItem "Comprimento"
    .AddItem "Descrição"
    .AddItem "Descrição Comercial"
    .AddItem "Dureza"
    .AddItem "Espessura"
    .AddItem "Família"
    .AddItem "Fornecedor"
    .AddItem "Largura"
    .AddItem "Número do desenho"
End With
If chkFiltrarEstoque.Value = 1 Then
    ListView1.ColumnHeaders(2).Width = 700
    ListView1.ColumnHeaders(4).Width = 4784
    cmbfiltrarpor.AddItem "RE"
    cmbfiltrarpor.AddItem "Lote"
Else
    ListView1.ColumnHeaders(2).Width = 0
    ListView1.ColumnHeaders(4).Width = 5484
End If
cmbfiltrarpor = "Código interno"

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub cmbVersao_Change()
On Error GoTo tratar_erro

ProcCarregaPosicao

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub cmbVersao_Click()
On Error GoTo tratar_erro

ProcCarregaPosicao

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcCarregaPosicao()
On Error GoTo tratar_erro

Set TBFIltro = CreateObject("adodb.recordset")
If PCP_Ordem = True Then
    If ID_Familia <> "0" Then TextoFiltro = "and IdMateriaPrima = " & ID_Familia
    TBFIltro.Open "Select Posicao from Producaomaterial where Ordem = " & frmprod.txtof & " and Versao = '" & cmbVersao & "' and Posicao IS NOT NULL " & TextoFiltro & " order by Posicao desc", Conexao, adOpenKeyset, adLockOptimistic
    If TBFIltro.EOF = False Then
        If ID_Familia = "0" Then Txt_posicao = TBFIltro!Posicao + 1 Else Txt_posicao = TBFIltro!Posicao
    Else
        Txt_posicao = 1
    End If
Else
    With frmproj_produto_estrutura
        If .Novo_Estrutura = True Then TextoFiltro = "codproduto = " & IDlista Else TextoFiltro = "codproduto = " & IDAntigo
        TBFIltro.Open "Select Posicao from ProjConjunto where " & TextoFiltro & " and Versao = '" & .VersaoEstrutura & "' and Desenho = '" & txtDesenho & "' and Versao_desenho = '" & cmbVersao & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBFIltro.EOF = True Then
    
            Set TBFIltro = CreateObject("adodb.recordset")
            TBFIltro.Open "Select Posicao from ProjConjunto where " & TextoFiltro & " and Versao = '" & .VersaoEstrutura & "' and Posicao IS NOT NULL order by Posicao desc", Conexao, adOpenKeyset, adLockOptimistic
            If TBFIltro.EOF = False Then
                Txt_posicao = TBFIltro!Posicao + 1
            Else
                Txt_posicao = 1
            End If
        Else
            Txt_posicao = TBFIltro!Posicao
        End If
    End With
End If
TBFIltro.Close

IDpedido = 0
Set TBItem = CreateObject("adodb.recordset")
TBItem.Open "Select CODIGO from ProjConjunto where Codproduto = " & IDlista & " and Versao = '" & frmproj_produto_estrutura.VersaoEstrutura & "' and Desenho = '" & txtDesenho & "' and Versao_desenho = '" & cmbVersao & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBItem.EOF = False Then
    IDpedido = TBItem!CODIGO
End If
TBItem.Close

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Cmd_excluir_Click()
On Error GoTo tratar_erro

Permitido = False
Inicio:
    With Lista
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                If Permitido = False Then
                    If MsgBox("Deseja realmente excluir este(s) filtro(s) para pesquisa?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
                End If
                Permitido = True
                .ListItems.Remove (InitFor)
                GoTo Inicio
            End If
        Next InitFor
    End With
    If Permitido = False Then
        MsgBox ("Informe o(s) filtro(s) para pesquisa antes de excluir."), vbExclamation
    Else
        MsgBox ("Filtro(s) para pesquisa excluído(s) com sucesso."), vbInformation
        If Lista.ListItems.Count = 0 Then
            Lista.Visible = False
            Me.Width = 10305
        End If
    End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Cmd_salvar_Click()
On Error GoTo tratar_erro

If txtTexto.Visible = True And txtTexto = "" Or cmbFamilia.Visible = True And cmbFamilia = "" Then
    MsgBox ("Informe o texto para pesquisa antes de adicionar o filtro na lista."), vbExclamation
    If txtTexto.Visible = True Then txtTexto.SetFocus Else cmbFamilia.SetFocus
    Exit Sub
End If

With Lista.ListItems
    .Add , , ""
    .Item(.Count).SubItems(1) = cmbfiltrarpor.Text
    If optInicio.Value = True Then .Item(.Count).SubItems(2) = "Início"
    If optMeio.Value = True Then .Item(.Count).SubItems(2) = "Meio"
    If optFim.Value = True Then .Item(.Count).SubItems(2) = "Fim"
    If txtTexto.Visible = True Then
        .Item(.Count).SubItems(3) = txtTexto
    Else
        .Item(.Count).SubItems(3) = cmbFamilia.Text
        .Item(.Count).SubItems(4) = cmbFamilia.ItemData(cmbFamilia.ListIndex)
    End If
End With
Lista.Visible = True
Me.Width = 14500

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub cmdPesoBruto_Click()
On Error GoTo tratar_erro

If txtDesenho = "" Then Exit Sub

If MsgBox("Deseja realmente carregar o peso bruto do produto principal?", vbQuestion + vbYesNo) = vbYes Then
    
    If frmproj_produto_estrutura.Novo_Estrutura = True Then
        TextoFiltro = "where P.codproduto = " & IDlista
    Else
        TextoFiltro = "INNER JOIN projconjunto PC on P.Codproduto = PC.codproduto where PC.codigo = " & IDpedido
    End If

    TBProduto.Open "Select P.PBruto from projproduto P " & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
        txtpeso = IIf(IsNull(TBProduto!PBruto), "0,00000", Format(TBProduto!PBruto, "###,##0.0000000000"))
    End If
    TBProduto.Close
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Lista
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then .ListItems.Item(InitFor).Checked = False Else .ListItems.Item(InitFor).Checked = True
        Next InitFor
    End With
Else
    OrdenaListView Lista, ColumnHeader
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If ListView1.ListItems.Count = 0 Then Exit Sub
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select * from projproduto where codproduto = " & ListView1.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    ProcLimpaCamposItem
    ProcCarregaDados
End If
TBProduto.Close

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub optIgual_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Txt_percenual_perda_Change()
On Error GoTo tratar_erro

If Txt_percenual_perda <> "" Then
    VerifNumero = Txt_percenual_perda
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_percenual_perda = ""
        Txt_percenual_perda.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Txt_percenual_perda_LostFocus()
On Error GoTo tratar_erro

Txt_percenual_perda = Format(Txt_percenual_perda, "###,##0.0000")

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Txt_posicao_Change()
On Error GoTo tratar_erro

If Txt_posicao <> "" Then
    VerifNumero = Txt_posicao
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_posicao = ""
        Txt_posicao.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Txt_posicao_GotFocus()
On Error GoTo tratar_erro
  
ProcGotFocus Txt_posicao

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub txtDesenho_Change()
On Error GoTo tratar_erro

If txtDesenho = "" Then Exit Sub
ProcCarregaComboCodRef cmbcodref, "P.desenho = '" & txtDesenho & "'", 0, "", False, True
ProcCarregaComboPartNumberFab Cmb_part_number_fabricante, "P.Desenho = '" & txtDesenho & "' and PF.Part_number IS NOT NULL"
ProcCarregaPosicao

'If frmProj_produto_Estrutura.Novo_Estrutura = True Then
'    With cmbVersao
'        .Clear
'        If ProcVerifVersaoCriada("A") = False Then .AddItem "A"
'        If ProcVerifVersaoCriada("B") = False Then .AddItem "B"
'        If ProcVerifVersaoCriada("C") = False Then .AddItem "C"
'        If ProcVerifVersaoCriada("D") = False Then .AddItem "D"
'        If ProcVerifVersaoCriada("E") = False Then .AddItem "E"
'        If ProcVerifVersaoCriada("F") = False Then .AddItem "F"
'        If ProcVerifVersaoCriada("G") = False Then .AddItem "G"
'        If ProcVerifVersaoCriada("H") = False Then .AddItem "H"
'        If ProcVerifVersaoCriada("I") = False Then .AddItem "I"
'        If ProcVerifVersaoCriada("J") = False Then .AddItem "J"
'        If ProcVerifVersaoCriada("K") = False Then .AddItem "K"
'        If ProcVerifVersaoCriada("L") = False Then .AddItem "L"
'        If ProcVerifVersaoCriada("M") = False Then .AddItem "M"
'        If ProcVerifVersaoCriada("N") = False Then .AddItem "N"
'        If ProcVerifVersaoCriada("O") = False Then .AddItem "O"
'        If ProcVerifVersaoCriada("P") = False Then .AddItem "P"
'        If ProcVerifVersaoCriada("Q") = False Then .AddItem "Q"
'        If ProcVerifVersaoCriada("R") = False Then .AddItem "R"
'        If ProcVerifVersaoCriada("S") = False Then .AddItem "S"
'        If ProcVerifVersaoCriada("T") = False Then .AddItem "T"
'        If ProcVerifVersaoCriada("U") = False Then .AddItem "U"
'        If ProcVerifVersaoCriada("V") = False Then .AddItem "V"
'        If ProcVerifVersaoCriada("W") = False Then .AddItem "W"
'        If ProcVerifVersaoCriada("X") = False Then .AddItem "X"
'        If ProcVerifVersaoCriada("Y") = False Then .AddItem "Y"
'        If ProcVerifVersaoCriada("Z") = False Then .AddItem "Z"
'    End With
'End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub cmbunkg_Click()
On Error GoTo tratar_erro

If cmbunkg = "Mt²" Then Label24.Caption = "Area" Else Label24.Caption = "Dim. / mm"
ProcVerificaValor

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub cmbVU_Change()
On Error GoTo tratar_erro

If cmbVU.Text <> "" Then
    VerifNumero = cmbVU.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        cmbVU.Text = ""
        cmbVU.SetFocus
        Exit Sub
    End If
End If
ProcVerificaValor

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub cmbVU_LostFocus()
On Error GoTo tratar_erro

cmbVU = Format(cmbVU, "###,##0.0000000000")

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 10125, 6, True
Formulario = "Engenharia/Estrutura"
If PCP_Ordem = True Then
    Caption = "PCP - Gerenciamento de ordem - Novo/alterar material requisitado"
    Formulario = "PCP/Gerenciamento de ordem"
    Familiatext = "P"
    If ID_Familia <> 0 Then
        Label4.Width = 1770
        txtpeso.Width = 1770
        cmdPesoBruto.Visible = False
        chkFiltrarEstoque.Visible = False
        Frame3.Visible = False
        ListView1.Visible = False
        Frame9.Visible = False
        PBLista.Visible = False
        Frame1.Top = 1020
        Frame2.Top = 3360
        Height = 4710
        chkFiltrarEstoque.Visible = False
        USToolBar1.ButtonState(1) = 5
    End If
Else
'    With cmbVersao
'        .Text = IIf(frmProj_produto_Estrutura.VersaoEstrutura = "0", "A", frmProj_produto_Estrutura.VersaoEstrutura)
'        .Locked = True
'        .TabStop = False
'    End With
    cmbVersao.Text = IIf(frmproj_produto_estrutura.VersaoEstrutura = "", "A", frmproj_produto_estrutura.VersaoEstrutura)
    Familiatext = "E"
    Frame3.Top = 990
    With ListView1
        .Top = 2520
        .Height = 2775
    End With
    Frame9.Top = 5340
    PBLista.Top = 5970
    Frame1.Top = 6240
    Frame2.Top = 8580
    Height = 9945
    chkFiltrarEstoque.Visible = False
End If
Direitos

ProcFiltroPadrao cmbfiltrarpor, optMeio, optFim, optIgual, 0, "Produtos/Serviços", Familiatext, False
If Permitido = False Then cmbfiltrarpor = "Código interno"

ProcCarregaComboFamilia cmbFamilia, "familia <> 'Null'", True
Cmb_ordenar = "Código interno"
If StrSqlLocProdPadrao <> "" Then ProcCarregaLista

txtCodigo = IDlista
If ProcLiberaCamposEstrutura = True Then
    With txtpeso
        .Locked = False
        .TabStop = True
    End With
    With cmbunkg
        .Locked = False
        .TabStop = True
    End With
End If
    
Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcSalvar()
On Error GoTo tratar_erro

If Alterar = False Then
    MsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation
    Exit Sub
End If
Acao = "salvar"
If cmbVersao = "" Then
    NomeCampo = "a versão"
    ProcVerificaAcao
    cmbVersao.SetFocus
    Exit Sub
End If
If Txt_posicao = "" Then
    NomeCampo = "a posição"
    ProcVerificaAcao
    Txt_posicao.SetFocus
    Exit Sub
End If
If txtDesenho.Text = "" Then
    NomeCampo = "o código interno na lista"
    ProcVerificaAcao
    Exit Sub
End If
If txtpeso.Text = "" Then
    NomeCampo = "o peso"
    ProcVerificaAcao
    txtpeso.SetFocus
    Exit Sub
End If
If cmbunkg.Text = "" Then
    NomeCampo = "a unidade do kilograma"
    ProcVerificaAcao
    cmbunkg.SetFocus
    Exit Sub
End If
If txtdimensao.Text = "" Then
    NomeCampo = "a dimensão"
    ProcVerificaAcao
    txtdimensao.SetFocus
    Exit Sub
End If
Valor = IIf(txtquantidade = "", 0, txtquantidade)
If Valor <= 0 Then
    NomeCampo = "a quantidade"
    ProcVerificaAcao
    txtquantidade.SetFocus
    Exit Sub
End If
'Verifica tipo do produto para ver se é obrigatório
If PCP_Ordem = True Then
    TextoFiltro = "P.Desenho = '" & frmprod.txtDesenho & "'"
Else
    If frmproj_produto_estrutura.Novo_Estrutura = True Then TextoFiltro = "P.codproduto = " & IDlista Else TextoFiltro = "P.codproduto = " & IDAntigo
End If
Set TBFI = CreateObject("adodb.recordset")
TBFI.Open "Select P.Codproduto from Projproduto P INNER JOIN projproduto_Tipo PT ON PT.ID = P.ID_Tipo where " & TextoFiltro & " and (PT.Codigo = '03' or PT.Codigo = '04')", Conexao, adOpenKeyset, adLockOptimistic
If TBFI.EOF = False Then
    Set TBFI = CreateObject("adodb.recordset")
    TBFI.Open "Select P.Codproduto from Projproduto P INNER JOIN projproduto_Tipo PT ON PT.ID = P.ID_Tipo where P.Desenho = '" & txtDesenho & "' and (PT.Codigo = '00' or PT.Codigo = '01' or PT.Codigo = '02' or PT.Codigo = '03' or PT.Codigo = '04' or PT.Codigo = '05' or PT.Codigo = '10')", Conexao, adOpenKeyset, adLockOptimistic
    If TBFI.EOF = False Then
        If Txt_percenual_perda = "" Then
            NomeCampo = "o percentual de perda"
            ProcVerificaAcao
            Txt_percenual_perda.SetFocus
            TBFI.Close
            Exit Sub
        End If
    End If
End If
TBFI.Close

If PCP_Ordem = True Then
    With frmprod
        Qtde = .txtquantidade
        Valor1 = txtquantidade
        Valor2 = txtpesototal
        Valor3 = txtdimensao
        
        If cmbunkg <> "N/a" And cmbunkg <> "" And (txtUN = "KG" Or txtUN = "MT" Or txtUN = "MM") Then
            Select Case txtUN
                Case "KG": Peso = Valor2
                Case "MT": Peso = (Valor3 * Valor1) / 1000
                Case "MM": Peso = Valor3 * Valor1
            End Select
        Else
            Peso = txtquantidade
        End If
                
        If ID_Familia = 0 Then
            Set TBMaterial = CreateObject("adodb.recordset")
            TBMaterial.Open "Select codigo from producaomaterial where codigo = '" & txtDesenho & "' and ordem = " & .txtof, Conexao, adOpenKeyset, adLockOptimistic
            If TBMaterial.EOF = False Then
                MsgBox ("O material " & txtDesenho & " já foi adicionado a esta ordem."), vbInformation
                TBMaterial.Close
                Exit Sub
            End If
            TBMaterial.Close
        Else
            'Verifica se a quantidade nova é menor que a quantidde empenhada na ordem
            qt = 0
            Set TBMaterial = CreateObject("adodb.recordset")
            TBMaterial.Open "Select ISNULL(Sum(Quantidade), 0) as Qt from Producao_NF_Consignada where Codinterno = '" & txtDesenho & "' and ordem = " & .txtof, Conexao, adOpenKeyset, adLockOptimistic
            If TBMaterial.EOF = False Then
                qt = TBMaterial!qt
            End If
            TBMaterial.Close
            If Format(qt, "###,##0.0000") > Format((Qtde * Peso), "###,##0.0000") Then
                MsgBox ("Não é permitido alterar para esta quantidade, pois a quantidade requisitada será menor que a quantidade empenhada." & vbCrLf & "Requisitada: " & Format(Qtde * Peso, "###,##0.0000") & " " & txtUN & vbCrLf & "Empenhada: " & Format(qt, "###,##0.0000") & " " & txtUN), vbExclamation
                Exit Sub
            End If
            
            'Verifica se a quantidade nova é menor que a quantidde baixada no estoque
            qt = 0
            Set TBMaterial = CreateObject("adodb.recordset")
            TBMaterial.Open "Select ISNULL(Sum(Saida), 0) as qt from estoque_movimentacao where oe = '" & .txtof & "' and desenho = '" & txtDesenho & "' and documento = '" & .txtof & "' and (operacao = 'SAIDA_ORDEM' or operacao = 'SAIDA_ORDEM_PARCIAL')", Conexao, adOpenKeyset, adLockOptimistic
            If TBMaterial.EOF = False Then
                qt = TBMaterial!qt
            End If
            TBMaterial.Close
            If Format(qt, "###,##0.0000") > Format((Qtde * Peso), "###,##0.0000") Then
                MsgBox ("Não é permitido alterar para esta quantidade, pois a quantidade requisitada será menor que quantidade baixada no estoque." & vbCrLf & "Requisitada: " & Format(Qtde * Peso, "###,##0.0000") & " " & txtUN & vbCrLf & "Baixada no estoque: " & Format(qt, "###,##0.0000") & " " & txtUN), vbExclamation
                Exit Sub
            End If
        End If
        
        Set TBItem = CreateObject("adodb.recordset")
        TBItem.Open "Select * from projproduto where desenho = '" & txtDesenho.Text & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBItem.EOF = False Then
            Set TBMaterial = CreateObject("adodb.recordset")
            TBMaterial.Open "Select * from producaomaterial where IdMateriaPrima = " & ID_Familia, Conexao, adOpenKeyset, adLockOptimistic
            If TBMaterial.EOF = False Then
                MsgBox ("Alteração efetuada com sucesso."), vbInformation
                Evento = "Alterar material"
            Else
                TBMaterial.AddNew
                TBMaterial!Versao = cmbVersao
                TBMaterial!Saida = "NÃO"
                MsgBox ("Novo material requisitado agregado a ordem com sucesso."), vbInformation
                Evento = "Adicionar material"
            End If
            TBMaterial!Posicao = Txt_posicao
            TBMaterial!Quantidade = Valor1 * Qtde
            TBMaterial!Unidade = txtUN
            TBMaterial!CODIGO = txtDesenho
            If Cmb_part_number_fabricante <> "" Then TBMaterial!ID_partnumber_fabricante = Cmb_part_number_fabricante.ItemData(Cmb_part_number_fabricante.ListIndex)
            TBMaterial!Descricao = txtdescricao
            TBMaterial!Obs = Txt_obs
            TBMaterial!Ordem = .txtof
            TBMaterial!PesoMetro = txtpeso
            TBMaterial!Pesounidade = txtkgpc
            TBMaterial!PesoTotal = Valor2 * Qtde
            TBMaterial!Percentual_perda = IIf(Txt_percenual_perda = "", 0, Txt_percenual_perda)
            TBMaterial!Dimensao = txtdimensao
            TBMaterial!Requisitado = Format(Peso * Qtde, "###,##0.0000")
            If TBItem!Un_Kg = "Mt²" Then TBMaterial!DimensaoTotal = ((Valor3 / 1000) / 1000) * TBMaterial!Quantidade Else TBMaterial!DimensaoTotal = (Valor3 / 1000) * TBMaterial!Quantidade
            
            If txtUN = "KG" Or TBItem!SubTipoItem = 1 Or TBItem!SubTipoItem = 2 Or TBItem!SubTipoItem = 3 Then
                If txtUN = "KG" And (TBItem!Un_Kg = "Mt²" Or TBItem!Un_Kg = "Mt/L") Then
                    If IsNull(TBItem!PBruto) = False And TBItem!PBruto > 0 And TBItem!PBruto <> "" Then TBMaterial!Total_pc = Format(TBMaterial!Requisitado / TBItem!PBruto, "###,##0.0000") Else TBMaterial!Total_pc = Null
                Else
                    If txtUN = "PÇ" Or txtUN = "PC" Or txtUN = "UN" Or txtUN = "CJ" Then TBMaterial!Total_pc = TBMaterial!Requisitado Else TBMaterial!Total_pc = Null
                End If
            End If
            
            'Verifica qtde. de saida
            qtdeliberada = 0
            qtdeliberadaPC = 0
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select Sum(Saida) as qtdeliberada, ISNULL(Sum(Saida_PC), 0) as qtdeliberadaPC from estoque_movimentacao where oe = '" & .txtof & "' and desenho = '" & txtDesenho & "' and documento = '" & .txtof & "' and (operacao = 'SAIDA_ORDEM' or operacao = 'SAIDA_ORDEM_PARCIAL')", Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                qtdeliberada = IIf(IsNull(TBAbrir!qtdeliberada), 0, Format(TBAbrir!qtdeliberada, "###,##0.0000"))
                qtdeliberadaPC = IIf(IsNull(TBAbrir!qtdeliberadaPC), 0, Format(TBAbrir!qtdeliberadaPC, "###,##0.0000"))
            End If
            
            If qtdeliberada = 0 And qtdeliberadaPC = 0 Then
                TBMaterial!Saida = "NÃO"
            ElseIf qtdeliberada >= TBMaterial!Requisitado Or qtdeliberadaPC >= TBMaterial!Total_pc Then
                    TBMaterial!Saida = "SIM"
                Else
                    TBMaterial!Saida = "PARCIAL"
            End If
            TBMaterial.Update
        End If
        '==================================
        Modulo = "PCP/Gerenciamento de ordem"
        ID_documento = TBMaterial!IdMateriaPrima
        Documento = "Ordem: " & .txtof.Text & " - Cód. interno: " & .txtDesenho
        Documento = "Cód. interno: " & txtDesenho
        ProcGravaEvento
        '==================================
        
        If chkFiltrarEstoque.Value = 1 Then
            Set TBproducao = CreateObject("adodb.recordset")
            TBproducao.Open "Select * from Producao_NF_Consignada", Conexao, adOpenKeyset, adLockOptimistic
            TBproducao.AddNew
            TBproducao!Data = Date
            TBproducao!Responsavel = pubUsuario
            TBproducao!Ordem = .txtof
            TBproducao!Codinterno = txtDesenho
            TBproducao!IDestoque = ListView1.SelectedItem.ListSubItems(1)
            TBproducao!Quantidade = Valor1 * Qtde
            TBproducao!Quantidade_PC = IIf(IsNull(TBMaterial!Total_pc), 0, TBMaterial!Total_pc)
            
            TBproducao.Update
            '==================================
            Modulo = "PCP/Gerenciamento de ordem"
            Evento = "Empenhar RE"
            ID_documento = TBproducao!ID
            Documento = "Ordem: " & .txtof.Text & " - Cód. interno: " & .txtDesenho
            Documento1 = "Cód. interno: " & txtDesenho & " - RE: " & ListView1.SelectedItem.ListSubItems(1)
            ProcGravaEvento
            '==================================
            TBproducao.Close
        End If
        TBMaterial.Close

        .ProcCarregaListaRequisicao
        If Evento = "Alterar material" Then Unload Me
    End With
Else
    With frmproj_produto_estrutura
        If .Novo_Estrutura = True Then TextoFiltro = "codproduto = " & IDlista Else TextoFiltro = "codproduto = " & IDAntigo
        
        Set TBProduto = CreateObject("adodb.recordset")
        TBProduto.Open "Select * from ProjConjunto where Codigo = " & IDpedido, Conexao, adOpenKeyset, adLockOptimistic
        If TBProduto.EOF = True Then
            TBProduto.AddNew
            TBProduto!Codproduto = IDlista
            MsgBox ("Novo registro agregado na estrutura com sucesso."), vbInformation
            Evento = "Novo"
            
            TextoFiltroPos = "Posicao = Posicao + 1 where Posicao >= " & Txt_posicao
        Else
            MsgBox ("Alteração efetuada com sucesso."), vbInformation
            Evento = "Alterar"
            
            TextoFiltroPos = ""
            If Txt_posicao < TBProduto!Posicao Then
                 TextoFiltroPos = "Posicao = Posicao + 1 where Posicao >= " & Txt_posicao & " and Posicao < " & TBProduto!Posicao
            ElseIf Txt_posicao > TBProduto!Posicao Then
                    TextoFiltroPos = "Posicao = Posicao - 1 where Posicao > " & TBProduto!Posicao & " and Posicao <= " & Txt_posicao
            End If
        End If
        If TextoFiltroPos <> "" Then Conexao.Execute "Update ProjConjunto Set " & TextoFiltroPos & " and Posicao IS NOT NULL and " & TextoFiltro & " and Versao = '" & .VersaoEstrutura & "'"
        
        TBProduto!Versao = .VersaoEstrutura
        TBProduto!Posicao = Txt_posicao
        TBProduto!Desenho = txtDesenho.Text
        TBProduto!Versao_desenho = cmbVersao
        If Cmb_part_number_fabricante <> "" Then TBProduto!ID_partnumber_fabricante = Cmb_part_number_fabricante.ItemData(Cmb_part_number_fabricante.ListIndex)
        TBProduto!Descricao = txtdescricao.Text
        TBProduto!Obs = Txt_obs
        TBProduto!Quantidade = txtquantidade.Text
        TBProduto!Dimensoes = txtdimensao.Text
        TBProduto!Peso = txtkgpc.Text
        TBProduto!PesoMetro = IIf(txtpeso.Text = "", 0, txtpeso)
        TBProduto!PesoTotal = txtpesototal.Text
        TBProduto!Percentual_perda = IIf(Txt_percenual_perda = "", 0, Txt_percenual_perda)
        TBProduto!Unidade = txtUN.Text
        TBProduto!Un_Kg = cmbunkg.Text
        TBProduto!Valor = cmbVU.Text
        TBProduto!ValorTotal = txtVT.Text
        TBProduto.Update
        IDpedido = TBProduto!CODIGO
        TBProduto.Close
        '==================================
        Modulo = "Engenharia/Estrutura"
        ID_documento = IDAntigo
        Set TBProduto = CreateObject("adodb.recordset")
        TBProduto.Open "Select Desenho from Projproduto where " & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
        If TBProduto.EOF = False Then
            Documento = "Cód. interno: " & TBProduto!Desenho
        End If
        TBProduto.Close
        Documento1 = "Cód. interno: " & txtDesenho
        ProcGravaEvento
        '==================================
        If .Novo_Estrutura = True Then ProcLimpaCamposItem
        .ProcCarregaLista
    End With
End If
'Unload Me

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro
  
Select Case KeyCode
    Case vbKeyF2: ProcFiltrar
    Case vbKeyF3: ProcSalvar
    Case vbKeyF4: Cmd_excluir_Click
    Case vbKeyF7: Cmd_salvar_Click
    Case vbKeyEscape: ProcSair
End Select

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Sub ProcLimpaCamposItem()
On Error GoTo tratar_erro

'If cmbVersao.Locked = False Then
    cmbVersao.ListIndex = -1
    Txt_posicao = ""
'End If
txtDesenho = ""
cmbcodref.ListIndex = -1
Cmb_part_number_fabricante.ListIndex = -1
txtUN.Text = ""
Txt_familia = ""
txtdescricao.Text = ""
Txt_obs = ""
txtpeso.Text = "0,00000"
cmbunkg.ListIndex = -1
txtdimensao.Text = "0,00000"
txtkgpc.Text = "0,00000"
txtquantidade.Text = "0,00000"
txtpesototal.Text = "0,00000"
Txt_percenual_perda = "0,0000"
txtVT.Text = "0,00"
cmbVU.Text = "0,00000"

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

txtCodigo = IDlista

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

Unload Me

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub txtdimensao_Change()
On Error GoTo tratar_erro

If txtdimensao.Text <> "" Then
    VerifNumero = txtdimensao.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtdimensao.Text = ""
        txtdimensao.SetFocus
        Exit Sub
    End If
End If
ProcCalculaPeso
ProcVerificaValor

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcCalculaPeso()
On Error GoTo tratar_erro

If txtpeso.Text <> "" And cmbunkg.Text <> "" And txtdimensao.Text <> "" And txtquantidade.Text <> "" Then
    If cmbunkg.Text = "Mt/L" Then txtkgpc.Text = Format(txtpeso.Text / 1000 * txtdimensao, "###,##0.0000000000")
    If cmbunkg.Text = "Pç" Then txtkgpc.Text = Format(txtpeso.Text, "###,##0.0000000000")
    If cmbunkg.Text = "Mt²" Then txtkgpc.Text = Format(((txtdimensao * txtpeso) / 1000) / 1000, "###,##0.0000000000")
    If cmbunkg.Text = "N/a" Then txtkgpc.Text = Format(0, "###,##0.0000000000")
    If txtdimensao.Text = "" Then txtdimensao.Text = Format(0, "###,##0.0000000000")
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcalculaPesoTotal()
On Error GoTo tratar_erro

If txtkgpc.Text <> "" And txtquantidade <> "" Then
    txtpesototal = Format(txtkgpc.Text * txtquantidade.Text, "###,##0.0000000000")
Else
    txtpesototal = "0,00000"
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Sub ProcVerificaValor()
On Error GoTo tratar_erro

ProcCalculaPeso
ProcalculaPesoTotal
If cmbVU.Text <> "" And txtquantidade.Text <> "" And txtdimensao.Text <> "" Then
    Select Case txtUN
        Case "KG": txtVT = Format(cmbVU * txtpesototal, "###,##0.00")
        Case "MM": txtVT = Format((cmbVU * txtdimensao) * txtquantidade, "###,##0.00")
        Case "MT": txtVT = Format(((cmbVU / 1000) * txtdimensao) * txtquantidade, "###,##0.00")
    End Select
    If txtUN <> "KG" And txtUN <> "MM" And txtUN <> "MT" Then txtVT = Format(cmbVU * txtquantidade, "###,##0.00")
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub txtdimensao_LostFocus()
On Error GoTo tratar_erro

txtdimensao.Text = Format(txtdimensao.Text, "###,##0.0000000000")

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub txtkgpc_Change()
On Error GoTo tratar_erro

ProcVerificaValor

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub txtkgpc_LostFocus()
On Error GoTo tratar_erro

If txtkgpc.Text <> "" Then
    VerifNumero = txtkgpc.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtkgpc.Text = ""
        txtkgpc.SetFocus
        Exit Sub
    End If
    txtkgpc.Text = Format(txtkgpc.Text, "###,##0.0000000000")
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
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
ProcVerificaValor

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub txtpeso_LostFocus()
On Error GoTo tratar_erro

txtpeso.Text = Format(txtpeso.Text, "###,##0.0000000000")

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub txtpesototal_Change()
On Error GoTo tratar_erro

ProcVerificaValor

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub txtQuantidade_Change()
On Error GoTo tratar_erro

If txtquantidade.Text <> "" Then
    VerifNumero = txtquantidade.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtquantidade.Text = ""
        txtquantidade.SetFocus
        Exit Sub
    End If
End If
ProcVerificaValor
txtVT.Text = Format(txtVT.Text, "###,##0.00")

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub txtquantidade_LostFocus()
On Error GoTo tratar_erro

txtquantidade.Text = Format(txtquantidade.Text, "###,##0.0000000000")

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal Key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcFiltrar
    Case 2: ProcSalvar
    'Case 4: ProcAjuda
    Case 5: ProcSair
End Select

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Cmb_ordenar_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub cmbFamilia_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
If cmbFamilia <> "" Then txtTexto = ""

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub cmbfiltrarpor_Click()
On Error GoTo tratar_erro

With cmbFamilia
    ListView1.ListItems.Clear
    If cmbfiltrarpor = "Família" Or cmbfiltrarpor = "Cliente" Or cmbfiltrarpor = "Fornecedor" Then
        txtTexto.Visible = False
        .Visible = True
        .Clear
        .AddItem ""
        If cmbfiltrarpor = "Família" Then
            ProcCarregaComboFamilia cmbFamilia, "familia <> 'Null'", True
        Else
            If cmbfiltrarpor = "Cliente" Then ProcCarregaComboCliForn cmbFamilia, True Else ProcCarregaComboCliForn cmbFamilia, False
        End If
    Else
        txtTexto.Visible = True
        .Visible = False
        If cmbfiltrarpor = "RE" And txtTexto <> "" Then
            VerifNumero = txtTexto
            ProcVerificaNumero
            If VerifNumero = False Then
                txtTexto = ""
                txtTexto.SetFocus
                Exit Sub
            End If
        End If
    End If
End With

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

If PCP_Ordem = True Then
     TextoFiltroPadrao = "P.Desenho IS NOT NULL"
Else
    If frmproj_produto_estrutura.Novo_Estrutura = True Then TextoFiltro = "codproduto = " & IDlista Else TextoFiltro = "codproduto = " & IDAntigo
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select Desenho from Projproduto where " & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
        TextoFiltroPadrao = "P.Desenho <> '" & TBProduto!Desenho & "' and P.Subtipoitem <> 4 and P.Bloqueado = 'False'"
    End If
    TBProduto.Close
End If

INNERJOIN_Estoque = ""
TextoFiltro_Estoque = ""
If chkFiltrarEstoque.Value = 1 Then
    CamposFiltro = "P.codProduto, P.Desenho, P.Descricao, P.Unidade, P.classe, E.IdEstoque"
    INNERJOIN_Estoque = " INNER JOIN Estoque_produtos E ON E.codproduto = P.codproduto"
    If frmprod.optconsignacao.Value = 1 Then TextoFiltro_Estoque2 = "(E.Consignacao = 'True' and E.id_cliente = " & frmprod.Txt_ID_cliente & " and E.Cliente = '" & frmprod.txtCliente & "' or E.Consignacao = 'False')" Else TextoFiltro_Estoque2 = "E.Consignacao = 'False'"
    TextoFiltro_Estoque = " and E.Estoque_real > 0 and " & TextoFiltro_Estoque2 & " and E.Liberado = 'SIM' and (Left(E.status, 7) = 'ENTRADA' or E.status = 'CONSIGNAÇÃO RECEBIDA')"
Else
    CamposFiltro = "P.codProduto, P.Desenho, P.Descricao, P.Unidade, P.classe"
End If
INNERJOINTEXTO = "Select " & CamposFiltro & " from (((Projproduto P LEFT JOIN item_aplicacoes IA ON P.codproduto = IA.codproduto) LEFT JOIN Projproduto_clientes PC ON P.codproduto = PC.codproduto) LEFT JOIN Projproduto_fornecedor PF ON P.codproduto = PF.codproduto) LEFT JOIN Projproduto_fabricante PFAB ON PFAB.Codproduto = P.codproduto" & INNERJOIN_Estoque
If Cmb_ordenar = "Código interno" Then Ordenar = "P.desenho" Else Ordenar = "P.Descricao"
TextoFiltroPadrao = TextoFiltroPadrao & " and P.Subtipoitem <> 4 and P.Bloqueado = 'False'" & TextoFiltro_Estoque & " group by " & CamposFiltro & " order by " & Ordenar

If Lista.ListItems.Count = 0 Then
    If txtTexto.Visible = True And txtTexto <> "" Or cmbFamilia.Visible = True And cmbFamilia <> "" Then
        If cmbfiltrarpor = "Cliente" Or cmbfiltrarpor = "Fornecedor" Or cmbfiltrarpor = "Família" Then
            Select Case cmbfiltrarpor
                Case "Cliente": TextoFiltro = "PC.IDCliente"
                Case "Fornecedor": TextoFiltro = "PF.IDfornecedor"
                Case "Família": TextoFiltro = "P.classe"
            End Select
            If cmbfiltrarpor = "Família" Then TextoFiltro = TextoFiltro & " = '" & cmbFamilia & "'" Else TextoFiltro = TextoFiltro & " = " & cmbFamilia.ItemData(cmbFamilia.ListIndex)
            StrSqlLocProdPadrao = INNERJOINTEXTO & " where " & TextoFiltro & " and " & TextoFiltroPadrao
        ElseIf cmbfiltrarpor = "Comprimento" Or cmbfiltrarpor = "Largura" Or cmbfiltrarpor = "Espessura" Then
                Select Case cmbfiltrarpor
                    Case "Comprimento": TextoFiltro = "P.Comprimento"
                    Case "Largura": TextoFiltro = "P.Largura"
                    Case "Espessura": TextoFiltro = "P.Espessura"
                End Select
                Valor = txtTexto
                NovoValor = Replace(Valor, ",", ".")
                StrSqlLocProdPadrao = INNERJOINTEXTO & " where " & TextoFiltro & " = " & NovoValor & " and " & TextoFiltroPadrao
            Else
                Select Case cmbfiltrarpor
                    Case "Código interno": TextoFiltro = "P.desenho"
                    Case "Descrição": TextoFiltro = "P.descricao"
                    Case "Descrição comercial": TextoFiltro = "P.Descricaotecnica"
                    Case "Dureza": TextoFiltro = "P.Dureza"
                    Case "Part number": TextoFiltro = "PFAB.Part_number"
                    Case "Código de referência": TextoFiltro = "IA.N_referencia"
                    Case "Número do desenho": TextoFiltro = "IA.desenho"
                    Case "RE": TextoFiltro = "E.IdEstoque"
                    Case "Lote": TextoFiltro = "E.Lote"
                End Select
                StrSqlLocProdPadrao = INNERJOINTEXTO & " where " & TextoFiltro & VerifTipoFiltroIMF(optInicio, optMeio, optFim, optIgual, txtTexto) & " and " & TextoFiltroPadrao
        End If
    Else
        StrSqlLocProdPadrao = INNERJOINTEXTO & " where " & TextoFiltroPadrao
    End If
Else
    TextoFiltroLista = ""
    With Lista
        For InitFor = 1 To .ListItems.Count
            If .ListItems(InitFor).ListSubItems(1) = "Cliente" Or .ListItems(InitFor).ListSubItems(1) = "Fornecedor" Then
                If .ListItems(InitFor).ListSubItems(1) = "Cliente" Then TextoFiltro = "PC.IDCliente" Else TextoFiltro = "PF.IDfornecedor"
                If TextoFiltroLista = "" Then TextoFiltroLista = INNERJOINTEXTO & " where " & TextoFiltro & " = " & .ListItems(InitFor).ListSubItems(4) Else TextoFiltroLista = TextoFiltroLista & " and " & TextoFiltro & " = " & .ListItems(InitFor).ListSubItems(4)
            ElseIf .ListItems(InitFor).ListSubItems(1) = "Família" Then
                    If TextoFiltroLista = "" Then TextoFiltroLista = INNERJOINTEXTO & " where P.Classe = '" & .ListItems(InitFor).ListSubItems(3) & "'" Else TextoFiltroLista = TextoFiltroLista & " and " & TextoFiltro & " = '" & .ListItems(InitFor).ListSubItems(3) & "'"
                ElseIf .ListItems(InitFor).ListSubItems(1) = "Comprimento" Or .ListItems(InitFor).ListSubItems(1) = "Largura" Or .ListItems(InitFor).ListSubItems(1) = "Espessura" Then
                        Select Case .ListItems(InitFor).ListSubItems(1)
                            Case "Comprimento": TextoFiltro = "P.Comprimento"
                            Case "Largura": TextoFiltro = "P.Largura"
                            Case "Espessura": TextoFiltro = "P.Espessura"
                        End Select
                        Valor = .ListItems(InitFor).ListSubItems(3)
                        NovoValor = Replace(Valor, ",", ".")
                        If TextoFiltroLista = "" Then TextoFiltroLista = INNERJOINTEXTO & " where " & TextoFiltro & " = " & NovoValor Else TextoFiltroLista = TextoFiltroLista & " and " & TextoFiltro & " = " & NovoValor
                    Else
                        Select Case .ListItems(InitFor).ListSubItems(1)
                            Case "Código interno": TextoFiltro = "P.desenho"
                            Case "Código de referência": TextoFiltro = "IA.N_referencia"
                            Case "Número do desenho": TextoFiltro = "IA.desenho"
                            Case "Descrição": TextoFiltro = "P.descricao"
                            Case "Descrição comercial": TextoFiltro = "P.Descricaotecnica"
                            Case "Dureza": TextoFiltro = "P.Dureza"
                            Case "Part number": TextoFiltro = "PFAB.Part_number"
                            Case "RE": TextoFiltro = "E.IdEstoque"
                            Case "Lote": TextoFiltro = "E.Lote"
                        End Select
                        If TextoFiltroLista = "" Then TextoFiltroLista = INNERJOINTEXTO & " where " & TextoFiltro & VerifTipoFiltroIMFLista(.ListItems(InitFor).ListSubItems(2), .ListItems(InitFor).ListSubItems(3)) Else TextoFiltroLista = TextoFiltroLista & " and " & TextoFiltro & VerifTipoFiltroIMFLista(.ListItems(InitFor).ListSubItems(2), .ListItems(InitFor).ListSubItems(3))
            End If
        Next InitFor
    End With
    StrSqlLocProdPadrao = TextoFiltroLista & " and " & TextoFiltroPadrao
End If
ProcCarregaLista

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub cmdPagAnt_Click()
On Error GoTo tratar_erro

If SóNumeros(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLocalizar_produto_padrao.AbsolutePage <> 2 Then
    If TBLocalizar_produto_padrao.AbsolutePage = -3 Then
        ProcExibePagina (TBLocalizar_produto_padrao.PageCount - 1)
    Else
        TBLocalizar_produto_padrao.AbsolutePage = TBLocalizar_produto_padrao.AbsolutePage - 2
        ProcExibePagina (TBLocalizar_produto_padrao.AbsolutePage)
    End If
Else
    ProcExibePagina (1)
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub cmdPagIr_Click()
On Error GoTo tratar_erro

If txtPagIr = "" Then Exit Sub
Quant = SóNumeros(Right(lblPaginas.Caption, 4))
If Quant <= 1 Or txtPagIr > Quant Then Exit Sub
If txtPagIr.Text >= 1 And txtPagIr.Text <= Quant Then
    TBLocalizar_produto_padrao.AbsolutePage = txtPagIr.Text
    ProcExibePagina (TBLocalizar_produto_padrao.AbsolutePage)
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub cmdPagPrim_Click()
On Error GoTo tratar_erro

If SóNumeros(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLocalizar_produto_padrao.AbsolutePage = 1
ProcExibePagina (TBLocalizar_produto_padrao.AbsolutePage)

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub cmdPagProx_Click()
On Error GoTo tratar_erro

If SóNumeros(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLocalizar_produto_padrao.AbsolutePage <> -3 Then
    If TBLocalizar_produto_padrao.AbsolutePage = 1 Then
        ProcExibePagina (2)
    Else
        ProcExibePagina (TBLocalizar_produto_padrao.AbsolutePage)
    End If
Else
    ProcExibePagina (TBLocalizar_produto_padrao.PageCount)
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub cmdPagUlt_Click()
On Error GoTo tratar_erro

If SóNumeros(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLocalizar_produto_padrao.AbsolutePage = TBLocalizar_produto_padrao.PageCount
ProcExibePagina (TBLocalizar_produto_padrao.AbsolutePage)

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

OrdenaListView ListView1, ColumnHeader

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Sub ProcCarregaLista()
On Error GoTo tratar_erro

lblRegistros.Caption = "Nº de registros: 0"
lblPaginas.Caption = "Página: 0 de: 0"
ListView1.ListItems.Clear
If StrSqlLocProdPadrao = "" Then Exit Sub
Set TBLocalizar_produto_padrao = CreateObject("adodb.recordset")
TBLocalizar_produto_padrao.Open StrSqlLocProdPadrao, Conexao, adOpenKeyset, adLockReadOnly
If TBLocalizar_produto_padrao.EOF = False Then ProcExibePagina (1)

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcExibePagina(Pagina)
On Error GoTo tratar_erro

ListView1.ListItems.Clear
TBLocalizar_produto_padrao.PageSize = IIf(txtNreg = "", 30, txtNreg)
TBLocalizar_produto_padrao.AbsolutePage = Pagina
TamanhoPagina = TBLocalizar_produto_padrao.PageSize
ContadorReg = 1

PBLista.Min = 0
PBLista.Max = VerifMaxPBListaPaginacao(TBLocalizar_produto_padrao.RecordCount - IIf(Pagina > 1, (TBLocalizar_produto_padrao.PageSize * (Pagina - 1)), 0), TBLocalizar_produto_padrao.PageSize)
PBLista.Value = 1
Contador = 0
Do While TBLocalizar_produto_padrao.EOF = False And (ContadorReg <= TamanhoPagina)
    With ListView1.ListItems
        .Add , , TBLocalizar_produto_padrao!Codproduto
        If chkFiltrarEstoque.Value = 1 Then .Item(.Count).SubItems(1) = IIf(IsNull(TBLocalizar_produto_padrao!IDestoque), "", TBLocalizar_produto_padrao!IDestoque)
        .Item(.Count).SubItems(2) = IIf(IsNull(TBLocalizar_produto_padrao!Desenho), "", TBLocalizar_produto_padrao!Desenho)
        .Item(.Count).SubItems(3) = IIf(IsNull(TBLocalizar_produto_padrao!Descricao), "", TBLocalizar_produto_padrao!Descricao)
        .Item(.Count).SubItems(4) = IIf(IsNull(TBLocalizar_produto_padrao!Unidade), "", TBLocalizar_produto_padrao!Unidade)
        .Item(.Count).SubItems(5) = IIf(IsNull(TBLocalizar_produto_padrao!classe), "", TBLocalizar_produto_padrao!classe)
    End With
    TBLocalizar_produto_padrao.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista.Value = Contador
Loop
lblRegistros.Caption = "Nº de registros: " & TBLocalizar_produto_padrao.RecordCount
If TBLocalizar_produto_padrao.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "Página: 1 de: " & TBLocalizar_produto_padrao.PageCount
ElseIf TBLocalizar_produto_padrao.AbsolutePage = adPosEOF Then
        lblPaginas.Caption = "Página: " & TBLocalizar_produto_padrao.PageCount & " de: " & TBLocalizar_produto_padrao.PageCount
    Else
        lblPaginas.Caption = "Página: " & TBLocalizar_produto_padrao.AbsolutePage - 1 & " de: " & TBLocalizar_produto_padrao.PageCount
End If


Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Optfim_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Optinicio_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Optmeio_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
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
    MsgBox ("Descrição do erro : " + Error()), vbCritical
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
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub txtTexto_Change()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
If txtTexto <> "" Then
    cmbFamilia.ListIndex = -1
    If cmbfiltrarpor = "RE" Then
        VerifNumero = txtTexto
        ProcVerificaNumero
        If VerifNumero = False Then
            txtTexto = ""
            txtTexto.SetFocus
            Exit Sub
        End If
    End If
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Function ProcVerifVersaoCriada(Versao As String) As Boolean
On Error GoTo tratar_erro

ProcVerifVersaoCriada = False
Set TBItem = CreateObject("adodb.recordset")
TBItem.Open "Select Codigo from ProjConjunto where codproduto = " & IDlista & " and Desenho = '" & txtDesenho & "' and Versao = '" & Versao & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBItem.EOF = False Then ProcVerifVersaoCriada = True
TBItem.Close

Exit Function
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Function
End Function
