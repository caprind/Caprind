VERSION 5.00
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form frmCompras_NaoConformidade 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Compras - Não conformidade"
   ClientHeight    =   10035
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   15360
   ControlBox      =   0   'False
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
   Begin VB.CheckBox chkRecomprado 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Produto/serviço recomprado"
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
      Left            =   3380
      TabIndex        =   31
      Top             =   4230
      Width           =   3015
   End
   Begin VB.TextBox txtID 
      Height          =   315
      Left            =   1500
      TabIndex        =   32
      Text            =   "0"
      Top             =   5400
      Visible         =   0   'False
      Width           =   885
   End
   Begin MSComctlLib.ListView Lista_ComprasNaoConformidade 
      Height          =   5265
      Left            =   60
      TabIndex        =   0
      Top             =   4470
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   9287
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
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Text            =   "ID"
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
         Text            =   "Pedido"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "Fornecedor"
         Object.Width           =   9352
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Cód. interno"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Tag             =   "T"
         Text            =   "Descrição"
         Object.Width           =   9352
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Object.Tag             =   "T"
         Text            =   "Un"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Object.Tag             =   "N"
         Text            =   "Qtde."
         Object.Width           =   2117
      EndProperty
   End
   Begin VB.Frame Frame1 
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
      Height          =   825
      Left            =   65
      TabIndex        =   33
      Top             =   990
      Width           =   15195
      Begin VB.TextBox txtResponsável 
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
         Left            =   1065
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Responsável pela inspeção."
         Top             =   370
         Width           =   3060
      End
      Begin VB.TextBox txtfornecedor 
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
         Left            =   5420
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Fornecedor."
         Top             =   370
         Width           =   9585
      End
      Begin VB.TextBox mskData 
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
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Data da inspeção."
         Top             =   370
         Width           =   870
      End
      Begin VB.TextBox txtIDpedido 
         Height          =   315
         Left            =   4140
         TabIndex        =   34
         Top             =   370
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.TextBox txtPedido 
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
         Left            =   4140
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Pedido de compra."
         Top             =   370
         Width           =   1260
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Fornecedor"
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
         Left            =   9800
         TabIndex        =   38
         Top             =   180
         Width           =   825
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Responsável pela inspeção"
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
         Left            =   1628
         TabIndex        =   37
         Top             =   180
         Width           =   1935
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Pedido"
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
         Left            =   4485
         TabIndex        =   36
         Top             =   180
         Width           =   570
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Data insp."
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
         Left            =   248
         TabIndex        =   35
         Top             =   180
         Width           =   735
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Check list de verificação"
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
      Height          =   2625
      Left            =   9240
      TabIndex        =   45
      Top             =   1800
      Width           =   6015
      Begin VB.TextBox txtobservacoes 
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
         Height          =   615
         Left            =   180
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   30
         ToolTipText     =   "Observações."
         Top             =   1860
         Width           =   5625
      End
      Begin VB.CheckBox chk_emb_ok 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ok"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   2520
         TabIndex        =   12
         Top             =   210
         Width           =   615
      End
      Begin VB.CheckBox chk_emb_nc 
         BackColor       =   &H00E0E0E0&
         Caption         =   "N/C"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   3870
         TabIndex        =   13
         Top             =   210
         Width           =   615
      End
      Begin VB.CheckBox chk_emb_na 
         BackColor       =   &H00E0E0E0&
         Caption         =   "N/A"
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
         Left            =   5220
         TabIndex        =   14
         Top             =   210
         Width           =   615
      End
      Begin VB.CheckBox chk_laudo_ok 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ok"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   2520
         TabIndex        =   15
         Top             =   450
         Width           =   615
      End
      Begin VB.CheckBox chk_laudo_nc 
         BackColor       =   &H00E0E0E0&
         Caption         =   "N/C"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   3870
         TabIndex        =   16
         Top             =   450
         Width           =   615
      End
      Begin VB.CheckBox chk_laudo_na 
         BackColor       =   &H00E0E0E0&
         Caption         =   "N/A"
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
         Left            =   5220
         TabIndex        =   17
         Top             =   450
         Width           =   615
      End
      Begin VB.CheckBox chk_qtd_ok 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ok"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   2520
         TabIndex        =   18
         Top             =   690
         Width           =   615
      End
      Begin VB.CheckBox chk_qtd_nc 
         BackColor       =   &H00E0E0E0&
         Caption         =   "N/C"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   3870
         TabIndex        =   19
         Top             =   690
         Width           =   615
      End
      Begin VB.CheckBox chk_qtd_na 
         BackColor       =   &H00E0E0E0&
         Caption         =   "N/A"
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
         Left            =   5220
         TabIndex        =   20
         Top             =   690
         Width           =   615
      End
      Begin VB.CheckBox chk_visual_ok 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ok"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   2520
         TabIndex        =   21
         Top             =   930
         Width           =   615
      End
      Begin VB.CheckBox chk_visual_nc 
         BackColor       =   &H00E0E0E0&
         Caption         =   "N/C"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   3870
         TabIndex        =   22
         Top             =   930
         Width           =   615
      End
      Begin VB.CheckBox chk_visual_na 
         BackColor       =   &H00E0E0E0&
         Caption         =   "N/A"
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
         Left            =   5220
         TabIndex        =   23
         Top             =   930
         Width           =   615
      End
      Begin VB.CheckBox chk_dim_ok 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ok"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   2520
         TabIndex        =   24
         Top             =   1170
         Width           =   615
      End
      Begin VB.CheckBox chk_dim_nc 
         BackColor       =   &H00E0E0E0&
         Caption         =   "N/C"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   3870
         TabIndex        =   25
         Top             =   1170
         Width           =   615
      End
      Begin VB.CheckBox chk_dim_na 
         BackColor       =   &H00E0E0E0&
         Caption         =   "N/A"
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
         Left            =   5220
         TabIndex        =   26
         Top             =   1170
         Width           =   615
      End
      Begin VB.CheckBox chk_outros_ok 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ok"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   2520
         TabIndex        =   27
         Top             =   1410
         Width           =   615
      End
      Begin VB.CheckBox chk_outros_nc 
         BackColor       =   &H00E0E0E0&
         Caption         =   "N/C"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   3870
         TabIndex        =   28
         Top             =   1410
         Width           =   615
      End
      Begin VB.CheckBox chk_outros_na 
         BackColor       =   &H00E0E0E0&
         Caption         =   "N/A"
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
         Left            =   5220
         TabIndex        =   29
         Top             =   1410
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "6 - Outros"
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
         Left            =   240
         TabIndex        =   52
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "5 - Dimensional"
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
         Left            =   240
         TabIndex        =   51
         Top             =   1200
         Width           =   1080
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "4 - Visual"
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
         Left            =   240
         TabIndex        =   50
         Top             =   960
         Width           =   645
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3 - Quantidade"
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
         Left            =   240
         TabIndex        =   49
         Top             =   720
         Width           =   1080
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2 - Laudos / certificados"
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
         Left            =   240
         TabIndex        =   48
         Top             =   480
         Width           =   1725
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1 - Embalagem"
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
         Left            =   240
         TabIndex        =   47
         Top             =   240
         Width           =   1050
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   2520
         TabIndex        =   46
         Top             =   1660
         Width           =   945
      End
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   60
      TabIndex        =   56
      Top             =   9750
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
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   13620
      Top             =   240
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frmCompras_NaoConformidade.frx":0000
      Count           =   1
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   65
      TabIndex        =   57
      Top             =   0
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   1720
      ButtonCount     =   7
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
      ButtonToolTipText1=   "Filtrar produto(s)/item(ns) devolvido(s)/refugado(s) (F2)"
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
      ButtonCaption3  =   "Recomprado"
      ButtonEnabled3  =   0   'False
      ButtonIconSize3 =   32
      ButtonToolTipText3=   "Filtrar produto(s)/item(ns) recomprado(s) (F7)"
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
      ButtonLeft3     =   80
      ButtonTop3      =   2
      ButtonWidth3    =   68
      ButtonHeight3   =   21
      ButtonUseMaskColor3=   0   'False
      ButtonEnabled4  =   0   'False
      ButtonIconSize4 =   32
      ButtonAlignment4=   2
      ButtonType4     =   1
      ButtonStyle4    =   -1
      BeginProperty ButtonFont4 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState4    =   -1
      ButtonLeft4     =   150
      ButtonTop4      =   4
      ButtonWidth4    =   2
      ButtonHeight4   =   54
      ButtonCaption5  =   "Ajuda"
      ButtonEnabled5  =   0   'False
      ButtonIconSize5 =   32
      ButtonToolTipText5=   "Ajuda (F1)"
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
      ButtonLeft5     =   154
      ButtonTop5      =   2
      ButtonWidth5    =   36
      ButtonHeight5   =   21
      ButtonUseMaskColor5=   0   'False
      ButtonCaption6  =   "Sair"
      ButtonEnabled6  =   0   'False
      ButtonIconSize6 =   32
      ButtonToolTipText6=   "Sair (Esc)"
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
      ButtonLeft6     =   192
      ButtonTop6      =   2
      ButtonWidth6    =   26
      ButtonHeight6   =   21
      ButtonUseMaskColor6=   0   'False
      ButtonEnabled7  =   0   'False
      ButtonIconSize7 =   32
      ButtonKey7      =   "7"
      ButtonAlignment7=   2
      BeginProperty ButtonFont7 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState7    =   5
      ButtonLeft7     =   220
      ButtonTop7      =   2
      ButtonWidth7    =   24
      ButtonHeight7   =   24
      ButtonUseMaskColor7=   0   'False
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Dados do produto/item"
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
      Height          =   1455
      Left            =   65
      TabIndex        =   39
      Top             =   1800
      Width           =   9165
      Begin VB.TextBox Txt_unidade 
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
         Left            =   3615
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Unidade."
         Top             =   430
         Width           =   600
      End
      Begin VB.TextBox txtInspecionada 
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
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade inspecionada."
         Top             =   430
         Width           =   2355
      End
      Begin VB.TextBox txtLote 
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
         Left            =   4230
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade."
         Top             =   430
         Width           =   2355
      End
      Begin VB.TextBox txtEspecificacoes 
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
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Descrição."
         Top             =   975
         Width           =   8775
      End
      Begin VB.TextBox txtNomenclatura 
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
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Código interno."
         Top             =   430
         Width           =   3420
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
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
         Left            =   3788
         TabIndex        =   44
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Qtde. insp."
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
         Left            =   7372
         TabIndex        =   43
         Top             =   240
         Width           =   810
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Qtde."
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
         Left            =   5197
         TabIndex        =   42
         Top             =   240
         Width           =   420
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Cód. interno*"
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
         Left            =   1328
         TabIndex        =   41
         Top             =   240
         Width           =   1125
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
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
         Left            =   4222
         TabIndex        =   40
         Top             =   780
         Width           =   690
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Laudo final de verificação"
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
      Height          =   885
      Left            =   65
      TabIndex        =   53
      Top             =   3240
      Width           =   9165
      Begin VB.TextBox Txt_liberado 
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
         Left            =   7050
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Liberado."
         Top             =   420
         Width           =   1905
      End
      Begin VB.TextBox Txt_laudo 
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
         MaxLength       =   255
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Laudo de inspeção."
         Top             =   420
         Width           =   6855
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Laudo"
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
         Left            =   3390
         TabIndex        =   55
         Top             =   240
         Width           =   435
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Liberado"
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
         Left            =   7695
         TabIndex        =   54
         Top             =   240
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmCompras_NaoConformidade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StrSql_comprasNC As String 'OK

Sub ProcAjuda()
On Error GoTo tratar_erro

FunAbrirVideoWeb ("http://www.youtube.com/watch?v=GJbpsWxB5rA&list=UUODGiDjQ-BCrxh0YXfJvoqg&index=15&feature=plcp")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaLista_ComprasNaoConformidade()
On Error GoTo tratar_erro

Lista_ComprasNaoConformidade.ListItems.Clear
If StrSql_comprasNC = "" Then Exit Sub
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open StrSql_comprasNC, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBAbrir.RecordCount
    PBLista.Value = 1
    contador = 0
    Do While TBAbrir.EOF = False
        With Lista_ComprasNaoConformidade.ListItems
            .Add , , TBAbrir!ID
            .Item(.Count).SubItems(1) = TBAbrir!IDEstoque
            .Item(.Count).SubItems(2) = IIf(IsNull(TBAbrir!LOTE), "", TBAbrir!LOTE)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBAbrir!Fornecedor), "", TBAbrir!Fornecedor)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBAbrir!Desenho), "", TBAbrir!Desenho)
            .Item(.Count).SubItems(5) = IIf(IsNull(TBAbrir!Descricao), "", TBAbrir!Descricao)
            .Item(.Count).SubItems(6) = IIf(IsNull(TBAbrir!Un), "", TBAbrir!Un)
            .Item(.Count).SubItems(7) = IIf(IsNull(TBAbrir!RJ), "", Format(TBAbrir!RJ, "###,##0.0000"))
        End With
        TBAbrir.MoveNext
        contador = contador + 1
        PBLista.Value = contador
    Loop
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyF2: ProcFiltrar
    Case vbKeyF3: ProcSalvar
    Case vbKeyF7: procFiltrar_Recomprado
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
Formulario = "Compras/Não conformidade"
Direitos
ProcLimpaVariaveisPrincipais

ProcRemoveObjetosResize Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

Formulario = "Compras/Não conformidade"
Direitos
ProcLimpaVariaveisPrincipais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

StrSql_comprasNC = "Select CR.*, EC.Lote, EC.Un, EC.Fornecedor, EC.id_cliente, EC.Desenho, EC.Descricao FROM Compras_recebimento CR INNER JOIN Estoque_controle EC on CR.IDestoque = EC.IDEstoque where CR.Recomprado = 'False' and LEFT(EC.status, 19) = 'ENTRADA_NOTA_FISCAL' and (CR.laudo = 'REFUGADO' or CR.laudo = 'DEVOLVIDO') order by EC.Lote, CR.IDestoque"
ProcCarregaLista_ComprasNaoConformidade

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procFiltrar_Recomprado()
On Error GoTo tratar_erro

StrSql_comprasNC = "Select CR.*, EC.Lote, EC.Un, EC.Fornecedor, EC.id_cliente, EC.Desenho, EC.Descricao FROM Compras_recebimento CR INNER JOIN Estoque_controle EC on CR.IDestoque = EC.IDEstoque where CR.Recomprado = 'True' and LEFT(EC.status, 19) = 'ENTRADA_NOTA_FISCAL' and (CR.laudo = 'REFUGADO' or CR.laudo = 'DEVOLVIDO') order by EC.Lote, CR.IDestoque"
ProcCarregaLista_ComprasNaoConformidade

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
If txtID = 0 Then
    USMsgBox ("Informe o produto na lista antes de alterar."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from compras_recebimento where ID = " & txtID, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = False Then
    If chkRecomprado.Value = 0 Then
        TBGravar!Recomprado = False
        TBGravar!Resp_recomprado = ""
        TBGravar!Data_Recomprado = Null
    Else
        TBGravar!Recomprado = True
        TBGravar!Resp_recomprado = pubUsuario
        TBGravar!Data_Recomprado = Format(Date, "dd/mm/yy")
    End If
    TBGravar.Update
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcCarregaLista_ComprasNaoConformidade
    '==================================
    Modulo = "Compras/Não conformidade"
    Evento = "Salvar"
    ID_documento = txtID
    Documento = "Pedido: " & txtPedido.Text & "- Cód. interno: " & txtCodinterno
    Documento1 = ""
    ProcGravaEvento
    '==================================
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_ComprasNaoConformidade_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView Lista_ComprasNaoConformidade, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_ComprasNaoConformidade_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista_ComprasNaoConformidade.ListItems.Count = 0 Then Exit Sub
ProcLimpaCamposPrincipal
ProcLimpaCampos
Set TBRecebidos = CreateObject("adodb.recordset")
TBRecebidos.Open "Select CR.*, EC.Lote, EC.Un, EC.Fornecedor, EC.id_cliente, EC.Desenho, EC.Descricao FROM Compras_recebimento CR INNER JOIN Estoque_controle EC on CR.IDestoque = EC.IDEstoque where CR.Id = " & Lista_ComprasNaoConformidade.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBRecebidos.EOF = False Then
    txtPedido = TBRecebidos!LOTE
    mskData = IIf(IsNull(TBRecebidos!Data), "", Format(TBRecebidos!Data, "dd/mm/yy"))
    txtResponsável = IIf(IsNull(TBRecebidos!Responsavel), "", TBRecebidos!Responsavel)
    txtID = TBRecebidos!ID
    txtFornecedor = TBRecebidos!Fornecedor
    txtNomenclatura.Text = TBRecebidos!Desenho
    Txt_unidade = IIf(IsNull(TBRecebidos!Un), "", TBRecebidos!Un)
    txtLote.Text = Format(TBRecebidos!quantidade, "###,##0.0000")
    txtInspecionada = Format(TBRecebidos!Enc, "###,##0.0000")
    txtEspecificacoes.Text = TBRecebidos!Descricao
    txtObservacoes.Text = TBRecebidos!Obs
    If IsNull(TBRecebidos!Laudo) = False And TBRecebidos!Laudo <> "" Then Txt_laudo.Text = TBRecebidos!Laudo
    If IsNull(TBRecebidos!Liberado) = False And TBRecebidos!Liberado <> "" Then Txt_liberado.Text = TBRecebidos!Liberado
    ProcEntradaList
    If TBRecebidos!Recomprado = True Then chkRecomprado.Value = 1 Else chkRecomprado.Value = 0
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcEntradaList()
On Error GoTo tratar_erro

'==========================
' Entrada de embalagem
'==========================
Select Case TBRecebidos!Embalagem
    Case 1: chk_emb_ok.Value = 1
    Case 2: chk_emb_nc.Value = 1
    Case 3: chk_emb_na.Value = 1
End Select
'==========================
' Entrada de laudo
'==========================
Select Case TBRecebidos!Laudos
    Case 1: chk_laudo_ok.Value = 1
    Case 2: chk_laudo_nc.Value = 1
    Case 3: chk_laudo_na.Value = 1
End Select
'==========================
' Entrada de quantidade
'==========================
Select Case TBRecebidos!quantidade
    Case 1: chk_qtd_ok.Value = 1
    Case 2: chk_qtd_nc.Value = 1
    Case 3: chk_qtd_na.Value = 1
End Select
'==========================
' Entrada de visual
'==========================
Select Case TBRecebidos!Visual
    Case 1: chk_visual_ok.Value = 1
    Case 2: chk_visual_nc.Value = 1
    Case 3: chk_visual_na.Value = 1
End Select
'==========================
' Entrada de dimensional
'==========================
Select Case TBRecebidos!dimensional
    Case 1: chk_dim_ok.Value = 1
    Case 2: chk_dim_nc.Value = 1
    Case 3: chk_dim_na.Value = 1
End Select
'==========================
' Entrada de outros
'==========================
Select Case TBRecebidos!Outros
    Case 1: chk_outros_ok.Value = 1
    Case 2: chk_outros_nc.Value = 1
    Case 3: chk_outros_na.Value = 1
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCamposPrincipal()
On Error GoTo tratar_erro

txtPedido.Text = ""
txtResponsável.Text = pubUsuario
txtLote.Text = "0,0000"
txtInspecionada = "0,0000"
mskData.Text = Format(Date, "dd/mm/yy")
txtFornecedor.Text = ""
txtNomenclatura.Text = ""
txtEspecificacoes.Text = ""
Txt_unidade = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCampos()
On Error GoTo tratar_erro

txtID = 0
txtObservacoes.Text = ""
chk_dim_na.Value = 0
chk_dim_nc.Value = 0
chk_dim_ok.Value = 0
chk_emb_na.Value = 0
chk_emb_nc.Value = 0
chk_emb_ok.Value = 0
chk_laudo_na.Value = 0
chk_laudo_nc.Value = 0
chk_laudo_ok.Value = 0
chk_visual_na.Value = 0
chk_visual_nc.Value = 0
chk_visual_ok.Value = 0
chk_outros_na.Value = 0
chk_outros_nc.Value = 0
chk_outros_ok.Value = 0
chk_qtd_na.Value = 0
chk_qtd_nc.Value = 0
chk_qtd_ok.Value = 0
Txt_laudo = ""
Txt_liberado = ""
chkRecomprado.Value = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcFiltrar
    Case 2: ProcSalvar
    Case 3: procFiltrar_Recomprado
    Case 5: ProcAjuda
    Case 6: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
