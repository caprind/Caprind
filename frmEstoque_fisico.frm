VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmestoque_fisico 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Estoque -  Inventário"
   ClientHeight    =   10035
   ClientLeft      =   1695
   ClientTop       =   1335
   ClientWidth     =   15270
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10035
   ScaleWidth      =   15270
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
      FormWidthDT     =   15390
      FormScaleHeightDT=   10035
      FormScaleWidthDT=   15270
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   55
      TabIndex        =   59
      Top             =   8250
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
         ItemData        =   "frmEstoque_fisico.frx":0000
         Left            =   6960
         List            =   "frmEstoque_fisico.frx":000A
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   180
         Width           =   1965
      End
      Begin VB.TextBox txtPagIr 
         Height          =   315
         Left            =   9540
         TabIndex        =   25
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
         Left            =   2730
         TabIndex        =   23
         Text            =   "30"
         ToolTipText     =   "Número de registros por página."
         Top             =   180
         Width           =   555
      End
      Begin DrawSuite2022.USButton cmdPagProx 
         Height          =   315
         Left            =   11760
         TabIndex        =   29
         ToolTipText     =   "Próxima página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmEstoque_fisico.frx":0022
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
         TabIndex        =   28
         ToolTipText     =   "Página anterior."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmEstoque_fisico.frx":37C6
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
         TabIndex        =   26
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
         TabIndex        =   27
         ToolTipText     =   "Primeira página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmEstoque_fisico.frx":72CF
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
         TabIndex        =   30
         ToolTipText     =   "Última página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmEstoque_fisico.frx":B3BE
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
      Begin VB.Label Label26 
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
         Left            =   3360
         TabIndex        =   73
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
         Left            =   5610
         TabIndex        =   67
         Top             =   180
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
         TabIndex        =   62
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
         TabIndex        =   61
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label Label6 
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
         Left            =   2040
         TabIndex        =   60
         Top             =   240
         Width           =   645
      End
   End
   Begin VB.Frame Frame12 
      BackColor       =   &H00E0E0E0&
      Height          =   855
      Left            =   55
      TabIndex        =   49
      Top             =   8880
      Width           =   15195
      Begin VB.TextBox Txt_valor_total_fisica 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   3735
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   35
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor total físico."
         Top             =   390
         Width           =   1725
      End
      Begin VB.TextBox Txt_valor_total_unitario 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   31
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor total unitário."
         Top             =   390
         Width           =   1725
      End
      Begin VB.TextBox Txt_qtde_total_estoque 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   13290
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   32
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Quantidade total estoque."
         Top             =   390
         Width           =   1725
      End
      Begin VB.TextBox Txt_valor_total_estoque 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   1965
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   33
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor total estoque."
         Top             =   390
         Width           =   1725
      End
      Begin VB.TextBox Txt_qtde_total_fisica 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   11505
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   34
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Quantidade total físico."
         Top             =   390
         Width           =   1725
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Valor total inventário"
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
         Left            =   3900
         TabIndex        =   54
         Top             =   180
         Width           =   1500
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Valor total estoque"
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
         Left            =   2100
         TabIndex        =   53
         Top             =   180
         Width           =   1365
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total em inventário"
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
         Left            =   11730
         TabIndex        =   52
         Top             =   180
         Width           =   1380
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total em estoque"
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
         Left            =   13515
         TabIndex        =   51
         Top             =   180
         Width           =   1245
      End
      Begin VB.Label Label42 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Valor total unitário"
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
         Left            =   465
         TabIndex        =   50
         Top             =   180
         Width           =   1320
      End
   End
   Begin VB.TextBox Txt_ID 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1200
      TabIndex        =   46
      Text            =   "0"
      Top             =   5310
      Visible         =   0   'False
      Width           =   675
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
      Height          =   3225
      Left            =   55
      TabIndex        =   36
      Top             =   990
      Width           =   15195
      Begin VB.TextBox txt_LA 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
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
         Left            =   10950
         Locked          =   -1  'True
         MaxLength       =   150
         TabIndex        =   82
         TabStop         =   0   'False
         ToolTipText     =   "Família do item"
         Top             =   2130
         Width           =   4065
      End
      Begin VB.TextBox txt_cod_Referencia 
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
         MaxLength       =   50
         TabIndex        =   81
         ToolTipText     =   "Código de referência do item"
         Top             =   1560
         Width           =   1965
      End
      Begin VB.TextBox txt_RE 
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
         Left            =   8140
         MaxLength       =   50
         TabIndex        =   80
         ToolTipText     =   "Numero da RE (Registro de estoque)"
         Top             =   2130
         Width           =   890
      End
      Begin VB.TextBox Txt_valor_total 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
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
         Left            =   13440
         Locked          =   -1  'True
         TabIndex        =   76
         TabStop         =   0   'False
         ToolTipText     =   "Valor unitário."
         Top             =   2730
         Width           =   1560
      End
      Begin MSComDlg.CommonDialog CD1 
         Left            =   1950
         Top             =   90
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.TextBox Txt_cli_forn 
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
         Left            =   2925
         Locked          =   -1  'True
         TabIndex        =   21
         ToolTipText     =   "Nome do fornecedor."
         Top             =   1020
         Width           =   11760
      End
      Begin VB.TextBox Txt_ID_cli_forn 
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
         Left            =   2160
         MaxLength       =   20
         TabIndex        =   20
         ToolTipText     =   "Código do fornecedor."
         Top             =   1020
         Width           =   750
      End
      Begin VB.ComboBox Cmb_tipo_cli_forn 
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
         Height          =   330
         ItemData        =   "frmEstoque_fisico.frx":EC4A
         Left            =   180
         List            =   "frmEstoque_fisico.frx":EC57
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   19
         ToolTipText     =   "Tipo."
         Top             =   1020
         Width           =   1965
      End
      Begin VB.TextBox Txt_numero_Serie 
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
         Left            =   4930
         MaxLength       =   60
         TabIndex        =   6
         ToolTipText     =   "Número de série."
         Top             =   2130
         Width           =   1670
      End
      Begin VB.CheckBox Chk_consignado 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Material de terceiro"
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
         Left            =   2400
         TabIndex        =   18
         Top             =   810
         Width           =   1695
      End
      Begin VB.TextBox txtRespValidacao 
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
         Left            =   11985
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Responsável pela validação."
         Top             =   375
         Width           =   3030
      End
      Begin VB.TextBox txtDtValidacao 
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
         Left            =   9940
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Data de validação."
         Top             =   375
         Width           =   2025
      End
      Begin VB.ComboBox cmbDestino 
         BackColor       =   &H00C0E0FF&
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
         ItemData        =   "frmEstoque_fisico.frx":EC72
         Left            =   4910
         List            =   "frmEstoque_fisico.frx":EC7C
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   14
         ToolTipText     =   "Destino."
         Top             =   2730
         Width           =   3225
      End
      Begin VB.ComboBox Cmb_empresa 
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
         ItemData        =   "frmEstoque_fisico.frx":ECAC
         Left            =   180
         List            =   "frmEstoque_fisico.frx":ECAE
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Empresa."
         Top             =   375
         Width           =   5475
      End
      Begin VB.TextBox Txt_un 
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
         Left            =   14040
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Unidade."
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox Txt_valor_unitario 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
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
         Left            =   9555
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Valor unitário."
         Top             =   2730
         Width           =   1320
      End
      Begin VB.TextBox Txt_familia 
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   180
         Locked          =   -1  'True
         MaxLength       =   150
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Família."
         Top             =   2130
         Width           =   4740
      End
      Begin VB.TextBox Txt_responsavel 
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
         Left            =   6880
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Responsável pelo cadastro."
         Top             =   375
         Width           =   3045
      End
      Begin VB.TextBox Txt_qtde_estoque 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
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
         Left            =   11775
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade em estoque."
         Top             =   2730
         Width           =   1410
      End
      Begin VB.TextBox Txt_qtde_fisica 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
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
         Left            =   8145
         TabIndex        =   17
         ToolTipText     =   "Quantidade no inventário."
         Top             =   2730
         Width           =   1410
      End
      Begin VB.TextBox Txt_etiqueta 
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   6615
         MaxLength       =   60
         TabIndex        =   10
         ToolTipText     =   "Número da etiqueta."
         Top             =   2130
         Width           =   1515
      End
      Begin VB.TextBox Txt_descricao 
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
         Left            =   3780
         Locked          =   -1  'True
         MaxLength       =   150
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Descrição."
         Top             =   1560
         Width           =   10245
      End
      Begin VB.TextBox Txt_cod_interno 
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
         TabIndex        =   5
         ToolTipText     =   "Código interno."
         Top             =   1560
         Width           =   1605
      End
      Begin MSComCtl2.DTPicker Txt_data 
         Height          =   315
         Left            =   5650
         TabIndex        =   1
         ToolTipText     =   "Data."
         Top             =   375
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
         Format          =   197918721
         CurrentDate     =   39057
      End
      Begin VB.TextBox Txt_certificado 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
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
         Left            =   2540
         MaxLength       =   60
         TabIndex        =   13
         ToolTipText     =   "Número do certificado."
         Top             =   2730
         Width           =   2345
      End
      Begin VB.TextBox Txt_corrida 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
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
         MaxLength       =   60
         TabIndex        =   12
         ToolTipText     =   "Número da corrida."
         Top             =   2730
         Width           =   2345
      End
      Begin VB.TextBox Txt_lote 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
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
         Left            =   9045
         MaxLength       =   60
         TabIndex        =   11
         ToolTipText     =   "Número do lote."
         Top             =   2130
         Width           =   1890
      End
      Begin DrawSuite2022.USButton Cmd_localizar_cli_forn 
         Height          =   315
         Left            =   14700
         TabIndex        =   78
         ToolTipText     =   "Localizar cliente ou fornecedor"
         Top             =   1020
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         DibPicture      =   "frmEstoque_fisico.frx":ECB0
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
         PicAlign        =   0
         ShowFocusRect   =   0   'False
         Theme           =   1
         ToolTipTitle    =   "CAPRIND v5.0"
      End
      Begin DrawSuite2022.USButton btnValor 
         Height          =   315
         Left            =   10890
         TabIndex        =   79
         ToolTipText     =   "Buscar valor unitário do cadastro do produto"
         Top             =   2730
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         DibPicture      =   "frmEstoque_fisico.frx":2CDB5
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
         PicAlign        =   0
         ShowFocusRect   =   0   'False
         Theme           =   1
         ToolTipTitle    =   "CAPRIND v5.0"
      End
      Begin VB.TextBox Txt_ID_prod 
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
         TabIndex        =   47
         Text            =   "0"
         Top             =   1560
         Visible         =   0   'False
         Width           =   345
      End
      Begin DrawSuite2022.USButton btnSalvarValorUnitario 
         Height          =   315
         Left            =   11220
         TabIndex        =   83
         ToolTipText     =   "Salvar valor unitário do produto na RE (Inventário)"
         Top             =   2730
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         DibPicture      =   "frmEstoque_fisico.frx":33F48
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
         PicAlign        =   0
         ShowFocusRect   =   0   'False
         Theme           =   1
         ToolTipTitle    =   "CAPRIND v5.0"
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "="
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   13230
         TabIndex        =   85
         Top             =   2730
         Width           =   180
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   11580
         TabIndex        =   84
         Top             =   2790
         Width           =   135
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Total"
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
         TabIndex        =   77
         Top             =   2550
         Width           =   795
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente/Fornecedor"
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
         Index           =   7
         Left            =   8100
         TabIndex        =   75
         Top             =   810
         Width           =   1380
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo"
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
         Index           =   5
         Left            =   975
         TabIndex        =   74
         Top             =   810
         Width           =   300
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Número de série"
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
         Left            =   5205
         TabIndex        =   72
         Top             =   1935
         Width           =   1170
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cód. de referência"
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
         Left            =   2107
         TabIndex        =   71
         Top             =   1365
         Width           =   1350
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N° do RE"
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
         Left            =   8295
         TabIndex        =   70
         Top             =   1935
         Width           =   645
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
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
         Index           =   4
         Left            =   12683
         TabIndex        =   69
         Top             =   180
         Width           =   1635
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
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
         Index           =   3
         Left            =   10112
         TabIndex        =   68
         Top             =   180
         Width           =   1680
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Destino*"
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
         Height          =   195
         Left            =   6232
         TabIndex        =   66
         Top             =   2535
         Width           =   630
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N° do certificado*"
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
         Height          =   195
         Left            =   3075
         TabIndex        =   65
         Top             =   2535
         Width           =   1290
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N° do lote*"
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
         Height          =   195
         Left            =   9570
         TabIndex        =   64
         Top             =   1935
         Width           =   825
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N° da corrida*"
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
         Height          =   195
         Left            =   840
         TabIndex        =   63
         Top             =   2535
         Width           =   1035
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Empresa"
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
         Left            =   2550
         TabIndex        =   58
         Top             =   180
         Width           =   735
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Local de armazenamento*"
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
         Height          =   195
         Left            =   12150
         TabIndex        =   55
         Top             =   1935
         Width           =   1875
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor unitário*"
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
         Height          =   195
         Left            =   9720
         TabIndex        =   48
         Top             =   2535
         Width           =   1065
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
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
         Index           =   2
         Left            =   9135
         TabIndex        =   45
         Top             =   1365
         Width           =   690
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cód. interno*"
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
         Height          =   195
         Index           =   1
         Left            =   480
         TabIndex        =   44
         Top             =   1365
         Width           =   990
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
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
         Index           =   0
         Left            =   2340
         TabIndex        =   43
         Top             =   1935
         Width           =   480
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
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
         Index           =   14
         Left            =   6085
         TabIndex        =   42
         Top             =   180
         Width           =   345
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
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
         Index           =   15
         Left            =   7950
         TabIndex        =   41
         Top             =   180
         Width           =   915
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo inventário*"
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
         Height          =   195
         Left            =   8235
         TabIndex        =   40
         Top             =   2535
         Width           =   1275
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo atual"
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
         Left            =   12075
         TabIndex        =   39
         Top             =   2535
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N° da etiqueta"
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
         TabIndex        =   38
         Top             =   1935
         Width           =   1050
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unidade"
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
         Left            =   14280
         TabIndex        =   37
         Top             =   1365
         Width           =   585
      End
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   55
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
      SearchText      =   ""
      Value           =   0
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   57
      Top             =   0
      Width           =   15165
      _ExtentX        =   26749
      _ExtentY        =   1720
      ButtonCount     =   11
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
      ButtonCaption6  =   "Validação"
      ButtonEnabled6  =   0   'False
      ButtonIconSize6 =   32
      ButtonToolTipText6=   "Validar quantidade física no estoque real (F7)"
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
      ButtonLeft6     =   209
      ButtonTop6      =   2
      ButtonWidth6    =   53
      ButtonHeight6   =   21
      ButtonUseMaskColor6=   0   'False
      ButtonCaption7  =   "Importar Excell"
      ButtonEnabled7  =   0   'False
      ButtonIconSize7 =   32
      ButtonToolTipText7=   "Importar planilha Excell"
      ButtonKey7      =   "7"
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
      ButtonLeft7     =   264
      ButtonTop7      =   2
      ButtonWidth7    =   80
      ButtonHeight7   =   21
      ButtonUseMaskColor7=   0   'False
      ButtonEnabled8  =   0   'False
      ButtonIconSize8 =   32
      ButtonAlignment8=   2
      ButtonType8     =   1
      ButtonStyle8    =   -1
      BeginProperty ButtonFont8 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState8    =   -1
      ButtonLeft8     =   346
      ButtonTop8      =   4
      ButtonWidth8    =   2
      ButtonHeight8   =   54
      ButtonCaption9  =   "Ajuda"
      ButtonEnabled9  =   0   'False
      ButtonIconSize9 =   32
      ButtonToolTipText9=   "Ajuda (F1)"
      ButtonKey9      =   "8"
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
      ButtonLeft9     =   350
      ButtonTop9      =   2
      ButtonWidth9    =   36
      ButtonHeight9   =   21
      ButtonUseMaskColor9=   0   'False
      ButtonCaption10 =   "Sair"
      ButtonEnabled10 =   0   'False
      ButtonIconSize10=   32
      ButtonToolTipText10=   "Sair (Esc)"
      ButtonKey10     =   "9"
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
      ButtonLeft10    =   388
      ButtonTop10     =   2
      ButtonWidth10   =   26
      ButtonHeight10  =   21
      ButtonUseMaskColor10=   0   'False
      ButtonEnabled11 =   0   'False
      ButtonIconSize11=   32
      ButtonKey11     =   "10"
      ButtonAlignment11=   2
      BeginProperty ButtonFont11 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState11   =   5
      ButtonLeft11    =   416
      ButtonTop11     =   2
      ButtonWidth11   =   24
      ButtonHeight11  =   24
      ButtonUseMaskColor11=   0   'False
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   6840
         Top             =   180
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmEstoque_fisico.frx":3C94D
         Count           =   1
      End
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   4035
      Left            =   60
      TabIndex        =   22
      Top             =   4200
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   7117
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
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   14
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Object.Tag             =   "T"
         Text            =   "RE"
         Object.Width           =   1234
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Object.Tag             =   "D"
         Text            =   "Data"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "Cód. interno"
         Object.Width           =   2116
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Cód. ref."
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Tag             =   "T"
         Text            =   "Descrição"
         Object.Width           =   8388
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Object.Tag             =   "T"
         Text            =   "Un."
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Object.Tag             =   "N"
         Text            =   "Valor unitário"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Object.Tag             =   "N"
         Text            =   "Saldo atual"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Object.Tag             =   "N"
         Text            =   "Valor total"
         Object.Width           =   2470
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Object.Tag             =   "N"
         Text            =   "Saldo. inventário"
         Object.Width           =   2734
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   11
         Text            =   "Valor inventário"
         Object.Width           =   2381
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   12
         Text            =   "Etiqueta"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   13
         Text            =   "Validado"
         Object.Width           =   1411
      EndProperty
   End
End
Attribute VB_Name = "frmestoque_fisico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Novo_Estoque_Fisico                     As Boolean 'OK
Public Sql_Estoque_Fisico_Localizar         As String 'OK
Public Sql_Estoque_Fisico_LocalizarTotal    As String 'OK
Public FormulaRel_Estoque_Fisico            As String 'OK
Dim TBLISTA_Estoque_Fisico                  As ADODB.Recordset 'OK
Dim QtdPC                                   As Double 'OK
Dim QtdePC                                  As Double 'OK
Dim Tipodest_NFcons                         As String 'OK
Dim xl As New Excel.Application
Dim xlw As Excel.Workbook

Sub ProcAtualizalista(Pagina As Integer)
On Error GoTo tratar_erro

lblRegistros.Caption = "Nº de registros: 0"
lblPaginas.Caption = "Página: 0 de: 0"
Lista.ListItems.Clear

Conexao.Execute ("update Estoque_fisico set IDestoque = Em.IDestoque from Estoque_fisico EF inner join Estoque_movimentacao EM on EF.ID = EM.ID_inventario")


If Sql_Estoque_Fisico_Localizar = "" Then Exit Sub
Set TBLISTA_Estoque_Fisico = CreateObject("adodb.recordset")
TBLISTA_Estoque_Fisico.Open Sql_Estoque_Fisico_Localizar, Conexao, adOpenKeyset, adLockReadOnly
If TBLISTA_Estoque_Fisico.EOF = False Then ProcExibePagina (1)
'ProcCarregaTotal

TabelaRel = 1
OrdenarRel = 1
CampoRel = 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina(Pagina)
On Error GoTo tratar_erro

Lista.ListItems.Clear
TBLISTA_Estoque_Fisico.PageSize = IIf(txtNreg = "", 30, txtNreg)
TBLISTA_Estoque_Fisico.AbsolutePage = Pagina
TamanhoPagina = TBLISTA_Estoque_Fisico.PageSize
ContadorReg = 1

PBLista.Min = 0
PBLista.Max = FunVerifMaxPBListaPaginacao(TBLISTA_Estoque_Fisico.RecordCount - IIf(Pagina > 1, (TBLISTA_Estoque_Fisico.PageSize * (Pagina - 1)), 0), TBLISTA_Estoque_Fisico.PageSize)
PBLista.Value = 1
Contador = 0
Saldo_Total = 0
Valor_Unitario_Total = 0
Valor_Lote_Total = 0
Saldo_total_Fisico = 0
Valor_Total_Fisico = 0



Do While TBLISTA_Estoque_Fisico.EOF = False And (ContadorReg <= TamanhoPagina)

'===============================================================
' Busca Saldo da RE na movimentação
'===============================================================
Set TBEstoque = CreateObject("adodb.recordset")
StrSql = "Select  ISNULL(Sum(Entrada) - Sum(Saida), 0) As Saldo from Estoque_Movimentacao where IDEstoque =  " & IIf(IsNull(TBLISTA_Estoque_Fisico!IDEstoque), 0, TBLISTA_Estoque_Fisico!IDEstoque)

TBEstoque.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
If TBEstoque.EOF = False Then
Saldo_Atual = TBEstoque!Saldo
End If
TBEstoque.Close
'===============================================================

    With Lista.ListItems
        .Add , , TBLISTA_Estoque_Fisico!ID
        .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA_Estoque_Fisico!IDEstoque), 0, TBLISTA_Estoque_Fisico!IDEstoque)
        .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA_Estoque_Fisico!Data), "", Format(TBLISTA_Estoque_Fisico!Data, "dd/mm/yy"))
        .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA_Estoque_Fisico!Desenho), "", TBLISTA_Estoque_Fisico!Desenho)
        .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA_Estoque_Fisico!Cod_ref), "", TBLISTA_Estoque_Fisico!Cod_ref)
        .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA_Estoque_Fisico!Descricao), "", TBLISTA_Estoque_Fisico!Descricao)
        .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA_Estoque_Fisico!Unidade), "", TBLISTA_Estoque_Fisico!Unidade)
        
        'Valor unitário
        .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA_Estoque_Fisico!valor_unitario), "", Format(TBLISTA_Estoque_Fisico!valor_unitario, "###,##0.000000"))
        valor_unitario = IIf(IsNull(TBLISTA_Estoque_Fisico!valor_unitario), 0, TBLISTA_Estoque_Fisico!valor_unitario)
        
        'Qtde. em estoque
        .Item(.Count).SubItems(8) = Format(Saldo_Atual, "###,##0.0000")
          
        'Valor total estoque
        .Item(.Count).SubItems(9) = Format(valor_unitario * Saldo_Atual, "###,##0.000000")
        Valor_Lote = Format(valor_unitario * Saldo_Atual, "###,##0.000000")
        
        'Qtde. físico
        .Item(.Count).SubItems(10) = IIf(IsNull(TBLISTA_Estoque_Fisico!qtde_fisica), "", Format(TBLISTA_Estoque_Fisico!qtde_fisica, "###,##0.0000"))
        Saldo_fisico = IIf(IsNull(TBLISTA_Estoque_Fisico!qtde_fisica), 0, TBLISTA_Estoque_Fisico!qtde_fisica)
        
        'Valor total físico
        .Item(.Count).SubItems(11) = Format(valor_unitario * Saldo_fisico, "###,##0.000000")
        Valor_Fisico = Format(valor_unitario * Saldo_fisico, "###,##0.000000")
        
        .Item(.Count).SubItems(12) = IIf(IsNull(TBLISTA_Estoque_Fisico!Etiqueta), "", TBLISTA_Estoque_Fisico!Etiqueta)
        .Item(.Count).SubItems(13) = IIf(IsNull(TBLISTA_Estoque_Fisico!DtValidacao) = False, "SIM", "NÃO")
        
        Saldo_Total = Saldo_Total + Saldo_Atual
        Valor_Unitario_Total = Valor_Unitario_Total + valor_unitario
        Valor_Lote_Total = Valor_Lote_Total + Valor_Lote
        Saldo_total_Fisico = Saldo_total_Fisico + Saldo_fisico
        Valor_Total_Fisico = Valor_Total_Fisico + Valor_Fisico
        
    End With
    TBLISTA_Estoque_Fisico.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista.Value = Contador
Loop

Txt_qtde_total_estoque = Format(Saldo_Total, "###,##0.0000")

Txt_valor_total_unitario = Format(Valor_Unitario_Total, "###,##0.0000")
Txt_valor_total_estoque = Format(Valor_Lote_Total, "###,##0.0000")
Txt_qtde_total_fisica = Format(Saldo_total_Fisico, "###,##0.0000")
Txt_valor_total_fisica = Format(Valor_Total_Fisico, "###,##0.0000")

lblRegistros.Caption = "Nº de registros: " & TBLISTA_Estoque_Fisico.RecordCount
If TBLISTA_Estoque_Fisico.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "Página: 1 de: " & TBLISTA_Estoque_Fisico.PageCount
ElseIf TBLISTA_Estoque_Fisico.AbsolutePage = adPosEOF Then
        lblPaginas.Caption = "Página: " & TBLISTA_Estoque_Fisico.PageCount & " de: " & TBLISTA_Estoque_Fisico.PageCount
    Else
        lblPaginas.Caption = "Página: " & TBLISTA_Estoque_Fisico.AbsolutePage - 1 & " de: " & TBLISTA_Estoque_Fisico.PageCount
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaTotal()
On Error GoTo tratar_erro

Valor_Cofins_Prod = 0
Valor_Cofins_Serv = 0
Valor_CSLL_Prod = 0
Valor_CSLL_Serv = 0
Valor_INSS_Serv = 0
Valor_IPI = 0
Valor_IRPJ_Prod = 0
Set TBContas = CreateObject("adodb.recordset")
TBContas.Open Sql_Estoque_Fisico_LocalizarTotal, Conexao, adOpenKeyset, adLockOptimistic
If TBContas.EOF = False Then
    Valor_Cofins_Prod = IIf(IsNull(TBContas!Valor_Cofins_Prod), 0, TBContas!Valor_Cofins_Prod)
    Valor_Cofins_Serv = IIf(IsNull(TBContas!Valor_Cofins_Serv), 0, TBContas!Valor_Cofins_Serv)
    Valor_CSLL_Prod = IIf(IsNull(TBContas!Valor_CSLL_Prod), 0, TBContas!Valor_CSLL_Prod)
    Valor_CSLL_Serv = IIf(IsNull(TBContas!Valor_CSLL_Serv), 0, TBContas!Valor_CSLL_Serv)
    Valor_INSS_Serv = IIf(IsNull(TBContas!Valor_INSS_Serv), 0, TBContas!Valor_INSS_Serv)
    Valor_IPI = IIf(IsNull(TBContas!Valor_IPI), 0, TBContas!Valor_IPI)
    Valor_IRPJ_Prod = IIf(IsNull(TBContas!Valor_IRPJ_Prod), 0, TBContas!Valor_IRPJ_Prod)
End If
TBContas.Close

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
                If USMsgBox("Deseja realmente excluir este(s) inventário(s)?" & vbCrLf & vbCrLf & "Lembre-se que serão excluidos a ficha de estoque, e a movimentação de entrada do inventário.", vbYesNo, "CAPRIND v5.0") = vbNo Then
                Exit Sub
                End If
            End If
            Permitido = True
            '==================================
            Modulo = "Estoque/Inventário"
            Evento = "Excluir"
            ID_documento = .ListItems(InitFor)
            Documento = "Data: " & .ListItems(InitFor).ListSubItems(2) & " - Código interno: " & .ListItems(InitFor).ListSubItems(3)
            Documento1 = ""
            ProcGravaEvento
            '==================================
            Conexao.Execute "DELETE from Estoque_fisico where id = " & .ListItems(InitFor)
            Conexao.Execute "DELETE from Estoque_controle where idestoque = " & .ListItems(InitFor).ListSubItems(1)
            Conexao.Execute "DELETE from Estoque_movimentacao where idestoque = " & .ListItems(InitFor).ListSubItems(1)
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) inventário(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Inventário(s) excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCampos
    ProcAtualizalista (1)
    Frame1.Enabled = False
    Novo_Estoque_Fisico = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

frmEstoque_fisico_abrir.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovo()
On Error GoTo tratar_erro
  
If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & ", voce não está autorizado a criar novo cadastro neste formulário."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
ProcLimpaCampos
Novo_Estoque_Fisico = True
Frame1.Enabled = True
cmbDestino = "Interno/Cliente"

frmEstoque_fisico_Novo.Show 1

Txt_data.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimpaCampos()
On Error GoTo tratar_erro

Txt_ID = 0
Txt_data = Date
Txt_responsavel = pubUsuario
txtDtValidacao = ""
txtRespValidacao = ""
Txt_cod_interno = ""
txt_cod_Referencia = ""
Txt_numero_serie = ""
Txt_descricao = ""
Txt_un = ""
Txt_familia = ""
Txt_etiqueta = ""
txt_RE.Text = ""
Txt_lote = ""
txt_LA.Text = ""
txt_Corrida = ""
txt_Certificado = ""
Txt_valor_unitario = ""
Txt_valor_total = ""
cmbDestino.ListIndex = -1
Txt_qtde_estoque = ""
Txt_qtde_fisica = ""
Chk_consignado.Value = 0
Cmb_tipo_cli_forn.ListIndex = -1
Txt_ID_cli_forn = ""
Txt_cli_forn = ""
Frame1.Enabled = False
CodigoLista = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcPuxaDados()
On Error GoTo tratar_erro

If IsNull(TBProduto!ID_empresa) = False And TBProduto!ID_empresa <> "" Then
ProcPuxaDadosComboEmpresa Cmb_empresa, TBProduto!ID_empresa
End If

Txt_ID = TBProduto!ID

Txt_data = IIf(IsNull(TBProduto!Data), "", TBProduto!Data)
Txt_responsavel = IIf(IsNull(TBProduto!Responsavel), "", TBProduto!Responsavel)
Txt_ID_prod = TBProduto!Codproduto
Txt_cod_interno = IIf(IsNull(TBProduto!Desenho), "", TBProduto!Desenho)
txt_cod_Referencia = IIf(IsNull(TBProduto!Cod_ref), "", TBProduto!Cod_ref)
Txt_descricao = IIf(IsNull(TBProduto!Descricao), "", TBProduto!Descricao)
Txt_un = IIf(IsNull(TBProduto!Unidade), "", TBProduto!Unidade)
Txt_familia = IIf(IsNull(TBProduto!Classe), "", TBProduto!Classe)
Txt_etiqueta = IIf(IsNull(TBProduto!Etiqueta), "", TBProduto!Etiqueta)
txt_RE.Text = IIf(IsNull(TBProduto!IDEstoque), "", TBProduto!IDEstoque)
Txt_lote.Text = IIf(IsNull(TBProduto!LOTE), "", TBProduto!LOTE)
txt_LA.Text = IIf(IsNull(TBProduto!local_armaz), "", TBProduto!local_armaz)
txt_Corrida = IIf(IsNull(TBProduto!Corrida), "", TBProduto!Corrida)
txt_Certificado = IIf(IsNull(TBProduto!Certificado), "", TBProduto!Certificado)
Txt_numero_serie = IIf(IsNull(TBProduto!Numero_serie), "", TBProduto!Numero_serie)

Txt_valor_unitario = IIf(IsNull(TBProduto!valor_unitario), "", Format(TBProduto!valor_unitario, "###,##0.000000"))
valor_unitario = IIf(Txt_valor_unitario <> "", Txt_valor_unitario, 0)

cmbDestino.Text = IIf(TBProduto!Destino = "Interno", "Interno/Cliente", "Terceiros (Remessa forn.)")

Txt_qtde_estoque = Format(Saldo_Atual, "###,##0.0000")
Txt_valor_total.Text = Format(Saldo_Atual * valor_unitario, "###,##0.000000")

Txt_qtde_fisica = IIf(IsNull(TBProduto!qtde_fisica), "", Format(TBProduto!qtde_fisica, "###,##0.0000"))
txtRespValidacao = IIf(IsNull(TBProduto!RespValidacao), "", TBProduto!RespValidacao)
txtDtValidacao = IIf(IsNull(TBProduto!DtValidacao), "", TBProduto!DtValidacao)

If TBProduto!Consignado = True Then
Chk_consignado.Value = 1
Else
Chk_consignado.Value = 0
End If

If IsNull(TBProduto!Tipo_cli_forn) = False And TBProduto!Tipo_cli_forn <> "" Then
Cmb_tipo_cli_forn = IIf(TBProduto!Tipo_cli_forn = "C", "Cliente", "Fornecedor")
End If

Txt_ID_cli_forn = IIf(IsNull(TBProduto!ID_cli_forn), "", TBProduto!ID_cli_forn)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub procValidar(Validar As Boolean)
On Error GoTo tratar_erro

With Lista
    Desenho = .ListItems.Item(InitFor).ListSubItems(3).Text
    DesenhoProduto = .ListItems.Item(InitFor).ListSubItems(5).Text
    
    IDCliente = 0
    Cliente = ""
    Permitido1 = False
    'Verifica qtde. fisica salva na tabela
    If IsNull(TBVendas!IDEstoque) = False And TBVendas!IDEstoque <> "" Then
        TextoFiltro = "IDestoque = " & TBVendas!IDEstoque
    Else
        If IsNull(TBVendas!LOTE) = True Or TBVendas!LOTE = "" Then
            TextoFiltro = "ID_empresa = " & TBVendas!ID_empresa & " and desenho = '" & Desenho & "'"
        Else
            TextoFiltro = "ID_empresa = " & TBVendas!ID_empresa & " and Data = '" & TBVendas!Data & "' and desenho = '" & Desenho & "' and lote = '" & TBVendas!LOTE & "' and local_armaz = '" & TBVendas!local_armaz & "' and Corrida = '" & TBVendas!Corrida & "' and Certificado = '" & TBVendas!Certificado & "'"
            If IsNull(TBVendas!Cod_ref) = False And TBVendas!Cod_ref <> "" Then TextoFiltro = TextoFiltro & " and Ref = '" & TBVendas!Cod_ref & "'"
            If IsNull(TBVendas!Numero_serie) = False And TBVendas!Numero_serie <> "" Then TextoFiltro = TextoFiltro & " and Numero_serie = '" & TBVendas!Numero_serie & "'"
        End If
    End If
    
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from Estoque_Controle where " & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        If TBVendas!Novo_lote = True Then
            Qtd = IIf(IsNull(TBVendas!qtde_fisica), 0, TBVendas!qtde_fisica)
            QtdPC = IIf(IsNull(TBVendas!qtde_fisica_PC), 0, TBVendas!qtde_fisica_PC)
        Else
            Qtd = IIf(IsNull(TBVendas!qtde_fisica), 0, TBVendas!qtde_fisica)
            QtdPC = IIf(IsNull(TBVendas!qtde_fisica_PC), 0, TBVendas!qtde_fisica_PC)
        End If
        IDCliente = IIf(IsNull(TBAbrir!ID_Cliente), 0, TBAbrir!ID_Cliente)
        Cliente = IIf(IsNull(TBAbrir!Cliente), "", TBAbrir!Cliente)
        Tipodest_NFcons = IIf(IsNull(TBAbrir!Tipodest_NFcons), "", TBAbrir!Tipodest_NFcons)
        If TBAbrir!Consignacao = True Then Permitido1 = True 'Consignado
    Else
        Qtd = IIf(IsNull(TBVendas!qtde_fisica), 0, TBVendas!qtde_fisica)
        QtdPC = IIf(IsNull(TBVendas!qtde_fisica_PC), 0, TBVendas!qtde_fisica_PC)
    End If
    Qtde = TBVendas!Qtde_estoque
    QtdePC = IIf(IsNull(TBVendas!Qtde_estoque_PC), 0, TBVendas!Qtde_estoque_PC)

    If Validar = True Then
    
        If IsNull(TBVendas!Novo_lote) = True Or TBVendas!Novo_lote = False Then
        
            If IsNull(TBVendas!LOTE) = True Or TBVendas!LOTE = "" Then
                TextoFiltro = "ID_empresa = " & TBVendas!ID_empresa & " and Desenho = '" & Desenho & "' and Estoque_real > 0 and Data <= '" & TBVendas!Data & "' and (Status <> 'ENTRADA_INVENTÁRIO' Or Status = 'ENTRADA_INVENTÁRIO' And Lote <> '" & "INV-" & Format(TBVendas!Data, "ddmmyy") & "')"
            ElseIf IsNull(TBVendas!IDEstoque) = False And TBVendas!IDEstoque <> "" Then
                    TextoFiltro = "IDestoque = " & TBVendas!IDEstoque
                Else
                    TextoFiltro = "ID_empresa = " & TBVendas!ID_empresa & " and Data = '" & TBVendas!Data & "' and Desenho = '" & Desenho & "' and Lote = '" & TBVendas!LOTE & "' and local_armaz = '" & TBVendas!local_armaz & "' and Corrida = '" & TBVendas!Corrida & "' and Certificado = '" & TBVendas!Certificado & "'"
            End If
            
            If IsNull(TBVendas!Cod_ref) = False And TBVendas!Cod_ref <> "" Then TextoFiltro = TextoFiltro & " and Ref = '" & TBVendas!Cod_ref & "'"
            If IsNull(TBVendas!Numero_serie) = False And TBVendas!Numero_serie <> "" Then TextoFiltro = TextoFiltro & " and Numero_serie = '" & TBVendas!Numero_serie & "'"
            
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select * from Estoque_Controle where " & TextoFiltro & " order by Data", Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                Do While TBAbrir.EOF = False
                
                    'Se existe empenho para outro lote, o sistema cria este empenho para o lote novo
                    If (IsNull(TBVendas!LOTE) = True Or TBVendas!LOTE = "") And TBVendas!Novo_lote = False Then
                        'Empenho de pedido interno
                        Set TBExecucao = CreateObject("adodb.recordset")
                        TBExecucao.Open "Select * from Estoque_Controle_Empenho_Vendas where id_estoque = " & TBAbrir!IDEstoque & " and (qtde_empenhada - ISNULL(qtde_saida, 0)) > 0", Conexao, adOpenKeyset, adLockOptimistic
                        If TBExecucao.EOF = False Then
                            Do While TBExecucao.EOF = False
                                Set TBCFOP = CreateObject("adodb.recordset")
                                TBCFOP.Open "Select * from Estoque_fisico_empenho", Conexao, adOpenKeyset, adLockOptimistic
                                TBCFOP.AddNew
                                TBCFOP!ID_fisico = TBVendas!ID
                                TBCFOP!ID_RE_antigo = TBAbrir!IDEstoque
                                TBCFOP!ID_carteira = TBExecucao!ID_carteira
                                TBCFOP!Data = TBExecucao!Data
                                TBCFOP!Responsavel = TBExecucao!Responsavel
                                TBCFOP!Qtde_empenho = TBExecucao!Qtde_empenhada - TBExecucao!Qtde_saida
                                TBCFOP!Ordem = False
                                TBCFOP.Update
                                TBCFOP.Close
                                TBExecucao.MoveNext
                            Loop
                        End If
                        TBExecucao.Close
                        
                        'Empenho de ordem
                        Set TBExecucao = CreateObject("adodb.recordset")
                        TBExecucao.Open "Select * from Producao_NF_Consignada where idestoque = " & TBAbrir!IDEstoque & " and (quantidade - ISNULL(qtde_saida, 0)) > 0", Conexao, adOpenKeyset, adLockOptimistic
                        If TBExecucao.EOF = False Then
                            Do While TBExecucao.EOF = False
                                Set TBCFOP = CreateObject("adodb.recordset")
                                TBCFOP.Open "Select * from Estoque_fisico_empenho", Conexao, adOpenKeyset, adLockOptimistic
                                TBCFOP.AddNew
                                TBCFOP!ID_fisico = TBVendas!ID
                                TBCFOP!ID_RE_antigo = TBAbrir!IDEstoque
                                TBCFOP!ID_carteira = TBExecucao!Ordem
                                TBCFOP!Data = TBExecucao!Data
                                TBCFOP!Responsavel = TBExecucao!Responsavel
                                TBCFOP!Qtde_empenho = TBExecucao!quantidade - TBExecucao!Qtde_saida
                                TBCFOP!Qtde_empenho_PC = IIf(IsNull(TBExecucao!Quantidade_PC), 0, TBExecucao!Quantidade_PC) - IIf(IsNull(TBExecucao!Qtde_saida_PC), 0, TBExecucao!Qtde_saida_PC)
                                TBCFOP!Ordem = True
                                TBCFOP.Update
                                TBCFOP.Close
                                TBExecucao.MoveNext
                            Loop
                        End If
                        TBExecucao.Close
                    End If
                If Novo_Estoque_Fisico = True Then
                    Set TBEstoque = CreateObject("adodb.recordset")
                    TBEstoque.Open "Select * from Estoque_movimentacao where IDestoque = " & TBAbrir!IDEstoque, Conexao, adOpenKeyset, adLockOptimistic
                    TBEstoque.AddNew
                    If IsNull(TBVendas!Destino) = True Or TBVendas!Destino = "" Then TBEstoque!Destino = "Interno" Else TBEstoque!Destino = TBVendas!Destino
                    TBEstoque!Terceiros = False
                    TBEstoque!LOTE = TBAbrir!LOTE
                    TBEstoque!Documento = TBAbrir!LOTE
                    TBEstoque!Desenho = TBAbrir!Desenho
                    Set TBItem = CreateObject("adodb.recordset")
                    TBItem.Open "Select classe from projproduto where desenho = '" & TBAbrir!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBItem.EOF = False Then
                        TBEstoque!Familia = TBItem!Classe
                    End If
                    TBItem.Close
                    
                    If IsNull(TBVendas!LOTE) = True Or TBVendas!LOTE = "" Then
                        qtdeliberada = TBAbrir!estoque_real
                        qtdeliberadaPC = IIf(IsNull(TBAbrir!estoque_real_PC), 0, TBAbrir!estoque_real_PC)
                        TBEstoque!Operacao = "SAIDA_INVENTÁRIO"
                        TBEstoque!Saida = qtdeliberada
                        TBEstoque!Saida_PC = qtdeliberadaPC
                        TBAbrir!estoque_real = TBAbrir!estoque_real - qtdeliberada
                        TBAbrir!estoque_real_PC = IIf(IsNull(TBAbrir!estoque_real_PC), 0, TBAbrir!estoque_real_PC) - qtdeliberadaPC
                    Else
                        If Txt_qtde_estoque > Qtd Then 'Saída
                            qtdeliberada = Txt_qtde_estoque - Qtd
                            qtdeliberadaPC = IIf(IsNull(TBAbrir!estoque_real_PC), 0, TBAbrir!estoque_real_PC) - QtdPC
                            TBEstoque!Operacao = "SAIDA_INVENTÁRIO"
                            TBEstoque!Saida = qtdeliberada
                            TBEstoque!Saida_PC = qtdeliberadaPC
                            
                            TBAbrir!estoque_real = Txt_qtde_estoque - qtdeliberada
                            TBAbrir!estoque_real_PC = IIf(IsNull(TBAbrir!estoque_real_PC), 0, TBAbrir!estoque_real_PC) - qtdeliberadaPC
                        Else
                            qtdeliberada = Qtd - Txt_qtde_estoque
                            qtdeliberadaPC = QtdPC - IIf(IsNull(TBAbrir!estoque_real_PC), 0, TBAbrir!estoque_real_PC)
                            TBEstoque!Operacao = "ENTRADA_INVENTÁRIO"
                            TBEstoque!Entrada = qtdeliberada
                            TBEstoque!Entrada_PC = qtdeliberadaPC
                            
                            TBAbrir!estoque_real = Txt_qtde_estoque + qtdeliberada
                            TBAbrir!estoque_real_PC = IIf(IsNull(TBAbrir!estoque_real_PC), 0, TBAbrir!estoque_real_PC) + qtdeliberadaPC
                        End If
                    End If
                    
                    TBEstoque!VlrUnit = Format(IIf(IsNull(TBAbrir!valor_unitario), 0, TBAbrir!valor_unitario), "###,##0.0000")
                    TBEstoque!vlrTotal = Format(TBEstoque!VlrUnit * qtdeliberada, "###,##0.00")
                    TBEstoque!Descricao = TBAbrir!Descricao
                    TBEstoque!Data = TBVendas!Data
                    TBEstoque!Responsavel = TBVendas!Responsavel
                    
                    TBEstoque!IDEstoque = TBAbrir!IDEstoque
                    TBEstoque!estoque_venda = 0
                    TBEstoque!ID_inventario = TBVendas!ID
                    TBEstoque.Update
                    TBEstoque.Close
                    
                    TBAbrir!estoque_venda = TBAbrir!estoque_real
                    TBAbrir!Valor_total = Format(IIf(IsNull(TBAbrir!valor_unitario), 0, TBAbrir!valor_unitario) * TBAbrir!estoque_real, "###,##0.00")
                    TBAbrir!Consignacao = TBVendas!Consignado
                    If IsNull(TBVendas!Tipo_cli_forn) = False And TBVendas!Tipo_cli_forn <> "" Then
                        TBAbrir!Tipodest_NFcons = TBVendas!Tipo_cli_forn
                        TBAbrir!ID_Cliente = TBVendas!ID_cli_forn
                        Set TBItem = CreateObject("adodb.recordset")
                        If TBVendas!Tipo_cli_forn = "C" Then
                            TBItem.Open "Select NomeRazao from Clientes where idcliente = " & TBVendas!ID_cli_forn & " and Prospecto = 'False' and DtValidacao IS NOT NULL and status <> 'Bloqueado'", Conexao, adOpenKeyset, adLockOptimistic
                            If TBItem.EOF = False Then
                                TBAbrir!Cliente = TBItem!NomeRazao
                            End If
                        Else
                            TBItem.Open "Select Nome_Razao from compras_fornecedores where idcliente = " & TBVendas!ID_cli_forn & " and Prospecto = 'False' and DtValidacao IS NOT NULL and status <> 'Bloqueado'", Conexao, adOpenKeyset, adLockOptimistic
                            If TBItem.EOF = False Then
                                TBAbrir!Cliente = TBItem!Nome_Razao
                            End If
                        End If
                        TBItem.Close
                    End If
                    If (TBVendas!Cod_ref) = False And TBVendas!Cod_ref <> "" Then TBAbrir!Ref = TBVendas!Cod_ref
                    TBAbrir!Numero_serie = TBVendas!Numero_serie
                    
                    TBAbrir.Update
                    End If
                    TBAbrir.MoveNext
                Loop
            End If
            TBAbrir.Close
            
        End If
            
        If TBVendas!Novo_lote = True Then
            ProcCriaMovimentacaoEntrada TBVendas!LOTE, TBVendas!local_armaz, TBVendas!Corrida, TBVendas!Certificado, .ListItems.Item(InitFor).Text, IIf(IsNull(TBVendas!Cod_ref), "", TBVendas!Cod_ref), IIf(IsNull(TBVendas!Numero_serie), "", TBVendas!Numero_serie)
        ElseIf IsNull(TBVendas!LOTE) = True Or TBVendas!LOTE = "" Then
        'Else
            ProcCriaMovimentacaoEntrada "INV-" & Format(TBVendas!Data, "ddmmyy"), TBVendas!local_armaz, TBVendas!Corrida, TBVendas!Certificado, .ListItems.Item(InitFor).Text, IIf(IsNull(TBVendas!Cod_ref), "", TBVendas!Cod_ref), IIf(IsNull(TBVendas!Numero_serie), "", TBVendas!Numero_serie)
        End If
        
    Else
        If IsNull(TBVendas!Novo_lote) = True Or TBVendas!Novo_lote = False Then
            Set TBEstoque = CreateObject("adodb.recordset")
            TBEstoque.Open "Select IDestoque, Operacao, Entrada, Entrada_PC, Saida, Saida_PC from Estoque_movimentacao where ID_inventario = " & TBVendas!ID & " and (operacao = 'SAIDA_INVENTÁRIO' or operacao = 'ENTRADA_INVENTÁRIO')", Conexao, adOpenKeyset, adLockOptimistic
            If TBEstoque.EOF = False Then
                Do While TBEstoque.EOF = False
                    Set TBAbrir = CreateObject("adodb.recordset")
                    TBAbrir.Open "Select Estoque_real, Estoque_real_PC, estoque_venda, qtde_fisica, valor_unitario, Valor_Total from Estoque_Controle where idestoque = " & TBEstoque!IDEstoque & " and (etiqueta <> '" & TBVendas!Etiqueta & "' or etiqueta IS NULL) order by Data", Conexao, adOpenKeyset, adLockOptimistic
                    If TBAbrir.EOF = False Then
                        If TBEstoque!Operacao = "SAIDA_INVENTÁRIO" Then 'Saída
                            TBAbrir!estoque_real = TBAbrir!estoque_real + TBEstoque!Saida
                            TBAbrir!estoque_real_PC = IIf(IsNull(TBAbrir!estoque_real_PC), 0, TBAbrir!estoque_real_PC) + IIf(IsNull(TBEstoque!Saida_PC), 0, TBEstoque!Saida_PC)
                        Else
                            qtdeliberada = Qtd - TBAbrir!estoque_real
                            qtdeliberadaPC = QtdPC - IIf(IsNull(TBAbrir!estoque_real_PC), 0, TBAbrir!estoque_real_PC)
                            TBAbrir!estoque_real = TBAbrir!estoque_real - TBEstoque!Entrada
                            TBAbrir!estoque_real_PC = IIf(IsNull(TBAbrir!estoque_real_PC), 0, TBAbrir!estoque_real_PC) - IIf(IsNull(TBEstoque!Entrada_PC), 0, TBEstoque!Entrada_PC)
                        End If
                        TBAbrir!estoque_venda = TBAbrir!estoque_real
                        TBAbrir!Valor_total = Format(IIf(IsNull(TBAbrir!valor_unitario), 0, TBAbrir!valor_unitario) * TBAbrir!estoque_real, "###,##0.00")
                        TBAbrir!qtde_fisica = 0
                                                        
                        TBAbrir.Update
                        TBEstoque.Delete
                    End If
                    TBAbrir.Close
                    TBEstoque.MoveNext
                Loop
            End If
            TBEstoque.Close
        End If

        If TBVendas!Novo_lote = True Or (IsNull(TBVendas!LOTE) = True Or TBVendas!LOTE = "") Then
   '     ProcExcluirMovimentacaoEntrada .ListItems.Item(InitFor).Text
        End If
        
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluirMovimentacaoEntrada(ID_documento As Long)
On Error GoTo tratar_erro

If TBVendas!Novo_lote = True Then TextoFiltro = " and Data = '" & TBVendas!Data & "'" Else TextoFiltro = ""
If IsNull(TBVendas!LOTE) = False And TBVendas!LOTE <> "" Then
    If TextoFiltro <> "" Then TextoFiltro = TextoFiltro & " and Lote = '" & TBVendas!LOTE & "'" Else TextoFiltro = " and Lote = '" & TBVendas!LOTE & "'"
End If
If IsNull(TBVendas!Etiqueta) = False And TBVendas!Etiqueta <> "" Then
    If TextoFiltro <> "" Then TextoFiltro = TextoFiltro & " and Etiqueta = '" & TBVendas!Etiqueta & "'" Else TextoFiltro = " and Etiqueta = '" & TBVendas!Etiqueta & "'"
End If
If IsNull(TBVendas!Cod_ref) = False And TBVendas!Cod_ref <> "" Then
    If TextoFiltro <> "" Then TextoFiltro = TextoFiltro & " and Ref = '" & TBVendas!Cod_ref & "'" Else TextoFiltro = " and Ref = '" & TBVendas!Cod_ref & "'"
End If
If IsNull(TBVendas!Numero_serie) = False And TBVendas!Numero_serie <> "" Then
    If TextoFiltro <> "" Then TextoFiltro = TextoFiltro & " and Numero_serie = '" & TBVendas!Numero_serie & "'" Else TextoFiltro = " and Numero_serie = '" & TBVendas!Numero_serie & "'"
End If
Set TBEstoque = CreateObject("adodb.recordset")
TBEstoque.Open "Select * from estoque_controle where ID_empresa = " & TBVendas!ID_empresa & " and Desenho = '" & Desenho & "' and local_armaz = '" & TBVendas!local_armaz & "' and Corrida = '" & TBVendas!Corrida & "' and Certificado = '" & TBVendas!Certificado & "'" & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
If TBEstoque.EOF = False Then

    'Verifica se existe empenho no inventario, se tiver o empenho volta para o lote antigo
    If (IsNull(TBVendas!LOTE) = True Or TBVendas!LOTE = "") And TBVendas!Novo_lote = False Then
        Set TBCFOP = CreateObject("adodb.recordset")
        TBCFOP.Open "Select * from Estoque_fisico_empenho where ID_fisico = " & TBVendas!ID, Conexao, adOpenKeyset, adLockOptimistic
        If TBCFOP.EOF = False Then
            Do While TBCFOP.EOF = False
                If TBCFOP!Ordem = False Then
                    'Empenho por pedido interno
                    Set TBExecucao = CreateObject("adodb.recordset")
                    TBExecucao.Open "Select * from Estoque_Controle_Empenho_Vendas where id_estoque = " & TBCFOP!ID_RE_antigo & " and ID_carteira = " & TBCFOP!ID_carteira, Conexao, adOpenKeyset, adLockOptimistic
                    If TBExecucao.EOF = True Then
                        TBExecucao.AddNew
                        TBExecucao!ID_estoque = TBCFOP!ID_RE_antigo
                        TBExecucao!ID_carteira = TBCFOP!ID_carteira
                        TBExecucao!Data = TBCFOP!Data
                        TBExecucao!Responsavel = TBCFOP!Responsavel
                    End If
                    TBExecucao!Qtde_empenhada = IIf(IsNull(TBExecucao!Qtde_empenhada), 0, TBExecucao!Qtde_empenhada) + TBCFOP!Qtde_empenho
                    TBExecucao.Update
                    TBExecucao.Close
                Else
                    'Empenho por Ordem
                    Set TBExecucao = CreateObject("adodb.recordset")
                    TBExecucao.Open "Select * from Producao_NF_Consignada where idestoque = " & TBCFOP!ID_RE_antigo & " and Ordem = " & TBCFOP!ID_carteira, Conexao, adOpenKeyset, adLockOptimistic
                    If TBExecucao.EOF = True Then
                        TBExecucao.AddNew
                        TBExecucao!Codinterno = TBEstoque!Desenho
                        TBExecucao!IDEstoque = TBCFOP!ID_RE_antigo
                        TBExecucao!Ordem = TBCFOP!ID_carteira
                        TBExecucao!Data = TBCFOP!Data
                        TBExecucao!Responsavel = TBCFOP!Responsavel
                    End If
                    TBExecucao!quantidade = IIf(IsNull(TBExecucao!quantidade), 0, TBExecucao!quantidade) + TBCFOP!Qtde_empenho
                    TBExecucao!Quantidade_PC = IIf(IsNull(TBExecucao!Quantidade_PC), 0, TBExecucao!Quantidade_PC) + IIf(IsNull(TBCFOP!Qtde_empenho_PC), 0, TBCFOP!Qtde_empenho_PC)
                    TBExecucao.Update
                    TBExecucao.Close
                End If
                TBCFOP.Delete
                TBCFOP.MoveNext
            Loop
        End If
        TBCFOP.Close
    End If

    '==================================
    Modulo = "Estoque/Inventário"
    Evento = "Cancelar validação estoque físico"
    ID_documento = ID_documento
    Documento = "Data: " & Format(TBVendas!Data, "dd/mm/yy") & " - Código interno: " & Desenho & " - Lote: " & TBEstoque!LOTE & " - Corrida: " & TBEstoque!Corrida & " - Certificado " & TBEstoque!Certificado
    Documento1 = ""
    ProcGravaEvento
    '==================================
    
    Conexao.Execute "DELETE from Estoque_movimentacao where ID_inventario = " & ID_documento
    
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from Estoque_movimentacao where IDestoque = " & TBEstoque!IDEstoque, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = True Then
        Conexao.Execute "DELETE Estoque_Controle_Empenho_Vendas where ID_estoque = " & TBEstoque!IDEstoque
        Conexao.Execute "DELETE Producao_NF_Consignada where IDestoque = " & TBEstoque!IDEstoque
        Conexao.Execute "DELETE from Estoque_Controle where IDestoque = " & TBEstoque!IDEstoque
    End If
    TBAbrir.Close
End If
TBEstoque.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCriaMovimentacaoEntrada(LOTE As String, LA As String, Corrida As String, Certificado As String, ID_documento As Long, CodRef As String, Nserie As String)
On Error GoTo tratar_erro

ValorTotal = TBVendas!valor_unitario
quantestoque = TBVendas!qtde_fisica
quantestoquePC = IIf(IsNull(TBVendas!qtde_fisica_PC), 0, TBVendas!qtde_fisica_PC)

TextoFiltro = ""
If CodRef <> "" Then TextoFiltro = " and Ref = '" & CodRef & "'"
If Nserie <> "" Then
    If TextoFiltro <> "" Then TextoFiltro = TextoFiltro & " and Numero_serie = '" & Nserie & "'" Else TextoFiltro = " and Numero_serie = '" & Nserie & "'"
End If
Set TBEstoque = CreateObject("adodb.recordset")
TBEstoque.Open "Select * from estoque_controle where ID_empresa = " & TBVendas!ID_empresa & " and Desenho = '" & Desenho & "' and Etiqueta = '" & TBVendas!Etiqueta & "' and Lote = '" & LOTE & "' and local_armaz = '" & LA & "' and Corrida = '" & Corrida & "' and Certificado = '" & Certificado & "'" & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
If TBEstoque.EOF = True Then TBEstoque.AddNew
TBEstoque!ID_empresa = TBVendas!ID_empresa
TBEstoque!Etiqueta = TBVendas!Etiqueta
TBEstoque!LOTE = LOTE
TBEstoque!Desenho = Desenho


Set TBItem = CreateObject("adodb.recordset")
TBItem.Open "Select classe, Unidade from projproduto where desenho = '" & Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBItem.EOF = False Then
    Familiatext = TBItem!Classe
    TBEstoque!Un = TBItem!Unidade
End If
TBItem.Close

TBEstoque!Classe = Familiatext
TBEstoque!valor_unitario = Format(ValorTotal, "###,##0.0000")
TBEstoque!Valor_total = Format(quantestoque * ValorTotal, "###,##0.00")
TBEstoque!Descricao = DesenhoProduto
TBEstoque!Data = TBVendas!Data
TBEstoque!Responsavel = TBVendas!Responsavel
TBEstoque!Certificado = Certificado
TBEstoque!Corrida = Certificado
TBEstoque!local_armaz = LA
TBEstoque!Qtde = TBVendas!qtde_fisica
TBEstoque!status = "ENTRADA_INVENTÁRIO"

If IDCliente <> 0 Then
    'Grava quando for correção de um RE
    TBEstoque!ID_Cliente = IDCliente
    TBEstoque!Cliente = Cliente
    TBEstoque!Tipodest_NFcons = IIf(Tipodest_NFcons = "", Null, Tipodest_NFcons)
    If Permitido1 = True Then TBEstoque!Consignacao = True Else TBEstoque!Consignacao = False
Else
    TBEstoque!Consignacao = TBVendas!Consignado
    If IsNull(TBVendas!Tipo_cli_forn) = False And TBVendas!Tipo_cli_forn <> "" Then
        TBEstoque!Tipodest_NFcons = TBVendas!Tipo_cli_forn
        TBEstoque!ID_Cliente = TBVendas!ID_cli_forn
        Set TBItem = CreateObject("adodb.recordset")
        If TBVendas!Tipo_cli_forn = "C" Then
            TBItem.Open "Select NomeRazao from Clientes where idcliente = " & TBVendas!ID_cli_forn & " and Prospecto = 'False' and DtValidacao IS NOT NULL and status <> 'Bloqueado'", Conexao, adOpenKeyset, adLockOptimistic
            If TBItem.EOF = False Then
                TBEstoque!Cliente = TBItem!NomeRazao
            End If
        Else
            TBItem.Open "Select Nome_Razao from compras_fornecedores where idcliente = " & TBVendas!ID_cli_forn & " and Prospecto = 'False' and DtValidacao IS NOT NULL and status <> 'Bloqueado'", Conexao, adOpenKeyset, adLockOptimistic
            If TBItem.EOF = False Then
                TBEstoque!Cliente = TBItem!Nome_Razao
            End If
        End If
        TBItem.Close
    End If
End If

TBEstoque!estoque_real = Qtd
TBEstoque!estoque_real_PC = QtdPC
TBEstoque!estoque_venda = Qtd
    
If IsNull(TBVendas!Cod_ref) = False And TBVendas!Cod_ref <> "" Then TBEstoque!Ref = TBVendas!Cod_ref
TBEstoque!Numero_serie = TBVendas!Numero_serie

TBEstoque.Update

'Corrige empenho
If (IsNull(TBVendas!LOTE) = True Or TBVendas!LOTE = "") And TBVendas!Novo_lote = False Then
    Conexao.Execute "UPDATE Estoque_fisico_empenho Set ID_RE = " & TBEstoque!IDEstoque & " where ID_fisico = " & TBVendas!ID
        
    Set TBCFOP = CreateObject("adodb.recordset")
    TBCFOP.Open "Select * from Estoque_fisico_empenho where ID_fisico = " & TBVendas!ID, Conexao, adOpenKeyset, adLockOptimistic
    If TBCFOP.EOF = False Then
        Do While TBCFOP.EOF = False
            If TBCFOP!Ordem = False Then
                'Empenho por pedido interno
                Set TBExecucao = CreateObject("adodb.recordset")
                TBExecucao.Open "Select * from Estoque_Controle_Empenho_Vendas where id_estoque = " & TBCFOP!ID_RE_antigo & " and id_carteira = " & TBCFOP!ID_carteira & " and Qtde_empenhada - ISNULL(Qtde_saida, 0) > 0", Conexao, adOpenKeyset, adLockOptimistic
                If TBExecucao.EOF = False Then
                    If quantestoque > 0 Then
                        Set TBCST = CreateObject("adodb.recordset")
                        TBCST.Open "Select * from Estoque_Controle_Empenho_Vendas", Conexao, adOpenKeyset, adLockOptimistic
                        TBCST.AddNew
                        TBCST!ID_estoque = TBEstoque!IDEstoque
                        TBCST!ID_carteira = TBExecucao!ID_carteira
                        If (TBExecucao!Qtde_empenhada - TBExecucao!Qtde_saida) > quantestoque Then
                            TBCST!Qtde_empenhada = quantestoque
                        Else
                            TBCST!Qtde_empenhada = TBExecucao!Qtde_empenhada - TBExecucao!Qtde_saida
                        End If
                        TBCST!Qtde_saida = 0
                        TBCST!Data = TBExecucao!Data
                        TBCST!Responsavel = TBExecucao!Responsavel
                        TBCST.Update
                        TBCST.Close
                        quantestoque = quantestoque - (TBExecucao!Qtde_empenhada - TBExecucao!Qtde_saida)
                    End If
            
                    If TBExecucao!Qtde_saida = 0 Then
                        TBExecucao.Delete
                    Else
                        TBExecucao!Qtde_empenhada = TBExecucao!Qtde_saida
                        TBExecucao.Update
                    End If
                End If
            Else
                'Empenho por Ordem
                Set TBExecucao = CreateObject("adodb.recordset")
                TBExecucao.Open "Select * from Producao_NF_Consignada where idestoque = " & TBCFOP!ID_RE_antigo & " and ordem = " & TBCFOP!ID_carteira & " and quantidade - ISNULL(Qtde_saida, 0) > 0", Conexao, adOpenKeyset, adLockOptimistic
                If TBExecucao.EOF = False Then
                    If quantestoque > 0 Then
                        Set TBCST = CreateObject("adodb.recordset")
                        TBCST.Open "Select * from Producao_NF_Consignada", Conexao, adOpenKeyset, adLockOptimistic
                        TBCST.AddNew
                        TBCST!IDEstoque = TBEstoque!IDEstoque
                        TBCST!Codinterno = TBEstoque!Desenho
                        TBCST!Ordem = TBExecucao!Ordem
                        If (TBExecucao!quantidade - TBExecucao!Qtde_saida) > quantestoque Then
                            TBCST!quantidade = quantestoque
                        Else
                            TBCST!quantidade = TBExecucao!quantidade - TBExecucao!Qtde_saida
                        End If
                        If (IIf(IsNull(TBExecucao!Quantidade_PC), 0, TBExecucao!Quantidade_PC) - IIf(IsNull(TBExecucao!Qtde_saida_PC), 0, TBExecucao!Qtde_saida_PC)) > quantestoquePC Then
                            TBCST!Quantidade_PC = quantestoquePC
                        Else
                            TBCST!Quantidade_PC = IIf(IsNull(TBExecucao!Quantidade_PC), 0, TBExecucao!Quantidade_PC) - IIf(IsNull(TBExecucao!Qtde_saida_PC), 0, TBExecucao!Qtde_saida_PC)
                        End If
                        TBCST!Qtde_saida = 0
                        TBCST!Qtde_saida_PC = 0
                        TBCST!Data = TBExecucao!Data
                        TBCST!Responsavel = TBExecucao!Responsavel
                        TBCST.Update
                        TBCST.Close
                        quantestoque = quantestoque - (TBExecucao!quantidade - TBExecucao!Qtde_saida)
                        quantestoquePC = quantestoquePC - (IIf(IsNull(TBExecucao!Quantidade_PC), 0, TBExecucao!Quantidade_PC) - IIf(IsNull(TBExecucao!Qtde_saida_PC), 0, TBExecucao!Qtde_saida_PC))
                    End If
            
                    If TBExecucao!Qtde_saida = 0 Then
                        TBExecucao.Delete
                    Else
                        TBExecucao!quantidade = TBExecucao!Qtde_saida
                        TBExecucao!Quantidade_PC = IIf(IsNull(TBExecucao!Qtde_saida_PC), 0, TBExecucao!Qtde_saida_PC)
                        TBExecucao.Update
                    End If
                End If
            End If
            TBCFOP.MoveNext
        Loop
    End If
    TBCFOP.Close
End If

Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select * from Estoque_movimentacao", Conexao, adOpenKeyset, adLockOptimistic
TBProduto.AddNew

If IsNull(TBVendas!Destino) = True Or TBVendas!Destino = "" Then
TBProduto!Destino = "Interno"
Else
TBProduto!Destino = TBVendas!Destino
End If

TBProduto!Terceiros = False
TBProduto!LOTE = LOTE
TBProduto!Documento = LOTE
TBProduto!Desenho = Desenho
TBProduto!Familia = Familiatext
quantestoque = TBVendas!qtde_fisica
TBProduto!VlrUnit = Format(ValorTotal, "###,##0.0000")
TBProduto!vlrTotal = Format(quantestoque * ValorTotal, "###,##0.00")
TBProduto!Descricao = DesenhoProduto
TBProduto!Data = TBVendas!Data
TBProduto!Responsavel = TBVendas!Responsavel
TBProduto!Entrada = TBVendas!qtde_fisica
TBProduto!Entrada_PC = IIf(IsNull(TBVendas!qtde_fisica_PC), 0, TBVendas!qtde_fisica_PC)
TBProduto!Operacao = "ENTRADA_INVENTÁRIO"
TBProduto!estoque_venda = Qtd
TBProduto!IDEstoque = TBEstoque!IDEstoque
TBProduto!ID_inventario = TBVendas!ID
TBProduto.Update
TBProduto.Close
'==================================
Modulo = "Estoque/Inventário"
Evento = "Validar estoque físico"
ID_documento = ID_documento
Documento = "Data: " & Format(TBVendas!Data, "dd/mm/yy") & " - Código interno: " & Desenho & " - Lote: " & LOTE & " - Corrida: " & Corrida & " - Certificado " & Certificado
Documento1 = ""
ProcGravaEvento
'==================================

TBEstoque.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub btnSalvarValorUnitario_Click()
On Error GoTo tratar_erro

If USMsgBox("Deseja realmente atualizar o valor unitario nessa RE e em todas movimentações dessa RE?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    StrSql = "Update Estoque_fisico Set Valor_unitario = '" & Replace(Txt_valor_unitario, ",", ".") & "' where idEstoque = '" & txt_RE & "'"
    'Debug.print StrSql
    
    Conexao.Execute (StrSql)
    StrSql = "Update Estoque_Movimentacao Set Vlrunit = '" & Replace(Txt_valor_unitario, ",", ".") & "' where idEstoque = '" & txt_RE & "'"
    'Debug.print StrSql
    
    Conexao.Execute (StrSql)
    ProcAtualizalista (1)
    USMsgBox "Valor atualizado com sucesso!", vbInformation, "CAPRIND v5.0"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub btnValor_Click()
On Error GoTo tratar_erro

If Txt_cod_interno <> "" Then
If USMsgBox("Deseja realmente buscar o valor unitário do cadastro do item a ser inventariado?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    Set TBProduto = CreateObject("adodb.recordset")
    StrSql = "Select * from projproduto where Desenho = '" & Txt_cod_interno & "'"
    TBProduto.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
     If TBProduto!Vendas = True Then
      If TBProduto!PConsumo = 0 Then
      USMsgBox "Cadastro do produto sem valor de venda", vbCritical, "CAPRIND v5.0"
      Exit Sub
      End If
      Txt_valor_unitario.Text = Format(TBProduto!PConsumo, "###,##0.0000")
     Else
      If TBProduto!PCusto = 0 Then
      USMsgBox "Cadastro do produto sem valor de compra", vbCritical, "CAPRIND v5.0"
      Exit Sub
      End If
      
      Txt_valor_unitario.Text = Format(TBProduto!PCusto, "###,##0.0000")
     End If
    End If
    TBProduto.Close
End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_consignado_Click()
On Error GoTo tratar_erro

If Chk_consignado.Value = 1 Then
    Label2(5).Caption = "Tipo*"
    Label2(6).Caption = "ID*"
    Label2(7).Caption = "Cliente/Fornecedor*"
Else
    Label2(5).Caption = "Tipo"
    Label2(6).Caption = "ID"
    Label2(7).Caption = "Cliente/Fornecedor"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub Cmb_certificado_Click()
On Error GoTo tratar_erro

ProcVerifDadosEstoque

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub Cmb_corrida_Click()
On Error GoTo tratar_erro

With Cmb_certificado
    .Clear
    If cmb_Lote <> "" And txt_LA <> "" And Cmb_corrida <> "" Then
        If Cmb_RE <> "" Then TextoFiltro = "IDestoque = " & Cmb_RE Else TextoFiltro = "ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and desenho = '" & Txt_cod_interno & "' and lote = '" & cmb_Lote & "' and local_armaz = '" & txt_LA & "' and Corrida = '" & Cmb_corrida & "'"
        Set TBAliquota = CreateObject("adodb.recordset")
        TBAliquota.Open "Select Certificado from estoque_controle where " & TextoFiltro & " group by Certificado", Conexao, adOpenKeyset, adLockOptimistic
        If TBAliquota.EOF = False Then
            Do While TBAliquota.EOF = False
                .AddItem TBAliquota!Certificado
                If Cmb_RE <> "" Then .Text = TBAliquota!Certificado
                TBAliquota.MoveNext
            Loop
        End If
        TBAliquota.Close
    End If
End With
ProcVerifDadosEstoque

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_empresa_Click()
On Error GoTo tratar_erro

If FunVerifMovimentacaoEstPC(Cmb_empresa.ItemData(Cmb_empresa.ListIndex)) = True Then
    With Txt_qtde_fisica
        .Locked = True
        .TabStop = False
    End With
Else
    With Txt_qtde_fisica
        .Locked = False
        .TabStop = True
    End With
End If
    
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
        .ButtonState(6) = 5
    Else
        .ButtonState(4) = 5
        .ButtonState(6) = 0
    End If
    .Refresh
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub Cmd_localizar_cli_forn_Click()
On Error GoTo tratar_erro

ProcLocalizarFornecedor

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLocalizarFornecedor()
On Error GoTo tratar_erro

ProcConfVariaveisLocCliente False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, True
ProcConfVariaveisLocForn False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, True
If Cmb_tipo_cli_forn = "Cliente" Then frmVendas_LocalizarCliente.Show 1 Else FrmCompras_localizafornecedor.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub



Private Sub ProcVerifDadosEstoque()
On Error GoTo tratar_erro

If Txt_cod_interno = "" Then Exit Sub
With Txt_valor_unitario
    .Text = ""
    Txt_qtde_estoque = "0,0000"
    Txt_qtde_estoque_PC = "0,0000"

    Set TBEstoque = CreateObject("adodb.recordset")
    If Chk_novo_lote.Value = 1 Then
        TBEstoque.Open "Select * from projproduto where desenho = '" & Txt_cod_interno & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBEstoque.EOF = False Then
            .Text = IIf(IsNull(TBEstoque!PCusto), "0,00000", Format(TBEstoque!PCusto, "###,##0.0000"))
        End If
        TBEstoque.Close
    ElseIf cmb_Lote = "" Then
            'Qtde. total em estoque
            Txt_qtde_estoque = Format(FunVerificaQtdeEstoque(Txt_cod_interno, 0, ""), "###,##0.0000")
            Txt_qtde_estoque_PC = Format(qt, "###,##0.0000")
            Estoquereal = Txt_qtde_estoque
            If Estoquereal > 0 And Valor_total > 0 Then Txt_valor_unitario = Format(Valor_total / Estoquereal, "###,##0.0000") Else Txt_valor_unitario = "0,00000"
        ElseIf Txt_cod_interno <> "" And txt_LA <> "" And Cmb_corrida <> "" And Cmb_certificado <> "" Then
                If Cmb_RE <> "" Then TextoFiltro = "IDestoque = " & Cmb_RE Else TextoFiltro = "ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and desenho = '" & Txt_cod_interno & "' and lote = '" & cmb_Lote & "' and local_armaz = '" & txt_LA & "' and Corrida = '" & Cmb_corrida & "' and Certificado = '" & Cmb_certificado & "'"
                TBEstoque.Open "Select valor_unitario, estoque_real, estoque_real_PC from estoque_controle where " & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
                If TBEstoque.EOF = False Then
                    Txt_valor_unitario = IIf(IsNull(TBEstoque!valor_unitario), "0,00000", Format(TBEstoque!valor_unitario, "###,##0.0000"))
                    Txt_qtde_estoque = IIf(IsNull(TBEstoque!estoque_real), "0,0000", Format(TBEstoque!estoque_real, "###,##0.0000"))
                    Txt_qtde_estoque_PC = IIf(IsNull(TBEstoque!estoque_real_PC), "0,0000", Format(TBEstoque!estoque_real_PC, "###,##0.0000"))
                End If
                TBEstoque.Close
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Estoque_Fisico.AbsolutePage <> 2 Then
    If TBLISTA_Estoque_Fisico.AbsolutePage = -3 Then
        ProcExibePagina (TBLISTA_Estoque_Fisico.PageCount - 1)
    Else
        TBLISTA_Estoque_Fisico.AbsolutePage = TBLISTA_Estoque_Fisico.AbsolutePage - 2
        ProcExibePagina (TBLISTA_Estoque_Fisico.AbsolutePage)
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
    TBLISTA_Estoque_Fisico.AbsolutePage = txtPagIr.Text
    ProcExibePagina (TBLISTA_Estoque_Fisico.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Estoque_Fisico.AbsolutePage = 1
ProcExibePagina (TBLISTA_Estoque_Fisico.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Estoque_Fisico.AbsolutePage <> -3 Then
    If TBLISTA_Estoque_Fisico.AbsolutePage = 1 Then
        ProcExibePagina (2)
    Else
        ProcExibePagina (TBLISTA_Estoque_Fisico.AbsolutePage)
    End If
Else
    ProcExibePagina (TBLISTA_Estoque_Fisico.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Estoque_Fisico.AbsolutePage = TBLISTA_Estoque_Fisico.PageCount
ProcExibePagina (TBLISTA_Estoque_Fisico.AbsolutePage)

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
    Case vbKeyF4: If Cmb_opcao_lista = "Excluir" Then ProcExcluir
    Case vbKeyF5: ProcImprimir
    Case vbKeyF7: If Cmb_opcao_lista = "Validação" Then ProcValidarRegistros Lista, "Estoque/Inventário"
    'Case vbKeyF1: ProcAjuda
    Case vbKeyEscape: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15192, 10, True
Formulario = "Estoque/Inventário"
Direitos
ProcLimpaVariaveisPrincipais

ProcCarregaComboEmpresa Cmb_empresa, False
Cmb_empresa_Click

cmbfiltrarpor = "Código interno"
Txt_data = Date
Cmb_opcao_lista = "Validação"

ProcRemoveObjetosResize Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

Formulario = "Estoque/Inventário"
Direitos
ProcLimpaVariaveisPrincipais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro

frmEstoque_fisico_MenuImpressao.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

If Novo_Estoque_Fisico = True Then
    If USMsgBox("O estoque físico ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvar
        If Novo_Estoque_Fisico = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
Novo_Estoque_Fisico = False
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvar()
On Error GoTo tratar_erro
  
If Frame1.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
Acao = "salvar"
If FunVerifValidacaoRegistro("alterar", txtDtValidacao, "mesmo", "o inventário", True) = False Then Exit Sub

If Txt_ID_prod = 0 Then
    NomeCampo = "o código interno"
    ProcVerificaAcao
    Txt_cod_interno.SetFocus
    Exit Sub
End If

'If Txt_etiqueta = "" Then
'    NomeCampo = "o número da etiqueta"
'    ProcVerificaAcao
'    Txt_etiqueta.SetFocus
'    Exit Sub
'End If
If txt_RE.Text = "" Then
    If Txt_lote = "" Then
        NomeCampo = "o número do lote"
        ProcVerificaAcao
        Txt_lote.SetFocus
        Exit Sub
    End If
    If txt_LA = "" Then
        NomeCampo = "o local de armazenamento"
        ProcVerificaAcao
        txt_LA.SetFocus
        Exit Sub
    End If
    If txt_Corrida = "" Then txt_Corrida = 0
    If txt_Certificado = "" Then txt_Certificado = 0
ElseIf txt_RE <> "" Or Cmb_RE = "" And cmb_Lote <> "" Then
        If Cmb_RE <> "" And cmb_Lote = "" Then
            NomeCampo = "o número do lote"
            ProcVerificaAcao
            cmb_Lote.SetFocus
            Exit Sub
        End If
        If txt_LA = "" Then
            NomeCampo = "o local de armazenamento"
            ProcVerificaAcao
            txt_LA.SetFocus
            Exit Sub
        End If
        If txt_Corrida = "" Then
            NomeCampo = "o número da corrida"
            ProcVerificaAcao
            txt_Corrida.SetFocus
            Exit Sub
        End If
        If txt_Certificado = "" Then
            NomeCampo = "o número do certificado"
            ProcVerificaAcao
            txt_Certificado.SetFocus
            Exit Sub
        End If
    Else
        If txt_LA = "" Then
            NomeCampo = "o local de armazenamento"
            ProcVerificaAcao
            txt_LA.SetFocus
            Exit Sub
        End If
        If txt_Corrida = "" Then txt_Corrida = 0
        If txt_Certificado = "" Then txt_Certificado = 0
End If
If cmbDestino = "" Then
    NomeCampo = "o destino"
    ProcVerificaAcao
    cmbDestino.SetFocus
    Exit Sub
End If

valor = IIf(Txt_valor_unitario = "", 0, Txt_valor_unitario)
If valor <= 0 And Txt_valor_unitario.Locked = False Then
    NomeCampo = "o valor unitário"
    ProcVerificaAcao
    Txt_valor_unitario.SetFocus
    Exit Sub
End If
Qtde = IIf(Txt_qtde_fisica = "", 0, Txt_qtde_fisica)
If Txt_qtde_fisica = "" Or Qtde < 0 Then
    NomeCampo = "a quantidade física"
    ProcVerificaAcao
    Txt_qtde_fisica.SetFocus
    Exit Sub
End If

If Chk_consignado.Value = 1 Then
    If Cmb_tipo_cli_forn = "" Then
        NomeCampo = "o tipo"
        ProcVerificaAcao
        Cmb_tipo_cli_forn.SetFocus
        Exit Sub
    End If
    If Txt_cli_forn = "" Then
        NomeCampo = "o cliente/fornecedor"
        ProcVerificaAcao
        Cmd_localizar_cli_forn_Click
        Exit Sub
    End If
End If

'Verifica se o número da etiqueta já foi utilizado
If Txt_etiqueta <> "" Then
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from Estoque_fisico where Etiqueta = '" & Txt_etiqueta & "' and ID <> " & Txt_ID, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        USMsgBox ("Este número de etiqueta já foi utilizado, favor alterar."), vbExclamation, "CAPRIND v5.0"
        Txt_etiqueta.SetFocus
        Exit Sub
    End If
    TBAbrir.Close
End If
'Verifica se o código de referencia está vinculado a outro produto
'If Cmb_cod_ref <> "" Then If FunVerifiCodRefUtilizado(Txt_cod_interno, Cmb_cod_ref) = True Then Exit Sub

If txt_RE.Text = "" Then
    'Verifica se já foi criado um inventario para este item na mesma data
    TextoFiltro = ""
    MsgTexto = ""
    If txt_cod_Referencia <> "" Then
        TextoFiltro = " and Cod_ref = '" & txt_cod_Referencia & "'"
        MsgTexto = " - Cód. de referência " & txt_cod_Referencia
    End If
    If Txt_numero_serie <> "" Then
        If TextoFiltro <> "" Then
            TextoFiltro = TextoFiltro & " and Numero_serie = '" & Txt_numero_serie & "'"
            MsgTexto = MsgTexto & " - Número de série " & Txt_numero_serie
        Else
            TextoFiltro = " and Numero_serie = '" & Txt_numero_serie & "'"
            MsgTexto = " - Número de série " & Txt_numero_serie
        End If
    End If
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select ID from Estoque_fisico where Data = '" & Format(Txt_data, "Short Date") & "' and ID <> " & Txt_ID & " and codproduto = " & Txt_ID_prod & " and Lote = '" & Txt_lote & "'" & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        USMsgBox ("Já existe um inventário nesta data para o produto " & Txt_cod_interno & " - Lote " & Txt_lote & MsgTexto & "."), vbExclamation, "CAPRIND v5.0"
        TBAbrir.Close
        Exit Sub
    End If
    TBAbrir.Close
    
    If txt_cod_Referencia <> "" Then
        'If FunVerifiCodRefUtilizado(Txt_cod_interno, Cmb_cod_ref) = True Then Exit Sub 'Verifica se o código de referencia está vinculado a outro produto
    
        'Verifica se já existe RE com o código de referencia no estoque
        Set TBCompras = CreateObject("adodb.recordset")
        TBCompras.Open "Select IDestoque from Estoque_controle where Desenho = '" & Txt_cod_interno & "' and Ref = '" & Cmb_cod_ref & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBCompras.EOF = False Then
            If USMsgBox("Já existe movimentação no estoque com esse código de referência, deseja prosseguir assim mesmo?", vbYesNo, "CAPRIND v5.0") = vbNo Then
                Cmb_cod_ref.SetFocus
                TBCompras.Close
                Exit Sub
            End If
        End If
        TBCompras.Close
    End If
End If

'Verifica se a quantidade fisica é menor que o total empenhado
If txt_RE.Text = "" Then
    If Txt_lote = "" Then
        TextoFiltro = "EC.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and EC.desenho = '" & Txt_cod_interno & "'"
    Else
        If txt_RE <> "" Then TextoFiltro = "EC.IDestoque = " & Cmb_RE Else TextoFiltro = "EC.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and EC.desenho = '" & Txt_cod_interno & "' and EC.lote = '" & Txt_lote & "' and EC.local_armaz = '" & txt_LA & "' and EC.Corrida = '" & Cmb_corrida & "' and EC.Certificado = '" & Cmb_certificado & "'"
    End If

    'Empenho de pedido interno
    qt = 0
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select Sum(Qtde_empenhada - ISNULL(Qtde_saida,0)) as qt from Estoque_Controle_Empenho_Vendas ECEV INNER JOIN Estoque_Controle EC ON ECEV.ID_estoque = EC.IdEstoque where " & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
        qt = IIf(IsNull(TBProduto!qt), 0, TBProduto!qt)
    End If
    
    'Empenho de ordem
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select Sum(Quantidade - ISNULL(Qtde_saida,0)) as qt from Producao_NF_Consignada PNC INNER JOIN Estoque_Controle EC ON PNC.IDestoque = EC.IdEstoque where " & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
        qt = qt + IIf(IsNull(TBProduto!qt), 0, TBProduto!qt)
    End If
    
    Qtd = Txt_qtde_fisica
    If qt > Qtd Then
        USMsgBox ("Não é possivel salvar inventário, porque a quantidade física informada é menor que a quantidade empenhada."), vbExclamation, "CAPRIND v5.0"
        Exit Sub
    End If
End If

Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from Estoque_fisico where ID = " & Txt_ID, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then TBGravar.AddNew
ProcEnviaDados

If txt_RE = "" Then
    Qtde = Txt_qtde_fisica
    NovoValor = Replace(Qtde, ",", ".")
    NovoValor1 = Replace(QtdePC, ",", ".")
    If Txt_lote = "" Then
        TextoFiltro = "ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and desenho = '" & Txt_cod_interno & "'"
    Else
        If txt_RE <> "" Then TextoFiltro = "IDestoque = " & txt_RE Else TextoFiltro = "ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and desenho = '" & Txt_cod_interno & "' and lote = '" & Txt_lote & "' and local_armaz = '" & txt_LA & "' and Corrida = '" & Cmb_corrida & "' and Certificado = '" & Cmb_certificado & "'"
    End If
    Conexao.Execute "Update estoque_controle Set qtde_fisica = " & NovoValor & ", qtde_fisica_PC = " & NovoValor1 & " where " & TextoFiltro
End If

TBGravar.Update
Txt_ID = TBGravar!ID
TBGravar.Close
If Novo_Estoque_Fisico = True Then
    USMsgBox ("Novo estoque físico cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo"
    Sql_Estoque_Fisico_Localizar = "Select Estoque_fisico.*, projproduto.Desenho, projproduto.Descricao, projproduto.Unidade from Estoque_fisico INNER JOIN projproduto on Estoque_fisico.Codproduto = projproduto.Codproduto where Estoque_fisico.ID = " & Txt_ID
    Sql_Estoque_Fisico_LocalizarTotal = "Select Sum(Estoque_fisico.valor_unitario) as Valor_Cofins_Prod, Sum(Estoque_fisico.Qtde_estoque) as Valor_Cofins_Serv, Sum(Estoque_fisico.valor_unitario * Estoque_fisico.Qtde_estoque) as Valor_CSLL_Prod, Sum(Estoque_fisico.Qtde_fisica) as Valor_CSLL_Serv, Sum(Estoque_fisico.valor_unitario * Estoque_fisico.Qtde_fisica) as Valor_INSS_Serv, Sum(Estoque_fisico.Qtde_estoque - Estoque_fisico.Qtde_fisica) as Valor_IPI, Sum(Estoque_fisico.valor_unitario * (Estoque_fisico.Qtde_estoque - Estoque_fisico.Qtde_fisica)) as Valor_IRPJ_Prod from (Estoque_fisico INNER JOIN projproduto on Estoque_fisico.Codproduto = projproduto.Codproduto) INNER JOIN Projfamilia on Projfamilia.Familia = projproduto.Classe where Estoque_fisico.ID = " & Txt_ID
    ProcAtualizalista (1)
Else
    str_Qtde = Replace(Txt_qtde_fisica.Text, ".", "")
    str_Qtde = Replace(Qtde, ",", ".")
    
    str_valor = Replace(Txt_valor_unitario.Text, ".", "")
    str_valor = Replace(valor, ",", ".")
    
    str_valortotal = Txt_qtde_fisica.Text * Txt_valor_unitario.Text
    str_valortotal = Replace(str_valortotal, ",", ".")
    
    StrSql = "update Estoque_Movimentacao SET data = '" & Txt_data.Value & "', entrada = " & str_Qtde & ", vlrUnit = " & str_valor & ", vlrTotal = " & str_valortotal & ", Lote = '" & Txt_lote.Text & "', documento = '" & Txt_lote.Text & "' where idestoque = " & txt_RE.Text & " and operacao = 'ENTRADA_INVENTÁRIO'"
    Conexao.Execute StrSql
    
    StrSql = "update Estoque_Movimentacao SET Lote = '" & Txt_lote.Text & "', documento = '" & Txt_lote.Text & "' where idestoque = " & txt_RE.Text & " and operacao ='SAIDA_LOCAL_DE_ARMAZENAMENTO'"
    Conexao.Execute StrSql
    
    StrSql = "update Estoque_Movimentacao SET Lote = '" & Txt_lote.Text & "'  where idestoque = " & txt_RE.Text & ""
    Conexao.Execute StrSql
    
    StrSql = "update Estoque_controle SET data = '" & Txt_data.Value & "', Lote = '" & Txt_lote.Text & "', valor_unitario = " & str_valor & " where idestoque = " & txt_RE.Text & ""
    Conexao.Execute StrSql
    
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar"
    ProcAtualizalista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
    If Lista.ListItems.Count <> 0 And CodigoLista <> 0 Then
        Lista.SelectedItem = Lista.ListItems(CodigoLista)
        Lista.SetFocus
    End If
End If
'ProcCriaMovimentacaoEntrada
'==================================
Modulo = "Estoque/Inventário"
ID_documento = Txt_ID
Documento = "Data: " & Format(Txt_data, "dd/mm/yy") & " - Código interno: " & Txt_cod_interno
Documento1 = ""
ProcGravaEvento
'==================================
'Novo_Estoque_Fisico = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviaDados()
On Error GoTo tratar_erro

TBGravar!ID_empresa = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
TBGravar!Data = Txt_data
TBGravar!Responsavel = Txt_responsavel
TBGravar!Codproduto = Txt_ID_prod
TBGravar!Etiqueta = Txt_etiqueta
If txt_RE.Text = "" Then
    TBGravar!Novo_lote = True
    TBGravar!LOTE = Txt_lote
    TBGravar!Corrida = txt_Corrida
    TBGravar!Certificado = txt_Certificado
Else
    TBGravar!Novo_lote = False
    TBGravar!IDEstoque = IIf(txt_RE = "", Null, txt_RE)
    If Txt_lote <> "" Then
        TBGravar!LOTE = Txt_lote
        TBGravar!Corrida = txt_Corrida
        TBGravar!Certificado = txt_Certificado
    Else
        TBGravar!LOTE = Null
        TBGravar!Corrida = txt_Corrida
        TBGravar!Certificado = txt_Certificado
    End If
End If
TBGravar!local_armaz = txt_LA.Text
TBGravar!valor_unitario = Txt_valor_unitario
TBGravar!Qtde_estoque = IIf(Txt_qtde_estoque <> "", Txt_qtde_estoque, 0)
'TBGravar!Qtde_estoque_PC = IIf(Txt_qtde_estoque_PC = "", Null, Txt_qtde_estoque_PC)
TBGravar!Destino = IIf(cmbDestino = "Interno/Cliente", "Interno", "Terceiros")
TBGravar!qtde_fisica = Txt_qtde_fisica

If Chk_consignado.Value = 1 Then TBGravar!Consignado = True Else TBGravar!Consignado = False
If Cmb_tipo_cli_forn = "" Then TBGravar!Tipo_cli_forn = Null Else TBGravar!Tipo_cli_forn = IIf(Cmb_tipo_cli_forn = "Cliente", "C", "F")
TBGravar!ID_cli_forn = IIf(Txt_ID_cli_forn = "", Null, Txt_ID_cli_forn)

TBGravar!Cod_ref = IIf(txt_cod_Referencia = "", Null, txt_cod_Referencia)
TBGravar!Numero_serie = Txt_numero_serie

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
                    Set TBFI = CreateObject("adodb.recordset")
                    TBFI.Open "Select * from Estoque_fisico where ID = " & .ListItems(InitFor) & " and RespValidacao IS NOT NULL", Conexao, adOpenKeyset, adLockOptimistic
                    If TBFI.EOF = False Then
                        TBFI.Close
                        .ListItems(InitFor).Checked = False
                        GoTo Proximo
                    End If
                    TBFI.Close
                Else
                    If .ListItems.Item(InitFor).ListSubItems(15) = "SIM" Then
                        'Verifica se é um inventario antigo, se for ele não deixa cancelar validação por aqui
                        Set TBAbrir = CreateObject("adodb.recordset")
                        TBAbrir.Open "Select EF.ID, EM.IDestoque,  EM.IDoperacao from Estoque_fisico EF INNER JOIN Estoque_movimentacao EM ON EM.Id_inventario = EF.ID where EF.id = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
                        If TBAbrir.EOF = True And .ListItems(InitFor).ListSubItems(14) <> "" Then
                            Set TBAbrir = CreateObject("adodb.recordset")
                            TBAbrir.Open "Select EM.IDestoque from Estoque_movimentacao EM INNER JOIN Estoque_controle EC ON EM.IDestoque = EC.IDestoque where EC.etiqueta = '" & .ListItems(InitFor).ListSubItems(14) & "'", Conexao, adOpenKeyset, adLockOptimistic
                            If TBAbrir.EOF = False Then
                                .ListItems.Item(InitFor).Checked = False
                                TBAbrir.Close
                                GoTo Proximo
                            End If
                        ElseIf TBAbrir.EOF = False Then
                                Set TBFI = CreateObject("adodb.recordset")
                                TBFI.Open "Select Idoperacao from Estoque_movimentacao where IDestoque = " & TBAbrir!IDEstoque & " and IDoperacao > " & TBAbrir!IDoperacao & " and ID_inventario <> " & TBAbrir!ID, Conexao, adOpenKeyset, adLockOptimistic
                                If TBFI.EOF = False Then
                                    .ListItems.Item(InitFor).Checked = False
                                    TBFI.Close
                                    GoTo Proximo
                                End If
                                TBFI.Close
                        End If
                        TBAbrir.Close
                        
                        'Verifica se foi feito alguma baixa no empenho depois de inventariar, se sim ele não deixa cancelar
                        Set TBAbrir = CreateObject("adodb.recordset")
                        TBAbrir.Open "Select EC.IDestoque from (Estoque_Controle_Empenho_Vendas EV INNER JOIN Estoque_controle EC ON EV.Id_estoque = EC.IDestoque) INNER JOIN Estoque_movimentacao EM ON EM.IDestoque = EC.IDestoque where EM.Id_inventario = " & .ListItems(InitFor) & " and EM.Operacao = 'ENTRADA_INVENTÁRIO' and EV.qtde_saida > 0", Conexao, adOpenKeyset, adLockOptimistic
                        If TBAbrir.EOF = False Then
                            .ListItems.Item(InitFor).Checked = False
                            GoTo Proximo
                        End If
                        Set TBAbrir = CreateObject("adodb.recordset")
                        TBAbrir.Open "Select EC.IDestoque from (Producao_NF_Consignada PNFC INNER JOIN Estoque_controle EC ON PNFC.Idestoque = EC.IDestoque) INNER JOIN Estoque_movimentacao EM ON EM.IDestoque = EC.IDestoque where EM.Id_inventario = " & .ListItems(InitFor) & " and EM.Operacao = 'ENTRADA_INVENTÁRIO' and PNFC.qtde_saida > 0", Conexao, adOpenKeyset, adLockOptimistic
                        If TBAbrir.EOF = False Then
                            .ListItems.Item(InitFor).Checked = False
                            GoTo Proximo
                        End If
                        'Verifica se foi criado instrumento no RE
                        Set TBAbrir = CreateObject("adodb.recordset")
                        TBAbrir.Open "Select EC.IDestoque from (Instrumentos I INNER JOIN Estoque_controle EC ON I.Idestoque = EC.IDestoque) INNER JOIN Estoque_movimentacao EM ON EM.IDestoque = EC.IDestoque where EM.Id_inventario = " & .ListItems(InitFor) & " and EM.Operacao = 'ENTRADA_INVENTÁRIO'", Conexao, adOpenKeyset, adLockOptimistic
                        If TBAbrir.EOF = False Then
                            .ListItems.Item(InitFor).Checked = False
                            GoTo Proximo
                        End If
                        TBAbrir.Close
                    End If
                End If
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView Lista, ColumnHeader
    TabelaRel = 1
    OrdenarRel = 1
    CampoRel = 1
    Select Case ColumnHeader
        Case "Cód. interno":
            TabelaRel = 3
            OrdenarRel = 1
            CampoRel = 1
        Case "Etiqueta":
            TabelaRel = 1
            OrdenarRel = 1
            CampoRel = 7
    End Select
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
                Set TBFI = CreateObject("adodb.recordset")
                TBFI.Open "Select * from Estoque_fisico where ID = " & .ListItems(InitFor) & " and RespValidacao IS NOT NULL", Conexao, adOpenKeyset, adLockOptimistic
                If TBFI.EOF = False Then
                    USMsgBox ("Não é permitido excluir este inventário, pois o mesmo já foi validado."), vbExclamation, "CAPRIND v5.0"
                    .ListItems.Item(InitFor).Checked = False
                    TBFI.Close
                    Exit Sub
                End If
                TBFI.Close
            Else
                If .ListItems.Item(InitFor).ListSubItems(13) = "SIM" Then
                    Set TBAbrir = CreateObject("adodb.recordset")
                    TBAbrir.Open "Select EF.ID, EM.IDestoque, EM.IDoperacao from Estoque_fisico EF INNER JOIN Estoque_movimentacao EM ON EM.Id_inventario = EF.ID where EF.id = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
                    If TBAbrir.EOF = True And .ListItems(InitFor).ListSubItems(12) <> "" Then
                        Set TBAbrir = CreateObject("adodb.recordset")
                        TBAbrir.Open "Select EM.IDestoque from Estoque_movimentacao EM INNER JOIN Estoque_controle EC ON EM.IDestoque = EC.IDestoque where EC.etiqueta = '" & .ListItems(InitFor).ListSubItems(12) & "'", Conexao, adOpenKeyset, adLockOptimistic
                        If TBAbrir.EOF = False Then
                            'USMsgBox ("Não é permitido cancelar a validação, é necessario excluir a(s) movimentação(ões)."), vbExclamation, "CAPRIND v5.0"
                            '.ListItems.Item(InitFor).Checked = False
                            'TBAbrir.Close
                            'Exit Sub
                        End If
                    ElseIf TBAbrir.EOF = False Then
                            Set TBFI = CreateObject("adodb.recordset")
                            TBFI.Open "Select Idoperacao from Estoque_movimentacao where IDestoque = " & TBAbrir!IDEstoque & " and IDoperacao > " & TBAbrir!IDoperacao & " and ID_inventario <> " & TBAbrir!ID, Conexao, adOpenKeyset, adLockOptimistic
                            If TBFI.EOF = False Then
                               ' USMsgBox ("Não é permitido cancelar a validação, é necessario excluir a(s) movimentação(ões) criadas após este inventário."), vbExclamation, "CAPRIND v5.0"
                               ' .ListItems.Item(InitFor).Checked = False
                               ' TBFI.Close
                               ' Exit Sub
                            End If
                            TBFI.Close
                    End If
                    TBAbrir.Close
                    
                    'Verifica se foi feito alguma baixa no empenho depois de inventariar, se sim ele não deixa cancelar
                    Set TBAbrir = CreateObject("adodb.recordset")
                    StrSql = "Select EC.IDestoque from (Estoque_Controle_Empenho_Vendas EV INNER JOIN Estoque_controle EC ON EV.Id_estoque = EC.IDestoque) INNER JOIN Estoque_movimentacao EM ON EM.IDestoque = EC.IDestoque where EM.Id_inventario = " & .ListItems(InitFor) & " and EM.Operacao = 'ENTRADA_INVENTÁRIO' and EV.qtde_saida > 0"
                    'Debug.print StrSql
                    
                    TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
                    If TBAbrir.EOF = False Then
                        'USMsgBox ("Não é permitido cancelar a validação, pois o inventário já sofreu movimentação(ões)."), vbExclamation, "CAPRIND v5.0"
                        '.ListItems.Item(InitFor).Checked = False
                        'Exit Sub
                    End If
                    Set TBAbrir = CreateObject("adodb.recordset")
                    TBAbrir.Open "Select EC.IDestoque from (Producao_NF_Consignada PNFC INNER JOIN Estoque_controle EC ON PNFC.Idestoque = EC.IDestoque) INNER JOIN Estoque_movimentacao EM ON EM.IDestoque = EC.IDestoque where EM.Id_inventario = " & .ListItems(InitFor) & " and EM.Operacao = 'ENTRADA_INVENTÁRIO' and PNFC.qtde_saida > 0", Conexao, adOpenKeyset, adLockOptimistic
                    If TBAbrir.EOF = False Then
                       ' USMsgBox ("Não é permitido cancelar a validação, pois o inventário já sofreu movimentação(ões)."), vbExclamation, "CAPRIND v5.0"
                        '.ListItems.Item(InitFor).Checked = False
                        'Exit Sub
                    End If
                    'Verifica se foi criado instrumento no RE
                    Set TBAbrir = CreateObject("adodb.recordset")
                    TBAbrir.Open "Select EC.IDestoque from (Instrumentos I INNER JOIN Estoque_controle EC ON I.Idestoque = EC.IDestoque) INNER JOIN Estoque_movimentacao EM ON EM.IDestoque = EC.IDestoque where EM.Id_inventario = " & .ListItems(InitFor) & " and EM.Operacao = 'ENTRADA_INVENTÁRIO'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBAbrir.EOF = False Then
                       ' USMsgBox ("Não é permitido cancelar a validação, pois o inventário já foi criado como instrumento."), vbExclamation, "CAPRIND v5.0"
                      '  .ListItems.Item(InitFor).Checked = False
                      '  Exit Sub
                    End If
                    TBAbrir.Close
                End If
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

'===============================================================
' Busca Saldo da RE na movimentação
'===============================================================
Set TBEstoque = CreateObject("adodb.recordset")
StrSql = "Select  ISNULL(Sum(Entrada) - Sum(Saida), 0) As Saldo from Estoque_Movimentacao where IDEstoque =  " & Lista.SelectedItem.ListSubItems.Item(1).Text
'Debug.print StrSql

TBEstoque.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
If TBEstoque.EOF = False Then
Saldo_Atual = TBEstoque!Saldo
End If
TBEstoque.Close
'===============================================================

Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select EF.*, P.Desenho, P.Descricao, P.Unidade, P.Classe from Estoque_fisico EF INNER JOIN projproduto P ON EF.Codproduto = P.Codproduto where EF.ID = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    ProcLimpaCampos
    ProcPuxaDados
    CodigoLista = Lista.SelectedItem.index
End If
TBProduto.Close
Frame1.Enabled = True
Novo_Estoque_Fisico = False
       
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_cod_interno_Change()
On Error GoTo tratar_erro

    Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select codproduto from projproduto where desenho = '" & Txt_cod_interno.Text & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Txt_ID_prod = TBAbrir!Codproduto
        End If
    TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_ID_cli_forn_Change()
On Error GoTo tratar_erro

Txt_cli_forn = ""
If Txt_ID_cli_forn <> "" Then
    VerifNumero = Txt_ID_cli_forn
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_ID_cli_forn = ""
        Txt_ID_cli_forn.SetFocus
        Exit Sub
    End If
    Set TBAbrir = CreateObject("adodb.recordset")
    If Cmb_tipo_cli_forn = "Cliente" Then
        TBAbrir.Open "Select NomeRazao from Clientes where idcliente = " & Txt_ID_cli_forn & " and Prospecto = 'False' and DtValidacao IS NOT NULL and status <> 'Bloqueado'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Txt_cli_forn = TBAbrir!NomeRazao
        End If
    Else
        TBAbrir.Open "Select Nome_Razao from compras_fornecedores where idcliente = " & Txt_ID_cli_forn & " and Prospecto = 'False' and DtValidacao IS NOT NULL and status <> 'Bloqueado'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Txt_cli_forn = TBAbrir!Nome_Razao
        End If
    End If
    TBAbrir.Close
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_qtde_fisica_Change()
On Error GoTo tratar_erro

If Txt_qtde_fisica <> "" Then
    VerifNumero = Txt_qtde_fisica
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_qtde_fisica = ""
        Txt_qtde_fisica.SetFocus
        Exit Sub
    End If
Saldo_Atual = Txt_valor_total.Text = Format(Saldo_Atual * valor_unitario, "###,##0.0000")
End If

If Saldo_Atual <> 0 And valor_unitario <> "" Then
Txt_valor_total.Text = Format(Saldo_Atual * valor_unitario, "###,##0.0000")
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_qtde_fisica_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus Txt_qtde_fisica

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_qtde_fisica_LostFocus()
On Error GoTo tratar_erro

Txt_qtde_fisica = Format(Txt_qtde_fisica, "###,##0.0000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub Txt_valor_unitario_Change()
On Error GoTo tratar_erro

If Txt_valor_unitario <> "" Then
    VerifNumero = Txt_valor_unitario
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_valor_unitario = ""
        Txt_valor_unitario.SetFocus
        Exit Sub
    End If
valor_unitario = Txt_valor_unitario.Text
End If

If Saldo_Atual <> 0 And valor_unitario <> "" Then
Txt_valor_total.Text = Format(Saldo_Atual * valor_unitario, "###,##0.000000")
End If

    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_valor_unitario_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus Txt_valor_unitario

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_valor_unitario_LostFocus()
On Error GoTo tratar_erro

Txt_valor_unitario = Format(Txt_valor_unitario, "###,##0.000000")

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

Private Sub ProcImportarPlanilha()
On Error GoTo tratar_erro

FrmEstoque_fisico_inventario.Show 1

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

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcNovo
    Case 2: ProcFiltrar
    Case 3: ProcSalvar
    Case 4: ProcExcluir
    Case 5: ProcImprimir
    Case 6: ProcValidarRegistros Lista, "Estoque/Inventário"
    Case 7: ProcImportarPlanilha
   
    'Case 8: ProcAjuda
    Case 10: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
