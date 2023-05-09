VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm_necessidade 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Estoque - Necessidade"
   ClientHeight    =   10035
   ClientLeft      =   1695
   ClientTop       =   1335
   ClientWidth     =   15360
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10035
   ScaleWidth      =   15360
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
      FormWidthDT     =   15480
      FormScaleHeightDT=   10035
      FormScaleWidthDT=   15360
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
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
      Height          =   870
      Left            =   55
      TabIndex        =   36
      Top             =   3015
      Width           =   15285
      Begin VB.TextBox txt_Necessidade_Estoque 
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
         Left            =   7076
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "Necessidade de estoque."
         Top             =   450
         Width           =   1309
      End
      Begin VB.TextBox Txt_necessidade_PC 
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
         Left            =   13680
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         ToolTipText     =   "Necessidade em peça."
         Top             =   450
         Width           =   1309
      End
      Begin VB.TextBox Txt_disponibilidade 
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
         Left            =   11036
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   "Disponibilidade."
         Top             =   450
         Width           =   1309
      End
      Begin VB.TextBox Txt_qtde_produzindo 
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
         Left            =   9716
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade produzindo."
         Top             =   450
         Width           =   1309
      End
      Begin VB.TextBox Txt_qtde_comprada 
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
         Left            =   8396
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade comprada."
         Top             =   450
         Width           =   1309
      End
      Begin VB.TextBox Txt_necessidade 
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
         Left            =   12356
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         ToolTipText     =   "Necessidade."
         Top             =   450
         Width           =   1309
      End
      Begin VB.TextBox Txt_qtde_estoque 
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
         Left            =   5756
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade em estoque."
         Top             =   450
         Width           =   1309
      End
      Begin VB.TextBox Txt_total_necessario 
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
         Left            =   4140
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Total necessário."
         Top             =   450
         Width           =   1605
      End
      Begin VB.TextBox Txt_estoque_minimo 
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
         Left            =   2820
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Estoque mínimo."
         Top             =   450
         Width           =   1309
      End
      Begin VB.TextBox Txt_necessidade_OP 
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
         Left            =   1500
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Necessidade da ordem de produção."
         Top             =   450
         Width           =   1309
      End
      Begin VB.TextBox Txt_necessidade_pedido 
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
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Necessidade do pedido."
         Top             =   450
         Width           =   1309
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Necess. est."
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
         Left            =   7228
         TabIndex        =   62
         Top             =   240
         Width           =   1005
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Necess. PÇ"
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
         Left            =   13877
         TabIndex        =   61
         Top             =   240
         Width           =   915
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Disponibilidade"
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
         Left            =   11038
         TabIndex        =   58
         Top             =   240
         Width           =   1305
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qtde. produzindo"
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
         Left            =   9733
         TabIndex        =   55
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qtde. comprada"
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
         Left            =   8458
         TabIndex        =   43
         Top             =   240
         Width           =   1185
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Necessidade"
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
         Left            =   12478
         TabIndex        =   42
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qtde. estoque"
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
         Left            =   5878
         TabIndex        =   41
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total necessário"
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
         Left            =   4230
         TabIndex        =   40
         Top             =   240
         Width           =   1425
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Estoque mínimo"
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
         Left            =   2912
         TabIndex        =   39
         Top             =   240
         Width           =   1125
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Necessidade OP"
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
         Left            =   1562
         TabIndex        =   38
         Top             =   240
         Width           =   1185
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Necess. PI"
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
         Left            =   437
         TabIndex        =   37
         Top             =   240
         Width           =   795
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
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
      Height          =   495
      Left            =   60
      TabIndex        =   63
      Top             =   1320
      Width           =   15285
      Begin VB.OptionButton Opt_PCP 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Por PCP"
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
         Left            =   180
         TabIndex        =   0
         Top             =   210
         Value           =   -1  'True
         Width           =   945
      End
      Begin VB.OptionButton Opt_vendas 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Por vendas"
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
         Left            =   1320
         TabIndex        =   1
         Top             =   210
         Width           =   1245
      End
      Begin VB.CheckBox Chk_estoque_minimo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Estoque mínimo"
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
         Left            =   3180
         TabIndex        =   9
         Top             =   210
         Width           =   1665
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   60
      TabIndex        =   52
      Top             =   9090
      Width           =   15285
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
         Left            =   3780
         TabIndex        =   27
         Text            =   "10"
         ToolTipText     =   "Número de registros por página."
         Top             =   180
         Width           =   555
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
         TabIndex        =   28
         ToolTipText     =   "Número da página."
         Top             =   180
         Width           =   555
      End
      Begin DrawSuite2022.USButton cmdPagIr 
         Height          =   315
         Left            =   10110
         TabIndex        =   29
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
      Begin DrawSuite2022.USButton cmdPagProx 
         Height          =   315
         Left            =   11760
         TabIndex        =   32
         ToolTipText     =   "Próxima página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "Frm_necessidade.frx":0000
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
         TabIndex        =   31
         ToolTipText     =   "Página anterior."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "Frm_necessidade.frx":37A4
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
      Begin DrawSuite2022.USButton cmdPagPrim 
         Height          =   315
         Left            =   10680
         TabIndex        =   30
         ToolTipText     =   "Primeira página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "Frm_necessidade.frx":72AD
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
         TabIndex        =   33
         ToolTipText     =   "Última página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "Frm_necessidade.frx":B39C
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
      Begin VB.Label Label16 
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
         Left            =   4410
         TabIndex        =   64
         Top             =   240
         Width           =   1440
      End
      Begin VB.Label Label24 
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
         Left            =   3090
         TabIndex        =   59
         Top             =   240
         Width           =   645
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
         TabIndex        =   54
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
         TabIndex        =   53
         Top             =   240
         Width           =   1275
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   51
      Top             =   330
      Width           =   15315
      _ExtentX        =   27014
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
      ButtonCaption2  =   "RM"
      ButtonEnabled2  =   0   'False
      ButtonIconSize2 =   32
      ButtonToolTipText2=   "Alterar requisição de material da ordem."
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
      ButtonState2    =   5
      ButtonLeft2     =   40
      ButtonTop2      =   2
      ButtonWidth2    =   23
      ButtonHeight2   =   21
      ButtonUseMaskColor2=   0   'False
      ButtonCaption3  =   "Relatório"
      ButtonEnabled3  =   0   'False
      ButtonIconSize3 =   32
      ButtonToolTipText3=   "Relatório (F5)"
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
      ButtonLeft3     =   65
      ButtonTop3      =   2
      ButtonWidth3    =   51
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
      ButtonLeft4     =   118
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
      ButtonLeft5     =   122
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
      ButtonLeft6     =   160
      ButtonTop6      =   2
      ButtonWidth6    =   26
      ButtonHeight6   =   21
      ButtonUseMaskColor6=   0   'False
      ButtonEnabled7  =   0   'False
      ButtonIconSize7 =   32
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
      ButtonState7    =   5
      ButtonLeft7     =   188
      ButtonTop7      =   2
      ButtonWidth7    =   24
      ButtonHeight7   =   24
      ButtonUseMaskColor7=   0   'False
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   13890
         Top             =   180
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "Frm_necessidade.frx":EC28
         Count           =   1
      End
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   60
      TabIndex        =   56
      Top             =   9720
      Width           =   15285
      _ExtentX        =   26961
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
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
      Height          =   1185
      Left            =   60
      TabIndex        =   44
      Top             =   1830
      Width           =   15285
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0E0&
         Height          =   1185
         Left            =   0
         TabIndex        =   47
         Top             =   0
         Width           =   1545
         Begin VB.OptionButton optIgual 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Igual"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   180
            TabIndex        =   16
            ToolTipText     =   "Física"
            Top             =   870
            Width           =   705
         End
         Begin VB.OptionButton Optmeio 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Meio frase"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   180
            TabIndex        =   13
            ToolTipText     =   "Física"
            Top             =   420
            Width           =   1305
         End
         Begin VB.OptionButton Optinicio 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Início frase"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   180
            TabIndex        =   11
            ToolTipText     =   "Jurídica"
            Top             =   180
            Value           =   -1  'True
            Width           =   1305
         End
         Begin VB.OptionButton Optfim 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Fim frase"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   180
            TabIndex        =   15
            ToolTipText     =   "Física"
            Top             =   660
            Width           =   1305
         End
      End
      Begin VB.ComboBox Cmb_filtrar 
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
         ItemData        =   "Frm_necessidade.frx":1202E
         Left            =   12540
         List            =   "Frm_necessidade.frx":12044
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         ToolTipText     =   "Tipo de necessidade."
         Top             =   510
         Width           =   2475
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
         ItemData        =   "Frm_necessidade.frx":120B2
         Left            =   1680
         List            =   "Frm_necessidade.frx":120B9
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "Empresa."
         Top             =   510
         Width           =   4035
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
         ItemData        =   "Frm_necessidade.frx":120CA
         Left            =   5730
         List            =   "Frm_necessidade.frx":120DD
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "Opções para filtro."
         Top             =   510
         Width           =   1965
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
         Left            =   7710
         TabIndex        =   4
         ToolTipText     =   "Texto para pesquisa."
         Top             =   510
         Width           =   4815
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
         ItemData        =   "Frm_necessidade.frx":12128
         Left            =   7710
         List            =   "Frm_necessidade.frx":1212F
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         ToolTipText     =   "Texto para pesquisa."
         Top             =   510
         Visible         =   0   'False
         Width           =   4815
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de necessidade"
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
         Left            =   12915
         TabIndex        =   60
         Top             =   300
         Width           =   1710
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
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
         Left            =   3315
         TabIndex        =   57
         Top             =   300
         Width           =   765
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
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
         Left            =   6285
         TabIndex        =   46
         Top             =   300
         Width           =   855
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Texto para pesquisa"
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
         Left            =   9360
         TabIndex        =   45
         Top             =   300
         Width           =   1485
      End
   End
   Begin TabDlg.SSTab SStab1 
      Height          =   10035
      Left            =   0
      TabIndex        =   35
      Top             =   0
      Width           =   15480
      _ExtentX        =   27305
      _ExtentY        =   17701
      _Version        =   393216
      Tabs            =   2
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
      TabCaption(0)   =   "Resumido"
      TabPicture(0)   =   "Frm_necessidade.frx":1213F
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Lista"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Lista_necessidade"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Detalhado"
      TabPicture(1)   =   "Frm_necessidade.frx":1215B
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame6"
      Tab(1).Control(1)=   "Lista_detalhado"
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame6 
         BackColor       =   &H00E0E0E0&
         Height          =   1185
         Left            =   -61860
         TabIndex        =   48
         Top             =   1830
         Width           =   2115
         Begin MSComCtl2.DTPicker msk_fltFim 
            Height          =   315
            Left            =   630
            TabIndex        =   8
            ToolTipText     =   "Data final."
            Top             =   660
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
            Format          =   199032833
            CurrentDate     =   39057
         End
         Begin MSComCtl2.DTPicker msk_fltInicio 
            Height          =   315
            Left            =   630
            TabIndex        =   7
            ToolTipText     =   "Data inicio."
            Top             =   300
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
            Format          =   199032833
            CurrentDate     =   39057
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Até :"
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
            TabIndex        =   50
            Top             =   720
            Width           =   360
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "De :"
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
            Top             =   360
            Width           =   300
         End
      End
      Begin MSComctlLib.ListView Lista_detalhado 
         Height          =   5175
         Left            =   -74940
         TabIndex        =   34
         Top             =   3900
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   9128
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList1"
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
         NumItems        =   18
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "N"
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Tag             =   "N"
            Text            =   "Cód. interno"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Descrição"
            Object.Width           =   3590
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Object.Tag             =   "N"
            Text            =   "Qtde."
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "Un."
            Object.Width           =   970
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Object.Tag             =   "N"
            Text            =   "Qtde. PÇ"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Object.Tag             =   "T"
            Text            =   "Ped. int."
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Object.Tag             =   "N"
            Text            =   "Ordem"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   8
            Object.Tag             =   "D"
            Text            =   "Prazo"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Object.Tag             =   "N"
            Text            =   "L. time prod."
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   10
            Object.Tag             =   "D"
            Text            =   "Pr. início OP"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Object.Tag             =   "N"
            Text            =   "L. time compras"
            Object.Width           =   2028
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   12
            Object.Tag             =   "D"
            Text            =   "Pr. emissão PC"
            Object.Width           =   2028
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Object.Tag             =   "T"
            Text            =   "Cód. interno"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   14
            Object.Tag             =   "T"
            Text            =   "Rev."
            Object.Width           =   970
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Object.Tag             =   "T"
            Text            =   "Cód. ref."
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Object.Tag             =   "T"
            Text            =   "Descrição"
            Object.Width           =   3590
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   17
            Object.Tag             =   "T"
            Text            =   "Cliente"
            Object.Width           =   3590
         EndProperty
      End
      Begin MSComctlLib.ListView Lista_necessidade 
         Height          =   2940
         Left            =   60
         TabIndex        =   26
         Top             =   6135
         Width           =   15285
         _ExtentX        =   26961
         _ExtentY        =   5186
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList1"
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
         NumItems        =   16
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "N"
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Object.Tag             =   "N"
            Text            =   "Qtde."
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Un."
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Object.Tag             =   "N"
            Text            =   "Qtde. PÇ"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "PI"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Object.Tag             =   "N"
            Text            =   "OP"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Object.Tag             =   "D"
            Text            =   "Prazo"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Object.Tag             =   "N"
            Text            =   "L. time prod."
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   8
            Object.Tag             =   "D"
            Text            =   "Pr. início OP"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Object.Tag             =   "N"
            Text            =   "L. time compras"
            Object.Width           =   2028
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   10
            Object.Tag             =   "D"
            Text            =   "Pr. emissão PC"
            Object.Width           =   2028
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Object.Tag             =   "T"
            Text            =   "Cód. interno"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   12
            Object.Tag             =   "T"
            Text            =   "Rev."
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Object.Tag             =   "T"
            Text            =   "Cód. de ref."
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Object.Tag             =   "T"
            Text            =   "Descrição"
            Object.Width           =   4503
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Object.Tag             =   "T"
            Text            =   "Cliente"
            Object.Width           =   4503
         EndProperty
      End
      Begin MSComctlLib.ListView Lista 
         Height          =   2220
         Left            =   60
         TabIndex        =   25
         Top             =   3900
         Width           =   15285
         _ExtentX        =   26961
         _ExtentY        =   3916
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
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   18
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "N"
            Text            =   "Cod. produto"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Tag             =   "T"
            Text            =   "Cód. interno"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Descrição"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Object.Tag             =   "T"
            Text            =   "Un."
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Object.Tag             =   "N"
            Text            =   "Lote mín."
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Object.Tag             =   "N"
            Text            =   "N. ped."
            Object.Width           =   1147
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Object.Tag             =   "N"
            Text            =   "N. OP."
            Object.Width           =   1147
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Object.Tag             =   "N"
            Text            =   "Est. mín."
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Object.Tag             =   "N"
            Text            =   "Neces,"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   9
            Object.Tag             =   "N"
            Text            =   "Est.."
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   10
            Object.Tag             =   "N"
            Text            =   "Nec.Est."
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   11
            Object.Tag             =   "N"
            Text            =   "Comp."
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   12
            Object.Tag             =   "N"
            Text            =   "Prod."
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   13
            Object.Tag             =   "N"
            Text            =   "Disp."
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   14
            Object.Tag             =   "N"
            Text            =   "Necess."
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   15
            Object.Tag             =   "N"
            Text            =   "Necess. PÇ"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   16
            Text            =   "Part Number"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   17
            Text            =   "Fabricante"
            Object.Width           =   4657
         EndProperty
      End
   End
End
Attribute VB_Name = "Frm_necessidade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public StrSql_EstoqueNecessidade As String 'OK
Public FormulaRel_EstoqueNecessidade As String 'OK
Dim TBLISTA_EstoqueNecessidade  As ADODB.Recordset 'OK

Sub ProcCarregaLista()
On Error GoTo tratar_erro

If StrSql_EstoqueNecessidade = "" Then Exit Sub
lblRegistros.Caption = "Nº de registros: 0"
lblPaginas.Caption = "Página: 0 de: 0"
Lista.ListItems.Clear
Lista_detalhado.ListItems.Clear
If StrSql_EstoqueNecessidade = "" Then Exit Sub
Set TBLISTA_EstoqueNecessidade = CreateObject("adodb.recordset")
TBLISTA_EstoqueNecessidade.Open StrSql_EstoqueNecessidade, Conexao, adOpenKeyset, adLockReadOnly
If TBLISTA_EstoqueNecessidade.EOF = False Then ProcExibePagina (1)
'Debug.print StrSql_EstoqueNecessidade

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina(Pagina)
On Error GoTo tratar_erro

Cotacao = 0
Lista.ListItems.Clear
Lista_detalhado.ListItems.Clear
TBLISTA_EstoqueNecessidade.PageSize = IIf(txtNreg = "", 30, txtNreg)
TBLISTA_EstoqueNecessidade.AbsolutePage = Pagina
TamanhoPagina = TBLISTA_EstoqueNecessidade.PageSize
ContadorReg = 1

PBLista.Min = 0
PBLista.Max = FunVerifMaxPBListaPaginacao(TBLISTA_EstoqueNecessidade.RecordCount - IIf(Pagina > 1, (TBLISTA_EstoqueNecessidade.PageSize * (Pagina - 1)), 0), TBLISTA_EstoqueNecessidade.PageSize)
PBLista.Value = 1
Contador = 0
Do While TBLISTA_EstoqueNecessidade.EOF = False And (ContadorReg <= TamanhoPagina)
    If SSTab1.Tab = 0 Then
        With Lista.ListItems.Add(, , TBLISTA_EstoqueNecessidade!Codproduto)
            .SubItems(1) = IIf(IsNull(TBLISTA_EstoqueNecessidade!Desenho), "", TBLISTA_EstoqueNecessidade!Desenho)
            .SubItems(2) = IIf(IsNull(TBLISTA_EstoqueNecessidade!Descricao), "", TBLISTA_EstoqueNecessidade!Descricao)
            .SubItems(3) = IIf(IsNull(TBLISTA_EstoqueNecessidade!Unidade), "", TBLISTA_EstoqueNecessidade!Unidade)
            .SubItems(4) = Format(TBLISTA_EstoqueNecessidade!qtde_LoteMinimo, "###,##0.00")
            If Opt_PCP.Value = True Then
                .SubItems(5) = TBLISTA_EstoqueNecessidade!Necessidade_PI
                .SubItems(6) = Format(TBLISTA_EstoqueNecessidade!Necessidade_OP, "###,##0.00")
            Else
                .SubItems(5) = Format(TBLISTA_EstoqueNecessidade!Necessidade_PI + TBLISTA_EstoqueNecessidade!Necessidade_PIEST, "###,##0.0000")
                .SubItems(6) = "0,0000"
            End If
            .SubItems(7) = Format(TBLISTA_EstoqueNecessidade!Estoque_minimo, "###,##0.00")
            .SubItems(8) = Format(TBLISTA_EstoqueNecessidade!Total_necessario, "###,##0.00")
            .SubItems(9) = Format(TBLISTA_EstoqueNecessidade!Qtde_estoque, "###,##0.00")
            .SubItems(10) = Format(TBLISTA_EstoqueNecessidade!Necessidade_estoque, "###,##0.00")
            .SubItems(11) = Format(TBLISTA_EstoqueNecessidade!Qtde_comprada, "###,##0.00")
            .SubItems(12) = TBLISTA_EstoqueNecessidade!Qtde_produzindo
            .SubItems(13) = Format(TBLISTA_EstoqueNecessidade!Disponibilidade, "###,##0.00")
            .SubItems(14) = Format(TBLISTA_EstoqueNecessidade!Necessidade, "###,##0.00")
            .SubItems(15) = TBLISTA_EstoqueNecessidade!Necessidade_PC
            
            Set TBItem = CreateObject("adodb.recordset")
            StrSql = "Select Part_number, Fabricante from Projproduto_fabricante PF Inner join Fabricante_Marca FM on PF.IDFabricante = FM.Id where codproduto = " & TBLISTA_EstoqueNecessidade!Codproduto
            'Debug.print StrSql
            
            TBItem.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
            If TBItem.EOF = False Then
                .SubItems(16) = IIf(IsNull(TBItem!Part_number), "", (TBItem!Part_number))
                .SubItems(17) = IIf(IsNull(TBItem!Fabricante), "", (TBItem!Fabricante))
            Else
                .SubItems(16) = ""
                .SubItems(17) = ""
            End If
            
            
            If Cmb_filtrar = "Todos necess. estoque" Or Cmb_filtrar = "Sem necess. estoque" Or Cmb_filtrar = "Com necess. estoque" Then
                NReal = Format(TBLISTA_EstoqueNecessidade!Necessidade_estoque, "###,##0.00")
            Else
                NReal = Format(TBLISTA_EstoqueNecessidade!Necessidade, "###,##0.00")
            End If
            If NReal > 0 Then
                .ForeColor = vbRed
                .ListSubItems(1).ForeColor = vbRed
                .ListSubItems(2).ForeColor = vbRed
                .ListSubItems(3).ForeColor = vbRed
                .ListSubItems(4).ForeColor = vbRed
                .ListSubItems(5).ForeColor = vbRed
                .ListSubItems(6).ForeColor = vbRed
                .ListSubItems(7).ForeColor = vbRed
                .ListSubItems(8).ForeColor = vbRed
                .ListSubItems(9).ForeColor = vbRed
                .ListSubItems(10).ForeColor = vbRed
                .ListSubItems(11).ForeColor = vbRed
                .ListSubItems(12).ForeColor = vbRed
                .ListSubItems(13).ForeColor = vbRed
                .ListSubItems(14).ForeColor = vbRed
                .ListSubItems(15).ForeColor = vbRed
            End If
        End With
    Else
        With Lista_detalhado.ListItems
            .Add , , TBLISTA_EstoqueNecessidade!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA_EstoqueNecessidade!Desenho), "", TBLISTA_EstoqueNecessidade!Desenho)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA_EstoqueNecessidade!Descricao), "", TBLISTA_EstoqueNecessidade!Descricao)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA_EstoqueNecessidade!Requisitado), "0,0000", Format(TBLISTA_EstoqueNecessidade!Requisitado, "###,##0.0000"))
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA_EstoqueNecessidade!Unidade), "", TBLISTA_EstoqueNecessidade!Unidade)
            .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA_EstoqueNecessidade!Requisitado_PC), 0, TBLISTA_EstoqueNecessidade!Requisitado_PC)
            .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA_EstoqueNecessidade!Ped_req), "", TBLISTA_EstoqueNecessidade!Ped_req)
            .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA_EstoqueNecessidade!Ordem_req), "", TBLISTA_EstoqueNecessidade!Ordem_req)
            .Item(.Count).SubItems(8) = IIf(IsNull(TBLISTA_EstoqueNecessidade!Prazo), "", Format(TBLISTA_EstoqueNecessidade!Prazo, "dd/mm/yy"))
'            .Item(.Count).SubItems(9) = IIf(IsNull(TBLISTA!LTimeProducao), "", TBLISTA!LTimeProducao)
'            .Item(.Count).SubItems(10) = IIf(IsNull(TBLISTA!PrazoInicioOP), "", TBLISTA!PrazoInicioOP)
'            .Item(.Count).SubItems(11) = IIf(IsNull(TBLISTA!LtimeCompras), "", TBLISTA!LtimeCompras)
'            .Item(.Count).SubItems(12) = IIf(IsNull(TBLISTA!PrazoEmissaoPC), "", TBLISTA!PrazoEmissaoPC)
            .Item(.Count).SubItems(13) = IIf(IsNull(TBLISTA_EstoqueNecessidade!Desenho_doc_req), "", TBLISTA_EstoqueNecessidade!Desenho_doc_req)
            .Item(.Count).SubItems(14) = IIf(IsNull(TBLISTA_EstoqueNecessidade!Revitem_doc_req), "", TBLISTA_EstoqueNecessidade!Revitem_doc_req)
            .Item(.Count).SubItems(15) = IIf(IsNull(TBLISTA_EstoqueNecessidade!Ref_doc_req), "", TBLISTA_EstoqueNecessidade!Ref_doc_req)
            .Item(.Count).SubItems(16) = IIf(IsNull(TBLISTA_EstoqueNecessidade!Descricao_doc_req), "", TBLISTA_EstoqueNecessidade!Descricao_doc_req)
            .Item(.Count).SubItems(17) = IIf(IsNull(TBLISTA_EstoqueNecessidade!Cliente_doc_req), "", TBLISTA_EstoqueNecessidade!Cliente_doc_req)
        End With
    End If
    TBLISTA_EstoqueNecessidade.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista.Value = Contador
Loop
lblRegistros.Caption = "Nº de registros: " & TBLISTA_EstoqueNecessidade.RecordCount
If TBLISTA_EstoqueNecessidade.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "Página: 1 de: " & TBLISTA_EstoqueNecessidade.PageCount
ElseIf TBLISTA_EstoqueNecessidade.AbsolutePage = adPosEOF Then
        lblPaginas.Caption = "Página: " & TBLISTA_EstoqueNecessidade.PageCount & " de: " & TBLISTA_EstoqueNecessidade.PageCount
    Else
        lblPaginas.Caption = "Página: " & TBLISTA_EstoqueNecessidade.AbsolutePage - 1 & " de: " & TBLISTA_EstoqueNecessidade.PageCount
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcVerificaQtdeSaida(Codinterno As String, Ordem As Long)
On Error GoTo tratar_erro

QtdeSaida = 0
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select Saida from Qtde_saida_estoque_produto where Desenho = '" & Codinterno & "' and Ordem = " & Ordem, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    QtdeSaida = IIf(IsNull(TBAbrir!Saida), 0, TBAbrir!Saida)
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ActiveResize1_ResizeComplete()
On Error GoTo tratar_erro

ProcCorrigeTela

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_estoque_minimo_Click()
On Error GoTo tratar_erro

ProcLimpaCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_empresa_Click()
On Error GoTo tratar_erro

ProcLimpaCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_filtrar_Click()
On Error GoTo tratar_erro

ProcLimpaCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbFamilia_Click()
On Error GoTo tratar_erro

txtTexto = ""
ProcLimpaCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbfiltrarpor_Click()
On Error GoTo tratar_erro

txtTexto = ""
cmbfamilia.ListIndex = -1
If cmbfiltrarpor = "Família" Then
    cmbfamilia.Visible = True
    txtTexto.Visible = False
Else
    cmbfamilia.Visible = False
    txtTexto.Visible = True
End If
ProcLimpaCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

With msk_fltFim
    If FunVerificaDataFinal(msk_fltInicio.Value, .Value) = False Then
        .Value = Date
        .SetFocus
        Exit Sub
    End If
End With
ProcLimpaCampos

IDAntigo = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
If SSTab1.Tab = 0 Then
    If Opt_PCP.Value = True Then NomeTabela = "Estoque_necessidade_resumido" Else NomeTabela = "Estoque_necessidade_resumido_PIEST"
    ApelidoTabela = "ENR"
Else
    NomeTabela = "Estoque_necessidade_detalhado"
    ApelidoTabela = "ENDET"
End If

TextoFiltroEmpresa = "(" & ApelidoTabela & ".ID_empresa = " & IDAntigo & " or " & ApelidoTabela & ".ID_empresa IS NULL)"
TextoFiltroEmpresaRel = "({" & NomeTabela & ".ID_empresa} = " & IDAntigo & " or ISNULL({" & NomeTabela & ".ID_empresa}) = True)"

If Compras_Necessidade = True Then
    TextoFiltroAplic = " and " & ApelidoTabela & ".Compras = 'True'"
    TextoFiltroAplicRel = " and {" & NomeTabela & ".Compras} = True"
ElseIf PCP_Necessidade = True Then
        TextoFiltroAplic = " and " & ApelidoTabela & ".Producao = 'True'"
        TextoFiltroAplicRel = " and {" & NomeTabela & ".Producao} = True"
    Else
        TextoFiltroAplic = " and " & ApelidoTabela & ".Desenho is not null"
        TextoFiltroAplicRel = " and {" & NomeTabela & ".Desenho} <> 'Null'"
End If

If Chk_estoque_minimo.Value = 1 Then
    TextoFiltroEstoqueMinimo = " and " & ApelidoTabela & ".Estoque_minimo > 0"
    TextoFiltroEstoqueMinimoRel = " and {" & NomeTabela & ".Estoque_minimo} > 0"
Else
    TextoFiltroEstoqueMinimo = ""
    TextoFiltroEstoqueMinimoRel = ""
End If

If SSTab1.Tab = 0 Then
    If Cmb_filtrar = "Com necessidade" Then
        TextoFiltroEstoque = " and " & ApelidoTabela & ".Necessidade > 0"
        TextoFiltroEstoqueRel = " and {" & NomeTabela & ".Necessidade} > 0"
    ElseIf Cmb_filtrar = "Sem necessidade" Then
            TextoFiltroEstoque = " and " & ApelidoTabela & ".Necessidade <= 0"
            TextoFiltroEstoqueRel = " and {" & NomeTabela & ".Necessidade} <= 0"
        ElseIf Cmb_filtrar = "Com necess. estoque" Then
                TextoFiltroEstoque = " and " & ApelidoTabela & ".Necessidade_estoque > 0"
                TextoFiltroEstoqueRel = " and {" & NomeTabela & ".Necessidade_estoque} > 0"
            ElseIf Cmb_filtrar = "Sem necess. estoque" Then
                    TextoFiltroEstoque = " and " & ApelidoTabela & ".Necessidade_estoque <= 0"
                    TextoFiltroEstoqueRel = " and {" & NomeTabela & ".Necessidade_estoque} <= 0"
                Else
                    TextoFiltroEstoque = ""
                    TextoFiltroEstoqueRel = ""
    End If
    CamposFiltro = "ENR.Codproduto, ENR.Desenho, ENR.Descricao, ENR.Unidade, ENR.qtde_LoteMinimo, ENR.Necessidade_PI, ENR.Estoque_minimo, ENR.Total_necessario, ENR.Qtde_estoque, ENR.Necessidade_estoque, ENR.Qtde_comprada, ENR.Qtde_produzindo, ENR.Disponibilidade, ENR.Necessidade, ENR.Necessidade_PC"
    If Opt_PCP.Value = True Then CamposFiltro = CamposFiltro & ", ENR.Necessidade_OP" Else CamposFiltro = CamposFiltro & ", ENR.Necessidade_PIEST"
    INNERJOINTEXTO = "Select " & CamposFiltro & " from (" & NomeTabela & " ENR LEFT JOIN item_aplicacoes IA ON ENR.codproduto = IA.codproduto) LEFT JOIN Projproduto_fabricante PFAB ON PFAB.Codproduto = ENR.codproduto"
    TextoFiltroPadrao = TextoFiltroEmpresa & TextoFiltroAplic & TextoFiltroEstoque & TextoFiltroEstoqueMinimo & " group by " & CamposFiltro & " order by ENR.desenho"
    TextoFiltroPadraoRel = TextoFiltroEmpresaRel & TextoFiltroAplicRel & TextoFiltroEstoqueRel & TextoFiltroEstoqueMinimoRel
    
    If txtTexto.Text <> "" Or cmbfamilia <> "" Then
        If cmbfiltrarpor = "Família" Then
            StrSql_EstoqueNecessidade = INNERJOINTEXTO & " where classe = '" & cmbfamilia & "' and " & TextoFiltroPadrao
            FormulaRel_EstoqueNecessidade = "{" & NomeTabela & ".classe} = '" & cmbfamilia.Text & "' and " & TextoFiltroPadraoRel
        Else
            Select Case cmbfiltrarpor
                Case "Código interno":
                    TextoFiltro = ApelidoTabela & ".Desenho"
                    TextoFiltroRel = NomeTabela & ".Desenho"
                Case "Código de referência":
                    TextoFiltro = "IA.n_referencia"
                    TextoFiltroRel = "item_aplicacoes.n_referencia"
                Case "Descrição":
                    TextoFiltro = ApelidoTabela & ".Descricao"
                    TextoFiltroRel = NomeTabela & ".Descricao"
                Case "Part number":
                    TextoFiltro = "PFAB.Part_number"
                    TextoFiltroRel = "Projproduto_fabricante.Part_number"
            End Select
            StrSql_EstoqueNecessidade = INNERJOINTEXTO & " where " & TextoFiltro & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and " & TextoFiltroPadrao
            FormulaRel_EstoqueNecessidade = "{" & TextoFiltroRel & "}" & FunVerifTipoFiltroIMFRel(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and " & TextoFiltroPadraoRel
        End If
    Else
        StrSql_EstoqueNecessidade = INNERJOINTEXTO & " where " & TextoFiltroPadrao
        FormulaRel_EstoqueNecessidade = TextoFiltroPadraoRel
    End If
Else
    If Opt_PCP.Value = True Then
        TextoFiltroTipo = "ENDET.Tipo <> 'PIEST'"
        TextoFiltroTipoRel = "{" & NomeTabela & ".Tipo} <> 'PIEST'"
    Else
        TextoFiltroTipo = "ENDET.Tipo <> 'OP'"
        TextoFiltroTipoRel = "{" & NomeTabela & ".Tipo} <> 'OP'"
    End If
    CamposFiltro = "ENDET.ID_empresa, ENDET.ID, ENDET.Desenho, ENDET.descricao, ENDET.Requisitado, ENDET.Unidade, ENDET.Requisitado_PC, ENDET.Ped_req, ENDET.Ordem_req, ENDET.Prazo, ENDET.Desenho_doc_req, ENDET.Revitem_doc_req, ENDET.Ref_doc_req, ENDET.Descricao_doc_req, ENDET.Cliente_doc_req"
    INNERJOINTEXTO = "Select " & CamposFiltro & " from (Estoque_necessidade_detalhado ENDET LEFT JOIN item_aplicacoes IA ON ENDET.codproduto = IA.codproduto) LEFT JOIN Projproduto_fabricante PFAB ON PFAB.Codproduto = ENDET.codproduto"
    TextoFiltroPadrao = TextoFiltroTipo & " and Prazo Between '" & Format(msk_fltInicio.Value, "dd-mm-yyyy") & "' and '" & Format(msk_fltFim.Value, "dd-mm-yyyy") & "' and " & TextoFiltroEmpresa & TextoFiltroAplic & TextoFiltroEstoqueMinimo & " order by Desenho"
    TextoFiltroPadraoRel = TextoFiltroTipoRel & " and {" & NomeTabela & ".Prazo} >= Date(" & Year(msk_fltInicio.Value) & "," & Month(msk_fltInicio.Value) & "," & Day(msk_fltInicio.Value) & ") and {" & NomeTabela & ".Prazo} <= Date(" & Year(msk_fltFim.Value) & "," & Month(msk_fltFim.Value) & "," & Day(msk_fltFim.Value) & ") and " & TextoFiltroEmpresaRel & TextoFiltroAplicRel & TextoFiltroEstoqueMinioRel
    
    If txtTexto.Text <> "" Or cmbfamilia <> "" Then
        If cmbfiltrarpor = "Família" Then
            StrSql_EstoqueNecessidade = INNERJOINTEXTO & " where Classe = '" & cmbfamilia & "' and " & TextoFiltroPadrao
            FormulaRel_EstoqueNecessidade = "{" & NomeTabela & ".Classe} = '" & cmbfamilia & "' and " & TextoFiltroPadraoRel
        Else
            Select Case cmbfiltrarpor
                Case "Código interno":
                    TextoFiltro = "ENDET.Desenho"
                    TextoFiltroRel = NomeTabela & ".Desenho"
                Case "Código de referência":
                    TextoFiltro = "IA.n_referencia"
                    TextoFiltroRel = "item_aplicacoes.n_referencia"
                Case "Descrição":
                    TextoFiltro = "ENDET.Descricao"
                    TextoFiltroRel = NomeTabela & ".Descricao"
                Case "Part number":
                    TextoFiltro = "PFAB.Part_number"
                    TextoFiltroRel = "Projproduto_fabricante.Part_number"
            End Select
            StrSql_EstoqueNecessidade = INNERJOINTEXTO & " where " & TextoFiltro & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and " & TextoFiltroPadrao
            FormulaRel_EstoqueNecessidade = "{" & TextoFiltroRel & "}" & FunVerifTipoFiltroIMFRel(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and " & TextoFiltroPadrao
        End If
    Else
        StrSql_EstoqueNecessidade = INNERJOINTEXTO & " where " & TextoFiltroPadrao
        FormulaRel_EstoqueNecessidade = TextoFiltroPadraoRel
    End If
End If
'Debug.print StrSql_EstoqueNecessidade
ProcCarregaLista

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_EstoqueNecessidade.AbsolutePage <> 2 Then
    If TBLISTA_EstoqueNecessidade.AbsolutePage = -3 Then
        ProcExibePagina (TBLISTA_EstoqueNecessidade.PageCount - 1)
    Else
        TBLISTA_EstoqueNecessidade.AbsolutePage = TBLISTA_EstoqueNecessidade.AbsolutePage - 2
        ProcExibePagina (TBLISTA_EstoqueNecessidade.AbsolutePage)
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
    TBLISTA_EstoqueNecessidade.AbsolutePage = txtPagIr.Text
    ProcExibePagina (TBLISTA_EstoqueNecessidade.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_EstoqueNecessidade.AbsolutePage = 1
ProcExibePagina (TBLISTA_EstoqueNecessidade.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_EstoqueNecessidade.AbsolutePage <> -3 Then
    If TBLISTA_EstoqueNecessidade.AbsolutePage = 1 Then
        ProcExibePagina (2)
    Else
        ProcExibePagina (TBLISTA_EstoqueNecessidade.AbsolutePage)
    End If
Else
    ProcExibePagina (TBLISTA_EstoqueNecessidade.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_EstoqueNecessidade.AbsolutePage = TBLISTA_EstoqueNecessidade.PageCount
ProcExibePagina (TBLISTA_EstoqueNecessidade.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyEscape: Unload Me
    Case vbKeyF2: ProcFiltrar
    Case vbKeyF5: ProcImprimir
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15195, 7, True
Cmb_filtrar = "Com necessidade"
If Compras_Necessidade = True Then
    Caption = "Compras - Necessidade"
    Formulario = "Compras/Necessidade"
    ProcCarregaComboFamilia cmbfamilia, "Compras = 'True' and familia <> 'Null'", True
    Familiatext = "C"
ElseIf PCP_Necessidade = True Then
        Caption = "PCP - Necessidade"
        Formulario = "PCP/Necessidade"
        ProcCarregaComboFamilia cmbfamilia, "familia <> 'Null'", True
        Familiatext = "P"
    Else
        Caption = "Estoque - Necessidade"
        Formulario = "Estoque/Necessidade"
        ProcCarregaComboFamilia cmbfamilia, "familia <> 'Null'", True
        Familiatext = "T"
        Cmb_filtrar = "Com necess. estoque"
End If
ProcCarregaComboEmpresa Cmb_empresa, False
ProcFiltroPadrao cmbfiltrarpor, Optmeio, Optfim, optIgual, Cmb_empresa.ItemData(Cmb_empresa.ListIndex), "Produtos/Serviços", Familiatext, True
If Permitido = False Then cmbfiltrarpor = "Código interno"
Direitos
ProcLimpaVariaveisPrincipais
SSTab1.Tab = 0
msk_fltInicio = Date
msk_fltFim = Date
Formulario_necessidade = Formulario

ProcRemoveObjetosResize Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

If Compras_Necessidade = True Then
    Caption = "Compras - Necessidade"
    Formulario = "Compras/Necessidade"
ElseIf PCP_Necessidade = True Then
        Caption = "PCP - Necessidade"
        Formulario = "PCP/Necessidade"
    Else
        Caption = "Estoque - Necessidade"
        Formulario = "Estoque/Necessidade"
End If
Formulario_necessidade = Formulario

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro
    
If SSTab1.Tab = 0 Then
    If Lista.ListItems.Count = 0 Then Exit Sub
    If Opt_PCP.Value = True Then NomeRel = "Estoque_necessidade_resumido.rpt" Else NomeRel = "Estoque_necessidade_resumido_vendas.rpt"
Else
    If Lista_detalhado.ListItems.Count = 0 Then Exit Sub
    NomeRel = "Estoque_necessidade_detalhado.rpt"
End If
ProcImprimirRel FormulaRel_EstoqueNecessidade, ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro
    
Formulario_necessidade = ""
Unload Me

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
ProcCarregaDados Lista.SelectedItem.SubItems(1)
ProcCarregaListaEmpenho
        
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcCarregaListaEmpenho()
On Error GoTo tratar_erro

Lista_necessidade.ListItems.Clear
If Opt_PCP.Value = True Then TextoFiltro = "Tipo <> 'PIEST'" Else TextoFiltro = "Tipo <> 'OP'"
Set TBLISTA = CreateObject("adodb.recordset")
StrSql = "Select * from Estoque_necessidade_detalhado where ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and Desenho = '" & Lista.SelectedItem.ListSubItems(1) & "' and " & TextoFiltro
'Debug.print StrSql

TBLISTA.Open StrSql, Conexao, adOpenKeyset, adLockReadOnly
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBLISTA.EOF = False
        With Lista_necessidade.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Requisitado), 0, Format(TBLISTA!Requisitado, "###,##0.0000"))
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Unidade), "", TBLISTA!Unidade)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Requisitado_PC), 0, TBLISTA!Requisitado_PC)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!Ped_req), "", TBLISTA!Ped_req)
            .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!Ordem_req), "", TBLISTA!Ordem_req)
            .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA!Prazo), "", Format(TBLISTA!Prazo, "dd/mm/yy"))
'            .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA!LTimeProducao), "", TBLISTA!LTimeProducao)
'            .Item(.Count).SubItems(8) = IIf(IsNull(TBLISTA!PrazoInicioOP), "", TBLISTA!PrazoInicioOP)
'            .Item(.Count).SubItems(9) = IIf(IsNull(TBLISTA!LtimeCompras), "", TBLISTA!LtimeCompras)
'            .Item(.Count).SubItems(10) = IIf(IsNull(TBLISTA!PrazoEmissaoPC), "", TBLISTA!PrazoEmissaoPC)
            .Item(.Count).SubItems(11) = IIf(IsNull(TBLISTA!Desenho_doc_req), "", TBLISTA!Desenho_doc_req)
            .Item(.Count).SubItems(12) = IIf(IsNull(TBLISTA!Revitem_doc_req), "", TBLISTA!Revitem_doc_req)
            .Item(.Count).SubItems(13) = IIf(IsNull(TBLISTA!Ref_doc_req), "", TBLISTA!Ref_doc_req)
            .Item(.Count).SubItems(14) = IIf(IsNull(TBLISTA!Descricao_doc_req), "", TBLISTA!Descricao_doc_req)
            .Item(.Count).SubItems(15) = IIf(IsNull(TBLISTA!Cliente_doc_req), "", TBLISTA!Cliente_doc_req)
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

Public Sub ProcLimpaCampos()
On Error GoTo tratar_erro

Txt_necessidade_pedido = "0,00"
Txt_necessidade_OP = "0,00"
Txt_estoque_minimo = "0,00"
Txt_total_necessario = "0,00"
Txt_qtde_estoque = "0,00"
Txt_qtde_comprada = "0,00"
Txt_qtde_produzindo = 0
Txt_disponibilidade = "0,00"
Txt_necessidade = "0,00"
Txt_necessidade_PC = "0,00"
txt_Necessidade_Estoque = "0,00"
Lista.ListItems.Clear
Lista_necessidade.ListItems.Clear
Lista_detalhado.ListItems.Clear
lblRegistros.Caption = "Nº de registros: 0"
lblPaginas.Caption = "Página: 0 de: 0"
USToolBar1.ButtonState(2) = 5

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_detalhado_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView Lista_detalhado, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_detalhado_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista_detalhado.ListItems.Count = 0 Then Exit Sub
ProcCarregaDados Lista_detalhado.SelectedItem.SubItems(1)

If Lista_detalhado.SelectedItem.SubItems(7) = "" Or Lista_detalhado.SelectedItem.SubItems(7) = "0" Then
    USToolBar1.ButtonState(2) = 5
Else
    If FunVerifProdSimiliar(Cmb_empresa.ItemData(Cmb_empresa.ListIndex)) = True Then USToolBar1.ButtonState(2) = 0
End If
        
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcCarregaDados(Codinterno As String)
On Error GoTo tratar_erro

If Opt_PCP.Value = True Then TabelaFiltro = "Estoque_necessidade_resumido" Else TabelaFiltro = "Estoque_necessidade_resumido_PIEST"
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from " & TabelaFiltro & " where Desenho = '" & Codinterno & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    If Opt_PCP.Value = True Then
        Txt_necessidade_pedido = Format(TBLISTA!Necessidade_PI, "###,##0.0000")
        Txt_necessidade_OP = Format(TBLISTA!Necessidade_OP, "###,##0.0000")
    Else
        Txt_necessidade_pedido = Format(TBLISTA!Necessidade_PI + TBLISTA!Necessidade_PIEST, "###,##0.0000")
        Txt_necessidade_OP = "0,00"
    End If
    Txt_estoque_minimo = Format(TBLISTA!Estoque_minimo, "###,##0.0000")
    Txt_total_necessario = Format(TBLISTA!Total_necessario, "###,##0.0000")
    Txt_qtde_estoque = Format(TBLISTA!Qtde_estoque, "###,##0.0000")
    txt_Necessidade_Estoque = Format(TBLISTA!Necessidade_estoque, "###,##0.0000")
    Txt_qtde_comprada = Format(TBLISTA!Qtde_comprada, "###,##0.0000")
    Txt_qtde_produzindo = TBLISTA!Qtde_produzindo
    Txt_disponibilidade = Format(TBLISTA!Disponibilidade, "###,##0.0000")
    Txt_necessidade = Format(TBLISTA!Necessidade, "###,##0.0000")
    Txt_necessidade_PC = TBLISTA!Necessidade_PC
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_necessidade_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView Lista_necessidade, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_necessidade_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista_necessidade.ListItems.Count = 0 Then Exit Sub
With USToolBar1
    If Lista_necessidade.SelectedItem.SubItems(5) = "" Or Lista_necessidade.SelectedItem.SubItems(5) = "0" Then
        .ButtonState(2) = 5
    Else
        If FunVerifProdSimiliar(Cmb_empresa.ItemData(Cmb_empresa.ListIndex)) = True Then .ButtonState(2) = 0 Else .ButtonState(2) = 5
    End If
    .Refresh
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub msk_fltFim_Change()
On Error GoTo tratar_erro

ProcLimpaCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub msk_fltInicio_Change()
On Error GoTo tratar_erro

ProcLimpaCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt_PCP_Click()
On Error GoTo tratar_erro

ProcLimpaCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt_vendas_Click()
On Error GoTo tratar_erro

ProcLimpaCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optIgual_Click()
On Error GoTo tratar_erro

ProcLimpaCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optinicio_Click()
On Error GoTo tratar_erro

ProcLimpaCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optmeio_Click()
On Error GoTo tratar_erro

ProcLimpaCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optfim_Click()
On Error GoTo tratar_erro

ProcLimpaCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error GoTo tratar_erro

USToolBar1.ButtonState(2) = 5
ProcCorrigeTela

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCorrigeTela()
On Error GoTo tratar_erro

Select Case SSTab1.Tab
    Case 0:
        Label5.Visible = True
        Cmb_filtrar.Visible = True
        Frame1.Width = Frame2.Width
        txtTexto.Width = Cmb_filtrar.Left - (cmbfiltrarpor.Left + cmbfiltrarpor.Width)
    Case 1:
        Label5.Visible = False
        Cmb_filtrar.Visible = False
        Frame1.Width = Frame1.Width - Frame6.Width
        txtTexto.Width = (txtTexto.Width + Cmb_filtrar.Width) - Frame6.Width
End Select
Label20.Left = txtTexto.Left + (txtTexto.Width / 2) - (Label20.Width / 2)
cmbfamilia.Width = txtTexto.Width

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

Private Sub txtTexto_Change()
On Error GoTo tratar_erro

cmbfamilia.ListIndex = -1
ProcLimpaCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcFiltrar
    Case 2: ProcRM
    Case 3: ProcImprimir
    Case 5: 'ProcAjuda
    Case 6: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcRM()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

If SSTab1.Tab = 0 Then
    If Lista_necessidade.ListItems.Count = 0 Then Exit Sub
Else
    If Lista_detalhado.ListItems.Count = 0 Then Exit Sub
End If
Frm_necessidade_RM.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
