VERSION 5.00
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVendas_PI_composicao_prodserv 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Proposta comercial - Composição do"
   ClientHeight    =   10065
   ClientLeft      =   1890
   ClientTop       =   1365
   ClientWidth     =   10215
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmVendas_PI_composicao_prodserv.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10065
   ScaleWidth      =   10215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Txt_ID 
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
      Left            =   1020
      Locked          =   -1  'True
      MaxLength       =   255
      MouseIcon       =   "frmVendas_PI_composicao_prodserv.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   49
      Text            =   "0"
      Top             =   7650
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Frame Frame3 
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
      Height          =   1515
      Left            =   55
      TabIndex        =   41
      Top             =   1020
      Width           =   10125
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0E0&
         Height          =   510
         Left            =   5160
         TabIndex        =   50
         Top             =   210
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
            TabIndex        =   15
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
            TabIndex        =   13
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
            TabIndex        =   12
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
            TabIndex        =   14
            Top             =   180
            Width           =   1155
         End
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
         TabIndex        =   1
         ToolTipText     =   "Texto para pesquisa."
         Top             =   1050
         Width           =   7485
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
         ItemData        =   "frmVendas_PI_composicao_prodserv.frx":0614
         Left            =   180
         List            =   "frmVendas_PI_composicao_prodserv.frx":063F
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Opções para filtro."
         Top             =   390
         Width           =   4905
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
         ItemData        =   "frmVendas_PI_composicao_prodserv.frx":06F0
         Left            =   7680
         List            =   "frmVendas_PI_composicao_prodserv.frx":06FA
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
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
         TabIndex        =   2
         ToolTipText     =   "Texto para pesquisa."
         Top             =   1050
         Visible         =   0   'False
         Width           =   7485
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
         Left            =   2212
         TabIndex        =   44
         Top             =   180
         Width           =   840
      End
      Begin VB.Label Label14 
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
         Left            =   3187
         TabIndex        =   43
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
         TabIndex        =   42
         Top             =   840
         Width           =   1020
      End
   End
   Begin VB.Frame Frame9 
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
      Height          =   615
      Left            =   55
      TabIndex        =   37
      Top             =   4800
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
         TabIndex        =   6
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
         TabIndex        =   5
         Text            =   "30"
         ToolTipText     =   "Número de registros por página."
         Top             =   180
         Width           =   555
      End
      Begin DrawSuite2022.USButton cmdPagProx 
         Height          =   315
         Left            =   7200
         TabIndex        =   10
         ToolTipText     =   "Próxima página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmVendas_PI_composicao_prodserv.frx":0719
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
         Left            =   6660
         TabIndex        =   9
         ToolTipText     =   "Página anterior."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmVendas_PI_composicao_prodserv.frx":3EC0
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
         Left            =   5550
         TabIndex        =   7
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
         Left            =   6120
         TabIndex        =   8
         ToolTipText     =   "Primeira página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmVendas_PI_composicao_prodserv.frx":79CF
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
         Left            =   7740
         TabIndex        =   11
         ToolTipText     =   "Última página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmVendas_PI_composicao_prodserv.frx":BAC3
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
         Left            =   8520
         TabIndex        =   40
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
         TabIndex        =   39
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Carregar               registros por página"
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
         Left            =   1980
         TabIndex        =   38
         Top             =   240
         Width           =   2760
      End
   End
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   8850
      Top             =   240
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frmVendas_PI_composicao_prodserv.frx":F350
      Count           =   1
   End
   Begin VB.Frame Frame1 
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
      Height          =   1485
      Left            =   55
      TabIndex        =   27
      Top             =   5430
      Width           =   10125
      Begin VB.ComboBox Cmb_un 
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
         ItemData        =   "frmVendas_PI_composicao_prodserv.frx":1293B
         Left            =   4245
         List            =   "frmVendas_PI_composicao_prodserv.frx":1293D
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   18
         ToolTipText     =   "Unidade."
         Top             =   370
         Width           =   735
      End
      Begin VB.TextBox Txt_quantidade 
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
         Left            =   6060
         MaxLength       =   50
         TabIndex        =   21
         ToolTipText     =   "Quantidade."
         Top             =   990
         Width           =   1125
      End
      Begin VB.TextBox Txt_valor_uni 
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
         Left            =   7200
         MaxLength       =   50
         MousePointer    =   99  'Custom
         TabIndex        =   22
         ToolTipText     =   "Valor unitário."
         Top             =   990
         Width           =   1335
      End
      Begin VB.TextBox Txt_valor_total 
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
         Left            =   8550
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   23
         TabStop         =   0   'False
         ToolTipText     =   "Valor total."
         Top             =   990
         Width           =   1365
      End
      Begin VB.TextBox Txt_descricao 
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
         MaxLength       =   255
         TabIndex        =   20
         ToolTipText     =   "Descrição."
         Top             =   990
         Width           =   5865
      End
      Begin VB.ComboBox Cmb_cod_ref 
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
         ItemData        =   "frmVendas_PI_composicao_prodserv.frx":1293F
         Left            =   2235
         List            =   "frmVendas_PI_composicao_prodserv.frx":12941
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   17
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
         Left            =   4975
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "Família."
         Top             =   370
         Width           =   4955
      End
      Begin VB.TextBox Txt_codigo_int 
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
         Left            =   210
         MaxLength       =   255
         TabIndex        =   16
         ToolTipText     =   "Código interno."
         Top             =   370
         Width           =   2010
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Quant."
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
         Left            =   6345
         TabIndex        =   48
         Top             =   780
         Width           =   555
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vlr. unit."
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
         Left            =   7515
         TabIndex        =   47
         Top             =   780
         Width           =   690
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vlr. total"
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
         Left            =   8872
         TabIndex        =   46
         Top             =   780
         Width           =   720
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
         Left            =   2700
         TabIndex        =   35
         Top             =   780
         Width           =   825
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Código interno"
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
         Left            =   600
         TabIndex        =   34
         Top             =   180
         Width           =   1230
      End
      Begin VB.Label Label3 
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
         Left            =   4485
         TabIndex        =   33
         Top             =   180
         Width           =   255
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
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
         Left            =   7212
         TabIndex        =   32
         Top             =   180
         Width           =   480
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
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
         Left            =   2565
         TabIndex        =   31
         Top             =   180
         Width           =   1350
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   990
      Left            =   55
      TabIndex        =   36
      Top             =   0
      Width           =   10125
      _ExtentX        =   17859
      _ExtentY        =   1746
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
      ButtonCaption3  =   "Excluir"
      ButtonEnabled3  =   0   'False
      ButtonIconSize3 =   32
      ButtonToolTipText3=   "Excluir (F4)"
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
      ButtonWidth3    =   39
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
      ButtonLeft4     =   121
      ButtonTop4      =   4
      ButtonWidth4    =   2
      ButtonHeight4   =   55
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
      ButtonLeft5     =   125
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
      ButtonLeft6     =   163
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
      ButtonLeft7     =   191
      ButtonTop7      =   2
      ButtonWidth7    =   24
      ButtonHeight7   =   24
      ButtonUseMaskColor7=   0   'False
   End
   Begin VB.Frame Frame2 
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
      Height          =   645
      Left            =   55
      TabIndex        =   28
      Top             =   9390
      Width           =   10125
      Begin VB.TextBox Txt_valor_total_composicao 
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
         Left            =   8550
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   26
         TabStop         =   0   'False
         ToolTipText     =   "Valor total da composição."
         Top             =   180
         Width           =   1365
      End
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   5490
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   25
         TabStop         =   0   'False
         ToolTipText     =   "Valor total do produto."
         Top             =   180
         Width           =   1410
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vlr. total comp. :"
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
         Left            =   7080
         TabIndex        =   30
         Top             =   180
         Width           =   1365
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vlr. total prod. :"
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
         Left            =   4155
         TabIndex        =   29
         Top             =   180
         Width           =   1290
      End
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   55
      TabIndex        =   45
      Top             =   9135
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
      Height          =   2235
      Left            =   55
      TabIndex        =   4
      Top             =   2550
      Width           =   10125
      _ExtentX        =   17859
      _ExtentY        =   3942
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Text            =   "Cód."
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "T"
         Text            =   "Cód. interno"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Descrição"
         Object.Width           =   9675
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "Un."
         Object.Width           =   970
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Família"
         Object.Width           =   4410
      EndProperty
   End
   Begin MSComctlLib.ListView Lista_compos 
      Height          =   2205
      Left            =   55
      TabIndex        =   24
      Top             =   6930
      Width           =   10125
      _ExtentX        =   17859
      _ExtentY        =   3889
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
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Object.Width           =   529
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
         Object.Width           =   5618
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "Un."
         Object.Width           =   970
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Família"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Object.Tag             =   "N"
         Text            =   "Vlr. unit."
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Object.Tag             =   "N"
         Text            =   "Qtde."
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Object.Tag             =   "N"
         Text            =   "Vlr. total"
         Object.Width           =   1587
      EndProperty
   End
End
Attribute VB_Name = "frmVendas_PI_composicao_prodserv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StrSqlLocProdPadraoCompos As String 'OK

Private Sub ProcCarregaDadosProdServ()
On Error GoTo tratar_erro

Txt_codigo_int = IIf(IsNull(TBProduto!Desenho), "", TBProduto!Desenho)
Cmb_cod_ref.Refresh
With Cmb_un
    If IsNull(TBProduto!Unidade) = False And TBProduto!Unidade <> "" Then .Text = TBProduto!Unidade
    .Locked = True
    .TabStop = False
End With
Txt_familia = IIf(IsNull(TBProduto!Classe), "", TBProduto!Classe)
With Txt_descricao
    .Text = IIf(IsNull(TBProduto!Descricao), "", TBProduto!Descricao)
    .Locked = True
    .TabStop = False
End With
Txt_valor_uni = IIf(IsNull(TBProduto!PCusto), "", Format(TBProduto!PCusto, "###,##0.0000000000"))

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaDadosCompos()
On Error GoTo tratar_erro

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from vendas_carteira_composicao where ID_carteira = " & IDlista & " and Codigo_interno = '" & Txt_codigo_int & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    Txt_ID = TBAbrir!ID
    If IsNull(TBAbrir!Un) = False And TBAbrir!Un <> "" Then Cmb_un = TBAbrir!Un
    Txt_descricao = TBAbrir!Descricao
    Txt_valor_uni = IIf(IsNull(TBAbrir!valor_unitario), "", Format(TBAbrir!valor_unitario, "###,##0.0000000000"))
    Txt_quantidade = IIf(IsNull(TBAbrir!quantidade), "", Format(TBAbrir!quantidade, "###,##0.0000"))
    Txt_valor_total = IIf(IsNull(TBAbrir!Valor_total), "", Format(TBAbrir!Valor_total, "###,##0.00"))
End If
TBAbrir.Close
   
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_compos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    For InitFor = 1 To Lista_compos.ListItems.Count
        If Lista_compos.ListItems.Item(InitFor).Checked = True Then
            Lista_compos.ListItems.Item(InitFor).Checked = False
        Else
            With IIf(Vendas_Proposta = True, frmVendas_proposta, frmVendas_PI)
                If .txtDtValidacao <> "" Then GoTo Proximo
            End With
            Lista_compos.ListItems.Item(InitFor).Checked = True
Proximo:
        End If
    Next InitFor
Else
    ProcOrdenaListView Lista_compos, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_compos_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

For InitFor = 1 To Lista_compos.ListItems.Count
    If Lista_compos.ListItems.Item(InitFor).Checked = True And Lista_compos.ListItems.Item(InitFor) = Item Then
        With IIf(Vendas_Proposta = True, frmVendas_proposta, frmVendas_PI)
            If .txtDtValidacao <> "" Then
                USMsgBox ("Não é permitido excluir esta composição, pois " & IIf(Vendas_Proposta = True, "a proposta está validada", "o pedido está validado") & "."), vbExclamation, "CAPRIND v5.0"
                Lista_compos.ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
        End With
    End If
Next InitFor

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_compos_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With Lista_compos
    If .ListItems.Count = 0 Then Exit Sub
    Txt_codigo_int = .SelectedItem.ListSubItems(1)
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With ListView1
    If .ListItems.Count = 0 Then Exit Sub
    Txt_codigo_int = .SelectedItem.ListSubItems(1)
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_codigo_int_Change()
On Error GoTo tratar_erro

ProcLimpaCamposItem
If Txt_codigo_int = "" Then Exit Sub
ProcCarregaComboCodRef Cmb_cod_ref, "P.desenho = '" & Txt_codigo_int & "'", 0, "", False, True
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select Desenho, Unidade, Classe, Descricao, PCusto from projproduto where desenho = '" & Txt_codigo_int & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    ProcCarregaDadosProdServ
End If
TBProduto.Close

ProcCarregaDadosCompos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 10125, 7, True
If PI_Produtos = True Then
    TextoCaption = " produto"
Else
    TextoCaption = " serviço"
    Label4.Caption = "Vlr. total serv. :"
    Txt_valor_total_prod_serv.ToolTipText = "Valor total do produto."
End If
If Vendas_PI = True Then Caption = "Pedido interno - Composição do " & TextoCaption Else Caption = Caption & TextoCaption
With IIf(Vendas_PI = True, frmVendas_PI, frmVendas_proposta)
    If PI_Produtos = True Then
        Caption = Caption & " - Cód. int.: " & .txtNomenclatura & " - Descrição: " & .txtdesctecnica
        IDlista = .txtid_produto
        Txt_valor_total_prod_serv = .txtvalor_total
    Else
        Caption = Caption & " - Cód. int.: " & .txtcodservico & " - Descrição: " & .txtdescservico
        IDlista = .txtid_servico
        Txt_valor_total_prod_serv = .txtvlrtotalservico
    End If
End With
ProcCarregaComboUnidade Cmb_un, False
ProcCarregaComboFamilia cmbfamilia, "familia <> 'Null'", True
cmbfiltrarpor = "Código interno"
Cmb_ordenar = "Código interno"
If StrSqlLocProdPadraoCompos <> "" Then ProcCarregaLista
ProcCarregaListaCompos

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
With IIf(Vendas_Proposta = True, frmVendas_proposta, frmVendas_PI)
    Set TBGravar = CreateObject("adodb.recordset")
    TBGravar.Open "Select * from vendas_carteira_composicao where ID = " & Txt_ID, Conexao, adOpenKeyset, adLockOptimistic
    If TBGravar.EOF = True Then
        If FunVerifValidacaoRegistro("criar nova", .txtDtValidacao, IIf(Vendas_Proposta = True, "proposta", "pedido interno"), "composição", IIf(Vendas_Proposta = True, False, True)) = False Then Exit Sub
        If ProcVerifCampos = False Then Exit Sub
        
        TBGravar.AddNew
        TBGravar!ID_carteira = IDlista
        USMsgBox ("Nova composição cadatrada com sucesso."), vbInformation, "CAPRIND v5.0"
        Evento = "Novo"
    Else
        If FunVerifValidacaoRegistro("alterar", .txtDtValidacao, IIf(Vendas_Proposta = True, "proposta", "pedido interno"), "a composição", IIf(Vendas_Proposta = True, False, True)) = False Then Exit Sub
        If ProcVerifCampos = False Then Exit Sub
        
        USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
        Evento = "Alterar"
    End If
    TBGravar!Codigo_interno = Txt_codigo_int
    TBGravar!Un = Cmb_un
    TBGravar!Descricao = Txt_descricao
    TBGravar!quantidade = Txt_quantidade
    TBGravar!valor_unitario = Txt_valor_uni
    TBGravar!Valor_total = Txt_valor_total
    TBGravar.Update
    Txt_ID = TBGravar!ID
    TBGravar.Close
    '==================================
    ID_documento = Txt_ID
    Documento = IIf(Vendas_PI = True, "Nº pedido: ", "Nº proposta: ") & .txtCotacao & " - Rev.: " & .txtrevisao & " - Cód. interno: " & IIf(PI_Produtos = True, .txtNomenclatura, .txtcodservico)
    Documento1 = "Cód. interno composição: " & Txt_codigo_int
    ProcGravaEvento
    '==================================
    ProcCarregaListaCompos
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Function ProcVerifCampos() As Boolean
On Error GoTo tratar_erro

ProcVerifCampos = True
Acao = "salvar"
If Txt_codigo_int = "" Then
    NomeCampo = "o código interno"
    ProcVerificaAcao
    Txt_codigo_int.SetFocus
    ProcVerifCampos = False
    Exit Function
End If
If Cmb_un = "" Then
    NomeCampo = "a unidade"
    ProcVerificaAcao
    Cmb_un.SetFocus
    ProcVerifCampos = False
    Exit Function
End If
If Txt_descricao = "" Then
    NomeCampo = "a descrição"
    ProcVerificaAcao
    Txt_descricao.SetFocus
    ProcVerifCampos = False
    Exit Function
End If
valor = IIf(Txt_quantidade = "", 0, Txt_quantidade)
If valor <= 0 Then
    NomeCampo = "a quantidade"
    ProcVerificaAcao
    Txt_quantidade.SetFocus
    ProcVerifCampos = False
    Exit Function
End If
valor = IIf(Txt_valor_uni = "", 0, Txt_valor_uni)
If valor <= 0 Then
    NomeCampo = "o valor unitário"
    ProcVerificaAcao
    Txt_valor_uni.SetFocus
    ProcVerifCampos = False
    Exit Function
End If

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Private Sub ProcExcluir()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With Lista_compos
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir esta(s) compisição(ões)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            Conexao.Execute "DELETE from vendas_carteira_composicao where ID = " & .ListItems(InitFor)
            
            '==================================
            ID_documento = .ListItems(InitFor)
            With IIf(Vendas_PI = True, frmVendas_PI, frmVendas_proposta)
                Documento = IIf(Vendas_PI = True, "Nº pedido: ", "Nº proposta: ") & .txtCotacao & " - Rev.: " & .txtrevisao & " - Cód. interno: " & IIf(PI_Produtos = True, .txtNomenclatura, .txtcodservico)
            End With
            Documento1 = "Cód. interno composição: " & .ListItems(InitFor).ListSubItems(1)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe a(s) composição(ões) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Composição(ões) excluída(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    Txt_codigo_int = ""
    ProcLimpaCamposItem
    ProcCarregaListaCompos
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
    Case vbKeyF4: ProcExcluir
    Case vbKeyEscape: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCamposItem()
On Error GoTo tratar_erro

Txt_ID = 0
Cmb_cod_ref.ListIndex = -1
With Cmb_un
    .ListIndex = -1
    .Locked = False
    .TabStop = True
End With
Txt_familia = ""
With Txt_descricao
    .Text = ""
    .Locked = False
    .TabStop = True
End With
Txt_valor_uni = ""
Txt_quantidade = ""
Txt_valor_total = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_quantidade_Change()
On Error GoTo tratar_erro

If Txt_quantidade <> "" Then
    VerifNumero = Txt_quantidade
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_quantidade = ""
        Txt_quantidade.SetFocus
        Exit Sub
    End If
End If
ProcCalculaVlrTotal

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_quantidade_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus Txt_quantidade

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_quantidade_LostFocus()
On Error GoTo tratar_erro

Txt_quantidade = Format(Txt_quantidade, "###,##0.0000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_valor_uni_Change()
On Error GoTo tratar_erro

If Txt_valor_uni <> "" Then
    VerifNumero = Txt_valor_uni
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_valor_uni = ""
        Txt_valor_uni.SetFocus
        Exit Sub
    End If
End If
ProcCalculaVlrTotal

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_valor_uni_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus Txt_valor_uni

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_valor_uni_LostFocus()
On Error GoTo tratar_erro

Txt_valor_uni = Format(Txt_valor_uni, "###,##0.0000000000")

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
    Case 3: ProcExcluir
    'Case 5: ProcAjuda
    Case 6: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_ordenar_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbFamilia_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
If cmbfamilia <> "" Then txtTexto = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbfiltrarpor_Click()
On Error GoTo tratar_erro

With cmbfamilia
    ListView1.ListItems.Clear
    If cmbfiltrarpor = "Família" Or cmbfiltrarpor = "Cliente" Or cmbfiltrarpor = "Fornecedor" Then
        txtTexto.Visible = False
        .Visible = True
        .Clear
        .AddItem ""
        If cmbfiltrarpor = "Família" Then
            ProcCarregaComboFamilia cmbfamilia, "familia <> 'Null'", True
        Else
            If cmbfiltrarpor = "Cliente" Then ProcCarregaComboCliForn cmbfamilia, True Else ProcCarregaComboCliForn cmbfamilia, False
        End If
    Else
        txtTexto.Visible = True
        .Visible = False
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

TipoProduto = "P.Subtipoitem IS NOT NULL"
CamposFiltro = "P.codProduto, P.Desenho, P.Descricao, P.Unidade, P.classe"
INNERJOINTEXTO = "Select " & CamposFiltro & " from (((Projproduto P LEFT JOIN item_aplicacoes IA ON IA.codproduto = P.codproduto) LEFT JOIN Projproduto_clientes PC ON PC.codproduto = P.codproduto) LEFT JOIN Projproduto_fornecedor PF ON PF.codproduto = P.codproduto) LEFT JOIN Projproduto_fabricante PFAB ON PFAB.Codproduto = P.codproduto"
Ordenar = " group by " & CamposFiltro & " order by "
If Cmb_ordenar = "Código interno" Then Ordenar = Ordenar & "P.desenho" Else Ordenar = Ordenar & "P.Descricao"
TextoFiltroPadrao = TipoProduto & " and P.Bloqueado = 'False' " & Ordenar

If txtTexto.Visible = True And txtTexto <> "" Or cmbfamilia.Visible = True And cmbfamilia <> "" Then
    If cmbfiltrarpor = "Cliente" Then
            StrSqlLocProdPadraoCompos = INNERJOINTEXTO & " where PC.IDCliente = " & cmbfamilia.ItemData(cmbfamilia.ListIndex) & " and " & TextoFiltroPadrao
        ElseIf cmbfiltrarpor = "Fornecedor" Then
                StrSqlLocProdPadraoCompos = INNERJOINTEXTO & " where PF.IDfornecedor = " & cmbfamilia.ItemData(cmbfamilia.ListIndex) & " and " & TextoFiltroPadrao
            ElseIf cmbfiltrarpor = "Família" Then
                    StrSqlLocProdPadraoCompos = INNERJOINTEXTO & " where P.classe = '" & cmbfamilia & "' and " & TextoFiltroPadrao
                ElseIf cmbfiltrarpor = "Comprimento" Or cmbfiltrarpor = "Largura" Or cmbfiltrarpor = "Espessura" Then
                        Select Case cmbfiltrarpor
                            Case "Comprimento": TextoFiltro = "P.Comprimento"
                            Case "Largura": TextoFiltro = "P.Largura"
                            Case "Espessura": TextoFiltro = "P.Espessura"
                        End Select
                        valor = txtTexto
                        NovoValor = Replace(valor, ",", ".")
                        StrSqlLocProdPadraoCompos = INNERJOINTEXTO & " where " & TextoFiltro & " = " & NovoValor & " and " & TextoFiltroPadrao
                    Else
                        Select Case cmbfiltrarpor
                            Case "Código interno": TextoFiltro = "P.desenho"
                            Case "Código de referência": TextoFiltro = "IA.N_referencia"
                            Case "Descrição": TextoFiltro = "P.descricao"
                            Case "Descrição comercial": TextoFiltro = "P.Descricaotecnica"
                            Case "Número do desenho": TextoFiltro = "IA.desenho"
                            Case "Dureza": TextoFiltro = "P.Dureza"
                            Case "Part number": TextoFiltro = "PFAB.Part_number"
                        End Select
                        StrSqlLocProdPadraoCompos = INNERJOINTEXTO & " where " & TextoFiltro & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, Optinicio, txtTexto) & " and " & TextoFiltroPadrao
    End If
Else
    StrSqlLocProdPadraoCompos = INNERJOINTEXTO & " where " & TextoFiltroPadrao
End If
ProcCarregaLista

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
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
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagIr_Click()
On Error GoTo tratar_erro

If txtPagIr = "" Then Exit Sub
Quant = ReturnNumbersOnly(Right(lblPaginas.Caption, 4))
If Quant <= 1 Or txtPagIr > Quant Then Exit Sub
If txtPagIr.Text >= 1 And txtPagIr.Text <= Quant Then
    TBLocalizar_produto_padrao.AbsolutePage = txtPagIr.Text
    ProcExibePagina (TBLocalizar_produto_padrao.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLocalizar_produto_padrao.AbsolutePage = 1
ProcExibePagina (TBLocalizar_produto_padrao.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
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
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLocalizar_produto_padrao.AbsolutePage = TBLocalizar_produto_padrao.PageCount
ProcExibePagina (TBLocalizar_produto_padrao.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView ListView1, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaLista()
On Error GoTo tratar_erro

lblRegistros.Caption = "Nº de registros: 0"
lblPaginas.Caption = "Página: 0 de: 0"
ListView1.ListItems.Clear
If StrSqlLocProdPadraoCompos = "" Then Exit Sub
Set TBLocalizar_produto_padrao = CreateObject("adodb.recordset")
TBLocalizar_produto_padrao.Open StrSqlLocProdPadraoCompos, Conexao, adOpenKeyset, adLockReadOnly
If TBLocalizar_produto_padrao.EOF = False Then ProcExibePagina (1)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
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
PBLista.Max = FunVerifMaxPBListaPaginacao(TBLocalizar_produto_padrao.RecordCount - IIf(Pagina > 1, (TBLocalizar_produto_padrao.PageSize * (Pagina - 1)), 0), TBLocalizar_produto_padrao.PageSize)
PBLista.Value = 1
Contador = 0
Do While TBLocalizar_produto_padrao.EOF = False And (ContadorReg <= TamanhoPagina)
    With ListView1.ListItems
        .Add , , TBLocalizar_produto_padrao!Codproduto
        .Item(.Count).SubItems(1) = IIf(IsNull(TBLocalizar_produto_padrao!Desenho), "", TBLocalizar_produto_padrao!Desenho)
        .Item(.Count).SubItems(2) = IIf(IsNull(TBLocalizar_produto_padrao!Descricao), "", TBLocalizar_produto_padrao!Descricao)
        .Item(.Count).SubItems(3) = IIf(IsNull(TBLocalizar_produto_padrao!Unidade), "", TBLocalizar_produto_padrao!Unidade)
        .Item(.Count).SubItems(4) = IIf(IsNull(TBLocalizar_produto_padrao!Classe), "", TBLocalizar_produto_padrao!Classe)
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
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaListaCompos()
On Error GoTo tratar_erro

valor = 0
Lista_compos.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select VCC.*, P.Classe from vendas_carteira_composicao VCC LEFT JOIN Projproduto P ON P.Desenho = VCC.Codigo_interno where VCC.ID_carteira = " & IDlista & " order by VCC.Codigo_interno", Conexao, adOpenKeyset, adLockReadOnly
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBLISTA.EOF = False
        With Lista_compos.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = TBLISTA!Codigo_interno
            .Item(.Count).SubItems(2) = TBLISTA!Descricao
            .Item(.Count).SubItems(3) = TBLISTA!Un
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!Classe), "", TBLISTA!Classe)
            .Item(.Count).SubItems(5) = Format(TBLISTA!valor_unitario, "###,##0.0000000000")
            .Item(.Count).SubItems(6) = Format(TBLISTA!quantidade, "###,##0.0000")
            .Item(.Count).SubItems(7) = Format(TBLISTA!Valor_total, "###,##0.00")
            valor = valor + TBLISTA!Valor_total
        End With
        TBLISTA.MoveNext
        Contador = Contador + 1
        PBLista.Value = Contador
    Loop
End If
TBLISTA.Close
Txt_valor_total_composicao = Format(valor, "###,##0.00")
 
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optfim_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optinicio_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optmeio_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear

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

ListView1.ListItems.Clear
If txtTexto <> "" Then cmbfamilia.ListIndex = -1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCalculaVlrTotal()
On Error GoTo tratar_erro

If Txt_valor_uni = "" Or Txt_quantidade = "" Then
    Txt_valor_total = ""
Else
    valor = Txt_valor_uni
    Valor1 = Txt_quantidade
    Txt_valor_total = Format(valor * Valor1, "###,##0.00")
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
