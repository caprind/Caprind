VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCompras_reqcot_abrir 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Administrativo - Compras - Cotação - Localizar produtos/serviços"
   ClientHeight    =   9060
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   14385
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
   ForeColor       =   &H8000000D&
   Icon            =   "frmCompras_Reqcot_abrir.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9060
   ScaleWidth      =   14385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   75
      TabIndex        =   69
      Top             =   8760
      Width           =   14325
      _ExtentX        =   25268
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
   Begin MSComctlLib.ListView Lista_Ncadastrados 
      Height          =   5655
      Left            =   60
      TabIndex        =   67
      Top             =   3090
      Visible         =   0   'False
      Width           =   14325
      _ExtentX        =   25268
      _ExtentY        =   9975
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
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
         Object.Width           =   13291
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "Un. est."
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Un. com."
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Tag             =   "T"
         Text            =   "Família"
         Object.Width           =   6174
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9120
      Left            =   0
      TabIndex        =   68
      Top             =   0
      Width           =   14445
      _ExtentX        =   25479
      _ExtentY        =   16087
      _Version        =   393216
      Tabs            =   4
      Tab             =   2
      TabsPerRow      =   4
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
      TabCaption(0)   =   "Necessidade"
      TabPicture(0)   =   "frmCompras_Reqcot_abrir.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1(0)"
      Tab(0).Control(1)=   "Frame3"
      Tab(0).Control(2)=   "Frame1(23)"
      Tab(0).Control(3)=   "USToolBar4"
      Tab(0).Control(4)=   "ListaNecessidade"
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Solicitação"
      TabPicture(1)   =   "frmCompras_Reqcot_abrir.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Lista_solicitados"
      Tab(1).Control(1)=   "USToolBar1"
      Tab(1).Control(2)=   "USImageList1"
      Tab(1).Control(3)=   "Frame1(2)"
      Tab(1).Control(4)=   "Frame1(1)"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Cadastrados"
      TabPicture(2)   =   "frmCompras_Reqcot_abrir.frx":047A
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frame5"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Lista_cadastrados"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "USToolBar2"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Frame6"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "USImageList2"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Frame1(4)"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "Não cadastrados"
      TabPicture(3)   =   "frmCompras_Reqcot_abrir.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "USToolBar3"
      Tab(3).Control(1)=   "Frame10"
      Tab(3).Control(2)=   "txtidlista"
      Tab(3).ControlCount=   3
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   615
         Index           =   4
         Left            =   60
         TabIndex        =   112
         Top             =   8130
         Width           =   14325
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
            Index           =   2
            Left            =   8220
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
            Index           =   2
            Left            =   3930
            TabIndex        =   53
            Text            =   "30"
            ToolTipText     =   "Número de registros por página."
            Top             =   180
            Width           =   555
         End
         Begin DrawSuite2022.USButton cmdPagProx 
            Height          =   315
            Index           =   2
            Left            =   10440
            TabIndex        =   58
            ToolTipText     =   "Próxima página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmCompras_Reqcot_abrir.frx":04B2
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
            Index           =   2
            Left            =   9900
            TabIndex        =   57
            ToolTipText     =   "Página anterior."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmCompras_Reqcot_abrir.frx":3C59
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
            Index           =   2
            Left            =   8790
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
            Index           =   2
            Left            =   9360
            TabIndex        =   56
            ToolTipText     =   "Primeira página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmCompras_Reqcot_abrir.frx":7768
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
            Index           =   2
            Left            =   10980
            TabIndex        =   59
            ToolTipText     =   "Última página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmCompras_Reqcot_abrir.frx":B859
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
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   2
            Left            =   12090
            TabIndex        =   115
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lblRegistros 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nº de registros: 0"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   2
            Left            =   180
            TabIndex        =   114
            Top             =   240
            Width           =   1275
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Carregar               registros por página"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   3240
            TabIndex        =   113
            Top             =   240
            Width           =   2760
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   615
         Index           =   1
         Left            =   -74940
         TabIndex        =   108
         Top             =   8130
         Width           =   14325
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
            Index           =   1
            Left            =   3930
            TabIndex        =   27
            Text            =   "30"
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
            Index           =   1
            Left            =   8220
            TabIndex        =   28
            ToolTipText     =   "Número da página."
            Top             =   180
            Width           =   555
         End
         Begin DrawSuite2022.USButton cmdPagProx 
            Height          =   315
            Index           =   1
            Left            =   10440
            TabIndex        =   32
            ToolTipText     =   "Próxima página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmCompras_Reqcot_abrir.frx":F0E7
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
            Index           =   1
            Left            =   9900
            TabIndex        =   31
            ToolTipText     =   "Página anterior."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmCompras_Reqcot_abrir.frx":1288E
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
            Index           =   1
            Left            =   8790
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
         Begin DrawSuite2022.USButton cmdPagPrim 
            Height          =   315
            Index           =   1
            Left            =   9360
            TabIndex        =   30
            ToolTipText     =   "Primeira página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmCompras_Reqcot_abrir.frx":1639D
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
            Index           =   1
            Left            =   10980
            TabIndex        =   33
            ToolTipText     =   "Última página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmCompras_Reqcot_abrir.frx":1A48F
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
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Carregar               registros por página"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   3240
            TabIndex        =   111
            Top             =   240
            Width           =   2760
         End
         Begin VB.Label lblRegistros 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nº de registros: 0"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   1
            Left            =   180
            TabIndex        =   110
            Top             =   240
            Width           =   1275
         End
         Begin VB.Label lblPaginas 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Página: 0 de: 0"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   1
            Left            =   12090
            TabIndex        =   109
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   615
         Index           =   0
         Left            =   -74940
         TabIndex        =   104
         Top             =   8130
         Width           =   14325
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
            Index           =   0
            Left            =   8220
            TabIndex        =   12
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
            Index           =   0
            Left            =   3930
            TabIndex        =   11
            Text            =   "30"
            ToolTipText     =   "Número de registros por página."
            Top             =   180
            Width           =   555
         End
         Begin DrawSuite2022.USButton cmdPagProx 
            Height          =   315
            Index           =   0
            Left            =   10440
            TabIndex        =   16
            ToolTipText     =   "Próxima página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmCompras_Reqcot_abrir.frx":1DD1D
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
            Index           =   0
            Left            =   9900
            TabIndex        =   15
            ToolTipText     =   "Página anterior."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmCompras_Reqcot_abrir.frx":214C4
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
            Index           =   0
            Left            =   8790
            TabIndex        =   13
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
            Index           =   0
            Left            =   9360
            TabIndex        =   14
            ToolTipText     =   "Primeira página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmCompras_Reqcot_abrir.frx":24FD3
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
            Index           =   0
            Left            =   10980
            TabIndex        =   17
            ToolTipText     =   "Última página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmCompras_Reqcot_abrir.frx":290C4
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
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   0
            Left            =   12090
            TabIndex        =   107
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lblRegistros 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nº de registros: 0"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   106
            Top             =   240
            Width           =   1275
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Carregar               registros por página"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   3240
            TabIndex        =   105
            Top             =   240
            Width           =   2760
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Height          =   855
         Index           =   2
         Left            =   -74940
         TabIndex        =   101
         Top             =   1320
         Width           =   14325
         Begin VB.Frame Frame2 
            BackColor       =   &H00E0E0E0&
            Height          =   510
            Left            =   2910
            TabIndex        =   117
            Top             =   210
            Width           =   4785
            Begin VB.OptionButton optIgual_Sol 
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
               TabIndex        =   22
               Top             =   180
               Width           =   705
            End
            Begin VB.OptionButton optMeio_Sol 
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
               TabIndex        =   20
               Top             =   180
               Width           =   1275
            End
            Begin VB.OptionButton Optinicio_Sol 
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
               TabIndex        =   19
               Top             =   180
               Value           =   -1  'True
               Width           =   1275
            End
            Begin VB.OptionButton optFim_sol 
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
               TabIndex        =   21
               Top             =   180
               Width           =   1155
            End
         End
         Begin VB.ComboBox cmbfiltrarpor_sol 
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
            ItemData        =   "frmCompras_Reqcot_abrir.frx":2C952
            Left            =   180
            List            =   "frmCompras_Reqcot_abrir.frx":2C96E
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   18
            ToolTipText     =   "Opções para filtro."
            Top             =   390
            Width           =   2655
         End
         Begin VB.TextBox txtTexto_sol 
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
            Left            =   7800
            TabIndex        =   23
            ToolTipText     =   "Texto para pesquisa."
            Top             =   390
            Width           =   6315
         End
         Begin MSComCtl2.DTPicker Txtprazo_sol 
            Height          =   315
            Left            =   7800
            TabIndex        =   25
            ToolTipText     =   "Texto para pesquisa."
            Top             =   390
            Visible         =   0   'False
            Width           =   6315
            _ExtentX        =   11139
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
            Format          =   197197825
            CurrentDate     =   39057
         End
         Begin VB.ComboBox cmbTexto_sol 
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
            ItemData        =   "frmCompras_Reqcot_abrir.frx":2C9DD
            Left            =   7800
            List            =   "frmCompras_Reqcot_abrir.frx":2C9DF
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   24
            ToolTipText     =   "Familia."
            Top             =   390
            Visible         =   0   'False
            Width           =   6315
         End
         Begin VB.Label Label1 
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
            Index           =   3
            Left            =   1087
            TabIndex        =   103
            Top             =   180
            Width           =   840
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Texto para pesquisa"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   1
            Left            =   10215
            TabIndex        =   102
            Top             =   180
            Width           =   1485
         End
      End
      Begin VB.Frame Frame3 
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
         Left            =   -74940
         TabIndex        =   100
         Top             =   1320
         Width           =   14325
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
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Height          =   855
         Index           =   23
         Left            =   -74940
         TabIndex        =   95
         Top             =   1830
         Width           =   14325
         Begin VB.Frame Frame4 
            BackColor       =   &H00E0E0E0&
            Height          =   510
            Left            =   2580
            TabIndex        =   116
            Top             =   210
            Width           =   4785
            Begin VB.OptionButton Optfim_necess 
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
               TabIndex        =   5
               Top             =   180
               Width           =   1155
            End
            Begin VB.OptionButton Optinicio_necess 
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
               TabIndex        =   3
               Top             =   180
               Value           =   -1  'True
               Width           =   1275
            End
            Begin VB.OptionButton Optmeio_necess 
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
               TabIndex        =   4
               Top             =   180
               Width           =   1275
            End
            Begin VB.OptionButton optIgual_necess 
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
               TabIndex        =   6
               Top             =   180
               Width           =   705
            End
         End
         Begin VB.TextBox txtTexto_necess 
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
            Left            =   7440
            TabIndex        =   7
            ToolTipText     =   "Texto para pesquisa."
            Top             =   390
            Width           =   4395
         End
         Begin VB.ComboBox cmbfiltrarpor_necess 
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
            ItemData        =   "frmCompras_Reqcot_abrir.frx":2C9E1
            Left            =   180
            List            =   "frmCompras_Reqcot_abrir.frx":2C9FD
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   2
            ToolTipText     =   "Opções para filtro."
            Top             =   390
            Width           =   2325
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
            ItemData        =   "frmCompras_Reqcot_abrir.frx":2CA75
            Left            =   11850
            List            =   "frmCompras_Reqcot_abrir.frx":2CA7F
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   9
            ToolTipText     =   "Tipo de necessidade."
            Top             =   390
            Width           =   2295
         End
         Begin VB.ComboBox cmbTexto_necess 
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
            ItemData        =   "frmCompras_Reqcot_abrir.frx":2CAA9
            Left            =   7440
            List            =   "frmCompras_Reqcot_abrir.frx":2CAAB
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   8
            ToolTipText     =   "Texto para pesquisa."
            Top             =   390
            Visible         =   0   'False
            Width           =   4395
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Texto para pesquisa"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   2
            Left            =   8895
            TabIndex        =   98
            Top             =   180
            Width           =   1485
         End
         Begin VB.Label Label1 
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
            Index           =   24
            Left            =   922
            TabIndex        =   97
            Top             =   180
            Width           =   840
         End
         Begin VB.Label Label4 
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
            Left            =   12135
            TabIndex        =   96
            Top             =   180
            Width           =   1710
         End
      End
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   -68430
         Top             =   540
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmCompras_Reqcot_abrir.frx":2CAAD
         Count           =   1
      End
      Begin DrawSuite2022.USImageList USImageList2 
         Left            =   6780
         Top             =   600
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmCompras_Reqcot_abrir.frx":2F837
         Count           =   1
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00E0E0E0&
         Height          =   855
         Left            =   60
         TabIndex        =   80
         Top             =   1320
         Width           =   14325
         Begin VB.Frame Frame7 
            BackColor       =   &H00E0E0E0&
            Height          =   510
            Left            =   2520
            TabIndex        =   118
            Top             =   210
            Width           =   4785
            Begin VB.OptionButton optFim_cad 
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
               TabIndex        =   37
               Top             =   180
               Width           =   1155
            End
            Begin VB.OptionButton optInicio_cad 
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
               TabIndex        =   35
               Top             =   180
               Value           =   -1  'True
               Width           =   1275
            End
            Begin VB.OptionButton optMeio_cad 
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
               TabIndex        =   36
               Top             =   180
               Width           =   1275
            End
            Begin VB.OptionButton optIgual_cad 
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
               TabIndex        =   38
               Top             =   180
               Width           =   705
            End
         End
         Begin VB.TextBox txtTexto_cad 
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
            Left            =   7410
            TabIndex        =   39
            ToolTipText     =   "Texto para pesquisa."
            Top             =   390
            Width           =   6735
         End
         Begin VB.ComboBox cmbfiltrarpor_cad 
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
            ItemData        =   "frmCompras_Reqcot_abrir.frx":325BA
            Left            =   180
            List            =   "frmCompras_Reqcot_abrir.frx":325D0
            Style           =   2  'Dropdown List
            TabIndex        =   34
            ToolTipText     =   "Opções para filtro."
            Top             =   390
            Width           =   2265
         End
         Begin VB.ComboBox cmbTexto_cad 
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
            Left            =   7410
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   40
            ToolTipText     =   "Texto para pesquisa."
            Top             =   390
            Visible         =   0   'False
            Width           =   6735
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Texto para pesquisa"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   0
            Left            =   10042
            TabIndex        =   82
            Top             =   180
            Width           =   1470
         End
         Begin VB.Label Label5 
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
            Left            =   892
            TabIndex        =   81
            Top             =   180
            Width           =   840
         End
      End
      Begin VB.TextBox txtidlista 
         Alignment       =   2  'Center
         BackColor       =   &H80000014&
         Enabled         =   0   'False
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
         Left            =   -73620
         MaxLength       =   50
         MouseIcon       =   "frmCompras_Reqcot_abrir.frx":32630
         MousePointer    =   99  'Custom
         TabIndex        =   79
         Text            =   "0"
         Top             =   2370
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   1755
         Left            =   -74940
         TabIndex        =   70
         Top             =   1320
         Width           =   14325
         Begin VB.TextBox txtCodproduto2 
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
            MaxLength       =   50
            TabIndex        =   60
            ToolTipText     =   "Código interno."
            Top             =   390
            Width           =   1665
         End
         Begin VB.TextBox txtQtde2 
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
            Left            =   12925
            MaxLength       =   50
            TabIndex        =   65
            ToolTipText     =   "Quantidade."
            Top             =   390
            Width           =   1220
         End
         Begin VB.ComboBox cmbFamilia3 
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
            Left            =   6900
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   62
            ToolTipText     =   "Família."
            Top             =   390
            Width           =   4350
         End
         Begin VB.ComboBox cmbun 
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
            ItemData        =   "frmCompras_Reqcot_abrir.frx":3293A
            Left            =   11265
            List            =   "frmCompras_Reqcot_abrir.frx":3293C
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   63
            ToolTipText     =   "Unidade de estoque."
            Top             =   390
            Width           =   825
         End
         Begin VB.TextBox txtDescricao_comercial2 
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
            MaxLength       =   5000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   66
            ToolTipText     =   "Descrição comercial."
            Top             =   990
            Width           =   13965
         End
         Begin VB.ComboBox Cmb_un_com 
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
            ItemData        =   "frmCompras_Reqcot_abrir.frx":3293E
            Left            =   12105
            List            =   "frmCompras_Reqcot_abrir.frx":32940
            Locked          =   -1  'True
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   64
            TabStop         =   0   'False
            ToolTipText     =   "Unidade comercial."
            Top             =   390
            Width           =   825
         End
         Begin VB.TextBox txtDescricao2 
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
            Left            =   1860
            MaxLength       =   255
            TabIndex        =   61
            ToolTipText     =   "Descrição."
            Top             =   390
            Width           =   5040
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Qtde."
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
            Index           =   0
            Left            =   13310
            TabIndex        =   77
            Top             =   180
            Width           =   450
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Un. est."
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   4
            Left            =   11385
            TabIndex        =   76
            Top             =   180
            Width           =   585
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cod. interno"
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
            Index           =   5
            Left            =   502
            TabIndex        =   75
            Top             =   180
            Width           =   1020
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
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
            Index           =   6
            Left            =   3968
            TabIndex        =   74
            Top             =   180
            Width           =   825
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Família"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   7
            Left            =   8835
            TabIndex        =   73
            Top             =   180
            Width           =   480
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descrição comercial"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   9
            Left            =   6465
            TabIndex        =   72
            Top             =   780
            Width           =   1395
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Un. com."
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   11
            Left            =   12195
            TabIndex        =   71
            Top             =   180
            Width           =   645
         End
      End
      Begin DrawSuite2022.USToolBar USToolBar3 
         Height          =   975
         Left            =   -74940
         TabIndex        =   78
         Top             =   330
         Width           =   14325
         _ExtentX        =   25268
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft2     =   40
         ButtonTop2      =   2
         ButtonWidth2    =   44
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft3     =   86
         ButtonTop3      =   2
         ButtonWidth3    =   45
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
         ButtonLeft4     =   133
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft5     =   137
         ButtonTop5      =   2
         ButtonWidth5    =   41
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft6     =   180
         ButtonTop6      =   2
         ButtonWidth6    =   30
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
         ButtonLeft7     =   212
         ButtonTop7      =   2
         ButtonWidth7    =   24
         ButtonHeight7   =   24
         Begin DrawSuite2022.USImageList USImageList3 
            Left            =   5730
            Top             =   210
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmCompras_Reqcot_abrir.frx":32942
            Count           =   1
         End
      End
      Begin DrawSuite2022.USToolBar USToolBar2 
         Height          =   975
         Left            =   60
         TabIndex        =   83
         Top             =   330
         Width           =   14325
         _ExtentX        =   25268
         _ExtentY        =   1720
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
         ButtonCaption2  =   "Adicionar"
         ButtonEnabled2  =   0   'False
         ButtonIconSize2 =   32
         ButtonToolTipText2=   "Adicionar produto na lista (F3)"
         ButtonKey2      =   "2"
         ButtonAlignment2=   2
         BeginProperty ButtonFont2 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft2     =   40
         ButtonTop2      =   2
         ButtonWidth2    =   61
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
         ButtonLeft3     =   103
         ButtonTop3      =   4
         ButtonWidth3    =   2
         ButtonHeight3   =   54
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft4     =   107
         ButtonTop4      =   2
         ButtonWidth4    =   41
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft5     =   150
         ButtonTop5      =   2
         ButtonWidth5    =   30
         ButtonHeight5   =   21
         ButtonUseMaskColor5=   0   'False
         ButtonEnabled6  =   0   'False
         ButtonIconSize6 =   32
         ButtonKey6      =   "6"
         ButtonAlignment6=   2
         BeginProperty ButtonFont6 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState6    =   5
         ButtonLeft6     =   182
         ButtonTop6      =   2
         ButtonWidth6    =   24
         ButtonHeight6   =   24
      End
      Begin MSComctlLib.ListView Lista_cadastrados 
         Height          =   4185
         Left            =   75
         TabIndex        =   52
         Top             =   3930
         Width           =   14325
         _ExtentX        =   25268
         _ExtentY        =   7382
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
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
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "N"
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
            Object.Width           =   10998
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Object.Tag             =   "T"
            Text            =   "Un. est."
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "Un. com."
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Object.Tag             =   "T"
            Text            =   "Família"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Object.Tag             =   "N"
            Text            =   "Qtde. estoque"
            Object.Width           =   2293
         EndProperty
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   1725
         Left            =   60
         TabIndex        =   84
         Top             =   2190
         Width           =   14325
         Begin VB.TextBox txtqtde_PC 
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
            Left            =   570
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   119
            TabStop         =   0   'False
            ToolTipText     =   "Quantidade em peça."
            Top             =   1110
            Visible         =   0   'False
            Width           =   945
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
            Left            =   8070
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   44
            TabStop         =   0   'False
            ToolTipText     =   "Unidade de estoque."
            Top             =   390
            Width           =   735
         End
         Begin VB.TextBox txtqtde 
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
            Left            =   9570
            MaxLength       =   50
            TabIndex        =   46
            ToolTipText     =   "Quantidade comercial."
            Top             =   390
            Width           =   945
         End
         Begin VB.CommandButton cmdcalc_peso 
            BackColor       =   &H00C0C0C0&
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
            Left            =   10500
            Picture         =   "frmCompras_Reqcot_abrir.frx":35D1A
            Style           =   1  'Graphical
            TabIndex        =   47
            ToolTipText     =   "Abrir calculadora para cálculo de peso."
            Top             =   390
            Width           =   315
         End
         Begin VB.TextBox txtcodproduto 
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
            TabIndex        =   41
            TabStop         =   0   'False
            ToolTipText     =   "Código interno."
            Top             =   390
            Width           =   1665
         End
         Begin VB.ComboBox Cmb_OS 
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
            ItemData        =   "frmCompras_Reqcot_abrir.frx":35F83
            Left            =   12930
            List            =   "frmCompras_Reqcot_abrir.frx":35F85
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   50
            ToolTipText     =   "Número da OS."
            Top             =   390
            Width           =   1215
         End
         Begin VB.TextBox txtOrdem 
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
            Left            =   11880
            TabIndex        =   49
            ToolTipText     =   "Número da ordem."
            Top             =   390
            Width           =   1035
         End
         Begin VB.TextBox txtDescricao_comercial 
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
            MaxLength       =   5000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   51
            ToolTipText     =   "Descrição comercial."
            Top             =   990
            Width           =   13965
         End
         Begin VB.TextBox Txt_un_com 
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
            Left            =   8820
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   45
            TabStop         =   0   'False
            ToolTipText     =   "Unidade comercial."
            Top             =   390
            Width           =   735
         End
         Begin VB.TextBox txtqtde_est 
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
            Left            =   10920
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   48
            TabStop         =   0   'False
            ToolTipText     =   "Quantidade da unidade de estoque."
            Top             =   390
            Width           =   945
         End
         Begin VB.CommandButton Cmd_visualizar_arquivo 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   1860
            Picture         =   "frmCompras_Reqcot_abrir.frx":35F87
            Style           =   1  'Graphical
            TabIndex        =   42
            ToolTipText     =   "Visualizar arquivo."
            Top             =   390
            Width           =   315
         End
         Begin VB.TextBox txtdesc 
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
            Left            =   2280
            Locked          =   -1  'True
            MaxLength       =   255
            TabIndex        =   43
            TabStop         =   0   'False
            ToolTipText     =   "Descrição."
            Top             =   390
            Width           =   5775
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
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
            Index           =   31
            Left            =   4755
            TabIndex        =   93
            Top             =   180
            Width           =   825
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cod. interno"
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
            Index           =   1
            Left            =   502
            TabIndex        =   92
            Top             =   180
            Width           =   1020
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Un. est."
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   2
            Left            =   8145
            TabIndex        =   91
            Top             =   180
            Width           =   585
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Qtde. com."
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
            Index           =   3
            Left            =   9592
            TabIndex        =   90
            Top             =   180
            Width           =   900
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "OS"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   13425
            TabIndex        =   89
            Top             =   180
            Width           =   210
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Ordem"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   12150
            TabIndex        =   88
            Top             =   180
            Width           =   480
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descrição comercial"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   8
            Left            =   6465
            TabIndex        =   87
            Top             =   780
            Width           =   1395
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Un. com."
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   10
            Left            =   8865
            TabIndex        =   86
            Top             =   180
            Width           =   645
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Qt. un. est."
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
            Index           =   12
            Left            =   10942
            TabIndex        =   85
            Top             =   180
            Width           =   900
         End
      End
      Begin DrawSuite2022.USToolBar USToolBar1 
         Height          =   975
         Left            =   -74940
         TabIndex        =   94
         Top             =   330
         Width           =   14325
         _ExtentX        =   25268
         _ExtentY        =   1720
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft1     =   2
         ButtonTop1      =   2
         ButtonWidth1    =   42
         ButtonHeight1   =   21
         ButtonUseMaskColor1=   0   'False
         ButtonCaption2  =   "Adicionar"
         ButtonEnabled2  =   0   'False
         ButtonIconSize2 =   32
         ButtonToolTipText2=   "Cotar todos os produtos listados (F3)"
         ButtonKey2      =   "2"
         ButtonAlignment2=   2
         BeginProperty ButtonFont2 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft2     =   46
         ButtonTop2      =   2
         ButtonWidth2    =   61
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
         ButtonLeft3     =   109
         ButtonTop3      =   4
         ButtonWidth3    =   2
         ButtonHeight3   =   54
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft4     =   113
         ButtonTop4      =   2
         ButtonWidth4    =   41
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft5     =   156
         ButtonTop5      =   2
         ButtonWidth5    =   30
         ButtonHeight5   =   21
         ButtonUseMaskColor5=   0   'False
         ButtonEnabled6  =   0   'False
         ButtonIconSize6 =   32
         ButtonKey6      =   "6"
         ButtonAlignment6=   2
         BeginProperty ButtonFont6 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState6    =   5
         ButtonLeft6     =   188
         ButtonTop6      =   2
         ButtonWidth6    =   24
         ButtonHeight6   =   24
         ButtonUseMaskColor6=   0   'False
      End
      Begin DrawSuite2022.USToolBar USToolBar4 
         Height          =   975
         Left            =   -74940
         TabIndex        =   99
         Top             =   330
         Width           =   14325
         _ExtentX        =   25268
         _ExtentY        =   1720
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft1     =   2
         ButtonTop1      =   2
         ButtonWidth1    =   42
         ButtonHeight1   =   21
         ButtonUseMaskColor1=   0   'False
         ButtonCaption2  =   "Adicionar"
         ButtonEnabled2  =   0   'False
         ButtonIconSize2 =   32
         ButtonToolTipText2=   "Adicionar selecionados (F3)"
         ButtonKey2      =   "2"
         ButtonAlignment2=   2
         BeginProperty ButtonFont2 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft2     =   46
         ButtonTop2      =   2
         ButtonWidth2    =   61
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
         ButtonLeft3     =   109
         ButtonTop3      =   4
         ButtonWidth3    =   2
         ButtonHeight3   =   54
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft4     =   113
         ButtonTop4      =   2
         ButtonWidth4    =   41
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft5     =   156
         ButtonTop5      =   2
         ButtonWidth5    =   30
         ButtonHeight5   =   21
         ButtonUseMaskColor5=   0   'False
         ButtonEnabled6  =   0   'False
         ButtonIconSize6 =   32
         ButtonKey6      =   "6"
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
         ButtonLeft6     =   188
         ButtonTop6      =   2
         ButtonWidth6    =   24
         ButtonHeight6   =   24
         ButtonUseMaskColor6=   0   'False
         Begin DrawSuite2022.USImageList USImageList4 
            Left            =   5520
            Top             =   210
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmCompras_Reqcot_abrir.frx":36549
            Count           =   1
         End
      End
      Begin MSComctlLib.ListView ListaNecessidade 
         Height          =   5415
         Left            =   -74940
         TabIndex        =   10
         Top             =   2700
         Width           =   14325
         _ExtentX        =   25268
         _ExtentY        =   9551
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "N"
            Object.Width           =   529
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
            Object.Width           =   16730
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Object.Tag             =   "T"
            Text            =   "Un."
            Object.Width           =   970
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Object.Tag             =   "N"
            Text            =   "Necessidade"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Object.Tag             =   "N"
            Text            =   "Necess. PÇ"
            Object.Width           =   2117
         EndProperty
      End
      Begin MSComctlLib.ListView Lista_solicitados 
         Height          =   5925
         Left            =   -74940
         TabIndex        =   26
         Top             =   2190
         Width           =   14325
         _ExtentX        =   25268
         _ExtentY        =   10451
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
         NumItems        =   12
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "N"
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Tag             =   "T"
            Text            =   "Status"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Nº solicitação"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Tag             =   "T"
            Text            =   "Cód. int."
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "Descrição"
            Object.Width           =   5369
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Object.Tag             =   "T"
            Text            =   "Un. est."
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Object.Tag             =   "T"
            Text            =   "Un. com."
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Object.Tag             =   "N"
            Text            =   "Quant. est."
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Object.Tag             =   "N"
            Text            =   "Quant. com."
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Object.Tag             =   "T"
            Text            =   "Detalhe"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   10
            Object.Tag             =   "D"
            Text            =   "Prazo entr."
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Object.Tag             =   "T"
            Text            =   "Obs."
            Object.Width           =   0
         EndProperty
      End
   End
End
Attribute VB_Name = "frmCompras_reqcot_abrir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StrSql_Cotacao_Localizar_Necessidade As String 'OK
Dim StrSql_Cotacao_Localizar_Solicitacao As String 'OK
Dim StrSql_Cotacao_Localizar_Cadastrados As String 'OK
Dim Novo_compras_reqcot_item As Boolean 'ok

Private Sub Cmb_filtrar_Click()
On Error GoTo tratar_erro

ListaNecessidade.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbTexto_cad_Click()
On Error GoTo tratar_erro

Lista_cadastrados.ListItems.Clear
ProcLimpaCampos
If cmbTexto_cad <> "" Then txtTexto_cad = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbfiltrarpor_cad_Click()
On Error GoTo tratar_erro

Lista_cadastrados.ListItems.Clear
ProcLimpaCampos
If cmbfiltrarpor_cad = "Família" Then
    txtTexto_cad.Visible = False
    cmbTexto_cad.Visible = True
    ProcCarregaComboFamilia cmbTexto_cad, "familia <> 'Null' and (Compras = 'True' or Fabricacao = 'True')", True
Else
    txtTexto_cad.Visible = True
    cmbTexto_cad.Visible = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbfiltrarpor_necess_Click()
On Error GoTo tratar_erro

ListaNecessidade.ListItems.Clear
If cmbfiltrarpor_necess = "Família" Then
    txtTexto_necess.Visible = False
    cmbTexto_necess.Visible = True
    ProcCarregaComboFamilia cmbTexto_necess, "familia <> 'Null' and Compras = 'True'", True
Else
    txtTexto_necess.Visible = True
    cmbTexto_necess.Visible = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procAdicionar_Cadastrados()
On Error GoTo tratar_erro

Acao = "adicionar na cotação"
If txtcodproduto.Text = "" Then
    NomeCampo = "o produto na lista"
    ProcVerificaAcao
    Exit Sub
End If
valor = IIf(txtQtde = "", 0, txtQtde)
If valor <= 0 Then
    NomeCampo = "a quantidade"
    ProcVerificaAcao
    txtQtde.SetFocus
    Exit Sub
End If
If Cmb_OS.Locked = False And Cmb_OS = "" Then
    NomeCampo = "o número da OS"
    ProcVerificaAcao
    Cmb_OS.SetFocus
    Exit Sub
End If

With frmcompras_reqcot
    Cont = .txtidcot
    Set TBCotacao = CreateObject("adodb.recordset")
    TBCotacao.Open "Select * FROM Cotacao_item where coditem = '" & txtcodproduto.Text & "' and Cotacao_item.idcot = " & Cont, Conexao, adOpenKeyset, adLockOptimistic
    If TBCotacao.EOF = False Then
        Set TBItem = CreateObject("adodb.recordset")
        TBItem.Open "Select * from compras_pedido_lista where IDlista = " & TBCotacao!iditemlista, Conexao, adOpenKeyset, adLockOptimistic
        If TBItem.EOF = True Then TBItem.AddNew
        USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
        Evento = "Alterar produto/serviço"
    Else
        TBCotacao.AddNew
        Set TBItem = CreateObject("adodb.recordset")
        TBItem.Open "Select * from compras_pedido_lista", Conexao, adOpenKeyset, adLockOptimistic
        TBItem.AddNew
        USMsgBox ("Novo produto adicionado com sucesso."), vbInformation, "CAPRIND v5.0"
        Evento = "Novo produto/serviço"
    End If
    
    TBItem!ID_cotacao = Cont
    TBItem!IDpedido = 0
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select * from Projproduto where Desenho = '" & txtcodproduto & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
        TBItem!Codproduto = TBProduto!Codproduto
        If TBProduto!Tipo = "P" Or TBProduto!Tipo = "I" Then TBItem!Tipo = "P" Else TBItem!Tipo = "S"
    End If
    TBProduto.Close
    TBItem!Desenho = txtcodproduto.Text
    TBItem!Descricao = txtdesc.Text
    TBItem!Descricao_comercial = txtDescricao_comercial.Text
    TBItem!quant_req = txtQtde.Text
    TBItem!quant_req_PC = IIf(txtQtde_PC = "", Null, txtQtde_PC)
    TBItem!Un = txtUN.Text
    TBItem!Unidade_com = Txt_un_com
    TBItem!Familia = Lista_cadastrados.SelectedItem.ListSubItems(5)
    TBItem!Status_Item = "COTANDO"
    TBItem!Ordem = IIf(txtOrdem = "", Null, txtOrdem)
    TBItem!OS = IIf(Cmb_OS = "", Null, Cmb_OS)
    TBItem.Update
    
    TBCotacao!idcot = Cont
    TBCotacao!iditemlista = TBItem!IDlista
    TBCotacao!coditem = txtcodproduto.Text
    TBCotacao.Update
    
    Set TBCotacao = CreateObject("adodb.recordset")
    TBCotacao.Open "Select * FROM Cotacao_item where coditem = '" & txtcodproduto.Text & "' and Idcot = " & Cont, Conexao, adOpenKeyset, adLockOptimistic
    If TBCotacao.EOF = False Then
        .ProcGravaFornecedores TBCotacao!ID, TBItem!IDlista, txtcodproduto
    End If
    TBCotacao.Close
    
    TBItem.Close
    .ProcCarregaListaItens
    .ProcCarregaListaItens1
End With
ProcLimpaCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbfiltrarpor_sol_Click()
On Error GoTo tratar_erro

Lista_solicitados.ListItems.Clear
Txtprazo_sol.Value = Date
If cmbfiltrarpor_sol = "Família" Then
    txtTexto_sol.Visible = False
    cmbTexto_sol.Visible = True
    Txtprazo_sol.Visible = False
    ProcCarregaComboFamilia cmbTexto_sol, "familia <> 'Null' and (Compras = 'True' or Fabricacao = 'True')", True
ElseIf cmbfiltrarpor_sol = "Prazo entrega" Then
        txtTexto_sol.Visible = False
        cmbTexto_sol.Visible = False
        Txtprazo_sol.Visible = True
    Else
        txtTexto_sol.Visible = True
        cmbTexto_sol.Visible = False
        Txtprazo_sol.Visible = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbTexto_necess_Click()
On Error GoTo tratar_erro

ListaNecessidade.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbTexto_sol_Click()
On Error GoTo tratar_erro

Lista_solicitados.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbun_Click()
On Error GoTo tratar_erro

If cmbun <> "" Then Cmb_un_com = cmbun

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_visualizar_arquivo_Click()
On Error GoTo tratar_erro

If txtcodproduto = "" Then Exit Sub
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select imagem from projproduto where desenho = '" & txtcodproduto & "' and imagem is not null", Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    If TBProduto!imagem <> "" Then ProcAbrirArquivo TBProduto!imagem
End If
TBProduto.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdcalc_peso_Click()
On Error GoTo tratar_erro

If txtcodproduto = "" Then Exit Sub
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select Desenho, Unidade, Un_Kg, peso_metro from projproduto where desenho = '" & txtcodproduto & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    Engenharia = False
    Compras_Requisicao = False
    Compras_Cotacao = True
    Compras_Pedido = False
    Estoque_recebimento = False
    Vendas_Proposta = False
    Vendas_PI = False
    FrmCalculo_Peso.Show 1
End If
TBProduto.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluir()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If TXTIDLista.Text = 0 Then
   USMsgBox ("Informe o produto/serviço antes de excluir."), vbExclamation, "CAPRIND v5.0"
   Exit Sub
End If
If USMsgBox("Deseja realmente excluir este produto/serviço da cotação?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    Conexao.Execute "DELETE from compras_pedido_lista where idlista = " & TXTIDLista
    Set TBCotacao = CreateObject("adodb.recordset")
    TBCotacao.Open "Select ID from cotacao_item where iditemlista = " & TXTIDLista & " and idcot = " & Cont, Conexao, adOpenKeyset, adLockOptimistic
    If TBCotacao.EOF = False Then
        Conexao.Execute "DELETE from Cotacao_fornecedor where iditem = " & TBCotacao!ID
    End If
    Conexao.Execute "DELETE from cotacao_item where iditemlista = " & TXTIDLista
    USMsgBox ("Produto/serviço excluído com sucesso."), vbInformation, "CAPRIND v5.0"
    '==================================
    Modulo = "Compras/Cotação"
    Evento = "Excluir produto/serviço"
    ID_documento = TXTIDLista
    Documento = "Nº cotação: " & Cont
    Documento1 = "Cód. interno: " & txtCodproduto2
    ProcGravaEvento
    '==================================
    ProcLimpaCampos2
    ProcCarregaLista3
    With frmcompras_reqcot
        .ProcCarregaListaItens
        .ProcCarregaListaItens1
    End With
    Novo_compras_reqcot_item = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar_Cadastrados()
On Error GoTo tratar_erro

CamposFiltro = "P.codProduto, P.Desenho, P.Descricao, P.Unidade, P.Unidade_com, P.Classe"
INNERJOINTEXTO = "Select " & CamposFiltro & " from ((Projproduto P LEFT JOIN item_aplicacoes IA ON IA.codproduto = P.codproduto) LEFT JOIN Compras_pedido_lista CPL ON CPL.Desenho = P.Desenho) LEFT JOIN Projproduto_fabricante PFAB ON PFAB.Codproduto = P.codproduto"
TextoFiltroPadrao = "P.Compras = 'True' and P.bloqueado = 'False' group by " & CamposFiltro & " order by P.desenho"
If txtTexto_cad.Visible = True And txtTexto_cad <> "" Or cmbTexto_cad.Visible = True And cmbTexto_cad <> "" Then
    If cmbfiltrarpor_cad = "Família" Then
        StrSql_Cotacao_Localizar_Cadastrados = INNERJOINTEXTO & " where P.classe = '" & cmbTexto_cad & "' and " & TextoFiltroPadrao
    Else
        Select Case cmbfiltrarpor_cad
            Case "Código interno": TextoFiltro = "P.desenho"
            Case "Descrição": TextoFiltro = "P.descricao"
            Case "Descrição comercial": TextoFiltro = "P.descricaotecnica"
            Case "Código de referência": TextoFiltro = "IA.N_referencia"
            Case "Part number": TextoFiltro = "PFAB.Part_number"
        End Select
        StrSql_Cotacao_Localizar_Cadastrados = INNERJOINTEXTO & " where " & TextoFiltro & FunVerifTipoFiltroIMF(optInicio_cad, optMeio_cad, optFim_cad, optIgual_cad, txtTexto_cad) & " and " & TextoFiltroPadrao
    End If
Else
    StrSql_Cotacao_Localizar_Cadastrados = INNERJOINTEXTO & " where " & TextoFiltroPadrao
End If
ProcCarregalista_Cadastrados

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar_Necessidade()
On Error GoTo tratar_erro

If Opt_PCP.Value = True Then NomeTabela = "Cotacao_Necessidade" Else NomeTabela = "Cotacao_Necessidade_PIEST"
CamposFiltro = "CN.Codproduto, CN.Desenho, CN.Descricao, CN.Unidade, CN.Unidade_com, CN.Necessidade, CN.Necessidade_estoque, CN.Necessidade_PC"
INNERJOINTEXTO = "Select " & CamposFiltro & " from (" & NomeTabela & " CN LEFT JOIN item_aplicacoes IA ON CN.codproduto = IA.codproduto) LEFT JOIN Projproduto_fabricante PFAB ON PFAB.Codproduto = CN.codproduto"
If Cmb_filtrar = "Com necessidade" Then TextoFiltroEstoque = " and CN.Necessidade > 0" Else TextoFiltroEstoque = " and CN.Necessidade_estoque > 0"
TextoFiltroPadrao = "CN.ID_empresa = " & frmcompras_reqcot.Cmb_empresa.ItemData(frmcompras_reqcot.Cmb_empresa.ListIndex) & TextoFiltroEstoque & " group by " & CamposFiltro & " order by CN.desenho"

If txtTexto_necess.Visible = True And txtTexto_necess <> "" Or cmbTexto_necess.Visible = True And cmbTexto_necess <> "" Then
    If cmbfiltrarpor_necess = "Família" Then
        StrSql_Cotacao_Localizar_Necessidade = INNERJOINTEXTO & " where CN.classe = '" & cmbTexto_necess & "' and " & TextoFiltroPadrao
    Else
        Select Case cmbfiltrarpor_necess
            Case "Código interno": TextoFiltro = "CN.Desenho"
            Case "Código de referência": TextoFiltro = "IA.n_referencia"
            Case "Descrição": TextoFiltro = "CN.Descricao"
            Case "Part number": TextoFiltro = "PFAB.Part_number"
        End Select
        StrSql_Cotacao_Localizar_Necessidade = INNERJOINTEXTO & " where " & TextoFiltro & FunVerifTipoFiltroIMF(optInicio_necess, Optmeio_necess, Optfim_necess, optIgual_necess, txtTexto_necess) & " and " & TextoFiltroPadrao
    End If
Else
    StrSql_Cotacao_Localizar_Necessidade = INNERJOINTEXTO & " where " & TextoFiltroPadrao
End If
ProcCarregalista_Necessidade

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregalista_Necessidade()
On Error GoTo tratar_erro

If StrSql_Cotacao_Localizar_Necessidade = "" Then Exit Sub
lblRegistros(0).Caption = "Nº de registros: 0"
lblPaginas(0).Caption = "Página: 0 de: 0"
ListaNecessidade.ListItems.Clear
Set TBLocalizar_produto_padrao = CreateObject("adodb.recordset")
TBLocalizar_produto_padrao.Open StrSql_Cotacao_Localizar_Necessidade, Conexao, adOpenKeyset, adLockReadOnly
If TBLocalizar_produto_padrao.EOF = False Then ProcExibePagina_Necessidade (1)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina_Necessidade(Pagina)
On Error GoTo tratar_erro

ListaNecessidade.ListItems.Clear
TBLocalizar_produto_padrao.PageSize = IIf(txtNreg(0) = "", 30, txtNreg(0))
TBLocalizar_produto_padrao.AbsolutePage = Pagina
TamanhoPagina = TBLocalizar_produto_padrao.PageSize
ContadorReg = 1

PBLista.Min = 0
PBLista.Max = FunVerifMaxPBListaPaginacao(TBLocalizar_produto_padrao.RecordCount - IIf(Pagina > 1, (TBLocalizar_produto_padrao.PageSize * (Pagina - 1)), 0), TBLocalizar_produto_padrao.PageSize)
PBLista.Value = 1
Contador = 0
Do While TBLocalizar_produto_padrao.EOF = False And (ContadorReg <= TamanhoPagina)
    With ListaNecessidade.ListItems.Add(, , TBLocalizar_produto_padrao!Codproduto)
        .SubItems(1) = IIf(IsNull(TBLocalizar_produto_padrao!Desenho), "", TBLocalizar_produto_padrao!Desenho)
        .SubItems(2) = IIf(IsNull(TBLocalizar_produto_padrao!Descricao), "", TBLocalizar_produto_padrao!Descricao)
        If TBLocalizar_produto_padrao!Unidade <> TBLocalizar_produto_padrao!Unidade_com Then
            If Cmb_filtrar = "Com necessidade" Then qt = Format(TBLocalizar_produto_padrao!Necessidade, "###,##0.0000") Else qt = Format(TBLocalizar_produto_padrao!Necessidade_estoque, "###,##0.0000")
            If FunVerifUNConversao(TBLocalizar_produto_padrao!Unidade, TBLocalizar_produto_padrao!Unidade_com) = True Then
                Qtde = FunConverteUN(TBLocalizar_produto_padrao!Unidade_com, TBLocalizar_produto_padrao!Unidade, qt, TBLocalizar_produto_padrao!Desenho)
                .SubItems(4) = Format(Qtde, "###,##0.0000")
            Else
                If Cmb_filtrar = "Com necessidade" Then .SubItems(4) = Format(TBLocalizar_produto_padrao!Necessidade, "###,##0.0000") Else .SubItems(4) = Format(TBLocalizar_produto_padrao!Necessidade_estoque, "###,##0.0000")
            End If
        Else
            If Cmb_filtrar = "Com necessidade" Then .SubItems(4) = Format(TBLocalizar_produto_padrao!Necessidade, "###,##0.0000") Else .SubItems(4) = Format(TBLocalizar_produto_padrao!Necessidade_estoque, "###,##0.0000")
        End If
        .SubItems(5) = TBLocalizar_produto_padrao!Necessidade_PC
        If Cmb_filtrar = "Com necess. estoque" Then NReal = Format(TBLocalizar_produto_padrao!Necessidade_estoque, "###,##0.0000") Else NReal = Format(TBLocalizar_produto_padrao!Necessidade, "###,##0.0000")
        If NReal > 0 Then
            .ForeColor = vbRed
            .ListSubItems(1).ForeColor = vbRed
            .ListSubItems(2).ForeColor = vbRed
            .ListSubItems(3).ForeColor = vbRed
            .ListSubItems(4).ForeColor = vbRed
            .ListSubItems(5).ForeColor = vbRed
        End If
    End With
    TBLocalizar_produto_padrao.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista.Value = Contador
Loop
lblRegistros(0).Caption = "Nº de registros: " & TBLocalizar_produto_padrao.RecordCount
If TBLocalizar_produto_padrao.AbsolutePage = adPosBOF Then
   lblPaginas(0).Caption = "Página: 1 de: " & TBLocalizar_produto_padrao.PageCount
ElseIf TBLocalizar_produto_padrao.AbsolutePage = adPosEOF Then
        lblPaginas(0).Caption = "Página: " & TBLocalizar_produto_padrao.PageCount & " de: " & TBLocalizar_produto_padrao.PageCount
    Else
        lblPaginas(0).Caption = "Página: " & TBLocalizar_produto_padrao.AbsolutePage - 1 & " de: " & TBLocalizar_produto_padrao.PageCount
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar_Solicitacao()
On Error GoTo tratar_erro

CamposFiltro = "CR.ID_requisicao, CR.Requisicaotexto, CPL.IdLista, CPL.Status_Item, CPL.desenho, CPL.descricao, CPL.Un, CPL.Unidade_com, CPL.quant_req, CPL.quant_req_PC, CPL.detalheitem, CPL.prazoreq, CPL.obs"
INNERJOINTEXTO = "Select " & CamposFiltro & " from (Compras_requisicao CR INNER JOIN Compras_pedido_lista CPL ON CR.ID_Requisicao = CPL.ID_Requisicao) LEFT JOIN Projproduto_fabricante PFAB ON PFAB.Codproduto = CPL.codproduto"
TextoFiltroPadrao = "(CPL.status_item = 'REQUISIT.' or CPL.status_item = 'COTANDO' and CPL.id_cotacao = " & frmcompras_reqcot.txtidcot & ")  and CR.Status = 'LIBERADA' group by " & CamposFiltro & " order by CR.ID_requisicao"

If txtTexto_sol.Visible = True And txtTexto_sol <> "" Or cmbTexto_sol.Visible = True And cmbTexto_sol <> "" Or Txtprazo_sol.Visible = True Then
    If cmbfiltrarpor_sol = "Família" Then
        StrSql_Cotacao_Localizar_Solicitacao = INNERJOINTEXTO & " where CPL.Familia = '" & cmbTexto_sol & "' and " & TextoFiltroPadrao
    ElseIf cmbfiltrarpor_sol = "Prazo entrega" Then
            StrSql_Cotacao_Localizar_Solicitacao = INNERJOINTEXTO & " where CPL.Prazoreq = '" & Format(Txtprazo_sol.Value, "Short Date") & "' and " & TextoFiltroPadrao
        Else
            Select Case cmbfiltrarpor_sol
                Case "Solicitação": TextoFiltro = "CR.Requisicaotexto"
                Case "Código interno": TextoFiltro = "CPL.desenho"
                Case "Descrição": TextoFiltro = "CPL.descricao"
                Case "Descrição comercial": TextoFiltro = "CPL.descricao_comercial"
                Case "Detalhe": TextoFiltro = "CPL.Detalheitem"
                Case "Part number": TextoFiltro = "PFAB.Part_number"
            End Select
            StrSql_Cotacao_Localizar_Solicitacao = INNERJOINTEXTO & " where " & TextoFiltro & FunVerifTipoFiltroIMF(optInicio_sol, optMeio_sol, optFim_sol, optIgual_sol, txtTexto_sol) & " and " & TextoFiltroPadrao
    End If
Else
    StrSql_Cotacao_Localizar_Solicitacao = INNERJOINTEXTO & " where " & TextoFiltroPadrao
End If
ProcCarregalista_Solicitacao

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregalista_Solicitacao()
On Error GoTo tratar_erro

lblRegistros(1).Caption = "Nº de reg.: 0"
lblPaginas(1).Caption = "Página: 0 de: 0"
Lista_solicitados.ListItems.Clear
If StrSql_Cotacao_Localizar_Solicitacao = "" Then Exit Sub
Set TBLocalizar_produto_padrao1 = CreateObject("adodb.recordset")
TBLocalizar_produto_padrao1.Open StrSql_Cotacao_Localizar_Solicitacao, Conexao, adOpenKeyset, adLockReadOnly
If TBLocalizar_produto_padrao1.EOF = False Then ProcExibePagina_Solicitacao (1)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina_Solicitacao(Pagina)
On Error GoTo tratar_erro

Lista_solicitados.ListItems.Clear
TBLocalizar_produto_padrao1.PageSize = IIf(txtNreg(1) = "", 30, txtNreg(1))
TBLocalizar_produto_padrao1.AbsolutePage = Pagina
TamanhoPagina = TBLocalizar_produto_padrao1.PageSize
ContadorReg = 1

PBLista.Min = 0
PBLista.Max = FunVerifMaxPBListaPaginacao(TBLocalizar_produto_padrao1.RecordCount - IIf(Pagina > 1, (TBLocalizar_produto_padrao1.PageSize * (Pagina - 1)), 0), TBLocalizar_produto_padrao1.PageSize)
PBLista.Value = 1
Contador = 0
Do While TBLocalizar_produto_padrao1.EOF = False And (ContadorReg <= TamanhoPagina)
    With Lista_solicitados.ListItems
        .Add , , TBLocalizar_produto_padrao1!IDlista
        .Item(.Count).SubItems(1) = IIf(IsNull(TBLocalizar_produto_padrao1!Status_Item), "", TBLocalizar_produto_padrao1!Status_Item)
        .Item(.Count).SubItems(2) = IIf(IsNull(TBLocalizar_produto_padrao1!Requisicaotexto), "", TBLocalizar_produto_padrao1!Requisicaotexto)
        .Item(.Count).SubItems(3) = IIf(IsNull(TBLocalizar_produto_padrao1!Desenho), "", TBLocalizar_produto_padrao1!Desenho)
        .Item(.Count).SubItems(4) = IIf(IsNull(TBLocalizar_produto_padrao1!Descricao), "", TBLocalizar_produto_padrao1!Descricao)
        .Item(.Count).SubItems(5) = IIf(IsNull(TBLocalizar_produto_padrao1!Un), "", TBLocalizar_produto_padrao1!Un)
        .Item(.Count).SubItems(6) = IIf(IsNull(TBLocalizar_produto_padrao1!Unidade_com), "", TBLocalizar_produto_padrao1!Unidade_com)
        If TBLocalizar_produto_padrao1!Un <> TBLocalizar_produto_padrao1!Unidade_com Then valor = FunConversaoFinalUn(TBLocalizar_produto_padrao1!Un, TBLocalizar_produto_padrao1!Unidade_com, TBLocalizar_produto_padrao1!quant_req, TBLocalizar_produto_padrao1!Desenho, True) Else valor = TBLocalizar_produto_padrao1!quant_req
        .Item(.Count).SubItems(7) = FunFormataCasasDecimais(4, valor)
        .Item(.Count).SubItems(8) = IIf(IsNull(TBLocalizar_produto_padrao1!quant_req), "", FunFormataCasasDecimais(4, TBLocalizar_produto_padrao1!quant_req))
        .Item(.Count).SubItems(9) = IIf(IsNull(TBLocalizar_produto_padrao1!detalheitem), "", TBLocalizar_produto_padrao1!detalheitem)
        .Item(.Count).SubItems(10) = IIf(IsNull(TBLocalizar_produto_padrao1!prazoreq), "", Format(TBLocalizar_produto_padrao1!prazoreq, "dd/mm/yy"))
        .Item(.Count).SubItems(11) = IIf(IsNull(TBLocalizar_produto_padrao1!Obs), "", TBLocalizar_produto_padrao1!Obs)
    End With
    ContadorReg = ContadorReg + 1
    TBLocalizar_produto_padrao1.MoveNext
    Contador = Contador + 1
    PBLista.Value = Contador
Loop
lblRegistros(1).Caption = "Nº de reg.: " & TBLocalizar_produto_padrao1.RecordCount
If TBLocalizar_produto_padrao1.AbsolutePage = adPosBOF Then
   lblPaginas(1).Caption = "Pág.: 1 de: " & TBLocalizar_produto_padrao1.PageCount
ElseIf TBLocalizar_produto_padrao1.AbsolutePage = adPosEOF Then
        lblPaginas(1).Caption = "Pág.: " & TBLocalizar_produto_padrao1.PageCount & " de: " & TBLocalizar_produto_padrao1.PageCount
    Else
        lblPaginas(1).Caption = "Pág.: " & TBLocalizar_produto_padrao1.AbsolutePage - 1 & " de: " & TBLocalizar_produto_padrao1.PageCount
End If

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
ProcLimpaCampos2
Novo_compras_reqcot_item = True
Frame10.Enabled = True
txtCodproduto2.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

If Novo_compras_reqcot_item = True Then
    If USMsgBox("O produto ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvar
        If Novo_compras_reqcot_item = True Then Exit Sub Else Unload Me
    End If
End If
Novo_compras_reqcot_item = False
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do$erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvar()
On Error GoTo tratar_erro
  
If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Frame10.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
Acao = "salvar"
If txtCodproduto2.Text = "" Then
    NomeCampo = "o código interno"
    ProcVerificaAcao
    txtCodproduto2.SetFocus
    Exit Sub
End If
Set TBItem = CreateObject("adodb.recordset")
TBItem.Open "Select * from projproduto where desenho = '" & txtCodproduto2 & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBItem.EOF = False Then
    USMsgBox ("Este código interno já foi adicionado na lista, favor alterar."), vbExclamation, "CAPRIND v5.0"
    TBItem.Close
    Exit Sub
End If
TBItem.Close
If cmbFamilia3.Text = "" Then
    NomeCampo = "a família"
    ProcVerificaAcao
    cmbFamilia3.SetFocus
    Exit Sub
End If
If cmbun.Text = "" Then
    NomeCampo = "a unidade de estoque"
    ProcVerificaAcao
    cmbun.SetFocus
    Exit Sub
End If
If Cmb_un_com.Text = "" Then
    NomeCampo = "a unidade comercial"
    ProcVerificaAcao
    Cmb_un_com.SetFocus
    Exit Sub
End If
valor = IIf(txtQtde2 = "", 0, txtQtde2)
If valor <= 0 Then
    NomeCampo = "a quantidade"
    ProcVerificaAcao
    txtQtde2.SetFocus
    Exit Sub
End If
If txtDescricao2.Text = "" Then
    NomeCampo = "a descrição"
    ProcVerificaAcao
    txtDescricao2.SetFocus
    Exit Sub
End If
If txtDescricao_comercial2.Text = "" Then
    NomeCampo = "a descrição comercial"
    ProcVerificaAcao
    txtDescricao_comercial2.SetFocus
    Exit Sub
End If
Set TBCompras = CreateObject("adodb.recordset")
TBCompras.Open "Select * from compras_pedido_lista where IDlista = " & TXTIDLista, Conexao, adOpenKeyset, adLockOptimistic
If TBCompras.EOF = True Then TBCompras.AddNew
TBCompras!ID_cotacao = Cont
TBCompras!IDpedido = 0
TBCompras!Desenho = txtCodproduto2.Text
TBCompras!Descricao = txtDescricao2.Text
TBCompras!Descricao_comercial = txtDescricao_comercial2.Text
TBCompras!quant_req = txtQtde2.Text
TBCompras!Un = cmbun.Text
TBCompras!Unidade_com = Cmb_un_com.Text
TBCompras!Familia = cmbFamilia3
TBCompras!Status_Item = "COTANDO"
TBCompras!Tipo = "P"
TBCompras.Update
TXTIDLista = TBCompras!IDlista

With frmcompras_reqcot
    Set TBCotacao = CreateObject("adodb.recordset")
    TBCotacao.Open "Select * from cotacao_item where iditemlista = " & TXTIDLista & " and idcot = " & Cont, Conexao, adOpenKeyset, adLockOptimistic
    If TBCotacao.EOF = True Then
        TBCotacao.AddNew
        Evento = "Novo produto/serviço"
    Else
        Evento = "Alterar produto/serviço"
    End If
    TBCotacao!idcot = Cont
    TBCotacao!iditemlista = TXTIDLista
    TBCotacao!coditem = txtCodproduto2.Text
    TBCotacao.Update
    
    .ProcGravaFornecedores TBCotacao!ID, TXTIDLista, txtCodproduto2
    
    TBCotacao.Close
    
    TBCompras.Close
    ProcCarregaLista3

    .ProcCarregaListaItens
    .ProcCarregaListaItens1
    
    If Novo_compras_reqcot_item = True Then
        USMsgBox ("Novo produto/serviço adicionado com sucesso a cotação."), vbInformation, "CAPRIND v5.0"
        Evento = "Novo produto/serviço"
    Else
        USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
        Evento = "Alterar produto/serviço"
        If Lista_Ncadastrados.ListItems.Count <> 0 And CodigoLista <> 0 Then
            Lista_Ncadastrados.SelectedItem = Lista_Ncadastrados.ListItems(CodigoLista)
            Lista_Ncadastrados.SetFocus
        End If
    End If
    Novo_compras_reqcot_item = False
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt_Click(index As Integer)
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas(index).Caption, 4)) <= 1 Then Exit Sub
Select Case SSTab1.Tab
    Case 0:
        If TBLocalizar_produto_padrao.AbsolutePage <> 2 Then
            If TBLocalizar_produto_padrao.AbsolutePage = -3 Then
                ProcExibePagina_Necessidade (TBLocalizar_produto_padrao.PageCount - 1)
            Else
                TBLocalizar_produto_padrao.AbsolutePage = TBLocalizar_produto_padrao.AbsolutePage - 2
                ProcExibePagina_Necessidade (TBLocalizar_produto_padrao.AbsolutePage)
            End If
        Else
            ProcExibePagina_Necessidade (1)
        End If
    Case 1:
        If TBLocalizar_produto_padrao.AbsolutePage <> 2 Then
            If TBLocalizar_produto_padrao.AbsolutePage = -3 Then
                ProcExibePagina_Solicitacao (TBLocalizar_produto_padrao.PageCount - 1)
            Else
                TBLocalizar_produto_padrao.AbsolutePage = TBLocalizar_produto_padrao.AbsolutePage - 2
                ProcExibePagina_Solicitacao (TBLocalizar_produto_padrao.AbsolutePage)
            End If
        Else
            ProcExibePagina_Solicitacao (1)
        End If
    Case 2:
        If TBLocalizar_produto_padrao2.AbsolutePage <> 2 Then
            If TBLocalizar_produto_padrao2.AbsolutePage = -3 Then
                ProcExibePagina_Cadastrados (TBLocalizar_produto_padrao2.PageCount - 1)
            Else
                TBLocalizar_produto_padrao2.AbsolutePage = TBLocalizar_produto_padrao2.AbsolutePage - 2
                ProcExibePagina_Cadastrados (TBLocalizar_produto_padrao2.AbsolutePage)
            End If
        Else
            ProcExibePagina_Cadastrados (1)
        End If
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagIr_Click(index As Integer)
On Error GoTo tratar_erro

If txtPagIr(index) = "" Then Exit Sub
Quant = ReturnNumbersOnly(Right(lblPaginas(index).Caption, 4))
If Quant <= 1 Or txtPagIr(index) > Quant Then Exit Sub
If txtPagIr(index).Text >= 1 And txtPagIr(index).Text <= Quant Then
    Select Case SSTab1.Tab
        Case 0:
            TBLocalizar_produto_padrao.AbsolutePage = txtPagIr(index).Text
            ProcExibePagina_Necessidade (TBLocalizar_produto_padrao.AbsolutePage)
        Case 1:
            TBLocalizar_produto_padrao.AbsolutePage = txtPagIr(index).Text
            ProcExibePagina_Solicitacao (TBLocalizar_produto_padrao.AbsolutePage)
        Case 2:
            TBLocalizar_produto_padrao2.AbsolutePage = txtPagIr(index).Text
            ProcExibePagina_Cadastrados (TBLocalizar_produto_padrao2.AbsolutePage)
    End Select
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim_Click(index As Integer)
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas(index).Caption, 4)) <= 1 Then Exit Sub
Select Case SSTab1.Tab
    Case 0:
        TBLocalizar_produto_padrao.AbsolutePage = 1
        ProcExibePagina_Necessidade (TBLocalizar_produto_padrao.AbsolutePage)
    Case 1:
        TBLocalizar_produto_padrao.AbsolutePage = 1
        ProcExibePagina_Solicitacao (TBLocalizar_produto_padrao.AbsolutePage)
    Case 2:
        TBLocalizar_produto_padrao2.AbsolutePage = 1
        ProcExibePagina_Cadastrados (TBLocalizar_produto_padrao2.AbsolutePage)
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx_Click(index As Integer)
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas(index).Caption, 4)) <= 1 Then Exit Sub
Select Case SSTab1.Tab
    Case 0:
        If TBLocalizar_produto_padrao.AbsolutePage <> -3 Then
            If TBLocalizar_produto_padrao.AbsolutePage = 1 Then
                ProcExibePagina_Necessidade (2)
            Else
                ProcExibePagina_Necessidade (TBLocalizar_produto_padrao.AbsolutePage)
            End If
        Else
            ProcExibePagina_Necessidade (TBLocalizar_produto_padrao.PageCount)
        End If
    Case 1:
        If TBLocalizar_produto_padrao.AbsolutePage <> -3 Then
            If TBLocalizar_produto_padrao.AbsolutePage = 1 Then
                ProcExibePagina_Solicitacao (2)
            Else
                ProcExibePagina_Solicitacao (TBLocalizar_produto_padrao.AbsolutePage)
            End If
        Else
            ProcExibePagina_Solicitacao (TBLocalizar_produto_padrao.PageCount)
        End If
    Case 2:
        If TBLocalizar_produto_padrao2.AbsolutePage <> -3 Then
            If TBLocalizar_produto_padrao2.AbsolutePage = 1 Then
                ProcExibePagina_Cadastrados (2)
            Else
                ProcExibePagina_Cadastrados (TBLocalizar_produto_padrao2.AbsolutePage)
            End If
        Else
            ProcExibePagina_Cadastrados (TBLocalizar_produto_padrao2.PageCount)
        End If
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt_Click(index As Integer)
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas(index).Caption, 4)) <= 1 Then Exit Sub
Select Case SSTab1.Tab
    Case 0:
        TBLocalizar_produto_padrao.AbsolutePage = TBLocalizar_produto_padrao.PageCount
        ProcExibePagina_Necessidade (TBLocalizar_produto_padrao.AbsolutePage)
    Case 1:
        TBLocalizar_produto_padrao.AbsolutePage = TBLocalizar_produto_padrao.PageCount
        ProcExibePagina_Solicitacao (TBLocalizar_produto_padrao.AbsolutePage)
    Case 2:
        TBLocalizar_produto_padrao2.AbsolutePage = TBLocalizar_produto_padrao2.PageCount
        ProcExibePagina_Cadastrados (TBLocalizar_produto_padrao2.AbsolutePage)
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case SSTab1.Tab
    Case 0:
        Select Case KeyCode
            Case vbKeyF2: ProcFiltrar_Necessidade
            Case vbKeyF3: ProcAdicionar_Necessidade
            Case vbKeyEscape: ProcSair
        End Select
    Case 1:
        Select Case KeyCode
            Case vbKeyF2: ProcFiltrar_Solicitacao
            Case vbKeyF3: ProcAdicionar_Solicitados
            Case vbKeyEscape: ProcSair
        End Select
    Case 2:
        Select Case KeyCode
            Case vbKeyF2: ProcFiltrar_Cadastrados
            Case vbKeyF3: procAdicionar_Cadastrados
            Case vbKeyEscape: ProcSair
        End Select
    Case 3:
        Select Case KeyCode
            Case vbKeyInsert: ProcNovo
            Case vbKeyF3: ProcSalvar
            Case vbKeyF4: ProcExcluir
            Case vbKeyEscape: ProcSair
        End Select
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar4 Me, 14325, 6, True
ProcCarregaToolBar1 Me, 14325, 6, True
ProcCarregaToolBar2 Me, 14325, 6, True
ProcCarregaToolBar3 Me, 14325, 7, True
SSTab1.Tab = 0
ProcCarregaComboFamilia cmbFamilia3, "familia IS NOT NULL and Compras = 'True'", False
ProcCarregaComboUnidade cmbun, False
ProcCarregaComboUnidade Cmb_un_com, False

ProcFiltroPadrao cmbfiltrarpor_necess, Optmeio_necess, Optfim_necess, optIgual_necess, frmcompras_reqcot.Cmb_empresa.ItemData(frmcompras_reqcot.Cmb_empresa.ListIndex), "Produtos/Serviços", "C", True
ProcFiltroPadrao cmbfiltrarpor_sol, optMeio_sol, optFim_sol, optIgual_sol, frmcompras_reqcot.Cmb_empresa.ItemData(frmcompras_reqcot.Cmb_empresa.ListIndex), "Produtos/Serviços", "C", True
ProcFiltroPadrao cmbfiltrarpor_cad, optMeio_cad, optFim_cad, optIgual_cad, frmcompras_reqcot.Cmb_empresa.ItemData(frmcompras_reqcot.Cmb_empresa.ListIndex), "Produtos/Serviços", "C", True
If Permitido = False Then
    cmbfiltrarpor_necess = "Código interno"
    cmbfiltrarpor_sol = "Código interno"
    cmbfiltrarpor_cad = "Código interno"
End If

Cmb_filtrar = "Com necessidade"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAdicionar_Solicitados()
On Error GoTo tratar_erro

With Lista_solicitados
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente adicionar este(s) produto(s)/serviço(s) nesta cotação?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            frmcompras_reqcot.ProcNovo_Solicitacao .ListItems.Item(InitFor)
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) produto(s)/serviço(s) antes de adicionar na cotação."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Produto(s)/serviço(s) adicionado(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    With frmcompras_reqcot
        .ProcCarregaListaItens
        .ProcCarregaListaItens1
    End With
    ProcCarregalista_Solicitacao
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaNecessidade_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With ListaNecessidade
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then .ListItems.Item(InitFor).Checked = False Else .ListItems.Item(InitFor).Checked = True
        Next InitFor
    End With
Else
    ProcOrdenaListView ListaNecessidade, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_solicitados_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Lista_solicitados
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then .ListItems.Item(InitFor).Checked = False Else .ListItems.Item(InitFor).Checked = True
        Next InitFor
    End With
Else
    ProcOrdenaListView Lista_solicitados, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_cadastrados_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView Lista_cadastrados, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_cadastrados_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

ProcLimpaCampos
ProcPuxaDados

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_Ncadastrados_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView Lista_Ncadastrados, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_Ncadastrados_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro
  
If Lista_Ncadastrados.ListItems.Count = 0 Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from compras_pedido_lista where IdLista = " & Lista_Ncadastrados.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    ProcLimpaCampos2
    TXTIDLista = IIf(IsNull(TBLISTA!IDlista), "", TBLISTA!IDlista)
    txtCodproduto2.Text = IIf(IsNull(TBLISTA!Desenho), "", TBLISTA!Desenho)
    If IsNull(TBLISTA!Un) = False And TBLISTA!Un <> "" Then cmbun.Text = TBLISTA!Un
    If IsNull(TBLISTA!Unidade_com) = False And TBLISTA!Unidade_com <> "" Then Cmb_un_com.Text = TBLISTA!Unidade_com
    txtDescricao2.Text = IIf(IsNull(TBLISTA!Descricao), "", TBLISTA!Descricao)
    txtDescricao_comercial2.Text = IIf(IsNull(TBLISTA!Descricao_comercial), "", TBLISTA!Descricao_comercial)
    If IsNull(TBLISTA!Familia) = False Then cmbFamilia3 = TBLISTA!Familia
    txtQtde2 = IIf(IsNull(TBLISTA!quant_req), "0,0000", Format(TBLISTA!quant_req, "###,##0.0000"))
    Novo_compras_reqcot_item = False
    Frame10.Enabled = True
    CodigoLista = Lista_Ncadastrados.SelectedItem.index
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt_PCP_Click()
On Error GoTo tratar_erro

ListaNecessidade.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt_vendas_Click()
On Error GoTo tratar_erro

ListaNecessidade.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optfim_cad_Click()
On Error GoTo tratar_erro

Lista_cadastrados.ListItems.Clear
ProcLimpaCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optfim_necess_Click()
On Error GoTo tratar_erro

ListaNecessidade.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub OptFim_sol_Click()
On Error GoTo tratar_erro

Lista_solicitados.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optIgual_cad_Click()
On Error GoTo tratar_erro

Lista_cadastrados.ListItems.Clear
ProcLimpaCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optIgual_necess_Click()
On Error GoTo tratar_erro

ListaNecessidade.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optIgual_Sol_Click()
On Error GoTo tratar_erro

Lista_solicitados.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optinicio_cad_Click()
On Error GoTo tratar_erro

Lista_cadastrados.ListItems.Clear
ProcLimpaCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optinicio_necess_Click()
On Error GoTo tratar_erro

ListaNecessidade.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optinicio_sol_Click()
On Error GoTo tratar_erro

Lista_solicitados.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optmeio_cad_Click()
On Error GoTo tratar_erro

Lista_cadastrados.ListItems.Clear
ProcLimpaCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optmeio_necess_Click()
On Error GoTo tratar_erro

ListaNecessidade.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optmeio_sol_Click()
On Error GoTo tratar_erro

Lista_solicitados.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error GoTo tratar_erro

Lista_Ncadastrados.Visible = False
Select Case SSTab1.Tab
    Case 0: If cmbfiltrarpor_necess.Visible = True Then cmbfiltrarpor_necess.SetFocus
    Case 1: cmbfiltrarpor_sol.SetFocus
    Case 2: cmbfiltrarpor_cad.SetFocus
    Case 3:
        With Lista_Ncadastrados
            .Visible = True
            .SetFocus
            .ListItems.Clear
        End With
        ProcCarregaLista3
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcPuxaDados()
On Error GoTo tratar_erro

Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from projproduto where codproduto = " & Lista_cadastrados.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    txtcodproduto.Text = IIf(IsNull(TBLISTA!Desenho), "", TBLISTA!Desenho)
    txtUN.Text = IIf(IsNull(TBLISTA!Unidade), "", TBLISTA!Unidade)
    Txt_un_com.Text = IIf(IsNull(TBLISTA!Unidade_com), "", TBLISTA!Unidade_com)
    txtdesc.Text = IIf(IsNull(TBLISTA!Descricao), "", TBLISTA!Descricao)
    txtDescricao_comercial.Text = IIf(IsNull(TBLISTA!descricaotecnica), "", TBLISTA!descricaotecnica)
End If
TBLISTA.Close
Frame5.Enabled = True

Exit Sub
tratar_erro:
    USMsgBox ("descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregalista_Cadastrados()
On Error GoTo tratar_erro

lblRegistros(2).Caption = "Nº de registros: 0"
lblPaginas(2).Caption = "Página: 0 de: 0"
Lista_cadastrados.ListItems.Clear
If StrSql_Cotacao_Localizar_Cadastrados = "" Then Exit Sub
Set TBLocalizar_produto_padrao2 = CreateObject("adodb.recordset")
TBLocalizar_produto_padrao2.Open StrSql_Cotacao_Localizar_Cadastrados, Conexao, adOpenKeyset, adLockReadOnly
If TBLocalizar_produto_padrao2.EOF = False Then ProcExibePagina_Cadastrados (1)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina_Cadastrados(Pagina)
On Error GoTo tratar_erro

Lista_cadastrados.ListItems.Clear
TBLocalizar_produto_padrao2.PageSize = IIf(txtNreg(2) = "", 30, txtNreg(2))
TBLocalizar_produto_padrao2.AbsolutePage = Pagina
TamanhoPagina = TBLocalizar_produto_padrao2.PageSize
ContadorReg = 1

PBLista.Min = 0
PBLista.Max = FunVerifMaxPBListaPaginacao(TBLocalizar_produto_padrao2.RecordCount - IIf(Pagina > 1, (TBLocalizar_produto_padrao2.PageSize * (Pagina - 1)), 0), TBLocalizar_produto_padrao2.PageSize)
PBLista.Value = 1
Contador = 0
Do While TBLocalizar_produto_padrao2.EOF = False And (ContadorReg <= TamanhoPagina)
    With Lista_cadastrados.ListItems
        .Add , , TBLocalizar_produto_padrao2!Codproduto
        .Item(.Count).SubItems(1) = IIf(IsNull(TBLocalizar_produto_padrao2!Desenho), "", TBLocalizar_produto_padrao2!Desenho)
        .Item(.Count).SubItems(2) = IIf(IsNull(TBLocalizar_produto_padrao2!Descricao), "", TBLocalizar_produto_padrao2!Descricao)
        .Item(.Count).SubItems(3) = IIf(IsNull(TBLocalizar_produto_padrao2!Unidade), "", TBLocalizar_produto_padrao2!Unidade)
        .Item(.Count).SubItems(4) = IIf(IsNull(TBLocalizar_produto_padrao2!Unidade_com), "", TBLocalizar_produto_padrao2!Unidade_com)
        .Item(.Count).SubItems(5) = IIf(IsNull(TBLocalizar_produto_padrao2!Classe), "", TBLocalizar_produto_padrao2!Classe)
        .Item(.Count).SubItems(6) = Format(FunVerificaQtdeEstoque(TBLocalizar_produto_padrao2!Desenho, 0, ""), "###,##0.0000")
    End With
    TBLocalizar_produto_padrao2.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista.Value = Contador
Loop
lblRegistros(2).Caption = "Nº de registros: " & TBLocalizar_produto_padrao2.RecordCount
If TBLocalizar_produto_padrao2.AbsolutePage = adPosBOF Then
   lblPaginas(2).Caption = "Página: 1 de: " & TBLocalizar_produto_padrao2.PageCount
ElseIf TBLocalizar_produto_padrao2.AbsolutePage = adPosEOF Then
        lblPaginas(2).Caption = "Página: " & TBLocalizar_produto_padrao2.PageCount & " de: " & TBLocalizar_produto_padrao2.PageCount
    Else
        lblPaginas(2).Caption = "Página: " & TBLocalizar_produto_padrao2.AbsolutePage - 1 & " de: " & TBLocalizar_produto_padrao2.PageCount
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCampos()
On Error GoTo tratar_erro

txtcodproduto = ""
txtUN.Text = ""
Txt_un_com.Text = ""
txtQtde.Text = ""
txtQtde_PC = ""
txtdesc.Text = ""
txtDescricao_comercial.Text = ""
txtOrdem = ""
Cmb_OS.ListIndex = -1
Cmb_OS.Locked = True
Cmb_OS.TabStop = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCampos2()
On Error GoTo tratar_erro

TXTIDLista = 0
txtCodproduto2 = ""
cmbun.ListIndex = -1
Cmb_un_com.ListIndex = -1
txtQtde2 = ""
txtDescricao2.Text = ""
txtDescricao_comercial2.Text = ""
cmbFamilia3.ListIndex = -1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txtprazo_sol_Change()
On Error GoTo tratar_erro

Lista_solicitados.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtNreg_Change(index As Integer)
On Error GoTo tratar_erro

If txtNreg(index) <> "" Then
    VerifNumero = txtNreg(index)
    ProcVerificaNumero
    If VerifNumero = False Then
        txtNreg(index) = ""
        txtNreg(index).SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtPagIr_Change(index As Integer)
On Error GoTo tratar_erro

If txtPagIr(index) <> "" Then
    VerifNumero = txtPagIr(index)
    ProcVerificaNumero
    If VerifNumero = False Then
        txtPagIr(index) = ""
        txtPagIr(index).SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtOrdem_LostFocus()
On Error GoTo tratar_erro

Cmb_OS.Clear
Cmb_OS.Locked = True
Cmb_OS.TabStop = False
If txtOrdem <> "" Then
    VerifNumero = txtOrdem
    ProcVerificaNumero
    If VerifNumero = False Then
        txtOrdem = ""
        txtOrdem.SetFocus
        Exit Sub
    End If
    ProcCarregaOS
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaOS()
On Error GoTo tratar_erro

If txtOrdem = "" Or txtOrdem = 0 Then Exit Sub
Cmb_OS.Clear
Set TBproducao = CreateObject("adodb.recordset")
TBproducao.Open "Select * from Producao where Ordem = " & txtOrdem & " and id_empresa = " & frmcompras_reqcot.Cmb_empresa.ItemData(frmcompras_reqcot.Cmb_empresa.ListIndex), Conexao, adOpenKeyset, adLockOptimistic
If TBproducao.EOF = False Then
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select * from projproduto where desenho = '" & txtcodproduto & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
        Set TBOrdem = CreateObject("adodb.recordset")
        TBOrdem.Open "Select * from ordemservico where Ordem = " & TBproducao!Ordem & " order by fase, retrabalho, IDproducao", Conexao, adOpenKeyset, adLockOptimistic
        If TBOrdem.EOF = False Then
            Do While TBOrdem.EOF = False
                IDFase = 0
                Set TBFI = CreateObject("adodb.recordset")
                TBFI.Open "Select Ordemservico_maq_utilizadas.* from Ordemservico_maq_utilizadas INNER JOIN CadMaquinas on Ordemservico_maq_utilizadas.Maquina = CadMaquinas.Maquina where Ordemservico_maq_utilizadas.Ordem = " & TBOrdem!IDProducao & " and CadMaquinas.Custos = 'False'", Conexao, adOpenKeyset, adLockOptimistic
                If TBFI.EOF = False Then
                    Do While TBFI.EOF = False
                        If IDFase <> TBFI!OS Then Cmb_OS.AddItem TBFI!OS
                        IDFase = TBFI!OS
                        TBFI.MoveNext
                    Loop
                Else
                    If TBOrdem!custos = False Then Cmb_OS.AddItem TBOrdem!IDProducao
                End If
                TBFI.Close
                TBOrdem.MoveNext
            Loop
        End If
        TBOrdem.Close
        Cmb_OS.Locked = False
        Cmb_OS.TabStop = True
    End If
    TBProduto.Close
Else
    USMsgBox ("Não foi encontrado nenhuma ordem de fabricação e montagem com este número."), vbExclamation, "CAPRIND v5.0"
    txtOrdem = ""
    txtOrdem.SetFocus
End If
TBproducao.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtQtde_Change()
On Error GoTo tratar_erro

If txtQtde <> "" Then
    VerifNumero = txtQtde
    ProcVerificaNumero
    If VerifNumero = False Then
        txtQtde = ""
        txtQtde.SetFocus
        Exit Sub
    End If
    If txtUN <> Txt_un_com Then
        txtqtde_est = FunFormataCasasDecimais(4, FunConversaoFinalUn(txtUN, Txt_un_com, txtQtde, txtcodproduto, True))
    Else
        txtqtde_est = FunFormataCasasDecimais(4, txtQtde)
    End If
    If FunVerifMovimentacaoEstPC(frmcompras_reqcot.Cmb_empresa.ItemData(frmcompras_reqcot.Cmb_empresa.ListIndex)) = True Then
        txtQtde_PC = FunCalculaQtdePC(txtcodproduto, txtQtde, True, Txt_un_com)
    Else
        txtQtde_PC = ""
    End If
Else
    txtqtde_est = ""
    txtQtde_PC = ""
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtQtde_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus txtQtde

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtqtde_LostFocus()
On Error GoTo tratar_erro

txtQtde.Text = Format(txtQtde.Text, "###,##0.0000")
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtqtde_PC_Change()
On Error GoTo tratar_erro

If txtQtde_PC <> "" Then
    VerifNumero = txtQtde_PC
    ProcVerificaNumero
    If VerifNumero = False Then
        txtQtde_PC = ""
        txtQtde_PC.SetFocus
        Exit Sub
    End If
    If txtQtde_PC.Locked = False Then txtQtde = FunCalculaQtdePC(txtcodproduto, txtQtde_PC, False, Txt_un_com)
Else
    If txtQtde_PC.Locked = False Then txtQtde = ""
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtqtde_PC_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus txtQtde_PC

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtQtde2_Change()
On Error GoTo tratar_erro

If txtQtde2 <> "" Then
    VerifNumero = txtQtde2
    ProcVerificaNumero
    If VerifNumero = False Then
        txtQtde2 = ""
        txtQtde2.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtQtde2_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus txtQtde2

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtQtde2_LostFocus()
On Error GoTo tratar_erro

txtQtde2.Text = Format(txtQtde2.Text, "###,##0.0000")
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtTexto_cad_Change()
On Error GoTo tratar_erro

Lista_cadastrados.ListItems.Clear
ProcLimpaCampos
If txtTexto_cad <> "" Then cmbTexto_cad.ListIndex = -1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaLista3()
On Error GoTo tratar_erro

Lista_Ncadastrados.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select CPL.* from compras_pedido_lista CPL LEFT JOIN projproduto P ON P.Desenho = CPL.Desenho where CPL.id_cotacao = " & Cont & " and P.Desenho IS NULL", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    With Lista_Ncadastrados.ListItems
        Do While TBLISTA.EOF = False
            .Add , , TBLISTA!IDlista
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Desenho), "", TBLISTA!Desenho)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Descricao), "", TBLISTA!Descricao)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Un), "", TBLISTA!Un)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!Unidade_com), "", TBLISTA!Unidade_com)
            .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!Familia), "", TBLISTA!Familia)
            TBLISTA.MoveNext
            Contador = Contador + 1
            PBLista.Value = Contador
        Loop
    End With
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtTexto_necess_Change()
On Error GoTo tratar_erro

ListaNecessidade.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtTexto_sol_Change()
On Error GoTo tratar_erro

Lista_solicitados.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtun_Change()
On Error GoTo tratar_erro

With txtQtde
    .Locked = False
    .TabStop = True
End With
With txtQtde_PC
    .Locked = True
    .TabStop = False
End With
If txtUN = "KG" And FunVerifMovimentacaoEstPC(frmcompras_reqcot.Cmb_empresa.ItemData(frmcompras_reqcot.Cmb_empresa.ListIndex)) = True Then
    With txtQtde
        .Locked = True
        .TabStop = False
    End With
    With txtQtde_PC
        .Locked = False
        .TabStop = True
    End With
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcFiltrar_Solicitacao
    Case 2: ProcAdicionar_Solicitados
    'Case 4: ProcAjuda
    Case 5: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar2_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcFiltrar_Cadastrados
    Case 2: procAdicionar_Cadastrados
    'Case 4: ProcAjuda
    Case 5: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar3_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcNovo
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

Private Sub ProcAdicionar_Necessidade()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
Permitido1 = False
With ListaNecessidade
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente adicionar este(s) produto(s)/serviço(s) nesta cotação?", vbYesNo, "CAPRIND v5.0") = vbYes Then
                    If USMsgBox("Algum produto/serviço selecionado será adicionado com quantidade parcial?", vbYesNo, "CAPRIND v5.0") = vbYes Then Permitido1 = True
                Else
                    Exit Sub
                End If
            End If
            Permitido = True
            IDlista = .ListItems.Item(InitFor)
            Desenho = .ListItems(InitFor).SubItems(1)
            If Permitido1 = True Then
                Compras_Pedido = False
                Vendas_PI = False
                Compras_Cotacao = True
                Faturamento = False
                Qtde = .ListItems(InitFor).SubItems(4)
                Permitido2 = True
                Sit_REG = 2
                frmVendas_PI_liberaritem.Show 1
                If Permitido2 = False Then Exit Sub
            Else
                valor = .ListItems(InitFor).SubItems(4)
                frmcompras_reqcot.ProcNovo_Necessidade Opt_vendas
            End If
        End If
    Next InitFor
End With

If Permitido = False Then
    USMsgBox ("Informe o(s) produto(s)/serviço(s) antes de adicionar na cotação."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Produto(s)/serviço(s) adicionado(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    With frmcompras_reqcot
        .ProcCarregaListaItens
        .ProcCarregaListaItens1
    End With
    ProcCarregalista_Necessidade
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar4_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcFiltrar_Necessidade
    Case 2: ProcAdicionar_Necessidade
    'Case 4: ProcAjuda
    Case 5: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
