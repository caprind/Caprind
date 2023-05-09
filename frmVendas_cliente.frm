VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVendas_cliente 
   Caption         =   "Administrativo - Vendas - Clientes"
   ClientHeight    =   10035
   ClientLeft      =   1965
   ClientTop       =   1755
   ClientWidth     =   15360
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   7.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmVendas_cliente.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10035
   ScaleWidth      =   15360
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ListView Lista 
      Height          =   4890
      Left            =   75
      TabIndex        =   37
      Top             =   4500
      Width           =   15225
      _ExtentX        =   26855
      _ExtentY        =   8625
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
         Object.Width           =   512
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Object.Tag             =   "D"
         Text            =   "Data"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Responsável"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "Razão social"
         Object.Width           =   19015
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Validado"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Tag             =   "N"
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Frame Frame15 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   75
      TabIndex        =   236
      Top             =   9360
      Width           =   15225
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
         ItemData        =   "frmVendas_cliente.frx":014A
         Left            =   6960
         List            =   "frmVendas_cliente.frx":0157
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   187
         Width           =   1965
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
         TabIndex        =   38
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
         Left            =   9540
         TabIndex        =   40
         ToolTipText     =   "Número da página."
         Top             =   180
         Width           =   555
      End
      Begin DrawSuite2022.USButton cmdPagProx 
         Height          =   315
         Left            =   11760
         TabIndex        =   44
         ToolTipText     =   "Próxima página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmVendas_cliente.frx":0177
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
         TabIndex        =   43
         ToolTipText     =   "Página anterior."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmVendas_cliente.frx":391B
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
         TabIndex        =   41
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
         TabIndex        =   42
         ToolTipText     =   "Primeira página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmVendas_cliente.frx":7424
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
         TabIndex        =   45
         ToolTipText     =   "Última página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmVendas_cliente.frx":B513
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
      Begin VB.Label Label15 
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
         TabIndex        =   283
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
         TabIndex        =   244
         Top             =   240
         Width           =   1260
      End
      Begin VB.Label Label14 
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
         TabIndex        =   243
         Top             =   240
         Width           =   645
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
         TabIndex        =   238
         Top             =   240
         Width           =   1275
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
         TabIndex        =   237
         Top             =   240
         Width           =   1095
      End
   End
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
   Begin DrawSuite2022.USToolBar USToolBar2 
      Height          =   975
      Left            =   75
      TabIndex        =   240
      Top             =   360
      Visible         =   0   'False
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   1720
      ButtonCount     =   12
      GradientColor1  =   16777215
      GradientColor2  =   14737632
      GradientColorDown1=   10802943
      GradientColorDown2=   7979263
      GradientColorDownRight1=   10802943
      GradientColorDownRight2=   7979263
      GradientColorOver1=   14417407
      GradientColorOver2=   12317439
      GradientColorOverRight1=   14417407
      GradientColorOverRight2=   12317439
      IsStrech        =   -1  'True
      RightColor1     =   14737632
      RightColor2     =   16777215
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
      ButtonCaption4  =   "Relatório"
      ButtonEnabled4  =   0   'False
      ButtonIconSize4 =   32
      ButtonToolTipText4=   "Relatório (F5)"
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
      ButtonLeft4     =   133
      ButtonTop4      =   2
      ButtonWidth4    =   60
      ButtonHeight4   =   21
      ButtonUseMaskColor4=   0   'False
      ButtonCaption5  =   "Anterior"
      ButtonEnabled5  =   0   'False
      ButtonIconSize5 =   32
      ButtonToolTipText5=   "Registro anterior."
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
      ButtonLeft5     =   195
      ButtonTop5      =   2
      ButtonWidth5    =   55
      ButtonHeight5   =   21
      ButtonUseMaskColor5=   0   'False
      ButtonCaption6  =   "Próximo"
      ButtonEnabled6  =   0   'False
      ButtonIconSize6 =   32
      ButtonToolTipText6=   "Próximo registro."
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
      ButtonLeft6     =   252
      ButtonTop6      =   2
      ButtonWidth6    =   55
      ButtonHeight6   =   21
      ButtonUseMaskColor6=   0   'False
      ButtonCaption7  =   "Atualizar"
      ButtonEnabled7  =   0   'False
      ButtonIconSize7 =   32
      ButtonToolTipText7=   "Utilizado pelo administrador do sistema."
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
      ButtonLeft7     =   309
      ButtonTop7      =   2
      ButtonWidth7    =   59
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
      ButtonLeft8     =   370
      ButtonTop8      =   4
      ButtonWidth8    =   2
      ButtonHeight8   =   54
      ButtonCaption9  =   "Ajuda"
      ButtonEnabled9  =   0   'False
      ButtonIconSize9 =   32
      ButtonToolTipText9=   "Ajuda (F1)"
      ButtonKey9      =   "9"
      ButtonAlignment9=   2
      BeginProperty ButtonFont9 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft9     =   374
      ButtonTop9      =   2
      ButtonWidth9    =   41
      ButtonHeight9   =   21
      ButtonUseMaskColor9=   0   'False
      ButtonCaption10 =   "Sair"
      ButtonEnabled10 =   0   'False
      ButtonIconSize10=   32
      ButtonToolTipText10=   "Sair (Esc)"
      ButtonKey10     =   "10"
      ButtonAlignment10=   2
      BeginProperty ButtonFont10 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft10    =   417
      ButtonTop10     =   2
      ButtonWidth10   =   30
      ButtonHeight10  =   21
      ButtonUseMaskColor10=   0   'False
      ButtonEnabled11 =   0   'False
      ButtonIconSize11=   32
      ButtonKey11     =   "11"
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
      ButtonLeft11    =   449
      ButtonTop11     =   2
      ButtonWidth11   =   24
      ButtonHeight11  =   24
      ButtonUseMaskColor11=   0   'False
      ButtonEnabled12 =   0   'False
      BeginProperty ButtonFont12 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft12    =   475
      ButtonTop12     =   2
      ButtonWidth12   =   24
      ButtonHeight12  =   24
      Begin DrawSuite2022.USImageList USImageList2 
         Left            =   12930
         Top             =   180
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmVendas_cliente.frx":ED9F
         Count           =   1
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   10035
      Left            =   0
      TabIndex        =   149
      Top             =   0
      Width           =   15360
      _ExtentX        =   27093
      _ExtentY        =   17701
      _Version        =   393216
      Tabs            =   9
      TabsPerRow      =   9
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
      TabCaption(0)   =   "Cliente"
      TabPicture(0)   =   "frmVendas_cliente.frx":14E83
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtidcliente"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "USToolBar1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Contatos"
      TabPicture(1)   =   "frmVendas_cliente.frx":14E9F
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Lista_contato"
      Tab(1).Control(1)=   "Frame2"
      Tab(1).Control(2)=   "txtIDContato"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Entrega"
      TabPicture(2)   =   "frmVendas_cliente.frx":14EBB
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "ListaEntrega"
      Tab(2).Control(1)=   "Frame6"
      Tab(2).Control(2)=   "txtid_entrega"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Cobrança"
      TabPicture(3)   =   "frmVendas_cliente.frx":14ED7
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txtid_cobranca"
      Tab(3).Control(1)=   "Frame8"
      Tab(3).Control(2)=   "listacobranca"
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "Famílias"
      TabPicture(4)   =   "frmVendas_cliente.frx":14EF3
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "txtid_familia"
      Tab(4).Control(1)=   "lista_familia"
      Tab(4).Control(2)=   "Frame12"
      Tab(4).ControlCount=   3
      TabCaption(5)   =   "Dados bancário"
      TabPicture(5)   =   "frmVendas_cliente.frx":14F0F
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "lista_banco"
      Tab(5).Control(1)=   "Frame10"
      Tab(5).Control(2)=   "txtID_banco"
      Tab(5).ControlCount=   3
      TabCaption(6)   =   "Impostos"
      TabPicture(6)   =   "frmVendas_cliente.frx":14F2B
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Lista_Impostos"
      Tab(6).Control(1)=   "txtid_impostos"
      Tab(6).Control(2)=   "Frame14"
      Tab(6).ControlCount=   3
      TabCaption(7)   =   "Comercial"
      TabPicture(7)   =   "frmVendas_cliente.frx":14F47
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "USToolBar3"
      Tab(7).Control(1)=   "Frame16"
      Tab(7).ControlCount=   2
      TabCaption(8)   =   "Outros"
      TabPicture(8)   =   "frmVendas_cliente.frx":14F63
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "USToolBar4"
      Tab(8).Control(1)=   "Frame19"
      Tab(8).ControlCount=   2
      Begin DrawSuite2022.USToolBar USToolBar1 
         Height          =   990
         Left            =   75
         TabIndex        =   239
         Top             =   330
         Width           =   15240
         _ExtentX        =   26882
         _ExtentY        =   1746
         ButtonCount     =   16
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
         ButtonCaption8  =   "Filtrar todos"
         ButtonEnabled8  =   0   'False
         ButtonIconSize8 =   32
         ButtonToolTipText8=   "Filtrar todos os clientes cadastrados."
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
         ButtonWidth8    =   66
         ButtonHeight8   =   21
         ButtonUseMaskColor8=   0   'False
         ButtonCaption9  =   "Status"
         ButtonEnabled9  =   0   'False
         ButtonIconSize9 =   32
         ButtonToolTipText9=   "Status (F7)"
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
         ButtonLeft9     =   415
         ButtonTop9      =   2
         ButtonWidth9    =   39
         ButtonHeight9   =   21
         ButtonUseMaskColor9=   0   'False
         ButtonCaption10 =   "Validação"
         ButtonEnabled10 =   0   'False
         ButtonIconSize10=   32
         ButtonToolTipText10=   "Validar/Cancelar validação (F10)"
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
         ButtonLeft10    =   456
         ButtonTop10     =   2
         ButtonWidth10   =   53
         ButtonHeight10  =   21
         ButtonUseMaskColor10=   0   'False
         ButtonCaption11 =   "Atualizar"
         ButtonEnabled11 =   0   'False
         ButtonIconSize11=   32
         ButtonToolTipText11=   "Utilizado pelo administrador do sistema."
         ButtonKey11     =   "10"
         ButtonAlignment11=   2
         BeginProperty ButtonFont11 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft11    =   511
         ButtonTop11     =   2
         ButtonWidth11   =   50
         ButtonHeight11  =   21
         ButtonUseMaskColor11=   0   'False
         ButtonCaption12 =   "Sincronizar"
         ButtonEnabled12 =   0   'False
         ButtonIconSize12=   32
         ButtonToolTipText12=   "Sincronizar clientes na nuvem"
         ButtonKey12     =   "11"
         ButtonAlignment12=   2
         BeginProperty ButtonFont12 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft12    =   563
         ButtonTop12     =   2
         ButtonWidth12   =   72
         ButtonHeight12  =   24
         ButtonUseMaskColor12=   0   'False
         ButtonEnabled13 =   0   'False
         ButtonIconSize13=   32
         ButtonAlignment13=   2
         ButtonType13    =   1
         ButtonStyle13   =   -1
         BeginProperty ButtonFont13 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState13   =   -1
         ButtonLeft13    =   637
         ButtonTop13     =   4
         ButtonWidth13   =   2
         ButtonHeight13  =   55
         ButtonCaption14 =   "Ajuda"
         ButtonEnabled14 =   0   'False
         ButtonIconSize14=   32
         ButtonToolTipText14=   "Ajuda (F1)"
         ButtonKey14     =   "13"
         ButtonAlignment14=   2
         BeginProperty ButtonFont14 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft14    =   641
         ButtonTop14     =   2
         ButtonWidth14   =   41
         ButtonHeight14  =   21
         ButtonUseMaskColor14=   0   'False
         ButtonCaption15 =   "Sair"
         ButtonEnabled15 =   0   'False
         ButtonIconSize15=   32
         ButtonToolTipText15=   "Sair (Esc)"
         ButtonKey15     =   "14"
         ButtonAlignment15=   2
         BeginProperty ButtonFont15 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft15    =   684
         ButtonTop15     =   2
         ButtonWidth15   =   30
         ButtonHeight15  =   21
         ButtonUseMaskColor15=   0   'False
         ButtonEnabled16 =   0   'False
         ButtonKey16     =   "15"
         BeginProperty ButtonFont16 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState16   =   5
         ButtonLeft16    =   716
         ButtonTop16     =   2
         ButtonWidth16   =   24
         ButtonHeight16  =   24
         ButtonUseMaskColor16=   0   'False
         Begin VB.CheckBox Chk_nao_contribuinte_ICMS 
            BackColor       =   &H80000004&
            Caption         =   "Não contribuinte ICMS"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   13410
            TabIndex        =   299
            Top             =   750
            Width           =   1755
         End
         Begin VB.CheckBox Chk_enviar_NF 
            BackColor       =   &H80000004&
            Caption         =   "Enviar NF"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   12300
            TabIndex        =   298
            Top             =   750
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.CheckBox Chk_prospecto 
            BackColor       =   &H80000004&
            Caption         =   "Prospecto"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   13410
            TabIndex        =   297
            Top             =   390
            Width           =   1035
         End
         Begin DrawSuite2022.USImageList USImageList1 
            Left            =   10740
            Top             =   90
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmVendas_cliente.frx":14F7F
            Count           =   1
         End
      End
      Begin VB.TextBox txtidcliente 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   225
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   33
         ToolTipText     =   "Código do cliente."
         Top             =   2880
         Width           =   675
      End
      Begin VB.Frame Frame19 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2115
         Left            =   -74925
         TabIndex        =   211
         Top             =   1285
         Width           =   15225
         Begin VB.TextBox txtLimiteCredito 
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
            MaxLength       =   50
            TabIndex        =   292
            TabStop         =   0   'False
            ToolTipText     =   "Código suframa."
            Top             =   1620
            Width           =   3120
         End
         Begin VB.CheckBox chkICMSST 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Utiliza cálculo simplificado de ICMS ST"
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
            Left            =   12030
            TabIndex        =   140
            Top             =   1095
            Width           =   3135
         End
         Begin VB.CommandButton Cmd_localizar_tipo_dcto 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   1260
            Picture         =   "frmVendas_cliente.frx":1E32B
            Style           =   1  'Graphical
            TabIndex        =   143
            ToolTipText     =   "Localizar tipo do documento."
            Top             =   1620
            Width           =   315
         End
         Begin VB.ComboBox cmbTipo_doc 
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
            ItemData        =   "frmVendas_cliente.frx":1E42D
            Left            =   180
            List            =   "frmVendas_cliente.frx":1E42F
            Style           =   2  'Dropdown List
            TabIndex        =   142
            ToolTipText     =   "Tipo do documento previsto para recebimento."
            Top             =   1620
            Width           =   1065
         End
         Begin VB.ComboBox cmbBanco 
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
            ItemData        =   "frmVendas_cliente.frx":1E431
            Left            =   1680
            List            =   "frmVendas_cliente.frx":1E433
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   144
            ToolTipText     =   "Instituição bancária prevista para recebimento."
            Top             =   1620
            Width           =   3525
         End
         Begin VB.CommandButton Cmd_limpar_grupo 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   11520
            Picture         =   "frmVendas_cliente.frx":1E435
            Style           =   1  'Graphical
            TabIndex        =   148
            ToolTipText     =   "Limpar grupo."
            Top             =   1620
            Width           =   315
         End
         Begin VB.CheckBox chkSuframa 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Código Suframa"
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
            Height          =   225
            Left            =   10065
            TabIndex        =   138
            Top             =   795
            Width           =   1515
         End
         Begin VB.TextBox txtSuframa 
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
            Left            =   9600
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   139
            TabStop         =   0   'False
            ToolTipText     =   "Código suframa."
            Top             =   1035
            Width           =   2370
         End
         Begin VB.CommandButton cmdGrupo 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   11190
            Picture         =   "frmVendas_cliente.frx":1E573
            Style           =   1  'Graphical
            TabIndex        =   147
            ToolTipText     =   "Localizar grupo."
            Top             =   1620
            Width           =   315
         End
         Begin VB.TextBox txtIDGrupo 
            BackColor       =   &H00FFFFFF&
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
            Left            =   8550
            MaxLength       =   60
            TabIndex        =   145
            Text            =   "0"
            Top             =   1620
            Visible         =   0   'False
            Width           =   950
         End
         Begin VB.TextBox Txt_ISSQN 
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
            Left            =   13755
            MaxLength       =   50
            TabIndex        =   141
            ToolTipText     =   "Alíquota de ISSQN."
            Top             =   390
            Width           =   1260
         End
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
            Height          =   960
            Left            =   180
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   134
            ToolTipText     =   "Observações."
            Top             =   390
            Width           =   9315
         End
         Begin VB.TextBox txtGrupo 
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
            Left            =   5220
            Locked          =   -1  'True
            TabIndex        =   146
            TabStop         =   0   'False
            ToolTipText     =   "Grupo."
            Top             =   1620
            Width           =   5955
         End
         Begin MSMask.MaskEdBox txttel02 
            Height          =   315
            Left            =   9600
            TabIndex        =   135
            ToolTipText     =   "Número do telefone."
            Top             =   390
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   0
            AllowPrompt     =   -1  'True
            AutoTab         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txttel03 
            Height          =   315
            Left            =   10985
            TabIndex        =   136
            ToolTipText     =   "Número do telefone."
            Top             =   390
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   0
            AllowPrompt     =   -1  'True
            AutoTab         =   -1  'True
            MaxLength       =   30
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txttel04 
            Height          =   315
            Left            =   12370
            TabIndex        =   137
            ToolTipText     =   "Número do telefone."
            Top             =   390
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   0
            AllowPrompt     =   -1  'True
            AutoTab         =   -1  'True
            MaxLength       =   30
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Limite de crédito"
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
            Index           =   58
            Left            =   12870
            TabIndex        =   293
            Top             =   1410
            Width           =   1170
         End
         Begin VB.Label Label13 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo docto."
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
            TabIndex        =   280
            Top             =   1410
            Width           =   885
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Instituição bancária prevista para recebimento"
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
            Index           =   22
            Left            =   1755
            TabIndex        =   279
            Top             =   1410
            Width           =   3345
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Tel.03"
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
            Index           =   12
            Left            =   12827
            TabIndex        =   217
            Top             =   180
            Width           =   450
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Tel.02"
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
            Index           =   11
            Left            =   11442
            TabIndex        =   216
            Top             =   180
            Width           =   450
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Tel.01"
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
            Index           =   10
            Left            =   10057
            TabIndex        =   215
            Top             =   180
            Width           =   450
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Grupo do cliente"
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
            Index           =   43
            Left            =   7605
            TabIndex        =   214
            Top             =   1410
            Width           =   1170
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "ISSQN (%)"
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
            Index           =   41
            Left            =   13988
            TabIndex        =   213
            Top             =   180
            Width           =   795
         End
         Begin VB.Label Label3 
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
            Index           =   45
            Left            =   4365
            TabIndex        =   212
            Top             =   180
            Width           =   945
         End
      End
      Begin VB.Frame Frame16 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   8685
         Left            =   -74925
         TabIndex        =   197
         Top             =   1285
         Width           =   15195
         Begin VB.CommandButton Cmd_limpar_CFOP 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   14700
            Picture         =   "frmVendas_cliente.frx":1E675
            Style           =   1  'Graphical
            TabIndex        =   119
            ToolTipText     =   "Limpar grupo."
            Top             =   240
            Width           =   315
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
            ItemData        =   "frmVendas_cliente.frx":1E7B3
            Left            =   2235
            List            =   "frmVendas_cliente.frx":1E7B5
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   115
            ToolTipText     =   "Empresa."
            Top             =   240
            Width           =   4245
         End
         Begin VB.CommandButton cmdcfop 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   14370
            Picture         =   "frmVendas_cliente.frx":1E7B7
            Style           =   1  'Graphical
            TabIndex        =   118
            ToolTipText     =   "Localizar CFOP."
            Top             =   240
            Width           =   315
         End
         Begin VB.TextBox txtoperacao 
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
            Left            =   8265
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   117
            TabStop         =   0   'False
            ToolTipText     =   "Descrição da natureza da operação."
            Top             =   240
            Width           =   6105
         End
         Begin VB.TextBox txtID_cfop 
            Alignment       =   2  'Center
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
            Height          =   315
            Left            =   7770
            MaxLength       =   100
            TabIndex        =   204
            ToolTipText     =   "Data da revisão."
            Top             =   240
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.CommandButton cmdCond_pag_padrao 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1130
            Left            =   14700
            Picture         =   "frmVendas_cliente.frx":1E8B9
            Style           =   1  'Graphical
            TabIndex        =   121
            ToolTipText     =   "Localizar cond. de pagamento."
            Top             =   660
            Width           =   315
         End
         Begin VB.CommandButton cmdValidade_Padrao 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1130
            Left            =   14700
            Picture         =   "frmVendas_cliente.frx":1E9BB
            Style           =   1  'Graphical
            TabIndex        =   133
            ToolTipText     =   "Localizar validade."
            Top             =   7435
            Width           =   315
         End
         Begin VB.CommandButton cmdReajuste_padrao 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1130
            Left            =   14700
            Picture         =   "frmVendas_cliente.frx":1EABD
            Style           =   1  'Graphical
            TabIndex        =   131
            ToolTipText     =   "Localizar reajuste."
            Top             =   6295
            Width           =   315
         End
         Begin VB.CommandButton cmdGarantia_padrao 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1130
            Left            =   14700
            Picture         =   "frmVendas_cliente.frx":1EBBF
            Style           =   1  'Graphical
            TabIndex        =   129
            ToolTipText     =   "Localizar garantia."
            Top             =   5156
            Width           =   315
         End
         Begin VB.CommandButton cmdImpostos_padrao 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1130
            Left            =   14700
            Picture         =   "frmVendas_cliente.frx":1ECC1
            Style           =   1  'Graphical
            TabIndex        =   127
            ToolTipText     =   "Localizar impostos."
            Top             =   4017
            Width           =   315
         End
         Begin VB.CommandButton cmdTransporte_padrao 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1130
            Left            =   14700
            Picture         =   "frmVendas_cliente.frx":1EDC3
            Style           =   1  'Graphical
            TabIndex        =   125
            ToolTipText     =   "Localizar transporte."
            Top             =   2878
            Width           =   315
         End
         Begin VB.CommandButton cmdDesenhos_calculos_padrao 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   885
            Left            =   14700
            Picture         =   "frmVendas_cliente.frx":1EEC5
            Style           =   1  'Graphical
            TabIndex        =   123
            ToolTipText     =   "Localizar desenhos e cálculos."
            Top             =   1890
            Width           =   315
         End
         Begin VB.TextBox txtgarantia 
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
            Height          =   1130
            Left            =   2235
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   128
            ToolTipText     =   "Garantia."
            Top             =   5156
            Width           =   12465
         End
         Begin VB.TextBox txttransporte 
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
            Height          =   1130
            Left            =   2235
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   124
            ToolTipText     =   "Transporte."
            Top             =   2878
            Width           =   12465
         End
         Begin VB.TextBox txtcalculos 
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
            Height          =   885
            Left            =   2235
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   122
            ToolTipText     =   "Desenhos e cálculos."
            Top             =   1890
            Width           =   12465
         End
         Begin VB.TextBox txtCondicoes 
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
            Height          =   1095
            Left            =   2235
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   120
            ToolTipText     =   "Condições de pagamento."
            Top             =   690
            Width           =   12465
         End
         Begin VB.TextBox txtReajuste 
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
            Height          =   1130
            Left            =   2235
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   130
            ToolTipText     =   "Reajuste dos preços."
            Top             =   6295
            Width           =   12465
         End
         Begin VB.TextBox txtValidade 
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
            Height          =   1130
            Left            =   2235
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   132
            ToolTipText     =   "Prazo de validade da proposta."
            Top             =   7435
            Width           =   12465
         End
         Begin VB.TextBox txtimpostos 
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
            Height          =   1130
            Left            =   2235
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   126
            ToolTipText     =   "Impostos."
            Top             =   4017
            Width           =   12465
         End
         Begin VB.TextBox txtCFOP 
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
            Left            =   7185
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   116
            TabStop         =   0   'False
            ToolTipText     =   "Natureza da operação."
            Top             =   240
            Width           =   1065
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Empresa:"
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
            Left            =   1380
            TabIndex        =   278
            Top             =   270
            Width           =   780
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Garantia :"
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
            TabIndex        =   206
            Top             =   5156
            Width           =   720
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "CFOP :"
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
            Left            =   6600
            TabIndex        =   205
            Top             =   240
            Width           =   510
         End
         Begin VB.Label Label38 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Transporte :"
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
            Left            =   1260
            TabIndex        =   203
            Top             =   2878
            Width           =   900
         End
         Begin VB.Label Label35 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Desenhos e calculos :"
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
            Height          =   345
            Left            =   615
            TabIndex        =   202
            Top             =   1890
            Width           =   1545
         End
         Begin VB.Label Label30 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Condições de pagamento :"
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
            TabIndex        =   201
            Top             =   660
            Width           =   1920
         End
         Begin VB.Label Label29 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Reajuste :"
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
            TabIndex        =   200
            Top             =   6295
            Width           =   750
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Validade :"
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
            Left            =   1455
            TabIndex        =   199
            Top             =   7435
            Width           =   705
         End
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Impostos :"
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
            Left            =   1395
            TabIndex        =   198
            Top             =   4017
            Width           =   765
         End
      End
      Begin VB.Frame Frame14 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   -74925
         TabIndex        =   193
         Top             =   1285
         Width           =   15195
         Begin VB.TextBox Txt_ID_CF 
            Alignment       =   2  'Center
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
            Height          =   315
            Left            =   4110
            TabIndex        =   109
            TabStop         =   0   'False
            ToolTipText     =   "ID CF."
            Top             =   390
            Visible         =   0   'False
            Width           =   525
         End
         Begin VB.TextBox TxtResponsavel6 
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
            Left            =   1335
            Locked          =   -1  'True
            TabIndex        =   108
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pelo cadastro."
            Top             =   390
            Width           =   2765
         End
         Begin VB.TextBox txtdata6 
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
            MaxLength       =   25
            TabIndex        =   107
            TabStop         =   0   'False
            ToolTipText     =   "Data do cadastro."
            Top             =   390
            Width           =   1140
         End
         Begin VB.TextBox txtIPI 
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
            Left            =   6390
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   112
            TabStop         =   0   'False
            ToolTipText     =   "Aliquota de IPI."
            Top             =   375
            Width           =   705
         End
         Begin VB.CommandButton cmdCF 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   5280
            Picture         =   "frmVendas_cliente.frx":1EFC7
            Style           =   1  'Graphical
            TabIndex        =   111
            ToolTipText     =   "Abrir módulo para consulta de classificação fiscal."
            Top             =   390
            Width           =   315
         End
         Begin VB.TextBox txtPorcentagemIPI 
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
            Left            =   10260
            MaxLength       =   50
            TabIndex        =   113
            ToolTipText     =   "Valor para cálculo do IPI."
            Top             =   375
            Width           =   795
         End
         Begin VB.TextBox Txt_CF 
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
            Left            =   4110
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   110
            TabStop         =   0   'False
            ToolTipText     =   "Classificação fiscal."
            Top             =   390
            Width           =   1155
         End
         Begin VB.Label Label2 
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
            Index           =   19
            Left            =   2260
            TabIndex        =   229
            Top             =   180
            Width           =   915
         End
         Begin VB.Label Label2 
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
            Index           =   18
            Left            =   578
            TabIndex        =   228
            Top             =   180
            Width           =   345
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "IPI (%) :"
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
            Left            =   5700
            MousePointer    =   1  'Arrow
            TabIndex        =   196
            Top             =   435
            Width           =   645
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "NCM"
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
            Left            =   4522
            TabIndex        =   195
            Top             =   180
            Width           =   330
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Regra para cálculo : Vlr. unitário c/ desc. /                     = Vlr. total c/ IPI - Vlr. unitário c/ desc. = Vlr. total IPI"
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
            Left            =   7170
            TabIndex        =   194
            Top             =   435
            Width           =   7860
         End
      End
      Begin VB.TextBox txtid_impostos 
         Height          =   315
         Left            =   -74010
         TabIndex        =   192
         Text            =   "0"
         Top             =   4350
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.TextBox txtid_familia 
         Height          =   315
         Left            =   -74295
         TabIndex        =   189
         Text            =   "0"
         Top             =   4230
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.TextBox txtID_banco 
         Height          =   315
         Left            =   -72345
         TabIndex        =   188
         Text            =   "0"
         Top             =   4140
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   -74925
         TabIndex        =   184
         Top             =   1285
         Width           =   15195
         Begin VB.TextBox TxtResponsavel5 
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
            Left            =   1395
            Locked          =   -1  'True
            TabIndex        =   102
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pelo cadastro."
            Top             =   390
            Width           =   3255
         End
         Begin VB.TextBox txtdata5 
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
            MaxLength       =   25
            TabIndex        =   101
            TabStop         =   0   'False
            ToolTipText     =   "Data do cadastro."
            Top             =   390
            Width           =   1185
         End
         Begin VB.TextBox txtBanco 
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
            Left            =   4680
            MaxLength       =   255
            TabIndex        =   103
            ToolTipText     =   "Banco recebedor."
            Top             =   390
            Width           =   5025
         End
         Begin VB.TextBox txtAgencia 
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
            Left            =   9720
            MaxLength       =   50
            TabIndex        =   104
            ToolTipText     =   "Número da agência."
            Top             =   390
            Width           =   2805
         End
         Begin VB.TextBox txtConta 
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
            Left            =   12540
            MaxLength       =   20
            TabIndex        =   105
            ToolTipText     =   "Número da conta."
            Top             =   390
            Width           =   2475
         End
         Begin VB.Label Label2 
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
            Index           =   17
            Left            =   2550
            TabIndex        =   227
            Top             =   180
            Width           =   915
         End
         Begin VB.Label Label2 
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
            Index           =   16
            Left            =   600
            TabIndex        =   226
            Top             =   180
            Width           =   345
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
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
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   7
            Left            =   6960
            TabIndex        =   187
            Top             =   180
            Width           =   435
         End
         Begin VB.Label Label20 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Conta"
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
            Left            =   13560
            TabIndex        =   186
            Top             =   180
            Width           =   435
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Agência"
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
            Index           =   8
            Left            =   10837
            TabIndex        =   185
            Top             =   180
            Width           =   570
         End
      End
      Begin VB.TextBox txtid_entrega 
         Height          =   315
         Left            =   -72900
         TabIndex        =   182
         Text            =   "0"
         Top             =   6030
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.TextBox txtIDContato 
         BackColor       =   &H00FFFFFF&
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
         Left            =   -71880
         MaxLength       =   60
         MouseIcon       =   "frmVendas_cliente.frx":1F0C9
         MousePointer    =   99  'Custom
         TabIndex        =   181
         Text            =   "0"
         ToolTipText     =   "Digite o nome para contato."
         Top             =   6420
         Visible         =   0   'False
         Width           =   950
      End
      Begin VB.TextBox txtid_cobranca 
         Height          =   315
         Left            =   -74280
         TabIndex        =   180
         Text            =   "0"
         Top             =   6420
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2085
         Left            =   -74925
         TabIndex        =   154
         Top             =   1285
         Width           =   15195
         Begin VB.TextBox txtemail_entrega 
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
            Left            =   5640
            MaxLength       =   60
            TabIndex        =   74
            ToolTipText     =   "E-mail."
            Top             =   1620
            Width           =   4695
         End
         Begin VB.TextBox txtsite_entrega 
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
            Left            =   10350
            MaxLength       =   60
            TabIndex        =   75
            ToolTipText     =   "Site."
            Top             =   1620
            Width           =   4665
         End
         Begin VB.TextBox txtComplemento_entrega 
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
            Left            =   13890
            MaxLength       =   30
            TabIndex        =   61
            ToolTipText     =   "Complemento."
            Top             =   390
            Width           =   1125
         End
         Begin VB.ComboBox cmbTipo_bairro_entrega 
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
            ItemData        =   "frmVendas_cliente.frx":1F3D3
            Left            =   180
            List            =   "frmVendas_cliente.frx":1F416
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   62
            ToolTipText     =   "Tipo do bairro."
            Top             =   990
            Width           =   1305
         End
         Begin VB.ComboBox cmbTipo_endereco_entrega 
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
            ItemData        =   "frmVendas_cliente.frx":1F4CD
            Left            =   5946
            List            =   "frmVendas_cliente.frx":1F507
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   58
            ToolTipText     =   "Tipo do endereço."
            Top             =   390
            Width           =   1260
         End
         Begin VB.TextBox TxtResponsavel2 
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
            Left            =   1382
            Locked          =   -1  'True
            TabIndex        =   56
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pelo cadastro."
            Top             =   390
            Width           =   2835
         End
         Begin VB.TextBox txtdata2 
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
            MaxLength       =   25
            TabIndex        =   55
            TabStop         =   0   'False
            ToolTipText     =   "Data do cadastro."
            Top             =   390
            Width           =   1185
         End
         Begin VB.TextBox txtNumero_entrega 
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
            Left            =   12880
            MaxLength       =   60
            TabIndex        =   60
            ToolTipText     =   "Número."
            Top             =   390
            Width           =   990
         End
         Begin VB.TextBox mskcep_entrega 
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
            Left            =   10380
            MaxLength       =   12
            TabIndex        =   68
            ToolTipText     =   "CEP."
            Top             =   990
            Width           =   1005
         End
         Begin VB.TextBox txtcxpostal_entrega 
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
            Left            =   9240
            MaxLength       =   50
            TabIndex        =   67
            ToolTipText     =   "Caixa postal."
            Top             =   990
            Width           =   1120
         End
         Begin VB.ComboBox cmbuf_entrega 
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
            ItemData        =   "frmVendas_cliente.frx":1F58F
            Left            =   4920
            List            =   "frmVendas_cliente.frx":1F591
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   64
            ToolTipText     =   "UF."
            Top             =   990
            Width           =   740
         End
         Begin VB.TextBox txtBairro_entrega 
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
            Left            =   1485
            MaxLength       =   60
            TabIndex        =   63
            ToolTipText     =   "Bairro."
            Top             =   990
            Width           =   3420
         End
         Begin VB.TextBox txtEndereco_entrega 
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
            Left            =   7200
            MaxLength       =   60
            TabIndex        =   59
            ToolTipText     =   "Endereço."
            Top             =   390
            Width           =   5670
         End
         Begin MSMask.MaskEdBox txttel2_entrega 
            Height          =   315
            Left            =   180
            TabIndex        =   71
            ToolTipText     =   "Número do telefone."
            Top             =   1620
            Width           =   1800
            _ExtentX        =   3175
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   0
            AllowPrompt     =   -1  'True
            AutoTab         =   -1  'True
            MaxLength       =   14
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txttel3_entrega 
            Height          =   315
            Left            =   1995
            TabIndex        =   72
            ToolTipText     =   "Número do telefone."
            Top             =   1620
            Width           =   1820
            _ExtentX        =   3201
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   0
            AllowPrompt     =   -1  'True
            AutoTab         =   -1  'True
            MaxLength       =   14
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txttel1_entrega 
            Height          =   315
            Left            =   11400
            TabIndex        =   69
            ToolTipText     =   "Número do telefone."
            Top             =   990
            Width           =   1785
            _ExtentX        =   3149
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   0
            AllowPrompt     =   -1  'True
            AutoTab         =   -1  'True
            MaxLength       =   14
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txttel4_entrega 
            Height          =   315
            Left            =   3825
            TabIndex        =   73
            ToolTipText     =   "Número do telefone."
            Top             =   1620
            Width           =   1800
            _ExtentX        =   3175
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   0
            AllowPrompt     =   -1  'True
            AutoTab         =   -1  'True
            MaxLength       =   14
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtfax_entrega 
            Height          =   315
            Left            =   13205
            TabIndex        =   70
            ToolTipText     =   "Número do fax."
            Top             =   990
            Width           =   1810
            _ExtentX        =   3201
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   0
            AllowPrompt     =   -1  'True
            AutoTab         =   -1  'True
            MaxLength       =   14
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtCNPJ_entrega 
            Height          =   315
            Left            =   4230
            TabIndex        =   57
            ToolTipText     =   "Número do CNPJ."
            Top             =   390
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   0
            MaxLength       =   18
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##.###.###/####-##"
            PromptChar      =   "_"
         End
         Begin VB.ComboBox cmbCidade_Entrega 
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
            ItemData        =   "frmVendas_cliente.frx":1F593
            Left            =   5670
            List            =   "frmVendas_cliente.frx":1F595
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   65
            ToolTipText     =   "Cidade."
            Top             =   990
            Width           =   3570
         End
         Begin VB.TextBox txtCidade_Entrega 
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
            MaxLength       =   60
            TabIndex        =   66
            ToolTipText     =   "Cidade."
            Top             =   990
            Visible         =   0   'False
            Width           =   3560
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Site"
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
            Index           =   56
            Left            =   12547
            TabIndex        =   245
            Top             =   1410
            Width           =   270
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Complemento"
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
            Index           =   51
            Left            =   13965
            TabIndex        =   234
            Top             =   180
            Width           =   975
         End
         Begin VB.Label Label3 
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
            Index           =   48
            Left            =   682
            TabIndex        =   231
            Top             =   780
            Width           =   300
         End
         Begin VB.Label Label3 
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
            Index           =   47
            Left            =   6426
            TabIndex        =   230
            Top             =   180
            Width           =   300
         End
         Begin VB.Label Label2 
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
            Index           =   11
            Left            =   2340
            TabIndex        =   221
            Top             =   180
            Width           =   915
         End
         Begin VB.Label Label2 
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
            Index           =   10
            Left            =   600
            TabIndex        =   220
            Top             =   180
            Width           =   345
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Número"
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
            Left            =   13110
            TabIndex        =   208
            Top             =   180
            Width           =   555
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
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
            Height          =   195
            Left            =   4890
            TabIndex        =   207
            Top             =   180
            Width           =   375
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Tel.04"
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
            Index           =   31
            Left            =   4500
            TabIndex        =   166
            Top             =   1410
            Width           =   450
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Tel.03"
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
            Index           =   30
            Left            =   2680
            TabIndex        =   165
            Top             =   1410
            Width           =   450
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Tel.02"
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
            Index           =   29
            Left            =   855
            TabIndex        =   164
            Top             =   1410
            Width           =   450
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Tel.01"
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
            Index           =   28
            Left            =   12075
            TabIndex        =   163
            Top             =   780
            Width           =   450
         End
         Begin VB.Label Label3 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "UF"
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
            Height          =   255
            Index           =   27
            Left            =   5185
            TabIndex        =   162
            Top             =   780
            Width           =   210
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Cidade"
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
            Index           =   26
            Left            =   7208
            TabIndex        =   161
            Top             =   780
            Width           =   495
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Bairro"
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
            Index           =   25
            Left            =   2985
            TabIndex        =   160
            Top             =   780
            Width           =   420
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Endereço"
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
            Index           =   24
            Left            =   9698
            TabIndex        =   159
            Top             =   180
            Width           =   675
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
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
            Index           =   22
            Left            =   7777
            TabIndex        =   158
            Top             =   1410
            Width           =   420
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Fax"
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
            Index           =   12
            Left            =   13975
            TabIndex        =   157
            Top             =   780
            Width           =   270
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "CEP"
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
            Index           =   21
            Left            =   10740
            TabIndex        =   156
            Top             =   780
            Width           =   285
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Cx. postal"
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
            Index           =   20
            Left            =   9435
            TabIndex        =   155
            Top             =   780
            Width           =   735
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   -74925
         TabIndex        =   150
         Top             =   1285
         Width           =   15195
         Begin VB.TextBox TxtEmail_Contato 
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
            Left            =   7320
            MaxLength       =   60
            TabIndex        =   51
            ToolTipText     =   "E-mail."
            Top             =   990
            Width           =   5925
         End
         Begin VB.CheckBox Chk_enviar_boleto 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Recebe o boleto?"
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
            Left            =   13320
            TabIndex        =   53
            Top             =   1110
            Width           =   1605
         End
         Begin VB.CheckBox Chk_enviar_NFe 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Recebe a NFe?"
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
            Left            =   13320
            TabIndex        =   52
            Top             =   840
            Width           =   1425
         End
         Begin VB.TextBox TxtResponsavel1 
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
            Left            =   1380
            Locked          =   -1  'True
            TabIndex        =   47
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pelo cadastro."
            Top             =   390
            Width           =   5175
         End
         Begin VB.TextBox txtdata1 
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
            MaxLength       =   25
            TabIndex        =   46
            TabStop         =   0   'False
            ToolTipText     =   "Data do cadastro."
            Top             =   390
            Width           =   1185
         End
         Begin VB.TextBox txtNomeContato 
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
            Left            =   6570
            MaxLength       =   255
            TabIndex        =   48
            ToolTipText     =   "Nome."
            Top             =   390
            Width           =   8445
         End
         Begin VB.TextBox txttelcontato 
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
            Left            =   3930
            MaxLength       =   30
            TabIndex        =   50
            ToolTipText     =   "Telefones."
            Top             =   990
            Width           =   3375
         End
         Begin VB.TextBox txtdepartamento 
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
            TabIndex        =   49
            ToolTipText     =   "Departamento."
            Top             =   990
            Width           =   3735
         End
         Begin VB.Label Label2 
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
            Index           =   7
            Left            =   3285
            TabIndex        =   219
            Top             =   180
            Width           =   915
         End
         Begin VB.Label Label2 
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
            Index           =   6
            Left            =   600
            TabIndex        =   218
            Top             =   180
            Width           =   345
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Nome"
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
            Left            =   10200
            TabIndex        =   183
            Top             =   180
            Width           =   405
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
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
            Index           =   16
            Left            =   10072
            TabIndex        =   153
            Top             =   780
            Width           =   420
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Telefones"
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
            Index           =   13
            Left            =   5280
            TabIndex        =   152
            Top             =   780
            Width           =   705
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Departamento"
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
            Left            =   1455
            TabIndex        =   151
            Top             =   780
            Width           =   1035
         End
      End
      Begin MSComctlLib.ListView ListaEntrega 
         Height          =   6315
         Left            =   -74925
         TabIndex        =   76
         Top             =   3390
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   11139
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "N"
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Tag             =   "T"
            Text            =   "Endereço"
            Object.Width           =   15046
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Bairro"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Tag             =   "T"
            Text            =   "Cidade"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "UF"
            Object.Width           =   882
         EndProperty
      End
      Begin MSComctlLib.ListView Lista_contato 
         Height          =   6940
         Left            =   -74925
         TabIndex        =   54
         Top             =   2760
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   12250
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "N"
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Tag             =   "T"
            Text            =   "Nome"
            Object.Width           =   12400
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Departamento"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Tag             =   "T"
            Text            =   "Telefones"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "E-mail"
            Object.Width           =   5292
         EndProperty
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2085
         Left            =   -74925
         TabIndex        =   167
         Top             =   1285
         Width           =   15195
         Begin VB.TextBox txtemail_cobranca 
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
            Left            =   5640
            MaxLength       =   60
            TabIndex        =   282
            ToolTipText     =   "E-mail."
            Top             =   1620
            Width           =   4695
         End
         Begin VB.TextBox txtSite_cobranca 
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
            Left            =   10350
            MaxLength       =   60
            TabIndex        =   281
            ToolTipText     =   "Site."
            Top             =   1620
            Width           =   4665
         End
         Begin VB.TextBox txtComplemento_cobranca 
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
            Left            =   13890
            MaxLength       =   30
            TabIndex        =   83
            ToolTipText     =   "Complemento."
            Top             =   390
            Width           =   1125
         End
         Begin VB.ComboBox cmbTipo_bairro_cobranca 
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
            ItemData        =   "frmVendas_cliente.frx":1F597
            Left            =   180
            List            =   "frmVendas_cliente.frx":1F5DA
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   84
            ToolTipText     =   "Tipo do bairro."
            Top             =   990
            Width           =   1305
         End
         Begin VB.ComboBox cmbTipo_endereco_cobranca 
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
            ItemData        =   "frmVendas_cliente.frx":1F691
            Left            =   5946
            List            =   "frmVendas_cliente.frx":1F6CB
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   80
            ToolTipText     =   "Tipo do endereço."
            Top             =   390
            Width           =   1260
         End
         Begin VB.TextBox TxtResponsavel3 
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
            Left            =   1382
            Locked          =   -1  'True
            TabIndex        =   78
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pelo cadastro."
            Top             =   390
            Width           =   2835
         End
         Begin VB.TextBox txtdata3 
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
            MaxLength       =   25
            TabIndex        =   77
            TabStop         =   0   'False
            ToolTipText     =   "Data do cadastro."
            Top             =   390
            Width           =   1185
         End
         Begin VB.TextBox txtNumero_cobranca 
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
            Left            =   12880
            MaxLength       =   60
            TabIndex        =   82
            ToolTipText     =   "Número."
            Top             =   390
            Width           =   990
         End
         Begin VB.TextBox mskcep_cobranca 
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
            Left            =   10380
            MaxLength       =   12
            TabIndex        =   90
            ToolTipText     =   "CEP."
            Top             =   990
            Width           =   1005
         End
         Begin VB.TextBox txtendereco_cobranca 
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
            Left            =   7200
            MaxLength       =   60
            TabIndex        =   81
            ToolTipText     =   "Endereço."
            Top             =   390
            Width           =   5670
         End
         Begin VB.TextBox txtbairro_cobranca 
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
            Left            =   1485
            MaxLength       =   60
            TabIndex        =   85
            ToolTipText     =   "Bairro."
            Top             =   990
            Width           =   3420
         End
         Begin VB.ComboBox cmbuf_cobranca 
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
            ItemData        =   "frmVendas_cliente.frx":1F753
            Left            =   4920
            List            =   "frmVendas_cliente.frx":1F755
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   86
            ToolTipText     =   "UF."
            Top             =   990
            Width           =   740
         End
         Begin VB.TextBox txtcxpostal_cobranca 
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
            Left            =   9240
            MaxLength       =   50
            TabIndex        =   89
            ToolTipText     =   "Caixa postal."
            Top             =   990
            Width           =   1120
         End
         Begin MSMask.MaskEdBox txttel2_cobranca 
            Height          =   315
            Left            =   180
            TabIndex        =   93
            ToolTipText     =   "Número do telefone."
            Top             =   1620
            Width           =   1800
            _ExtentX        =   3175
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   0
            AllowPrompt     =   -1  'True
            AutoTab         =   -1  'True
            MaxLength       =   14
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txttel3_cobranca 
            Height          =   315
            Left            =   1995
            TabIndex        =   94
            ToolTipText     =   "Número do telefone."
            Top             =   1620
            Width           =   1820
            _ExtentX        =   3201
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   0
            AllowPrompt     =   -1  'True
            AutoTab         =   -1  'True
            MaxLength       =   14
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txttel1_cobranca 
            Height          =   315
            Left            =   11400
            TabIndex        =   91
            ToolTipText     =   "Número do telefone."
            Top             =   990
            Width           =   1785
            _ExtentX        =   3149
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   0
            AllowPrompt     =   -1  'True
            AutoTab         =   -1  'True
            MaxLength       =   14
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txttel4_cobranca 
            Height          =   315
            Left            =   3825
            TabIndex        =   95
            ToolTipText     =   "Número do telefone."
            Top             =   1620
            Width           =   1800
            _ExtentX        =   3175
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   0
            AllowPrompt     =   -1  'True
            AutoTab         =   -1  'True
            MaxLength       =   14
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtfax_cobranca 
            Height          =   315
            Left            =   13205
            TabIndex        =   92
            ToolTipText     =   "Número do fax."
            Top             =   990
            Width           =   1810
            _ExtentX        =   3201
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   0
            AllowPrompt     =   -1  'True
            AutoTab         =   -1  'True
            MaxLength       =   14
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtCNPJ_cobranca 
            Height          =   315
            Left            =   4230
            TabIndex        =   79
            ToolTipText     =   "Número do CNPJ."
            Top             =   390
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   0
            MaxLength       =   18
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##.###.###/####-##"
            PromptChar      =   "_"
         End
         Begin VB.ComboBox cmbCidade_cobranca 
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
            ItemData        =   "frmVendas_cliente.frx":1F757
            Left            =   5670
            List            =   "frmVendas_cliente.frx":1F759
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   87
            ToolTipText     =   "Cidade."
            Top             =   990
            Width           =   3570
         End
         Begin VB.TextBox txtcidade_cobranca 
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
            MaxLength       =   60
            TabIndex        =   88
            ToolTipText     =   "Cidade."
            Top             =   990
            Visible         =   0   'False
            Width           =   3560
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Site"
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
            Index           =   57
            Left            =   12547
            TabIndex        =   246
            Top             =   1410
            Width           =   270
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Complemento"
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
            Index           =   52
            Left            =   13965
            TabIndex        =   235
            Top             =   180
            Width           =   975
         End
         Begin VB.Label Label3 
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
            Index           =   50
            Left            =   682
            TabIndex        =   233
            Top             =   780
            Width           =   300
         End
         Begin VB.Label Label3 
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
            Index           =   49
            Left            =   6426
            TabIndex        =   232
            Top             =   180
            Width           =   300
         End
         Begin VB.Label Label2 
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
            Index           =   13
            Left            =   2340
            TabIndex        =   223
            Top             =   180
            Width           =   915
         End
         Begin VB.Label Label2 
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
            Index           =   12
            Left            =   600
            TabIndex        =   222
            Top             =   180
            Width           =   345
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Número"
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
            Left            =   13098
            TabIndex        =   210
            Top             =   180
            Width           =   555
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
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
            Height          =   195
            Left            =   4875
            TabIndex        =   209
            Top             =   180
            Width           =   405
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Cx. postal"
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
            Index           =   39
            Left            =   9435
            TabIndex        =   179
            Top             =   780
            Width           =   735
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "CEP"
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
            Index           =   38
            Left            =   10740
            TabIndex        =   178
            Top             =   780
            Width           =   285
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Fax"
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
            Index           =   9
            Left            =   13975
            TabIndex        =   177
            Top             =   780
            Width           =   270
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
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
            Index           =   37
            Left            =   7777
            TabIndex        =   176
            Top             =   1410
            Width           =   420
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Endereço"
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
            Index           =   36
            Left            =   9698
            TabIndex        =   175
            Top             =   180
            Width           =   675
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Bairro"
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
            Index           =   35
            Left            =   2985
            TabIndex        =   174
            Top             =   780
            Width           =   420
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Cidade"
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
            Index           =   34
            Left            =   7215
            TabIndex        =   173
            Top             =   780
            Width           =   495
         End
         Begin VB.Label Label3 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "UF"
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
            Height          =   255
            Index           =   33
            Left            =   5185
            TabIndex        =   172
            Top             =   780
            Width           =   210
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Tel.01"
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
            Index           =   32
            Left            =   12067
            TabIndex        =   171
            Top             =   780
            Width           =   450
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Tel.02"
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
            Index           =   23
            Left            =   855
            TabIndex        =   170
            Top             =   1410
            Width           =   450
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Tel.03"
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
            Index           =   19
            Left            =   2680
            TabIndex        =   169
            Top             =   1410
            Width           =   450
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Tel.04"
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
            Index           =   18
            Left            =   4500
            TabIndex        =   168
            Top             =   1410
            Width           =   450
         End
      End
      Begin MSComctlLib.ListView lista_familia 
         Height          =   7540
         Left            =   -74925
         TabIndex        =   100
         Top             =   2160
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   13309
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "N"
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Tag             =   "T"
            Text            =   "Família"
            Object.Width           =   25629
         EndProperty
      End
      Begin MSComctlLib.ListView listacobranca 
         Height          =   6315
         Left            =   -74925
         TabIndex        =   96
         Top             =   3390
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   11139
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "N"
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Tag             =   "T"
            Text            =   "Endereço"
            Object.Width           =   15046
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Bairro"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Tag             =   "T"
            Text            =   "Cidade"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "UF"
            Object.Width           =   882
         EndProperty
      End
      Begin VB.Frame Frame12 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   -74925
         TabIndex        =   190
         Top             =   1285
         Width           =   15195
         Begin VB.TextBox TxtResponsavel4 
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
            Left            =   1395
            Locked          =   -1  'True
            TabIndex        =   98
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pelo cadastro."
            Top             =   390
            Width           =   4305
         End
         Begin VB.TextBox txtdata4 
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
            MaxLength       =   25
            TabIndex        =   97
            TabStop         =   0   'False
            ToolTipText     =   "Data do cadastro."
            Top             =   390
            Width           =   1185
         End
         Begin VB.ComboBox cmbFamilia 
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
            Left            =   5730
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   99
            ToolTipText     =   "Família."
            Top             =   390
            Width           =   9315
         End
         Begin VB.Label Label2 
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
            Left            =   3090
            TabIndex        =   225
            Top             =   180
            Width           =   915
         End
         Begin VB.Label Label2 
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
            Left            =   600
            TabIndex        =   224
            Top             =   180
            Width           =   345
         End
         Begin VB.Label Label3 
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
            Index           =   8
            Left            =   10147
            TabIndex        =   191
            Top             =   180
            Width           =   480
         End
      End
      Begin MSComctlLib.ListView lista_banco 
         Height          =   7540
         Left            =   -74925
         TabIndex        =   106
         Top             =   2160
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   13309
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "N"
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Tag             =   "T"
            Text            =   "Banco"
            Object.Width           =   13282
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "N"
            Text            =   "Agência"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Tag             =   "N"
            Text            =   "Conta"
            Object.Width           =   6174
         EndProperty
      End
      Begin MSComctlLib.ListView Lista_Impostos 
         Height          =   7540
         Left            =   -74925
         TabIndex        =   114
         Top             =   2160
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   13309
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "N"
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Tag             =   "T"
            Text            =   "NCM"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Object.Tag             =   "N"
            Text            =   "IPI (%)"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Object.Tag             =   "N"
            Text            =   "Vlr. p/ cálc. IPI"
            Object.Width           =   2646
         EndProperty
      End
      Begin DrawSuite2022.USToolBar USToolBar3 
         Height          =   975
         Left            =   -74925
         TabIndex        =   241
         Top             =   330
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   1720
         ButtonCount     =   9
         GradientColor2  =   14737632
         GradientColorOverRight1=   16315633
         GradientColorOverRight2=   15195350
         GripperColor    =   15195350
         IsStrech        =   -1  'True
         RightColor1     =   0
         RightColor2     =   0
         ShowEndPanel    =   0   'False
         Theme           =   1
         ButtonCaption1  =   "Salvar"
         ButtonEnabled1  =   0   'False
         ButtonIconSize1 =   32
         ButtonToolTipText1=   "Salvar (F3)"
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
         ButtonWidth1    =   44
         ButtonHeight1   =   21
         ButtonUseMaskColor1=   0   'False
         ButtonCaption2  =   "Excluir"
         ButtonEnabled2  =   0   'False
         ButtonIconSize2 =   32
         ButtonToolTipText2=   "Excluir (F4)"
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
         ButtonLeft2     =   48
         ButtonTop2      =   2
         ButtonWidth2    =   45
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft3     =   95
         ButtonTop3      =   2
         ButtonWidth3    =   60
         ButtonHeight3   =   21
         ButtonUseMaskColor3=   0   'False
         ButtonCaption4  =   "Anterior"
         ButtonEnabled4  =   0   'False
         ButtonIconSize4 =   32
         ButtonToolTipText4=   "Registro anterior."
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
         ButtonLeft4     =   157
         ButtonTop4      =   2
         ButtonWidth4    =   55
         ButtonHeight4   =   21
         ButtonUseMaskColor4=   0   'False
         ButtonCaption5  =   "Próximo"
         ButtonEnabled5  =   0   'False
         ButtonIconSize5 =   32
         ButtonToolTipText5=   "Próximo registro."
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
         ButtonLeft5     =   214
         ButtonTop5      =   2
         ButtonWidth5    =   55
         ButtonHeight5   =   21
         ButtonUseMaskColor5=   0   'False
         ButtonEnabled6  =   0   'False
         ButtonIconSize6 =   32
         ButtonAlignment6=   2
         ButtonType6     =   1
         ButtonStyle6    =   -1
         BeginProperty ButtonFont6 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState6    =   -1
         ButtonLeft6     =   271
         ButtonTop6      =   4
         ButtonWidth6    =   2
         ButtonHeight6   =   54
         ButtonCaption7  =   "Ajuda"
         ButtonEnabled7  =   0   'False
         ButtonIconSize7 =   32
         ButtonToolTipText7=   "Ajuda (F1)"
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
         ButtonLeft7     =   275
         ButtonTop7      =   2
         ButtonWidth7    =   41
         ButtonHeight7   =   21
         ButtonUseMaskColor7=   0   'False
         ButtonCaption8  =   "Sair"
         ButtonEnabled8  =   0   'False
         ButtonIconSize8 =   32
         ButtonToolTipText8=   "Sair (Esc)"
         ButtonKey8      =   "8"
         ButtonAlignment8=   2
         BeginProperty ButtonFont8 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft8     =   318
         ButtonTop8      =   2
         ButtonWidth8    =   30
         ButtonHeight8   =   21
         ButtonUseMaskColor8=   0   'False
         ButtonEnabled9  =   0   'False
         ButtonIconSize9 =   32
         ButtonKey9      =   "9"
         ButtonAlignment9=   2
         BeginProperty ButtonFont9 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState9    =   5
         ButtonLeft9     =   350
         ButtonTop9      =   2
         ButtonWidth9    =   24
         ButtonHeight9   =   24
         Begin DrawSuite2022.USImageList USImageList3 
            Left            =   12120
            Top             =   210
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmVendas_cliente.frx":1F75B
            Count           =   1
         End
      End
      Begin DrawSuite2022.USToolBar USToolBar4 
         Height          =   975
         Left            =   -74925
         TabIndex        =   242
         Top             =   330
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   1720
         ButtonCount     =   8
         GradientColor2  =   14737632
         GradientColorOverRight1=   16315633
         GradientColorOverRight2=   15195350
         GripperColor    =   15195350
         IsStrech        =   -1  'True
         RightColor1     =   0
         RightColor2     =   0
         ShowEndPanel    =   0   'False
         Theme           =   1
         ButtonCaption1  =   "Salvar"
         ButtonEnabled1  =   0   'False
         ButtonIconSize1 =   32
         ButtonToolTipText1=   "Salvar (F3)"
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
         ButtonWidth1    =   44
         ButtonHeight1   =   21
         ButtonUseMaskColor1=   0   'False
         ButtonCaption2  =   "Relatório"
         ButtonEnabled2  =   0   'False
         ButtonIconSize2 =   32
         ButtonToolTipText2=   "Relatório (F5)"
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
         ButtonLeft2     =   48
         ButtonTop2      =   2
         ButtonWidth2    =   60
         ButtonHeight2   =   21
         ButtonUseMaskColor2=   0   'False
         ButtonCaption3  =   "Anterior"
         ButtonEnabled3  =   0   'False
         ButtonIconSize3 =   32
         ButtonToolTipText3=   "Registro anterior."
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
         ButtonLeft3     =   110
         ButtonTop3      =   2
         ButtonWidth3    =   55
         ButtonHeight3   =   21
         ButtonUseMaskColor3=   0   'False
         ButtonCaption4  =   "Próximo"
         ButtonEnabled4  =   0   'False
         ButtonIconSize4 =   32
         ButtonToolTipText4=   "Próximo registro."
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
         ButtonLeft4     =   167
         ButtonTop4      =   2
         ButtonWidth4    =   55
         ButtonHeight4   =   21
         ButtonUseMaskColor4=   0   'False
         ButtonEnabled5  =   0   'False
         ButtonIconSize5 =   32
         ButtonAlignment5=   2
         ButtonType5     =   1
         ButtonStyle5    =   -1
         BeginProperty ButtonFont5 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState5    =   -1
         ButtonLeft5     =   224
         ButtonTop5      =   4
         ButtonWidth5    =   2
         ButtonHeight5   =   54
         ButtonCaption6  =   "Ajuda"
         ButtonEnabled6  =   0   'False
         ButtonIconSize6 =   32
         ButtonToolTipText6=   "Ajuda (F1)"
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
         ButtonLeft6     =   228
         ButtonTop6      =   2
         ButtonWidth6    =   41
         ButtonHeight6   =   21
         ButtonUseMaskColor6=   0   'False
         ButtonCaption7  =   "Sair"
         ButtonEnabled7  =   0   'False
         ButtonIconSize7 =   32
         ButtonToolTipText7=   "Sair (Esc)"
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
         ButtonLeft7     =   271
         ButtonTop7      =   2
         ButtonWidth7    =   30
         ButtonHeight7   =   21
         ButtonUseMaskColor7=   0   'False
         ButtonEnabled8  =   0   'False
         ButtonIconSize8 =   32
         ButtonKey8      =   "8"
         ButtonAlignment8=   2
         BeginProperty ButtonFont8 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState8    =   5
         ButtonLeft8     =   303
         ButtonTop8      =   2
         ButtonWidth8    =   24
         ButtonHeight8   =   24
         Begin DrawSuite2022.USImageList USImageList4 
            Left            =   13530
            Top             =   210
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmVendas_cliente.frx":24211
            Count           =   1
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3210
         Left            =   75
         TabIndex        =   247
         Top             =   1290
         Width           =   15225
         Begin VB.ComboBox cmbVendedor 
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
            ItemData        =   "frmVendas_cliente.frx":27F95
            Left            =   11910
            List            =   "frmVendas_cliente.frx":27F97
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   286
            ToolTipText     =   "Vendedor responsável pelo cliente"
            Top             =   990
            Width           =   2475
         End
         Begin VB.TextBox txtIDVendedor 
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
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   11940
            Locked          =   -1  'True
            MaxLength       =   5
            TabIndex        =   290
            ToolTipText     =   "% Comissão"
            Top             =   990
            Width           =   390
         End
         Begin VB.TextBox txtComissao 
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
            Height          =   330
            Left            =   14370
            MaxLength       =   5
            TabIndex        =   288
            ToolTipText     =   "% Comissão"
            Top             =   990
            Width           =   390
         End
         Begin VB.ComboBox cmbTipo_bairro 
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
            ItemData        =   "frmVendas_cliente.frx":27F99
            Left            =   3870
            List            =   "frmVendas_cliente.frx":27FD6
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   15
            ToolTipText     =   "Tipo do bairro."
            Top             =   2190
            Width           =   1125
         End
         Begin VB.TextBox txtBairro 
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
            Height          =   325
            Left            =   5010
            MaxLength       =   60
            TabIndex        =   16
            ToolTipText     =   "Bairro."
            Top             =   2190
            Width           =   3180
         End
         Begin VB.TextBox txtemail 
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
            Height          =   330
            Left            =   3480
            MaxLength       =   60
            TabIndex        =   23
            ToolTipText     =   "E-mail."
            Top             =   2775
            Width           =   2745
         End
         Begin VB.TextBox txtSite 
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
            Height          =   330
            Left            =   6240
            MaxLength       =   100
            TabIndex        =   24
            ToolTipText     =   "Site."
            Top             =   2775
            Width           =   2775
         End
         Begin VB.ComboBox TXTcategoria 
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
            ItemData        =   "frmVendas_cliente.frx":28084
            Left            =   11340
            List            =   "frmVendas_cliente.frx":28086
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   9
            ToolTipText     =   "Categoria."
            Top             =   990
            Width           =   555
         End
         Begin VB.ComboBox cmbUF 
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
            ItemData        =   "frmVendas_cliente.frx":28088
            Left            =   4035
            List            =   "frmVendas_cliente.frx":2808A
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   3
            ToolTipText     =   "UF."
            Top             =   990
            Width           =   615
         End
         Begin VB.TextBox txtcep 
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
            Left            =   120
            MaxLength       =   12
            TabIndex        =   12
            Text            =   "13.339-575"
            ToolTipText     =   "CEP."
            Top             =   2190
            Width           =   915
         End
         Begin VB.TextBox txtcaixapostal 
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
            Height          =   330
            Left            =   120
            MaxLength       =   50
            TabIndex        =   20
            ToolTipText     =   "Caixa postal."
            Top             =   2775
            Width           =   795
         End
         Begin VB.TextBox txtIM_IE 
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
            Left            =   10020
            MaxLength       =   15
            TabIndex        =   8
            ToolTipText     =   "Inscrição municipal."
            Top             =   990
            Width           =   1305
         End
         Begin VB.TextBox txtStatus 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            ForeColor       =   &H00000080&
            Height          =   315
            Left            =   11550
            Locked          =   -1  'True
            TabIndex        =   30
            TabStop         =   0   'False
            ToolTipText     =   "Status."
            Top             =   390
            Width           =   3525
         End
         Begin VB.TextBox txtnomefantasia 
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
            Height          =   330
            Left            =   6690
            MaxLength       =   60
            TabIndex        =   11
            ToolTipText     =   "Nome fantasia."
            Top             =   1590
            Width           =   8385
         End
         Begin VB.TextBox txtnomerazao 
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
            Height          =   330
            Left            =   840
            MaxLength       =   60
            TabIndex        =   10
            ToolTipText     =   "Razão social."
            Top             =   1590
            Width           =   5835
         End
         Begin VB.ComboBox cmbPessoa 
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
            ItemData        =   "frmVendas_cliente.frx":2808C
            Left            =   165
            List            =   "frmVendas_cliente.frx":28096
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   0
            ToolTipText     =   "Tipo."
            Top             =   990
            Width           =   1020
         End
         Begin VB.TextBox txtRG_IE 
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
            Left            =   8730
            MaxLength       =   15
            TabIndex        =   7
            ToolTipText     =   "Inscrição estadual."
            Top             =   990
            Width           =   1275
         End
         Begin VB.TextBox txtEndereco 
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
            Height          =   325
            Left            =   9255
            MaxLength       =   60
            TabIndex        =   18
            ToolTipText     =   "Endereço."
            Top             =   2190
            Width           =   4185
         End
         Begin VB.ComboBox cmbtransportadora 
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
            Left            =   10260
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   25
            ToolTipText     =   "Nome da transportadora."
            Top             =   2775
            Width           =   4455
         End
         Begin VB.TextBox txtNumero 
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
            Height          =   325
            Left            =   13455
            MaxLength       =   5
            TabIndex        =   35
            ToolTipText     =   "Número."
            Top             =   2190
            Width           =   570
         End
         Begin VB.TextBox TxtResponsavel 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   1170
            Locked          =   -1  'True
            TabIndex        =   29
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pelo cadastro."
            Top             =   390
            Width           =   4575
         End
         Begin VB.TextBox txtdata 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   180
            Locked          =   -1  'True
            MaxLength       =   25
            TabIndex        =   28
            TabStop         =   0   'False
            ToolTipText     =   "Data do cadastro."
            Top             =   390
            Width           =   975
         End
         Begin VB.ComboBox Txt_pais 
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
            ItemData        =   "frmVendas_cliente.frx":280AC
            Left            =   2400
            List            =   "frmVendas_cliente.frx":280AE
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   2
            ToolTipText     =   "País."
            Top             =   990
            Width           =   1635
         End
         Begin VB.ComboBox cmbTipo_endereco 
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
            ItemData        =   "frmVendas_cliente.frx":280B0
            Left            =   8190
            List            =   "frmVendas_cliente.frx":280F0
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   17
            ToolTipText     =   "Tipo do endereço."
            Top             =   2190
            Width           =   1050
         End
         Begin VB.TextBox txtComplemento 
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
            Height          =   325
            Left            =   14040
            MaxLength       =   30
            TabIndex        =   19
            ToolTipText     =   "Complemento."
            Top             =   2190
            Width           =   1035
         End
         Begin VB.ComboBox Cmb_tipo_transp 
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
            ItemData        =   "frmVendas_cliente.frx":28187
            Left            =   9030
            List            =   "frmVendas_cliente.frx":28197
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   26
            ToolTipText     =   "Tipo da transportadora."
            Top             =   2775
            Width           =   1215
         End
         Begin VB.TextBox txtRespValidacao 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   7800
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   32
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pela validação."
            Top             =   390
            Width           =   3735
         End
         Begin VB.TextBox txtDtValidacao 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   5760
            Locked          =   -1  'True
            TabIndex        =   31
            TabStop         =   0   'False
            ToolTipText     =   "Data e hora da validação."
            Top             =   390
            Width           =   2025
         End
         Begin VB.ComboBox cmbRegimeTributario 
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
            ItemData        =   "frmVendas_cliente.frx":281BB
            Left            =   7230
            List            =   "frmVendas_cliente.frx":281BD
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   5
            ToolTipText     =   "Regime tributário."
            Top             =   990
            Width           =   1485
         End
         Begin MSMask.MaskEdBox txttel01 
            Height          =   330
            Left            =   930
            TabIndex        =   21
            ToolTipText     =   "Número do telefone."
            Top             =   2775
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   582
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   0
            AllowPrompt     =   -1  'True
            AutoTab         =   -1  'True
            MaxLength       =   14
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtfax 
            Height          =   330
            Left            =   2370
            TabIndex        =   22
            ToolTipText     =   "Número do telefone(2)"
            Top             =   2775
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   582
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   0
            AllowPrompt     =   -1  'True
            AutoTab         =   -1  'True
            MaxLength       =   14
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtcnpj 
            Height          =   315
            Left            =   4650
            TabIndex        =   4
            ToolTipText     =   "Número do CNPJ."
            Top             =   990
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   0
            MaxLength       =   18
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##.###.###/####-##"
            PromptChar      =   "_"
         End
         Begin VB.ComboBox cmbCidade 
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
            ItemData        =   "frmVendas_cliente.frx":281BF
            Left            =   1395
            List            =   "frmVendas_cliente.frx":281C1
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   14
            ToolTipText     =   "Cidade."
            Top             =   2190
            Width           =   2490
         End
         Begin VB.TextBox txtCidade 
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
            Left            =   2295
            MaxLength       =   60
            TabIndex        =   36
            ToolTipText     =   "Cidade."
            Top             =   2190
            Visible         =   0   'False
            Width           =   2580
         End
         Begin VB.ComboBox cmbOrigem 
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
            ItemData        =   "frmVendas_cliente.frx":281C3
            Left            =   1200
            List            =   "frmVendas_cliente.frx":281D0
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   1
            ToolTipText     =   "Origem."
            Top             =   990
            Width           =   1185
         End
         Begin MSMask.MaskEdBox txtCpf 
            Height          =   315
            Left            =   4650
            TabIndex        =   34
            ToolTipText     =   "Número do CPF."
            Top             =   990
            Visible         =   0   'False
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   0
            MaxLength       =   14
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "###.###.###-##"
            PromptChar      =   "_"
         End
         Begin DrawSuite2022.USButton cmdConsultar 
            Height          =   315
            Left            =   6240
            TabIndex        =   6
            ToolTipText     =   "Consultar cadastro no SEFAZ"
            Top             =   990
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   556
            DibPicture      =   "frmVendas_cliente.frx":281ED
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
         Begin DrawSuite2022.USButton Cmd_buscarCEP 
            Height          =   315
            Left            =   1050
            TabIndex        =   13
            ToolTipText     =   "Consultar endereço por CEP nos correios"
            Top             =   2190
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   556
            DibPicture      =   "frmVendas_cliente.frx":2B83D
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
            Theme           =   1
            ToolTipTitle    =   "CAPRIND v5.0"
         End
         Begin DrawSuite2022.USButton cmdLocTransp 
            Height          =   315
            Left            =   14730
            TabIndex        =   27
            ToolTipText     =   "Consultar cadastro transportadora"
            Top             =   2775
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   556
            DibPicture      =   "frmVendas_cliente.frx":329D0
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
            BorderColor     =   1154291
            BorderColorDisabled=   13160660
            BorderColorDown =   16576
            BorderColorOver =   8438015
            GradientColor1  =   1154291
            GradientColor2  =   1154291
            GradientColor3  =   1154291
            GradientColor4  =   1154291
            GradientColorDisabled1=   14215660
            GradientColorDisabled2=   14215660
            GradientColorDisabled3=   14215660
            GradientColorDisabled4=   14215660
            GradientColorOver1=   8438015
            GradientColorOver2=   8438015
            GradientColorOver3=   8438015
            GradientColorOver4=   8438015
            GradientColorDown1=   16576
            GradientColorDown2=   16576
            GradientColorDown3=   16576
            GradientColorDown4=   16576
            PicAlign        =   8
            Theme           =   5
            ToolTipTitle    =   "CAPRIND v5.0"
         End
         Begin DrawSuite2022.USButton btnSintegra 
            Height          =   315
            Left            =   6570
            TabIndex        =   294
            ToolTipText     =   "Consultar cadastro no Sintegra."
            Top             =   990
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   556
            DibPicture      =   "frmVendas_cliente.frx":38435
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
         Begin DrawSuite2022.USButton btnRF 
            Height          =   315
            Left            =   6900
            TabIndex        =   295
            ToolTipText     =   "Consultar cadastro na receita federal."
            Top             =   990
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   556
            DibPicture      =   "frmVendas_cliente.frx":40C65
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
         Begin DrawSuite2022.USButton btnWeb 
            Height          =   315
            Left            =   14760
            TabIndex        =   296
            ToolTipText     =   "Atualizar cliente na WEB"
            Top             =   990
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   556
            DibPicture      =   "frmVendas_cliente.frx":46ACF
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
            BorderColor     =   1154291
            BorderColorDisabled=   13160660
            BorderColorDown =   16576
            BorderColorOver =   8438015
            GradientColor1  =   1154291
            GradientColor2  =   1154291
            GradientColor3  =   1154291
            GradientColor4  =   1154291
            GradientColorDisabled1=   14215660
            GradientColorDisabled2=   14215660
            GradientColorDisabled3=   14215660
            GradientColorDisabled4=   14215660
            GradientColorOver1=   8438015
            GradientColorOver2=   8438015
            GradientColorOver3=   8438015
            GradientColorOver4=   8438015
            GradientColorDown1=   16576
            GradientColorDown2=   16576
            GradientColorDown3=   16576
            GradientColorDown4=   16576
            PicAlign        =   8
            Theme           =   5
            ToolTipTitle    =   "CAPRIND v5.0"
         End
         Begin VB.TextBox txtVendedor 
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
            Height          =   330
            Left            =   11940
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   291
            ToolTipText     =   "Vendedor Interno"
            Top             =   990
            Visible         =   0   'False
            Width           =   2340
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
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
            Left            =   14475
            TabIndex        =   289
            Top             =   780
            Width           =   165
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vendedor interno"
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
            Index           =   20
            Left            =   12495
            TabIndex        =   287
            Top             =   780
            Width           =   1245
         End
         Begin VB.Label Label3 
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
            Index           =   40
            Left            =   4215
            TabIndex        =   285
            Top             =   1980
            Width           =   300
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Bairro"
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
            Left            =   6390
            TabIndex        =   284
            Top             =   1980
            Width           =   420
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CNPJ*"
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
            Height          =   195
            Left            =   5160
            TabIndex        =   277
            Top             =   780
            Width           =   510
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CPF"
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
            Left            =   5295
            TabIndex        =   276
            Top             =   780
            Visible         =   0   'False
            Width           =   300
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Nome fantasia"
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
            Left            =   10335
            TabIndex        =   275
            Top             =   1380
            Width           =   1095
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Telefone"
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
            Index           =   6
            Left            =   1327
            TabIndex        =   274
            Top             =   2580
            Width           =   630
         End
         Begin VB.Label Label3 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "UF*"
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
            Height          =   255
            Index           =   5
            Left            =   4245
            TabIndex        =   273
            Top             =   780
            Width           =   315
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Cidade"
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
            Left            =   2393
            TabIndex        =   272
            Top             =   1980
            Width           =   495
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Endereço"
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
            Left            =   11160
            TabIndex        =   271
            Top             =   1980
            Width           =   675
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Razão social (60 caracteres máximo NFe)*"
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
            Height          =   195
            Index           =   0
            Left            =   1957
            TabIndex        =   270
            Top             =   1380
            Width           =   3600
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Insc. estadual*"
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
            Index           =   2
            Left            =   8820
            TabIndex        =   269
            Top             =   780
            Width           =   1110
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
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
            Index           =   7
            Left            =   4635
            TabIndex        =   268
            Top             =   2580
            Width           =   420
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Telefone (2)"
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
            Left            =   2505
            TabIndex        =   267
            Top             =   2580
            Width           =   885
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "CEP"
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
            Left            =   495
            TabIndex        =   266
            Top             =   1980
            Width           =   285
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Cx. postal"
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
            Index           =   9
            Left            =   150
            TabIndex        =   265
            Top             =   2580
            Width           =   735
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
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
            Index           =   17
            Left            =   13080
            TabIndex        =   264
            Top             =   180
            Width           =   465
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo*"
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
            TabIndex        =   263
            Top             =   780
            Width           =   390
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Insc. municipal"
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
            Left            =   10155
            TabIndex        =   262
            Top             =   780
            Width           =   1050
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Transportadora*"
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
            Index           =   4
            Left            =   11925
            TabIndex        =   261
            Top             =   2580
            Width           =   1215
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "País*"
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
            Index           =   42
            Left            =   3030
            TabIndex        =   260
            Top             =   780
            Width           =   375
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "N°"
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
            Left            =   13680
            TabIndex        =   259
            Top             =   1980
            Width           =   180
         End
         Begin VB.Label Label2 
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
            Index           =   9
            Left            =   3015
            TabIndex        =   258
            Top             =   180
            Width           =   915
         End
         Begin VB.Label Label2 
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
            Index           =   5
            Left            =   495
            TabIndex        =   257
            Top             =   180
            Width           =   345
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Complemento"
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
            Index           =   44
            Left            =   14070
            TabIndex        =   256
            Top             =   1980
            Width           =   975
         End
         Begin VB.Label Label3 
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
            Index           =   46
            Left            =   8535
            TabIndex        =   255
            Top             =   1980
            Width           =   300
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo*"
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
            Index           =   3
            Left            =   9420
            TabIndex        =   254
            Top             =   2580
            Width           =   390
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Responsável pela validação"
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
            Left            =   8670
            TabIndex        =   253
            Top             =   180
            Width           =   1980
         End
         Begin VB.Label Label49 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
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
            Left            =   5932
            TabIndex        =   252
            Top             =   180
            Width           =   1680
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Site"
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
            Index           =   53
            Left            =   7492
            TabIndex        =   251
            Top             =   2580
            Width           =   270
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Regime tributário*"
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
            Index           =   54
            Left            =   7320
            TabIndex        =   250
            Top             =   780
            Width           =   1320
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Origem*"
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
            Index           =   55
            Left            =   1530
            TabIndex        =   249
            Top             =   780
            Width           =   600
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cat*"
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
            Height          =   195
            Index           =   21
            Left            =   11430
            TabIndex        =   248
            Top             =   780
            Width           =   390
         End
      End
   End
End
Attribute VB_Name = "frmVendas_cliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Novo_Cliente     As Boolean 'OK
Public Novo_Cliente1    As Boolean 'OK
Public Novo_Cliente2    As Boolean 'OK
Public Novo_Cliente3    As Boolean 'OK
Public Novo_Cliente4    As Boolean 'OK
Public Novo_Cliente5    As Boolean 'OK
Public Novo_Cliente6    As Boolean 'OK
Public StrSql_Cliente   As String 'OK
Public FormulaRel_Cliente   As String 'OK
Dim TBLISTA_Cliente     As ADODB.Recordset 'OK

Sub ProcAjuda()
On Error GoTo tratar_erro

FunAbrirVideoWeb ("http://www.youtube.com/watch?v=AJYBrtXnRCw&list=UUODGiDjQ-BCrxh0YXfJvoqg&index=54&feature=plcp")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaComboVendedor()
On Error GoTo tratar_erro

cmbVendedor.Clear

Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Vendas_Vendedores order by Vendedor", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    cmbVendedor.AddItem ""
    Do While TBLISTA.EOF = False
        cmbVendedor.AddItem TBLISTA!vendedor
        TBLISTA.MoveNext
    Loop
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaListaContatos()
On Error GoTo tratar_erro

Lista_contato.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Clientes_contatos where Idcliente = " & IIf(txtIDcliente = "", 0, txtIDcliente) & " order by nomecontato", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    TBLISTA.MoveLast
    Contador = 0
    TBLISTA.MoveFirst
    Do While TBLISTA.EOF = False
        With Lista_contato.ListItems
            .Add = TBLISTA!idcontato
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!NomeContato), "", TBLISTA!NomeContato)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Departamento), "", TBLISTA!Departamento)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!telefone), "", TBLISTA!telefone)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!Email), "", TBLISTA!Email)
        End With
        TBLISTA.MoveNext
        Contador = Contador + 1
        'PBLista.Value = Contador
    Loop
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub procLimpacamposContatos()
On Error GoTo tratar_erro

txtIDContato.Text = 0
txtData1 = Format(Date, "dd/mm/yy")
txtResponsavel1 = pubUsuario
txtNomeContato.Text = ""
txtdepartamento.Text = ""
txttelcontato.Text = ""
TxtEmail_Contato.Text = ""
Chk_enviar_NFe.Value = 0
Chk_enviar_boleto.Value = 0
CodigoLista1 = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaFamilia()
On Error GoTo tratar_erro

ProcCarregaComboFamilia cmbfamilia, "familia <> 'Null' and vendas = 'True'", False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ActiveResize1_ResizeComplete()
On Error GoTo tratar_erro

If SSTab1.Tab = 0 Or SSTab1.Tab = 8 Then
    With Lista
        .Visible = True
        If SSTab1.Tab = 0 Then
            .Top = Frame1.Top + Frame1.Height
            .Height = Frame15.Top - .Top
        Else
            .Top = Frame19.Top + Frame19.Height
            .Height = Frame15.Top - .Top
        End If
    End With
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub btnRF_Click()
On Error GoTo tratar_erro

Dim resposta As String
Dim obj As MSXML2.ServerXMLHTTP50
Set obj = New MSXML2.ServerXMLHTTP50
Dim Plugin As String

If cmbPessoa.Text = "" Then
      USMsgBox "Não será possivel buscar o cadastro desse cnpj pois não foi informado o tipo", vbOKOnly, "CAPRIND v5.0"
      cmbPessoa.SetFocus
      Exit Sub
End If

If cmbOrigem.Text = "" Then
      USMsgBox "Não será possivel buscar o cadastro desse cnpj pois não foi informado a Origem", vbOKOnly, "CAPRIND v5.0"
      cmbOrigem.SetFocus
      Exit Sub
End If

If Txt_pais.Text = "" Then
      USMsgBox "Não será possivel buscar o cadastro desse cnpj pois não foi informado o país", vbOKOnly, "CAPRIND v5.0"
      Txt_pais.SetFocus
      Exit Sub
End If


If cmbuf.Text = "" Then
      USMsgBox "Não será possivel buscar o cadastro desse cnpj pois não foi informado a UF", vbOKOnly, "CAPRIND v5.0"
      cmbuf.SetFocus
      Exit Sub
End If

If txtcnpj.Text = "__.___.___/____-__" Then
      USMsgBox "Não será possivel buscar o cadastro desse cnpj pois não foi informado O CNPJ", vbOKOnly, "CAPRIND v5.0"
      txtcnpj.SetFocus
      Exit Sub
End If

Plugin = "RF"

CnpjDados = ReturnNumbersOnly(txtcnpj.Text)

obj.Open "GET", "https://www.sintegraws.com.br/api/v1/execute-api.php?token=1F718E4E-3222-42F1-95D6-995FC9E69C9C&cnpj=" & CnpjDados & "&plugin=" & Plugin & ""

conteudo = CnpjDados
obj.send conteudo

resposta = obj.responseText
'Debug.print resposta

If LerDadosJSON(resposta, "status", "", "") = "OK" And LerDadosJSON(resposta, "code", "", "") = "0" Then

USMsgBox LerDadosJSON(resposta, "message", "", ""), vbInformation, "CAPRIND v5.0"
txtnomerazao.Text = UCase(LerDadosJSON(resposta, "nome", "", ""))
cmbuf.Text = UCase(LerDadosJSON(resposta, "uf", "", ""))
txttel01 = LerDadosJSON(resposta, "telefone", "", "")
txtBairro = UCase(LerDadosJSON(resposta, "bairro", "", ""))
txtendereco = UCase(LerDadosJSON(resposta, "logradouro", "", ""))
txtNumero = LerDadosJSON(resposta, "numero", "", "")
txtCEP = LerDadosJSON(resposta, "cep", "", "")
txtEmail = LerDadosJSON(resposta, "email", "", "")
Cidade = UCase(LerDadosJSON(resposta, "municipio", "", ""))
If Cidade = "SANTA BARBARA D'OESTE" Then
cmbCidade.Text = "SANTA BARBARA DO OESTE"
Else
cmbCidade.Text = UCase(LerDadosJSON(resposta, "municipio", "", ""))
End If


txtnomefantasia = UCase(LerDadosJSON(resposta, "fantasia", "", ""))
'cmbRegimeTributario.Text = IIf(LerDadosJSON(resposta, "regime_tributacao", "", "") = "Normal - regime periódico de apuração", "Lucro presumido", "Simples Nacional")
'txtRG_IE = Trim(LerDadosJSON(resposta, "inscricao_estadual", "", ""))

txtCategoria.Text = "A"

Cmd_buscarCEP_Click
Else
USMsgBox LerDadosJSON(resposta, "message", "", ""), vbInformation, "CAPRIND v5.0"
txtnomerazao.Text = ""
cmbuf.ListIndex = -1
txtBairro = ""
txtendereco = ""
txtNumero = ""
txtCEP = ""
txtnomefantasia = ""
cmbRegimeTributario.ListIndex = -1
txtRG_IE = ""
txtCategoria.ListIndex = -1

End If


Exit Sub
tratar_erro:
    MousePointer = 0
    If Err.Number = 91 Then
        USMsgBox ("Não foi possível carregar todos os dados referentes a este CEP."), vbInformation, "CAPRIND v5.0"
        Exit Sub
    End If
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub btnWeb_Click()
On Error GoTo tratar_erro

If txtIDcliente.Text <> "" Then
    procAtualizaClienteNuvem
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_nao_contribuinte_ICMS_Click()
On Error GoTo tratar_erro

If Chk_nao_contribuinte_ICMS.Value = 1 Then


End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkSuframa_Click()
On Error GoTo tratar_erro

With txtSuframa
    If chkSuframa.Value = 1 Then
        .Locked = False
        .TabStop = True
    Else
        .Locked = True
        .TabStop = False
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_empresa_Click()
On Error GoTo tratar_erro

If Cmb_empresa <> "" Then
    ProcLimpaCamposComerciais
    procPuxadados_Comerciais
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
    If Cmb_opcao_lista = "Validação" Then
        .ButtonState(4) = 5
        .ButtonState(9) = 5
        .ButtonState(10) = 0
    ElseIf Cmb_opcao_lista = "Excluir" Then
            .ButtonState(4) = 0
            .ButtonState(9) = 5
            .ButtonState(10) = 5
        Else
            .ButtonState(4) = 5
            .ButtonState(9) = 0
            .ButtonState(10) = 5
    End If
    .Refresh
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_tipo_transp_Click()
On Error GoTo tratar_erro

ProcCarregaComboTransp

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaComboTransp()
On Error GoTo tratar_erro

With cmbtransportadora
    .Clear
    If Cmb_tipo_transp <> "" Then
        If Cmb_tipo_transp = "Empresa" Then
            NomeTabela = "Empresa"
            NomeCampo = "Empresa"
            NomeCampo1 = "Codigo"
        Else
            NomeCampo1 = "IDCliente"
            If Cmb_tipo_transp = "Cliente" Then
                NomeTabela = "Clientes"
                NomeCampo = "NomeRazao"
            Else
                NomeTabela = "Compras_fornecedores"
                NomeCampo = "Nome_Razao"
            End If
        End If
        
        Set TBLISTA = CreateObject("adodb.recordset")
        TBLISTA.Open "Select " & NomeCampo & ", " & NomeCampo1 & " FROM " & NomeTabela & " where " & NomeCampo & " is not null group by " & NomeCampo & ", " & NomeCampo1, Conexao, adOpenKeyset, adLockOptimistic
        If TBLISTA.EOF = False Then
            .AddItem ""
            Do While TBLISTA.EOF = False
                Select Case Cmb_tipo_transp
                    Case "Cliente":
                        .AddItem TBLISTA!NomeRazao
                        .ItemData(.NewIndex) = TBLISTA!IDCliente
                    Case "Fornecedor":
                        .AddItem TBLISTA!Nome_Razao
                        .ItemData(.NewIndex) = TBLISTA!IDCliente
                    Case "Empresa":
                        .AddItem TBLISTA!Empresa
                        .ItemData(.NewIndex) = TBLISTA!CODIGO
                End Select
                TBLISTA.MoveNext
            Loop
        End If
        TBLISTA.Close
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub cmbVendedor_Change()
On Error GoTo tratar_erro

If cmbVendedor.Text <> "" Then
   Set TBUsuarios = CreateObject("adodb.recordset")
   TBUsuarios.Open "Select * from Vendas_Vendedores where Vendedor = '" & cmbVendedor.Text & "'", Conexao, adOpenKeyset, adLockOptimistic
   If TBUsuarios.EOF = False Then
   txtIDVendedor.Text = TBUsuarios!ID
   End If
   TBUsuarios.Close
Else
txtIDVendedor.Text = ""
txtComissao.Text = ""
End If
   
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbVendedor_Click()
On Error GoTo tratar_erro

If cmbVendedor.Text <> "" Then
   Set TBUsuarios = CreateObject("adodb.recordset")
   TBUsuarios.Open "Select * from Vendas_Vendedores where Vendedor = '" & cmbVendedor.Text & "'", Conexao, adOpenKeyset, adLockOptimistic
   If TBUsuarios.EOF = False Then
   txtIDVendedor.Text = TBUsuarios!ID
   End If
   TBUsuarios.Close
Else
txtIDVendedor.Text = ""
txtComissao.Text = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub

End Sub

Private Sub Cmd_buscarCEP_Click()
On Error GoTo tratar_erro

If txtCEP = "" Then Exit Sub

If Len(ReturnNumbersOnly(txtCEP)) < 8 Then
    USMsgBox "CEP informado com menos de oito dígitos é inválido, favor verificar!", vbCritical, "CAPRIND V5.0"
    Exit Sub
End If

If cmbOrigem = "" Or cmbOrigem = "Estrangeiro" Then
    USMsgBox ("Só é permitido carregar os dados pelo CEP se a origem for Nacional."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If FunBuscaEndereco(txtCEP) = False Then
'    usMsgbox ("Não foi encontrado nenhuma informação pelo CEP informado."), vbExclamation, "CAPRIND v5.0"
'    cmbTipo_endereco.ListIndex = -1
'    txtEndereco = ""
'    cmbTipo_bairro.ListIndex = -1
'    txtBairro = ""
'    cmbUF.ListIndex = -1
'    cmbCidade.ListIndex = -1
    Exit Sub
Else
    Permitido = True
'    cmbTipo_endereco = Tipo_endereco
    txtendereco = Trim(IIf(Permitido = True, UCase(Endereco), Endereco))
    txtBairro = Trim(IIf(Permitido = True, UCase(Bairro), Bairro))
    cmbuf = UF
    If Cidade = "Santa Bárbara D'Oeste" Then
    cmbCidade = "SANTA BARBARA DO OESTE"
    Else
   cmbCidade = Trim(FunTiraAcentosTexto(Cidade))
   End If
   
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_limpar_CFOP_Click()
On Error GoTo tratar_erro

txtID_cfop = ""
txtCFOP = ""
txtOperacao = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_limpar_grupo_Click()
On Error GoTo tratar_erro

txtIDGrupo = 0
txtGrupo = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_localizar_tipo_dcto_Click()
On Error GoTo tratar_erro

Financeiro_Contas_Pagar = False
Financeiro_Contas_Receber = False
Clientes = True
Compras_Fornecedores = False
frmContas_Tipo_Dcto.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdConsultar_Click()
On Error GoTo tratar_erro
Dim resposta As String

If cmbPessoa.Text = "" Then
      USMsgBox "Não será possivel buscar o cadastro desse cnpj pois não foi informado o tipo", vbOKOnly, "CAPRIND v5.0"
      cmbPessoa.SetFocus
      Exit Sub
End If

If cmbOrigem.Text = "" Then
      USMsgBox "Não será possivel buscar o cadastro desse cnpj pois não foi informado a Origem", vbOKOnly, "CAPRIND v5.0"
      cmbOrigem.SetFocus
      Exit Sub
End If

If Txt_pais.Text = "" Then
      USMsgBox "Não será possivel buscar o cadastro desse cnpj pois não foi informado o país", vbOKOnly, "CAPRIND v5.0"
      Txt_pais.SetFocus
      Exit Sub
End If


If cmbuf.Text = "" Then
      USMsgBox "Não será possivel buscar o cadastro desse cnpj pois não foi informado a UF", vbOKOnly, "CAPRIND v5.0"
      cmbuf.SetFocus
      Exit Sub
End If

If txtcnpj.Text = "__.___.___/____-__" Then
      USMsgBox "Não será possivel buscar o cadastro desse cnpj pois não foi informado O CNPJ", vbOKOnly, "CAPRIND v5.0"
      txtcnpj.SetFocus
      Exit Sub
End If

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from empresa where Empresa = '" & Cmb_empresa.Text & "'", Conexao, adOpenKeyset, adLockReadOnly
If TBAbrir.EOF = False Then
CnpjNF = ReturnNumbersOnly(TBAbrir!CNPJ)
End If
TBAbrir.Close

If CnpjNF = "34270461000104" Then
CnpjNF = ReturnNumbersOnly("17.966.412/0001-83")
End If

If cmbPessoa.Text = "Jurídica" Then
  resposta = consultarCadastroContribuinte(CnpjNF, cmbuf.Text, ReturnNumbersOnly(txtcnpj.Text), "CNPJ")
Else
  resposta = consultarCadastroContribuinte(CnpjNF, cmbuf.Text, ReturnNumbersOnly(txtCpf.Text), "CPF")
End If

'Debug.print resposta
status = LerDadosJSON(resposta, "status", "", "")
   If status = "200" Then
      Set p = JSON.parse(resposta)
      
      If p.Item("retConsCad").Item("infCons").Item("xMotivo") = "CNPJ da consulta nao cadastrado como contribuinte na UF. CNPJ: 16740838000151" Then
      USMsgBox "Não será possivel buscar o cadastro desse cnpj pois não é contribuinte do ICMS"
      Exit Sub
      End If
      
      If p.Item("retConsCad").Item("infCons").Item("xMotivo") = "Rejeição: CNPJ da consulta não cadastrado como contribuinte na UF" Then
      'FunBuscaDadosCNPJ (ReturnNumbersOnly(txtcnpj.Text))
      USMsgBox "Não será possivel buscar o cadastro desse cnpj pois não é contribuinte do ICMS"
      Exit Sub
      End If
      
      If p.Item("retConsCad").Item("infCons").Item("xMotivo") = "Rejeicao: CNPJ da consulta nao cadastrado como contribuinte na UF" Then
      USMsgBox "Rejeicao: CNPJ da consulta nao cadastrado como contribuinte na UF"
      Exit Sub
      End If
      
      If p.Item("retConsCad").Item("infCons").Item("xMotivo") = "Rejeicao: Sigla da UF da consulta difere da UF do Web Service" Then
      USMsgBox "Rejeicao: Sigla da UF da consulta difere da UF do Web Service"
      Exit Sub
      End If
      
      If p.Item("retConsCad").Item("infCons").Item("xMotivo") = "Rejeicao: CPF da consulta nao cadastrado como contribuinte na UF" Then
      USMsgBox "Rejeicao: CPF da consulta nao cadastrado como contribuinte na UF"
      Exit Sub
      End If
      
      If p.Item("retConsCad").Item("infCons").Item("xMotivo") = "Rejeição: UF não fornece consulta por CPF" Then
      USMsgBox "Rejeição: UF não fornece consulta por CPF"
      Exit Sub
      End If

      
      '"Rejeicao: Sigla da UF da consulta difere da UF do Web Service"

      cmbuf.Text = Trim(p.Item("retConsCad").Item("infCons").Item("UF"))
      
      If cmbPessoa.Text = "Jurídica" Then
      txtRG_IE = Trim(p.Item("retConsCad").Item("infCons").Item("infCad").Item(1).Item("IE"))
      End If
      
      txtnomerazao.Text = Trim(p.Item("retConsCad").Item("infCons").Item("infCad").Item(1).Item("xNome"))
      'txtnomefantasia = Trim(p.Item("retConsCad").Item("infCons").Item("infCad").Item(1).Item("xNome"))
      txtendereco = Trim(p.Item("retConsCad").Item("infCons").Item("infCad").Item(1).Item("ender").Item("xLgr"))
      txtNumero = Trim(p.Item("retConsCad").Item("infCons").Item("infCad").Item(1).Item("ender").Item("nro"))
      txtBairro = Trim(p.Item("retConsCad").Item("infCons").Item("infCad").Item(1).Item("ender").Item("xBairro"))
      If Trim(p.Item("retConsCad").Item("infCons").Item("infCad").Item(1).Item("ender").Item("xMun")) <> "" Then
      cmbCidade.Text = Trim(p.Item("retConsCad").Item("infCons").Item("infCad").Item(1).Item("ender").Item("xMun"))
      End If
      txtCEP = Trim(p.Item("retConsCad").Item("infCons").Item("infCad").Item(1).Item("ender").Item("CEP"))
      cmbRegimeTributario.Text = IIf(p.Item("retConsCad").Item("infCons").Item("infCad").Item(1).Item("xRegApur") = "NORMAL - REGIME PERIÓDICO DE APURAÇÃO", "Lucro presumido", "Simples Nacional")
      'txtnomefantasia = p.Item("retConsCad").Item("infCons").Item("infCad").Item(1).Item("xNome")
      If txtCEP <> "" Then
      Cmd_buscarCEP_Click
      End If
      txtCategoria.Text = "A"
      USMsgBox "Consulta relizada com sucesso, dados carregados", vbInformation, "CAPRIND v5.0"
      
   Else
      USMsgBox resposta, vbCritical, "CAPRIND v5.0"
   End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Command1_Click()
    On Error GoTo SAI
    Dim retorno As String
     
    Dim status As String
    Dim restContCad As String
    Dim infCons As String
    Dim auxInfCad As Variant
    Dim respInfCad As String
    Dim infCad As String
    
    Dim IE As String
    Dim CNPJ As String
    Dim xNome As String
    Dim xLgr As String
    Dim nro As String
    Dim xCpl As String
    Dim xBairro As String
    Dim cMun As String
    Dim CEP As String
    
Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from empresa where Empresa = '" & Cmb_empresa.Text & "'", Conexao, adOpenKeyset, adLockReadOnly
    If TBAbrir.EOF = False Then
    CnpjNF = ReturnNumbersOnly(TBAbrir!CNPJ)
    End If
TBAbrir.Close

If CnpjNF = "34270461000104" Then
    CnpjNF = ReturnNumbersOnly("07.758.985/0001-68")
End If

    retorno = consultarCadastroContribuinte(CnpjNF, cmbuf.Text, ReturnNumbersOnly(txtcnpj.Text), "CNPJ")
   ' retorno = consultarCadastroContribuinte(txtCNPJCont.Text, txtuf.Text, txtCNPJ_CPF.Text, cbTtipoContrib.Text)
    
    status = LerDadosJSON(retorno, "status", "", "")
    
    If (status = 200) Then
        cStat = LerDadosJSON(retorno, "retConsCad", "infCons", "cStat")

        If (cStat = "111") Or (cStat = "112") Then
            motivo = LerDadosJSON(retorno, "motivo", "", "")
            MsgBox (motivo)
            auxInfCad = Split(retorno, """infCad"":[")
            auxInfCad = Split(auxInfCad(1), "]")
            auxInfCad = Split(auxInfCad(0), "},")

            If (UBound(auxInfCad) = 0) Then
            
                infCad = auxInfCad(0)
                
                xMotivo = LerDadosJSON(infCad, xMotivo, "", "")
                
                txtRetornoXmotivo.Text = xMotivo

            Else
                Dim i As Integer
                For i = 0 To UBound(auxInfCad)
                    infCad = auxInfCad(i)

                    If (i <> UBound(auxInfCad)) Then
                        infCad = infCad & "}"
                    End If

                    IE = LerDadosJSON(infCad, "IE", "", "")
                    CNPJ = LerDadosJSON(infCad, "CNPJ", "", "")
                    UF = LerDadosJSON(infCad, "UF", "", "")
                    xNome = LerDadosJSON(infCad, "xNome", "", "")
                    xLgr = LerDadosJSON(infCad, "ender", "xLgr", "")
                    CEP = LerDadosJSON(infCad, "ender", "CEP", "")
                Next
            End If
        End If
        
        If (cStat <> "111") Then
            xMotivo = LerDadosJSON(retorno, "retConsCad", "infCons", "xMotivo")
            MsgBox (xMotivo)
        End If

        
    End If
     
    If (status <> 200) Then
        motivo = LerDadosJSON(retorno, "motivo", "", "")
        MsgBox (xMotivo)
    End If
    
Exit Sub
    
SAI:
    MsgBox (vbNewLine & Err.Description), vbInformation, titleNFeAPI
End Sub

Private Sub lista_banco_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With lista_banco
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If FunVerificaRegistroValidado("Clientes", "IDcliente = " & txtIDcliente, "cliente", "dados bancário", "excluir estes", True, True) = False Then
                .ListItems.Item(InitFor).Checked = False
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

Private Sub lista_familia_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With lista_familia
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If FunVerificaRegistroValidado("Clientes", "IDcliente = " & txtIDcliente, "cliente", "família", "excluir esta", True, True) = False Then
                .ListItems.Item(InitFor).Checked = False
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

Private Sub Lista_Impostos_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With Lista_Impostos
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If FunVerificaRegistroValidado("Clientes", "IDcliente = " & txtIDcliente, "cliente", "imposto", "excluir este", True, True) = False Then
                .ListItems.Item(InitFor).Checked = False
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

Private Sub Lista_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With Lista
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True And .ListItems.Item(InitFor) = Item Then
            If Cmb_opcao_lista = "Excluir" Then
                If .ListItems(InitFor).ListSubItems(4) = "Sim" Then
                    USMsgBox ("Não é permitido excluir este cliente, pois o mesmo está validado."), vbExclamation, "CAPRIND v5.0"
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
                End If
                
                Mensagem = "Não é possível excluir este cliente, pois o mesmo está sendo utilizado no módulo"
                IDCliente = .ListItems(InitFor)
                
                ProcVerificaRegistroUtilizado "Certificado_qualidade", "IDcliente = " & IDCliente, "Qualidade/Ensaios/Controle de certificados"
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
                End If
                ProcVerificaRegistroUtilizado "Estoque_Controle", "id_cliente = " & IDCliente, "Estoque/Consignação/Recebimento"
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
                End If
                ProcVerificaRegistroUtilizado "Liquido_penetrante", "IDcliente = " & IDCliente, "Qualidade/Ensaios/Líquido penetrante"
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
                End If
                ProcVerificaRegistroUtilizado "Producao", "Cliente = '" & .ListItems(InitFor).ListSubItems(3) & "'", "PCP/Gerenciamento de ordem"
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
                End If
                ProcVerificaRegistroUtilizado "Projproduto_clientes", "IDcliente = " & IDCliente, "Engenharia/Produtos e serviços"
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
                End If
                ProcVerificaRegistroUtilizado "tbl_contas_receber", "IDcliente = " & IDCliente, "Financeiro/Contas a receber"
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
                End If
                ProcVerificaRegistroUtilizado "tbl_Dados_Nota_Fiscal", "Id_Int_Cliente = " & IDCliente & " and txt_Razao_Nome = '" & .ListItems(InitFor).ListSubItems(3) & "'", "Faturamento/Nota fiscal"
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
                End If
                ProcVerificaRegistroUtilizado "vendas_proposta", "IDcliente = " & IDCliente, "Vendas/Proposta comercial"
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
                End If
                ProcVerificaRegistroUtilizado "Vendas_Tele", "IDcliente = " & IDCliente, "Vendas/Telemarketing"
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
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

Private Sub Txt_ID_CF_change()
On Error GoTo tratar_erro

txtIPI = ""
Set TBFI = CreateObject("adodb.recordset")
TBFI.Open "Select * from tbl_ClassificacaoFiscal where Idclass = " & IIf(Txt_ID_CF = "", 0, Txt_ID_CF), Conexao, adOpenKeyset, adLockOptimistic
If TBFI.EOF = False Then
    txtIPI = IIf(IsNull(TBFI!dbl_IPI), "", TBFI!dbl_IPI)
End If
TBFI.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbOrigem_Click()
On Error GoTo tratar_erro

ProcCarregaComboUF cmbuf, "UF is not null", cmbOrigem
ProcCarregaComboUF cmbuf_entrega, "UF is not null", cmbOrigem
ProcCarregaComboUF cmbuf_cobranca, "UF is not null", cmbOrigem
If cmbOrigem = "Estrangeiro" Then
    With txtcnpj
        .Text = "__.___.___/____-__"
        .Enabled = False
    End With
    With txtCpf
        .Text = "___.___.___-__"
        .Enabled = False
    End With
    cmbCidade.Visible = False
    cmbCidade_Entrega.Visible = False
    cmbCidade_cobranca.Visible = False
    txtCidade.Visible = True
    txtCidade_Entrega.Visible = True
    txtcidade_cobranca.Visible = True
Else
    Txt_pais.Text = "BRASIL"
    txtcnpj.Enabled = True
    txtCpf.Enabled = True
    cmbCidade.Visible = True
    cmbCidade_Entrega.Visible = True
    cmbCidade_cobranca.Visible = True
    txtCidade.Visible = False
    txtCidade_Entrega.Visible = False
    txtcidade_cobranca.Visible = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbPessoa_Click()
On Error GoTo tratar_erro

ProcVerifPessoa

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcVerifPessoa()
On Error GoTo tratar_erro

If Left(cmbPessoa, 6) = "Física" Then
    With txtCpf
        .Text = "___.___.___-__"
        .Enabled = True
        .Visible = True
    End With
    With txtcnpj
        .Text = "__.___.___/____-__"
        .Enabled = False
        .Visible = False
    End With
    With txtIM_IE
        .Text = ""
        .ToolTipText = "Inscrição estadual."
    End With
    Label2(0).Caption = "IE"
    With txtRG_IE
        .Text = ""
        .ToolTipText = "Registro geral."
    End With
    Label5.Visible = True
    Label1.Visible = False
    Label2(2).Caption = "RG"
    
    'Entrega
    Label9.Caption = "CPF"
    With txtCNPJ_entrega
        .Mask = "###.###.###-##"
        .ToolTipText = "Número do CPF."
    End With
    
    'Cobrança
    Label10.Caption = "CPF"
    With txtCNPJ_cobranca
        .Mask = "###.###.###-##"
        .ToolTipText = "Número do CPF."
    End With
    
    'Regime
    With cmbRegimeTributario
        .Clear
        .AddItem ""
        .AddItem "MEI"
    End With
    
    'Pessoa física não é contribuinte ICMS
    With Chk_nao_contribuinte_ICMS
       .Value = 1
       .Enabled = False
    End With
    
    With txtRG_IE
        .Text = "ISENTO"
        .Enabled = False
    End With
    
Else
    With txtCpf
        .Text = "___.___.___-__"
        .Enabled = False
        .Visible = False
    End With
    With txtcnpj
        .Text = "__.___.___/____-__"
        .Enabled = True
        .Visible = True
    End With
    With txtIM_IE
        .Text = ""
        .ToolTipText = "Inscrição municipal."
    End With
    Label2(0).Caption = "Insc. municipal"
    With txtRG_IE
        .Text = ""
        .ToolTipText = "Inscrição estadual."
    End With
    Label5.Visible = False
    Label1.Visible = True
    Label2(2).Caption = "Insc. estadual*"
    
    'Entrega
    Label9.Caption = "CNPJ"
    With txtCNPJ_entrega
        .Mask = "##.###.###/####-##"
        .ToolTipText = "Número do CNPJ."
    End With
    
    'Cobrança
    Label10.Caption = "CNPJ"
    With txtCNPJ_cobranca
        .Mask = "##.###.###/####-##"
        .ToolTipText = "Número do CNPJ."
    End With
    
    'Regime
    With cmbRegimeTributario
        .Clear
        .AddItem ""
        .AddItem "Lucro presumido"
        .AddItem "Lucro real"
        .AddItem "Simples nacional"
    End With
    
    'Pessoa física não é contribuinte ICMS
    With Chk_nao_contribuinte_ICMS
       .Value = 0
       .Enabled = True
    End With
    
    With txtRG_IE
        .Text = ""
        .Enabled = True
    End With
    
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procFiltrar_todos()
On Error GoTo tratar_erro

StrSql_Cliente = "Select * from clientes order by nomerazao"
FormulaRel_Cliente = ""
ProcCarregaLista (1)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLocalizar()
On Error GoTo tratar_erro

frmVendas_cliente_localizar.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcProximo()
On Error GoTo tratar_erro

If txtIDcliente = "" Then Exit Sub
Set TBClientes = CreateObject("adodb.recordset")
TBClientes.Open "Select * from clientes order by NomeRazao", Conexao, adOpenKeyset, adLockOptimistic
If TBClientes.BOF = False Then
    TBClientes.Find ("IDCliente = " & txtIDcliente)
    TBClientes.MoveNext
    If TBClientes.EOF = False Then
        txtIDcliente = TBClientes!IDCliente
        Set TBClientes = CreateObject("adodb.recordset")
        TBClientes.Open "Select * from clientes where IDCliente = " & IIf(txtIDcliente = "", 0, txtIDcliente), Conexao, adOpenKeyset, adLockOptimistic
        ProcLimpar
        procLimpacamposContatos
        ProcLimpacamposEntrega
        ProcLimpacamposCobranca
        ProcLimpaFamilia
        ProcLimpacampos_banco
        ProcLimpaCampos_Impostos
        ProcLimpaCamposComerciais
        ProcLimpaCampos_Outros
        ProcPuxaDados
        ProcCarregaListaContatos
        Proccarregalistaentrega
        ProcCarregalistacobranca
        ProcCarregaListaFamilia
        ProcCarregalista_banco
        ProcCarregaLista_Impostos
        procPuxadados_Comerciais
        ProcPuxaDados_Outros
    Else
        USMsgBox ("Fim dos cadastros de clientes."), vbInformation, "CAPRIND v5.0"
    End If
End If
Novo_Cliente1 = False
Novo_Cliente2 = False
Novo_Cliente3 = False
Novo_Cliente4 = False
Novo_Cliente5 = False
Novo_Cliente6 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procAtualizaClientesNuvem()
On Error GoTo tratar_erro

FunAbreBDSite
If ConexaoMySql.State = 1 Then

If USMsgBox("Deseja realmente atualizar todos os cadastros dos clientes na nuvem?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    
    Set TBAbrir = CreateObject("adodb.recordset")
    StrSql = "select * from Vendas_Clientes_Vendedores where vendedor = 'JAMES CARLETTI MARSON' order by NomeRazao"
    TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
    
    If TBAbrir.EOF = False Then
    'Debug.print TBAbrir.RecordCount
    
        Do While TBAbrir.EOF = False
            Set TBMySQL = New ADODB.Recordset
            '=================================================================
            ' Salvar Cliente na nuvem
            '=================================================================
            If TBAbrir!Tipo = "JP" Then
                StrSql = "Select * From Vendas_Clientes where CNPJ = '" & ReturnNumbersOnly(TBAbrir!CPF_CNPJ) & "' and CNPJ_Empresa = '" & CNPJ_Empresa & "'"
            Else
                StrSql = "Select * From Vendas_Clientes where CPF = '" & ReturnNumbersOnly(TBAbrir!CPF_CNPJ) & "' and CNPJ_Empresa = '" & CNPJ_Empresa & "'"
            End If
            
            'Debug.print StrSql
             TBMySQL.Open StrSql, ConexaoMySql, adOpenKeyset, adLockOptimistic, adCmdText
              If TBMySQL.EOF = False Then
                    TBMySQL.Fields!vendedor = TBAbrir!vendedor
                    TBMySQL.Update
                End If
                    
        TBAbrir.MoveNext
        Loop
        End If
        TBAbrir.Close
        USMsgBox ("Atualização efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
        '==================================
        Modulo = "Vendas/Clientes"
        Evento = "Atualizar clientes nuvem"
        ID_documento = 0
        Documento = ""
        Documento1 = ""
        ProcGravaEvento
        '==================================
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procAtualizaClienteNuvem()
On Error GoTo tratar_erro

FunAbreBDSite
If ConexaoMySql.State = 1 Then

If USMsgBox("Deseja realmente atualizar o vendedor no cadastro do cliente na nuvem?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
    
    Set TBAbrir = CreateObject("adodb.recordset")
    StrSql = "select CL.*, VV.Vendedor from Clientes CL Left join Vendas_Vendedores_Clientes VVC on VVC.IDCliente = CL.IDCliente left Join Vendas_Vendedores VV on VVC.IDVendedor = VV.Id where CL.IDCliente = '" & txtIDcliente.Text & "' and vv.vendedor = '" & cmbVendedor.Text & "' order by CL.NomeRazao"
    'Debug.print StrSql
    
    TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
    
    If TBAbrir.EOF = False Then
    'Debug.print TBAbrir.RecordCount
            Set TBMySQL = New ADODB.Recordset
            '=================================================================
            ' Salvar Cliente na nuvem
            '=================================================================
            If TBAbrir!Tipo = "JP" Then
                StrSql = "Select * From Vendas_Clientes where CNPJ = '" & ReturnNumbersOnly(TBAbrir!CPF_CNPJ) & "' and CNPJ_Empresa = '" & CNPJ_Empresa & "'"
            Else
                StrSql = "Select * From Vendas_Clientes where CPF = '" & ReturnNumbersOnly(TBAbrir!CPF_CNPJ) & "' and CNPJ_Empresa = '" & CNPJ_Empresa & "'"
            End If
            
            'Debug.print StrSql
             TBMySQL.Open StrSql, ConexaoMySql, adOpenKeyset, adLockOptimistic, adCmdText
              If TBMySQL.EOF = False Then
                    TBMySQL.Fields!vendedor = TBAbrir!vendedor
                    TBMySQL.Update
        End If
        TBAbrir.Close
        USMsgBox ("Vendedor salvo no cadastro do cliente Web com sucesso."), vbInformation, "CAPRIND v5.0"
        '==================================
        Modulo = "Vendas/Clientes"
        Evento = "Atualizar vendedor no cadastro do cliente na nuvem"
        ID_documento = 0
        Documento = ""
        Documento1 = ""
        ProcGravaEvento
        '==================================
    Else
    USMsgBox "Esse vendedor não está vinculado a esse cliente, altere o vendedor no cliente primeiro e depois atualize na WEB!", vbCritical, "CAPRIND v5,0"
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub ProcStatus()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With Lista
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then Permitido = True
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) cliente(s) antes de alterar o status."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
frmVendas_clientes_bloq.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbUF_Click()
On Error GoTo tratar_erro

If cmbOrigem.Text = "Nacional" And cmbuf.Text = "EX" Or cmbOrigem = "Estrangeiro" And cmbuf <> "EX" And cmbuf <> "" Then
    cmbuf.ListIndex = -1
    cmbCidade.Clear
    txtCidade.Text = ""
ElseIf cmbuf.Text <> "EX" Then
        ProcCarregaComboCidade cmbCidade, "Sigla_UF = '" & cmbuf & "'", False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbuf_cobranca_Click()
On Error GoTo tratar_erro

If cmbOrigem.Text = "Nacional" And cmbuf_cobranca.Text = "EX" Or cmbOrigem = "Estrangeiro" And cmbuf_cobranca <> "EX" And cmbuf_cobranca <> "" Then
    cmbuf_cobranca.ListIndex = -1
    cmbCidade_cobranca.Clear
    txtcidade_cobranca.Text = ""
ElseIf cmbuf_cobranca <> "EX" Then
        ProcCarregaComboCidade cmbCidade_cobranca, "Sigla_UF = '" & cmbuf_cobranca & "'", False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbuf_entrega_Click()
On Error GoTo tratar_erro

If cmbOrigem.Text = "Nacional" And cmbuf_entrega.Text = "EX" Or cmbOrigem = "Estrangeiro" And cmbuf_entrega <> "EX" And cmbuf_entrega <> "" Then
    cmbuf_entrega.ListIndex = -1
    cmbCidade_Entrega.Clear
    txtCidade_Entrega.Text = ""
ElseIf cmbuf_entrega <> "EX" Then
        ProcCarregaComboCidade cmbCidade_Entrega, "Sigla_UF = '" & cmbuf_entrega & "'", False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub CmdCF_Click()
On Error GoTo tratar_erro

Vendas_Proposta = False
Vendas_PI = False
Faturamento = False
Clientes = True
Compras_Pedido = False
Familia_NCM = False
ClassFiscal = False
frmProj_Classificacao_Fiscal.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdcfop_Click()
On Error GoTo tratar_erro

Clientes = True
Vendas_Proposta = False
Vendas_PI = False
Faturamento = False
Compras_Pedido = False
Sit_REG = 2
frm_ListaNatureza.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdCond_pag_padrao_Click()
On Error GoTo tratar_erro

Aplic = 1
Compras_Cotacao = False
Compras_Pedido = False
Vendas_Proposta = False
Vendas_PI = False
Estoque_recebimento = False
Clientes = True
Sit_REG = 0
frmCompras_pedido_DadosComerciais.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdDesenhos_calculos_padrao_Click()
On Error GoTo tratar_erro

Aplic = 3
Compras_Cotacao = False
Compras_Pedido = False
Vendas_Proposta = False
Vendas_PI = False
Estoque_recebimento = False
Clientes = True
Sit_REG = 0
frmCompras_pedido_DadosComerciais.Show

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
Permitido = False
With Lista
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) cliente(s)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            
            Conexao.Execute "DELETE from Clientes WHERE IDCLIENTE = " & .ListItems(InitFor) & ""
            Conexao.Execute "DELETE from Clientes_Contatos WHERE IDCLIENTE = " & .ListItems(InitFor) & ""
            Conexao.Execute "DELETE from clientes_entrega WHERE IDCLIENTE = " & .ListItems(InitFor) & ""
            Conexao.Execute "DELETE from clientes_cobranca WHERE IDCLIENTE = " & .ListItems(InitFor) & ""
            Conexao.Execute "DELETE from compras_fornecedores_familia where IDCLIENTE = " & .ListItems(InitFor) & " and tipo = 'C'"
            Conexao.Execute "DELETE from Compras_fornecedores_banco where id_fornecedor = " & .ListItems(InitFor) & " and tipo = 'C'"
            Conexao.Execute "DELETE from Clientes_Impostos where IDCLIENTE = " & .ListItems(InitFor) & ""
            
            '==================================
            Modulo = "Vendas/Clientes"
            Evento = "Excluir"
            ID_documento = .ListItems(InitFor)
            Documento = "Cliente: " & .ListItems(InitFor).SubItems(2)
            Documento1 = ""
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) cliente(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Cliente(s) excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    With txtIDcliente
        .Text = ""
        .Locked = False
        .TabStop = True
    End With
    ProcLimpar
    ProcCarregaLista (1)
    Frame1.Enabled = False
    ProcLimparTudo
    Novo_Cliente = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procExcluir_cobranca()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With listacobranca
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) local(is) de cobrança?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            
            Conexao.Execute "DELETE from clientes_cobranca where Idcobranca = " & .ListItems(InitFor)
            '==================================
            Modulo = "Vendas/Clientes"
            Evento = "Excluir local para cobrança"
            ID_documento = .ListItems(InitFor)
            Documento = "Cliente: " & txtnomerazao
            Documento1 = "Endereço: " & .ListItems(InitFor).SubItems(1)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) local(is) de cobrança antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Local(is) de cobrança excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcCarregalistacobranca
    ProcLimpacamposCobranca
    Frame8.Enabled = False
    Novo_Cliente3 = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procExcluir_banco()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With lista_banco
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) banco(s)?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            
            Conexao.Execute "DELETE from Compras_fornecedores_banco where id_fornecedor = " & IIf(txtIDcliente = "", 0, txtIDcliente) & " and id = " & .ListItems(InitFor) & " and tipo = 'C'"
            '==================================
            Modulo = "Vendas/Clientes"
            Evento = "Excluir banco"
            ID_documento = .ListItems(InitFor)
            Documento = "Cliente: " & txtnomerazao
            Documento1 = "Banco: " & .ListItems(InitFor).SubItems(1)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) banco(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Banco(s) excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpacampos_banco
    Frame10.Enabled = False
    ProcCarregalista_banco
    Novo_Cliente5 = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procExcluir_comerciais()
On Error GoTo tratar_erro

If USMsgBox("Deseja realmente excluir os dados comerciais do cliente?", vbYesNo) = vbYes Then
    Acao = "excluir os dados comerciais"
    If Cmb_empresa = "" Then
        NomeCampo = "empresa"
        ProcVerificaAcao
        Cmb_empresa.SetFocus
        Exit Sub
    End If
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select * FROM Clientes_DadosComerciais WHERE IDcliente = " & txtIDcliente.Text & " and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex), Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
        USMsgBox ("Dados comerciais excluídos com sucesso."), vbInformation, "CAPRIND v5.0"
        '==================================
        Modulo = "Vendas/Clientes"
        Evento = "Excluir dados comerciais"
        ID_documento = TBProduto!ID
        Documento = "Cliente: " & txtnomerazao & " - Empresa: " & Cmb_empresa
        Documento1 = ""
        ProcGravaEvento
        '==================================
        TBProduto.Delete
    End If
    ProcLimpaCamposComerciais
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procExcluir_entrega()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With ListaEntrega
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) local(is) de entrega?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            
            Conexao.Execute "DELETE from clientes_entrega where Identrega = " & .ListItems(InitFor)
            '==================================
            Modulo = "Vendas/Clientes"
            Evento = "Excluir local para entrega"
            ID_documento = .ListItems(InitFor)
            Documento = "Cliente: " & txtnomerazao
            Documento1 = "Endereço: " & .ListItems(InitFor).SubItems(1)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) local(is) de entrega antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Local(is) de entrega excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    Proccarregalistaentrega
    ProcLimpacamposEntrega
    Frame6.Enabled = False
    Novo_Cliente2 = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procExcluir_familia()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With lista_familia
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir esta(s) família(s)?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            
            Conexao.Execute "DELETE from compras_fornecedores_familia where idfamilia = " & .ListItems(InitFor) & " and tipo = 'C'"
            '==================================
            Modulo = "Vendas/Clientes"
            Evento = "Excluir família"
            ID_documento = .ListItems(InitFor)
            Documento = "Cliente: " & txtnomerazao
            Documento1 = "Família: " & .ListItems(InitFor).SubItems(1)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe a(s) família(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Família(s) excluída(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    cmbfamilia.ListIndex = -1
    txtid_familia = 0
    Frame12.Enabled = False
    ProcCarregaListaFamilia
    Novo_Cliente4 = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procExcluir_impostos()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With Lista_Impostos
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) imposto(s)?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            
            Conexao.Execute "DELETE from Clientes_Impostos where Id = " & .ListItems(InitFor)
            '==================================
            Modulo = "Vendas/Clientes"
            Evento = "Excluir imposto"
            ID_documento = .ListItems(InitFor)
            Documento = "Cliente: " & txtnomerazao
            Documento1 = "NCM: " & .ListItems(InitFor).SubItems(1)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) imposto(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Imposto(s) excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCampos_Impostos
    Frame14.Enabled = False
    ProcCarregaLista_Impostos
    Novo_Cliente6 = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procExcluir_contato()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With Lista_contato
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) contato(s)?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            Conexao.Execute "DELETE from clientes_contatos where Idcontato = " & .ListItems(InitFor)
            '==================================
            Modulo = "Vendas/Clientes"
            Evento = "Excluir contato"
            ID_documento = .ListItems(InitFor)
            Documento = "Cliente: " & txtnomerazao
            Documento1 = "Contato: " & .ListItems(InitFor).SubItems(1)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) contato(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Contato(s) excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    Frame2.Enabled = False
    ProcCarregaListaContatos
    procLimpacamposContatos
    Novo_Cliente1 = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdGarantia_padrao_Click()
On Error GoTo tratar_erro

Aplic = 8
Compras_Cotacao = False
Compras_Pedido = False
Vendas_Proposta = False
Vendas_PI = False
Estoque_recebimento = False
Clientes = True
Sit_REG = 0
frmCompras_pedido_DadosComerciais.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdGrupo_Click()
On Error GoTo tratar_erro
  
frmVendas_cliente_grupos.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdImpostos_padrao_Click()
On Error GoTo tratar_erro

Aplic = 7
Compras_Cotacao = False
Compras_Pedido = False
Vendas_Proposta = False
Vendas_PI = False
Estoque_recebimento = False
Clientes = True
Sit_REG = 0
frmCompras_pedido_DadosComerciais.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdLocTransp_Click()
On Error GoTo tratar_erro

With Cmb_tipo_transp
    If .Text = "Cliente" Then
        ProcConfVariaveisLocCliente True, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False
        frmVendas_LocalizarCliente.Show 1
    ElseIf .Text = "Fornecedor" Then
            ProcConfVariaveisLocForn True, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False
            FrmCompras_localizafornecedor.Show 1
        Else
            frmFaturamento_Prod_Serv_Localizar_Empresa.Show 1
    End If
End With

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
With txtIDcliente
    .Text = ""
    .Locked = True
    .TabStop = False
End With
ProcLimpar
Novo_Cliente = True
Frame1.Enabled = True
cmbPessoa.SetFocus
ProcLimparTudo
USMsgBox "Após escolher o Tipo a origem o Pais a UF e digitar o CNPJ, clique no botão ao lado do CNPJ para buscar os dados do novo cliente na receita federal.", vbInformation, "CAPRIND v5.0"
cmbVendedor.Visible = True
cmbVendedor.ListIndex = -1
txtComissao.Text = ""
txtIDVendedor.Text = ""
txtVendedor.Visible = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procNovo_cobranca()
On Error GoTo tratar_erro

If USMsgBox("Deseja aproveitar os dados principais do cliente?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    txtID_cobranca = 0
    If Left(cmbPessoa, 8) = "Jurídica" Then txtCNPJ_cobranca = txtcnpj.Text Else txtCNPJ_cobranca = txtCpf
    If cmbTipo_endereco <> "" Then cmbTipo_endereco_cobranca = cmbTipo_endereco
    txtendereco_cobranca = txtendereco
    txtNumero_cobranca = txtNumero
    txtComplemento_cobranca = txtComplemento
    If cmbTipo_bairro <> "" Then cmbTipo_bairro_cobranca = cmbTipo_bairro
    txtbairro_cobranca = txtBairro
    If cmbuf <> "EX" Then
        cmbuf_cobranca = cmbuf
        cmbCidade_cobranca.Visible = True
        txtcidade_cobranca.Visible = False
        cmbCidade_cobranca = cmbCidade
    Else
        cmbuf_cobranca = cmbuf
        cmbCidade_cobranca.Visible = False
        txtcidade_cobranca.Visible = True
        txtcidade_cobranca = txtCidade
    End If
    txtcxpostal_cobranca = txtcaixapostal
    mskcep_cobranca = txtCEP
    txttel1_cobranca = txttel01
    txttel2_cobranca = txttel02
    txttel3_cobranca = txttel03
    txttel4_cobranca = txttel04
    txtfax_cobranca = txtFax
    txtemail_cobranca.Text = txtEmail.Text
    txtSite_cobranca.Text = txtSite.Text
    CodigoLista3 = 0
Else
    ProcLimpacamposCobranca
End If
Frame8.Enabled = True
txtendereco_cobranca.SetFocus
Novo_Cliente3 = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procNovo_Banco()
On Error GoTo tratar_erro

ProcLimpacampos_banco
Novo_Cliente5 = True
Frame10.Enabled = True
txtBanco.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procNovo_Contato()
On Error GoTo tratar_erro

procLimpacamposContatos
Novo_Cliente1 = True
Frame2.Enabled = True
txtNomeContato.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimpacamposEntrega()
On Error GoTo tratar_erro

txtID_entrega.Text = 0
txtData2 = Format(Date, "dd/mm/yy")
txtResponsavel2 = pubUsuario
If cmbPessoa <> "" Then
    If Left(cmbPessoa, 8) = "Jurídica" Then txtCNPJ_entrega.Text = "__.___.___/____-__" Else txtCNPJ_entrega.Text = "___.___.___-__"
End If
cmbTipo_endereco_entrega.ListIndex = -1
txtEndereco_entrega.Text = ""
txtNumero_entrega.Text = ""
txtComplemento_entrega = ""
cmbTipo_bairro_entrega.ListIndex = -1
txtBairro_entrega.Text = ""
cmbCidade_Entrega.ListIndex = -1
txtCidade_Entrega.Text = ""
cmbuf_entrega.ListIndex = -1
txtcxpostal_entrega.Text = ""
mskcep_entrega.Text = ""
txttel1_entrega.Text = ""
txttel2_entrega.Text = ""
txttel3_entrega.Text = ""
txttel4_entrega.Text = ""
txtfax_entrega.Text = ""
txtemail_entrega.Text = ""
txtsite_entrega.Text = ""
CodigoLista2 = 0
 
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimpacamposCobranca()
On Error GoTo tratar_erro

txtID_cobranca.Text = 0
txtData3 = Format(Date, "dd/mm/yy")
txtResponsavel3 = pubUsuario
If cmbPessoa <> "" Then
    If Left(cmbPessoa, 8) = "Jurídica" Then txtCNPJ_cobranca.Text = "__.___.___/____-__" Else txtCNPJ_cobranca.Text = "___.___.___-__"
End If
cmbTipo_endereco_cobranca.ListIndex = -1
txtendereco_cobranca.Text = ""
txtNumero_cobranca.Text = ""
txtComplemento_cobranca = ""
cmbTipo_bairro_cobranca.ListIndex = -1
txtbairro_cobranca.Text = ""
txtcidade_cobranca.Text = ""
cmbCidade_cobranca.ListIndex = -1
cmbuf_cobranca.ListIndex = -1
txtcxpostal_cobranca.Text = ""
mskcep_cobranca = ""
txttel1_cobranca.Text = ""
txttel2_cobranca.Text = ""
txttel3_cobranca.Text = ""
txttel4_cobranca.Text = ""
txtfax_cobranca.Text = ""
txtemail_cobranca.Text = ""
CodigoLista3 = 0
 
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procNovo_entrega()
On Error GoTo tratar_erro

If USMsgBox("Deseja aproveitar os dados principais do cliente?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    txtID_entrega = 0
    If Left(cmbPessoa, 8) = "Jurídica" Then txtCNPJ_entrega = txtcnpj.Text Else txtCNPJ_entrega = txtCpf
    If cmbTipo_endereco <> "" Then cmbTipo_endereco_entrega = cmbTipo_endereco
    txtEndereco_entrega = txtendereco
    txtNumero_entrega = txtNumero
    txtComplemento_entrega = txtComplemento
    If cmbTipo_bairro <> "" Then cmbTipo_bairro_entrega = cmbTipo_bairro
    txtBairro_entrega = txtBairro
    If cmbuf <> "EX" Then
        cmbuf_entrega = cmbuf
        cmbCidade_Entrega.Visible = True
        txtCidade_Entrega.Visible = False
        cmbCidade_Entrega = cmbCidade
    Else
        cmbuf_entrega = cmbuf
        cmbCidade_Entrega.Visible = False
        txtCidade_Entrega.Visible = True
        txtCidade_Entrega = txtCidade
    End If
    txtcxpostal_entrega = txtcaixapostal
    mskcep_entrega = txtCEP
    txttel1_entrega = txttel01
    txttel2_entrega = txttel02
    txttel3_entrega = txttel03
    txttel4_entrega = txttel04
    txtfax_entrega = txtFax
    txtemail_entrega.Text = txtEmail.Text
    txtsite_entrega.Text = txtSite.Text
    CodigoLista3 = 0
Else
    ProcLimpacamposEntrega
End If
Frame6.Enabled = True
txtEndereco_entrega.SetFocus
Novo_Cliente2 = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procNovo_familia()
On Error GoTo tratar_erro
  
txtid_familia = 0
cmbfamilia.ListIndex = -1
Novo_Cliente4 = True
Frame12.Enabled = True
cmbfamilia.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procNovo_impostos()
On Error GoTo tratar_erro

ProcLimpaCampos_Impostos
Novo_Cliente6 = True
Frame14.Enabled = True
CmdCF_Click

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimpaCampos_Impostos()
On Error GoTo tratar_erro

txtid_impostos = 0
txtData6 = Format(Date, "dd/mm/yy")
txtResponsavel6 = pubUsuario
Txt_ID_CF = ""
Txt_CF = ""
txtIPI = ""
txtPorcentagemIPI = ""
CodigoLista6 = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimpaCampos_Outros()
On Error GoTo tratar_erro

txtObservacoes.Text = ""
txttel02.Text = ""
txttel03.Text = ""
txttel04.Text = ""
Txt_ISSQN = ""
txtIDGrupo = "0"
txtGrupo = ""
chkSuframa.Value = 0
txtSuframa = ""
chkICMSST.Value = 0
cmbBanco.ListIndex = -1
cmbTipo_doc.ListIndex = -1
txtLimiteCredito = ""
CodigoLista8 = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Cliente.AbsolutePage <> 2 Then
    If TBLISTA_Cliente.AbsolutePage = -3 Then
        ProcExibePagina (TBLISTA_Cliente.PageCount - 1)
    Else
        TBLISTA_Cliente.AbsolutePage = TBLISTA_Cliente.AbsolutePage - 2
        ProcExibePagina (TBLISTA_Cliente.AbsolutePage)
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
    TBLISTA_Cliente.AbsolutePage = txtPagIr.Text
    ProcExibePagina (TBLISTA_Cliente.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Cliente.AbsolutePage = 1
ProcExibePagina (TBLISTA_Cliente.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Cliente.AbsolutePage <> -3 Then
    If TBLISTA_Cliente.AbsolutePage = 1 Then
        ProcExibePagina (2)
    Else
        ProcExibePagina (TBLISTA_Cliente.AbsolutePage)
    End If
Else
    ProcExibePagina (TBLISTA_Cliente.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Cliente.AbsolutePage = TBLISTA_Cliente.PageCount
ProcExibePagina (TBLISTA_Cliente.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdReajuste_padrao_Click()
On Error GoTo tratar_erro

Aplic = 9
Compras_Cotacao = False
Compras_Pedido = False
Vendas_Proposta = False
Vendas_PI = False
Estoque_recebimento = False
Clientes = True
Sit_REG = 0
frmCompras_pedido_DadosComerciais.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

If Novo_Cliente = True Then
    If USMsgBox("O cliente ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvar
        If Novo_Cliente = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
If Novo_Cliente1 = True Then
    If USMsgBox("O contato ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        procSalvar_contato
        If Novo_Cliente1 = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
If Novo_Cliente2 = True Then
    If USMsgBox("O local para entrega ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        procSalvar_entrega
        If Novo_Cliente2 = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
If Novo_Cliente3 = True Then
    If USMsgBox("O local para cobrança ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        procSalvar_cobranca
        If Novo_Cliente3 = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
If Novo_Cliente4 = True Then
    If USMsgBox("A família ainda não foi salva, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        procSalvar_familia
        If Novo_Cliente4 = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
If Novo_Cliente5 = True Then
    If USMsgBox("O banco ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        procSalvar_comerciais
        If Novo_Cliente5 = True Then
            Exit Sub
        Else
            Unload Me
        End If
    Else
        If txtID_banco <> "" Then
            Sair = True
            procExcluir_comerciais
        End If
    End If
End If
If Novo_Cliente6 = True Then
    If USMsgBox("O imposto ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        procSalvar_impostos
        If Novo_Cliente6 = True Then
            Exit Sub
        Else
            Unload Me
        End If
    Else
        If txtid_impostos <> "" Then
            Sair = True
            procExcluir_impostos
        End If
    End If
End If
Novo_Cliente = False
Novo_Cliente1 = False
Novo_Cliente2 = False
Novo_Cliente3 = False
Novo_Cliente4 = False
Novo_Cliente5 = False
Novo_Cliente6 = False
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
If Frame1.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
If cmbOrigem = "Nacional" And Left(cmbPessoa, 8) = "Jurídica" Then
    If cmbRegimeTributario.Text = "" Then
        If USMsgBox("O cliente será cadastrado sem regime tributário, deseja prosseguir assim mesmo?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
    End If
End If

Acao = "salvar"
If txtnomerazao.Text = "" Then
    NomeCampo = "o nome"
    ProcVerificaAcao
    txtnomerazao.SetFocus
    Exit Sub
End If
If Chk_prospecto.Value = 0 And Chk_enviar_NF.Value = 1 Then
    If cmbPessoa.Text = "" Then
        NomeCampo = "o tipo"
        ProcVerificaAcao
        If Frame1.Enabled = True Then cmbPessoa.SetFocus
        Exit Sub
    End If
    If txtCategoria.Text = "" Then
        NomeCampo = "a categoria"
        ProcVerificaAcao
        txtCategoria.SetFocus
        Exit Sub
    End If
    If cmbOrigem.Text = "" Then
        NomeCampo = "a origem"
        ProcVerificaAcao
        cmbOrigem.SetFocus
        Exit Sub
    End If
    If cmbOrigem = "Nacional" Then
        If Left(cmbPessoa, 8) = "Jurídica" And txtcnpj = "__.___.___/____-__" Then
            NomeCampo = "o CNPJ"
            ProcVerificaAcao
            txtcnpj.SetFocus
            Exit Sub
        ElseIf Left(cmbPessoa, 6) = "Física" And txtCpf.Text = "___.___.___-__" Then
                NomeCampo = "o CPF"
                ProcVerificaAcao
                txtCpf.SetFocus
                Exit Sub
        End If
    End If
    If Txt_pais = "" Then
        NomeCampo = "o país"
        ProcVerificaAcao
        Txt_pais.SetFocus
        Exit Sub
    End If
    If txtendereco <> "" And cmbOrigem = "Nacional" And cmbTipo_endereco = "" Then
        NomeCampo = "o tipo"
        ProcVerificaAcao
        cmbTipo_endereco.SetFocus
        Exit Sub
    End If
    If txtendereco = "" Then
        NomeCampo = "o endereço"
        ProcVerificaAcao
        txtendereco.SetFocus
        Exit Sub
    End If
    If txtNumero.Text = "" Then
        NomeCampo = "o número"
        ProcVerificaAcao
        txtNumero.SetFocus
        Exit Sub
    End If
    If txtBairro <> "" And cmbOrigem = "Nacional" And cmbTipo_bairro = "" Then
        NomeCampo = "o tipo"
        ProcVerificaAcao
        cmbTipo_bairro.SetFocus
        Exit Sub
    End If
    If txtBairro = "" Then
        NomeCampo = "o bairro"
        ProcVerificaAcao
        txtBairro.SetFocus
        Exit Sub
    End If
    If cmbOrigem = "Nacional" Then
        If cmbuf.Text = "" Then
            NomeCampo = "o estado"
            ProcVerificaAcao
            cmbuf.SetFocus
            Exit Sub
        End If
        If cmbuf.Text = "EX" Then
            USMsgBox ("Não é permitido informar UF de exportação para cliente nacional."), vbInformation, "CAPRIND v5.0"
            cmbuf.SetFocus
            Exit Sub
        End If
        If cmbCidade = "" Then
            NomeCampo = "a cidade"
            ProcVerificaAcao
            cmbCidade.SetFocus
            Exit Sub
        End If
        'If FunVerificaCidade(cmbCidade, cmbuf) = False Then Exit Sub
        
        If txtCEP.Text = "" Then
            NomeCampo = "o CEP"
            ProcVerificaAcao
            txtCEP.SetFocus
            Exit Sub
        End If
    End If
End If

If txtcnpj.Text <> "__.___.___/____-__" Then
    If Funconsistir_CgcCpf(txtcnpj) = False Then
        USMsgBox ("O número do CNPJ digitado não é válido, digite o número correto."), vbExclamation, "CAPRIND v5.0"
        txtcnpj.SetFocus
        Exit Sub
    End If
    Set TBProposta = CreateObject("adodb.recordset")
    TBProposta.Open "Select * from clientes where cpf_cnpj = '" & txtcnpj.Text & "' and idcliente <> " & IIf(txtIDcliente = "", 0, txtIDcliente), Conexao, adOpenKeyset, adLockOptimistic
    If TBProposta.EOF = False Then
        If Novo_Cliente = True Or TBProposta!NomeRazao <> txtnomerazao.Text Then
            USMsgBox ("Já existe cadastro deste CNPJ para o cliente " & TBProposta("Nomerazao") & ", favor alterar o número do CNPJ."), vbExclamation, "CAPRIND v5.0"
            txtcnpj.Text = "__.___.___/____-__"
            txtcnpj.SetFocus
            TBProposta.Close
            Exit Sub
        End If
    End If
    TBProposta.Close
ElseIf txtCpf.Text <> "___.___.___-__" Then
        Set TBClientes = CreateObject("adodb.recordset")
        TBClientes.Open "Select * from clientes where cpf_cnpj = '" & txtCpf.Text & "' and idcliente <> " & IIf(txtIDcliente = "", 0, txtIDcliente), Conexao, adOpenKeyset, adLockOptimistic
        If TBClientes.EOF = False Then
            If Novo_Cliente = True Or TBClientes!NomeRazao <> txtnomerazao.Text Then
                USMsgBox ("Já existe cadastro deste CPF para o cliente " & TBClientes("Nomerazao") & ", favor alterar o número do CPF."), vbExclamation, "CAPRIND v5.0"
                txtCpf = "___.___.___-__"
                txtCpf.SetFocus
                TBClientes.Close
                Exit Sub
            End If
        End If
        TBClientes.Close
End If

Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from clientes where IDCLIENTE = " & IIf(txtIDcliente = "", 0, txtIDcliente), Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then
    TBGravar.AddNew
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select IDCliente from clientes order by idcliente desc", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        txtIDcliente.Text = TBAbrir!IDCliente + 1
    Else
        txtIDcliente.Text = 1
    End If
    TBAbrir.Close
    TBGravar!IDCliente = txtIDcliente.Text
    TBGravar!status = "Liberado"
Else
    If FunVerifValidacaoRegistro("alterar", txtDtValidacao, "mesmo", "o cliente", True) = False Then Exit Sub
    If txtnomerazao <> TBGravar!NomeRazao Or cmbuf <> TBGravar!UF Or cmbCidade <> TBGravar!Cidade Then
        If USMsgBox("Deseja atualizar os dados deste cliente em todos os módulos?", vbYesNo, "CAPRIND v5.0") = vbYes Then
            If txtnomerazao <> TBGravar!NomeRazao Or cmbuf <> TBGravar!UF Or cmbCidade <> TBGravar!Cidade Then Conexao.Execute "Update tbl_contas_receber Set Nome_Razao = '" & txtnomerazao & "', Cidade = '" & cmbCidade & "', Estado = '" & cmbuf & "' where idcliente = " & IIf(txtIDcliente = "", 0, txtIDcliente) & " and Tipo = 'CL'"
            
            If txtnomerazao <> TBGravar!NomeRazao Then
                Conexao.Execute "Update Certificado_qualidade Set Cliente = '" & txtnomerazao & "' where cliente = '" & TBGravar!NomeRazao & "'"
                Conexao.Execute "Update CQ_RNC Set Cliente_forn = '" & txtnomerazao & "' where ID_forn = " & IIf(txtIDcliente = "", 0, txtIDcliente) & " and Tipo = 'C'"
                Conexao.Execute "Update Estoque_Controle Set Cliente = '" & txtnomerazao & "' where id_cliente = " & IIf(txtIDcliente = "", 0, txtIDcliente)
                Conexao.Execute "Update EM set EM.Cliente = '" & txtnomerazao & "' from Estoque_movimentacao EM INNER JOIN estoque_controle EC ON EM.IDestoque = EC.Idestoque where EC.id_cliente = " & IIf(txtIDcliente = "", 0, txtIDcliente)
                Conexao.Execute "Update Liquido_penetrante Set Cliente = '" & txtnomerazao & "' where IDcliente = " & IIf(txtIDcliente = "", 0, txtIDcliente)
                Conexao.Execute "Update Producao Set Cliente = '" & txtnomerazao & "' where IDcliente = " & IIf(txtIDcliente = "", 0, txtIDcliente)
                Conexao.Execute "Update item_aplicacoes Set aplicacao = '" & txtnomerazao & "' where ID_cliente_forn = " & IIf(txtIDcliente = "", 0, txtIDcliente) & " and Tipo = 'C'"
                
                Conexao.Execute "Update tbl_Dados_Nota_Fiscal Set txt_Razao_Nome = '" & txtnomerazao & "' where Id_Int_Cliente = " & IIf(txtIDcliente = "", 0, txtIDcliente) & " and txt_Razao_Nome = '" & TBGravar!NomeRazao & "'"
                Conexao.Execute "Update Vendas_analise Set Cliente = '" & txtnomerazao & "' where idcliente = " & IIf(txtIDcliente = "", 0, txtIDcliente)
                Conexao.Execute "Update Vendas_proposta Set Cliente = '" & txtnomerazao & "' where idcliente = " & IIf(txtIDcliente = "", 0, txtIDcliente)
                Conexao.Execute "Update Vendas_Tele Set Cliente = '" & txtnomerazao & "' where IDcliente = " & IIf(txtIDcliente = "", 0, txtIDcliente)
                Conexao.Execute "Update UltraSom Set Cliente = '" & txtnomerazao & "' where IDcliente = " & IIf(txtIDcliente = "", 0, txtIDcliente)
                
                'Transportadora
                Conexao.Execute "Update Compras_fornecedores Set Transportadora = '" & txtnomerazao & "' where IDTransp = " & IIf(txtIDcliente = "", 0, txtIDcliente) & " and Tipo_transp = 'C'"
                Conexao.Execute "Update vendas_comercial Set Transportadora = '" & txtnomerazao & "' where IDInttransp = " & IIf(txtIDcliente = "", 0, txtIDcliente) & " and Tipo_transp = 'C'"
                Conexao.Execute "Update tbl_Dados_Transp Set txt_Razao = '" & txtnomerazao & "' where IdIntTransp = " & IIf(txtIDcliente = "", 0, txtIDcliente) & " and Tipo_transp = 'C'"
            End If
        End If
    End If
End If

'    Set TBAbrir = CreateObject("adodb.recordset")
'    TBAbrir.Open "Select * from Vendas_Vendedores_Clientes Where IDCliente = " & txtidcliente.Text & "", Conexao, adOpenKeyset, adLockOptimistic
'    If TBAbrir.EOF = False Then
'    USMsgBox "Esse cliente já pertence a um vendedor, não é permitido alterar o vendedor, fale com o responsável", vbInformation, "CAPRIND v5.0"
'    GoTo Continua
'    End If
'    TBAbrir.Close

'===============================================================================
' Grava o vendedor interno no cliente
'===============================================================================
If txtIDVendedor.Text <> "" Then
    Set TBAbrir = CreateObject("adodb.recordset")
     TBAbrir.Open "Select * from Vendas_Vendedores_Clientes Where IDCliente = " & txtIDcliente.Text & "", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = True Then
    TBAbrir.AddNew
    End If
   
    TBAbrir!IDvendedor = txtIDVendedor.Text
    TBAbrir!IDCliente = Int(txtIDcliente.Text)
    TBAbrir!Comissao = IIf(txtComissao.Text <> "", txtComissao.Text, 0)
    TBAbrir!tipocomissao = "C"
    TBAbrir.Update
    TBAbrir.Close
'===============================================================================
End If

Continua:

ProcEnviaDados
TBGravar.Update
TBGravar.Close

If Novo_Cliente = True Then
    USMsgBox ("Novo cliente cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo"
    StrSql_Cliente = "Select * from clientes where idcliente = " & txtIDcliente
    ProcCarregaLista (1)
    
    Novo_Cliente = False
    USMsgBox "Não se esqueça de cadastrar o(s) contato(s) para este cliente.", vbInformation, "CAPRIND v5.0"
    SSTab1.Tab = 1
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar"
    ProcCarregaLista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
    If Lista.ListItems.Count <> 0 And CodigoLista <> 0 Then
        Lista.SelectedItem = Lista.ListItems(CodigoLista)
        Lista.SetFocus
    End If
End If
1:
    '==================================
    Modulo = "Vendas/Clientes"
    ID_documento = txtIDcliente
    Documento = "Cliente: " & txtnomerazao
    Documento1 = ""
    ProcGravaEvento
    '==================================
    With txtIDcliente
        .Locked = False
        .TabStop = True
    End With
'    With cmbPessoa
'        .Locked = True
'        .TabStop = False
'    End With
    Novo_Cliente = False

Exit Sub
tratar_erro:
    If Err.Number = "35600" Then GoTo 1
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procSalvar_cobranca()
On Error GoTo tratar_erro

If Frame8.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
If Left(cmbPessoa, 8) = "Jurídica" Then
    If txtCNPJ_cobranca.Text <> "__.___.___/____-__" Then
        If Funconsistir_CgcCpf(txtCNPJ_cobranca) = False Then
            USMsgBox ("O número do CNPJ digitado não é válido, digite o número correto."), vbExclamation, "CAPRIND v5.0"
            txtCNPJ_cobranca.SetFocus
            Exit Sub
        End If
    End If
End If
Acao = "salvar"
If txtendereco_cobranca <> "" And cmbRegimeTributario.Text = "Simples nacional" And cmbTipo_endereco_cobranca = "" Then
    NomeCampo = "o tipo"
    ProcVerificaAcao
    cmbTipo_endereco_cobranca.SetFocus
    Exit Sub
End If
If txtendereco_cobranca.Text = "" Then
    NomeCampo = "o endereço"
    ProcVerificaAcao
    txtendereco_cobranca.SetFocus
    Exit Sub
End If
If txtbairro_cobranca <> "" And cmbRegimeTributario.Text = "Simples nacional" And cmbTipo_bairro_cobranca = "" Then
    NomeCampo = "o tipo"
    ProcVerificaAcao
    cmbTipo_bairro_cobranca.SetFocus
    Exit Sub
End If
If cmbCidade_cobranca <> "" And cmbuf_cobranca <> "" And cmbuf_cobranca <> "EX" Then
    If cmbRegimeTributario.Text = "Simples nacional" And FunVerificaCidade(cmbCidade_cobranca, cmbuf_cobranca) = False Then Exit Sub
End If
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from clientes_cobranca where idcobranca = " & txtID_cobranca, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then
    TBGravar.AddNew
Else
    If FunVerificaRegistroValidado("Clientes", "IDcliente = " & txtIDcliente, "cliente", "este endereço de cobrança", "alterar", True, True) = False Then Exit Sub
End If
TBGravar!IDCliente = txtIDcliente
If txtData3 = "" Then TBGravar!Data = Date Else TBGravar!Data = txtData3
If txtResponsavel3 = "" Then TBGravar!Responsavel = pubUsuario Else TBGravar!Responsavel = txtResponsavel3
TBGravar!Tipo = "C"
TBGravar!CNPJ = txtCNPJ_cobranca.Text
TBGravar!Tipo_endereco = cmbTipo_endereco_cobranca
TBGravar!endereco_Cobranca = txtendereco_cobranca
TBGravar!Numero = txtNumero_cobranca.Text
TBGravar!complemento = txtComplemento_cobranca
TBGravar!Tipo_bairro = cmbTipo_bairro_cobranca
TBGravar!bairro_Cobranca = txtbairro_cobranca
If cmbuf_cobranca <> "EX" Then TBGravar!cidade_Cobranca = IIf(cmbCidade_cobranca.Text = "", Null, cmbCidade_cobranca.Text) Else TBGravar!cidade_Cobranca = IIf(txtcidade_cobranca.Text = "", Null, txtcidade_cobranca.Text)
TBGravar!uf_Cobranca = cmbuf_cobranca
TBGravar!cxpostal_Cobranca = txtcxpostal_cobranca
TBGravar!cep_Cobranca = mskcep_cobranca
TBGravar!tel1_Cobranca = txttel1_cobranca
TBGravar!tel2_Cobranca = txttel2_cobranca
TBGravar!tel3_Cobranca = txttel3_cobranca
TBGravar!tel4_Cobranca = txttel4_cobranca
TBGravar!fax_Cobranca = txtfax_cobranca
TBGravar!email_Cobranca = IIf(txtemail_cobranca.Text = "", Null, LCase(txtemail_cobranca.Text))
TBGravar!site_cobranca = IIf(txtSite_cobranca.Text = "", Null, LCase(txtSite_cobranca.Text))
TBGravar.Update
txtID_cobranca = TBGravar!idCobranca
TBGravar.Close
ProcCarregalistacobranca
If Novo_Cliente3 = True Then
    USMsgBox ("Novo local para cobrança cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo local para cobrança"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar local para cobrança"
    If listacobranca.ListItems.Count <> 0 And CodigoLista3 <> 0 Then
        listacobranca.SelectedItem = listacobranca.ListItems(CodigoLista3)
        listacobranca.SetFocus
    End If
End If
'==================================
Modulo = "Vendas/Clientes"
ID_documento = txtID_cobranca
Documento = "Cliente: " & txtnomerazao
Documento1 = "Endereço: " & txtendereco_cobranca
ProcGravaEvento
'==================================
Novo_Cliente3 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procSalvar_Banco()
On Error GoTo tratar_erro

If Frame10.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
Acao = "salvar"
If txtBanco = "" Then
    NomeCampo = "o banco"
    ProcVerificaAcao
    txtBanco.SetFocus
    Exit Sub
End If
If txtAgencia = "" Then
    NomeCampo = "a agência"
    ProcVerificaAcao
    txtAgencia.SetFocus
    Exit Sub
End If
If txtConta = "" Then
    NomeCampo = "a conta"
    ProcVerificaAcao
    txtConta.SetFocus
    Exit Sub
End If
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Compras_fornecedores_banco where banco = '" & txtBanco & "' and id <> " & txtID_banco & " and id_fornecedor = " & IIf(txtIDcliente = "", 0, txtIDcliente) & " and Tipo = 'C'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    USMsgBox ("Este banco já está cadastrado."), vbExclamation, "CAPRIND v5.0"
    txtBanco.SetFocus
    Exit Sub
End If
TBAbrir.Close
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from Compras_fornecedores_banco where id = " & txtID_banco, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then
    TBGravar.AddNew
Else
    If FunVerificaRegistroValidado("Clientes", "IDcliente = " & txtIDcliente, "cliente", "estes dados bancários", "alterar", True, True) = False Then Exit Sub
End If
TBGravar!id_fornecedor = txtIDcliente
If txtData5 = "" Then TBGravar!Data = Date Else TBGravar!Data = txtData5
If txtResponsavel5 = "" Then TBGravar!Responsavel = pubUsuario Else TBGravar!Responsavel = txtResponsavel5
TBGravar!Tipo = "C"
TBGravar!Banco = txtBanco
TBGravar!Agencia = txtAgencia
TBGravar!Conta = txtConta
TBGravar.Update
txtID_banco = TBGravar!ID
TBGravar.Close
ProcCarregalista_banco
If Novo_Cliente5 = True Then
    USMsgBox "Novo banco cadastrado com sucesso.", vbInformation, "CAPRIND v5.0"
    Evento = "Novo banco"
Else
    USMsgBox "Alteração efetuada com sucesso.", vbInformation, "CAPRIND v5.0"
    Evento = "Alterar banco"
    If lista_banco.ListItems.Count <> 0 And CodigoLista5 <> 0 Then
        lista_banco.SelectedItem = lista_banco.ListItems(CodigoLista5)
        lista_banco.SetFocus
    End If
End If
'==================================
Modulo = "Vendas/Clientes"
ID_documento = txtID_banco
Documento = "Cliente: " & txtnomerazao
Documento1 = "Banco: " & txtBanco
ProcGravaEvento
'==================================
Novo_Cliente5 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procSalvar_contato()
On Error GoTo tratar_erro

If Frame2.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
If txtNomeContato.Text = "" Then
    USMsgBox ("Informe o nome do contato antes de salvar."), vbExclamation, "CAPRIND v5.0"
    txtNomeContato.SetFocus
    Exit Sub
End If
Set TBClientes = CreateObject("adodb.recordset")
TBClientes.Open "Select * from clientes_contatos where idcontato = " & txtIDContato.Text, Conexao, adOpenKeyset, adLockOptimistic
If TBClientes.EOF = True Then
    TBClientes.AddNew
Else
    If txtNomeContato <> TBClientes!NomeContato Then
        Conexao.Execute "Update Vendas_tele Set Contato = '" & txtNomeContato & "' where IDcliente = " & txtIDcliente & " and Contato = '" & TBClientes!NomeContato & "'"
        Conexao.Execute "Update Vendas_tele Set contato_visita = '" & txtNomeContato & "' where IDcliente = " & txtIDcliente & " and contato_visita = '" & TBClientes!NomeContato & "'"
    End If
    If FunVerificaRegistroValidado("Clientes", "IDcliente = " & txtIDcliente, "cliente", "este contato", "alterar", True, True) = False Then Exit Sub
End If
TBClientes!IDCliente = txtIDcliente
If txtData1 = "" Then TBClientes!Data = Date Else TBClientes!Data = txtData1
If txtResponsavel1 = "" Then TBClientes!Responsavel = pubUsuario Else TBClientes!Responsavel = txtResponsavel1
TBClientes!NomeContato = txtNomeContato
TBClientes!Departamento = txtdepartamento
TBClientes!telefone = txttelcontato
TBClientes!Email = IIf(TxtEmail_Contato.Text = "", Null, LCase(Trim(TxtEmail_Contato)))
If Chk_enviar_NFe.Value = 1 Then TBClientes!Enviar_NFe = True Else TBClientes!Enviar_NFe = False
If Chk_enviar_boleto.Value = 1 Then TBClientes!Enviar_boleto = True Else TBClientes!Enviar_boleto = False
TBClientes.Update
txtIDContato = TBClientes!idcontato
TBClientes.Close
ProcCarregaListaContatos
If Novo_Cliente1 = True Then
    USMsgBox ("Novo contato cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo contato"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar contato"
    If Lista_contato.ListItems.Count <> 0 And CodigoLista1 <> 0 Then
        Lista_contato.SelectedItem = Lista_contato.ListItems(CodigoLista1)
        Lista_contato.SetFocus
    End If
End If
'==================================
Modulo = "Vendas/Clientes"
ID_documento = txtIDContato
Documento = "Cliente: " & txtnomerazao
Documento1 = "Contato: " & txtNomeContato
ProcGravaEvento
'==================================
Novo_Cliente1 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Proccarregalistaentrega()
On Error GoTo tratar_erro

ListaEntrega.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from clientes_entrega where idcliente = " & IIf(txtIDcliente = "", 0, txtIDcliente) & " and Tipo = 'C' order by endereco_entrega", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    TBLISTA.MoveLast
    'PBLista.Min = 0
    'PBLista.Max = TBLISTA.RecordCount
    'PBLista.Value = 1
    Contador = 0
    TBLISTA.MoveFirst
    Do While TBLISTA.EOF = False
        With ListaEntrega.ListItems
            .Add = TBLISTA!identrega
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!endereco_entrega), "", TBLISTA!endereco_entrega)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!bairro_entrega), "", TBLISTA!bairro_entrega)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!cidade_entrega), "", TBLISTA!cidade_entrega)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!uf_entrega), "", TBLISTA!uf_entrega)
            TBLISTA.MoveNext
            Contador = Contador + 1
            'PBLista.Value = Contador
        End With
    Loop
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregalistacobranca()
On Error GoTo tratar_erro

listacobranca.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from clientes_cobranca where idcliente = " & IIf(txtIDcliente = "", 0, txtIDcliente) & " and Tipo = 'C' order by endereco_cobranca", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    TBLISTA.MoveLast
    'PBLista.Min = 0
    'PBLista.Max = TBLISTA.RecordCount
    'PBLista.Value = 1
    Contador = 0
    TBLISTA.MoveFirst
    Do While TBLISTA.EOF = False
        With listacobranca.ListItems
            .Add = TBLISTA!idCobranca
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!endereco_Cobranca), "", TBLISTA!endereco_Cobranca)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!bairro_Cobranca), "", TBLISTA!bairro_Cobranca)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!cidade_Cobranca), "", TBLISTA!cidade_Cobranca)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!uf_Cobranca), "", TBLISTA!uf_Cobranca)
            TBLISTA.MoveNext
            Contador = Contador + 1
            'PBLista.Value = Contador
        End With
    Loop
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procSalvar_comerciais()
On Error GoTo tratar_erro
  
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select * FROM Clientes_DadosComerciais WHERE IDcliente = " & txtIDcliente.Text & " and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex), Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    If FunVerificaRegistroValidado("Clientes", "IDcliente = " & txtIDcliente, "cliente", "estes dados comerciais", "alterar", True, True) = False Then Exit Sub
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar dados comerciais"
Else
    TBProduto.AddNew
    USMsgBox ("Dados comerciais do cliente cadastrados com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo dados comerciais"
    TBProduto!IDCliente = txtIDcliente
End If
ProcEnviadadosComercial
TBProduto.Update
ID_documento = TBProduto!ID
TBProduto.Close
'==================================
Modulo = "Vendas/Clientes"
Documento = "Cliente: " & txtnomerazao & " - Empresa: " & Cmb_empresa
Documento1 = ""
ProcGravaEvento
'==================================

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procSalvar_entrega()
On Error GoTo tratar_erro

If Frame6.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
If Left(cmbPessoa, 8) = "Jurídica" Then
    If txtCNPJ_entrega.Text <> "__.___.___/____-__" Then
        If Funconsistir_CgcCpf(txtCNPJ_entrega) = False Then
            USMsgBox ("O número do CNPJ digitado não é válido, digite o número correto."), vbExclamation, "CAPRIND v5.0"
            txtCNPJ_entrega.SetFocus
            Exit Sub
        End If
    End If
End If
Acao = "salvar"
If txtEndereco_entrega <> "" And cmbRegimeTributario.Text = "Simples nacional" And cmbTipo_endereco_entrega = "" Then
    NomeCampo = "o tipo"
    ProcVerificaAcao
    cmbTipo_endereco_entrega.SetFocus
    Exit Sub
End If
If txtEndereco_entrega.Text = "" Then
    NomeCampo = "o endereço"
    ProcVerificaAcao
    txtEndereco_entrega.SetFocus
    Exit Sub
End If
If txtBairro_entrega <> "" And cmbRegimeTributario.Text = "Simples nacional" And cmbTipo_bairro_entrega = "" Then
    NomeCampo = "o tipo"
    ProcVerificaAcao
    cmbTipo_bairro_entrega.SetFocus
    Exit Sub
End If
If cmbCidade_Entrega <> "" And cmbuf_entrega <> "" And cmbuf_entrega <> "EX" Then
    'If cmbRegimeTributario.Text = "Simples nacional" And FunVerificaCidade(cmbCidade_Entrega, cmbuf_entrega) = False Then Exit Sub
End If
Set TBClientes = CreateObject("adodb.recordset")
TBClientes.Open "Select * from clientes_entrega where identrega = " & txtID_entrega, Conexao, adOpenKeyset, adLockOptimistic
If TBClientes.EOF = True Then
    TBClientes.AddNew
Else
    If FunVerificaRegistroValidado("Clientes", "IDcliente = " & txtIDcliente, "cliente", "este endereço de entrega", "alterar", True, True) = False Then Exit Sub
End If
TBClientes!IDCliente = txtIDcliente
If txtData2 = "" Then TBClientes!Data = Date Else TBClientes!Data = txtData2
If txtResponsavel2 = "" Then TBClientes!Responsavel = pubUsuario Else TBClientes!Responsavel = txtResponsavel2
TBClientes!Tipo = "C"
TBClientes!CNPJ = txtCNPJ_entrega.Text
TBClientes!Tipo_endereco = cmbTipo_endereco_entrega
TBClientes!endereco_entrega = txtEndereco_entrega
TBClientes!Numero = txtNumero_entrega.Text
TBClientes!complemento = txtComplemento_entrega
TBClientes!Tipo_bairro = cmbTipo_bairro_entrega
TBClientes!bairro_entrega = txtBairro_entrega
If cmbuf_cobranca <> "EX" Then TBClientes!cidade_entrega = IIf(cmbCidade_Entrega.Text = "", Null, cmbCidade_Entrega.Text) Else TBClientes!cidade_Cobranca = IIf(txtCidade_Entrega.Text = "", Null, txtCidade_Entrega.Text)
TBClientes!uf_entrega = cmbuf_entrega
TBClientes!cxpostal_entrega = txtcxpostal_entrega
TBClientes!cep_entrega = mskcep_entrega
TBClientes!tel1_entrega = txttel1_entrega
TBClientes!tel2_entrega = txttel2_entrega
TBClientes!tel3_entrega = txttel3_entrega
TBClientes!tel4_entrega = txttel4_entrega
TBClientes!fax_entrega = txtfax_entrega
TBClientes!email_entrega = IIf(txtemail_entrega.Text = "", Null, LCase(txtemail_entrega.Text))
TBClientes!Site_Entrega = IIf(txtsite_entrega.Text = "", Null, LCase(txtsite_entrega.Text))
TBClientes.Update
txtID_entrega = TBClientes!identrega
TBClientes.Close
Proccarregalistaentrega
If Novo_Cliente2 = True Then
    USMsgBox ("Novo local para entrega cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo local para entrega"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar local para entrega"
    If ListaEntrega.ListItems.Count <> 0 And CodigoLista2 <> 0 Then
        ListaEntrega.SelectedItem = ListaEntrega.ListItems(CodigoLista2)
        ListaEntrega.SetFocus
    End If
End If
'==================================
Modulo = "Vendas/Clientes"
ID_documento = txtID_entrega
Documento = "Cliente: " & txtnomerazao
Documento1 = "Endereço: " & txtEndereco_entrega
ProcGravaEvento
'==================================
Novo_Cliente2 = False
 
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procSalvar_familia()
On Error GoTo tratar_erro

If Frame12.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
If cmbfamilia = "" Then
    USMsgBox ("Informe a família antes de salvar."), vbExclamation, "CAPRIND v5.0"
    cmbfamilia.Enabled = True
    cmbfamilia.SetFocus
    Exit Sub
End If
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from compras_fornecedores_familia where idfamilia = " & txtid_familia, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then
    TBGravar.AddNew
Else
    If FunVerificaRegistroValidado("Clientes", "IDcliente = " & txtIDcliente, "cliente", "esta família", "alterar", True, True) = False Then Exit Sub
End If
TBGravar!IDCliente = txtIDcliente
If txtData4 = "" Then TBGravar!Data = Date Else TBGravar!Data = txtData4
If txtResponsavel4 = "" Then TBGravar!Responsavel = pubUsuario Else TBGravar!Responsavel = txtResponsavel4
TBGravar!Tipo = "C"
TBGravar!Familia = cmbfamilia
TBGravar.Update
txtid_familia = TBGravar!idFamilia
TBGravar.Close
ProcCarregaListaFamilia
If Novo_Cliente4 = True Then
    USMsgBox ("Nova família cadastrada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Nova familia"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar familia"
    If lista_familia.ListItems.Count <> 0 And CodigoLista4 <> 0 Then
        lista_familia.SelectedItem = lista_familia.ListItems(CodigoLista4)
        lista_familia.SetFocus
    End If
End If
'==================================
Modulo = "Vendas/Clientes"
ID_documento = txtid_familia
Documento = "Cliente: " & txtnomerazao
Documento1 = "Família: " & cmbfamilia
ProcGravaEvento
'==================================
Novo_Cliente4 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procSalvar_impostos()
On Error GoTo tratar_erro

If Frame14.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
Acao = "salvar"
If Txt_CF = "" Then
    NomeCampo = "a classificação fiscal"
    ProcVerificaAcao
    cmdCF.SetFocus
    Exit Sub
End If
If txtPorcentagemIPI = "" Then
    NomeCampo = "o valor para cáculo do IPI"
    ProcVerificaAcao
    txtPorcentagemIPI.SetFocus
    Exit Sub
End If
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from Clientes_Impostos where id = " & txtid_impostos, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then
    TBGravar.AddNew
Else
    If FunVerificaRegistroValidado("Clientes", "IDcliente = " & txtIDcliente, "cliente", "este imposto", "alterar", True, True) = False Then Exit Sub
End If
TBGravar!IDCliente = txtIDcliente
If txtData6 = "" Then TBGravar!Data = Date Else TBGravar!Data = txtData6
If txtResponsavel6 = "" Then TBGravar!Responsavel = pubUsuario Else TBGravar!Responsavel = txtResponsavel6
TBGravar!ID_CF = Txt_ID_CF
TBGravar!PorcentagemIPI = txtPorcentagemIPI
TBGravar.Update
txtid_impostos = TBGravar!ID
TBGravar.Close
ProcCarregaLista_Impostos
If Novo_Cliente6 = True Then
    USMsgBox "Novo imposto cadastrado com sucesso.", vbInformation, "CAPRIND v5.0"
    Evento = "Novo imposto"
Else
    USMsgBox "Alteração efetuada com sucesso.", vbInformation, "CAPRIND v5.0"
    Evento = "Alterar imposto"
    If Lista_Impostos.ListItems.Count <> 0 And CodigoLista6 <> 0 Then
        Lista_Impostos.SelectedItem = Lista_Impostos.ListItems(CodigoLista6)
        Lista_Impostos.SetFocus
    End If
End If
'==================================
Modulo = "Vendas/Clientes"
ID_documento = txtid_impostos
Documento = "Cliente: " & txtnomerazao
Documento1 = "NCM: " & Txt_CF
ProcGravaEvento
'==================================
Novo_Cliente6 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaLista_Impostos()
On Error GoTo tratar_erro

Lista_Impostos.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Clientes_Impostos where IDCLiente = " & IIf(txtIDcliente = "", 0, txtIDcliente), Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    TBLISTA.MoveLast
    'PBLista.Min = 0
    'PBLista.Max = TBLISTA.RecordCount
    'PBLista.Value = 1
    Contador = 0
    TBLISTA.MoveFirst
    Do While TBLISTA.EOF = False
        With Lista_Impostos.ListItems
            .Add , , TBLISTA!ID
            
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select * from tbl_ClassificacaoFiscal where Idclass = " & IIf(IsNull(TBLISTA!ID_CF), 0, TBLISTA!ID_CF), Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                .Item(.Count).SubItems(1) = IIf(IsNull(TBAbrir!IDIntClasse), "", TBAbrir!IDIntClasse)
                .Item(.Count).SubItems(2) = IIf(IsNull(TBAbrir!dbl_IPI), "", TBAbrir!dbl_IPI)
            End If
            TBAbrir.Close
            
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!PorcentagemIPI), "", Format(TBLISTA!PorcentagemIPI, "###,##0.00"))
        End With
        TBLISTA.MoveNext
        Contador = Contador + 1
        'PBLista.Value = Contador
    Loop
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procSalvar_outros()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If FunVerificaRegistroValidado("Clientes", "IDcliente = " & txtIDcliente, "cliente", "o mesmo", "alterar", True, True) = False Then Exit Sub

If chkSuframa.Value = 1 And txtSuframa = "" Then
    NomeCampo = "o código suframa"
    ProcVerificaAcao
    txtSuframa.SetFocus
    Exit Sub
End If
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from clientes where IDCLIENTE = " & IIf(txtIDcliente = "", 0, txtIDcliente), Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = False Then
    TBGravar!txt_observacoes = txtObservacoes.Text
    TBGravar!tel02 = txttel02.Text
    TBGravar!tel03 = txttel03.Text
    TBGravar!tel04 = txttel04.Text
    TBGravar!idgrupo = txtIDGrupo
    If txtIDGrupo <> "0" Then ProcAtualizaGrupo
    If chkSuframa.Value = 1 Then
        TBGravar!chkSuframa = True
        TBGravar!Suframa = txtSuframa
    Else
        TBGravar!chkSuframa = False
        TBGravar!Suframa = ""
    End If
    If chkICMSST.Value = 1 Then TBGravar!SimplesICMSST = True Else TBGravar!SimplesICMSST = False
    TBGravar!ISSQN = IIf(Txt_ISSQN = "", Null, Txt_ISSQN)
    TBGravar!Banco = IIf(cmbBanco = "", Null, cmbBanco)
    TBGravar!Tipo_doc = IIf(cmbTipo_doc = "", Null, cmbTipo_doc)
    TBGravar!txtLimiteCredito = IIf(txtLimiteCredito = "", Null, txtLimiteCredito)
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    TBGravar.Update
End If
TBGravar.Close
'==================================
Modulo = "Vendas/Clientes"
Evento = "Alterar"
ID_documento = txtIDcliente
Documento = "Cliente: " & txtnomerazao
Documento1 = ""
ProcGravaEvento
'==================================

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdTransporte_padrao_Click()
On Error GoTo tratar_erro

Aplic = 6
Compras_Cotacao = False
Compras_Pedido = False
Vendas_Proposta = False
Vendas_PI = False
Estoque_recebimento = False
Clientes = True
Sit_REG = 0
frmCompras_pedido_DadosComerciais.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdValidade_Padrao_Click()
On Error GoTo tratar_erro

Aplic = 10
Compras_Cotacao = False
Compras_Pedido = False
Vendas_Proposta = False
Vendas_PI = False
Estoque_recebimento = False
Clientes = True
Sit_REG = 0
frmCompras_pedido_DadosComerciais.Show

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
            Case vbKeyInsert: ProcNovo
            Case vbKeyF2: ProcLocalizar
            Case vbKeyF3: ProcSalvar
            Case vbKeyF4: ProcExcluir
            Case vbKeyF5: ProcImprimir
            Case vbKeyF7: If Cmb_opcao_lista = "Status" Then ProcStatus
            Case vbKeyF10: If Cmb_opcao_lista = "Validação" Then ProcValidarRegistros Lista, "Vendas/Clientes"
            Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
    Case 1:
        Select Case KeyCode
            Case vbKeyInsert: procNovo_Contato
            Case vbKeyF3: procSalvar_contato
            Case vbKeyF4: procExcluir_contato
            Case vbKeyF5: ProcImprimir
            Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
    Case 2:
        Select Case KeyCode
            Case vbKeyInsert: procNovo_entrega
            Case vbKeyF3: procSalvar_entrega
            Case vbKeyF4: procExcluir_entrega
            Case vbKeyF5: ProcImprimir
            Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
    Case 3:
        Select Case KeyCode
            Case vbKeyInsert: procNovo_cobranca
            Case vbKeyF3: procSalvar_cobranca
            Case vbKeyF4: procExcluir_cobranca
            Case vbKeyF5: ProcImprimir
            Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
    Case 4:
        Select Case KeyCode
            Case vbKeyInsert: procNovo_familia
            Case vbKeyF3: procSalvar_familia
            Case vbKeyF4: procExcluir_familia
            Case vbKeyF5: ProcImprimir
            Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
    Case 5:
        Select Case KeyCode
            Case vbKeyInsert: procNovo_Banco
            Case vbKeyF3: procSalvar_Banco
            Case vbKeyF4: procExcluir_banco
            Case vbKeyF5: ProcImprimir
            Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
    Case 6:
        Select Case KeyCode
            Case vbKeyInsert: procNovo_impostos
            Case vbKeyF3: procSalvar_impostos
            Case vbKeyF4: procExcluir_impostos
            Case vbKeyF5: ProcImprimir
            Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
    Case 7:
        Select Case KeyCode
            Case vbKeyF3: procSalvar_comerciais
            Case vbKeyF4: procExcluir_comerciais
            Case vbKeyF5: ProcImprimir
            Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
    Case 8:
        Select Case KeyCode
            Case vbKeyF3: procSalvar_outros
            Case vbKeyF5: ProcImprimir
            Case vbKeyF1: ProcAjuda
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

ProcCarregaToolBar1 Me, 15195, 15, True
ProcCarregaToolBar2 Me, 15195, 11, True
ProcCarregaToolBar3 Me, 15195, 9, True
ProcCarregaToolBar4 Me, 15195, 8, True
Formulario = "Vendas/Clientes"
Direitos
ProcCarregaCombos
USToolBar2.Visible = False
SSTab1.Tab = 0
ProcLimpaVariaveisPrincipais
Cmb_opcao_lista.Text = "Validação"
ProcCarregaComboVendedor
ProcRemoveObjetosResize Me

If IDempresa <> 0 Then
    Set TBFIltro = CreateObject("adodb.recordset")
    TBFIltro.Open "Select * from Empresa where Codigo = " & IDempresa, Conexao, adOpenKeyset, adLockOptimistic
    If TBFIltro.EOF = False Then
    CNPJ_Empresa = TBFIltro!CNPJ
    End If
    TBFIltro.Close
End If


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimparTudo()
On Error GoTo tratar_erro

Frame2.Enabled = False
Frame6.Enabled = False
Frame8.Enabled = False
Frame12.Enabled = False
Frame10.Enabled = False
Frame14.Enabled = False
procLimpacamposContatos
ProcLimpacamposEntrega
ProcLimpacamposCobranca
ProcLimpaFamilia
ProcLimpacampos_banco
ProcLimpaCampos_Impostos
ProcLimpaCamposComerciais
ProcLimpaCampos_Outros
Novo_Cliente1 = False
Novo_Cliente2 = False
Novo_Cliente3 = False
Novo_Cliente4 = False
Novo_Cliente5 = False
Novo_Cliente6 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpar()
On Error GoTo tratar_erro
  
txtData = Format(Date, "dd/mm/yy")
txtResponsavel = pubUsuario
txtStatus = "Liberado"
txtDtValidacao = ""
txtRespValidacao = ""
Chk_prospecto.Value = 0
Chk_enviar_NF.Value = 1
With cmbPessoa
    .ListIndex = -1
'    .Locked = False
'    .TabStop = True
End With
txtnomerazao.Text = ""
txtCategoria.ListIndex = -1
cmbOrigem.ListIndex = -1
txtcnpj.Text = "__.___.___/____-__"
txtCpf = "___.___.___-__"
txtRG_IE.Text = ""
txtIM_IE = ""
txtnomefantasia.Text = ""
Txt_pais.ListIndex = -1
cmbTipo_endereco.ListIndex = -1
txtendereco.Text = ""
txtNumero = ""
cmbTipo_bairro.ListIndex = -1
txtBairro.Text = ""
cmbuf.ListIndex = -1
cmbCidade.ListIndex = -1
txtCidade.Text = ""
txtComplemento = ""
txtcaixapostal.Text = ""
txtCEP.Text = ""
txttel01.Text = ""
txtFax.Text = ""
txtEmail.Text = ""
txtSite.Text = ""
cmbRegimeTributario.ListIndex = -1
Cmb_tipo_transp.ListIndex = -1
cmbtransportadora.ListIndex = -1
Chk_nao_contribuinte_ICMS.Value = 0
CodigoLista = 0
Caption = "Administrativo - Vendas - Clientes"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcPuxaDados()
On Error GoTo tratar_erro

txtData.Text = IIf(IsNull(TBClientes!Data), "", Format(TBClientes!Data, "dd/mm/yy"))
txtResponsavel = IIf(IsNull(TBClientes!Responsavel), "", TBClientes!Responsavel)
txtDtValidacao.Text = IIf(IsNull(TBClientes!DtValidacao), "", TBClientes!DtValidacao)
txtRespValidacao.Text = IIf(IsNull(TBClientes!RespValidacao), "", TBClientes!RespValidacao)
txtnomerazao.Text = IIf(IsNull(TBClientes!NomeRazao), "", TBClientes!NomeRazao)
Caption = "Administrativo - Clientes (Cliente : " & IIf(IsNull(TBClientes!NomeRazao), "", TBClientes!NomeRazao) & ")"
With cmbPessoa
    If TBClientes!Tipo <> "" Then
        Select Case TBClientes!Tipo
            Case "FP": .Text = "Física"
            Case "JP": .Text = "Jurídica"
            Case "FR": .Text = "Física"
            Case "JR": .Text = "Jurídica"
        End Select
'        .Locked = True
'        .TabStop = False
    Else
        .ListIndex = -1
'        .Locked = False
'        .TabStop = True
    End If
End With
If TBClientes!idTipoEmpresa = 1 Then
    If TBClientes!Tipo = "JP" Or TBClientes!Tipo = "JR" Then
        If TBClientes!CPF_CNPJ <> "" Then
            CNPJ = Trim(TBClientes!CPF_CNPJ)
            With txtcnpj
                .PromptInclude = False
                .Text = CNPJ
                .PromptInclude = True
            End With
        End If
    End If
    If TBClientes!Tipo = "FP" Or TBClientes!Tipo = "FR" Then
        If TBClientes!CPF_CNPJ <> "" Then
            CNPJ = Trim(TBClientes!CPF_CNPJ)
            With txtCpf
                .PromptInclude = False
                .Text = CNPJ
                .PromptInclude = True
            End With
        End If
    End If
End If
If TBClientes!idTipoEmpresa <> "" Then
    If TBClientes!idTipoEmpresa = 1 Then cmbOrigem.Text = "Nacional"
    If TBClientes!idTipoEmpresa = 0 Then cmbOrigem.Text = "Estrangeiro"
End If
If TBClientes!Presumido = True Then cmbRegimeTributario.Text = "Lucro presumido"
If TBClientes!Simples = True Then cmbRegimeTributario.Text = "Simples nacional"
If TBClientes!Real = True Then cmbRegimeTributario.Text = "Lucro real"
If TBClientes!MEI = True Then cmbRegimeTributario.Text = "MEI"
txtendereco.Text = IIf(IsNull(TBClientes!Endereco), "", TBClientes!Endereco)
txtBairro.Text = IIf(IsNull(TBClientes!Bairro), "", TBClientes!Bairro)
txtRG_IE.Text = IIf(IsNull(TBClientes!RG_IE), "", TBClientes!RG_IE)
txtEmail.Text = IIf(IsNull(TBClientes!Email), "", TBClientes!Email)
txtSite.Text = IIf(IsNull(TBClientes!Site), "", TBClientes!Site)
txtcaixapostal.Text = IIf(IsNull(TBClientes!cxpostal), "", TBClientes!cxpostal)
txttel01.Text = IIf(IsNull(TBClientes!Tel01), "", TBClientes!Tel01)
txtFax.Text = IIf(IsNull(TBClientes!Fax), "", TBClientes!Fax)
txtnomefantasia.Text = IIf(IsNull(TBClientes!NomeFantasia), "", TBClientes!NomeFantasia)
txtIM_IE = IIf(IsNull(TBClientes!RG_IM), "", TBClientes!RG_IM)
txtComplemento.Text = IIf(IsNull(TBClientes!complemento), "", TBClientes!complemento)
txtNumero = IIf(IsNull(TBClientes!Numero), "", TBClientes!Numero)
If TBClientes!Prospecto = True Then Chk_prospecto.Value = 1 Else Chk_prospecto.Value = 0
If TBClientes!Enviar_NF = True Then Chk_enviar_NF.Value = 1 Else Chk_enviar_NF.Value = 0
txtCEP = IIf(IsNull(TBClientes!CEP), "", TBClientes!CEP)
If TBClientes!status = "Liberado" Or TBClientes!status = "" Or IsNull(TBClientes!status) = True Then
    txtStatus.Text = "Liberado"
ElseIf TBClientes!status = "Bloqueado" Then
        txtStatus.Text = "Bloqueado"
    Else
        txtStatus.Text = "Parcial"
End If

NomeCampo = "a categoria"
If IsNull(TBClientes!Categoria) = False And TBClientes!Categoria <> "" Then txtCategoria = TBClientes!Categoria
NomeCampo = "o país"
If IsNull(TBClientes!Pais) = False And TBClientes!Pais <> "" Then Txt_pais = TBClientes!Pais
NomeCampo = "o tipo do endereço"
If IsNull(TBClientes!Tipo_endereco) = False And TBClientes!Tipo_endereco <> "" Then cmbTipo_endereco.Text = TBClientes!Tipo_endereco
NomeCampo = "o tipo do bairro"
If IsNull(TBClientes!Tipo_bairro) = False And TBClientes!Tipo_bairro <> "" Then cmbTipo_bairro.Text = TBClientes!Tipo_bairro
NomeCampo = "o estado"
If IsNull(TBClientes!UF) = False And TBClientes!UF <> "" Then cmbuf.Text = TBClientes!UF
NomeCampo = "a Cidade"
If TBClientes!UF <> "EX" Then
    cmbCidade.Visible = True
    txtCidade.Visible = False
    'Debug.print TBClientes!Cidade
    Cidade = TBClientes!Cidade
    
    If Cidade = "santa barbara d'oeste" Then
    cmbCidade = "SANTA BARBARA DO OESTE"
    Else
    If IsNull(TBClientes!Cidade) = False And TBClientes!Cidade <> "" Then cmbCidade = TBClientes!Cidade
    End If
    
Else
    cmbCidade.Visible = False
    txtCidade.Visible = True
    
    If TBClientes!Cidade = "SANTA BARBARA D'OESTE" Then
    txtCidade.Text = "SANTA BARBARA DO OESTE"
    Else
    txtCidade.Text = IIf(IsNull(TBClientes!Cidade), "", TBClientes!Cidade)
    End If
    
End If
NomeCampo = "o tipo da transportadora"
If IsNull(TBClientes!Tipo_transp) = False And TBClientes!Tipo_transp <> "" Then
    Select Case TBClientes!Tipo_transp
        Case "C": Cmb_tipo_transp = "Cliente"
        Case "F": Cmb_tipo_transp = "Fornecedor"
        Case "E": Cmb_tipo_transp = "Empresa"
    End Select
End If
NomeCampo = "a transportadora"
If IsNull(TBClientes!txt_transportadora) = False And TBClientes!txt_transportadora <> "" Then cmbtransportadora = TBClientes!txt_transportadora
If TBClientes!Nao_contribuinte_ICMS = True Then Chk_nao_contribuinte_ICMS.Value = 1 Else Chk_nao_contribuinte_ICMS.Value = 0
2:
    
    Novo_Cliente = False
    With txtIDcliente
        .Locked = False
        .TabStop = True
    End With
    ProcLimparTudo


'================================================================
' Busca vendedor interno
'================================================================

StrSql = "Select CL.IDCliente,VV.Id as IDVendedor, CL.NomeRazao , VV.Vendedor, VVC.Comissao from Clientes as CL inner join Vendas_Vendedores_Clientes as VVC on Cl.IDCliente = VVC.IDCliente Inner Join Vendas_Vendedores as VV On VVC.IDVendedor = VV.Id Where CL.IDCliente = " & txtIDcliente.Text & " order by CL.NomeRazao"

'Debug.print StrSql

    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
    cmbVendedor.Visible = True
    txtVendedor.Visible = False
    txtIDVendedor.Text = TBAbrir!IDvendedor
    txtVendedor.Text = TBAbrir!vendedor
    cmbVendedor.Text = TBAbrir!vendedor
    txtComissao.Text = IIf(IsNull(TBAbrir!Comissao), 0, TBAbrir!Comissao)

    Else
    
    ' If USMsgBox("Vendedor externo não cadastrado para esse cliente." & vbCrLf & "Deseja buscar do pedido interno?", vbYesNo, "CAPRIND v5.0") = vbYes Then
        StrSql = "select VP.IDcliente,VP.Cidade, cliente, Vend_ext, VE, VV.Comissao from vendas_proposta VP Inner join Vendas_Vendedores VV on VP.VE = VV.Id Where VP.idcliente = '" & txtIDcliente.Text & "' group by VP.idcliente, VP.cidade, cliente, vend_ext,VE, VV.Comissao order by Vend_ext "
        'Debug.print StrSql
        
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
        cmbVendedor.Visible = True
        txtVendedor.Visible = False
        txtIDVendedor.Text = TBAbrir!VE
        txtVendedor.Text = TBAbrir!Vend_ext
        cmbVendedor.Text = TBAbrir!Vend_ext
        
        txtComissao.Text = IIf(IsNull(TBAbrir!Comissao), 0, TBAbrir!Comissao)
        'End If
        
        Else
    
        txtIDVendedor.Text = ""
        txtVendedor.Text = ""
        txtComissao.Text = ""
        txtVendedor.Visible = False
        cmbVendedor.Visible = True
        cmbVendedor.ListIndex = -1
        End If
        TBAbrir.Close
   End If
'================================================================
'procAtualizaClienteNuvem

Exit Sub
tratar_erro:
    If Err.Number = "383" Then
        USMsgBox ("Não foi encontrado " & NomeCampo & " desse cliente, favor revisar."), vbExclamation, "CAPRIND v5.0"
        GoTo 2
    End If
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

Formulario = "Vendas/Clientes"
Direitos
ProcCarregaCombos
ProcLimpaVariaveisPrincipais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procAtualiza()
On Error GoTo tratar_erro

If InputBox("Informe a senha para liberar.") = "280362C" Then
    If USMsgBox("Deseja realmente atualizar os dados do endereço de entrega, cobrança e o código do país?", vbYesNo, "CAPRIND v5.0") = vbYes Then
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from clientes_entrega order by endereco_entrega", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            TBAbrir.MoveLast
            'PBLista.Min = 0
            'PBLista.Max = TBAbrir.RecordCount
            'PBLista.Value = 1
            Contador = 0
            TBAbrir.MoveFirst
            Do While TBAbrir.EOF = False
                If IsNull(TBAbrir!Tipo) = True Or TBAbrir!Tipo <> "" Then
                    TBAbrir!Tipo = "C"
                    TBAbrir.Update
                End If
                TBAbrir.MoveNext
                Contador = Contador + 1
                'PBLista.Value = Contador
            Loop
        End If
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from clientes_cobranca order by endereco_cobranca", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            TBAbrir.MoveLast
            'PBLista.Min = 0
            'PBLista.Max = TBAbrir.RecordCount
            'PBLista.Value = 1
            Contador = 0
            TBAbrir.MoveFirst
            Do While TBAbrir.EOF = False
                If IsNull(TBAbrir!Tipo) = True Or TBAbrir!Tipo <> "" Then
                    TBAbrir!Tipo = "C"
                    TBAbrir.Update
                End If
                TBAbrir.MoveNext
                Contador = Contador + 1
                'PBLista.Value = Contador
            Loop
        End If
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from Clientes order by Pais", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            TBAbrir.MoveLast
            'PBLista.Min = 0
            'PBLista.Max = TBAbrir.RecordCount
            'PBLista.Value = 1
            Contador = 0
            TBAbrir.MoveFirst
            Do While TBAbrir.EOF = False
                If IsNull(TBAbrir!Pais) = True Or TBAbrir!Pais <> "" Then
                    Set TBFI = CreateObject("adodb.recordset")
                    TBFI.Open "Select * from Codigos_pais where Pais = '" & TBAbrir!Pais & "'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBFI.EOF = False Then
                        TBAbrir!Codigo_pais = TBFI!CODIGO
                    Else
                        TBAbrir!Pais = ""
                        TBAbrir!Codigo_pais = Null
                    End If
                    TBAbrir.Update
                    TBFI.Close
                End If
                TBAbrir.MoveNext
                Contador = Contador + 1
                'PBLista.Value = Contador
            Loop
        End If
        TBAbrir.Close
        USMsgBox ("Atualização efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
        '==================================
        Modulo = "Vendas/Clientes"
        Evento = "Atualizar"
        ID_documento = 0
        Documento = ""
        Documento1 = ""
        ProcGravaEvento
        '==================================
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub lista_banco_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With lista_banco
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If FunVerificaRegistroValidadoSemMsg("Clientes", "IDcliente = " & txtIDcliente, True) = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
Proximo:
                .ListItems.Item(InitFor).Checked = True
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView lista_banco, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub lista_banco_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If lista_banco.ListItems.Count = 0 Then Exit Sub
ProcLimpacampos_banco
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Compras_fornecedores_banco where id = " & lista_banco.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    txtID_banco = TBAbrir!ID
    txtData5 = IIf(IsNull(TBAbrir!Data), "", Format(TBAbrir!Data, "dd/mm/yy"))
    txtResponsavel5 = IIf(IsNull(TBAbrir!Responsavel), "", TBAbrir!Responsavel)
    txtBanco = IIf(IsNull(TBAbrir!Banco), "", TBAbrir!Banco)
    txtAgencia = IIf(IsNull(TBAbrir!Agencia), "", TBAbrir!Agencia)
    txtConta = IIf(IsNull(TBAbrir!Conta), "", TBAbrir!Conta)
    CodigoLista5 = lista_banco.SelectedItem.index
End If
TBAbrir.Close
Frame10.Enabled = True
Novo_Cliente5 = False
    
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
                    If .ListItems(InitFor).ListSubItems(4) = "Sim" Then
                        .ListItems.Item(InitFor).Checked = False
                        GoTo Proximo
                    End If
                    
                    IDCliente = .ListItems(InitFor)
                    ProcVerificaRegistroUtilizadoSemMsg "Certificado_qualidade", "IDcliente = " & IDCliente
                    If Permitido = False Then
                        .ListItems.Item(InitFor).Checked = False
                        GoTo Proximo
                    End If
                    ProcVerificaRegistroUtilizadoSemMsg "Estoque_Controle", "id_cliente = " & IDCliente
                    If Permitido = False Then
                        .ListItems.Item(InitFor).Checked = False
                        GoTo Proximo
                    End If
                    ProcVerificaRegistroUtilizadoSemMsg "Liquido_penetrante", "IDcliente = " & IDCliente
                    If Permitido = False Then
                        .ListItems.Item(InitFor).Checked = False
                        GoTo Proximo
                    End If
                    ProcVerificaRegistroUtilizadoSemMsg "Producao", "Cliente = '" & .ListItems(InitFor).ListSubItems(3) & "'"
                    If Permitido = False Then
                        .ListItems.Item(InitFor).Checked = False
                        GoTo Proximo
                    End If
                    ProcVerificaRegistroUtilizadoSemMsg "Projproduto_clientes", "IDcliente = " & IDCliente
                    If Permitido = False Then
                        .ListItems.Item(InitFor).Checked = False
                        GoTo Proximo
                    End If
                    ProcVerificaRegistroUtilizadoSemMsg "tbl_contas_receber", "IDcliente = " & IDCliente
                    If Permitido = False Then
                        .ListItems.Item(InitFor).Checked = False
                        GoTo Proximo
                    End If
                    ProcVerificaRegistroUtilizadoSemMsg "tbl_Dados_Nota_Fiscal", "Id_Int_Cliente = " & IDCliente & " and txt_Razao_Nome = '" & .ListItems(InitFor).ListSubItems(3) & "'"
                    If Permitido = False Then
                        .ListItems.Item(InitFor).Checked = False
                        GoTo Proximo
                    End If
                    ProcVerificaRegistroUtilizadoSemMsg "vendas_proposta", "IDcliente = " & IDCliente
                    If Permitido = False Then
                        .ListItems.Item(InitFor).Checked = False
                        GoTo Proximo
                    End If
                    ProcVerificaRegistroUtilizadoSemMsg "Vendas_Tele", "IDcliente = " & IDCliente
                    If Permitido = False Then
                        .ListItems.Item(InitFor).Checked = False
                        GoTo Proximo
                    End If
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

Private Sub Lista_contato_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Lista_contato
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If FunVerificaRegistroValidadoSemMsg("Clientes", "IDcliente = " & txtIDcliente, True) = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView Lista_contato, ColumnHeader
End If


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_contato_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With Lista_contato
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If FunVerificaRegistroValidado("Clientes", "IDcliente = " & txtIDcliente, "cliente", "contato", "excluir este", True, True) = False Then .ListItems.Item(InitFor).Checked = False
        End If
    Next InitFor
End With


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_contato_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista_contato.ListItems.Count = 0 Then Exit Sub
procLimpacamposContatos
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Clientes_contatos where IDContato = " & Lista_contato.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    txtIDContato.Text = TBAbrir!idcontato
    txtData1 = IIf(IsNull(TBAbrir!Data), "", Format(TBAbrir!Data, "dd/mm/yy"))
    txtResponsavel1 = IIf(IsNull(TBAbrir!Responsavel), "", TBAbrir!Responsavel)
    txtNomeContato.Text = IIf(IsNull(TBAbrir!NomeContato), "", TBAbrir!NomeContato)
    txtdepartamento.Text = IIf(IsNull(TBAbrir!Departamento), "", TBAbrir!Departamento)
    txttelcontato.Text = IIf(IsNull(TBAbrir!telefone), "", TBAbrir!telefone)
    TxtEmail_Contato.Text = IIf(IsNull(TBAbrir!Email), "", TBAbrir!Email)
    If TBAbrir!Enviar_NFe = True Then Chk_enviar_NFe.Value = 1 Else Chk_enviar_NFe.Value = 0
    If TBAbrir!Enviar_boleto = True Then Chk_enviar_boleto.Value = 1 Else Chk_enviar_boleto.Value = 0
    CodigoLista1 = Lista_contato.SelectedItem.index
End If
TBAbrir.Close
Frame2.Enabled = True
Novo_Cliente1 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub lista_familia_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With lista_familia
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If FunVerificaRegistroValidadoSemMsg("Clientes", "IDcliente = " & txtIDcliente, True) = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
Proximo:
                .ListItems.Item(InitFor).Checked = True
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView lista_familia, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub lista_familia_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If lista_familia.ListItems.Count = 0 Then Exit Sub
ProcLimpaFamilia
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from compras_fornecedores_familia where idFamilia = " & lista_familia.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    txtid_familia = TBAbrir!idFamilia
    txtData4 = IIf(IsNull(TBAbrir!Data), "", Format(TBAbrir!Data, "dd/mm/yy"))
    txtResponsavel4 = IIf(IsNull(TBAbrir!Responsavel), "", TBAbrir!Responsavel)
    If IsNull(TBAbrir!Familia) = False And TBAbrir!Familia <> "" Then cmbfamilia = TBAbrir!Familia
End If
1:
    TBAbrir.Close
    CodigoLista4 = lista_familia.SelectedItem.index
    Novo_Cliente4 = False
    Frame12.Enabled = True

Exit Sub
tratar_erro:
    If Err.Number = "383" Then
        USMsgBox ("Não foi encontrado a família deste cliente."), vbExclamation, "CAPRIND v5.0"
        GoTo 1
    End If
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_Impostos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Lista_Impostos
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If FunVerificaRegistroValidadoSemMsg("Clientes", "IDcliente = " & txtIDcliente, True) = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
Proximo:
                .ListItems.Item(InitFor).Checked = True
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView lista_familia, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_Impostos_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista_Impostos.ListItems.Count = 0 Then Exit Sub
ProcLimpaCampos_Impostos
Set TBFIltro = CreateObject("adodb.recordset")
TBFIltro.Open "Select * from Clientes_Impostos where id = " & Lista_Impostos.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBFIltro.EOF = False Then
    txtid_impostos = TBFIltro!ID
    txtData6 = IIf(IsNull(TBFIltro!Data), "", Format(TBFIltro!Data, "dd/mm/yy"))
    txtResponsavel6 = IIf(IsNull(TBFIltro!Responsavel), "", TBFIltro!Responsavel)
    
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from tbl_ClassificacaoFiscal where Idclass = " & IIf(IsNull(TBFIltro!ID_CF), 0, TBFIltro!ID_CF), Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Txt_ID_CF = TBAbrir!Idclass
        Txt_CF = IIf(IsNull(TBAbrir!IDIntClasse), "", TBAbrir!IDIntClasse)
    End If
    TBAbrir.Close
    
    txtPorcentagemIPI = IIf(IsNull(TBFIltro!PorcentagemIPI), "", Format(TBFIltro!PorcentagemIPI, "###,##0.00"))
    CodigoLista6 = Lista_Impostos.SelectedItem.index
End If
TBFIltro.Close
Frame14.Enabled = True
Novo_Cliente6 = False
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 Then Exit Sub
Novo_Cliente = False
txtIDcliente = ""
txtIDcliente = Lista.SelectedItem
CodigoLista = Lista.SelectedItem.index
Frame1.Enabled = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub listacobranca_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With listacobranca
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If FunVerificaRegistroValidadoSemMsg("Clientes", "IDcliente = " & txtIDcliente, True) = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                ProcVerificaRegistroUtilizadoSemMsg "vendas_comercial", "ID_cobranca = " & .ListItems(InitFor)
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                ProcVerificaRegistroUtilizadoSemMsg "tbl_Dados_Nota_Fiscal_NFe", "ID_cobranca = " & .ListItems(InitFor)
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView Lista_contato, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub listacobranca_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With listacobranca
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If FunVerificaRegistroValidado("Clientes", "IDcliente = " & txtIDcliente, "cliente", "endereço de cobrança", "excluir este", True, True) = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            Mensagem = "Não é permitido excluir este local de cobrança, pois o mesmo está sendo utilizado no módulo"
            ProcVerificaRegistroUtilizado "vendas_comercial", "ID_cobranca = " & .ListItems(InitFor), "Vendas"
            If Permitido = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            ProcVerificaRegistroUtilizado "tbl_Dados_Nota_Fiscal_NFe", "ID_cobranca = " & .ListItems(InitFor), "Faturamento/Nota fiscal"
            If Permitido = False Then .ListItems.Item(InitFor).Checked = False
        End If
    Next InitFor
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub listacobranca_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If listacobranca.ListItems.Count = 0 Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from clientes_cobranca where idcobranca = " & listacobranca.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    txtID_cobranca.Text = TBLISTA!idCobranca
    txtData3 = IIf(IsNull(TBLISTA!Data), "", Format(TBLISTA!Data, "dd/mm/yy"))
    txtResponsavel3 = IIf(IsNull(TBLISTA!Responsavel), "", TBLISTA!Responsavel)
    If Left(cmbPessoa, 8) = "Jurídica" Then
        If TBLISTA!CNPJ <> "" Then
            txtCNPJ_entrega = IIf(IsNull(TBLISTA!CNPJ), "__.___.___/____-__", IIf(Len(TBLISTA!CNPJ) > 14, TBLISTA!CNPJ, "__.___.___/____-__"))
        End If
    Else
       If TBLISTA!CNPJ <> "" Then
            txtCNPJ_entrega = IIf(IsNull(TBLISTA!CNPJ), "___.___.___-__", IIf(Len(TBLISTA!CNPJ) > 14, "___.___.___-__", TBLISTA!CNPJ))
       End If
    End If
    txtendereco_cobranca.Text = IIf(IsNull(TBLISTA!endereco_Cobranca), "", TBLISTA!endereco_Cobranca)
    txtNumero_cobranca.Text = IIf(IsNull(TBLISTA!Numero), "", TBLISTA!Numero)
    txtComplemento_cobranca.Text = IIf(IsNull(TBLISTA!complemento), "", TBLISTA!complemento)
    txtbairro_cobranca.Text = IIf(IsNull(TBLISTA!bairro_Cobranca), "", TBLISTA!bairro_Cobranca)
    txtcxpostal_cobranca.Text = IIf(IsNull(TBLISTA!cxpostal_Cobranca), "", TBLISTA!cxpostal_Cobranca)
    mskcep_cobranca = IIf(IsNull(TBLISTA!cep_Cobranca), "", TBLISTA!cep_Cobranca)
    txttel1_cobranca.Text = IIf(IsNull(TBLISTA!tel1_Cobranca), "", TBLISTA!tel1_Cobranca)
    txttel2_cobranca.Text = IIf(IsNull(TBLISTA!tel2_Cobranca), "", TBLISTA!tel2_Cobranca)
    txttel3_cobranca.Text = IIf(IsNull(TBLISTA!tel3_Cobranca), "", TBLISTA!tel3_Cobranca)
    txttel4_cobranca.Text = IIf(IsNull(TBLISTA!tel4_Cobranca), "", TBLISTA!tel4_Cobranca)
    txtfax_cobranca.Text = IIf(IsNull(TBLISTA!fax_Cobranca), "", TBLISTA!fax_Cobranca)
    txtemail_cobranca.Text = IIf(IsNull(TBLISTA!email_Cobranca), "", TBLISTA!email_Cobranca)
    txtSite_cobranca.Text = IIf(IsNull(TBLISTA!site_cobranca), "", TBLISTA!site_cobranca)
    NomeCampo = "o tipo do endereço"
    If IsNull(TBLISTA!Tipo_endereco) = False And TBLISTA!Tipo_endereco <> "" Then cmbTipo_endereco_cobranca.Text = TBLISTA!Tipo_endereco
    NomeCampo = "o tipo do bairro"
    If IsNull(TBLISTA!Tipo_bairro) = False And TBLISTA!Tipo_bairro <> "" Then cmbTipo_bairro_cobranca.Text = TBLISTA!Tipo_bairro
    NomeCampo = "o estado"
    If IsNull(TBLISTA!uf_Cobranca) = False And TBLISTA!uf_Cobranca <> "" Then cmbuf_cobranca = TBLISTA!uf_Cobranca
    NomeCampo = "a ciade"
    If TBLISTA!uf_Cobranca <> "EX" Then
        cmbCidade_cobranca.Visible = True
        txtcidade_cobranca.Visible = False
        If UCase(TBLISTA!cidade_Cobranca) = "SANTA BARBARA D'OESTE" Then
        cmbCidade_cobranca.Text = "SANTA BARBARA DO OESTE"
        Else
        cmbCidade_cobranca.Text = IIf(IsNull(TBLISTA!cidade_Cobranca), "", TBLISTA!cidade_Cobranca)
        End If
    Else
        cmbCidade_cobranca.Visible = False
        txtcidade_cobranca.Visible = True
        
        If UCase(TBLISTA!cidade_Cobranca) = "SANTA BARBARA D'OESTE" Then
        cmbCidade_cobranca.Text = "SANTA BARBARA DO OESTE"
        Else
        cmbCidade_cobranca.Text = IIf(IsNull(TBLISTA!cidade_Cobranca), "", TBLISTA!cidade_Cobranca)
        End If
    End If
1:
    Frame8.Enabled = True
    CodigoLista3 = listacobranca.SelectedItem.index
    Novo_Cliente3 = False
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    If Err.Number = "383" Then
        USMsgBox ("Não foi encontrado " & NomeCampo & " do endereço de cobrança desse cliente."), vbExclamation, "CAPRIND v5.0"
        GoTo 1
    End If
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaEntrega_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With ListaEntrega
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If FunVerificaRegistroValidadoSemMsg("Clientes", "IDcliente = " & txtIDcliente, True) = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                ProcVerificaRegistroUtilizadoSemMsg "tbl_Dados_Nota_Fiscal_NFe", "ID_entrega = " & .ListItems(InitFor)
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                ProcVerificaRegistroUtilizadoSemMsg "Vendas_comercial", "ID_entrega = " & .ListItems(InitFor)
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView Lista_contato, ColumnHeader
End If


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaEntrega_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With ListaEntrega
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If FunVerificaRegistroValidado("Clientes", "IDcliente = " & txtIDcliente, "cliente", "endereço de entrega", "excluir este", True, True) = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            Mensagem = "Não é permitido excluir este local para entrega, pois o mesmo está sendo utilizado no módulo"
            ProcVerificaRegistroUtilizado "Vendas_comercial", "ID_entrega = " & .ListItems(InitFor), "Faturamento/Nota fiscal"
            If Permitido = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            ProcVerificaRegistroUtilizado "tbl_Dados_Nota_Fiscal_NFe", "ID_entrega = " & .ListItems(InitFor), "Faturamento/Nota fiscal"
            If Permitido = False Then .ListItems.Item(InitFor).Checked = False
        End If
    Next InitFor
End With

With Lista_contato
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            
            If Permitido = False Then .ListItems.Item(InitFor).Checked = False
        End If
    Next InitFor
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaEntrega_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If ListaEntrega.ListItems.Count = 0 Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from clientes_entrega where identrega = " & ListaEntrega.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    txtID_entrega.Text = TBLISTA!identrega
    txtData2 = IIf(IsNull(TBLISTA!Data), "", Format(TBLISTA!Data, "dd/mm/yy"))
    txtResponsavel2 = IIf(IsNull(TBLISTA!Responsavel), "", TBLISTA!Responsavel)
    If Left(cmbPessoa, 8) = "Jurídica" Then
        If TBLISTA!CNPJ <> "" Then
            txtCNPJ_entrega = IIf(IsNull(TBLISTA!CNPJ), "__.___.___/____-__", IIf(Len(TBLISTA!CNPJ) > 14, TBLISTA!CNPJ, "__.___.___/____-__"))
        End If
    Else
       If TBLISTA!CNPJ <> "" Then
            txtCNPJ_entrega = IIf(IsNull(TBLISTA!CNPJ), "___.___.___-__", IIf(Len(TBLISTA!CNPJ) > 14, "___.___.___-__", TBLISTA!CNPJ))
       End If
    End If
    txtEndereco_entrega.Text = IIf(IsNull(TBLISTA!endereco_entrega), "", TBLISTA!endereco_entrega)
    txtNumero_entrega.Text = IIf(IsNull(TBLISTA!Numero), "", TBLISTA!Numero)
    txtComplemento_entrega.Text = IIf(IsNull(TBLISTA!complemento), "", TBLISTA!complemento)
    txtBairro_entrega.Text = IIf(IsNull(TBLISTA!bairro_entrega), "", TBLISTA!bairro_entrega)
    txtcxpostal_entrega.Text = IIf(IsNull(TBLISTA!cxpostal_entrega), "", TBLISTA!cxpostal_entrega)
    mskcep_entrega = IIf(IsNull(TBLISTA!cep_entrega), "", TBLISTA!cep_entrega)
    txttel1_entrega.Text = IIf(IsNull(TBLISTA!tel1_entrega), "", TBLISTA!tel1_entrega)
    txttel2_entrega.Text = IIf(IsNull(TBLISTA!tel2_entrega), "", TBLISTA!tel2_entrega)
    txttel3_entrega.Text = IIf(IsNull(TBLISTA!tel3_entrega), "", TBLISTA!tel3_entrega)
    txttel4_entrega.Text = IIf(IsNull(TBLISTA!tel4_entrega), "", TBLISTA!tel4_entrega)
    txtfax_entrega.Text = IIf(IsNull(TBLISTA!fax_entrega), "", TBLISTA!fax_entrega)
    txtemail_entrega.Text = IIf(IsNull(TBLISTA!email_entrega), "", TBLISTA!email_entrega)
    txtsite_entrega.Text = IIf(IsNull(TBLISTA!Site_Entrega), "", TBLISTA!Site_Entrega)
    NomeCampo = "o tipo do endereço"
    If IsNull(TBLISTA!Tipo_endereco) = False And TBLISTA!Tipo_endereco <> "" Then cmbTipo_endereco_entrega.Text = TBLISTA!Tipo_endereco
    NomeCampo = "o tipo do bairro"
    If IsNull(TBLISTA!Tipo_bairro) = False And TBLISTA!Tipo_bairro <> "" Then cmbTipo_bairro_entrega.Text = TBLISTA!Tipo_bairro
    NomeCampo = "o estado"
    If IsNull(TBLISTA!uf_entrega) = False And TBLISTA!uf_entrega <> "" Then cmbuf_entrega = TBLISTA!uf_entrega
    NomeCampo = "a cidade"
    If TBLISTA!uf_entrega <> "EX" Then
        cmbCidade_Entrega.Visible = True
        txtCidade_Entrega.Visible = False
        
        If UCase(TBLISTA!cidade_entrega) = "SANTA BARBARA D'OESTE" Then
        cmbCidade_Entrega.Text = "SANTA BARBARA DO OESTE"
        Else
        cmbCidade_Entrega.Text = IIf(IsNull(TBLISTA!cidade_entrega), "", TBLISTA!cidade_entrega)
        End If
        
    Else
        cmbCidade_Entrega.Visible = False
        txtCidade_Entrega.Visible = True
        txtCidade_Entrega.Text = IIf(IsNull(TBLISTA!cidade_entrega), "", TBLISTA!cidade_entrega)
    End If
1:
    Frame6.Enabled = True
    CodigoLista2 = ListaEntrega.SelectedItem.index
    Novo_Cliente2 = False
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    If Err.Number = "383" Then
        USMsgBox ("Não foi encontrado " & NomeCampo & " do endereço de entrega desse cliente."), vbExclamation, "CAPRIND v5.0"
        GoTo 1
    End If
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error GoTo tratar_erro

USToolBar2.Visible = False
If txtIDcliente = "" Or txtIDcliente = "0" Then
    SSTab1.Tab = 0
    Exit Sub
End If
Select Case SSTab1.Tab
    Case 0:
        With Lista
            .Visible = True
            .Top = Frame1.Top + Frame1.Height
            .Height = Frame15.Top - .Top
        End With
        Frame15.Visible = True
        'PBLista.Visible = True
       ' If txtIDCliente.Visible = True Then txtIDCliente.SetFocus
    Case 1:
        Lista.Visible = False
        Frame15.Visible = False
        'PBLista.Visible = True
        USToolBar2.Visible = True
        ProcVerificaProsseguir
        If Permitido = False Then Exit Sub
        Lista_contato.SetFocus
        ProcCarregaListaContatos
    Case 2:
        Lista.Visible = False
        Frame15.Visible = False
        'PBLista.Visible = True
        USToolBar2.Visible = True
        ProcVerificaProsseguir
        If Permitido = False Then Exit Sub
        ListaEntrega.SetFocus
        Proccarregalistaentrega
    Case 3:
        Lista.Visible = False
        Frame15.Visible = False
        'PBLista.Visible = True
        USToolBar2.Visible = True
        ProcVerificaProsseguir
        If Permitido = False Then Exit Sub
        listacobranca.SetFocus
        ProcCarregalistacobranca
    Case 4:
        Lista.Visible = False
        Frame15.Visible = False
        'PBLista.Visible = True
        USToolBar2.Visible = True
        ProcVerificaProsseguir
        If Permitido = False Then Exit Sub
        lista_familia.SetFocus
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from clientes where IDCliente = " & IIf(txtIDcliente = "", 0, txtIDcliente), Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            ProcCarregaListaFamilia
        End If
        TBAbrir.Close
    Case 5:
        Lista.Visible = False
        Frame15.Visible = False
        'PBLista.Visible = True
        USToolBar2.Visible = True
        ProcVerificaProsseguir
        If Permitido = False Then Exit Sub
        lista_banco.SetFocus
        ProcCarregalista_banco
    Case 6:
        Lista.Visible = False
        Frame15.Visible = False
        'PBLista.Visible = True
        USToolBar2.Visible = True
        ProcVerificaProsseguir
        If Permitido = False Then Exit Sub
        Lista_Impostos.SetFocus
        ProcCarregaLista_Impostos
    Case 7:
        Lista.Visible = False
        Frame15.Visible = False
        'PBLista.Visible = False
        ProcVerificaProsseguir
        If Permitido = False Then Exit Sub
        Cmb_empresa.SetFocus
    Case 8:
        With Lista
            .Visible = True
            .Top = Frame19.Top + Frame19.Height
            .Height = Frame15.Top - .Top
        End With
        Frame15.Visible = True
        'PBLista.Visible = True
        ProcVerificaProsseguir
        If Permitido = False Then Exit Sub
        Lista.SetFocus
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcVerificaProsseguir()
On Error GoTo tratar_erro

Permitido = True
If Novo_Cliente = True Then
    USMsgBox ("Salve o cliente antes de prosseguir."), vbExclamation, "CAPRIND v5.0"
    Permitido = False
    SSTab1.Tab = 0
    Exit Sub
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaFamilia()
On Error GoTo tratar_erro

txtid_familia = 0
txtData4 = Format(Date, "dd/mm/yy")
txtResponsavel4 = pubUsuario
cmbfamilia.ListIndex = -1
CodigoLista4 = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaCombos()
On Error GoTo tratar_erro

ProcCarregaFamilia
ProcCarregaComboPais Txt_pais
ProcCarregaComboEmpresa Cmb_empresa, False
ProcCarregaComboBancoFinanceiro cmbBanco, "txt_Descricao is not null", True
ProcCarregaComboTipoDocto cmbTipo_doc, "Tipo = 'R'"
With txtCategoria
    .Clear
    .AddItem "A"
    .AddItem "B"
    .AddItem "C"
    .AddItem "D"
    .AddItem "E"
    .AddItem "F"
    .AddItem "G"
    .AddItem "H"
    .AddItem "I"
    .AddItem "J"
End With
If txtIDcliente <> "" Then ProcCarregaCamposCombo

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcCarregaCamposCombo()
On Error GoTo tratar_erro

Txt_pais.ListIndex = -1
cmbuf.ListIndex = -1
cmbCidade.ListIndex = -1
cmbtransportadora.ListIndex = -1
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from clientes where idcliente = " & txtIDcliente, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    If IsNull(TBAbrir!Categoria) = False And TBAbrir!Categoria <> "" Then txtCategoria = TBAbrir!Categoria
    If IsNull(TBAbrir!Pais) = False And TBAbrir!Pais <> "" Then Txt_pais = TBAbrir!Pais
    If IsNull(TBAbrir!UF) = False And TBAbrir!UF <> "" Then cmbuf = TBAbrir!UF
    If IsNull(TBAbrir!Cidade) = False And TBAbrir!Cidade <> "" Then cmbCidade = TBAbrir!Cidade
    If IsNull(TBAbrir!txt_transportadora) = False And TBAbrir!txt_transportadora <> "" Then
        Select Case TBAbrir!Tipo_transp
            Case "C": Cmb_tipo_transp = "Cliente"
            Case "F": Cmb_tipo_transp = "Fornecedor"
            Case "E": Cmb_tipo_transp = "Empresa"
        End Select
        cmbtransportadora = TBAbrir!txt_transportadora
    End If
1:
End If

Exit Sub
tratar_erro:
    If Err.Number = "383" Then GoTo 1
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_ISSQN_LostFocus()
On Error GoTo tratar_erro

If Txt_ISSQN.Text <> "" Then
    VerifNumero = Txt_ISSQN.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_ISSQN.Text = ""
        Txt_ISSQN.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregalista_banco()
On Error GoTo tratar_erro

lista_banco.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Compras_fornecedores_banco where id_fornecedor = " & IIf(txtIDcliente = "", 0, txtIDcliente) & " and Tipo = 'C' order by banco", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    TBLISTA.MoveLast
    'PBLista.Min = 0
    'PBLista.Max = TBLISTA.RecordCount
    'PBLista.Value = 1
    Contador = 0
    TBLISTA.MoveFirst
    Do While TBLISTA.EOF = False
        With lista_banco.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Banco), "", TBLISTA!Banco)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Agencia), "", TBLISTA!Agencia)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Conta), "", TBLISTA!Conta)
        End With
        TBLISTA.MoveNext
        Contador = Contador + 1
        'PBLista.Value = Contador
    Loop
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpacampos_banco()
On Error GoTo tratar_erro

txtID_banco = 0
txtData5 = Format(Date, "dd/mm/yy")
txtResponsavel5 = pubUsuario
txtBanco = ""
txtAgencia = ""
txtConta = ""
CodigoLista5 = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAnterior()
On Error GoTo tratar_erro

If txtIDcliente = "" Then Exit Sub
Set TBClientes = CreateObject("adodb.recordset")
TBClientes.Open "Select * from clientes order by NomeRazao", Conexao, adOpenKeyset, adLockOptimistic
If TBClientes.BOF = False Then
    TBClientes.Find ("IDCliente = " & txtIDcliente)
    TBClientes.MovePrevious
    If TBClientes.BOF = False Then
        txtIDcliente = TBClientes!IDCliente
        Set TBClientes = CreateObject("adodb.recordset")
        TBClientes.Open "Select * from clientes where IDCliente = " & IIf(txtIDcliente = "", 0, txtIDcliente), Conexao, adOpenKeyset, adLockOptimistic
        ProcLimpar
        procLimpacamposContatos
        ProcLimpacamposEntrega
        ProcLimpacamposCobranca
        ProcLimpaFamilia
        ProcLimpacampos_banco
        ProcLimpaCampos_Impostos
        ProcLimpaCamposComerciais
        ProcLimpaCampos_Outros
        ProcPuxaDados
        ProcCarregaListaContatos
        Proccarregalistaentrega
        ProcCarregalistacobranca
        ProcCarregaListaFamilia
        ProcCarregalista_banco
        ProcCarregaLista_Impostos
        procPuxadados_Comerciais
        ProcPuxaDados_Outros
    Else
        USMsgBox ("Fim dos cadastros de clientes."), vbInformation, "CAPRIND v5.0"
    End If
End If
Novo_Cliente1 = False
Novo_Cliente2 = False
Novo_Cliente3 = False
Novo_Cliente4 = False
Novo_Cliente5 = False
Novo_Cliente6 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaListaFamilia()
On Error GoTo tratar_erro

lista_familia.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from compras_fornecedores_familia where IDCliente = " & IIf(txtIDcliente = "", 0, txtIDcliente) & " and tipo = 'C'", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    TBLISTA.MoveLast
    'PBLista.Min = 0
    'PBLista.Max = TBLISTA.RecordCount
    'PBLista.Value = 1
    Contador = 0
    TBLISTA.MoveFirst
    Do While TBLISTA.EOF = False
        With lista_familia.ListItems
            .Add , , TBLISTA!idFamilia
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Familia), "", (TBLISTA!Familia))
        End With
        TBLISTA.MoveNext
        Contador = Contador + 1
        'PBLista.Value = Contador
    Loop
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaLista(Pagina As Integer)
On Error GoTo tratar_erro

lblRegistros.Caption = "Nº de registros: 0"
lblPaginas.Caption = "Página: 0 de: 0"
Lista.ListItems.Clear
If StrSql_Cliente = "" Then Exit Sub
Set TBLISTA_Cliente = CreateObject("adodb.recordset")
TBLISTA_Cliente.Open StrSql_Cliente, Conexao, adOpenKeyset, adLockReadOnly
If TBLISTA_Cliente.EOF = False Then ProcExibePagina (Pagina)
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina(Pagina)
On Error GoTo tratar_erro

Lista.ListItems.Clear
TBLISTA_Cliente.PageSize = IIf(txtNreg = "", 30, txtNreg)
TBLISTA_Cliente.AbsolutePage = Pagina
TamanhoPagina = TBLISTA_Cliente.PageSize
ContadorReg = 1

'PBLista.Min = 0
'PBLista.Max = FunVerifMax'PBListaPaginacao(TBLISTA_Cliente.RecordCount - IIf(Pagina > 1, (TBLISTA_Cliente.PageSize * (Pagina - 1)), 0), TBLISTA_Cliente.PageSize)
'PBLista.Value = 1
Contador = 0
Do While TBLISTA_Cliente.EOF = False And (ContadorReg <= TamanhoPagina)
    With Lista.ListItems
        .Add , , TBLISTA_Cliente!IDCliente
        .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA_Cliente!Data), "", Format(TBLISTA_Cliente!Data, "dd/mm/yy"))
        .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA_Cliente!Responsavel), "", TBLISTA_Cliente!Responsavel)
        .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA_Cliente!NomeRazao), "", TBLISTA_Cliente!NomeRazao)
        .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA_Cliente!DtValidacao), "Não", "Sim")
        .Item(.Count).SubItems(5) = TBLISTA_Cliente!ID
    End With
    TBLISTA_Cliente.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    'PBLista.Value = Contador
Loop
lblRegistros.Caption = "Nº de registros: " & TBLISTA_Cliente.RecordCount
If TBLISTA_Cliente.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "Página: 1 de: " & TBLISTA_Cliente.PageCount
ElseIf TBLISTA_Cliente.AbsolutePage = adPosEOF Then
        lblPaginas.Caption = "Página: " & TBLISTA_Cliente.PageCount & " de: " & TBLISTA_Cliente.PageCount
    Else
        lblPaginas.Caption = "Página: " & TBLISTA_Cliente.AbsolutePage - 1 & " de: " & TBLISTA_Cliente.PageCount
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub



Private Sub txtemail_cobranca_LostFocus()
On Error GoTo tratar_erro

txtemail_cobranca = LCase(txtemail_cobranca)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub TxtEmail_Contato_LostFocus()
On Error GoTo tratar_erro

ValidaEmail (TxtEmail_Contato.Text)

TxtEmail_Contato = LCase(TxtEmail_Contato)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtemail_entrega_LostFocus()
On Error GoTo tratar_erro

txtemail_entrega = LCase(txtemail_entrega)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtEmail_LostFocus()
On Error GoTo tratar_erro

txtEmail = LCase(txtEmail)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtIDcliente_Change()
On Error GoTo tratar_erro

If Novo_Cliente = False Then
    ProcLimpar
    ProcLimpaCamposComerciais
    ProcLimpaCampos_Outros
End If
If txtIDcliente <> "" Then
    VerifNumero = txtIDcliente
    ProcVerificaNumero
    If VerifNumero = False Then
        txtIDcliente = ""
        txtIDcliente.SetFocus
        Exit Sub
    End If
    If Novo_Cliente = False Then
        Set TBClientes = CreateObject("adodb.recordset")
        TBClientes.Open "Select * from clientes where idcliente = " & IIf(txtIDcliente = "", 0, txtIDcliente), Conexao, adOpenKeyset, adLockOptimistic
        If TBClientes.EOF = False Then
            ProcPuxaDados
            procPuxadados_Comerciais
            ProcPuxaDados_Outros
            Frame1.Enabled = True
        Else
            Frame1.Enabled = False
        End If
        TBClientes.Close
    End If
End If

If txtIDcliente.Text <> "" Then
    'procAtualizaClienteNuvem
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtLimiteCredito_Change()
On Error GoTo tratar_erro

If txtLimiteCredito.Text <> "" Then
    VerifNumero = txtLimiteCredito.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtLimiteCredito.Text = ""
        txtLimiteCredito.SetFocus
        Exit Sub
    End If
    If txtqtservico = "" Then Exit Sub
End If
    
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

Private Sub txtPorcentagemIPI_Change()
On Error GoTo tratar_erro

If txtPorcentagemIPI.Text <> "" Then
    VerifNumero = txtPorcentagemIPI.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtPorcentagemIPI.Text = ""
        txtPorcentagemIPI.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcEnviaDados()
On Error GoTo tratar_erro

TBGravar!NomeFantasia = IIf(txtnomefantasia = "", Null, txtnomefantasia)
TBGravar!IDCliente = txtIDcliente.Text
TBGravar!Data = IIf(txtData = "", Date, txtData)
TBGravar!Responsavel = IIf(txtResponsavel = "", pubUsuario, txtResponsavel)
TBGravar!NomeRazao = Replace(txtnomerazao, "'", " ")
TBGravar!Endereco = IIf(txtendereco = "", Null, txtendereco)
TBGravar!Bairro = IIf(txtBairro = "", Null, txtBairro)
If cmbuf.Text <> "EX" Then TBGravar!Cidade = cmbCidade Else TBGravar!Cidade = IIf(txtCidade = "", Null, txtCidade)
If cmbRegimeTributario = "Lucro presumido" Then TBGravar!Presumido = True Else TBGravar!Presumido = False
If cmbRegimeTributario = "Simples nacional" Then TBGravar!Simples = True Else TBGravar!Simples = False
If cmbRegimeTributario = "Lucro real" Then TBGravar!Real = True Else TBGravar!Real = False
If cmbRegimeTributario = "MEI" Then TBGravar!MEI = True Else TBGravar!MEI = False
If cmbtransportadora <> "" Then
    Select Case Cmb_tipo_transp
        Case "Cliente": TBGravar!Tipo_transp = "C"
        Case "Fornecedor": TBGravar!Tipo_transp = "F"
        Case "Empresa": TBGravar!Tipo_transp = "E"
    End Select
    TBGravar!txt_transportadora = cmbtransportadora
    TBGravar!idTransp = cmbtransportadora.ItemData(cmbtransportadora.ListIndex)
Else
    TBGravar!Tipo_transp = ""
    TBGravar!txt_transportadora = ""
     TBGravar!idTransp = 0
End If

Select Case cmbPessoa
    Case "Física": TipoPessoa = "FP"
    'Case "Física - Revenda": TipoPessoa = "FR"
    Case "Jurídica": TipoPessoa = "JP"
    'Case "Jurídica - Revenda": TipoPessoa = "JR"
End Select
TBGravar!Tipo = TipoPessoa

TBGravar!Pais = Txt_pais
If Txt_pais <> "" Then TBGravar!Codigo_pais = Txt_pais.ItemData(Txt_pais.ListIndex)
TBGravar!Tipo_endereco = cmbTipo_endereco
TBGravar!Tipo_bairro = cmbTipo_bairro
TBGravar!complemento = IIf(txtComplemento.Text = "", Null, txtComplemento.Text)
TBGravar!Tel01 = txttel01.Text
TBGravar!Fax = txtFax.Text
TBGravar!RG_IE = IIf(txtRG_IE.Text = "", Null, txtRG_IE.Text)
TBGravar!RG_IM = IIf(txtIM_IE = "", Null, txtIM_IE.Text)

If cmbOrigem.Text = "Nacional" Then
    If txtcnpj.Text <> "__.___.___/____-__" Then TBGravar!CPF_CNPJ = txtcnpj.Text Else TBGravar!CPF_CNPJ = txtCpf
ElseIf Novo_Cliente = True Then
        Contador = 0
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from clientes where idTipoEmpresa = 0", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Contador = TBAbrir.RecordCount
        End If
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from Compras_fornecedores where idTipoEmpresa = 0", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Contador = Contador + TBAbrir.RecordCount
        End If
        TBAbrir.Close
        TBGravar!CPF_CNPJ = Contador + 1
End If

TBGravar!Email = IIf(txtEmail.Text = "", Null, LCase(txtEmail.Text))
TBGravar!Site = IIf(txtSite.Text = "", Null, LCase(txtSite.Text))
TBGravar!cxpostal = txtcaixapostal.Text
If cmbOrigem = "" Then
    TBGravar!idTipoEmpresa = Null
ElseIf cmbOrigem.Text = "Nacional" Then
        TBGravar!idTipoEmpresa = 1
    Else
        TBGravar!idTipoEmpresa = 0
End If
If Chk_prospecto.Value = 1 Then TBGravar!Prospecto = True Else TBGravar!Prospecto = False
If Chk_enviar_NF.Value = 1 Then TBGravar!Enviar_NF = True Else TBGravar!Enviar_NF = False
TBGravar!UF = cmbuf.Text
TBGravar!CEP = txtCEP.Text
TBGravar!Numero = txtNumero
TBGravar!Categoria = txtCategoria
If Chk_nao_contribuinte_ICMS.Value = 1 Then TBGravar!Nao_contribuinte_ICMS = True Else TBGravar!Nao_contribuinte_ICMS = False
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcAtualizaGrupo()
On Error GoTo tratar_erro

Conexao.Execute "DELETE from item_aplicacoes where IDGrupo <> " & txtIDGrupo & " and id_cliente_forn = " & txtIDcliente & " and IDGrupo <> 0"
CODIGO = ""
Set TBClientes = CreateObject("adodb.recordset")
TBClientes.Open "select * from item_aplicacoes where IDGrupo = " & txtIDGrupo & " order by n_referencia", Conexao, adOpenKeyset, adLockOptimistic
Do While TBClientes.EOF = False
    If TBClientes!N_referencia <> CODIGO Then
        Set TBFIltro = CreateObject("adodb.recordset")
        TBFIltro.Open "select * from item_aplicacoes", Conexao, adOpenKeyset, adLockOptimistic
        TBFIltro.AddNew
        TBFIltro!Codproduto = IIf(IsNull(TBClientes!Codproduto), "0", TBClientes!Codproduto)
        TBFIltro!ID_cliente_forn = IIf(IsNull(txtIDcliente), "0", txtIDcliente)
        TBFIltro!Tipo = IIf(IsNull(TBClientes!Tipo), "", TBClientes!Tipo)
        TBFIltro!N_referencia = IIf(IsNull(TBClientes!N_referencia), "", TBClientes!N_referencia)
        TBFIltro!Rev = IIf(IsNull(TBClientes!Rev), "", TBClientes!Rev)
        TBFIltro!Desenho = IIf(IsNull(TBClientes!Desenho), "", TBClientes!Desenho)
        TBFIltro!Aplicacao = IIf(IsNull(txtnomerazao), "", txtnomerazao)
        TBFIltro!Descricao = IIf(IsNull(TBClientes!Descricao), "", TBClientes!Descricao)
        TBFIltro!idgrupo = IIf(IsNull(TBClientes!idgrupo), "0", TBClientes!idgrupo)
        TBFIltro.Update
        TBFIltro.Close
    End If
    CODIGO = TBClientes!N_referencia
    TBClientes.MoveNext
Loop
TBClientes.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro: " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcEnviadadosComercial()
On Error GoTo tratar_erro

TBProduto!ID_empresa = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
TBProduto!condicoes = txtCondicoes.Text
TBProduto!calculos = txtcalculos.Text
TBProduto!transporte = txttransporte.Text
TBProduto!impostos = txtimpostos.Text
TBProduto!garantia = txtgarantia.Text
TBProduto!reajuste = txtReajuste.Text
TBProduto!validade = txtValidade.Text
TBProduto!IDCFOP = IIf(txtID_cfop = "", Null, txtID_cfop)
TBProduto!CFOP = txtCFOP
TBProduto!descricaoCFOP = txtOperacao

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCamposComerciais()
On Error GoTo tratar_erro

txtcalculos.Text = "N/A"
txtimpostos.Text = "N/A"
txtCondicoes.Text = "N/A"
txtgarantia.Text = "N/A"
txtReajuste.Text = "N/A"
txttransporte.Text = "N/A"
txtValidade.Text = "N/A"
txtID_cfop.Text = ""
txtCFOP.Text = ""
txtOperacao = ""
CodigoLista7 = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub procPuxadados_Comerciais()
On Error GoTo tratar_erro

Set TBCotacao = CreateObject("adodb.recordset")
TBCotacao.Open "Select * FROM Clientes_DadosComerciais WHERE IDCliente = " & IIf(txtIDcliente = "", 0, txtIDcliente) & " and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex), Conexao, adOpenKeyset, adLockOptimistic
If TBCotacao.EOF = False Then
    txtcalculos.Text = IIf(IsNull(TBCotacao!calculos), "", TBCotacao!calculos)
    txtimpostos.Text = IIf(IsNull(TBCotacao!impostos), "", TBCotacao!impostos)
    txtCondicoes.Text = IIf(IsNull(TBCotacao!condicoes), "", TBCotacao!condicoes)
    txtgarantia.Text = IIf(IsNull(TBCotacao!garantia), "", TBCotacao!garantia)
    txtReajuste.Text = IIf(IsNull(TBCotacao!reajuste), "", TBCotacao!reajuste)
    txttransporte.Text = IIf(IsNull(TBCotacao!transporte), "", TBCotacao!transporte)
    txtValidade.Text = IIf(IsNull(TBCotacao!validade), "", TBCotacao!validade)
    txtID_cfop = IIf(IsNull(TBCotacao!IDCFOP), "", TBCotacao!IDCFOP)
    txtCFOP = IIf(IsNull(TBCotacao!CFOP), "", TBCotacao!CFOP)
    txtOperacao = IIf(IsNull(TBCotacao!descricaoCFOP), "", TBCotacao!descricaoCFOP)
End If
TBCotacao.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcPuxaDados_Outros()
On Error GoTo tratar_erro

txtObservacoes.Text = IIf(IsNull(TBClientes!txt_observacoes), "", TBClientes!txt_observacoes)
txttel02.Text = IIf(IsNull(TBClientes!tel02), "", TBClientes!tel02)
txttel03.Text = IIf(IsNull(TBClientes!tel03), "", TBClientes!tel03)
txttel04.Text = IIf(IsNull(TBClientes!tel04), "", TBClientes!tel04)
If IsNull(TBClientes!idgrupo) = False Then
    Set TBFIltro = CreateObject("adodb.recordset")
    TBFIltro.Open "select * from clientes_grupos where id = " & TBClientes!idgrupo, Conexao, adOpenKeyset, adLockOptimistic
    If TBFIltro.EOF = False Then
        txtIDGrupo = TBClientes!idgrupo
        txtGrupo = IIf(IsNull(TBFIltro!Texto), "", TBFIltro!Texto)
    End If
End If
If TBClientes!chkSuframa = True Then
    txtSuframa = IIf(IsNull(TBClientes!Suframa), 0, TBClientes!Suframa)
    chkSuframa.Value = 1
End If
If TBClientes!SimplesICMSST = True Then chkICMSST.Value = 1 Else chkICMSST.Value = 0
Txt_ISSQN = IIf(IsNull(TBClientes!ISSQN), "", TBClientes!ISSQN)
If IsNull(TBClientes!Banco) = False And TBClientes!Banco <> "" Then cmbBanco = TBClientes!Banco
If IsNull(TBClientes!Tipo_doc) = False And TBClientes!Tipo_doc <> "" Then cmbTipo_doc = TBClientes!Tipo_doc
If IsNull(TBClientes!txtLimiteCredito) = False And TBClientes!txtLimiteCredito <> "" Then txtLimiteCredito = TBClientes!txtLimiteCredito

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtSite_cobranca_LostFocus()
On Error GoTo tratar_erro

txtSite_cobranca = LCase(txtSite_cobranca)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtsite_entrega_LostFocus()
On Error GoTo tratar_erro

txtsite_entrega = LCase(txtsite_entrega)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtSite_LostFocus()
On Error GoTo tratar_erro

Txt_site = LCase(Txt_site)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub btnSintegra_Click()
On Error GoTo tratar_erro

Dim resposta As String
Dim obj As MSXML2.ServerXMLHTTP50
Set obj = New MSXML2.ServerXMLHTTP50
Dim Plugin As String

If cmbPessoa.Text = "" Then
      USMsgBox "Não será possivel buscar o cadastro desse cnpj pois não foi informado o tipo", vbOKOnly, "CAPRIND v5.0"
      cmbPessoa.SetFocus
      Exit Sub
End If

If cmbOrigem.Text = "" Then
      USMsgBox "Não será possivel buscar o cadastro desse cnpj pois não foi informado a Origem", vbOKOnly, "CAPRIND v5.0"
      cmbOrigem.SetFocus
      Exit Sub
End If

If Txt_pais.Text = "" Then
      USMsgBox "Não será possivel buscar o cadastro desse cnpj pois não foi informado o país", vbOKOnly, "CAPRIND v5.0"
      Txt_pais.SetFocus
      Exit Sub
End If


If cmbuf.Text = "" Then
      USMsgBox "Não será possivel buscar o cadastro desse cnpj pois não foi informado a UF", vbOKOnly, "CAPRIND v5.0"
      cmbuf.SetFocus
      Exit Sub
End If

If txtcnpj.Text = "__.___.___/____-__" Then
      USMsgBox "Não será possivel buscar o cadastro desse cnpj pois não foi informado O CNPJ", vbOKOnly, "CAPRIND v5.0"
      txtcnpj.SetFocus
      Exit Sub
End If

Plugin = "ST"

CnpjDados = ReturnNumbersOnly(txtcnpj.Text)

obj.Open "GET", "https://www.sintegraws.com.br/api/v1/execute-api.php?token=1F718E4E-3222-42F1-95D6-995FC9E69C9C&cnpj=" & CnpjDados & "&plugin=" & Plugin & ""

conteudo = CnpjDados
obj.send conteudo

resposta = obj.responseText

If LerDadosJSON(resposta, "status", "", "") = "OK" And LerDadosJSON(resposta, "code", "", "") = "0" Then

USMsgBox LerDadosJSON(resposta, "message", "", ""), vbInformation, "CAPRIND v5.0"
txtnomerazao.Text = UCase(LerDadosJSON(resposta, "nome_empresarial", "", ""))
cmbuf.Text = UCase(LerDadosJSON(resposta, "uf", "", ""))
'txttel01 = LerDadosJSON(resposta, "telefone", "", "")
txtBairro = UCase(LerDadosJSON(resposta, "bairro", "", ""))
txtendereco = UCase(LerDadosJSON(resposta, "logradouro", "", ""))
txtNumero = LerDadosJSON(resposta, "numero", "", "")
txtCEP = LerDadosJSON(resposta, "cep", "", "")
'cmbCidade.Text = UCase(LerDadosJSON(resposta, "municipio", "", ""))
txtnomefantasia = UCase(LerDadosJSON(resposta, "nome_fantasia", "", ""))
cmbRegimeTributario.Text = IIf(LerDadosJSON(resposta, "regime_tributacao", "", "") = "Normal - regime periódico de apuração", "Lucro presumido", "Simples Nacional")
txtRG_IE = Trim(LerDadosJSON(resposta, "inscricao_estadual", "", ""))

txtCategoria.Text = "A"

Cmd_buscarCEP_Click
Else
USMsgBox LerDadosJSON(resposta, "message", "", ""), vbInformation, "CAPRIND v5.0"
txtnomerazao.Text = ""
cmbuf.ListIndex = -1
txtBairro = ""
txtendereco = ""
txtNumero = ""
txtCEP = ""
txtnomefantasia = ""
cmbRegimeTributario.ListIndex = -1
txtRG_IE = ""
txtCategoria.ListIndex = -1

End If




Exit Sub
tratar_erro:
    MousePointer = 0
    If Err.Number = 91 Then
        USMsgBox ("Não foi possível carregar todos os dados referentes a este CEP."), vbInformation, "CAPRIND v5.0"
        Exit Sub
    End If
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
    Case 8: procFiltrar_todos
    Case 9: ProcStatus
    Case 10: ProcValidarRegistros Lista, "Vendas/Clientes"
    Case 11: 'procAtualiza
    Case 12: 'procAtualizaVendedorCliente
    Case 14: ProcAjuda
    Case 15: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar2_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcNovoTab
    Case 2: ProcSalvarTab
    Case 3: ProcExcluirTab
    Case 4: ProcImprimir
    Case 5: ProcAnterior
    Case 6: ProcProximo
    Case 7: 'procAtualizaVendedorCliente
    Case 9: ProcAjuda
    Case 10: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procAtualizaVendedorCliente()
On Error GoTo tratar_erro

Set TBClientes = CreateObject("adodb.recordset")
TBClientes.Open "Select * FROM Vendas_Vendedores_Clientes Order By IDVendedor, IDCliente", Conexao, adOpenKeyset, adLockOptimistic
    If TBClientes.EOF = False Then
    Do While TBClientes.EOF = False
    Set TBAfericao = CreateObject("adodb.recordset")
        TBAfericao.Open "Select * FROM Vendas_Vendedores_Clientes Where IDVendedor = '" & TBClientes!IDvendedor & "' and IDCliente = '" & TBClientes!IDCliente & "' Order By IDVendedor,IDCliente", Conexao, adOpenKeyset, adLockOptimistic
            If TBAfericao.EOF = False Then
            Contador = TBAfericao.RecordCount
            Do While Contador > 1
            'Debug.print TBAfericao!ID
            StrSql = "Delete from Vendas_Vendedores_Clientes where ID = '" & TBAfericao!ID & "'"
            'Debug.print StrSql
            
            Conexao.Execute StrSql
            TBAfericao.MoveNext
            Contador = Contador - 1
            Loop
            End If
            TBAfericao.Close
    TBClientes.MoveNext
    Loop
    End If
    TBClientes.Close

procAtualizaClientesNuvem
USMsgBox "Vendedores atualizados com sucesso!", vbInformation, "CAPRIND v5.0"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar3_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcSalvarTab
    Case 2: ProcExcluirTab
    Case 3: ProcImprimir
    Case 4: ProcAnterior
    Case 5: ProcProximo
    Case 7: ProcAjuda
    Case 8: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar4_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcSalvarTab
    Case 2: ProcImprimir
    Case 3: ProcAnterior
    Case 4: ProcProximo
    Case 6: ProcAjuda
    Case 7: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro
  
frmVendas_Cliente_RelUF.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcNovoTab()
On Error GoTo tratar_erro

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Select Case SSTab1.Tab
    Case 1: If FunVerificaRegistroValidado("Clientes", "IDcliente = " & txtIDcliente, "cliente", "contato", "criar novo", True, True) = False Then Exit Sub
            procNovo_Contato
    Case 2: If FunVerificaRegistroValidado("Clientes", "IDcliente = " & txtIDcliente, "cliente", "endereço de entrega", "criar novo", True, True) = False Then Exit Sub
            procNovo_entrega
    Case 3: If FunVerificaRegistroValidado("Clientes", "IDcliente = " & txtIDcliente, "cliente", "endereço de cobrança", "criar novo", True, True) = False Then Exit Sub
            procNovo_cobranca
    Case 4: If FunVerificaRegistroValidado("Clientes", "IDcliente = " & txtIDcliente, "cliente", "família", "criar novo", True, True) = False Then Exit Sub
            procNovo_familia
    Case 5: If FunVerificaRegistroValidado("Clientes", "IDcliente = " & txtIDcliente, "cliente", "dados bancários", "criar novos", True, True) = False Then Exit Sub
            procNovo_Banco
    Case 6: If FunVerificaRegistroValidado("Clientes", "IDcliente = " & txtIDcliente, "cliente", "imposto", "criar novo", True, True) = False Then Exit Sub
            procNovo_impostos
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcSalvarTab()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Select Case SSTab1.Tab
    Case 1: procSalvar_contato
    Case 2: procSalvar_entrega
    Case 3: procSalvar_cobranca
    Case 4: procSalvar_familia
    Case 5: procSalvar_Banco
    Case 6: procSalvar_impostos
    Case 7: procSalvar_comerciais
    Case 8: procSalvar_outros
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcExcluirTab()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Select Case SSTab1.Tab
    Case 1: procExcluir_contato
    Case 2: procExcluir_entrega
    Case 3: procExcluir_cobranca
    Case 4: procExcluir_familia
    Case 5: procExcluir_banco
    Case 6: procExcluir_impostos
    Case 7: procExcluir_comerciais
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
