VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCFI 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Estoque - Almoxarifado"
   ClientHeight    =   10035
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15360
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MousePointer    =   99  'Custom
   ScaleHeight     =   10035
   ScaleWidth      =   15360
   WindowState     =   2  'Maximized
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   99
      ScreenHeight    =   1080
      ScreenWidth     =   2560
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   60
      TabIndex        =   26
      Top             =   5430
      Width           =   15195
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
         TabIndex        =   8
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
         Left            =   3750
         TabIndex        =   7
         Text            =   "30"
         ToolTipText     =   "Número de registros por página."
         Top             =   180
         Width           =   555
      End
      Begin DrawSuite2022.USButton cmdPagProx 
         Height          =   315
         Left            =   11760
         TabIndex        =   12
         ToolTipText     =   "Próxima página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmCFI.frx":0000
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
         TabIndex        =   11
         ToolTipText     =   "Página anterior."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmCFI.frx":37A4
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
         TabIndex        =   9
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
         TabIndex        =   10
         ToolTipText     =   "Primeira página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmCFI.frx":72AD
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
         TabIndex        =   13
         ToolTipText     =   "Última página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmCFI.frx":B39C
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
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "registros por página"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   4380
         TabIndex        =   31
         Top             =   240
         Width           =   1440
      End
      Begin VB.Label lblRegistros 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nº de registros: 0"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   180
         TabIndex        =   29
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label lblPaginas 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Página: 0 de: 0"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   13050
         TabIndex        =   28
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Carregar"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   3060
         TabIndex        =   27
         Top             =   240
         Width           =   645
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   23
      Top             =   0
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   1720
      ButtonCount     =   12
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
      ButtonLeft2     =   37
      ButtonTop2      =   2
      ButtonWidth2    =   36
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
      ButtonLeft3     =   75
      ButtonTop3      =   2
      ButtonWidth3    =   39
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft4     =   116
      ButtonTop4      =   2
      ButtonWidth4    =   51
      ButtonHeight4   =   21
      ButtonUseMaskColor4=   0   'False
      ButtonCaption5  =   "Filtrar pendentes"
      ButtonEnabled5  =   0   'False
      ButtonIconSize5 =   32
      ButtonToolTipText5=   "Filtrar produtos/itens pendentes (F7)"
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
      ButtonLeft5     =   169
      ButtonTop5      =   2
      ButtonWidth5    =   90
      ButtonHeight5   =   21
      ButtonUseMaskColor5=   0   'False
      ButtonCaption6  =   "Retirar"
      ButtonEnabled6  =   0   'False
      ButtonIconSize6 =   32
      ButtonToolTipText6=   "Retirar (F8)"
      ButtonKey6      =   "5"
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
      ButtonLeft6     =   261
      ButtonTop6      =   2
      ButtonWidth6    =   41
      ButtonHeight6   =   21
      ButtonUseMaskColor6=   0   'False
      ButtonCaption7  =   "Devolver"
      ButtonEnabled7  =   0   'False
      ButtonIconSize7 =   32
      ButtonToolTipText7=   "Devolver (F9)"
      ButtonKey7      =   "6"
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
      ButtonLeft7     =   304
      ButtonTop7      =   2
      ButtonWidth7    =   51
      ButtonHeight7   =   21
      ButtonUseMaskColor7=   0   'False
      ButtonCaption8  =   "Atualizar"
      ButtonEnabled8  =   0   'False
      ButtonIconSize8 =   32
      ButtonToolTipText8=   "Utilizado pelo administrador do sistema."
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
      ButtonLeft8     =   357
      ButtonTop8      =   2
      ButtonWidth8    =   50
      ButtonHeight8   =   21
      ButtonUseMaskColor8=   0   'False
      ButtonEnabled9  =   0   'False
      ButtonIconSize9 =   32
      ButtonAlignment9=   2
      ButtonType9     =   1
      ButtonStyle9    =   -1
      BeginProperty ButtonFont9 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState9    =   -1
      ButtonLeft9     =   409
      ButtonTop9      =   4
      ButtonWidth9    =   2
      ButtonHeight9   =   54
      ButtonCaption10 =   "Ajuda"
      ButtonEnabled10 =   0   'False
      ButtonIconSize10=   32
      ButtonToolTipText10=   "Ajuda (F1)"
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
      ButtonLeft10    =   413
      ButtonTop10     =   2
      ButtonWidth10   =   36
      ButtonHeight10  =   21
      ButtonUseMaskColor10=   0   'False
      ButtonCaption11 =   "Sair"
      ButtonEnabled11 =   0   'False
      ButtonIconSize11=   32
      ButtonToolTipText11=   "Sair (Esc)"
      ButtonKey11     =   "11"
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
      ButtonLeft11    =   451
      ButtonTop11     =   2
      ButtonWidth11   =   26
      ButtonHeight11  =   21
      ButtonUseMaskColor11=   0   'False
      ButtonEnabled12 =   0   'False
      ButtonIconSize12=   32
      ButtonKey12     =   "12"
      ButtonAlignment12=   2
      BeginProperty ButtonFont12 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState12   =   5
      ButtonLeft12    =   479
      ButtonTop12     =   2
      ButtonWidth12   =   24
      ButtonHeight12  =   24
      ButtonUseMaskColor12=   0   'False
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   13770
         Top             =   180
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmCFI.frx":EC28
         Count           =   1
      End
   End
   Begin VB.TextBox txtId 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   510
      TabIndex        =   21
      Text            =   "0"
      Top             =   3540
      Visible         =   0   'False
      Width           =   435
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   2700
      Left            =   60
      TabIndex        =   6
      Top             =   2700
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   4763
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
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   17
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "T"
         Text            =   "Status"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Cód. interno"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "Descrição"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Família"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Tag             =   "N"
         Text            =   "N° RE"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Object.Tag             =   "T"
         Text            =   "Nº lote"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Object.Tag             =   "T"
         Text            =   "Cód. de ref."
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Object.Tag             =   "T"
         Text            =   "Funcionário"
         Object.Width           =   4234
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Object.Tag             =   "T"
         Text            =   "Posto de trabalho"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Object.Tag             =   "N"
         Text            =   "Qtde. retir."
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   11
         Object.Tag             =   "D"
         Text            =   "Dt. retir."
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   12
         Object.Tag             =   "D"
         Text            =   "Prev. devol."
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   13
         Object.Tag             =   "N"
         Text            =   "Qtde. devol."
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   14
         Object.Tag             =   "N"
         Text            =   "Qtde. dev. c/ probl."
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   15
         Object.Tag             =   "D"
         Text            =   "Dt. devol."
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Object.Tag             =   "T"
         Text            =   "Observação"
         Object.Width           =   0
      EndProperty
   End
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
      Height          =   1695
      Left            =   60
      TabIndex        =   16
      Top             =   990
      Width           =   15195
      Begin VB.TextBox txtobs 
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
         Height          =   615
         Left            =   180
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Observações."
         Top             =   945
         Width           =   14805
      End
      Begin VB.TextBox txtcodinterno 
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
         MaxLength       =   50
         TabIndex        =   0
         TabStop         =   0   'False
         ToolTipText     =   "Código interno."
         Top             =   370
         Width           =   1605
      End
      Begin VB.TextBox txtquantestoque 
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
         Left            =   13170
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade em estoque."
         Top             =   370
         Width           =   1815
      End
      Begin VB.CommandButton cmdLocalizar_produto 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1800
         Picture         =   "frmCFI.frx":157CA
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Localizar produtos."
         Top             =   370
         Width           =   315
      End
      Begin VB.TextBox txtfamilia 
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
         Left            =   9090
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Família."
         Top             =   370
         Width           =   4065
      End
      Begin VB.TextBox txtDescricao 
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
         Left            =   2220
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Descrição."
         Top             =   370
         Width           =   6855
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Observações"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   7110
         TabIndex        =   22
         Top             =   750
         Width           =   945
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Qtd. estoque"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   13597
         TabIndex        =   20
         Top             =   180
         Width           =   960
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Cód. interno*"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   487
         TabIndex        =   19
         Top             =   180
         Width           =   990
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Família"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   10882
         TabIndex        =   18
         Top             =   180
         Width           =   480
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   5302
         TabIndex        =   17
         Top             =   180
         Width           =   690
      End
   End
   Begin DrawSuite2022.USProgressBar PBLista1 
      Height          =   255
      Left            =   75
      TabIndex        =   24
      Top             =   9750
      Width           =   15105
      _ExtentX        =   26644
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
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   60
      TabIndex        =   30
      Top             =   6060
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   4635
      Left            =   60
      TabIndex        =   25
      Top             =   6360
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   8176
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
      TabCaption(0)   =   "Lista de movimentação do estoque"
      TabPicture(0)   =   "frmCFI.frx":158CC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Lista_movimentacao"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Destino/aplicação"
      TabPicture(1)   =   "frmCFI.frx":158E8
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Lista_destino"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin MSComctlLib.ListView Lista_movimentacao 
         Height          =   3060
         Left            =   45
         TabIndex        =   14
         Top             =   330
         Width           =   15105
         _ExtentX        =   26644
         _ExtentY        =   5398
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
            Object.Tag             =   "T"
            Text            =   "Código"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Operação"
            Object.Width           =   6747
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Object.Tag             =   "D"
            Text            =   "Data"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Object.Tag             =   "N"
            Text            =   "Entrada"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Object.Tag             =   "N"
            Text            =   "Saída"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Object.Tag             =   "T"
            Text            =   "Documento"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Object.Tag             =   "T"
            Text            =   "Responsável"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Object.Tag             =   "T"
            Text            =   "Requisitante"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "IDestoque"
            Object.Width           =   0
         EndProperty
      End
      Begin MSComctlLib.ListView Lista_destino 
         Height          =   3060
         Left            =   -74955
         TabIndex        =   15
         Top             =   330
         Width           =   15105
         _ExtentX        =   26644
         _ExtentY        =   5398
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
            Text            =   "Cod produto"
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
            Object.Width           =   13626
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "Família"
            Object.Width           =   8996
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Object.Tag             =   "T"
            Text            =   "Un."
            Object.Width           =   1058
         EndProperty
      End
   End
End
Attribute VB_Name = "frmCFI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public IDFI As Long 'OK
Public SQL_almoxarifado As String 'OK
Public CFI_devolucao As Boolean 'OK
Public CFI_saida As Boolean 'OK
Public FormulaRel_CFI As String 'OK
Dim TBLISTA_CFI As ADODB.Recordset 'OK

Sub ProcAjuda()
On Error GoTo tratar_erro

FunAbrirVideoWeb ("http://www.youtube.com/watch?v=-yg4rW8aF1Y&list=UUODGiDjQ-BCrxh0YXfJvoqg&index=42&feature=plcp")

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
Lista_Movimentacao.ListItems.Clear
Lista_destino.ListItems.Clear
If SQL_almoxarifado = "" Then Exit Sub
Set TBLISTA_CFI = CreateObject("adodb.recordset")
TBLISTA_CFI.Open SQL_almoxarifado, Conexao, adOpenKeyset, adLockReadOnly
If TBLISTA_CFI.EOF = False Then ProcExibePagina (Pagina)
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina(Pagina)
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista_Movimentacao.ListItems.Clear
Lista_destino.ListItems.Clear
TBLISTA_CFI.PageSize = IIf(txtNreg = "", 30, txtNreg)
TBLISTA_CFI.AbsolutePage = Pagina
TamanhoPagina = TBLISTA_CFI.PageSize
ContadorReg = 1

PBLista.Min = 0
PBLista.Max = FunVerifMaxPBListaPaginacao(TBLISTA_CFI.RecordCount - IIf(Pagina > 1, (TBLISTA_CFI.PageSize * (Pagina - 1)), 0), TBLISTA_CFI.PageSize)
PBLista.Value = 1
Contador = 0
Do While TBLISTA_CFI.EOF = False And (ContadorReg <= TamanhoPagina)
    With Lista.ListItems
        .Add , , TBLISTA_CFI!IDCFI
        .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA_CFI!status), "", TBLISTA_CFI!status)
        .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA_CFI!Codigo_produto), "", TBLISTA_CFI!Codigo_produto)
        .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA_CFI!Descricao), "", TBLISTA_CFI!Descricao)
        .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA_CFI!Familia), "", TBLISTA_CFI!Familia)
        .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA_CFI!IDEstoque), "", TBLISTA_CFI!IDEstoque)
        .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA_CFI!LOTE), "", TBLISTA_CFI!LOTE)
        .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA_CFI!Ref), "", TBLISTA_CFI!Ref)
        .Item(.Count).SubItems(8) = IIf(IsNull(TBLISTA_CFI!Funcionario), "", TBLISTA_CFI!Funcionario)
        .Item(.Count).SubItems(9) = IIf(IsNull(TBLISTA_CFI!maquina), "", TBLISTA_CFI!maquina)
        .Item(.Count).SubItems(10) = Format(TBLISTA_CFI!Quantretirada, "###,##0.0000")
        .Item(.Count).SubItems(11) = Format(TBLISTA_CFI!Dataretirada, "dd/mm/yy")
        .Item(.Count).SubItems(12) = Format(TBLISTA_CFI!dataprevisao, "dd/mm/yy")
        .Item(.Count).SubItems(13) = IIf(IsNull(TBLISTA_CFI!Quantdevolvido), "", Format(TBLISTA_CFI!Quantdevolvido, "###,##0.0000"))
        .Item(.Count).SubItems(14) = IIf(IsNull(TBLISTA_CFI!Quantdevolvidoprobl), "", Format(TBLISTA_CFI!Quantdevolvidoprobl, "###,##0.0000"))
        .Item(.Count).SubItems(15) = IIf(IsNull(TBLISTA_CFI!Datadevolucao), "", Format(TBLISTA_CFI!Datadevolucao, "dd/mm/yy"))
        .Item(.Count).SubItems(16) = IIf(IsNull(TBLISTA_CFI!Observacao), "", TBLISTA_CFI!Observacao)
    End With
    TBLISTA_CFI.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista.Value = Contador
Loop
lblRegistros.Caption = "Nº de registros: " & TBLISTA_CFI.RecordCount
If TBLISTA_CFI.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "Página: 1 de: " & TBLISTA_CFI.PageCount
ElseIf TBLISTA_CFI.AbsolutePage = adPosEOF Then
        lblPaginas.Caption = "Página: " & TBLISTA_CFI.PageCount & " de: " & TBLISTA_CFI.PageCount
    Else
        lblPaginas.Caption = "Página: " & TBLISTA_CFI.AbsolutePage - 1 & " de: " & TBLISTA_CFI.PageCount
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCampos()
On Error GoTo tratar_erro

txtId = 0
txtCodinterno.Text = ""
txtdescricao.Text = ""
txtquantestoque.Text = ""
txtfamilia.Text = ""
txtObs.Text = ""
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAbrir()
On Error GoTo tratar_erro

frmCFI_Localizar.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procFiltrar_todos()
On Error GoTo tratar_erro

ProcLimpaCampos
INNERJOINTEXTO = "Select CFI.*, CM.maquina, EC.Ref from ((CFI INNER JOIN Projproduto P ON P.desenho = CFI.Codigo_produto) LEFT JOIN Cadmaquinas CM ON CM.idmaquina = CFI.ID_Maquina) LEFT JOIN Estoque_controle EC ON EC.IDestoque = CFI.IDestoque"
If Qualidade_Almox = True Then
    Caption = "Qualidade - Almoxarifado"
    InstFiltro = "and P.Instrumento = 'True'"
    InstFiltroRel = "and {Projproduto.Instrumento} = True"
Else
    InstFiltro = ""
    InstFiltroRel = ""
End If
SQL_almoxarifado = INNERJOINTEXTO & " where (CFI.status = 'EM ABERTO' or CFI.status = 'DEVOLVIDO PARCIAL') " & InstFiltro & " order by CFI.dataretirada, CFI.codigo_produto, CFI.lote, CFI.idcfi"
FormulaRel_CFI = "({CFI.Status} = 'EM ABERTO' OR {CFI.Status} = 'DEVOLVIDO PARCIAL') " & InstFiltroRel
ProcCarregaLista (1)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdLocalizar_produto_Click()
On Error GoTo tratar_erro

CFI_saida = False
CFI_devolucao = False
frmCFI_locprod.Show 1

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
ProcLimpaCampos
Frame2.Enabled = True
CFI_devolucao = False
CFI_saida = False
frmCFI_locprod.Show 1
    
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

Private Sub ProcAtualizar()
On Error GoTo tratar_erro

If InputBox("Informe a senha para liberar.") = "280362A" Then
    If USMsgBox("Deseja realmente atualizar as movimentações?", vbYesNo, "CAPRIND v5.0") = vbYes Then
        Set TBFI = CreateObject("adodb.recordset")
        TBFI.Open "Select * from CFI order by IDCFI", Conexao, adOpenKeyset, adLockOptimistic
        If TBFI.EOF = False Then
            TBFI.MoveLast
            PBLista1.Min = 0
            PBLista1.Max = TBFI.RecordCount
            PBLista1.Value = 1
            Contador = 0
            TBFI.MoveFirst
            Do While TBFI.EOF = False
                Set TBEstoque = CreateObject("adodb.recordset")
                TBEstoque.Open "Select * from Estoque_controle where Lote = '" & TBFI!LOTE & "' and Desenho  = '" & TBFI!Codigo_produto & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBEstoque.EOF = False Then
                    If TBFI!Quantretirada > 0 Then
                        Quant = TBFI!Quantretirada
                        NovoValor = Replace(Quant, ",", ".")
                        Set TBGravar = CreateObject("adodb.recordset")
                        TBGravar.Open "Select * from Estoque_movimentacao where ID_cfi = " & TBFI!IDCFI & " and IDestoque = " & TBEstoque!IDEstoque & " and Lote = '" & TBEstoque!LOTE & "' and Operacao = 'SAIDA_ALMOXARIFADO' and Data = '" & TBFI!Dataretirada & "' and Saida = " & NovoValor, Conexao, adOpenKeyset, adLockOptimistic
                        If TBGravar.EOF = True Then
                            TBGravar.AddNew
                            TBGravar!Destino = "Interno"
                            TBGravar!Terceiros = False
                        End If
                        TBGravar!Id_cfi = TBFI!IDCFI
                        TBGravar!IDEstoque = TBEstoque!IDEstoque
                        TBGravar!Operacao = "SAIDA_ALMOXARIFADO"
                        TBGravar!Desenho = TBEstoque!Desenho
                        TBGravar!Descricao = TBEstoque!Descricao
                        TBGravar!Data = TBFI!Dataretirada
                        TBGravar!Saida = Quant
                        TBGravar!Responsavel = IIf(IsNull(TBFI!Responsavel), pubUsuario, TBFI!Responsavel)
                        TBGravar!LOTE = TBEstoque!LOTE
                        TBGravar!Requisitante = TBFI!Funcionario
                        TBGravar!VlrUnit = IIf(IsNull(TBEstoque!valor_unitario), 0, TBEstoque!valor_unitario)
                        TBGravar!vlrTotal = Format(IIf(IsNull(TBEstoque!valor_unitario), 0, TBEstoque!valor_unitario) * Quant, "###,##0.00")
                        TBGravar!Familia = IIf(IsNull(TBEstoque!Classe), "", TBEstoque!Classe)
                        TBGravar!Destino = "Interno"
                        TBGravar.Update
                        TBGravar.Close
                    End If
                        
                    If TBFI!Quantdevolvido > 0 Then
                        Quant = TBFI!Quantdevolvido
                        NovoValor = Replace(Quant, ",", ".")
                        Set TBGravar = CreateObject("adodb.recordset")
                        TBGravar.Open "Select * from Estoque_movimentacao where ID_cfi = " & TBFI!IDCFI & " and IDestoque = " & TBEstoque!IDEstoque & " and Lote = '" & TBEstoque!LOTE & "' and Operacao = 'DEVOLUCAO_ALMOXARIFADO' and Data = '" & TBFI!Datadevolucao & "' and Entrada = " & NovoValor, Conexao, adOpenKeyset, adLockOptimistic
                        If TBGravar.EOF = True Then
                            TBGravar.AddNew
                            TBGravar!Destino = "Interno"
                            TBGravar!Terceiros = False
                        End If
                        TBGravar!Id_cfi = TBFI!IDCFI
                        TBGravar!IDEstoque = TBEstoque!IDEstoque
                        TBGravar!Operacao = "DEVOLUCAO_ALMOXARIFADO"
                        TBGravar!Desenho = TBEstoque!Desenho
                        TBGravar!Descricao = TBEstoque!Descricao
                        TBGravar!Data = TBFI!Datadevolucao
                        TBGravar!Entrada = Quant
                        TBGravar!Responsavel = IIf(IsNull(TBFI!Responsavel), pubUsuario, TBFI!Responsavel)
                        TBGravar!LOTE = TBEstoque!LOTE
                        TBGravar!VlrUnit = IIf(IsNull(TBEstoque!valor_unitario), 0, TBEstoque!valor_unitario)
                        TBGravar!vlrTotal = Format(IIf(IsNull(TBEstoque!valor_unitario), 0, TBEstoque!valor_unitario) * Quant, "###,##0.00")
                        TBGravar!Familia = IIf(IsNull(TBEstoque!Classe), "", TBEstoque!Classe)
                        TBGravar!Destino = "Interno"
                        TBGravar.Update
                        TBGravar.Close
                    End If
                    
                    If TBFI!Quantdevolvidoprobl > 0 Then
                        Quant = TBFI!Quantdevolvidoprobl
                        NovoValor = Replace(Quant, ",", ".")
                        Set TBGravar = CreateObject("adodb.recordset")
                        TBGravar.Open "Select * from Estoque_movimentacao where ID_cfi = " & TBFI!IDCFI & " and IDestoque = " & TBEstoque!IDEstoque & " and Lote = '" & TBEstoque!LOTE & "' and Operacao = 'DEVOLUCAO_ALMOXARIFADO C/ PROB.' and Data = '" & TBFI!Datadevolucao & "' and Entrada = " & NovoValor, Conexao, adOpenKeyset, adLockOptimistic
                        If TBGravar.EOF = True Then
                            TBGravar.AddNew
                            TBGravar!Destino = "Interno"
                            TBGravar!Terceiros = False
                        End If
                        TBGravar!Id_cfi = TBFI!IDCFI
                        TBGravar!IDEstoque = TBEstoque!IDEstoque
                        TBGravar!Operacao = "DEVOLUCAO_ALMOXARIFADO C/ PROB."
                        TBGravar!Desenho = TBEstoque!Desenho
                        TBGravar!Descricao = TBEstoque!Descricao
                        TBGravar!Data = TBFI!Datadevolucao
                        TBGravar!Entrada = Quant
                        TBGravar!Responsavel = IIf(IsNull(TBFI!Responsavel), pubUsuario, TBFI!Responsavel)
                        TBGravar!LOTE = TBEstoque!LOTE
                        TBGravar!VlrUnit = IIf(IsNull(TBEstoque!valor_unitario), 0, TBEstoque!valor_unitario)
                        TBGravar!vlrTotal = Format(IIf(IsNull(TBEstoque!valor_unitario), 0, TBEstoque!valor_unitario) * Quant, "###,##0.00")
                        TBGravar!Familia = IIf(IsNull(TBEstoque!Classe), "", TBEstoque!Classe)
                        TBGravar!Destino = "Interno"
                        TBGravar.Update
                        TBGravar.Close
                    End If
                End If
                TBEstoque.Close
                
                TBFI.MoveNext
                Contador = Contador + 1
                PBLista1.Value = Contador
            Loop
        End If
        
        USMsgBox ("Atualização efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
        '==================================
        Modulo = Formulario
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

Private Sub cmdPagAnt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_CFI.AbsolutePage <> 2 Then
    If TBLISTA_CFI.AbsolutePage = -3 Then
        ProcExibePagina (TBLISTA_CFI.PageCount - 1)
    Else
        TBLISTA_CFI.AbsolutePage = TBLISTA_CFI.AbsolutePage - 2
        ProcExibePagina (TBLISTA_CFI.AbsolutePage)
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
    TBLISTA_CFI.AbsolutePage = txtPagIr.Text
    ProcExibePagina (TBLISTA_CFI.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_CFI.AbsolutePage = 1
ProcExibePagina (TBLISTA_CFI.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_CFI.AbsolutePage <> -3 Then
    If TBLISTA_CFI.AbsolutePage = 1 Then
        ProcExibePagina (2)
    Else
        ProcExibePagina (TBLISTA_CFI.AbsolutePage)
    End If
Else
    ProcExibePagina (TBLISTA_CFI.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_CFI.AbsolutePage = TBLISTA_CFI.PageCount
ProcExibePagina (TBLISTA_CFI.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyInsert: ProcNovo
    Case vbKeyF2: ProcAbrir
    Case vbKeyF4: ProcExcluir
    Case vbKeyF5: ProcImprimir
    Case vbKeyF7: procFiltrar_todos
    Case vbKeyF8: ProcRetirada
    Case vbKeyF9: ProcDevolucao
    Case vbKeyF1: ProcAjuda
    Case vbKeyEscape: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15195, 12, True
If Qualidade_Almox = False Then Formulario = "Estoque/Almoxarifado" Else Formulario = "Qualidade/Almoxarifado"
Direitos
ProcLimpaVariaveisPrincipais
SSTab1.Tab = 0
CFI_devolucao = False
CFI_saida = False
procFiltrar_todos

ProcRemoveObjetosResize Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcDevolucao()
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 Or txtCodinterno.Text = "" Or txtId.Text = "0" Then Exit Sub
CFI_saida = False
CFI_devolucao = True
frmCFI_Devolucao.Show 1

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

qtdeliberada = 0
qtdeliberar = 0
With Lista_Movimentacao
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If .ListItems(InitFor).SubItems(2) <> "DEVOLUCAO_ALMOXARIFADO C/ PROB." Then qtdeliberada = qtdeliberada + .ListItems.Item(InitFor).SubItems(4)            'Entrada
            qtdeliberar = qtdeliberar + .ListItems.Item(InitFor).SubItems(5) 'Saída
        End If
    Next InitFor
End With
Qtde = IIf(txtquantestoque = "", 0, txtquantestoque)
If ((Qtde + qtdeliberar) - qtdeliberada) < 0 Then
    USMsgBox ("Não é permitido excluir essa(s) movimentação(ões), pois o saldo ficará negativo."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

Permitido = False
With Lista_Movimentacao
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir esta(s) movimentação(ões)?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            
            Conexao.Execute "DELETE from Estoque_movimentacao where IDoperacao = " & .ListItems(InitFor)
            
            lista_v = .ListItems(InitFor).SubItems(2)
            If lista_v = "SAIDA_ALMOXARIFADO" Or lista_v = "DEVOLUCAO_ALMOXARIFADO" Then
                Set TBEstoque = CreateObject("adodb.recordset")
                TBEstoque.Open "Select * from estoque_controle where IDestoque = " & .ListItems.Item(InitFor).ListSubItems(9), Conexao, adOpenKeyset, adLockOptimistic
                If TBEstoque.EOF = False Then
                    If lista_v = "SAIDA_ALMOXARIFADO" Then
                        QuantSolicitado = .ListItems.Item(InitFor).ListSubItems(5)
                        TBEstoque!estoque_real = Format(TBEstoque!estoque_real + QuantSolicitado, "###,##0.00")
                    Else
                        QuantSolicitado = .ListItems.Item(InitFor).ListSubItems(4)
                        TBEstoque!estoque_real = Format(TBEstoque!estoque_real - QuantSolicitado, "###,##0.00")
                    End If
                    TBEstoque!Valor_total = Format(TBEstoque!valor_unitario * TBEstoque!estoque_real, "###,##0.00")
                    TBEstoque.Update
                End If
                TBEstoque.Close
            End If
            
            Set TBEstoque = CreateObject("adodb.recordset")
            TBEstoque.Open "Select * from Estoque_movimentacao where ID_cfi = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
            If TBEstoque.EOF = False Then
                Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open "Select Sum(Entrada) as quantidade, Sum(Saida) as QuantEmpenho from Estoque_movimentacao where ID_cfi = " & txtId & " and (Operacao = 'SAIDA_ALMOXARIFADO' or Operacao = 'DEVOLUCAO_ALMOXARIFADO')", Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    quantidade = IIf(IsNull(TBAbrir!quantidade), 0, TBAbrir!quantidade)
                    QuantEmpenho = IIf(IsNull(TBAbrir!QuantEmpenho), 0, TBAbrir!QuantEmpenho)
                End If
                
                Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open "Select Sum(Entrada) as quantestoque from Estoque_movimentacao where ID_cfi = " & txtId & " and Operacao = 'DEVOLUCAO_ALMOXARIFADO C/ PROB.'", Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    quantestoque = IIf(IsNull(TBAbrir!quantestoque), 0, TBAbrir!quantestoque)
                End If
                
                Dataini = 0
                Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open "Select Data from Estoque_movimentacao where ID_cfi = " & txtId & " and (Operacao = 'DEVOLUCAO_ALMOXARIFADO' or Operacao = 'DEVOLUCAO_ALMOXARIFADO C/ PROB.') order by Data", Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    TBAbrir.MoveLast
                    Dataini = TBAbrir!Data
                End If
                TBAbrir.Close
                
                NovoValor = Replace(quantidade, ",", ".") 'Devolução
                NovoValor1 = Replace(QuantEmpenho, ",", ".") 'Saida
                NovoValor2 = Replace(quantestoque, ",", ".") 'Devolução c/ prob.
                
                CamposFiltro = "Quantretirada = " & NovoValor1 & ", Quantdevolvido = " & NovoValor & ", Quantdevolvidoprobl = " & NovoValor2
                If QuantEmpenho > (quantidade + quantestoque) And (quantidade + quantestoque) > 0 Then
                    If quantestoque = 0 Then
                        CamposFiltro = CamposFiltro & ", DataDevolucao = '" & Dataini & "', Status = 'DEVOLVIDO PARCIAL', Observacao = NULL, Restricao = 'False'"
                    Else
                        CamposFiltro = CamposFiltro & ", DataDevolucao = '" & Dataini & "', Status = 'DEVOLVIDO PARCIAL',"
                    End If
                ElseIf quantidade + quantestoque = 0 Then
                        CamposFiltro = CamposFiltro & ", DataDevolucao = NULL, Status = 'EM ABERTO', Observacao = NULL, Restricao = 'False'"
                    Else
                        CamposFiltro = CamposFiltro & ", DataDevolucao = '" & Dataini & "', Status = 'DEVOLVIDO'"
                End If
                Conexao.Execute "Update CFI Set " & CamposFiltro & " where IDcfi = " & txtId
            Else
                Conexao.Execute "DELETE from cfi where idcfi = " & txtId
            End If
            TBEstoque.Close
            
            '==================================
            Modulo = Formulario
            Evento = "Excluir movimentação"
            ID_documento = Lista_Movimentacao.SelectedItem
            Documento = "Cód. interno: " & Lista.SelectedItem.ListSubItems(2) & " - RE: " & Lista.SelectedItem.ListSubItems(5) & " - Lote: " & Lista.SelectedItem.ListSubItems(6)
            Documento1 = "Operação: " & .ListItems.Item(InitFor).SubItems(2) & " - Data: " & .ListItems.Item(InitFor).SubItems(3) & " - Entrada: " & .ListItems.Item(InitFor).SubItems(3) & " - Saída: " & .ListItems.Item(InitFor).SubItems(4) & " - Documento: " & .ListItems.Item(InitFor).SubItems(6)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe a(s) movimentação(ões) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Movimentação(ões) excluída(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    Frame2.Enabled = False
    ProcCarregaLista (1)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 Then Exit Sub
NomeRel = "Almoxarifado.rpt"
ProcImprimirRel FormulaRel_CFI, ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcRetirada()
On Error GoTo tratar_erro

If txtCodinterno.Text = "" Then Exit Sub
CFI_saida = True
CFI_devolucao = False
frmCFI_Saida.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

If Qualidade_Almox = False Then Formulario = "Estoque/Almoxarifado" Else Formulario = "Qualidade/Almoxarifado"
Direitos
ProcLimpaVariaveisPrincipais

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

With Lista
    If .ListItems.Count = 0 Then Exit Sub
    txtId.Text = .SelectedItem
    txtCodinterno.Text = .SelectedItem.SubItems(2)
    txtdescricao.Text = .SelectedItem.SubItems(3)
    txtObs.Text = .SelectedItem.SubItems(16)
    
    Set TBEstoque = CreateObject("adodb.recordset")
    TBEstoque.Open "Select Estoque_real from estoque_controle where IDestoque = " & .SelectedItem.ListSubItems(5), Conexao, adOpenKeyset, adLockOptimistic
    If TBEstoque.EOF = False Then
        txtquantestoque = IIf(IsNull(TBEstoque!estoque_real), 0, Format(TBEstoque!estoque_real, "###,##0.0000"))
    End If
    TBEstoque.Close
    
    txtfamilia.Text = .SelectedItem.SubItems(4)
End With
Frame2.Enabled = True
ProcCarregaListaMovimentacao
ProcCarregaLista_destino

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaListaMovimentacao()
On Error GoTo tratar_erro

'Carrega movimentações na lista
Lista_Movimentacao.ListItems.Clear
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from estoque_movimentacao where ID_CFI = " & Lista.SelectedItem & " order by Data desc, Idoperacao desc", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    PBLista1.Min = 0
    PBLista1.Max = TBAbrir.RecordCount
    PBLista1.Value = 1
    Contador = 0
    Do While TBAbrir.EOF = False
        With Lista_Movimentacao.ListItems
            .Add , , TBAbrir!IDoperacao
            .Item(.Count).SubItems(1) = IIf(IsNull(TBAbrir!Desenho), "", TBAbrir!Desenho)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBAbrir!Operacao), "", TBAbrir!Operacao)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBAbrir!Data), "", Format(TBAbrir!Data, "dd/mm/yy"))
            .Item(.Count).SubItems(4) = IIf(IsNull(TBAbrir!Entrada), "0,0000", Format(TBAbrir!Entrada, "###,##0.0000"))
            .Item(.Count).SubItems(5) = IIf(IsNull(TBAbrir!Saida), "0,0000", Format(TBAbrir!Saida, "###,##0.0000"))
            .Item(.Count).SubItems(6) = IIf(IsNull(TBAbrir!Documento), "", TBAbrir!Documento)
            .Item(.Count).SubItems(7) = IIf(IsNull(TBAbrir!Responsavel), "", TBAbrir!Responsavel)
            .Item(.Count).SubItems(8) = IIf(IsNull(TBAbrir!Requisitante), "", TBAbrir!Requisitante)
            .Item(.Count).SubItems(9) = TBAbrir!IDEstoque
        End With
        TBAbrir.MoveNext
        Contador = Contador + 1
        PBLista1.Value = Contador
    Loop
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaLista_destino()
On Error GoTo tratar_erro

Lista_destino.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "select * from CFI_Itens where ID_CFI = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista1.Min = 0
    PBLista1.Max = TBLISTA.RecordCount
    PBLista1.Value = 1
    Contador = 0
    Do While TBLISTA.EOF = False
        With Lista_destino.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Codproduto), "", TBLISTA!Codproduto)
            Set TBItem = CreateObject("adodb.recordset")
            TBItem.Open "Select * from Projproduto where codproduto = " & TBLISTA!Codproduto, Conexao, adOpenKeyset, adLockOptimistic
            If TBItem.EOF = False Then
                .Item(.Count).SubItems(2) = IIf(IsNull(TBItem!Desenho), "", TBItem!Desenho)
                .Item(.Count).SubItems(3) = IIf(IsNull(TBItem!Descricao), "", Trim(TBItem!Descricao))
                .Item(.Count).SubItems(4) = IIf(IsNull(TBItem!Classe), "", Trim(TBItem!Classe))
                .Item(.Count).SubItems(5) = IIf(IsNull(TBItem!Unidade), "", TBItem!Unidade)
            End If
            TBItem.Close
            TBLISTA.MoveNext
            Contador = Contador + 1
            PBLista1.Value = Contador
        End With
    Loop
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_movimentacao_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With Lista_Movimentacao
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            lista_v = .ListItems(InitFor).SubItems(2)
            If lista_v <> "SAIDA_ALMOXARIFADO" And lista_v <> "DEVOLUCAO_ALMOXARIFADO" And lista_v <> "DEVOLUCAO_ALMOXARIFADO C/ PROB." Then
                USMsgBox ("Não é permitido excluir este tipo de movimentação neste módulo."), vbExclamation, "CAPRIND v5.0"
                .ListItems.Item(InitFor).Checked = False
            End If
        End If
    Next InitFor
End With

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

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcNovo
    Case 2: ProcAbrir
    Case 3: ProcExcluir
    Case 4: ProcImprimir
    Case 5: procFiltrar_todos
    Case 6: ProcRetirada
    Case 7: ProcDevolucao
    Case 8: ProcAtualizar
    Case 10: ProcAjuda
    Case 11: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
