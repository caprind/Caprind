VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMinuta 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Administrativo - Faturamento - Minuta de despacho"
   ClientHeight    =   10035
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15360
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
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   75
      TabIndex        =   65
      Top             =   9720
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
   Begin TabDlg.SSTab SStab1 
      Height          =   10065
      Left            =   0
      TabIndex        =   36
      Top             =   0
      Width           =   15360
      _ExtentX        =   27093
      _ExtentY        =   17754
      _Version        =   393216
      Tab             =   2
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
      TabCaption(0)   =   "Minuta de despacho"
      TabPicture(0)   =   "frmMinuta.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "USToolBar1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Lista"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Notas fiscais"
      TabPicture(1)   =   "frmMinuta.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame7"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lista_nota"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "USToolBar2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Dados da transportadora"
      TabPicture(2)   =   "frmMinuta.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "USToolBar3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Lista_transportadora"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "txtID_transp"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Frame4"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Frame9"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).ControlCount=   5
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   -74925
         TabIndex        =   66
         Top             =   9090
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
            Left            =   3780
            TabIndex        =   6
            Text            =   "30"
            ToolTipText     =   "Número de registros por página."
            Top             =   180
            Width           =   555
         End
         Begin DrawSuite2022.USButton cmdPagProx 
            Height          =   315
            Left            =   11760
            TabIndex        =   11
            ToolTipText     =   "Próxima página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmMinuta.frx":0054
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
            TabIndex        =   10
            ToolTipText     =   "Página anterior."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmMinuta.frx":37FB
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
            TabIndex        =   8
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
            TabIndex        =   9
            ToolTipText     =   "Primeira página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmMinuta.frx":7313
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
            TabIndex        =   12
            ToolTipText     =   "Última página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmMinuta.frx":B409
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
         Begin VB.Label Label12 
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
            TabIndex        =   72
            Top             =   240
            Width           =   1440
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
            TabIndex        =   69
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
            TabIndex        =   68
            Top             =   240
            Width           =   1275
         End
         Begin VB.Label Label4 
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
            TabIndex        =   67
            Top             =   240
            Width           =   645
         End
      End
      Begin DrawSuite2022.USToolBar USToolBar1 
         Height          =   975
         Left            =   -74925
         TabIndex        =   64
         Top             =   330
         Width           =   15195
         _ExtentX        =   26802
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft2     =   40
         ButtonTop2      =   2
         ButtonWidth2    =   42
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
         ButtonLeft3     =   84
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
         ButtonLeft4     =   130
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
         ButtonLeft5     =   177
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
         ButtonLeft6     =   239
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
         ButtonLeft7     =   296
         ButtonTop7      =   2
         ButtonWidth7    =   55
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
         ButtonLeft8     =   353
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
         ButtonLeft9     =   357
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
         ButtonLeft10    =   400
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
         ButtonLeft11    =   432
         ButtonTop11     =   2
         ButtonWidth11   =   24
         ButtonHeight11  =   24
         ButtonUseMaskColor11=   0   'False
         Begin DrawSuite2022.USImageList USImageList1 
            Left            =   12330
            Top             =   120
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmMinuta.frx":ECC0
            Count           =   1
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   1485
         Left            =   75
         TabIndex        =   56
         Top             =   2820
         Width           =   15200
         Begin VB.TextBox txtFrete 
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
            Left            =   9975
            MaxLength       =   50
            TabIndex        =   34
            ToolTipText     =   "Frete."
            Top             =   975
            Width           =   5010
         End
         Begin VB.TextBox txtColeta 
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
            TabIndex        =   33
            ToolTipText     =   "Tipo de coleta."
            Top             =   975
            Width           =   9780
         End
         Begin VB.TextBox txtMotorista 
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
            Left            =   1470
            MaxLength       =   255
            TabIndex        =   31
            ToolTipText     =   "Motorista."
            Top             =   390
            Width           =   10590
         End
         Begin VB.TextBox txtPlaca 
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
            Left            =   12075
            MaxLength       =   50
            TabIndex        =   32
            ToolTipText     =   "Placa do veiculo."
            Top             =   390
            Width           =   2910
         End
         Begin VB.OptionButton optSim 
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
            Height          =   195
            Left            =   180
            TabIndex        =   29
            Top             =   450
            Width           =   555
         End
         Begin VB.OptionButton optNao 
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
            Height          =   195
            Left            =   750
            TabIndex        =   30
            Top             =   450
            Width           =   585
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Frete"
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
            Left            =   12285
            TabIndex        =   61
            Top             =   780
            Width           =   390
         End
         Begin VB.Label Label24 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de coleta"
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
            Left            =   4568
            TabIndex        =   60
            Top             =   780
            Width           =   1005
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Motorista"
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
            Left            =   6428
            TabIndex        =   59
            Top             =   180
            Width           =   675
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Placa do veículo"
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
            Left            =   12960
            TabIndex        =   58
            Top             =   180
            Width           =   1140
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Redespacho"
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
            Left            =   270
            TabIndex        =   57
            Top             =   180
            Width           =   1035
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0E0&
         Height          =   1515
         Left            =   75
         TabIndex        =   48
         Top             =   1305
         Width           =   15200
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
            ItemData        =   "frmMinuta.frx":14C0E
            Left            =   180
            List            =   "frmMinuta.frx":14C1E
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   19
            ToolTipText     =   "Tipo da transportadora."
            Top             =   390
            Width           =   1335
         End
         Begin VB.TextBox txtTranportadora 
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
            Left            =   1530
            MaxLength       =   255
            TabIndex        =   20
            ToolTipText     =   "Transportadora."
            Top             =   390
            Width           =   5910
         End
         Begin VB.CommandButton Cmd_cli_forn_transp 
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   7455
            Picture         =   "frmMinuta.frx":14C42
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Localizar cliente/fornecedor."
            Top             =   390
            Width           =   315
         End
         Begin VB.TextBox txtfax 
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
            Left            =   6330
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   26
            TabStop         =   0   'False
            ToolTipText     =   "Número do fax."
            Top             =   990
            Width           =   2280
         End
         Begin VB.TextBox txttelefone 
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
            Left            =   4050
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   25
            TabStop         =   0   'False
            ToolTipText     =   "Número do telefone."
            Top             =   990
            Width           =   2270
         End
         Begin VB.TextBox txtcnpj 
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
            Left            =   8625
            Locked          =   -1  'True
            MaxLength       =   200
            TabIndex        =   27
            TabStop         =   0   'False
            ToolTipText     =   "CNPJ/CPF."
            Top             =   990
            Width           =   3240
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
            Left            =   180
            Locked          =   -1  'True
            MaxLength       =   255
            TabIndex        =   23
            TabStop         =   0   'False
            ToolTipText     =   "Cidade."
            Top             =   990
            Width           =   3350
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
            Height          =   315
            Left            =   7875
            Locked          =   -1  'True
            MaxLength       =   255
            TabIndex        =   22
            TabStop         =   0   'False
            ToolTipText     =   "Endereço."
            Top             =   390
            Width           =   7125
         End
         Begin VB.TextBox txtIE 
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
            Locked          =   -1  'True
            MaxLength       =   200
            TabIndex        =   28
            TabStop         =   0   'False
            ToolTipText     =   "Inscrição estadual."
            Top             =   990
            Width           =   3120
         End
         Begin VB.TextBox cmbuf 
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
            Left            =   3540
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   24
            TabStop         =   0   'False
            ToolTipText     =   "Uf."
            Top             =   990
            Width           =   500
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo da transp."
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
            TabIndex        =   71
            Top             =   180
            Width           =   1095
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Transportadora"
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
            Left            =   3923
            TabIndex        =   35
            Top             =   180
            Width           =   1125
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "IE"
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
            Left            =   13365
            TabIndex        =   55
            Top             =   780
            Width           =   150
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
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
            Left            =   7335
            TabIndex        =   54
            Top             =   780
            Width           =   270
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
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
            Left            =   4875
            TabIndex        =   53
            Top             =   780
            Width           =   630
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "CNPJ/CPF"
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
            Left            =   9885
            TabIndex        =   52
            Top             =   780
            Width           =   720
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
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
            Left            =   1608
            TabIndex        =   51
            Top             =   780
            Width           =   495
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
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
            Left            =   11115
            TabIndex        =   50
            Top             =   180
            Width           =   675
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
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
            Height          =   195
            Index           =   5
            Left            =   3698
            TabIndex        =   49
            Top             =   780
            Width           =   195
         End
      End
      Begin VB.TextBox txtID_transp 
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
         Height          =   345
         Left            =   3255
         Locked          =   -1  'True
         MaxLength       =   50
         MouseIcon       =   "frmMinuta.frx":14D44
         MousePointer    =   99  'Custom
         TabIndex        =   47
         TabStop         =   0   'False
         Text            =   "0"
         ToolTipText     =   "Número do telefone."
         Top             =   6930
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   885
         Left            =   -74925
         TabIndex        =   41
         Top             =   8820
         Width           =   15200
         Begin VB.TextBox txtBruto 
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
            Left            =   8660
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   16
            TabStop         =   0   'False
            ToolTipText     =   "Peso bruto total."
            Top             =   420
            Width           =   2100
         End
         Begin VB.TextBox txtQtde 
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
            Left            =   180
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   14
            TabStop         =   0   'False
            ToolTipText     =   "Quantidade total."
            Top             =   420
            Width           =   2100
         End
         Begin VB.TextBox txtLiq 
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
            Left            =   4420
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   15
            TabStop         =   0   'False
            ToolTipText     =   "Peso líquido total."
            Top             =   420
            Width           =   2100
         End
         Begin VB.TextBox txtNF 
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
            Left            =   12900
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   17
            TabStop         =   0   'False
            ToolTipText     =   "Valor total da(s) nota(s) fiscal(ais)."
            Top             =   420
            Width           =   2100
         End
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Peso bruto total"
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
            Left            =   9028
            TabIndex        =   45
            Top             =   210
            Width           =   1365
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Quantidade total"
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
            Left            =   518
            TabIndex        =   44
            Top             =   210
            Width           =   1425
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Peso liquído total"
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
            Left            =   4743
            TabIndex        =   43
            Top             =   210
            Width           =   1455
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Valor total NF"
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
            Left            =   13388
            TabIndex        =   42
            Top             =   210
            Width           =   1125
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   4035
         Left            =   -74925
         TabIndex        =   37
         Top             =   1305
         Width           =   15200
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
            ItemData        =   "frmMinuta.frx":1504E
            Left            =   180
            List            =   "frmMinuta.frx":15050
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   0
            ToolTipText     =   "Empresa."
            Top             =   390
            Width           =   6990
         End
         Begin VB.TextBox txtobs 
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
            Height          =   2895
            Left            =   180
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   4
            ToolTipText     =   "Observações."
            Top             =   990
            Width           =   14820
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
            Left            =   9240
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   2
            TabStop         =   0   'False
            ToolTipText     =   "Data."
            Top             =   390
            Width           =   1110
         End
         Begin VB.TextBox txtResponsavel 
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
            Left            =   10365
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   3
            TabStop         =   0   'False
            ToolTipText     =   "Responsável."
            Top             =   390
            Width           =   4635
         End
         Begin VB.TextBox txtid 
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
            Left            =   7170
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   1
            TabStop         =   0   'False
            ToolTipText     =   "Número da minuta."
            Top             =   390
            Width           =   2055
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
            Left            =   3308
            TabIndex        =   70
            Top             =   180
            Width           =   735
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
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
            Left            =   7133
            TabIndex        =   46
            Top             =   780
            Width           =   945
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
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
            Left            =   9623
            TabIndex        =   40
            Top             =   180
            Width           =   345
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
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
            Left            =   12225
            TabIndex        =   39
            Top             =   180
            Width           =   915
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Número da minuta"
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
            Left            =   7417
            TabIndex        =   38
            Top             =   180
            Width           =   1560
         End
      End
      Begin MSComctlLib.ListView Lista 
         Height          =   3720
         Left            =   -74925
         TabIndex        =   5
         Top             =   5355
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   6562
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
            Text            =   "Nº minuta"
            Object.Width           =   3334
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
            Object.Width           =   20135
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Tag             =   "N"
            Text            =   "IDempresa"
            Object.Width           =   0
         EndProperty
      End
      Begin MSComctlLib.ListView lista_nota 
         Height          =   7470
         Left            =   -74925
         TabIndex        =   13
         Top             =   1335
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   13176
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
            Object.Tag             =   "N"
            Text            =   "Nota fiscal"
            Object.Width           =   2028
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Tipo"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Object.Tag             =   "N"
            Text            =   "Série"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "Cliente"
            Object.Width           =   3272
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Object.Tag             =   "T"
            Text            =   "Cidade"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Object.Tag             =   "T"
            Text            =   "Telefone"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Object.Tag             =   "N"
            Text            =   "Qtde."
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Espécie"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   9
            Object.Tag             =   "N"
            Text            =   "Peso liq."
            Object.Width           =   2028
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   10
            Object.Tag             =   "N"
            Text            =   "Peso bruto"
            Object.Width           =   2028
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   11
            Object.Tag             =   "N"
            Text            =   "Valor total"
            Object.Width           =   2117
         EndProperty
      End
      Begin MSComctlLib.ListView Lista_transportadora 
         Height          =   5385
         Left            =   75
         TabIndex        =   18
         Top             =   4320
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   9499
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
            Text            =   "ID_transp"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Tag             =   "T"
            Text            =   "Transportadora"
            Object.Width           =   10294
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Endereço"
            Object.Width           =   7937
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Tag             =   "N"
            Text            =   "CNPJ"
            Object.Width           =   3944
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Tag             =   "N"
            Text            =   "IE"
            Object.Width           =   3944
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Object.Tag             =   "T"
            Text            =   "Tipo"
            Object.Width           =   0
         EndProperty
      End
      Begin DrawSuite2022.USToolBar USToolBar3 
         Height          =   975
         Left            =   75
         TabIndex        =   62
         Top             =   330
         Width           =   15200
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
         ButtonUseMaskColor9=   0   'False
         Begin DrawSuite2022.USImageList USImageList3 
            Left            =   12330
            Top             =   180
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmMinuta.frx":15052
            Count           =   1
         End
      End
      Begin DrawSuite2022.USToolBar USToolBar2 
         Height          =   975
         Left            =   -74925
         TabIndex        =   63
         Top             =   330
         Width           =   15200
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
         ButtonLeft2     =   40
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
         ButtonLeft3     =   87
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
         ButtonLeft4     =   149
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
         ButtonLeft5     =   206
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
         ButtonLeft6     =   263
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
         ButtonLeft7     =   267
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
         ButtonLeft8     =   310
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
         ButtonLeft9     =   342
         ButtonTop9      =   2
         ButtonWidth9    =   24
         ButtonHeight9   =   24
         ButtonUseMaskColor9=   0   'False
         Begin DrawSuite2022.USImageList USImageList2 
            Left            =   12330
            Top             =   180
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmMinuta.frx":19B0D
            Count           =   1
         End
      End
   End
End
Attribute VB_Name = "frmMinuta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Novo_Minuta As Boolean 'OK
Public StrSql_Localizar_Minuta As String 'OK
Dim TBLISTA_Minuta As ADODB.Recordset 'OK

Sub ProcAjuda()
On Error GoTo tratar_erro

FunAbrirVideoWeb ("http://www.youtube.com/watch?v=Df-Fyi_w_Oo&list=UUODGiDjQ-BCrxh0YXfJvoqg&index=1&feature=plcp")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_tipo_transp_Click()
On Error GoTo tratar_erro

ProcCarregaTransportadora

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_cli_forn_transp_Click()
On Error GoTo tratar_erro

With Cmb_tipo_transp
    Acao = "localizar a transportadora"
    If .Text = "" Then
        NomeCampo = "o tipo da transportadora"
        ProcVerificaAcao
        .SetFocus
        Exit Sub
    End If

    Sit_REG = 4
    If .Text = "Cliente" Then
        ProcConfVariaveisLocCliente False, False, False, False, False, False, False, False, False, False, False, True, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False
        frmVendas_LocalizarCliente.Show 1
    ElseIf .Text = "Fornecedor" Then
            ProcConfVariaveisLocForn False, False, False, False, False, False, False, False, True, False, False, False, False, False, False, False, False, False, False, False
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
                If USMsgBox("Deseja realmente excluir esta(s) minuta(s)?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            Conexao.Execute "Update tbl_Detalhes_Nota Set ID_Minuta = 0 where ID_Minuta = " & .ListItems(InitFor)
            Conexao.Execute "DELETE from minuta WHERE id = " & .ListItems(InitFor)
            Conexao.Execute "DELETE from minuta_notas where id_minuta = " & .ListItems(InitFor)
            '==================================
            Modulo = "Faturamento/Minuta de despacho"
            Evento = "Excluir"
            ID_documento = .ListItems(InitFor)
            Documento = "Número da minuta: " & .ListItems(InitFor)
            Documento1 = ""
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe a(s) minuta(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Minuta(s) excluída(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCampos
    ProcCarregaLista (1)
    Frame3.Enabled = False
    Novo_Minuta = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluir_Nota()
On Error GoTo tratar_erro
    
If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With lista_nota
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir esta(s) nota(s)?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            Conexao.Execute "Update tbl_Detalhes_Nota set tbl_Detalhes_Nota.ID_Minuta = 0 from tbl_Detalhes_Nota INNER JOIN Minuta_notas on tbl_Detalhes_Nota.ID_nota = Minuta_notas.ID_nota where Minuta_notas.ID = " & .ListItems(InitFor)
            Conexao.Execute "DELETE from minuta_notas where ID = " & .ListItems(InitFor)
            '==================================
            Modulo = "Faturamento/Minuta de despacho"
            Evento = "Excluir nota"
            ID_documento = .ListItems(InitFor)
            Documento = "Número da minuta: " & txtId
            Documento1 = "Nº nota: " & .ListItems(InitFor).ListSubItems(1) & " - Tipo: " & .ListItems(InitFor).ListSubItems(2) & " - Série: " & .ListItems(InitFor).ListSubItems(3)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe a(s) nota(s) fiscal(ais) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Nota(s) fiscal(ais) excluída(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcCarregalista_nota
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro

If txtId = "" Then
    USMsgBox ("Informe a minuta antes de visualizar impressão."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
NomeRel = "Faturamento_minuta.rpt"
ProcImprimirRel "{Minuta.id}= " & txtId, ""

NomeRel = "Faturamento_romaneio.rpt"
ProcImprimirRel "{Minuta.id}= " & txtId, ""

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
Novo_Minuta = True
Frame3.Enabled = True
Cmb_empresa.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovo_Nota()
On Error GoTo tratar_erro
  
If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Minuta = True
Faturamento = False
frmMinuta_notas.Show 1

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
If Frame3.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * FROM minuta where id = " & IIf(txtId = "", 0, txtId), Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then
    TBGravar.AddNew
    TBGravar!Redespacho = False
End If
TBGravar!ID_empresa = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
If txtResponsavel = "" Then TBGravar!Responsavel = pubUsuario Else TBGravar!Responsavel = txtResponsavel
If txtData = "" Then TBGravar!Data = Date Else TBGravar!Data = txtData
TBGravar!Observacao = txtObs
TBGravar.Update
txtId = TBGravar!ID
TBGravar.Close
If Novo_Minuta = True Then
    USMsgBox ("Nova minuta cadastrada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo"
    StrSql_Localizar_Minuta = "Select * FROM minuta where id = " & txtId
    ProcCarregaLista (1)
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar"
    ProcCarregaLista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
    If Lista.ListItems.Count <> 0 And CodigoLista <> 0 Then
        Lista.SelectedItem = Lista.ListItems(CodigoLista)
        Lista.SetFocus
    End If
End If
'==================================
Modulo = "Faturamento/Minuta de despacho"
ID_documento = txtId
Documento = "Número da minuta: " & txtId
Documento1 = ""
ProcGravaEvento
'==================================
Novo_Minuta = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Minuta.AbsolutePage <> 2 Then
    If TBLISTA_Minuta.AbsolutePage = -3 Then
        ProcExibePagina (TBLISTA_Minuta.PageCount - 1)
    Else
        TBLISTA_Minuta.AbsolutePage = TBLISTA_Minuta.AbsolutePage - 2
        ProcExibePagina (TBLISTA_Minuta.AbsolutePage)
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
    TBLISTA_Minuta.AbsolutePage = txtPagIr.Text
    ProcExibePagina (TBLISTA_Minuta.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Minuta.AbsolutePage = 1
ProcExibePagina (TBLISTA_Minuta.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Minuta.AbsolutePage <> -3 Then
    If TBLISTA_Minuta.AbsolutePage = 1 Then
        ProcExibePagina (2)
    Else
        ProcExibePagina (TBLISTA_Minuta.AbsolutePage)
    End If
Else
    ProcExibePagina (TBLISTA_Minuta.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Minuta.AbsolutePage = TBLISTA_Minuta.PageCount
ProcExibePagina (TBLISTA_Minuta.AbsolutePage)

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
            Case vbKeyF2: ProcFiltrar
            Case vbKeyF3: ProcSalvar
            Case vbKeyF4: ProcExcluir
            Case vbKeyF5: ProcImprimir
            Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
    Case 1:
        Select Case KeyCode
            Case vbKeyInsert: ProcNovo_Nota
            Case vbKeyF4: ProcExcluir_Nota
            Case vbKeyF5: ProcImprimir
            Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
    Case 2:
        Select Case KeyCode
            Case vbKeyF3: ProcSalvar_Transp
            Case vbKeyF4: ProcExcluir_Transp
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

ProcCarregaToolBar1 Me, 15200, 11, True
ProcCarregaToolBar2 Me, 15200, 9, True
ProcCarregaToolBar3 Me, 15200, 9, True
Formulario = "Faturamento/Minuta de despacho"
Direitos
SSTab1.Tab = 0
ProcCarregaComboEmpresa Cmb_empresa, False
ProcLimpaVariaveisPrincipais

ProcRemoveObjetosResize Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro
    
Formulario = "Faturamento/Minuta de despacho"
Direitos
ProcLimpaVariaveisPrincipais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluir_Transp()
On Error GoTo tratar_erro
    
If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If txtID_transp.Text = "" Then
    USMsgBox ("Informe a transportadora antes de excluir."), vbInformation, "CAPRIND v5.0"
    Exit Sub
End If
If USMsgBox("Deseja realmente excluir esta transportadora?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    Conexao.Execute "Update minuta Set ID_transportadora = 0, Motorista = Null, placa = Null, Coleta = Null, Frete = Null where id = " & txtId
    USMsgBox ("Transportadora excluída com sucesso."), vbInformation, "CAPRIND v5.0"
    '==================================
    Modulo = "Faturamento/Minuta de despacho"
    Evento = "Excluir transportadora"
    ID_documento = txtID_transp
    Documento = "Número da minuta: " & txtId
    Documento1 = "Transportadora: " & txtTranportadora
    ProcGravaEvento
    '==================================
    ProcLimpacampos_transp
    txtTranportadora = ""
    Frame9.Enabled = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub



Private Sub ProcFiltrar()
On Error GoTo tratar_erro

frmMinuta_Localizar.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

If Novo_Minuta = True Then
    If USMsgBox("A minuta ainda não foi salva, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvar
        If Novo_Minuta = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
Novo_Minuta = False
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvar_Transp()
On Error GoTo tratar_erro
  
If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If txtID_transp.Text = "" Then
    USMsgBox ("Informe a transportadora antes de salvar."), vbInformation, "CAPRIND v5.0"
    Exit Sub
End If
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * FROM minuta where id = " & txtId.Text, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = False Then
    Select Case Cmb_tipo_transp
        Case "Cliente": TipoTransp = "C"
        Case "Fornecedor": TipoTransp = "F"
        Case "Empresa": TipoTransp = "E"
    End Select
    TBGravar!Tipo_transp = TipoTransp
    TBGravar!ID_transportadora = txtID_transp
    If optSim.Value = True Then TBGravar!Redespacho = True Else TBGravar!Redespacho = False
    TBGravar!Motorista = txtMotorista
    TBGravar!placa = txtPlaca
    TBGravar!Coleta = txtColeta
    TBGravar!Frete = txtFrete
    TBGravar.Update
    TBGravar.Close
    USMsgBox ("Transportadora cadastrada com sucesso."), vbInformation, "CAPRIND v5.0"
    '==================================
    Modulo = "Faturamento/Minuta de despacho"
    Evento = "Nova transportadora"
    ID_documento = txtID_transp
    Documento = "Número da minuta: " & txtId
    Documento1 = "Transportadora: " & txtTranportadora
    ProcGravaEvento
    '==================================
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "Nº minuta" Then
    With Lista
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then .ListItems.Item(InitFor).Checked = False Else .ListItems.Item(InitFor).Checked = True
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
 
Private Sub Lista_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from minuta where id = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    ProcLimpaCampos
    ProcCarregaDados
    CodigoLista = Lista.SelectedItem.index
End If
TBLISTA.Close
Frame3.Enabled = True
Novo_Minuta = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub lista_nota_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With lista_nota
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then .ListItems.Item(InitFor).Checked = False Else .ListItems.Item(InitFor).Checked = True
        Next InitFor
    End With
Else
    ProcOrdenaListView lista_nota, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_transportadora_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView Lista_transportadora, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_transportadora_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista_transportadora.ListItems.Count = 0 Then Exit Sub
txtID_transp = Lista_transportadora.SelectedItem
If Cmb_tipo_transp = "Empresa" Then
    NomeTabela = "Empresa"
    NomeCampo = "Codigo"
Else
    NomeCampo = "IDCliente"
    If Cmb_tipo_transp = "Cliente" Then NomeTabela = "Clientes" Else NomeTabela = "Compras_fornecedores"
End If
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * FROM " & NomeTabela & " WHERE " & NomeCampo & " = " & txtID_transp, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    If Cmb_tipo_transp = "Empresa" Then
        txtTranportadora = IIf(IsNull(TBAbrir!Empresa), "", TBAbrir!Empresa)
        If IsNull(TBAbrir!CNPJ) = True Or TBAbrir!CNPJ = "__.___.___/____-__" Or TBAbrir!CNPJ = "" Then txtcnpj = "" Else txtcnpj = TBAbrir!CNPJ
        txtIE = IIf(IsNull(TBAbrir!IE), "", TBAbrir!IE)
        cmbuf = IIf(IsNull(TBAbrir!UF), "", TBAbrir!UF)
        txttelefone = IIf(IsNull(TBAbrir!telefone), "", TBAbrir!telefone)
    Else
        If TBAbrir!idTipoEmpresa = 1 Then
            If IsNull(TBAbrir!CPF_CNPJ) = True Or TBAbrir!CPF_CNPJ = "__.___.___/____-__" Or TBAbrir!CPF_CNPJ = "" Then txtcnpj = "" Else txtcnpj = TBAbrir!CPF_CNPJ
        End If
        txtIE = IIf(IsNull(TBAbrir!RG_IE), "", TBAbrir!RG_IE)
        If Cmb_tipo_transp = "Cliente" Then
            txtTranportadora = IIf(IsNull(TBAbrir!NomeRazao), "", TBAbrir!NomeRazao)
            cmbuf = IIf(IsNull(TBAbrir!UF), "", TBAbrir!UF)
            txttelefone = IIf(IsNull(TBAbrir!Tel01), "", TBAbrir!Tel01)
        Else
            txtTranportadora = IIf(IsNull(TBAbrir!Nome_Razao), "", TBAbrir!Nome_Razao)
            cmbuf = IIf(IsNull(TBAbrir!Estado), "", TBAbrir!Estado)
            txttelefone = IIf(IsNull(TBAbrir!Telefones), "", TBAbrir!Telefones)
        End If
    End If
1:
    If IsNull(TBAbrir!Tipo_endereco) = False And TBAbrir!Tipo_endereco <> "" Then
        Endereco = TBAbrir!Tipo_endereco & ": " & IIf(IsNull(TBAbrir!Endereco), "", TBAbrir!Endereco)
    Else
        Endereco = IIf(IsNull(TBAbrir!Endereco), "", TBAbrir!Endereco)
    End If
    txtendereco = Endereco
    txtCidade = IIf(IsNull(TBAbrir!Cidade), "", TBAbrir!Cidade)
    txtFax = IIf(IsNull(TBAbrir!Fax), "", TBAbrir!Fax)
    
    Frame9.Enabled = True
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaDadosTransporte()
On Error GoTo tratar_erro

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from minuta where id = " & txtId & " and ID_transportadora <> 0 and ID_transportadora is not null", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    txtID_transp = TBAbrir!ID_transportadora
    If TBAbrir!Tipo_transp = "E" Then
        NomeTabela = "Empresa"
        NomeCampo = "Codigo"
    Else
        NomeCampo = "IDCliente"
        If TBAbrir!Tipo_transp = "C" Then NomeTabela = "Clientes" Else NomeTabela = "Compras_fornecedores"
    End If
    Set TBFI = CreateObject("adodb.recordset")
    TBFI.Open "Select * from " & NomeTabela & " where " & NomeCampo & " = " & TBAbrir!ID_transportadora, Conexao, adOpenKeyset, adLockOptimistic
    If TBFI.EOF = False Then
        Select Case TBAbrir!Tipo_transp
            Case "C":
                Cmb_tipo_transp = "Cliente"
                txtTranportadora = TBFI!NomeRazao
                If IsNull(TBFI!Tipo_endereco) = False And TBFI!Tipo_endereco <> "" Then
                    Endereco = TBFI!Tipo_endereco & ": " & IIf(IsNull(TBFI!Endereco), "", TBFI!Endereco)
                Else
                    Endereco = IIf(IsNull(TBFI!Endereco), "", TBFI!Endereco)
                End If
                txtendereco = Endereco
                txtCidade = IIf(IsNull(TBFI!Cidade), "", TBFI!Cidade)
                cmbuf = IIf(IsNull(TBFI!UF), "", TBFI!UF)
                txttelefone = IIf(IsNull(TBFI!Tel01), "", TBFI!Tel01)
                txtFax = IIf(IsNull(TBFI!Fax), "", TBFI!Fax)
                If TBFI!idTipoEmpresa = 1 Then
                    If IsNull(TBFI!CPF_CNPJ) = True Or TBFI!CPF_CNPJ = "__.___.___/____-__" Or TBFI!CPF_CNPJ = "" Then txtcnpj = "" Else txtcnpj = TBFI!CPF_CNPJ
                End If
                txtIE = IIf(IsNull(TBFI!RG_IE), "", TBFI!RG_IE)
            Case "F":
                Cmb_tipo_transp = "Fornecedor"
                txtTranportadora = TBFI!Nome_Razao
                If IsNull(TBFI!Tipo_endereco) = False And TBFI!Tipo_endereco <> "" Then
                    Endereco = TBFI!Tipo_endereco & ": " & IIf(IsNull(TBFI!Endereco), "", TBFI!Endereco)
                Else
                    Endereco = IIf(IsNull(TBFI!Endereco), "", TBFI!Endereco)
                End If
                txtendereco = Endereco
                txtCidade = IIf(IsNull(TBFI!Cidade), "", TBFI!Cidade)
                cmbuf = IIf(IsNull(TBFI!Estado), "", TBFI!Estado)
                txttelefone = IIf(IsNull(TBFI!Telefones), "", TBFI!Telefones)
                txtFax = IIf(IsNull(TBFI!Fax), "", TBFI!Fax)
                If TBFI!idTipoEmpresa = 1 Then
                    If IsNull(TBFI!CPF_CNPJ) = True Or TBFI!CPF_CNPJ = "__.___.___/____-__" Or TBFI!CPF_CNPJ = "" Then txtcnpj = "" Else txtcnpj = TBFI!CPF_CNPJ
                End If
                txtIE = IIf(IsNull(TBFI!RG_IE), "", TBFI!RG_IE)
            Case "E":
                Cmb_tipo_transp = "Empresa"
                txtTranportadora = TBFI!Empresa
                If IsNull(TBFI!Tipo_endereco) = False And TBFI!Tipo_endereco <> "" Then
                    Endereco = TBFI!Tipo_endereco & ": " & IIf(IsNull(TBFI!Endereco), "", TBFI!Endereco)
                Else
                    Endereco = IIf(IsNull(TBFI!Endereco), "", TBFI!Endereco)
                End If
                txtendereco = Endereco
                txtCidade = IIf(IsNull(TBFI!Cidade), "", TBFI!Cidade)
                cmbuf = IIf(IsNull(TBFI!UF), "", TBFI!UF)
                txttelefone = IIf(IsNull(TBFI!telefone), "", TBFI!telefone)
                txtFax = IIf(IsNull(TBFI!Fax), "", TBFI!Fax)
                If IsNull(TBFI!CNPJ) = True Or TBFI!CNPJ = "__.___.___/____-__" Or TBFI!CNPJ = "" Then txtcnpj = "" Else txtcnpj = TBFI!CNPJ
                txtIE = IIf(IsNull(TBFI!IE), "", TBFI!IE)
        End Select
    End If
    TBFI.Close
    If TBAbrir!Redespacho = True Then optSim.Value = True Else optNao.Value = True
    txtMotorista = IIf(IsNull(TBAbrir!Motorista), "", TBAbrir!Motorista)
    txtPlaca = IIf(IsNull(TBAbrir!placa), "", TBAbrir!placa)
    txtColeta = IIf(IsNull(TBAbrir!Coleta), "", TBAbrir!Coleta)
    txtFrete = IIf(IsNull(TBAbrir!Frete), "", TBAbrir!Frete)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error GoTo tratar_erro

If txtId = "" Then
    SSTab1.Tab = 0
    Exit Sub
End If
Select Case SSTab1.Tab
    Case 0: Lista.SetFocus
    Case 1:
        ProcVerificaProsseguir
        If Permitido = False Then Exit Sub
        lista_nota.SetFocus
        ProcCarregalista_nota
    Case 2:
        ProcVerificaProsseguir
        If Permitido = False Then Exit Sub
        Lista_transportadora.SetFocus
        ProcLimpacampos_transp
        ProcCarregaDadosTransporte
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcVerificaProsseguir()
On Error GoTo tratar_erro

Permitido = True
If Novo_Minuta = True Then
    USMsgBox ("Salve a minuta antes de prosseguir."), vbInformation, "CAPRIND v5.0"
    Permitido = False
    SSTab1.Tab = 0
    Exit Sub
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCampos()
On Error GoTo tratar_erro

txtId = ""
txtResponsavel = pubUsuario
txtData = Format(Date, "dd/mm/yy")
txtObs = ""
CodigoLista = 0
Caption = "Administrativo - Faturamento - Minuta de despacho"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaLista(Pagina As Integer)
On Error GoTo tratar_erro

Lista.ListItems.Clear
lblRegistros.Caption = "Nº de registros: 0"
lblPaginas.Caption = "Página: 0 de: 0"
If StrSql_Localizar_Minuta = "" Then Exit Sub
Set TBLISTA_Minuta = CreateObject("adodb.recordset")
TBLISTA_Minuta.Open StrSql_Localizar_Minuta, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA_Minuta.EOF = False Then ProcExibePagina (Pagina)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina(Pagina)
On Error GoTo tratar_erro

Lista.ListItems.Clear
TBLISTA_Minuta.PageSize = IIf(txtNreg = "", 30, txtNreg)
TBLISTA_Minuta.AbsolutePage = Pagina
TamanhoPagina = TBLISTA_Minuta.PageSize
ContadorReg = 1

PBLista.Min = 0
PBLista.Max = FunVerifMaxPBListaPaginacao(TBLISTA_Minuta.RecordCount - IIf(Pagina > 1, (TBLISTA_Minuta.PageSize * (Pagina - 1)), 0), TBLISTA_Minuta.PageSize)
PBLista.Value = 1
Contador = 0
Do While TBLISTA_Minuta.EOF = False And (ContadorReg <= TamanhoPagina)
    With Lista.ListItems
        .Add , , TBLISTA_Minuta!ID
        .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA_Minuta!Data), "", Format(TBLISTA_Minuta!Data, "dd/mm/yy"))
        .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA_Minuta!Responsavel), "", TBLISTA_Minuta!Responsavel)
        .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA_Minuta!ID_empresa), 0, TBLISTA_Minuta!ID_empresa)
    End With
    TBLISTA_Minuta.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista.Value = Contador
Loop
lblRegistros.Caption = "Nº de registros: " & TBLISTA_Minuta.RecordCount
If TBLISTA_Minuta.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "Página: 1 de: " & TBLISTA_Minuta.PageCount
ElseIf TBLISTA_Minuta.AbsolutePage = adPosEOF Then
        lblPaginas.Caption = "Página: " & TBLISTA_Minuta.PageCount & " de: " & TBLISTA_Minuta.PageCount
    Else
        lblPaginas.Caption = "Página: " & TBLISTA_Minuta.AbsolutePage - 1 & " de: " & TBLISTA_Minuta.PageCount
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregalista_nota()
On Error GoTo tratar_erro

Quant = 0
PesoBruto = 0
PesoLiquido = 0
ValorTotal = 0
lista_nota.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from minuta_notas where id_minuta = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    TBLISTA.MoveLast
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    TBLISTA.MoveFirst
    Do While TBLISTA.EOF = False
        With lista_nota.ListItems
            .Add , , TBLISTA!ID
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select * from tbl_Dados_Nota_Fiscal where id = " & TBLISTA!ID_nota, Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                .Item(.Count).SubItems(1) = IIf(IsNull(TBAbrir!int_NotaFiscal), "", TBAbrir!int_NotaFiscal)
                If IsNull(TBAbrir!TipoNF) = False Then
                    If TBAbrir!TipoNF = "M1" Then Tipo = "Produto(s)"
                    If TBAbrir!TipoNF = "SA" Then Tipo = "Serviço(s)"
                    If TBAbrir!TipoNF = "M1SA" Then Tipo = "Prod./Serv."
                End If
                .Item(.Count).SubItems(2) = IIf(Tipo = "", "", Tipo)
                .Item(.Count).SubItems(3) = IIf(IsNull(TBAbrir!Serie), "", TBAbrir!Serie)
                .Item(.Count).SubItems(4) = IIf(IsNull(TBAbrir!txt_Razao_Nome), "", TBAbrir!txt_Razao_Nome)
                .Item(.Count).SubItems(5) = IIf(IsNull(TBAbrir!txt_Municipio), "", TBAbrir!txt_Municipio)
                .Item(.Count).SubItems(6) = IIf(IsNull(TBAbrir!txt_Fone_Fax), "", TBAbrir!txt_Fone_Fax)
            End If
            TBAbrir.Close
            Set TBTransporte = CreateObject("adodb.recordset")
            TBTransporte.Open "Select * from tbl_Dados_Transp where id_nota = " & TBLISTA!ID_nota, Conexao, adOpenKeyset, adLockOptimistic
            If TBTransporte.EOF = False Then
                .Item(.Count).SubItems(7) = IIf(IsNull(TBTransporte!int_Qtd_Transp), "", Format(TBTransporte!int_Qtd_Transp, "###,##0.0000"))
                .Item(.Count).SubItems(8) = IIf(IsNull(TBTransporte!txt_Especie), "", TBTransporte!txt_Especie)
                .Item(.Count).SubItems(9) = IIf(IsNull(TBTransporte!dbl_Peso_Liquido), "", Format(TBTransporte!dbl_Peso_Liquido, "###,##0.0000"))
                .Item(.Count).SubItems(10) = IIf(IsNull(TBTransporte!dbl_Peso_Bruto), "", Format(TBTransporte!dbl_Peso_Bruto, "###,##0.0000"))
                Quant = Quant + IIf(IsNull(TBTransporte!int_Qtd_Transp), 0, Format(TBTransporte!int_Qtd_Transp, "###,##0.0000"))
                PesoLiquido = PesoLiquido + IIf(IsNull(TBTransporte!dbl_Peso_Liquido), 0, Format(TBTransporte!dbl_Peso_Liquido, "###,##0.0000"))
                PesoBruto = PesoBruto + IIf(IsNull(TBTransporte!dbl_Peso_Bruto), 0, Format(TBTransporte!dbl_Peso_Bruto, "###,##0.0000"))
            End If
            TBTransporte.Close
            Set TBTotaisnota = CreateObject("adodb.recordset")
            TBTotaisnota.Open "Select dbl_Valor_Total_Nota from tbl_Totais_Nota where id_nota = " & TBLISTA!ID_nota, Conexao, adOpenKeyset, adLockOptimistic
            If TBTotaisnota.EOF = False Then
                .Item(.Count).SubItems(11) = IIf(IsNull(TBTotaisnota!dbl_Valor_Total_Nota), 0, Format(TBTotaisnota!dbl_Valor_Total_Nota, "###,##0.00"))
                ValorTotal = ValorTotal + IIf(IsNull(TBTotaisnota!dbl_Valor_Total_Nota), 0, TBTotaisnota!dbl_Valor_Total_Nota)
            End If
            TBTotaisnota.Close
        End With
        TBLISTA.MoveNext
        Contador = Contador + 1
        PBLista.Value = Contador
    Loop
End If
TBLISTA.Close
ProcLimpacampos_nota
txtQtde = Format(Quant, "###,##0.0000")
txtLiq = Format(PesoLiquido, "###,##0.0000")
txtBruto = Format(PesoBruto, "###,##0.0000")
txtNF = Format(ValorTotal, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpacampos_nota()
On Error GoTo tratar_erro

txtQtde = ""
txtBruto = ""
txtLiq = ""
txtNF = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAnterior()
On Error GoTo tratar_erro

If txtId = "" Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from minuta order by id", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.BOF = False Then
    TBLISTA.Find ("id = " & txtId)
    TBLISTA.MovePrevious
    If TBLISTA.BOF = False Then
        ProcLimpaCampos
        ProcLimpacampos_nota
        ProcLimpacampos_transp
        txtTranportadora = ""
        ProcCarregaDados
        ProcCarregalista_nota
        ProcCarregaDadosTransporte
    Else
        USMsgBox ("Fim dos cadastros da minuta."), vbInformation, "CAPRIND v5.0"
    End If
End If
Novo_Minuta = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcProximo()
On Error GoTo tratar_erro

If txtId = "" Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from minuta order by id", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.BOF = False Then
    TBLISTA.Find ("id = " & txtId)
    TBLISTA.MoveNext
    If TBLISTA.EOF = False Then
        ProcLimpaCampos
        ProcLimpacampos_nota
        ProcLimpacampos_transp
        txtTranportadora = ""
        ProcCarregaDados
        ProcCarregalista_nota
        ProcCarregaDadosTransporte
    Else
        USMsgBox ("Fim dos cadastros da minuta."), vbInformation, "CAPRIND v5.0"
    End If
End If
Novo_Minuta = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaDados()
On Error GoTo tratar_erro

txtId.Text = TBLISTA!ID
If IsNull(TBLISTA!ID_empresa) = False And TBLISTA!ID_empresa <> "" Then ProcPuxaDadosComboEmpresa Cmb_empresa, TBLISTA!ID_empresa
txtData = Format(TBLISTA!Data, "dd/mm/yy")
txtResponsavel = TBLISTA!Responsavel
txtObs = IIf(IsNull(TBLISTA!Observacao), "", TBLISTA!Observacao)
Caption = "Administrativo - Faturamento - Minuta de despacho (Número : " & TBLISTA!ID & ")"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpacampos_transp()
On Error GoTo tratar_erro

txtID_transp = ""
Cmb_tipo_transp.ListIndex = -1
txtTranportadora = ""
txtendereco = ""
txtBairro = ""
txtCidade = ""
cmbuf = ""
txttelefone = ""
txtFax = ""
txtcnpj = ""
txtIE = ""
optSim.Value = False
optNao.Value = True
txtMotorista = ""
txtPlaca = ""
txtColeta = ""
txtFrete = ""

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

Private Sub txtTranportadora_Change()
On Error GoTo tratar_erro

ProcCarregaTransportadora

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaTransportadora()
On Error GoTo tratar_erro

txtendereco.Text = ""
txtCidade.Text = ""
cmbuf.Text = ""
txttelefone.Text = ""
txtFax.Text = ""
txtcnpj.Text = ""
txtIE.Text = ""

TextoFiltro = " and DtValidacao IS NOT NULL and status <> 'Bloqueado'"
Lista_transportadora.ListItems.Clear
If Cmb_tipo_transp = "Cliente" Then
    NomeTabela = "Clientes"
    NomeCampo = "NomeRazao"
ElseIf Cmb_tipo_transp = "Fornecedor" Then
        NomeTabela = "Compras_fornecedores"
        NomeCampo = "Nome_Razao"
    Else
        NomeTabela = "Empresa"
        NomeCampo = "Empresa"
        TextoFiltro = ""
End If

Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * FROM " & NomeTabela & " where " & NomeCampo & " like '" & txtTranportadora & "%' " & TextoFiltro & " order by " & NomeCampo, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    With Lista_transportadora.ListItems
        Do While TBLISTA.EOF = False
            Select Case Cmb_tipo_transp
                Case "Cliente":
                    .Add , , TBLISTA!IDCliente
                    .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!NomeRazao), "", TBLISTA!NomeRazao)
                    If IsNull(TBLISTA!Tipo_endereco) = False And TBLISTA!Tipo_endereco <> "" Then
                        Endereco = TBLISTA!Tipo_endereco & ": " & IIf(IsNull(TBLISTA!Endereco), "", TBLISTA!Endereco)
                    Else
                        Endereco = IIf(IsNull(TBLISTA!Endereco), "", TBLISTA!Endereco)
                    End If
                    .Item(.Count).SubItems(2) = Endereco
                    If TBLISTA!idTipoEmpresa = 1 Then
                        If IsNull(TBLISTA!CPF_CNPJ) = True Or TBLISTA!CPF_CNPJ = "__.___.___/____-__" Or TBLISTA!CPF_CNPJ = "" Then
                            .Item(.Count).SubItems(3) = ""
                        Else
                            .Item(.Count).SubItems(3) = TBLISTA!CPF_CNPJ
                        End If
                    End If
                    .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!RG_IE), "", TBLISTA!RG_IE)
                    .Item(.Count).SubItems(5) = "C"
                Case "Fornecedor":
                    .Add , , TBLISTA!IDCliente
                    .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Nome_Razao), "", TBLISTA!Nome_Razao)
                    If IsNull(TBLISTA!Tipo_endereco) = False And TBLISTA!Tipo_endereco <> "" Then
                        Endereco = TBLISTA!Tipo_endereco & ": " & IIf(IsNull(TBLISTA!Endereco), "", TBLISTA!Endereco)
                    Else
                        Endereco = IIf(IsNull(TBLISTA!Endereco), "", TBLISTA!Endereco)
                    End If
                    .Item(.Count).SubItems(2) = Endereco
                    If TBLISTA!idTipoEmpresa = 1 Then
                        If IsNull(TBLISTA!CPF_CNPJ) = True Or TBLISTA!CPF_CNPJ = "__.___.___/____-__" Or TBLISTA!CPF_CNPJ = "" Then
                            .Item(.Count).SubItems(3) = ""
                        Else
                            .Item(.Count).SubItems(3) = TBLISTA!CPF_CNPJ
                        End If
                    End If
                    .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!RG_IE), "", TBLISTA!RG_IE)
                    .Item(.Count).SubItems(5) = "F"
                Case "Empresa":
                    .Add , , TBLISTA!CODIGO
                    .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Empresa), "", TBLISTA!Empresa)
                    If IsNull(TBLISTA!Tipo_endereco) = False And TBLISTA!Tipo_endereco <> "" Then
                        Endereco = TBLISTA!Tipo_endereco & ": " & IIf(IsNull(TBLISTA!Endereco), "", TBLISTA!Endereco)
                    Else
                        Endereco = IIf(IsNull(TBLISTA!Endereco), "", TBLISTA!Endereco)
                    End If
                    .Item(.Count).SubItems(2) = Endereco
                    If IsNull(TBLISTA!CNPJ) = True Or TBLISTA!CNPJ = "__.___.___/____-__" Or TBLISTA!CNPJ = "" Then
                        .Item(.Count).SubItems(3) = ""
                    Else
                        .Item(.Count).SubItems(3) = TBLISTA!CNPJ
                    End If
                    .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!IE), "", TBLISTA!IE)
                    .Item(.Count).SubItems(5) = "E"
            End Select
            TBLISTA.MoveNext
        Loop
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
    Case 1: ProcNovo
    Case 2: ProcFiltrar
    Case 3: ProcSalvar
    Case 4: ProcExcluir
    Case 5: ProcImprimir
    Case 6: ProcAnterior
    Case 7: ProcProximo
    Case 9: ProcAjuda
    Case 10: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar2_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcNovo_Nota
    Case 2: ProcExcluir_Nota
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

Private Sub USToolBar3_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcSalvar_Transp
    Case 2: ProcExcluir_Transp
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

