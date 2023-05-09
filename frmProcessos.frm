VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{50D37AD9-8D3C-43DD-ADD5-7C957C951843}#1.9#0"; "FlexCell.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProcessos 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Engenharia - Processos"
   ClientHeight    =   10035
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15360
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProcessos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10035
   ScaleWidth      =   15360
   WindowState     =   2  'Maximized
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   10035
      Left            =   0
      TabIndex        =   57
      Top             =   0
      Width           =   15360
      _ExtentX        =   27093
      _ExtentY        =   17701
      _Version        =   393216
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
      TabCaption(0)   =   "Dados principais"
      TabPicture(0)   =   "frmProcessos.frx":014A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame17"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "PBLista"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "ListaProcessos"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "USToolBar1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Estrutura do produto"
      TabPicture(1)   =   "frmProcessos.frx":0166
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(1)=   "Grid1"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Fases"
      TabPicture(2)   =   "frmProcessos.frx":0182
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Framedetalhes"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame14"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "USToolBar2"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "PBlista1"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "ListaFases"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Frame11"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "CommonDialog1"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Txt_ID_fase"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Frame5"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).ControlCount=   9
      Begin VB.Frame Frame5 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   -74925
         TabIndex        =   155
         Top             =   9090
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
            ItemData        =   "frmProcessos.frx":019E
            Left            =   6960
            List            =   "frmProcessos.frx":01A8
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   56
            TabStop         =   0   'False
            Top             =   180
            Width           =   1965
         End
         Begin VB.Label Label1 
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
            Index           =   34
            Left            =   5610
            TabIndex        =   156
            Top             =   240
            Width           =   1260
         End
      End
      Begin VB.TextBox Txt_ID_fase 
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
         Height          =   325
         Left            =   -72330
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   144
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   6360
         Visible         =   0   'False
         Width           =   615
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   -67140
         Top             =   6030
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
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
         Height          =   675
         Left            =   -74925
         TabIndex        =   138
         Top             =   330
         Width           =   15195
         Begin VB.ComboBox cmbVersao_pesquisar_estrutura 
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
            ItemData        =   "frmProcessos.frx":01BD
            Left            =   2070
            List            =   "frmProcessos.frx":020F
            Style           =   2  'Dropdown List
            TabIndex        =   22
            ToolTipText     =   "Versão."
            Top             =   210
            Width           =   795
         End
         Begin VB.Image imgFile 
            Height          =   240
            Left            =   14550
            Picture         =   "frmProcessos.frx":0261
            Top             =   270
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image imgFolder 
            Height          =   240
            Left            =   14280
            Picture         =   "frmProcessos.frx":07EB
            Top             =   270
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pesquisa por versão :"
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
            Index           =   19
            Left            =   180
            TabIndex        =   139
            Top             =   210
            Width           =   1800
         End
      End
      Begin DrawSuite2022.USToolBar USToolBar1 
         Height          =   975
         Left            =   75
         TabIndex        =   136
         Top             =   330
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   1720
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft6     =   209
         ButtonTop6      =   2
         ButtonWidth6    =   47
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft7     =   258
         ButtonTop7      =   2
         ButtonWidth7    =   46
         ButtonHeight7   =   21
         ButtonUseMaskColor7=   0   'False
         ButtonCaption8  =   "Copiar"
         ButtonEnabled8  =   0   'False
         ButtonIconSize8 =   32
         ButtonToolTipText8=   "Copiar (F7)"
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
         ButtonLeft8     =   306
         ButtonTop8      =   2
         ButtonWidth8    =   39
         ButtonHeight8   =   21
         ButtonUseMaskColor8=   0   'False
         ButtonCaption9  =   "Revisar"
         ButtonEnabled9  =   0   'False
         ButtonIconSize9 =   32
         ButtonToolTipText9=   "Revisar (F8)"
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
         ButtonLeft9     =   347
         ButtonTop9      =   2
         ButtonWidth9    =   44
         ButtonHeight9   =   21
         ButtonUseMaskColor9=   0   'False
         ButtonCaption10 =   "Status"
         ButtonEnabled10 =   0   'False
         ButtonIconSize10=   32
         ButtonToolTipText10=   "Status (F9)"
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
         ButtonLeft10    =   393
         ButtonTop10     =   2
         ButtonWidth10   =   39
         ButtonHeight10  =   21
         ButtonUseMaskColor10=   0   'False
         ButtonCaption11 =   "Validação"
         ButtonEnabled11 =   0   'False
         ButtonIconSize11=   32
         ButtonToolTipText11=   "Validar/cancelar validação do processo (F10)"
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
         ButtonLeft11    =   434
         ButtonTop11     =   2
         ButtonWidth11   =   53
         ButtonHeight11  =   21
         ButtonUseMaskColor11=   0   'False
         ButtonCaption12 =   "Atualizar"
         ButtonEnabled12 =   0   'False
         ButtonIconSize12=   32
         ButtonToolTipText12=   "Utilizado pelo administrador do sistema."
         ButtonKey12     =   "12"
         ButtonAlignment12=   2
         BeginProperty ButtonFont12 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft12    =   489
         ButtonTop12     =   2
         ButtonWidth12   =   50
         ButtonHeight12  =   21
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
         ButtonLeft13    =   541
         ButtonTop13     =   4
         ButtonWidth13   =   2
         ButtonHeight13  =   54
         ButtonCaption14 =   "Ajuda"
         ButtonEnabled14 =   0   'False
         ButtonIconSize14=   32
         ButtonToolTipText14=   "Ajuda (F1)"
         ButtonKey14     =   "14"
         ButtonAlignment14=   2
         BeginProperty ButtonFont14 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft14    =   545
         ButtonTop14     =   2
         ButtonWidth14   =   36
         ButtonHeight14  =   21
         ButtonUseMaskColor14=   0   'False
         ButtonCaption15 =   "Sair"
         ButtonEnabled15 =   0   'False
         ButtonIconSize15=   32
         ButtonToolTipText15=   "Sair (Esc)"
         ButtonKey15     =   "15"
         ButtonAlignment15=   2
         BeginProperty ButtonFont15 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft15    =   583
         ButtonTop15     =   2
         ButtonWidth15   =   26
         ButtonHeight15  =   21
         ButtonUseMaskColor15=   0   'False
         ButtonEnabled16 =   0   'False
         ButtonIconSize16=   32
         ButtonKey16     =   "16"
         ButtonAlignment16=   2
         BeginProperty ButtonFont16 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState16   =   5
         ButtonLeft16    =   611
         ButtonTop16     =   2
         ButtonWidth16   =   24
         ButtonHeight16  =   24
         ButtonUseMaskColor16=   0   'False
         Begin DrawSuite2022.USImageList USImageList1 
            Left            =   11640
            Top             =   195
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmProcessos.frx":0D75
            Count           =   1
         End
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Tempos do processo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   -71805
         TabIndex        =   120
         Top             =   1290
         Width           =   12075
         Begin VB.TextBox txtA5 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   9540
            Locked          =   -1  'True
            TabIndex        =   153
            TabStop         =   0   'False
            Text            =   "000"
            Top             =   270
            Width           =   885
         End
         Begin VB.TextBox txtA4 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            Height          =   210
            Left            =   5040
            Locked          =   -1  'True
            TabIndex        =   151
            TabStop         =   0   'False
            Text            =   "000"
            Top             =   270
            Width           =   1815
         End
         Begin VB.TextBox txtA6 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            Height          =   210
            Left            =   1740
            Locked          =   -1  'True
            TabIndex        =   149
            TabStop         =   0   'False
            Text            =   "00:00:00"
            Top             =   270
            Width           =   1485
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Peça(s) x hora :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   23
            Left            =   8130
            TabIndex        =   154
            Top             =   270
            Width           =   1365
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Segundo(s) x peça :"
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
            Index           =   22
            Left            =   3270
            TabIndex        =   152
            Top             =   270
            Width           =   1665
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total x peça :"
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
            Index           =   21
            Left            =   540
            TabIndex        =   150
            Top             =   270
            Width           =   1125
         End
      End
      Begin VB.Frame Framelista 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   3405
         Left            =   -74925
         TabIndex        =   73
         Top             =   1230
         Width           =   11745
         Begin VB.TextBox txtdetalheitem 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   10260
            MouseIcon       =   "frmProcessos.frx":A698
            MousePointer    =   99  'Custom
            TabIndex        =   92
            ToolTipText     =   "Detalhe."
            Top             =   2340
            Width           =   1305
         End
         Begin VB.TextBox txtdescricaoproduto 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   2250
            Locked          =   -1  'True
            MouseIcon       =   "frmProcessos.frx":A9A2
            MousePointer    =   99  'Custom
            TabIndex        =   91
            TabStop         =   0   'False
            ToolTipText     =   "Descrição do produto."
            Top             =   1020
            Width           =   9315
         End
         Begin VB.TextBox txtcodintproduto 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   175
            Locked          =   -1  'True
            MouseIcon       =   "frmProcessos.frx":ACAC
            MousePointer    =   99  'Custom
            TabIndex        =   90
            TabStop         =   0   'False
            ToolTipText     =   "Código interno do produto."
            Top             =   1020
            Width           =   1965
         End
         Begin VB.TextBox txtrev 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1410
            Locked          =   -1  'True
            MouseIcon       =   "frmProcessos.frx":AFB6
            MousePointer    =   99  'Custom
            TabIndex        =   89
            TabStop         =   0   'False
            ToolTipText     =   "Número da revisão do pedido interno."
            Top             =   390
            Width           =   405
         End
         Begin VB.TextBox txtidcarteira 
            Height          =   315
            Left            =   2250
            MouseIcon       =   "frmProcessos.frx":B2C0
            MousePointer    =   99  'Custom
            TabIndex        =   88
            ToolTipText     =   "Número do pedido interno."
            Top             =   390
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.ComboBox cmbfamilia 
            BackColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   5175
            MouseIcon       =   "frmProcessos.frx":B5CA
            MousePointer    =   99  'Custom
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   87
            ToolTipText     =   "Família."
            Top             =   1740
            Width           =   5460
         End
         Begin VB.CheckBox chkItem 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Criar novo item (cód. automático) ?"
            Height          =   225
            Left            =   2280
            MouseIcon       =   "frmProcessos.frx":B71C
            MousePointer    =   99  'Custom
            TabIndex        =   86
            ToolTipText     =   "Criar um novo item (código automático)"
            Top             =   1770
            Width           =   2775
         End
         Begin VB.TextBox txtObs 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   4530
            MaxLength       =   255
            MouseIcon       =   "frmProcessos.frx":B86E
            MousePointer    =   99  'Custom
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   85
            ToolTipText     =   "Observações."
            Top             =   2962
            Width           =   7035
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   180
            MouseIcon       =   "frmProcessos.frx":BB78
            MousePointer    =   99  'Custom
            TabIndex        =   84
            ToolTipText     =   "Descrição."
            Top             =   2340
            Width           =   7245
         End
         Begin VB.TextBox txtN_Estoque 
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
            Height          =   330
            Left            =   180
            MouseIcon       =   "frmProcessos.frx":BE82
            MousePointer    =   99  'Custom
            TabIndex        =   83
            ToolTipText     =   "Código interno."
            Top             =   1740
            Width           =   1635
         End
         Begin VB.TextBox txtQS 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   8880
            MouseIcon       =   "frmProcessos.frx":C18C
            MousePointer    =   99  'Custom
            TabIndex        =   82
            ToolTipText     =   "Quantidade solicitado."
            Top             =   2340
            Width           =   1305
         End
         Begin VB.TextBox txtQE 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   7500
            Locked          =   -1  'True
            MouseIcon       =   "frmProcessos.frx":C496
            MousePointer    =   99  'Custom
            TabIndex        =   81
            TabStop         =   0   'False
            ToolTipText     =   "Quantidade em estoque."
            Top             =   2340
            Width           =   1305
         End
         Begin VB.ComboBox cmbun 
            BackColor       =   &H00FFFFFF&
            Height          =   330
            ItemData        =   "frmProcessos.frx":C7A0
            Left            =   10710
            List            =   "frmProcessos.frx":C7A2
            MouseIcon       =   "frmProcessos.frx":C7A4
            MousePointer    =   99  'Custom
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   80
            ToolTipText     =   "Unidade."
            Top             =   1740
            Width           =   855
         End
         Begin VB.TextBox cmbProposta 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   180
            Locked          =   -1  'True
            MouseIcon       =   "frmProcessos.frx":C8F6
            MousePointer    =   99  'Custom
            TabIndex        =   79
            TabStop         =   0   'False
            ToolTipText     =   "Número do pedido interno."
            Top             =   390
            Width           =   1215
         End
         Begin VB.TextBox Cmbcliente 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   2250
            Locked          =   -1  'True
            MouseIcon       =   "frmProcessos.frx":CC00
            MousePointer    =   99  'Custom
            TabIndex        =   78
            TabStop         =   0   'False
            ToolTipText     =   "Nome do cliente."
            Top             =   390
            Width           =   9315
         End
         Begin VB.CommandButton cmdselprop 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1825
            MouseIcon       =   "frmProcessos.frx":CF0A
            MousePointer    =   99  'Custom
            Picture         =   "frmProcessos.frx":D05C
            Style           =   1  'Graphical
            TabIndex        =   77
            ToolTipText     =   "Localizar pedido interno."
            Top             =   390
            Width           =   315
         End
         Begin VB.CommandButton cmdEscolher_item 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1825
            MouseIcon       =   "frmProcessos.frx":D15E
            MousePointer    =   99  'Custom
            Picture         =   "frmProcessos.frx":D2B0
            Style           =   1  'Graphical
            TabIndex        =   76
            ToolTipText     =   "Localizar item."
            Top             =   1740
            Width           =   315
         End
         Begin VB.TextBox cmbStatus 
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
            Height          =   330
            Left            =   180
            Locked          =   -1  'True
            MouseIcon       =   "frmProcessos.frx":D3B2
            MousePointer    =   99  'Custom
            TabIndex        =   74
            TabStop         =   0   'False
            ToolTipText     =   "Status do item."
            Top             =   2962
            Width           =   2775
         End
         Begin MSMask.MaskEdBox txtprazo 
            Height          =   315
            Left            =   3030
            TabIndex        =   75
            ToolTipText     =   "Prazo de entrega."
            Top             =   2962
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Status"
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
            Left            =   1290
            TabIndex        =   107
            Top             =   2760
            Width           =   555
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
            BackStyle       =   0  'Transparent
            Caption         =   "Prazo entrega"
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
            Left            =   3082
            TabIndex        =   106
            Top             =   2760
            Width           =   1020
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Detalhe"
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
            Left            =   10635
            TabIndex        =   105
            Top             =   2130
            Width           =   555
         End
         Begin VB.Label Label33 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Cód. interno produto"
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
            Left            =   400
            TabIndex        =   104
            Top             =   810
            Width           =   1515
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Descriçao do produto"
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
            Left            =   5422
            TabIndex        =   103
            Top             =   810
            Width           =   1530
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Rev."
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
            Left            =   1425
            TabIndex        =   102
            Top             =   180
            Width           =   345
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Observação"
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
            Left            =   7612
            TabIndex        =   101
            Top             =   2760
            Width           =   870
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
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
            Left            =   7620
            TabIndex        =   100
            Top             =   2130
            Width           =   1050
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
            BackStyle       =   0  'Transparent
            Caption         =   "Un"
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
            Left            =   11040
            TabIndex        =   99
            Top             =   1530
            Width           =   195
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
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
            Left            =   7635
            TabIndex        =   98
            Top             =   1530
            Width           =   540
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Cliente"
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
            Left            =   5940
            TabIndex        =   97
            Top             =   180
            Width           =   495
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "N° pedido int."
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
            Left            =   292
            TabIndex        =   96
            Top             =   180
            Width           =   990
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
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
            Index           =   1
            Left            =   487
            TabIndex        =   95
            Top             =   2130
            Width           =   690
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Qtde. solicitada"
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
            Left            =   8970
            TabIndex        =   94
            Top             =   2130
            Width           =   1125
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Codigo interno"
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
            Left            =   382
            TabIndex        =   93
            Top             =   1530
            Width           =   1230
         End
         Begin VB.Image imgCalendario 
            Height          =   360
            Left            =   4170
            MouseIcon       =   "frmProcessos.frx":D6BC
            MousePointer    =   99  'Custom
            Picture         =   "frmProcessos.frx":D80E
            Stretch         =   -1  'True
            ToolTipText     =   "Abrir calendário."
            Top             =   2925
            Width           =   330
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00E0E0E0&
         Height          =   6615
         Left            =   -74970
         TabIndex        =   58
         Top             =   1200
         Width           =   11820
         Begin VB.TextBox TxtEmail_Contato 
            Enabled         =   0   'False
            Height          =   330
            Left            =   1770
            MouseIcon       =   "frmProcessos.frx":DC91
            MousePointer    =   99  'Custom
            TabIndex        =   63
            ToolTipText     =   "E-mail do cliente."
            Top             =   1440
            Width           =   9855
         End
         Begin VB.TextBox txttelcontato 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   1770
            MaxLength       =   40
            MouseIcon       =   "frmProcessos.frx":DF9B
            MousePointer    =   99  'Custom
            TabIndex        =   62
            ToolTipText     =   "Ramal do contato."
            Top             =   1020
            Width           =   9855
         End
         Begin VB.TextBox txtdepartamento 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   1770
            MaxLength       =   60
            MouseIcon       =   "frmProcessos.frx":E2A5
            MousePointer    =   99  'Custom
            TabIndex        =   61
            ToolTipText     =   "Departamento do contato."
            Top             =   630
            Width           =   9855
         End
         Begin VB.TextBox txtNomeContato 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   1770
            MaxLength       =   60
            MouseIcon       =   "frmProcessos.frx":E5AF
            MousePointer    =   99  'Custom
            TabIndex        =   60
            ToolTipText     =   "Nome do contato."
            Top             =   240
            Width           =   9855
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
            Left            =   1800
            MaxLength       =   60
            MouseIcon       =   "frmProcessos.frx":E8B9
            MousePointer    =   99  'Custom
            TabIndex        =   59
            ToolTipText     =   "Digite o nome para contato."
            Top             =   240
            Visible         =   0   'False
            Width           =   950
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   4515
            Left            =   150
            TabIndex        =   64
            Top             =   1860
            Width           =   11460
            _ExtentX        =   20214
            _ExtentY        =   7964
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483628
            BorderStyle     =   1
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
               Text            =   "IDContato"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Nome contato"
               Object.Width           =   4939
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Text            =   "Departamento"
               Object.Width           =   4410
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Text            =   "Telefone"
               Object.Width           =   3175
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "E-Mail"
               Object.Width           =   7558
            EndProperty
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "E-Mail:"
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
            Left            =   1215
            TabIndex        =   68
            Top             =   1478
            Width           =   480
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Ramal:"
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
            Left            =   1200
            TabIndex        =   67
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Nome do contato:"
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
            Left            =   405
            TabIndex        =   66
            Top             =   300
            Width           =   1290
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Departamento:"
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
            Left            =   600
            TabIndex        =   65
            Top             =   690
            Width           =   1095
         End
      End
      Begin MSComctlLib.ListView ListaEntrega 
         Height          =   3825
         Left            =   -74970
         TabIndex        =   69
         ToolTipText     =   "Lista de entrega"
         Top             =   4020
         Width           =   11835
         _ExtentX        =   20876
         _ExtentY        =   6747
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
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
            Text            =   "idcliente"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "identrega"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Endereço"
            Object.Width           =   10583
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Bairro"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Cidade"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "UF"
            Object.Width           =   882
         EndProperty
      End
      Begin MSComctlLib.ListView listacobranca 
         Height          =   3825
         Left            =   -74970
         TabIndex        =   70
         Top             =   4020
         Width           =   11835
         _ExtentX        =   20876
         _ExtentY        =   6747
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         MousePointer    =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmProcessos.frx":EBC3
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "idcliente"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "idcobranca"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Endereço"
            Object.Width           =   10583
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Bairro"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Cidade"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "UF"
            Object.Width           =   882
         EndProperty
      End
      Begin MSComctlLib.ListView lista_comerciais 
         Height          =   5055
         Left            =   -74970
         TabIndex        =   71
         Top             =   2730
         Width           =   11790
         _ExtentX        =   20796
         _ExtentY        =   8916
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483628
         BorderStyle     =   1
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
            Text            =   "Status"
            Object.Width           =   5680
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Proposta"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Rev."
            Object.Width           =   970
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Valor total"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Data de emissão"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Data de venda"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "Data de faturamento"
            Object.Width           =   3175
         EndProperty
      End
      Begin MSComctlLib.ListView lista 
         Height          =   2400
         Left            =   -74925
         TabIndex        =   72
         Top             =   4620
         Width           =   11745
         _ExtentX        =   20717
         _ExtentY        =   4233
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         MousePointer    =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmProcessos.frx":ED25
         NumItems        =   12
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Id_lista"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nº do item"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Nº pedido int."
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Rev."
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Cliente"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Cód. interno"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "Un."
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Descrição"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Qtde. solic."
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   9
            Text            =   "Detalhe"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   10
            Text            =   "Prazo entr."
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Status"
            Object.Width           =   2646
         EndProperty
      End
      Begin MSComctlLib.ListView ListaFases 
         Height          =   3825
         Left            =   -74925
         TabIndex        =   55
         Top             =   5250
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   6747
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
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
         NumItems        =   11
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "N"
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Object.Tag             =   "N"
            Text            =   "Fase"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Rev."
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Object.Tag             =   "D"
            Text            =   "Dt. rev."
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "Posto de trab."
            Object.Width           =   2999
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Object.Tag             =   "T"
            Text            =   "Descrição"
            Object.Width           =   8298
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Object.Tag             =   "T"
            Text            =   "Grupo/op."
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   7
            Object.Tag             =   "D"
            Text            =   "T. prep."
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   8
            Object.Tag             =   "D"
            Text            =   "T. exec."
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   9
            Object.Tag             =   "T"
            Text            =   "Coletado"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   10
            Object.Tag             =   "T"
            Text            =   "Tem plano"
            Object.Width           =   2117
         EndProperty
      End
      Begin MSComctlLib.ListView ListaProcessos 
         Height          =   6285
         Left            =   75
         TabIndex        =   13
         Top             =   2790
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   11086
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
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
         NumItems        =   11
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "N"
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Tag             =   "T"
            Text            =   "Processo"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Rev."
            Object.Width           =   970
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Object.Tag             =   "D"
            Text            =   "Dt. rev."
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "Tipo"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Object.Tag             =   "D"
            Text            =   "Dt. emissão"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Object.Tag             =   "T"
            Text            =   "Cód. interno"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Object.Tag             =   "T"
            Text            =   "Descrição"
            Object.Width           =   9093
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Object.Tag             =   "T"
            Text            =   "Status"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   9
            Object.Tag             =   "T"
            Text            =   "Coletado"
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   10
            Object.Tag             =   "T"
            Text            =   "Validado"
            Object.Width           =   1499
         EndProperty
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   1485
         Left            =   75
         TabIndex        =   108
         Top             =   1290
         Width           =   15195
         Begin VB.TextBox txtUn 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   4830
            Locked          =   -1  'True
            MaxLength       =   4
            TabIndex        =   169
            TabStop         =   0   'False
            ToolTipText     =   "Unidade do item"
            Top             =   1020
            Width           =   490
         End
         Begin VB.TextBox txtidprocesso 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   180
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   0
            TabStop         =   0   'False
            ToolTipText     =   "ID do processo."
            Top             =   390
            Width           =   825
         End
         Begin VB.TextBox Txt_numero_processo 
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
            Left            =   1020
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   1
            TabStop         =   0   'False
            ToolTipText     =   "Número do processo."
            Top             =   390
            Width           =   1215
         End
         Begin VB.TextBox txtRespValidacao 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   10395
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   7
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pela validação."
            Top             =   390
            Width           =   3300
         End
         Begin VB.TextBox txtDtValidacao 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   8370
            Locked          =   -1  'True
            TabIndex        =   6
            TabStop         =   0   'False
            ToolTipText     =   "Data de validação."
            Top             =   390
            Width           =   2015
         End
         Begin VB.ComboBox cmbtipoprocesso 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   330
            ItemData        =   "frmProcessos.frx":EE87
            Left            =   2750
            List            =   "frmProcessos.frx":EE9A
            Locked          =   -1  'True
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   3
            TabStop         =   0   'False
            ToolTipText     =   "Tipo."
            Top             =   390
            Width           =   1365
         End
         Begin VB.TextBox txtDtImplantacao 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   4140
            Locked          =   -1  'True
            TabIndex        =   4
            TabStop         =   0   'False
            ToolTipText     =   "Data do cadastro."
            Top             =   390
            Width           =   900
         End
         Begin VB.TextBox txtElaborado 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   5055
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   5
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pelo cadastro."
            Top             =   390
            Width           =   3300
         End
         Begin VB.TextBox txtProduto 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   5340
            Locked          =   -1  'True
            MaxLength       =   255
            TabIndex        =   12
            TabStop         =   0   'False
            ToolTipText     =   "Descrição."
            Top             =   1020
            Width           =   9675
         End
         Begin VB.TextBox txtdesenho 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   180
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   9
            TabStop         =   0   'False
            ToolTipText     =   "Código interno."
            Top             =   1020
            Width           =   1560
         End
         Begin VB.TextBox txtrevdesenho 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   1753
            Locked          =   -1  'True
            MaxLength       =   4
            TabIndex        =   10
            TabStop         =   0   'False
            ToolTipText     =   "Revisão."
            Top             =   1020
            Width           =   490
         End
         Begin VB.TextBox txtrevproc 
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
            Left            =   2250
            Locked          =   -1  'True
            MaxLength       =   4
            TabIndex        =   2
            TabStop         =   0   'False
            ToolTipText     =   "Revisão."
            Top             =   390
            Width           =   480
         End
         Begin VB.ComboBox txtreferencia 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   2265
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   11
            ToolTipText     =   "Código de referência."
            Top             =   1020
            Width           =   2545
         End
         Begin VB.TextBox txtstatus 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   13705
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   8
            TabStop         =   0   'False
            ToolTipText     =   "Status."
            Top             =   390
            Width           =   1310
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Un"
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
            Left            =   4905
            TabIndex        =   170
            Top             =   810
            Width           =   315
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Rev."
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
            Left            =   1826
            TabIndex        =   161
            Top             =   810
            Width           =   345
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Rev."
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
            Left            =   2303
            TabIndex        =   160
            Top             =   180
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "ID"
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
            Left            =   510
            TabIndex        =   157
            Top             =   180
            Width           =   165
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
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
            Index           =   7
            Left            =   11055
            TabIndex        =   147
            Top             =   180
            Width           =   1980
         End
         Begin VB.Label Label49 
            Alignment       =   1  'Right Justify
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
            Left            =   8537
            TabIndex        =   146
            Top             =   180
            Width           =   1680
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
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
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   4
            Left            =   3237
            TabIndex        =   118
            Top             =   180
            Width           =   390
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Dt. emissão"
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
            Left            =   4155
            TabIndex        =   115
            Top             =   180
            Width           =   870
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
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
            Index           =   6
            Left            =   6248
            TabIndex        =   114
            Top             =   180
            Width           =   915
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
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
            Index           =   12
            Left            =   9570
            TabIndex        =   113
            Top             =   810
            Width           =   705
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
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
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   9
            Left            =   465
            TabIndex        =   112
            Top             =   810
            Width           =   990
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Processo"
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
            Left            =   1245
            TabIndex        =   111
            Top             =   180
            Width           =   765
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Cód. referencia"
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
            Left            =   2970
            TabIndex        =   110
            Top             =   810
            Width           =   1125
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
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
            Index           =   8
            Left            =   14113
            TabIndex        =   109
            Top             =   180
            Width           =   495
         End
      End
      Begin DrawSuite2022.USProgressBar PBLista 
         Height          =   255
         Left            =   75
         TabIndex        =   131
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
         SearchText      =   ""
         Value           =   0
      End
      Begin DrawSuite2022.USProgressBar PBlista1 
         Height          =   255
         Left            =   -74925
         TabIndex        =   135
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
         SearchText      =   ""
         Value           =   0
      End
      Begin DrawSuite2022.USToolBar USToolBar2 
         Height          =   975
         Left            =   -74925
         TabIndex        =   137
         Top             =   330
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   1720
         ButtonCount     =   17
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
         ButtonLeft2     =   37
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
         ButtonLeft3     =   77
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
         ButtonLeft4     =   118
         ButtonTop4      =   2
         ButtonWidth4    =   51
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft5     =   171
         ButtonTop5      =   2
         ButtonWidth5    =   47
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft6     =   220
         ButtonTop6      =   2
         ButtonWidth6    =   46
         ButtonHeight6   =   21
         ButtonUseMaskColor6=   0   'False
         ButtonCaption7  =   "Copiar"
         ButtonEnabled7  =   0   'False
         ButtonIconSize7 =   32
         ButtonToolTipText7=   "Copiar (F7)"
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
         ButtonLeft7     =   268
         ButtonTop7      =   2
         ButtonWidth7    =   39
         ButtonHeight7   =   21
         ButtonUseMaskColor7=   0   'False
         ButtonCaption8  =   "Revisar"
         ButtonEnabled8  =   0   'False
         ButtonIconSize8 =   32
         ButtonToolTipText8=   "Revisar (F8)"
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
         ButtonLeft8     =   309
         ButtonTop8      =   2
         ButtonWidth8    =   44
         ButtonHeight8   =   21
         ButtonUseMaskColor8=   0   'False
         ButtonCaption9  =   "Renumerar"
         ButtonEnabled9  =   0   'False
         ButtonIconSize9 =   32
         ButtonToolTipText9=   "Renumerar fases (F9)"
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
         ButtonLeft9     =   355
         ButtonTop9      =   2
         ButtonWidth9    =   61
         ButtonHeight9   =   21
         ButtonUseMaskColor9=   0   'False
         ButtonCaption10 =   "Programas"
         ButtonEnabled10 =   0   'False
         ButtonIconSize10=   32
         ButtonToolTipText10=   "Abrir programas da fase (F10)"
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
         ButtonLeft10    =   418
         ButtonTop10     =   2
         ButtonWidth10   =   59
         ButtonHeight10  =   21
         ButtonUseMaskColor10=   0   'False
         ButtonCaption11 =   "Utensílios"
         ButtonEnabled11 =   0   'False
         ButtonIconSize11=   32
         ButtonToolTipText11=   "Abrir utensílios da fase (F11)"
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
         ButtonLeft11    =   479
         ButtonTop11     =   2
         ButtonWidth11   =   53
         ButtonHeight11  =   21
         ButtonUseMaskColor11=   0   'False
         ButtonCaption12 =   "Tempos"
         ButtonEnabled12 =   0   'False
         ButtonIconSize12=   32
         ButtonToolTipText12=   "Validar os tempos da fase (F12)"
         ButtonKey12     =   "12"
         ButtonAlignment12=   2
         BeginProperty ButtonFont12 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft12    =   534
         ButtonTop12     =   2
         ButtonWidth12   =   45
         ButtonHeight12  =   21
         ButtonUseMaskColor12=   0   'False
         ButtonCaption13 =   "Plano de insp."
         ButtonEnabled13 =   0   'False
         ButtonIconSize13=   32
         ButtonToolTipText13=   "Abrir plano de inspeção (F12)"
         ButtonKey13     =   "13"
         ButtonAlignment13=   2
         BeginProperty ButtonFont13 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft13    =   581
         ButtonTop13     =   2
         ButtonWidth13   =   75
         ButtonHeight13  =   21
         ButtonUseMaskColor13=   0   'False
         ButtonEnabled14 =   0   'False
         ButtonIconSize14=   32
         ButtonAlignment14=   2
         ButtonType14    =   1
         ButtonStyle14   =   -1
         BeginProperty ButtonFont14 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState14   =   -1
         ButtonLeft14    =   658
         ButtonTop14     =   4
         ButtonWidth14   =   2
         ButtonHeight14  =   54
         ButtonCaption15 =   "Ajuda"
         ButtonEnabled15 =   0   'False
         ButtonIconSize15=   32
         ButtonToolTipText15=   "Ajuda (F1)"
         ButtonKey15     =   "15"
         ButtonAlignment15=   2
         BeginProperty ButtonFont15 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft15    =   662
         ButtonTop15     =   2
         ButtonWidth15   =   36
         ButtonHeight15  =   21
         ButtonUseMaskColor15=   0   'False
         ButtonCaption16 =   "Sair"
         ButtonEnabled16 =   0   'False
         ButtonIconSize16=   32
         ButtonToolTipText16=   "Sair (Esc)"
         ButtonKey16     =   "16"
         ButtonAlignment16=   2
         BeginProperty ButtonFont16 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft16    =   700
         ButtonTop16     =   2
         ButtonWidth16   =   26
         ButtonHeight16  =   21
         ButtonUseMaskColor16=   0   'False
         ButtonEnabled17 =   0   'False
         ButtonIconSize17=   32
         ButtonKey17     =   "17"
         ButtonAlignment17=   2
         BeginProperty ButtonFont17 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState17   =   5
         ButtonLeft17    =   728
         ButtonTop17     =   2
         ButtonWidth17   =   24
         ButtonHeight17  =   24
         ButtonUseMaskColor17=   0   'False
         Begin DrawSuite2022.USImageList USImageList2 
            Left            =   12480
            Top             =   195
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmProcessos.frx":EED2
            Count           =   1
         End
      End
      Begin VB.Frame Frame14 
         BackColor       =   &H00E0E0E0&
         Height          =   585
         Left            =   -74925
         TabIndex        =   129
         Top             =   1290
         Width           =   3105
         Begin VB.ComboBox cmbVersao_pesquisar 
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
            ItemData        =   "frmProcessos.frx":188A5
            Left            =   2040
            List            =   "frmProcessos.frx":188A7
            Style           =   2  'Dropdown List
            TabIndex        =   23
            ToolTipText     =   "Versão."
            Top             =   180
            Width           =   795
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pesquisa por versão :"
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
            Index           =   20
            Left            =   180
            TabIndex        =   130
            Top             =   180
            Width           =   1800
         End
      End
      Begin VB.Frame Framedetalhes 
         BackColor       =   &H00E0E0E0&
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
         ForeColor       =   &H00000000&
         Height          =   3375
         Left            =   -74925
         TabIndex        =   116
         Top             =   1860
         Width           =   15195
         Begin VB.CheckBox chkRastreavel 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Rastreável?"
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
            Left            =   13440
            TabIndex        =   172
            Top             =   990
            Width           =   1155
         End
         Begin VB.CheckBox chkPlano_montagem 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Tem plano de montagem?"
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
            Left            =   10740
            TabIndex        =   54
            Top             =   2970
            Width           =   2505
         End
         Begin VB.CommandButton cmdImportar 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   9660
            Picture         =   "frmProcessos.frx":188A9
            Style           =   1  'Graphical
            TabIndex        =   51
            ToolTipText     =   "Localizar arquivo."
            Top             =   2910
            Width           =   315
         End
         Begin VB.TextBox Txt_caminho 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   150
            Locked          =   -1  'True
            MaxLength       =   255
            TabIndex        =   50
            TabStop         =   0   'False
            ToolTipText     =   "Caminho do arquivo."
            Top             =   2910
            Width           =   9495
         End
         Begin VB.CommandButton Cmd_limpar_caminho 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   9990
            Picture         =   "frmProcessos.frx":189AB
            Style           =   1  'Graphical
            TabIndex        =   52
            ToolTipText     =   "Limpar caminho."
            Top             =   2910
            Width           =   315
         End
         Begin VB.CommandButton Cmd_visualizar_arquivo 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   10320
            Picture         =   "frmProcessos.frx":18AE9
            Style           =   1  'Graphical
            TabIndex        =   53
            ToolTipText     =   "Visualizar arquivo."
            Top             =   2910
            Width           =   315
         End
         Begin VB.CheckBox chk_N_apontamento 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Não apontada?"
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
            Left            =   13440
            TabIndex        =   47
            Top             =   1650
            Width           =   1575
         End
         Begin VB.CheckBox Chk_tem_plano 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Tem plano?"
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
            Height          =   195
            Left            =   13440
            TabIndex        =   46
            Top             =   1425
            Width           =   1275
         End
         Begin VB.PictureBox Cor_fonte 
            BackColor       =   &H00000000&
            Height          =   285
            Left            =   12990
            ScaleHeight     =   225
            ScaleWidth      =   195
            TabIndex        =   143
            Top             =   1950
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.CommandButton Cmd_cor 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Cor"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   13230
            Style           =   1  'Graphical
            TabIndex        =   44
            ToolTipText     =   "Mudar cor das instruções de trabalho."
            Top             =   1935
            Width           =   765
         End
         Begin VB.ComboBox Cmb_tamanho_fonte 
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
            ItemData        =   "frmProcessos.frx":190AB
            Left            =   11130
            List            =   "frmProcessos.frx":190AD
            Style           =   2  'Dropdown List
            TabIndex        =   40
            ToolTipText     =   "Tamanho."
            Top             =   2250
            Width           =   735
         End
         Begin VB.ComboBox Cmb_fonte 
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
            Left            =   6510
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   39
            ToolTipText     =   "Fonte."
            Top             =   2250
            Width           =   3495
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00E0E0E0&
            Height          =   1035
            Left            =   11970
            TabIndex        =   140
            Top             =   840
            Width           =   1395
            Begin VB.CheckBox Chk_sublinhado 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Sublinhado"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   90
               TabIndex        =   43
               Top             =   720
               Width           =   1245
            End
            Begin VB.CheckBox Chk_italico 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Itálico"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   90
               TabIndex        =   42
               Top             =   465
               Width           =   915
            End
            Begin VB.CheckBox Chk_negrito 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Negrito"
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
               Left            =   90
               TabIndex        =   41
               Top             =   210
               Width           =   915
            End
         End
         Begin RichTextLib.RichTextBox txtDescricao 
            Height          =   1245
            Left            =   5850
            TabIndex        =   38
            ToolTipText     =   "Instruções de trabalho."
            Top             =   960
            Width           =   6015
            _ExtentX        =   10610
            _ExtentY        =   2196
            _Version        =   393217
            BorderStyle     =   0
            Enabled         =   -1  'True
            ScrollBars      =   2
            TextRTF         =   $"frmProcessos.frx":190AF
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
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
            ItemData        =   "frmProcessos.frx":1912D
            Left            =   180
            List            =   "frmProcessos.frx":1912F
            Style           =   2  'Dropdown List
            TabIndex        =   24
            ToolTipText     =   "Versão."
            Top             =   390
            Width           =   705
         End
         Begin VB.ComboBox Cmb_grupo_posto 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   2950
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   28
            ToolTipText     =   "Grupo do posto de trabalho."
            Top             =   390
            Width           =   2400
         End
         Begin VB.TextBox txtdescmaquina 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   7440
            Locked          =   -1  'True
            TabIndex        =   30
            TabStop         =   0   'False
            ToolTipText     =   "Descrição do posto de trabalho."
            Top             =   390
            Width           =   6195
         End
         Begin VB.TextBox Txt_rev_fase 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   1380
            TabIndex        =   26
            ToolTipText     =   "Revisão."
            Top             =   390
            Width           =   540
         End
         Begin VB.TextBox Txt_data_rev 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   1935
            Locked          =   -1  'True
            MaxLength       =   255
            TabIndex        =   27
            TabStop         =   0   'False
            ToolTipText     =   "Data da revisão."
            Top             =   390
            Width           =   1005
         End
         Begin VB.CommandButton Cmd_grupo_op 
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
            Left            =   14700
            Picture         =   "frmProcessos.frx":19131
            Style           =   1  'Graphical
            TabIndex        =   32
            ToolTipText     =   "Localizar grupo/op."
            Top             =   390
            Width           =   315
         End
         Begin VB.CheckBox chkCronometrado 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Coletado?"
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
            Left            =   13440
            TabIndex        =   45
            Top             =   1200
            Width           =   1155
         End
         Begin VB.CommandButton Cmd_abrir_instrucao 
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
            Height          =   285
            Left            =   11970
            Picture         =   "frmProcessos.frx":19233
            Style           =   1  'Graphical
            TabIndex        =   48
            ToolTipText     =   "Localizar instruções de trabalho."
            Top             =   1935
            Width           =   1005
         End
         Begin VB.Frame Frame12 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1335
            Left            =   150
            TabIndex        =   121
            Top             =   870
            Width           =   5595
            Begin VB.CheckBox chkPchora 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Pçs x Tempo de execução?"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   120
               TabIndex        =   33
               Top             =   30
               Width           =   2535
            End
            Begin VB.Frame Frame15 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Height          =   375
               Left            =   1230
               TabIndex        =   127
               Top             =   240
               Width           =   1095
            End
            Begin VB.TextBox txtPcHora 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   2625
               MaxLength       =   6
               TabIndex        =   36
               Text            =   "1,000"
               ToolTipText     =   "Total de peças por tempo de execução prevista."
               Top             =   810
               Width           =   1035
            End
            Begin VB.TextBox TxtA3 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Height          =   210
               Left            =   4035
               Locked          =   -1  'True
               TabIndex        =   37
               TabStop         =   0   'False
               Text            =   "00:00:00"
               Top             =   855
               Width           =   1305
            End
            Begin MSMask.MaskEdBox txtpreparacao 
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "H:mm:ss"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1046
                  SubFormatType   =   4
               EndProperty
               Height          =   315
               Left            =   195
               TabIndex        =   34
               ToolTipText     =   "Tempo de preparação previsto."
               Top             =   810
               Width           =   1065
               _ExtentX        =   1879
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
               ForeColor       =   0
               AutoTab         =   -1  'True
               MaxLength       =   9
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   "###:##:##"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox txtexecucao 
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "H:mm:ss"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1046
                  SubFormatType   =   4
               EndProperty
               Height          =   315
               Left            =   1290
               TabIndex        =   35
               ToolTipText     =   "Tempo de execução previsto."
               Top             =   810
               Width           =   1065
               _ExtentX        =   1879
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
               ForeColor       =   0
               AutoTab         =   -1  'True
               MaxLength       =   9
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   "###:##:##"
               PromptChar      =   "_"
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "Preparação*"
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
               Index           =   26
               Left            =   180
               TabIndex        =   126
               Top             =   600
               Width           =   1095
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "Execução*"
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
               Index           =   27
               Left            =   1365
               TabIndex        =   125
               Top             =   600
               Width           =   915
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "Pç(s) x exec.*"
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
               Index           =   28
               Left            =   2550
               TabIndex        =   124
               Top             =   600
               Width           =   1185
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Execução x peça"
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
               Index           =   29
               Left            =   3975
               TabIndex        =   123
               Top             =   615
               Width           =   1395
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   " /                             ="
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
               Left            =   2415
               TabIndex        =   122
               Top             =   870
               Width           =   1530
            End
         End
         Begin VB.CommandButton cmdsimbolos 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Símbolos"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   14010
            Style           =   1  'Graphical
            TabIndex        =   49
            ToolTipText     =   "Inserir símbolos especiais nas instruções de trabalho."
            Top             =   1935
            Width           =   1005
         End
         Begin VB.TextBox txtgrupo_op 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   13650
            Locked          =   -1  'True
            TabIndex        =   31
            TabStop         =   0   'False
            ToolTipText     =   "Grupo/operação."
            Top             =   390
            Width           =   1035
         End
         Begin VB.TextBox txtFase 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   900
            TabIndex        =   25
            ToolTipText     =   "Fase."
            Top             =   390
            Width           =   465
         End
         Begin VB.ComboBox cmbMaquina 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   5370
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   29
            ToolTipText     =   "Posto de trabalho."
            Top             =   390
            Width           =   2055
         End
         Begin DrawSuite2022.USButton btnSugestoes 
            Height          =   525
            Left            =   13350
            TabIndex        =   171
            ToolTipText     =   "Abrir sugestões de melhorias do processo"
            Top             =   2700
            Width           =   1665
            _ExtentX        =   2937
            _ExtentY        =   926
            DibPicture      =   "frmProcessos.frx":19335
            Caption         =   "Sugestões"
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
            BorderColorDisabled=   13160660
            BorderColorDown =   4013465
            BorderColorOver =   4408288
            GradientColor1  =   5263559
            GradientColor2  =   5263559
            GradientColor3  =   5263559
            GradientColor4  =   5263559
            GradientColorDisabled1=   13160660
            GradientColorDisabled2=   13160660
            GradientColorDisabled3=   13160660
            GradientColorDisabled4=   13160660
            GradientColorOver1=   4408288
            GradientColorOver2=   4408288
            GradientColorOver3=   4408288
            GradientColorOver4=   4408288
            GradientColorDown1=   4013465
            GradientColorDown2=   4013465
            GradientColorDown3=   4013465
            GradientColorDown4=   4013465
            ShowFocusRect   =   0   'False
            Theme           =   4
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Rev."
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
            Left            =   1478
            TabIndex        =   166
            Top             =   180
            Width           =   345
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Dt. revisão"
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
            Left            =   2040
            TabIndex        =   165
            Top             =   180
            Width           =   795
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Grupo do posto*"
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
            Left            =   3550
            TabIndex        =   164
            Top             =   180
            Width           =   1200
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Posto de trabalho*"
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
            Left            =   5715
            TabIndex        =   163
            Top             =   180
            Width           =   1365
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Descrição do posto de trabalho"
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
            Left            =   9420
            TabIndex        =   162
            Top             =   180
            Width           =   2235
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Caminho do arquivo"
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
            Index           =   33
            Left            =   4185
            TabIndex        =   159
            Top             =   2700
            Width           =   1425
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Grupo/op."
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
            Left            =   13800
            TabIndex        =   158
            Top             =   180
            Width           =   735
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Tamanho :"
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
            Left            =   10260
            TabIndex        =   142
            Top             =   2250
            Width           =   765
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Fonte :"
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
            Left            =   5895
            TabIndex        =   141
            Top             =   2250
            Width           =   525
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
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
            Index           =   24
            Left            =   187
            TabIndex        =   128
            Top             =   180
            Width           =   690
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Instruções de trabalho*"
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
            Left            =   5955
            TabIndex        =   119
            Top             =   780
            Width           =   1725
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Fase*"
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
            Left            =   915
            TabIndex        =   117
            Top             =   180
            Width           =   435
         End
      End
      Begin VB.Frame Frame17 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   75
         TabIndex        =   132
         Top             =   9090
         Width           =   15195
         Begin VB.ComboBox Cmb_opcao_lista2 
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
            ItemData        =   "frmProcessos.frx":1B4E9
            Left            =   6960
            List            =   "frmProcessos.frx":1B4F3
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   187
            Width           =   1965
         End
         Begin VB.TextBox txtNreg 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   2730
            TabIndex        =   14
            Text            =   "30"
            ToolTipText     =   "Número de registros por página."
            Top             =   180
            Width           =   555
         End
         Begin VB.TextBox txtPagIr 
            Height          =   315
            Left            =   9540
            TabIndex        =   16
            ToolTipText     =   "Número da página."
            Top             =   180
            Width           =   555
         End
         Begin DrawSuite2022.USButton cmdPagProx 
            Height          =   315
            Left            =   11760
            TabIndex        =   20
            ToolTipText     =   "Próxima página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmProcessos.frx":1B50B
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
            TabIndex        =   19
            ToolTipText     =   "Página anterior."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmProcessos.frx":1ECAF
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
            TabIndex        =   17
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
            TabIndex        =   18
            ToolTipText     =   "Primeira página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmProcessos.frx":227B8
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
            TabIndex        =   21
            ToolTipText     =   "Última página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmProcessos.frx":268A7
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
            TabIndex        =   167
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
            TabIndex        =   148
            Top             =   240
            Width           =   1260
         End
         Begin VB.Label Label18 
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
            TabIndex        =   145
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
            TabIndex        =   134
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
            TabIndex        =   133
            Top             =   240
            Width           =   1095
         End
      End
      Begin FlexCell.Grid Grid1 
         Height          =   8955
         Left            =   -74925
         TabIndex        =   168
         Top             =   1020
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   15796
         Cols            =   2
         DefaultFontSize =   8.25
         GridColor       =   12632256
         ReadOnly        =   -1  'True
         Rows            =   2
      End
   End
End
Attribute VB_Name = "frmProcessos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TemProcesso                As Boolean 'OK
Public Novo_Processo           As Boolean 'OK
Dim Novo_Processo1             As Boolean 'OK
Public Sql_Processo_Localizar  As String 'OK
Public TBLISTA_Processos       As ADODB.Recordset 'OK
Public Processo_Rastreavel     As Boolean

'GridEstrutura
Public m_Tree As New Node
Public m_Row As Long
Public m_Col As Long
Dim tempNode As Node
Dim intIndex, i As Integer
Dim CodRef As String, DataValidacao As String, RespValidacao As String

Sub ProcAjuda()
On Error GoTo tratar_erro

FunAbrirVideoWeb ("http://www.youtube.com/watch?v=wY9_GSKjErg&list=UUODGiDjQ-BCrxh0YXfJvoqg&index=9&feature=plcp")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ActiveResize1_ResizeComplete()
On Error GoTo tratar_erro

If Txt_ID_fase <> "0" Then
    Set TBTempo = CreateObject("adodb.recordset")
    TBTempo.Open "Select Descricao FROM Fases where IDFase = " & Txt_ID_fase, Conexao, adOpenKeyset, adLockOptimistic
    If TBTempo.EOF = False Then
        txtdescricao.TextRTF = ""
        txtdescricao = TBTempo!Descricao
    End If
    TBTempo.Close
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub btnSugestoes_Click()
On Error GoTo tratar_erro

If Txt_ID_fase.Text <> "" Or IsNumeric(Txt_ID_fase.Text) = True Then
frmProcessos_Sugestoes.Show 1
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_italico_Click()
On Error GoTo tratar_erro

txtdescricao.SelItalic = Chk_italico

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_negrito_Click()
On Error GoTo tratar_erro

txtdescricao.SelBold = Chk_negrito

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_sublinhado_Click()
On Error GoTo tratar_erro

txtdescricao.SelUnderline = Chk_sublinhado

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkPchora_Click()
On Error GoTo tratar_erro:

With txtPcHora
    If chkPchora.Value = 1 Then
        .Locked = False
        .TabStop = True
    Else
        .Locked = True
        .TabStop = False
        .Text = 1
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_fonte_Click()
On Error GoTo tratar_erro

txtdescricao.SelFontName = Cmb_fonte

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_grupo_posto_Click()
On Error GoTo tratar_erro

ProcCarregaMaquinas

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_opcao_lista_Click()
On Error GoTo tratar_erro

With ListaFases
    For InitFor = 1 To .ListItems.Count
        .ListItems.Item(InitFor).Checked = False
    Next InitFor
End With

With USToolBar2
    If Cmb_opcao_lista = "Excluir" Then
        .ButtonState(3) = 0
        .ButtonState(7) = 5
    Else
        .ButtonState(3) = 5
        .ButtonState(7) = 0
    End If
    .Refresh
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_opcao_lista2_Click()
On Error GoTo tratar_erro

With ListaProcessos
    For InitFor = 1 To .ListItems.Count
        .ListItems.Item(InitFor).Checked = False
    Next InitFor
End With

With USToolBar1
    If Cmb_opcao_lista2 = "Excluir" Then
        .ButtonState(4) = 0
        .ButtonState(11) = 5
    Else
        .ButtonState(4) = 5
        .ButtonState(11) = 0
    End If
    .Refresh
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_tamanho_fonte_Click()
On Error GoTo tratar_erro

txtdescricao.SelFontSize = Cmb_tamanho_fonte

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbVersao_Click()
On Error GoTo tratar_erro

If Novo_Processo1 = True Then
    ProcLimparFase
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from fases where idprocesso = " & txtidprocesso & " and versao = '" & cmbVersao & "' order by fase", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        TBAbrir.MoveLast
        Fase = Int(TBAbrir!Fase) + 10
    Else
        Fase = "10"
    End If
    TBAbrir.Close
    txtFase.Text = Fase
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbVersao_pesquisar_Click()
On Error GoTo tratar_erro

'ProcLimparFase
ProcAtualizaFases

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbVersao_pesquisar_estrutura_Click()
On Error GoTo tratar_erro

ProcCarregaEstrutura

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_abrir_instrucao_Click()
On Error GoTo tratar_erro

If cmbMaquina = "" Then
    USMsgBox ("Informe o posto de trabalho antes de localizar as instruções de trabalho."), vbExclamation, "CAPRIND v5.0"
    cmbMaquina.SetFocus
    Exit Sub
End If
Processos_instrucoes = True
RNC_Nao_Conformidade = False
FrmInstrucoes_trabalho.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procCopiar_fase()
On Error GoTo tratar_erro

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
Permitido1 = False
With ListaFases
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente copiar esta(s) fases(s)?", vbYesNo, "CAPRIND v5.0") = vbYes Then
                    With frmProcessos_copiar_fase
                        .Show 1
                        If .Versao_fase = "" Then Exit Sub
                    End With
                    GoTo 1
                Else
                    Exit Sub
                End If
            End If
1:
            Permitido = True
            ProcCopiarFase "Select * from fases where idprocesso = " & txtidprocesso.Text & " and IDFase = " & .ListItems(InitFor), False, txtidprocesso, txtdesenho, txtProduto, False

            '==================================
            Evento = "Nova fase"
            Modulo = "Engenharia/Processos"
            ID_documento = IDlista
            Documento = "Processo: " & txtidprocesso & " - Rev.: " & txtrevproc & " - Cód. interno: " & txtdesenho & " - Rev.: " & txtrevdesenho
            Documento1 = "Versão: " & frmProcessos_copiar_fase.Versao_fase & " - Fase: " & .ListItems.Item(InitFor).ListSubItems(1) & " - Rev.: " & .ListItems.Item(InitFor).ListSubItems(2) & " - Grupo/op.: " & .ListItems.Item(InitFor).ListSubItems(4) & " - Posto: " & .ListItems.Item(InitFor).ListSubItems(5)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe a(s) fases(s) antes de copiar."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Fase(s) copiada(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcCarregaVersao
    cmbVersao_pesquisar = frmProcessos_copiar_fase.Versao_fase
    ProcAtualizaFases
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_cor_Click()
On Error GoTo tratar_erro

With CommonDialog1
    .Color = Cor_fonte.BackColor
    .ShowColor
End With
Cor_fonte.BackColor = CommonDialog1.Color
txtdescricao.SelColor = CommonDialog1.Color

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_grupo_op_Click()
On Error GoTo tratar_erro

frmProcessos_gupoop.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procRevisar_Fase()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If txtFase.Text = "" Then
    USMsgBox ("Informe a fase antes de revisar."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Novo_Processo1 = True Then
    USMsgBox ("Salve a fase antes de revisar."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
frmProcessos_fases_revisao.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAnterior()
On Error GoTo tratar_erro

If txtidprocesso = "" Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Processos order by idprocesso", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.BOF = False Then
    TBLISTA.Find ("idprocesso = " & txtidprocesso)
    TBLISTA.MovePrevious
    If TBLISTA.BOF = False Then
        txtidprocesso = TBLISTA!IDPROCESSO
        Set TBProcessos = CreateObject("adodb.recordset")
        TBProcessos.Open "Select * from Processos where idprocesso = " & txtidprocesso, Conexao, adOpenKeyset, adLockOptimistic
        ProcPuxaDados
        ProcVerificaTipoProcesso
        ProcCarregaMaquinas
        ProcAtualizaFases
    Else
        USMsgBox ("Fim dos cadastros de processo."), vbInformation, "CAPRIND v5.0"
    End If
End If
Novo_Processo = False
Novo_Processo1 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procTempos()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If txtFase.Text = "" Then
    USMsgBox ("Informe a fase antes de validar os tempos."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
frmProcessos_ordens.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcPlanoinspecao()
On Error GoTo tratar_erro

Formulario = "Qualidade/Plano de inspeção"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub

With frmPlanoinspecao
    .Show
    Set TBplano = CreateObject("adodb.recordset")
    TBplano.Open "Select * from Plano where Desenho = '" & txtdesenho & "' and IDFase = " & ListaFases.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
    If TBplano.EOF = False Then
        .ProcLimpar
        .ProcCarregaDados
        .Novo_Plano = False
        .Frame1.Enabled = True
        .StrSql_Plano_Localizar = "Select * from plano where idplano = " & TBplano!IDPlano
        .ProcCarregaLista (1)
    Else
        Direitos
        If Incluir = False Then
            USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
            Exit Sub
        End If
        
        TBplano.AddNew
        TBplano!Data = Date
        TBplano!Inspetor = pubUsuario
        TBplano!Desenho = txtdesenho
        TBplano!Fase = ListaFases.SelectedItem.ListSubItems(1)
        TBplano!Grupo_op = ListaFases.SelectedItem.ListSubItems(4)
        TBplano!Descricao = txtProduto
        TBplano!IDFase = ListaFases.SelectedItem
        TBplano.Update
        .txtPI = TBplano!IDPlano
        
        Set TBplano = CreateObject("adodb.recordset")
        TBplano.Open "Select * from plano where idplano = " & .txtPI, Conexao, adOpenKeyset, adLockOptimistic
        If TBplano.EOF = False Then
            .ProcCarregaDados
        End If
        TBplano.Close
        .Frame1.Enabled = True
        .Novo_Plano = True
    End If
End With

Formulario = "Engenharia/Processos"
Direitos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcProximo()
On Error GoTo tratar_erro

If txtidprocesso = "" Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Processos order by idprocesso", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.BOF = False Then
    TBLISTA.Find ("idprocesso = " & txtidprocesso)
    TBLISTA.MoveNext
    If TBLISTA.EOF = False Then
        txtidprocesso = TBLISTA!IDPROCESSO
        Set TBProcessos = CreateObject("adodb.recordset")
        TBProcessos.Open "Select * from Processos where idprocesso = " & txtidprocesso, Conexao, adOpenKeyset, adLockOptimistic
        ProcPuxaDados
        ProcVerificaTipoProcesso
        ProcCarregaMaquinas
        ProcAtualizaFases
    Else
        USMsgBox ("Fim dos cadastros de processo."), vbInformation, "CAPRIND v5.0"
    End If
End If
Novo_Processo = False
Novo_Processo1 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_limpar_caminho_Click()
On Error GoTo tratar_erro

txt_Caminho = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_visualizar_arquivo_Click()
On Error GoTo tratar_erro

If txt_Caminho <> "" Then ProcAbrirArquivo txt_Caminho

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdImportar_Click()
On Error GoTo tratar_erro

ProcCarregaCaminhoNomeArquivo CommonDialog1, "*.*", "*.*"
txt_Caminho = caminho

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Processos.AbsolutePage <> 2 Then
    If TBLISTA_Processos.AbsolutePage = -3 Then
        ProcExibePagina (TBLISTA_Processos.PageCount - 1)
    Else
        TBLISTA_Processos.AbsolutePage = TBLISTA_Processos.AbsolutePage - 2
        ProcExibePagina (TBLISTA_Processos.AbsolutePage)
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
    TBLISTA_Processos.AbsolutePage = txtPagIr.Text
    ProcExibePagina (TBLISTA_Processos.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Processos.AbsolutePage = 1
ProcExibePagina (TBLISTA_Processos.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Processos.AbsolutePage <> -3 Then
    If TBLISTA_Processos.AbsolutePage = 1 Then
        ProcExibePagina (2)
    Else
        ProcExibePagina (TBLISTA_Processos.AbsolutePage)
    End If
Else
    ProcExibePagina (TBLISTA_Processos.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Processos.AbsolutePage = TBLISTA_Processos.PageCount
ProcExibePagina (TBLISTA_Processos.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaFases_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With ListaFases
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If FunVerificaRegistroValidadoSemMsg("Processos", "IDprocesso = " & txtidprocesso, True) = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
            
                .ListItems.Item(InitFor).Checked = True
            End If
Proximo:
        Next InitFor
    End With
Else
    ProcOrdenaListView ListaFases, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Listafases_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With ListaFases
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True And .ListItems.Item(InitFor) = Item Then
            If FunVerificaRegistroValidado("Processos", "IDProcesso = " & txtidprocesso, "processo", "fase", IIf(Cmb_opcao_lista = "Excluir", "excluir esta", "copiar esta"), True, True) = False Then
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

Private Sub ListaProcessos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With ListaProcessos
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If Cmb_opcao_lista2 = "Excluir" Then
                    If .ListItems(InitFor).SubItems(9) = "SIM" Then GoTo Proximo
                    If FunVerificaRegistroValidadoSemMsg("Processos", "IDProcesso = " & .ListItems(InitFor), True) = False Then GoTo Proximo
                    ProcVerificaRegistroUtilizadoSemMsg "Producao", "idprocesso = " & .ListItems(InitFor)
                    If Permitido = False Then GoTo Proximo
                    ProcVerificaRegistroUtilizadoSemMsg "plano", "desenho = '" & .ListItems(InitFor).SubItems(5) & "'"
                    If Permitido = False Then GoTo Proximo
                Else
                    Set TBProcessos = CreateObject("adodb.recordset")
                    TBProcessos.Open "Select IDprocesso from processos where NProcesso = '" & .ListItems(InitFor).ListSubItems(1) & "' and Revisao > " & .ListItems(InitFor).ListSubItems(2), Conexao, adOpenKeyset, adLockOptimistic
                    If TBProcessos.EOF = False Then
                        TBProcessos.Close
                        GoTo Proximo
                    End If
                    Set TBProcessos = CreateObject("adodb.recordset")
                    TBProcessos.Open "Select P.IDprocesso from processos P LEFT JOIN Fases F ON F.IDprocesso = P.IDProcesso where P.IDProcesso = " & .ListItems(InitFor) & " and DtValidacao IS NULL and F.IDFase IS NULL", Conexao, adOpenKeyset, adLockOptimistic
                    If TBProcessos.EOF = False Then
                        TBProcessos.Close
                        GoTo Proximo
                    End If
                    TBProcessos.Close
                End If
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView ListaProcessos, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaProcessos_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With ListaProcessos
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True And .ListItems.Item(InitFor) = Item Then
            If Cmb_opcao_lista2 = "Excluir" Then
                If FunVerificaRegistroValidado("Processos", "IDProcesso = " & .ListItems(InitFor), "mesmo", "o processo", "excluir", True, True) = False Then
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
                End If
                
                Mensagem = "Não é permitido excluir este processo, pois o mesmo está sendo utilizado no módulo"
                ProcVerificaRegistroUtilizado "Producao", "idprocesso = " & .ListItems(InitFor), "PCP/Gerenciamento de ordem"
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
                End If
                
                Set TBplano = CreateObject("adodb.recordset")
                TBplano.Open "Select P.IDplano from plano P INNER JOIN Fases F ON F.IDFase = P.IDFase INNER JOIN Processos PRO ON PRO.IDprocesso = F.IDprocesso where PRO.IDProcesso = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
                If TBplano.EOF = False Then
                    If USMsgBox("A(s) fase(s) deste processo está(ão) sendo usada(s) no módulo plano de inspeção, deseja realmente excluir a(s) fase(s), o(s) plano(s) e todos utensílios e programas vinculados?", vbYesNo, "CAPRIND v5.0") = vbNo Then
                        .ListItems.Item(InitFor).Checked = False
                    End If
                End If
                TBplano.Close
            Else
                Set TBProcessos = CreateObject("adodb.recordset")
                TBProcessos.Open "Select IDprocesso from processos where NProcesso = '" & .ListItems(InitFor).ListSubItems(1) & "' and Revisao > " & .ListItems(InitFor).ListSubItems(2), Conexao, adOpenKeyset, adLockOptimistic
                If TBProcessos.EOF = False Then
                    USMsgBox ("Não é permitido alterar a validação deste processo, pois o mesmo já foi revisado."), vbExclamation, "CAPRIND v5.0"
                    .ListItems.Item(InitFor).Checked = False
                    TBProcessos.Close
                    Exit Sub
                End If
                Set TBProcessos = CreateObject("adodb.recordset")
                TBProcessos.Open "Select P.IDprocesso from processos P LEFT JOIN Fases F ON F.IDprocesso = P.IDProcesso where P.IDProcesso = " & .ListItems(InitFor) & " and DtValidacao IS NULL and F.IDFase IS NULL", Conexao, adOpenKeyset, adLockOptimistic
                If TBProcessos.EOF = False Then
                    USMsgBox ("Não é permitido validar este processo, pois não existe fase cadastrada para o mesmo."), vbExclamation, "CAPRIND v5.0"
                    .ListItems.Item(InitFor).Checked = False
                End If
                TBProcessos.Close
            End If
        End If
    Next InitFor
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaProcessos_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If ListaProcessos.ListItems.Count = 0 Then Exit Sub
Set TBProcessos = CreateObject("adodb.recordset")
TBProcessos.Open "Select * from processos where idprocesso = " & ListaProcessos.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBProcessos.EOF = False Then
    ProcPuxaDados
    CodigoLista = ListaProcessos.SelectedItem.index
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error GoTo tratar_erro

If txtidprocesso.Text = "" Then
    SSTab1.Tab = 0
    Exit Sub
End If
Select Case SSTab1.Tab
    Case 0:
        If ListaProcessos.Visible = True Then ListaProcessos.SetFocus
    Case 1:
        cmbVersao_pesquisar_estrutura.SetFocus
        ProcVerificaProsseguir
        If Permitido = False Then Exit Sub
        ProcVerificaTipoProcesso
        ProcCarregaEstrutura
    Case 2:
        cmbVersao_pesquisar.SetFocus
        ProcVerificaProsseguir
        If Permitido = False Then Exit Sub
        ProcVerificaTipoProcesso
        ProcCarregaVersao
        ProcLimparFase
        ProcCarregaGrupos
        'ProcAtualizaFases
        ProcAtualizaCustoProcesso
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimparTudo()
On Error GoTo tratar_erro

Framedetalhes.Enabled = False
ProcLimparFase
ListaFases.ListItems.Clear
Novo_Processo1 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcVerificaProsseguir()
On Error GoTo tratar_erro

Permitido = True
If Novo_Processo = True Then
    USMsgBox ("Salve o processo antes de prosseguir."), vbExclamation, "CAPRIND v5.0"
    SSTab1.Tab = 0
    Permitido = False
    Exit Sub
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcVerificaTipoProcesso()
On Error GoTo tratar_erro

Set TBProcessos = CreateObject("adodb.recordset")
TBProcessos.Open "Select tipo from processos where idprocesso = " & txtidprocesso.Text, Conexao, adOpenKeyset, adLockOptimistic
If TBProcessos.EOF = False Then
    If IsNull(TBProcessos!Tipo) = True Or TBProcessos!Tipo = "" Then
        USMsgBox ("Informe o tipo do processo e salve antes de prosseguir."), vbExclamation, "CAPRIND v5.0"
        SSTab1.Tab = 0
        TBProcessos.Close
        Exit Sub
    End If
End If
TBProcessos.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcAtualizaFases()
On Error GoTo tratar_erro

Total = 0
Total1 = 0
ListaFases.ListItems.Clear
Set TBFases = CreateObject("adodb.recordset")
TBFases.Open "Select * FROM Fases WHERE IDProcesso = " & txtidprocesso & " and versao = '" & cmbVersao_pesquisar & "' order by fase, IDFase", Conexao, adOpenKeyset, adLockOptimistic
If TBFases.EOF = False Then
    PBLista1.Min = 0
    PBLista1.Max = TBFases.RecordCount
    PBLista1.Value = 1
    Contador = 0
    Do While TBFases.EOF = False
        With ListaFases.ListItems
            .Add , , TBFases!IDFase
            .Item(.Count).SubItems(1) = IIf(IsNull(TBFases!Fase) = False, TBFases!Fase, "")
            
            .Item(.Count).SubItems(2) = IIf(IsNull(TBFases!Revisao), 0, TBFases!Revisao)
            Set TBEventos = CreateObject("adodb.recordset")
            TBEventos.Open "Select revisao, data from Fases_revisao where IDFase = " & TBFases!IDFase & " order by ID desc", Conexao, adOpenKeyset, adLockOptimistic
            If TBEventos.EOF = False Then
                .Item(.Count).SubItems(3) = IIf(IsNull(TBEventos!Data), "", Format(TBEventos!Data, "dd/mm/yy"))
            End If
            TBEventos.Close
            
            .Item(.Count).SubItems(4) = IIf(IsNull(TBFases!maquina) = False, TBFases!maquina, "")
            Set TBMaquinas = CreateObject("adodb.recordset")
            TBMaquinas.Open "Select Descricao from CadMaquinas where Maquina = '" & TBFases!maquina & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBMaquinas.EOF = False Then
                .Item(.Count).SubItems(5) = TBMaquinas!Descricao
            End If
            TBMaquinas.Close
            .Item(.Count).SubItems(6) = IIf(IsNull(TBFases!Grupo_op) = False, TBFases!Grupo_op, "")
            .Item(.Count).SubItems(7) = IIf(IsNull(TBFases!TempoPreparacao), "00:00:00", TBFases!TempoPreparacao)
            
            If IsNull(TBFases!pc_te) = True Then
                TBFases!pc_te = 1
                TBFases.Update
            End If

            Total1 = Total1 + IIf(IsNull(TBFases!TESegundos), 0, TBFases!TESegundos)
            .Item(.Count).SubItems(8) = IIf(IsNull(TBFases!TempoExecucao), "00:00:00", TBFases!TempoExecucao)
            .Item(.Count).SubItems(9) = IIf((TBFases!cronometrado = True), "SIM", "NÃO")
            .Item(.Count).SubItems(10) = IIf((TBFases!Plano_inspecao = True), "SIM", "NÃO")
        End With
        TBFases.MoveNext
        Contador = Contador + 1
        PBLista1.Value = Contador
    Loop
End If
1:
    TBFases.Close
    txtA4.Text = Total1
    ProcCalculaA5
    ProcCalculaA6

Exit Sub
tratar_erro:
    If Err.Number = 365 Then GoTo 1
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub Procentrada()
On Error GoTo tratar_erro

TempoPreparacao = ""
TempoExecucao = ""

'Rotina de tranformacao de preparacao
If IsNull(TBFases!Preparacao) = False And TBFases!Preparacao <> "__:__:__" Then
    ElapsedTime (TBFases!Preparacao)
    TempoPreparacao = HoraTotal
    If Len(TempoPreparacao) = 8 Then
        TempoPreparacao = "0" & TempoPreparacao
    End If
End If

'Rotina de tranformacao de execucao
If IsNull(TBFases!Execucao) = False And TBFases!Execucao <> "__:__:__" Then
    ElapsedTime (TBFases!Execucao)
    TempoExecucao = HoraTotal
    If Len(TempoExecucao) = 8 Then
        TempoExecucao = "0" & TempoExecucao
    End If
End If
        
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub Procsaida()
On Error GoTo tratar_erro

'=======================================================
' CALCULO DE TEMPO DE PREPARACAO MAIOR QUE 23:59:59 HS =
'=======================================================
DataResultado = 0
If txtpreparacao > "023:59:59" Then
    ProcFormataHora (txtpreparacao)
Else
    DataResultado = txtpreparacao
End If
ElapsedTime (DataResultado)
Preparacao = DataResultado
TempoPreparacao = HoraTotal
'=====================================================
' CALCULO DE TEMPO DE EXECUÇÃO MAIOR QUE 23:59:59 HS =
'=====================================================
DataResultado = 0
If txtexecucao > "023:59:59" Then
    ProcFormataHora (txtexecucao)
Else
    DataResultado = txtexecucao
End If
ElapsedTime (DataResultado)
Execucao = DataResultado
TempoExecucao = HoraTotal

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbMaquina_Click()
On Error GoTo tratar_erro

Set TBMaquinas = CreateObject("adodb.recordset")
TBMaquinas.Open "Select * FROM CadMaquinas where maquina = '" & cmbMaquina.Text & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBMaquinas.EOF = False Then
    txtdescmaquina.Text = Trim(TBMaquinas!Descricao)
End If
TBMaquinas.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procExcluir_fase()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With ListaFases
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir esta(s) fase(s)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            
            'Exclui dados do plano de inspeção
            Conexao.Execute "DELETE from PDI from ((plano P INNER JOIN Fases F ON F.IDFase = P.IDFase) INNER JOIN Planodimensao PD ON PD.IdPlano = P.IdPlano) INNER JOIN Planodimensao_instrumentos PDI ON PDI.ID_dimensao = PD.idDimensao where P.IDfase = " & .ListItems(InitFor)
            Conexao.Execute "DELETE from PD from (plano P INNER JOIN Fases F ON F.IDFase = P.IDFase) INNER JOIN Planodimensao PD ON PD.IdPlano = P.IdPlano where P.IDfase = " & .ListItems(InitFor)
            Conexao.Execute "DELETE from PR from (plano P INNER JOIN Fases F ON F.IDFase = P.IDFase) INNER JOIN Plano_revisao PR ON PR.IdPlano = P.IdPlano where P.IDfase = " & .ListItems(InitFor)
            Conexao.Execute "DELETE from plano where IDfase = " & .ListItems(InitFor)
                       
            Conexao.Execute "DELETE from Fases WHERE IDFase = " & .ListItems(InitFor)
            Conexao.Execute "DELETE from Fases_revisao WHERE IDFase = " & .ListItems(InitFor)
            Conexao.Execute "DELETE from Ferramentas WHERE IDProcesso = " & txtidprocesso & " AND IDFase = " & .ListItems(InitFor)
            Conexao.Execute "DELETE from Programas WHERE IDProcesso = " & txtidprocesso & " AND IDFase = " & .ListItems(InitFor)
            
            '==================================
            Modulo = "Engenharia/Processos"
            Evento = "Excluir fase"
            ID_documento = .ListItems(InitFor)
            Documento = "Processo: " & txtidprocesso & " - Rev.: " & txtrevproc & " - Cód. interno: " & txtdesenho & " - Rev.: " & txtrevdesenho
            Documento1 = "Versão: " & cmbVersao_pesquisar & " - Fase: " & .ListItems(InitFor).SubItems(1) & " - Rev.: " & .ListItems(InitFor).SubItems(2) & " - Grupo/op.: " & .ListItems(InitFor).SubItems(4) & " - Posto: " & .ListItems(InitFor).SubItems(5)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe a(s) fase(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Fase(s) excluída(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    
    ProcCarregaVersao
    Set TBItem = CreateObject("adodb.recordset")
    TBItem.Open "Select versao from fases WHERE IDProcesso = " & txtidprocesso & " and Versao = '" & cmbVersao & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBItem.EOF = False Then
        cmbVersao_pesquisar = cmbVersao
    End If
    TBItem.Close
    
    ProcAtualizaFases
    ProcAtualizaCustoProcesso
    
    ProcLimparFase
    Framedetalhes.Enabled = False
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
If txtidprocesso = "" Then
    USMsgBox ("Informe o processo antes de alterar o status."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If ProcVerifiProcRevisado(Txt_numero_processo, txtrevproc, "alterar o status deste processo", "mesmo", True) = True Then Exit Sub

frmProcessos_bloq.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Function ProcVerifiProcRevisado(Nprocesso As String, Revisao As Long, MsgemPadrao As String, MsgemProc As String, MostrarMsgem As Boolean) As Boolean
On Error GoTo tratar_erro

ProcVerifiProcRevisado = False
Set TBTempo = CreateObject("adodb.recordset")
TBTempo.Open "Select IDprocesso from processos where NProcesso = '" & Nprocesso & "' and Revisao > " & Revisao, Conexao, adOpenKeyset, adLockOptimistic
If TBTempo.EOF = False Then
    If MostrarMsgem = True Then USMsgBox ("Não é permitido " & MsgemPadrao & ", pois o " & MsgemProc & " já foi revisado."), vbExclamation, "CAPRIND v5.0"
    ProcVerifiProcRevisado = True
End If
TBTempo.Close

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Private Sub ProcFerramentas()
On Error GoTo tratar_erro
  
If txtFase.Text = "" Then
    USMsgBox ("Informe a fase antes de abrir o módulo de utensílios."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Novo_Processo1 = True Then
    USMsgBox ("Salve a fase antes de abrir o módulo de utensílios."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
frmFerramentasdafase.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procSalvar_fase()
On Error GoTo tratar_erro
Dim UltSel As Long 'OK

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Framedetalhes.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
Acao = "salvar"
If cmbVersao = "" Then
    NomeCampo = "a versão"
    ProcVerificaAcao
    cmbVersao.SetFocus
    Exit Sub
End If
If txtFase = "" Then
    NomeCampo = "a fase"
    ProcVerificaAcao
    txtFase.SetFocus
    Exit Sub
End If
If Cmb_grupo_posto.Text = "" Then
    NomeCampo = "o grupo do posto de trabalho"
    ProcVerificaAcao
    Cmb_grupo_posto.SetFocus
    Exit Sub
End If
If cmbMaquina.Text = "" Then
    NomeCampo = "o posto de trabalho"
    ProcVerificaAcao
    cmbMaquina.SetFocus
    Exit Sub
End If

txtpreparacao.PromptInclude = False
If Len(txtpreparacao.Text) < 7 Then
    txtpreparacao.PromptInclude = True
    USMsgBox "Verifique se faltam dados no campo preparação ( " & txtpreparacao.Text & " ) á serem preenchidos.", vbExclamation, "CAPRIND v5.0"
    txtpreparacao.SetFocus
    Exit Sub
End If
txtpreparacao.PromptInclude = True
txtexecucao.PromptInclude = False
If Len(txtexecucao.Text) < 7 Then
    txtexecucao.PromptInclude = True
    USMsgBox "Verifique se faltam dados no campo execução ( " & txtexecucao.Text & " ) á serem preenchidos.", vbExclamation, "CAPRIND v5.0"
    txtexecucao.SetFocus
    Exit Sub
End If
txtexecucao.PromptInclude = True

valor = IIf(txtPcHora = "", 0, txtPcHora)
If valor <= 0 Then
    NomeCampo = "a quantidade de peças por hora"
    ProcVerificaAcao
    txtPcHora.SetFocus
    Exit Sub
End If
If txtdescricao.Text = "" Then
    NomeCampo = "a descriçao da fase"
    ProcVerificaAcao
    txtdescricao.SetFocus
    Exit Sub
End If

Set TBFases = CreateObject("adodb.recordset")
TBFases.Open "Select * from fases where IDfase = " & IIf(Txt_ID_fase = "", 0, Txt_ID_fase), Conexao, adOpenKeyset, adLockOptimistic
If TBFases.EOF = True Then
    TBFases.AddNew
Else
    If FunVerificaRegistroValidado("Processos", "IDProcesso = " & txtidprocesso, "processo", "a fase", "alterar", True, True) = False Then Exit Sub
    
    'Verifica se existe plano de insepção para fase
    If txtFase <> TBFases!Fase Then
        Set TBplano = CreateObject("adodb.recordset")
        TBplano.Open "Select IDplano from Plano where IDFase = " & TBFases!IDFase, Conexao, adOpenKeyset, adLockOptimistic
        If TBplano.EOF = False Then
            USMsgBox ("Não é permitido alterar esta fase, pois a mesma está sendo utilizada no módulo Qualidade/Plano de inspeção."), vbExclamation, "CAPRIND v5.0"
            TBplano.Close
            Exit Sub
        End If
        TBplano.Close
    End If
    If txtgrupo_op <> TBFases!Grupo_op Then Conexao.Execute "Update Plano Set Grupo_op = '" & txtgrupo_op & "' where IDfase = " & TBFases!IDFase
End If
TBFases!IDPROCESSO = txtidprocesso
TBFases!Fase = txtFase.Text
TBFases!Revisao = IIf(Txt_rev_fase = "", 0, Txt_rev_fase)
TBFases!Grupo_op = txtgrupo_op.Text
TBFases!maquina = cmbMaquina.Text
TBFases!Descricao = txtdescricao.TextRTF
TBFases!caminho = txt_Caminho

Procsaida
TBFases!Preparacao = Preparacao
ElapsedTime (Preparacao)
TBFases!TPSegundos = s
TBFases!TempoPreparacao = TempoPreparacao

TBFases!Execucao = Execucao
TBFases!TESegundos = FunCalculaSegPC(TBFases!Execucao, txtPcHora)
TBFases!TempoExecucao = TxtA3.Text

TBFases!versao = cmbVersao.Text
TBFases!pc_te = txtPcHora.Text
If chkCronometrado.Value = 1 Then TBFases!cronometrado = True Else TBFases!cronometrado = False
If chkPchora.Value = 1 Then TBFases!pecahora = True Else TBFases!pecahora = False
If chk_N_apontamento.Value = 1 Then TBFases!Nao_aponta = True Else TBFases!Nao_aponta = False
If chkPlano_montagem.Value = 1 Then TBFases!Plano_montagem = True Else TBFases!Plano_montagem = False
If chkRastreavel.Value = 1 Then TBFases!rastreavel = True Else TBFases!rastreavel = False

TBFases.Update
Txt_ID_fase = TBFases!IDFase
Conexao.Execute "UPDATE plano Set Fase = " & txtFase & ", Grupo_op = '" & txtgrupo_op & "' where IDFase = " & Txt_ID_fase
TBFases.Close

ProcCarregaVersao
If Novo_Processo1 = True Then
    USMsgBox ("Nova fase cadastrada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Nova fase"
    
    cmbVersao_pesquisar = cmbVersao
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar fase"
    If CodigoLista1 <> 0 And ListaFases.ListItems.Count <> 0 Then
        ListaFases.SelectedItem = ListaFases.ListItems(CodigoLista1)
        ListaFases.SetFocus
    End If
End If
'==================================
Modulo = "Engenharia/Processos"
ID_documento = IDFase
Documento = "Processo: " & txtidprocesso & " - Rev.: " & txtrevproc & " - Cód. interno: " & txtdesenho & " - Rev.: " & txtrevdesenho
Documento1 = "Versão: " & cmbVersao & " - Fase: " & txtFase & " - Rev.: " & Txt_rev_fase & " - Grupo/op.: " & txtgrupo_op & " - Posto: " & cmbMaquina
ProcGravaEvento
'==================================
Novo_Processo1 = False
ProcAtualizaCustoProcesso
  
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procNovaFase()
On Error GoTo tratar_erro

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If FunVerificaRegistroValidado("Processos", "IDProcesso = " & txtidprocesso, "processo", "fase", "criar nova", True, True) = False Then Exit Sub
If ProcVerifiProcRevisado(Txt_numero_processo, txtrevproc, "criar nova fase", "processo", True) = True Then Exit Sub
Txt_ID_fase = 0
Novo_Processo1 = True
Framedetalhes.Enabled = True
If cmbVersao_pesquisar = "" Then cmbVersao = "A" Else cmbVersao = cmbVersao_pesquisar

cmbVersao.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procProgramas()
On Error GoTo tratar_erro
  
If txtFase.Text = "" Then
    USMsgBox ("Informe a fase antes de abrir o módulo de programas."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Novo_Processo1 = True Then
    USMsgBox ("Salve a fase antes de abrir o módulo de programas."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

frmProgramas.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcRevisar()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If txtidprocesso = "" Then
    USMsgBox ("Informe o processo antes de criar revisão."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
frmProcessos_revisao.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcRenumerar()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If txtidprocesso = 0 Or ListaFases.ListItems.Count = 0 Then Exit Sub
If FunVerificaRegistroValidado("Processos", "IDProcesso = " & txtidprocesso, "processo", "a fase", "renumerar", True, True) = False Then Exit Sub
Fase = 10
Set TBFases = CreateObject("adodb.recordset")
TBFases.Open "Select * from fases where idprocesso = " & txtidprocesso & " and versao = '" & cmbVersao_pesquisar & "' order by fase", Conexao, adOpenKeyset, adLockOptimistic
If TBFases.EOF = False Then
    Do While TBFases.EOF = False
        Conexao.Execute "UPDATE Plano Set Fase = " & Fase & " where IDFase = " & TBFases!IDFase
        
        TBFases!Fase = Fase
        TBFases.Update
        Fase = Fase + 10
        TBFases.MoveNext
    Loop
End If
TBFases.Close
USMsgBox ("Fases renumeradas com sucesso."), vbInformation, "CAPRIND v5.0"
'==================================
Modulo = "Engenharia/Processos"
Evento = "Renumerar fases"
ID_documento = txtidprocesso.Text
Documento = "Processo: " & txtidprocesso & " - Rev.: " & txtrevproc & " - Cód. interno: " & txtdesenho & " - Rev.: " & txtrevdesenho
Documento1 = ""
ProcGravaEvento
'==================================
ProcLimparFase
ProcAtualizaFases

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdsimbolos_Click()
On Error GoTo tratar_erro

RNC_Nao_Conformidade = False
frmsimbolos.Show 1

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
            Case vbKeyF4: If Cmb_opcao_lista2 = "Excluir" Then ProcExcluir
            Case vbKeyF5: ProcImprimir
            Case vbKeyF7: procCopiarProcesso
            Case vbKeyF8: ProcRevisar
            Case vbKeyF9: ProcStatus
            Case vbKeyF10: If Cmb_opcao_lista2 = "Validação" Then ProcValidarRegistros ListaProcessos, "Engenharia/Processos"
            Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
    Case 1:
        Select Case KeyCode
            Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
    Case 2:
        Select Case KeyCode
            Case vbKeyInsert: procNovaFase
            Case vbKeyF3: procSalvar_fase
            Case vbKeyF4: If Cmb_opcao_lista = "Excluir" Then procExcluir_fase
            Case vbKeyF5: ProcImprimir
            Case vbKeyF7: If Cmb_opcao_lista = "Copiar" Then procCopiar_fase
            Case vbKeyF8: procRevisar_Fase
            Case vbKeyF9: ProcRenumerar
            Case vbKeyF10: procProgramas
            Case vbKeyF11: ProcFerramentas
            Case vbKeyF12: procTempos
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

ProcCarregaToolBar1 Me, 15195, 16, True
ProcCarregaToolBar2 Me, 15195, 17, True
Formulario = "Engenharia/Processos"
Direitos
ProcCarregaComboFontes Cmb_fonte
ProcCarregaComboTamanhoFonte Cmb_tamanho_fonte, 8, 16
ProcCarregaComboVersao cmbVersao, False, False, False, False, ""
SSTab1.Tab = 0
ProcLimpaVariaveisPrincipais
cmbVersao_pesquisar_estrutura = "A"
Cmb_opcao_lista2 = "Validação"
Cmb_opcao_lista = "Excluir"

ProcRemoveObjetosResize Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaGrupos()
On Error GoTo tratar_erro

With Cmb_grupo_posto
    .Clear
    Set TBMaquinas = CreateObject("adodb.recordset")
    TBMaquinas.Open "Select Grupo FROM CadMaquinas where Grupo is not null and Bloqueado = 'False' group by grupo", Conexao, adOpenKeyset, adLockOptimistic
    If TBMaquinas.EOF = False Then
        Do While TBMaquinas.EOF = False
            .AddItem TBMaquinas!Grupo
            TBMaquinas.MoveNext
        Loop
    End If
    TBMaquinas.Close
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaMaquinas()
On Error GoTo tratar_erro

txtdescmaquina = ""
ProcCarregaComboPostoTrab cmbMaquina, "Grupo = '" & Cmb_grupo_posto & "' and Bloqueado = 'False'", False, False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaListaProcessos(Pagina As Integer)
On Error GoTo tratar_erro

lblRegistros.Caption = "Nº de registros: 0"
lblPaginas.Caption = "Página: 0 de: 0"
ListaProcessos.ListItems.Clear
If Sql_Processo_Localizar = "" Then Exit Sub
Set TBLISTA_Processos = CreateObject("adodb.recordset")
TBLISTA_Processos.Open Sql_Processo_Localizar, Conexao, adOpenKeyset, adLockReadOnly
If TBLISTA_Processos.EOF = False Then ProcExibePagina (Pagina)
        
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina(Pagina)
On Error GoTo tratar_erro

ListaProcessos.ListItems.Clear
TBLISTA_Processos.PageSize = IIf(txtNreg = "", 30, txtNreg)
TBLISTA_Processos.AbsolutePage = Pagina
TamanhoPagina = TBLISTA_Processos.PageSize
ContadorReg = 1

PBLista.Min = 0
PBLista.Max = FunVerifMaxPBListaPaginacao(TBLISTA_Processos.RecordCount - IIf(Pagina > 1, (TBLISTA_Processos.PageSize * (Pagina - 1)), 0), TBLISTA_Processos.PageSize)
PBLista.Value = 1
Contador = 0
Do While TBLISTA_Processos.EOF = False And (ContadorReg <= TamanhoPagina)
    With ListaProcessos.ListItems
        .Add , , TBLISTA_Processos!IDPROCESSO
        .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA_Processos!Nprocesso), 0, TBLISTA_Processos!Nprocesso)
        .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA_Processos!Revisao), 0, TBLISTA_Processos!Revisao)
        Set TBProduto = CreateObject("adodb.recordset")
        TBProduto.Open "Select data from HistProc where IDProcesso = " & TBLISTA_Processos!IDPROCESSO & " order by IDHistorico desc", Conexao, adOpenKeyset, adLockOptimistic
        If TBProduto.EOF = False Then
            .Item(.Count).SubItems(3) = IIf(IsNull(TBProduto!Data), "", Format(TBProduto!Data, "dd/mm/yy"))
        End If
        TBProduto.Close
        
        If IsNull(TBLISTA_Processos!Tipo) = False And TBLISTA_Processos!Tipo <> "" Then
            If TBLISTA_Processos!Tipo = "C" Then .Item(.Count).SubItems(4) = "Custos"
            If TBLISTA_Processos!Tipo = "F" Then .Item(.Count).SubItems(4) = "Componente"
            If TBLISTA_Processos!Tipo = "M" Then .Item(.Count).SubItems(4) = "Subconjunto"
            If TBLISTA_Processos!Tipo = "E" Then .Item(.Count).SubItems(4) = "Produto final"
            If TBLISTA_Processos!Tipo = "S" Then .Item(.Count).SubItems(4) = "Serviço"
        End If
        .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA_Processos!DtImplantacao), "", Format(TBLISTA_Processos!DtImplantacao, "dd/mm/yy"))
        .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA_Processos!Desenho), "", TBLISTA_Processos!Desenho)
        .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA_Processos!Descricao), "", TBLISTA_Processos!Descricao)
        If ProcVerifiProcRevisado(TBLISTA_Processos!Nprocesso, TBLISTA_Processos!Revisao, "", "", False) = True Then .Item(.Count).SubItems(8) = "Revisado" Else .Item(.Count).SubItems(8) = IIf(TBLISTA_Processos!Bloqueado = True, "Bloqueado", "Liberado")
        .Item(.Count).SubItems(9) = IIf(IsNull(TBLISTA_Processos!cronometrado), "", TBLISTA_Processos!cronometrado)
        .Item(.Count).SubItems(10) = IIf(IsNull(TBLISTA_Processos!DtValidacao), "NÃO", "SIM")
    End With
    TBLISTA_Processos.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista.Value = Contador
Loop
lblRegistros.Caption = "Nº de registros: " & TBLISTA_Processos.RecordCount
If TBLISTA_Processos.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "Página: 1 de: " & TBLISTA_Processos.PageCount
ElseIf TBLISTA_Processos.AbsolutePage = adPosEOF Then
        lblPaginas.Caption = "Página: " & TBLISTA_Processos.PageCount & " de: " & TBLISTA_Processos.PageCount
    Else
        lblPaginas.Caption = "Página: " & TBLISTA_Processos.AbsolutePage - 1 & " de: " & TBLISTA_Processos.PageCount
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcPuxaDados()
On Error GoTo tratar_erro

txtidprocesso.Text = TBProcessos!IDPROCESSO
Txt_numero_processo = IIf(IsNull(TBProcessos!Nprocesso), "", TBProcessos!Nprocesso)
txtrevproc = IIf(IsNull(TBProcessos!Revisao), 0, TBProcessos!Revisao)
With cmbtipoprocesso
    If IsNull(TBProcessos!Tipo) = False And TBProcessos!Tipo <> "" Then
        Select Case TBProcessos!Tipo
            Case "C":
                .Clear
                .AddItem "Componente"
                .AddItem "Subconjunto"
                .AddItem "Produto final"
                .AddItem "Custos"
                .Text = "Custos"
            Case "F": .Text = "Componente"
            Case "M": .Text = "Subconjunto"
            Case "E": .Text = "Produto final"
            Case "S": .Text = "Serviço"
        End Select
        .Locked = True
        .TabStop = False
    Else
        .Locked = False
        .TabStop = True
        .ListIndex = -1
    End If
End With
If TBProcessos!Contador <> "" Then txtrevdesenho.Text = TBProcessos!Contador

Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select rastreavel, codproduto, desenho,Unidade, descricao from projproduto where codproduto = " & TBProcessos!Codproduto, Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    Caption = "Engenharia - Processos - Gerenciamento - (Processo : " & TBProcessos!IDPROCESSO & " - Cód. interno : " & TBProduto!Desenho & ")"
    txtdesenho.Text = TBProduto!Desenho
    txtProduto.Text = IIf(IsNull(TBProduto!Descricao), "", (TBProduto!Descricao))
    txtUN.Text = IIf(IsNull(TBProduto!Unidade), "", (TBProduto!Unidade))
    Label1(21).Caption = "Total x " & txtUN.Text & " :"
    Label1(22).Caption = "Segundos x " & txtUN.Text & " :"
    Label1(23).Caption = txtUN.Text & " x hora : "
    Label1(28).Caption = txtUN.Text & " x execução"
    Label1(29).Caption = "Execução x " & txtUN.Text
    Processo_Rastreavel = IIf(IsNull(TBProduto!rastreavel), 0, TBProduto!rastreavel)
End If


StrSql = "UPDATE fases SET rastreavel = '" & IIf(TBProduto!rastreavel = True, 1, 0) & "' WHERE rastreavel is null AND IDProcesso = '" & TBProcessos!IDPROCESSO & "'"


'Debug.print StrSql

Conexao.Execute (StrSql)
TBProduto.Close

If ProcVerifiProcRevisado(TBProcessos!Nprocesso, TBProcessos!Revisao, "", "", False) = True Then txtStatus = "Revisado" Else txtStatus = IIf(TBProcessos!Bloqueado = True, "Bloqueado", "Liberado")

txtElaborado.Text = IIf(IsNull(TBProcessos!elaborado), "", (Format(TBProcessos!elaborado, "dd/mm/yy")))
txtDtImplantacao.Text = IIf(IsNull(TBProcessos!DtImplantacao), "", (Format(TBProcessos!DtImplantacao, "dd/mm/yy")))
txtDtValidacao.Text = IIf(IsNull(TBProcessos!DtValidacao), "", TBProcessos!DtValidacao)
txtRespValidacao.Text = IIf(IsNull(TBProcessos!RespValidacao), "", TBProcessos!RespValidacao)

Novo_Processo = False
ProcLimparTudo

ProcAtualizaFases
Frame1.Enabled = True
txtreferencia.Enabled = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

Formulario = "Engenharia/Processos"
Direitos
ProcLimpaVariaveisPrincipais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAtualizacao()
On Error GoTo tratar_erro

If USMsgBox("Deseja realmente executar essa(s) atualização(ões)?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    With frmProcessos_atualizar
        If .Chk1.Value = 1 Then
            Set TBFases = CreateObject("adodb.recordset")
            TBFases.Open "Select * from fases where idprocesso = " & txtidprocesso & " order by grupo_op", Conexao, adOpenKeyset, adLockOptimistic
            If TBFases.EOF = False Then
                PBLista.Min = 0
                PBLista.Max = TBFases.RecordCount
                PBLista.Value = 1
                Contador = 0
                Do While TBFases.EOF = False
                    If TBFases!Grupo_op <> "" Then
                        Set TBplano = CreateObject("adodb.recordset")
                        TBplano.Open "Select * from plano where desenho = '" & txtdesenho & "' and fase = " & TBFases!Grupo_op, Conexao, adOpenKeyset, adLockOptimistic
                        If TBplano.EOF = False Then
                            TBplano!Fase = TBFases!Fase
                            TBplano!Descricao = txtProduto
                            TBplano.Update
                        End If
                        TBplano.Close
                    End If
                    TBFases.MoveNext
                    Contador = Contador + 1
                    PBLista.Value = Contador
                Loop
            End If
            TBFases.Close
            
            Set TBFases = CreateObject("adodb.recordset")
            TBFases.Open "Select * from fases where idprocesso = " & txtidprocesso & " order by fase", Conexao, adOpenKeyset, adLockOptimistic
            If TBFases.EOF = False Then
                PBLista.Min = 0
                PBLista.Max = TBFases.RecordCount
                PBLista.Value = 1
                Contador = 0
                Do While TBFases.EOF = False
                    If TBFases!Grupo_op <> "" Then
                        Set TBplano = CreateObject("adodb.recordset")
                        TBplano.Open "Select * from plano where desenho = '" & txtdesenho & "' and fase = " & TBFases!Fase, Conexao, adOpenKeyset, adLockOptimistic
                        If TBplano.EOF = False Then
                            TBplano!Grupo_op = TBFases!Grupo_op
                            TBplano!Descricao = txtProduto
                            TBplano.Update
                        End If
                        TBplano.Close
                    End If
                    TBFases.MoveNext
                    Contador = Contador + 1
                    PBLista.Value = Contador
                Loop
            End If
            TBFases.Close
           
            'Exclui o plano mais antigo qdo tiver dois planos iguais
Inicio:
            Set TBplano = CreateObject("adodb.recordset")
            TBplano.Open "Select * from plano where desenho = '" & txtdesenho & "' order by fase", Conexao, adOpenKeyset, adLockOptimistic
            If TBplano.EOF = False Then
                TBplano.MoveLast
                PBLista.Min = 0
                PBLista.Max = TBplano.RecordCount
                PBLista.Value = 1
                Contador = 0
                TBplano.MoveFirst
                Do While TBplano.EOF = False
                    Set TBAbrir = CreateObject("adodb.recordset")
                    TBAbrir.Open "Select * from plano where idplano <> " & TBplano!IDPlano & " and desenho = '" & txtdesenho & "' and fase = " & TBplano!Fase & " order by fase", Conexao, adOpenKeyset, adLockOptimistic
                    If TBAbrir.EOF = False Then
                        Dataini = Format(TBplano!Data, "dd/mm/yy")
                        DataFim = Format(TBAbrir!Data, "dd/mm/yy")
                        If Dataini > DataFim Then
                            Conexao.Execute "DELETE from planodimensao where idplano = " & TBAbrir!IDPlano
                            Conexao.Execute "DELETE from plano where idplano = " & TBAbrir!IDPlano
                            TBplano.Close
                            GoTo Inicio
                        Else
                            Conexao.Execute "DELETE from planodimensao where idplano = " & TBplano!IDPlano
                            Conexao.Execute "DELETE from plano where idplano = " & TBplano!IDPlano
                            TBplano.Close
                            GoTo Inicio
                        End If
                    End If
                    TBAbrir.Close
                    TBplano.MoveNext
                    Contador = Contador + 1
                    PBLista.Value = Contador
                Loop
            End If
            TBplano.Close
            USMsgBox ("Atualização efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
            '==================================
            Modulo = "Engenharia/Processos"
            Evento = "Atualizar plano insp."
            ID_documento = 0
            Documento = ""
            Documento1 = ""
            ProcGravaEvento
            '=================================
        End If
        
        If .Chk2 Then
            Set TBProcessos = CreateObject("adodb.recordset")
            TBProcessos.Open "Select codproduto, tipo from processos order by idprocesso", Conexao, adOpenKeyset, adLockOptimistic
            If TBProcessos.EOF = False Then
                TBProcessos.MoveLast
                PBLista.Min = 0
                PBLista.Max = TBProcessos.RecordCount
                PBLista.Value = 1
                Contador = 0
                TBProcessos.MoveFirst
                Do While TBProcessos.EOF = False
                    Set TBProduto = CreateObject("adodb.recordset")
                    TBProduto.Open "Select * from projproduto where codproduto = " & TBProcessos!Codproduto & " and SubTipoItem is not null and SubTipoItem <> 0 and SubTipoItem <> 4", Conexao, adOpenKeyset, adLockOptimistic
                    If TBProduto.EOF = False Then
                        Select Case TBProduto!SubTipoItem
                            Case 1: TBProcessos!Tipo = "C"
                            Case 2: TBProcessos!Tipo = "S"
                            Case 3: TBProcessos!Tipo = "P"
                        End Select
                    End If
                    TBProduto.Close
                    TBProcessos.Update
                    TBProcessos.MoveNext
                    Contador = Contador + 1
                    PBLista.Value = Contador
                Loop
                USMsgBox ("Atualização efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
                '==================================
                Modulo = "Engenharia/Processos"
                Evento = "Atualizar tipo"
                ID_documento = 0
                Documento = ""
                Documento1 = ""
                ProcGravaEvento
                '==================================
            End If
            TBProcessos.Close
        End If
        
        If .Chk3 Then
            Set TBProcessos = CreateObject("adodb.recordset")
            TBProcessos.Open "Select * from processos order by idprocesso", Conexao, adOpenKeyset, adLockOptimistic
            If TBProcessos.EOF = False Then
                TBProcessos.MoveLast
                PBLista.Min = 0
                PBLista.Max = TBProcessos.RecordCount
                PBLista.Value = 1
                Contador = 0
                TBProcessos.MoveFirst
                Do While TBProcessos.EOF = False
                    DecimoSegundos = 0
                    TotalFaseSeg = 0
                    CustoFase = 0
                    CustohoraSeg = 0
                    CustoProcesso = 0
                    CustoTotalPrep = 0
                    'Localiza fases do processo para atualizar custo
                    Set TBFases = CreateObject("adodb.recordset")
                    TBFases.Open "Select * FROM Fases WHERE IDProcesso = " & TBProcessos!IDPROCESSO & " order by fase", Conexao, adOpenKeyset, adLockOptimistic
                    If TBFases.EOF = False Then
                        Do While TBFases.EOF = False
                            Set TBMaquinas = CreateObject("adodb.recordset")
                            TBMaquinas.Open "Select * FROM CadMaquinas WHERE Maquina = '" & TBFases("maquina") & "'", Conexao, adOpenKeyset, adLockOptimistic
                            If TBMaquinas.EOF = False Then
                                Fase = Int(TBFases!Fase)
                                If Len(Fase) = 2 Then Fase = "00" & Fase
                                If Len(Fase) = 3 Then Fase = "0" & Fase
                                If Len(Fase) = 4 Then Fase = Fase
                                TBFases!Fase = Fase
                                If IsNull(TBFases!pc_te) = True Or TBFases!pc_te = "" Then TBFases!pc_te = 1
                                
                                'Calcula e grava tempo de preparação
                                ElapsedTime (IIf(IsNull(TBFases!Preparacao), 0, TBFases!Preparacao))
                                TBFases!TPSegundos = s
                                TBFases!TempoPreparacao = HoraTotal
                                
                                'Calcula e grava tempo de execução
                                If IsNull(TBFases!TESegundos) = True Or TBFases!TESegundos = "" Or TBFases!TESegundos = 0 Then TBFases!TESegundos = FunCalculaSegPC(IIf(IsNull(TBFases!Execucao), 0, TBFases!Execucao), TBFases!pc_te)
                                TBFases!TempoExecucao = FormataTempo(IIf(IsNull(TBFases!TESegundos), 0, TBFases!TESegundos))
                                
                                'Calcula custo fase
                                TotalFaseSeg = IIf(IsNull(TBFases!TESegundos), 0, TBFases!TESegundos)
                                CustohoraSeg = TBMaquinas!PrecoHora / 3600
                                CustoFase = CustohoraSeg * TotalFaseSeg
                                
                                'Soma total de segundos no processo
                                DecimoSegundos = DecimoSegundos + (IIf(IsNull(TBFases!TESegundos), 0, TBFases!TESegundos))
                                CustoProcesso = CustoProcesso + CustoFase
                                                            
                                TBFases!Custo = CustoFase
                                
                                If IsNull(TBMaquinas!PrecoHora_Setup) = False And TBMaquinas!PrecoHora_Setup <> "" Then CustohoraSeg = TBMaquinas!PrecoHora_Setup / 3600
                                ElapsedTime (IIf(IsNull(TBFases!Preparacao), 0, TBFases!Preparacao))
                                TBFases!Custoprep = CustohoraSeg * s
                                TBFases.Update
                                CustoTotalPrep = CustoTotalPrep + IIf(IsNull(TBFases!Custoprep), 0, TBFases!Custoprep)
                            End If
                            TBFases.MoveNext
                        Loop
                    End If
                    'Atualiza Tempo total e custo total do processo
                    TBProcessos!Custo = CustoProcesso
                    TBProcessos!Custoprep = CustoTotalPrep
                    TBProcessos!TTotalSEG = DecimoSegundos
                    TBProcessos!TTotal = FormataTempo(DecimoSegundos)
                    TBProcessos.Update
                    TBProcessos.MoveNext
                    Contador = Contador + 1
                    PBLista.Value = Contador
                Loop
            End If
            TBProcessos.Close
            USMsgBox ("Atualização efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
            '==================================
            Modulo = "Engenharia/Processos"
            Evento = "Atualizar fases"
            ID_documento = 0
            Documento = ""
            Documento1 = ""
            ProcGravaEvento
            '==================================
        End If
    End With
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procCopiarProcesso()
On Error GoTo tratar_erro

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If txtidprocesso = "" Then
    USMsgBox ("Informe o processo antes de copiar."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
frmprocessos_copiar.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro

If txtidprocesso = "" Then
    USMsgBox ("Informe o processo antes de visualizar impressão."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
NomeRel = "processos.rpt"
ProcImprimirRel "{processos.idprocesso}= " & txtidprocesso, ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcAtualizaCustoProcesso()
On Error GoTo tratar_erro

TempoTotalProcesso = 0
TotalFaseSeg = 0
CustoFase = 0
CustohoraSeg = 0
CustoProcesso = 0
CustoTotalPrep = 0
TotalSegundos = 0

'Abre processo para atualizar custo
Set TBProcessos = CreateObject("adodb.recordset")
TBProcessos.Open "Select * from processos where idprocesso = " & txtidprocesso, Conexao, adOpenKeyset, adLockOptimistic
If TBProcessos.EOF = False Then
    'Localiza fases do processo para atualizar custo
    Set TBFases = CreateObject("adodb.recordset")
    TBFases.Open "Select * FROM Fases WHERE IDProcesso = " & txtidprocesso & " order by fase", Conexao, adOpenKeyset, adLockOptimistic
    If TBFases.EOF = False Then
        Do While TBFases.EOF = False
            'Busca custo hora da maquina
            Set TBMaquinas = CreateObject("adodb.recordset")
            TBMaquinas.Open "Select * FROM CadMaquinas WHERE Maquina = '" & TBFases("maquina") & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBMaquinas.EOF = False Then
                ElapsedTime (IIf(IsNull(TBFases!Execucao), 0, TBFases!Execucao))
                PcHora = TBFases!pc_te
                If PcHora <> 0 Then TotalFaseSeg = s / PcHora
                TotalSegundos = TotalSegundos + s
                
                CustohoraSeg = TBMaquinas!PrecoHora / 3600
                CustoFase = CustohoraSeg * TotalFaseSeg
                TempoTotalProcesso = TempoTotalProcesso + IIf(IsNull(TBFases!Execucao), 0, TBFases!Execucao)
                CustoProcesso = CustoProcesso + CustoFase
                TBFases!Custo = CustoFase
                
                If IsNull(TBMaquinas!PrecoHora_Setup) = False And TBMaquinas!PrecoHora_Setup <> "" Then CustohoraSeg = TBMaquinas!PrecoHora_Setup / 3600
                ElapsedTime (IIf(IsNull(TBFases!Preparacao), 0, TBFases!Preparacao))
                TBFases!Custoprep = CustohoraSeg * s
                TBFases.Update
                CustoTotalPrep = CustoTotalPrep + TBFases!Custoprep
            End If
            TBFases.MoveNext
        Loop
    End If
    'Atualiza Tempo total e custo total do processo
    TBProcessos!Custo = CustoProcesso
    TBProcessos!Custoprep = CustoTotalPrep
    
    'Preças por hora
    TBProcessos!PcHora = IIf(txtA5.Text <> "", txtA5.Text, 0)
    
   ' ProcFormataHora (IIf(TxtTempoProcesso.Text <> "", TxtTempoProcesso, 0))
    TBProcessos!TTotalSEG = s + DecimoSegundos
    TBProcessos!TTotal = FormataTempo(TBProcessos!TTotalSEG)

    TBProcessos.Update
End If
TBProcessos.Close
                
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
With ListaProcessos
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) processo(s)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            
            'Exclui dados do plano de inspeção
            Conexao.Execute "DELETE from PDI from ((plano P INNER JOIN Fases F ON F.IDFase = P.IDFase) INNER JOIN Planodimensao PD ON PD.IdPlano = P.IdPlano) INNER JOIN Planodimensao_instrumentos PDI ON PDI.ID_dimensao = PD.idDimensao where F.IDProcesso = " & .ListItems(InitFor)
            Conexao.Execute "DELETE from PD from (plano P INNER JOIN Fases F ON F.IDFase = P.IDFase) INNER JOIN Planodimensao PD ON PD.IdPlano = P.IdPlano where F.IDProcesso = " & .ListItems(InitFor)
            Conexao.Execute "DELETE from PR from (plano P INNER JOIN Fases F ON F.IDFase = P.IDFase) INNER JOIN Plano_revisao PR ON PR.IdPlano = P.IdPlano where F.IDProcesso = " & .ListItems(InitFor)
            Conexao.Execute "DELETE from P from plano P INNER JOIN Fases F ON F.IDFase = P.IDFase where F.IDProcesso = " & .ListItems(InitFor)
            
            Conexao.Execute "DELETE from H from HISTPROC H INNER JOIN Processos PRO ON PRO.IDprocesso = H.IDprocesso where PRO.NProcesso = '" & .ListItems(InitFor).ListSubItems(1) & "' and PRO.Revisao = " & .ListItems(InitFor).ListSubItems(2) - 1
            Conexao.Execute "DELETE from Processos where IDProcesso = " & .ListItems(InitFor)
            Conexao.Execute "DELETE from HISTPROC where IDProcesso = " & .ListItems(InitFor)
            Conexao.Execute "DELETE from fases where IDProcesso = " & .ListItems(InitFor)
            Conexao.Execute "DELETE from fases_revisao where IDProcesso = " & .ListItems(InitFor)
            Conexao.Execute "DELETE from Ferramentas where IDProcesso = " & .ListItems(InitFor)
            Conexao.Execute "DELETE from programas where IDProcesso = " & .ListItems(InitFor)
            
            
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select P.Codproduto from Processos PR INNER JOIN Projproduto P ON PR.Codproduto = P.Codproduto where P.Desenho = '" & .ListItems(InitFor).SubItems(6) & "' and PR.Tipo <> 'C'", Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                TemProcessoTexto = "Processo = 'True'"
            Else
                TemProcessoTexto = "Processo = 'False'"
            End If
            TBAbrir.Close
            Conexao.Execute "Update projproduto Set " & TemProcessoTexto & " where Desenho = '" & .ListItems(InitFor).SubItems(6) & "'"
            
            '==================================
            Modulo = "Engenharia/Processos"
            Evento = "Excluir"
            ID_documento = .ListItems(InitFor)
            Documento = "Processo: " & .ListItems(InitFor).SubItems(1) & " - Rev.: " & .ListItems(InitFor).SubItems(2) & " - Cód. interno: " & .ListItems(InitFor).SubItems(6)
            Documento1 = ""
            ProcGravaEvento
            '==================================
            
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) processo(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Processo(s) excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpar
    ProcCarregaListaProcessos (1)
    ProcLimparTudo
    Frame1.Enabled = False
    txtreferencia.Enabled = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLocalizar()
On Error GoTo tratar_erro
  
frmprocessos_Abrir.Show 1

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
ProcLimpar
ProcLimparFase
Sit_REG = 1
frmprocessos_Novo.Show 1

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
If txtdesenho.Text = "" Then
    ProcVerificaSalvar
    Exit Sub
End If
If cmbtipoprocesso.Text = "" Then
    USMsgBox ("Informe o tipo do processo antes de salvar."), vbExclamation, "CAPRIND v5.0"
    cmbtipoprocesso.SetFocus
    Exit Sub
End If

Set TBProcessos = CreateObject("adodb.recordset")
TBProcessos.Open "Select * FROM Processos WHERE idprocesso = " & txtidprocesso, Conexao, adOpenKeyset, adLockOptimistic
If TBProcessos.EOF = True Then
    TBProcessos.AddNew
    TBProcessos!cronometrado = "NÃO"
    TBProcessos!Bloqueado = False
    TBProcessos!Nprocesso = FunCriarNovoNumeroProcesso
Else
    If FunVerifValidacaoRegistro("alterar", txtDtValidacao, "mesmo", "o processo", True) = False Then Exit Sub
End If
TBProcessos!Revisao = IIf(txtrevproc.Text <> "", txtrevproc, 0)
Select Case cmbtipoprocesso.Text
    Case "Custos": TBProcessos!Tipo = "C"
    Case "Componente": TBProcessos!Tipo = "F"
    Case "Subconjunto": TBProcessos!Tipo = "M"
    Case "Produto final": TBProcessos!Tipo = "E"
    Case "Serviço": TBProcessos!Tipo = "S"
End Select

If txtDtImplantacao = "" Then TBProcessos!DtImplantacao = Date Else TBProcessos!DtImplantacao = txtDtImplantacao
If txtElaborado = "" Then TBProcessos!elaborado = pubUsuario Else TBProcessos!elaborado = txtElaborado

Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select Codproduto, RevDesenho, Processo from projproduto where Desenho = '" & txtdesenho & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    Codproduto = TBProduto!Codproduto
    TBProcessos!Contador = TBProduto!RevDesenho
    TBProcessos!Codproduto = TBProduto!Codproduto
    TBProduto!Processo = True
    TBProduto.Update
End If
TBProduto.Close

TBProcessos.Update
txtidprocesso = TBProcessos!IDPROCESSO
If Novo_Processo = True Then Conexao.Execute "Update Processos set ordenarprocesso = " & TBProcessos!IDPROCESSO & " where IDprocesso = " & TBProcessos!IDPROCESSO
TBProcessos.Close
If Novo_Processo = True Then
    USMsgBox ("Novo processo de " & cmbtipoprocesso & " cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo"
    Sql_Processo_Localizar = "Select PRO.IDprocesso, PRO.Nprocesso, PRO.Revisao, PRO.Tipo, PRO.DtImplantacao, PRO.Bloqueado, PRO.cronometrado, PRO.DtValidacao, PRO.ordenarprocesso, P.Desenho, P.Descricao from Processos PRO INNER JOIN projproduto P ON P.codproduto = PRO.CodProduto where PRO.idprocesso = " & txtidprocesso
    ProcCarregaListaProcessos (1)
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar"
    ProcCarregaListaProcessos (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
    If CodigoLista <> 0 And ListaProcessos.ListItems.Count <> 0 Then
        ListaProcessos.SelectedItem = ListaProcessos.ListItems(CodigoLista)
        ListaProcessos.SetFocus
    End If
End If
'==================================
Modulo = "Engenharia/Processos"
ID_documento = txtidprocesso.Text
Documento = "Processo: " & txtidprocesso & " - Rev.: " & txtrevproc & " - Cód. interno: " & txtdesenho & " - Rev.: " & txtrevdesenho
Documento1 = ""
ProcGravaEvento
'==================================
Novo_Processo = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaFases_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro
Dim Prep As Variant 'OK
Dim Exec As Variant 'OK

If ListaFases.ListItems.Count = 0 Then Exit Sub
ProcLimparFase
Set TBFases = CreateObject("adodb.recordset")
TBFases.Open "Select * FROM Fases WHERE IDFase = " & ListaFases.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBFases.EOF = False Then
    ProcPuxaDadosFase
    CodigoLista1 = ListaFases.SelectedItem.index
End If
TBFases.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcPuxaDadosFase()
On Error GoTo tratar_erro

Txt_ID_fase = TBFases!IDFase
If TBFases!versao <> "" Then cmbVersao.Text = TBFases!versao

If TBFases!pecahora = False Then chkPchora.Value = 0 Else chkPchora.Value = 1
txtFase.Text = TBFases!Fase

Txt_rev_fase = IIf(IsNull(TBFases!Revisao), "", TBFases!Revisao)
If IsNull(TBFases!Revisao) = False And TBFases!Revisao <> "" Then
    'Verifica data da revisão
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select Data from Fases_revisao where IDFase = " & TBFases!IDFase & " and Revisao = '" & TBFases!Revisao & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Txt_data_rev = Format(TBAbrir!Data, "dd/mm/yy")
    End If
    TBAbrir.Close
End If

txtgrupo_op = IIf(IsNull(TBFases!Grupo_op), "", TBFases!Grupo_op)

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select Grupo from CadMaquinas where Maquina = '" & TBFases!maquina & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    Cmb_grupo_posto.Text = TBAbrir!Grupo
End If

NomeCampo = "a máquina"
If TBFases!maquina <> "" Then cmbMaquina.Text = TBFases!maquina
1:
    Novo_Processo1 = False
    Framedetalhes.Enabled = True
    txtdescricao.TextRTF = IIf(IsNull(TBFases!Descricao), "", TBFases!Descricao)
    txt_Caminho = IIf(IsNull(TBFases!caminho), "", TBFases!caminho)
    If TBFases!cronometrado = True Then chkCronometrado.Value = 1 Else chkCronometrado.Value = 0
    If TBFases!Plano_inspecao = True Then Chk_tem_plano.Value = 1 Else Chk_tem_plano.Value = 0
    If TBFases!Nao_aponta = True Then chk_N_apontamento.Value = 1 Else chk_N_apontamento.Value = 0
    If TBFases!rastreavel = True Then chkRastreavel = 1 Else chkRastreavel.Value = 0
    
    Procentrada
    If TempoPreparacao = "" Then txtpreparacao.Text = "___:__:__" Else txtpreparacao.Text = TempoPreparacao
    If TempoExecucao = "" Then txtexecucao.Text = "___:__:__" Else txtexecucao.Text = TempoExecucao
    FaseNovo_Processo = TBFases!IDFase
    txtPcHora.Text = Format(IIf(IsNull(TBFases!pc_te) = False, TBFases!pc_te, 0), "###,##0.0000")
    
    TxtA3 = FunCalculaSegPC(TBFases!Execucao, txtPcHora)
    TxtA3 = FormataTempo(TxtA3.Text)
    If TBFases!Plano_montagem = True Then chkPlano_montagem.Value = 1 Else chkPlano_montagem.Value = 0

Exit Sub
tratar_erro:
    If Err.Number = "383" Then
        USMsgBox ("Não foi encontrado " & NomeCampo & " desta fase."), vbExclamation, "CAPRIND v5.0"
        GoTo 1
    End If
    Exit Sub
End Sub

Sub ProcCopiarFase(TextoFiltro As String, CopiarProcesso As Boolean, IDPROCESSO As Long, Desenho As String, Descricao As String, Revisar As Boolean)
On Error GoTo tratar_erro

Set TBFases = CreateObject("adodb.recordset")
TBFases.Open TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
If TBFases.EOF = False Then
    Do While TBFases.EOF = False
        IDFase = TBFases!IDFase
        Set TBMaquinas = CreateObject("adodb.recordset")
        TBMaquinas.Open "Select * from fases", Conexao, adOpenKeyset, adLockOptimistic
        TBMaquinas.AddNew
        TBMaquinas!IDPROCESSO = IDPROCESSO
        TBMaquinas!Fase = TBFases!Fase
        TBMaquinas!Grupo_op = TBFases!Grupo_op
        TBMaquinas!maquina = TBFases!maquina
        TBMaquinas!Preparacao = TBFases!Preparacao
        TBMaquinas!Execucao = TBFases!Execucao
        TBMaquinas!Descricao = TBFases!Descricao
        TBMaquinas!TESegundos = IIf(IsNull(TBFases!TESegundos), 0, TBFases!TESegundos)
        TBMaquinas!TempoPreparacao = TBFases!TempoPreparacao
        TBMaquinas!TempoExecucao = TBFases!TempoExecucao
        TBMaquinas!Custoprep = TBFases!Custoprep
        If CopiarProcesso = True Then TBMaquinas!versao = TBFases!versao Else TBMaquinas!versao = frmProcessos_copiar_fase.Versao_fase
        TBMaquinas!pc_te = IIf(IsNull(TBFases!pc_te), 0, TBFases!pc_te)
        TBMaquinas!TPSegundos = TBFases!TPSegundos
        TBMaquinas!Revisao = 0
        TBMaquinas!pecahora = TBFases!pecahora
        TBMaquinas!cronometrado = False
        TBMaquinas.Update
        IDlista = TBMaquinas!IDFase
        
        Set TBFerramenta = CreateObject("adodb.recordset")
        TBFerramenta.Open "Select * from ferramentas where idprocesso = " & txtidprocesso.Text & " and idfase = " & IDFase, Conexao, adOpenKeyset, adLockOptimistic
        If TBFerramenta.EOF = False Then
            Set TBFerramentas = CreateObject("adodb.recordset")
            TBFerramentas.Open "Select * from ferramentas", Conexao, adOpenKeyset, adLockOptimistic
            Do While TBFerramenta.EOF = False
                TBFerramentas.AddNew
                TBFerramentas!IDPROCESSO = IDPROCESSO
                TBFerramentas!IDFase = TBMaquinas!IDFase
                TBFerramentas!Numero = TBFerramenta!Numero
                TBFerramentas!quantidade = TBFerramenta!quantidade
                TBFerramentas!ID_acessorio = TBFerramenta!ID_acessorio
                TBFerramentas.Update
                TBFerramenta.MoveNext
            Loop
            TBFerramentas.Close
        End If
        TBFerramenta.Close
        
        Set TBPrograma = CreateObject("adodb.recordset")
        TBPrograma.Open "Select * from programas where idprocesso = " & txtidprocesso.Text & " and idfase = " & IDFase, Conexao, adOpenKeyset, adLockOptimistic
        If TBPrograma.EOF = False Then
            Set TBProgramas = CreateObject("adodb.recordset")
            TBProgramas.Open "Select * from programas", Conexao, adOpenKeyset, adLockOptimistic
            Do While TBPrograma.EOF = False
                TBProgramas.AddNew
                TBProgramas!IDPROCESSO = IDPROCESSO
                TBProgramas!IDFase = TBMaquinas!IDFase
                TBProgramas!Ciclo = TBPrograma!Ciclo
                TBProgramas!Descricao = TBPrograma!Descricao
                TBProgramas!programa = TBPrograma!programa
                TBProgramas!maquina = TBPrograma!maquina
                TBProgramas!Desenho = TBPrograma!Desenho
                TBProgramas!Data = TBPrograma!Data
                TBProgramas!Programacao = TBPrograma!Programacao
                TBProgramas.Update
                TBPrograma.MoveNext
            Loop
            TBProgramas.Close
        End If
        TBPrograma.Close
        
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select P.nivel, PD.* FROM plano P INNER JOIN Planodimensao PD ON P.idplano = PD.idplano where P.IDfase = " & IDFase, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            If Permitido1 = False And Revisar = False Then
                If USMsgBox("Deseja copiar o plano de inspeção da(s) fase(s)?", vbYesNo, "CAPRIND v5.0") = vbYes Then Permitido1 = True
            End If
            If Permitido1 = True Or Revisar = True Then ProcCopiarPlanoInspecao Desenho, Descricao, TBMaquinas!IDFase, TBFases!Fase, TBFases!Grupo_op
        End If
        TBAbrir.Close
        TBFases.MoveNext
    Loop
End If
TBFases.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCalculaA5()
On Error GoTo tratar_erro

Valor1 = 0
Valor2 = 0
Valor3 = 0
If IsNumeric(txtA4.Text) = True Then
    If txtA4 <> 0 Then
        Valor1 = txtA4
        Valor2 = 3600
        Valor3 = Valor2 / Valor1
    End If
End If
txtA5.Text = Format(Valor3, "###,##0.0000000000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCalculaA6()
On Error GoTo tratar_erro

If IsNumeric(txtA4.Text) = True Then
    txtA6.Text = FormataTempo(txtA4.Text)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcSair()
On Error GoTo tratar_erro
  
If Novo_Processo = True Then
    If USMsgBox("O processo ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvar
        If Novo_Processo = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
If Novo_Processo1 = True Then
    If USMsgBox("A fase ainda não foi salva, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        procSalvar_fase
        If Novo_Processo1 = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
Novo_Processo = False
Novo_Processo1 = False
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcAbrir()
On Error GoTo tratar_erro

ProcLimpar
Set TBProcessos = CreateObject("adodb.recordset")
TBProcessos.Open "Select * from Processos order by IDProcesso", Conexao, adOpenKeyset, adLockOptimistic
If TBProcessos.EOF = False Then
    TBProcessos.MoveLast
    'Set TBProduto = CreateObject("adodb.recordset")
    'TBProduto.Open "Select * from projproduto where codproduto = " & TBProcessos!codProduto , Conexao, adOpenKeyset, adLockOptimistic
    'If TBProduto.EOF = False then
        'codProduto = TBProduto!codProduto
        'If TBProduto!Revdesenho <> TBProcessos!Contador Then
            'If usMsgbox("Revisão do código interno do processo ( " & TBProcessos!Contador & " ) em desacordo com revisão de desenho do projeto ( " & TBProduto!Revdesenho & " ), deseja atualizar automaticamento a revisão no processo?.", vbyesno, "CAPRIND v5.0") = vbYes Then
                'TBProcessos!Contador = TBProduto!Revdesenho
                'TBProcessos.Update
                'usMsgbox "Revisão do código interno atualizada com sucesso.", vbInformation, "CAPRIND v5.0"
            'Else
                'usMsgbox "Revise o processo e atualize a revisão do código interno na próxima abertura do processo.", vbInformation, "CAPRIND v5.0"
            'End If
        'End If
        'TBProduto.Close
    'End If
    txtidprocesso = TBProcessos!IDPROCESSO
    Set TBHistProc = CreateObject("adodb.recordset")
    TBHistProc.Open "Select * FROM histproc WHERE idprocesso = " & txtidprocesso, Conexao, adOpenKeyset, adLockOptimistic
    If TBHistProc.EOF = False Then
        TBHistProc.MoveLast
        txtrevproc.Text = ""
        txtrevproc.Text = TBHistProc!Ordem
        TBHistProc.Close
    Else
        txtrevproc.Text = 0
    End If
    ProcPuxaDados
    ProcVerificaItem
End If
    
Exit Sub
tratar_erro:
   USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpar()
On Error GoTo tratar_erro
    
Codproduto = ""
txtidprocesso.Text = ""
Txt_numero_processo = ""
cmbtipoprocesso.ListIndex = -1
txtrevproc.Text = ""
txtElaborado.Text = ""
txtDtImplantacao.Text = Format(Date, "dd/mm/yy")
txtdesenho.Text = ""
txtrevdesenho.Text = ""
txtreferencia.Clear
txtStatus.Text = ""
txtProduto.Text = ""
txtDtValidacao = ""
txtRespValidacao = ""
CodigoLista = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimparFase()
On Error GoTo tratar_erro
    
Txt_ID_fase = 0
txtFase.Text = ""
Txt_rev_fase = 0
Txt_data_rev = ""
txtgrupo_op.Text = ""
Cmb_grupo_posto.ListIndex = -1
cmbMaquina.ListIndex = -1
txtdescmaquina.Text = ""
chkPchora.Value = 0
chkCronometrado.Value = 0
Chk_tem_plano.Value = 0
chkRastreavel.Value = 0

chk_N_apontamento.Value = 0
txtpreparacao.Text = "000:00:00"
txtexecucao.Text = "000:00:00"
txtPcHora = 1
TxtA3 = "00:00:00"
txtdescricao.Text = ""
txt_Caminho = ""
chkPlano_montagem.Value = 0
CodigoLista1 = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcVerificaItem()
On Error GoTo tratar_erro

Desenho = txtdesenho.Text
If Desenho = "" Then Exit Sub
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select * from projproduto where desenho = '" & Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    Set TBProcessos = CreateObject("adodb.recordset")
    TBProcessos.Open "Select * from processos where codproduto = " & TBProduto!Codproduto, Conexao, adOpenKeyset, adLockOptimistic
    TBProcessos!Codproduto = TBProduto!Codproduto
    Codproduto = TBProduto!Codproduto
    TBProcessos.Update
    TBProduto.Close
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub txtDescricao_Change()
On Error GoTo tratar_erro

With txtdescricao
    Cor_fonte.BackColor = IIf(IsNull(.SelColor), Cor_fonte.BackColor, .SelColor)
    Chk_negrito.Value = IIf(IsNull(.SelBold), 2, Abs(.SelBold))
    Chk_italico.Value = IIf(IsNull(.SelItalic), 2, Abs(.SelItalic))
    Chk_sublinhado.Value = IIf(IsNull(.SelUnderline), 2, Abs(.SelUnderline))
    Cmb_fonte = IIf(IsNull(.SelFontName), "", .SelFontName)
1:
    Cmb_tamanho_fonte = IIf(IsNull(.SelFontSize), "", .SelFontSize)
End With

Exit Sub
tratar_erro:
    If Err.Number = 383 Then
        Cmb_tamanho_fonte.AddItem IIf(IsNull(txtdescricao.SelFontSize), "", txtdescricao.SelFontSize)
        GoTo 1
    End If
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtDesenho_Change()
On Error GoTo tratar_erro

txtreferencia.Clear
If txtdesenho <> "" Then
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from item_aplicacoes where codproduto = " & TBProduto!Codproduto, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Do While TBAbrir.EOF = False
            txtreferencia.AddItem TBAbrir!N_referencia
            TBAbrir.MoveNext
        Loop
    End If
    TBAbrir.Close
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtexecucao_Change()
On Error GoTo tratar_erro

ProcCalculaExecucaoPeca

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCalculaExecucaoPeca()
On Error GoTo tratar_erro

txtpreparacao.PromptInclude = True
txtexecucao.PromptInclude = False
If Len(txtexecucao.Text) < 7 Then
    txtexecucao.PromptInclude = True
    Exit Sub
End If
txtexecucao.PromptInclude = True
If txtexecucao > "023:59:59" Then
    ProcFormataHora (txtexecucao)
    Familiatext = DataResultado
    TxtA3 = FunCalculaSegPC(Familiatext, IIf(txtPcHora = "", 0, txtPcHora))
Else
    TxtA3 = FunCalculaSegPC(txtexecucao, IIf(txtPcHora = "", 0, txtPcHora))
End If
TxtA3 = FormataTempo(TxtA3.Text)
txtexecucao.PromptInclude = False

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

Private Sub txtPcHora_Change()
On Error GoTo tratar_erro

If txtPcHora.Text <> "" Then
    VerifNumero = txtPcHora.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtPcHora.Text = ""
        txtPcHora.SetFocus
        Exit Sub
    End If
End If
ProcCalculaExecucaoPeca
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaVersao()
On Error GoTo tratar_erro

With cmbVersao_pesquisar
    Tipo = .Text
    .Clear
    Permitido = False
    Set TBItem = CreateObject("adodb.recordset")
    TBItem.Open "Select versao from fases WHERE IDProcesso = " & txtidprocesso & " and versao is not null GROUP BY versao", Conexao, adOpenKeyset, adLockOptimistic
    If TBItem.EOF = False Then
        Do While TBItem.EOF = False
            .AddItem TBItem!versao
            If TBItem!versao = Tipo Then Permitido = True
            TBItem.MoveNext
        Loop
        TBItem.MoveFirst
        .Text = IIf(Tipo <> "" And Permitido = True, Tipo, TBItem!versao)
    End If
    TBItem.Close
End With

Exit Sub
tratar_erro:
    If Err.Number = "383" Then
        cmbVersao_pesquisar.Text = TBItem!versao
        Exit Sub
    End If
    USMsgBox ("Descrição do erro : " + Error() & Err.Number), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub procAtualiza()
On Error GoTo tratar_erro

If InputBox("Informe a senha para liberar.") = "280362P" Then frmProcessos_atualizar.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaEstrutura()
On Error GoTo tratar_erro
''ReDim arrNodes(2000)

Call m_Tree.Nodes.Clear
Grid1.rows = 1

m_Row = 1
m_Col = 1
Desenho = ""
Contador1 = -1
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Projproduto where Desenho = '" & txtdesenho & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    CodRef = ""
    Set TBFI = CreateObject("adodb.recordset")
    TBFI.Open "Select n_referencia from item_aplicacoes where codproduto = " & TBLISTA!Codproduto, Conexao, adOpenKeyset, adLockOptimistic
    If TBFI.EOF = False Then
        CodRef = TBFI!N_referencia
    End If
    TBFI.Close
    
    DataValidacao = ""
    RespValidacao = ""
    If TBLISTA!SubTipoItem <> 0 Then
        Set TBFI = CreateObject("adodb.recordset")
        TBFI.Open "Select * from Projconjunto_desc_versao where codproduto = " & TBLISTA!Codproduto & " and Versao = '" & cmbVersao_pesquisar_estrutura & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBFI.EOF = False Then
            DataValidacao = IIf(IsNull(TBFI!DtValidacao), "", TBFI!DtValidacao)
            RespValidacao = IIf(IsNull(TBFI!RespValidacao), "", TBFI!RespValidacao)
        End If
    End If
    Contador1 = Contador1 + 1
    arrNodes(Contador1).Level = 0
    arrNodes(Contador1).Text = TBLISTA!Desenho & vbTab & "" & vbTab & TBLISTA!Codproduto & vbTab & CodRef & vbTab & TBLISTA!Descricao & vbTab & TBLISTA!Unidade & vbTab & cmbVersao_pesquisar_estrutura & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & DataValidacao & vbTab & RespValidacao
    
    Codproduto = TBLISTA!Codproduto
    
    ProcNivel2Estrutura frmProcessos, cmbVersao_pesquisar_estrutura, False, False, True, False
    
    With Grid1
        .AutoRedraw = False
        .AllowUserPaste = cellTextOnly
        .ExtendLastCol = True
        .DrawMode = cellOwnerDraw
        .Cols = 18
        .rows = m_Row
        .Cell(0, 1).Text = "Cód. interno"
        .Cell(0, 2).Text = "Pos."
        .Cell(0, 3).Text = "ID"
        .Cell(0, 4).Text = "Cód. de ref."
        .Cell(0, 5).Text = "Descrição"
        .Cell(0, 6).Text = "Un."
        .Cell(0, 7).Text = "Ver."
        .Cell(0, 8).Text = "Vlr./un"
        .Cell(0, 9).Text = "Un/vlr."
        .Cell(0, 10).Text = "Dim/mm"
        .Cell(0, 11).Text = "Vlr./pç"
        .Cell(0, 12).Text = "Qtde."
        .Cell(0, 13).Text = "Total"
        .Cell(0, 14).Text = "Dt. validação"
        .Cell(0, 15).Text = "Resp. validação"
        .Cell(0, 16).Text = "Part number"
        .Cell(0, 17).Text = "Observações"
        .Range(0, 1, 0, 17).Alignment = cellCenterCenter
        .Column(1).Width = 200
        .Column(2).Width = 30
        .Column(3).Width = 0
        .Column(4).Width = 80
        .Column(5).Width = 300
        .Column(6).Width = 40
        .Column(6).Alignment = cellCenterCenter
        .Column(7).Width = 40
        .Column(7).Alignment = cellCenterCenter
        .Column(8).Width = 100
        .Column(8).Alignment = cellRightCenter
        .Column(9).Width = 40
        .Column(10).Width = 100
        .Column(10).Alignment = cellRightCenter
        .Column(11).Width = 100
        .Column(11).Alignment = cellRightCenter
        .Column(12).Width = 100
        .Column(12).Alignment = cellRightCenter
        .Column(13).Width = 100
        .Column(13).Alignment = cellRightCenter
        .Column(14).Width = 120
        .Column(15).Width = 100
        .Column(16).Width = 150
        .Column(17).Width = 400

        'First node
        Set tempNode = m_Tree.Nodes.Add("")
        .AddItem arrNodes(0).Text
        
        'Other nodes
        For intIndex = 1 To Contador1 'UBound(arrNodes)
            If arrNodes(intIndex).Level = arrNodes(intIndex - 1).Level Then
                Set tempNode = tempNode.Parent.Nodes.Add("")
            ElseIf arrNodes(intIndex).Level > arrNodes(intIndex - 1).Level Then
                Set tempNode = tempNode.Nodes.Add("")
            ElseIf arrNodes(intIndex).Level < arrNodes(intIndex - 1).Level Then
                For i = arrNodes(intIndex).Level To arrNodes(intIndex - 1).Level
                    Set tempNode = tempNode.Parent
                Next
                Set tempNode = tempNode.Nodes.Add("")
            End If
            .AddItem arrNodes(intIndex).Text
        Next
        
        .AutoRedraw = True
        .Refresh
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
    Case 2: ProcLocalizar
    Case 3: ProcSalvar
    Case 4: ProcExcluir
    Case 5: ProcImprimir
    Case 6: ProcAnterior
    Case 7: ProcProximo
    Case 8: procCopiarProcesso
    Case 9: ProcRevisar
    Case 10: ProcStatus
    Case 11: ProcValidarRegistros ListaProcessos, "Engenharia/Processos"
    Case 12: procAtualiza
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
    Case 1: procNovaFase
    Case 2: procSalvar_fase
    Case 3: procExcluir_fase
    Case 4: ProcImprimir
    Case 5: ProcAnterior
    Case 6: ProcProximo
    Case 7: procCopiar_fase
    Case 8: procRevisar_Fase
    Case 9: ProcRenumerar
    Case 10: procProgramas
    Case 11: ProcFerramentas
    Case 12: procTempos
    Case 13: ProcPlanoinspecao
    Case 15: ProcAjuda
    Case 16: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Grid1_Click()
On Error GoTo tratar_erro
Dim point As POINTAPI
Dim objCell As FlexCell.Cell
Dim intWidth As Integer

If FunCheckEditStatus() Then Exit Sub
intWidth = 20

Call GetCursorPos(point)
Call ScreenToClient(Grid1.hWnd, point)
Set objCell = Grid1.HitTest(point.x, point.Y)

If Not objCell Is Nothing Then
    If objCell.Row >= m_Row And objCell.Col = m_Col Then
        Dim objNode As Node
        Set objNode = m_Tree.FindNode(objCell.Row - m_Row + 2)
        If Not objNode Is Nothing Then
            Dim i As Long, x As Long, Y As Long
            x = objCell.Left + 2 + (objNode.Level - 1) * intWidth
            Y = objCell.Top + (objCell.Height - 9) / 2
            If point.x >= x And point.x <= x + 9 And point.Y >= Y And point.Y <= Y + 9 Then
                If objNode.Expanded Then
                    objNode.Collapse
                    Grid1.AutoRedraw = False
                    For i = 1 To objNode.ChildrenCount
                        Grid1.RowHeight(objCell.Row + i) = 0
                    Next
                    Grid1.AutoRedraw = True
                    Grid1.Refresh
                Else
                    objNode.Expand
                    Grid1.AutoRedraw = False
                    For i = 1 To objNode.ChildrenCount
                        If objNode.FindNode(i + 1).Visible Then
                            Grid1.RowHeight(objCell.Row + i) = -1 'DefaultRowHeight
                        End If
                    Next
                    Grid1.AutoRedraw = True
                    Grid1.Refresh
                End If
            End If
        End If
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Grid1_OwnerDrawCell(ByVal Row As Long, ByVal Col As Long, ByVal hdc As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Handled As Boolean)
On Error GoTo tratar_erro
Dim i As Long, j As Long
Dim x As Long, Y As Long
Dim hPen As Long, hOldPen As Long
Dim hBrush As Long, hOldBrush As Long
Dim lngLevel As Long
Dim blnDrawLine As Boolean
Dim objNode As Node, tmpNode As Node
Dim intWidth As Integer
Dim intAdd As Integer

If Row < m_Row Or Col <> m_Col Then Exit Sub

intWidth = 20
intAdd = 26
    
Set objNode = m_Tree.FindNode(Row - m_Row + 2)
If Not objNode Is Nothing Then
    lngLevel = objNode.Level - 1

    'Tree lines
    hPen = CreatePen(0, 1, RGB(128, 128, 128))
    hOldPen = SelectObject(hdc, hPen)
    For i = 0 To lngLevel
        If i < lngLevel - 1 Then
            blnDrawLine = True
            Set tmpNode = objNode
            For j = i To lngLevel - 2
                Set tmpNode = tmpNode.Parent
            Next
            If tmpNode.NextNode Is Nothing Then
                blnDrawLine = False
            End If
            If blnDrawLine Then
                'All
                Call DrawLine(hdc, Left + intWidth * i + intAdd, Top - 1, Left + intWidth * i + intAdd, Bottom + 1)
            End If
        ElseIf i = lngLevel - 1 Then
            'Top
            Call DrawLine(hdc, Left + intWidth * i + intAdd, Top - 1, Left + intWidth * i + intAdd, Top + (Bottom - Top) / 2)
            If Not objNode.NextNode Is Nothing Then
                'Bottom
                Call DrawLine(hdc, Left + intWidth * i + intAdd, Top + (Bottom - Top) / 2, Left + intWidth * i + intAdd, Bottom + 1)
            End If
        ElseIf i = lngLevel Then
            'Top
            If objNode.VisibleNodesCount > 1 Then
                Call DrawLine(hdc, Left + intWidth * i + intAdd, Top + (Bottom - Top) / 2, Left + intWidth * i + intAdd, Bottom + 1)
            End If
        End If
        'Horizontal line
        If lngLevel > 0 Then
            Call DrawLine(hdc, Left + intWidth * (lngLevel - 1) + intAdd, Top + (Bottom - Top) / 2, Left + intWidth * (lngLevel - 1) + intAdd + 10, Top + (Bottom - Top) / 2)
        End If
    Next
    
    Call SelectObject(hdc, hOldPen)
    Call DeleteObject(hPen)

    '+/-
    If objNode.ChildrenCount > 0 Then
        hPen = CreatePen(0, 1, 0)
        hOldPen = SelectObject(hdc, hPen)
        hBrush = CreateSolidBrush(RGB(255, 255, 255))
        hOldPen = SelectObject(hdc, hBrush)
        
        x = Left + 2 + intWidth * lngLevel
        Y = Top + (Bottom - Top - 9) / 2
        
        Call Rectangle(hdc, x, Y, x + 9, Y + 9)
        If objNode.Expanded Then
            Call DrawLine(hdc, x + 2, Y + 4, x + 7, Y + 4)
        Else
            Call DrawLine(hdc, x + 2, Y + 4, x + 7, Y + 4)
            Call DrawLine(hdc, x + 4, Y + 2, x + 4, Y + 7)
        End If
    
        Call SelectObject(hdc, hOldPen)
        Call DeleteObject(hPen)
        Call SelectObject(hdc, hOldBrush)
        Call DeleteObject(hBrush)
    End If
    
    'Icon
    If objNode.ChildrenCount > 0 Then
        DrawIconEx hdc, Left + intWidth * lngLevel + 18, Top + (Bottom - Top - 16) / 2, imgFolder.Picture, 16, 16, 0, 0, DI_NORMAL
    Else
        DrawIconEx hdc, Left + intWidth * lngLevel + 18, Top + (Bottom - Top - 16) / 2, imgFile.Picture, 16, 16, 0, 0, DI_NORMAL
    End If
    
    'Text
    With Grid1.Cell(Row, Col)
        Dim rc As rect
        Call SetRect(rc, Left + intWidth * lngLevel + 37, Top, Right, Bottom)
        Call DrawText(hdc, .Text, -1, rc, DT_SINGLELINE Or DT_VCENTER)
    End With

    Handled = True
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Function FunCheckEditStatus() As Boolean
On Error GoTo tratar_erro
Dim hWnd As Long
Dim strClassName As String
Dim intPos As Integer

strClassName = Space(256)
hWnd = GetFocus()
Call GetClassName(hWnd, strClassName, 256)
intPos = InStr(1, strClassName, Chr(0))
strClassName = Left(strClassName, intPos - 1)
If strClassName = "ThunderRT6TextBox" Then FunCheckEditStatus = True    'Editing Else    FunCheckEditStatus = False

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function
