VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUsuarios 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Configurações do sistema - Usuários"
   ClientHeight    =   10035
   ClientLeft      =   240
   ClientTop       =   450
   ClientWidth     =   15360
   ClipControls    =   0   'False
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
   Icon            =   "frmUsuarios.frx":0000
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
   Begin MSComctlLib.ListView ListView1 
      Height          =   6345
      Left            =   60
      TabIndex        =   17
      Top             =   3390
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   11192
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
         Text            =   "Usuário"
         Object.Width           =   15787
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Setor"
         Object.Width           =   4586
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "Expira"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Status"
         Object.Width           =   2646
      EndProperty
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   60
      TabIndex        =   49
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   10605
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   15360
      _ExtentX        =   27093
      _ExtentY        =   18706
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
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
      TabCaption(0)   =   "Usuários"
      TabPicture(0)   =   "frmUsuarios.frx":014A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frameusuario"
      Tab(0).Control(1)=   "USToolBar1"
      Tab(0).Control(2)=   "USImageList1"
      Tab(0).Control(3)=   "txtid"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Acessos"
      TabPicture(1)   =   "frmUsuarios.frx":0166
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Lista"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "frameacesso"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "USToolBar2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame1"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "USImageList2"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "txtIdAcesso"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      Begin VB.TextBox txtIdAcesso 
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
         Left            =   1410
         MaxLength       =   20
         MouseIcon       =   "frmUsuarios.frx":0182
         MousePointer    =   99  'Custom
         TabIndex        =   45
         Text            =   "0"
         Top             =   1710
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.TextBox txtid 
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
         Height          =   335
         Left            =   -73410
         MaxLength       =   20
         MouseIcon       =   "frmUsuarios.frx":048C
         MousePointer    =   99  'Custom
         TabIndex        =   28
         Text            =   "0"
         Top             =   5130
         Visible         =   0   'False
         Width           =   315
      End
      Begin DrawSuite2022.USImageList USImageList2 
         Left            =   12480
         Top             =   480
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmUsuarios.frx":0796
         Count           =   1
      End
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   -63030
         Top             =   450
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmUsuarios.frx":6B33
         Count           =   1
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Direitos"
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
         Left            =   11435
         TabIndex        =   44
         Top             =   1320
         Width           =   3825
         Begin VB.CheckBox chkValidacao 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Validar"
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
            Left            =   2910
            TabIndex        =   25
            Top             =   390
            Value           =   1  'Checked
            Width           =   795
         End
         Begin VB.CheckBox chkExcluir 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Excluir"
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
            Height          =   216
            Left            =   2085
            TabIndex        =   24
            Top             =   390
            Value           =   1  'Checked
            Width           =   765
         End
         Begin VB.CheckBox ChkAlterar 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Alterar"
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
            Height          =   216
            Left            =   1230
            TabIndex        =   23
            Top             =   390
            Value           =   1  'Checked
            Width           =   795
         End
         Begin VB.CheckBox Chknovo 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Criar novo"
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
            Height          =   216
            Left            =   150
            TabIndex        =   22
            Top             =   390
            Value           =   1  'Checked
            Width           =   1065
         End
      End
      Begin DrawSuite2022.USToolBar USToolBar2 
         Height          =   975
         Left            =   65
         TabIndex        =   43
         Top             =   330
         Width           =   15200
         _ExtentX        =   26802
         _ExtentY        =   1720
         ButtonCount     =   13
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
         ButtonCaption7  =   "Configurar listas"
         ButtonEnabled7  =   0   'False
         ButtonIconSize7 =   32
         ButtonToolTipText7=   "Configurar lista dos módulos (F7)"
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
         ButtonWidth7    =   100
         ButtonHeight7   =   21
         ButtonUseMaskColor7=   0   'False
         ButtonCaption8  =   "Acesso total"
         ButtonEnabled8  =   0   'False
         ButtonIconSize8 =   32
         ButtonToolTipText8=   "Definir acesso total (F8)"
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
         ButtonLeft8     =   411
         ButtonTop8      =   2
         ButtonWidth8    =   78
         ButtonHeight8   =   21
         ButtonUseMaskColor8=   0   'False
         ButtonCaption9  =   "Acesso por módulo"
         ButtonEnabled9  =   0   'False
         ButtonIconSize9 =   32
         ButtonToolTipText9=   "Definir acessos por módulo (F9)"
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
         ButtonLeft9     =   491
         ButtonTop9      =   2
         ButtonWidth9    =   115
         ButtonHeight9   =   21
         ButtonUseMaskColor9=   0   'False
         ButtonEnabled10 =   0   'False
         ButtonIconSize10=   32
         ButtonAlignment10=   2
         ButtonType10    =   1
         ButtonStyle10   =   -1
         BeginProperty ButtonFont10 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState10   =   -1
         ButtonLeft10    =   608
         ButtonTop10     =   4
         ButtonWidth10   =   2
         ButtonHeight10  =   54
         ButtonCaption11 =   "Ajuda"
         ButtonEnabled11 =   0   'False
         ButtonIconSize11=   32
         ButtonToolTipText11=   "Ajuda (F1)"
         ButtonKey11     =   "11"
         ButtonAlignment11=   2
         BeginProperty ButtonFont11 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft11    =   612
         ButtonTop11     =   2
         ButtonWidth11   =   41
         ButtonHeight11  =   21
         ButtonUseMaskColor11=   0   'False
         ButtonCaption12 =   "Sair"
         ButtonEnabled12 =   0   'False
         ButtonIconSize12=   32
         ButtonToolTipText12=   "Sair (Esc)"
         ButtonKey12     =   "12"
         ButtonAlignment12=   2
         BeginProperty ButtonFont12 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft12    =   655
         ButtonTop12     =   2
         ButtonWidth12   =   30
         ButtonHeight12  =   21
         ButtonUseMaskColor12=   0   'False
         ButtonEnabled13 =   0   'False
         ButtonIconSize13=   32
         ButtonKey13     =   "13"
         ButtonAlignment13=   2
         BeginProperty ButtonFont13 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState13   =   5
         ButtonLeft13    =   687
         ButtonTop13     =   2
         ButtonWidth13   =   24
         ButtonHeight13  =   24
         ButtonUseMaskColor13=   0   'False
      End
      Begin DrawSuite2022.USToolBar USToolBar1 
         Height          =   975
         Left            =   -74940
         TabIndex        =   42
         Top             =   330
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   1720
         ButtonCount     =   13
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
         ButtonCaption8  =   "Status"
         ButtonEnabled8  =   0   'False
         ButtonIconSize8 =   32
         ButtonToolTipText8=   "Status (F7)"
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
         ButtonLeft8     =   353
         ButtonTop8      =   2
         ButtonWidth8    =   45
         ButtonHeight8   =   21
         ButtonUseMaskColor8=   0   'False
         ButtonCaption9  =   "Atualizar"
         ButtonEnabled9  =   0   'False
         ButtonIconSize9 =   32
         ButtonToolTipText9=   "Utilizado pelo administrador do sistema."
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
         ButtonLeft9     =   400
         ButtonTop9      =   2
         ButtonWidth9    =   59
         ButtonHeight9   =   21
         ButtonUseMaskColor9=   0   'False
         ButtonEnabled10 =   0   'False
         ButtonIconSize10=   32
         ButtonAlignment10=   2
         ButtonType10    =   1
         ButtonStyle10   =   -1
         BeginProperty ButtonFont10 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState10   =   -1
         ButtonLeft10    =   461
         ButtonTop10     =   4
         ButtonWidth10   =   2
         ButtonHeight10  =   54
         ButtonCaption11 =   "Ajuda"
         ButtonEnabled11 =   0   'False
         ButtonIconSize11=   32
         ButtonToolTipText11=   "Ajuda (F1)"
         ButtonKey11     =   "11"
         ButtonAlignment11=   2
         BeginProperty ButtonFont11 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft11    =   465
         ButtonTop11     =   2
         ButtonWidth11   =   41
         ButtonHeight11  =   21
         ButtonUseMaskColor11=   0   'False
         ButtonCaption12 =   "Sair"
         ButtonEnabled12 =   0   'False
         ButtonIconSize12=   32
         ButtonToolTipText12=   "Sair (Esc)"
         ButtonKey12     =   "12"
         ButtonAlignment12=   2
         BeginProperty ButtonFont12 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft12    =   508
         ButtonTop12     =   2
         ButtonWidth12   =   30
         ButtonHeight12  =   21
         ButtonUseMaskColor12=   0   'False
         ButtonEnabled13 =   0   'False
         ButtonIconSize13=   32
         ButtonKey13     =   "13"
         ButtonAlignment13=   2
         BeginProperty ButtonFont13 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState13   =   5
         ButtonLeft13    =   540
         ButtonTop13     =   2
         ButtonWidth13   =   24
         ButtonHeight13  =   24
         ButtonUseMaskColor13=   0   'False
      End
      Begin VB.Frame Frameusuario 
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
         Height          =   2055
         Left            =   -74935
         TabIndex        =   31
         Top             =   1320
         Width           =   15200
         Begin VB.TextBox Txt_email 
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
            MaxLength       =   60
            TabIndex        =   9
            ToolTipText     =   "E-mail."
            Top             =   990
            Width           =   5925
         End
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   6030
            Top             =   660
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.CommandButton cmdImportar 
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
            Left            =   12150
            Picture         =   "frmUsuarios.frx":DFF4
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Localizar assinatura."
            Top             =   1590
            Width           =   315
         End
         Begin VB.CommandButton Cmd_limpar_caminho 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   12480
            Picture         =   "frmUsuarios.frx":E0F6
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Limpar caminho."
            Top             =   1590
            Width           =   315
         End
         Begin VB.CommandButton Cmd_visualizar_arquivo 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   12810
            Picture         =   "frmUsuarios.frx":E234
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Visualizar arquivo."
            Top             =   1590
            Width           =   315
         End
         Begin VB.TextBox Txt_caminho 
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
            Left            =   6660
            Locked          =   -1  'True
            TabIndex        =   12
            TabStop         =   0   'False
            ToolTipText     =   "Caminho da assinatura."
            Top             =   1590
            Width           =   5475
         End
         Begin VB.CheckBox chkAtivar_AvisosDiario 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Caption         =   "Ativar avisos diário?"
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
            Left            =   13200
            TabIndex        =   16
            Top             =   1650
            Width           =   1815
         End
         Begin VB.CommandButton cmdSetor 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   6270
            Picture         =   "frmUsuarios.frx":E7F6
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Localizar setor."
            Top             =   1590
            Width           =   315
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
            Left            =   1350
            MaxLength       =   50
            TabIndex        =   1
            ToolTipText     =   "Responsável pelo cadastro."
            Top             =   390
            Width           =   4485
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
            Left            =   180
            MaxLength       =   20
            TabIndex        =   0
            ToolTipText     =   "Data do cadastro."
            Top             =   390
            Width           =   1155
         End
         Begin VB.TextBox txtcracha 
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
            Left            =   7410
            MaxLength       =   8
            TabIndex        =   8
            ToolTipText     =   "Número do cracha."
            Top             =   990
            Width           =   1665
         End
         Begin VB.TextBox txtRepetir 
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
            IMEMode         =   3  'DISABLE
            Left            =   11760
            MaxLength       =   20
            PasswordChar    =   "*"
            TabIndex        =   4
            ToolTipText     =   "Senha do usuário."
            Top             =   390
            Width           =   1635
         End
         Begin VB.TextBox txtSenha 
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
            IMEMode         =   3  'DISABLE
            Left            =   10110
            MaxLength       =   20
            PasswordChar    =   "*"
            TabIndex        =   3
            ToolTipText     =   "Senha do usuário."
            Top             =   390
            Width           =   1635
         End
         Begin VB.TextBox Txt_setor 
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
            TabIndex        =   10
            TabStop         =   0   'False
            ToolTipText     =   "Setor."
            Top             =   1590
            Width           =   6075
         End
         Begin VB.CheckBox chkExpirar 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Expirar"
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
            Left            =   13635
            TabIndex        =   5
            Top             =   180
            Width           =   825
         End
         Begin VB.TextBox txtNome 
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
            TabIndex        =   7
            ToolTipText     =   "Nome."
            Top             =   990
            Width           =   7215
         End
         Begin VB.TextBox txtUsuario 
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
            Left            =   5850
            MaxLength       =   50
            TabIndex        =   2
            ToolTipText     =   "Usuário."
            Top             =   390
            Width           =   4245
         End
         Begin MSMask.MaskEdBox txtexpiracao 
            Height          =   315
            Left            =   13410
            TabIndex        =   6
            ToolTipText     =   "Data de expiração da senha."
            Top             =   390
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Enabled         =   0   'False
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
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Caminho da assinatura"
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
            Left            =   8580
            TabIndex        =   48
            Top             =   1380
            Width           =   1635
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
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
            Left            =   11842
            TabIndex        =   46
            Top             =   780
            Width           =   420
         End
         Begin VB.Label Label12 
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
            Left            =   3135
            TabIndex        =   41
            Top             =   180
            Width           =   915
         End
         Begin VB.Label Label11 
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
            Left            =   585
            TabIndex        =   40
            Top             =   180
            Width           =   345
         End
         Begin VB.Image imgCalendario 
            Height          =   360
            Left            =   14685
            Picture         =   "frmUsuarios.frx":E8F8
            Stretch         =   -1  'True
            ToolTipText     =   "Abrir calendário."
            Top             =   360
            Width           =   330
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nº do cracha"
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
            Left            =   7710
            TabIndex        =   37
            Top             =   780
            Width           =   1065
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Repetir"
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
            Left            =   12390
            TabIndex        =   36
            Top             =   180
            Width           =   525
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Senha"
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
            Left            =   10605
            TabIndex        =   35
            Top             =   180
            Width           =   450
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Setor"
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
            Left            =   3015
            TabIndex        =   34
            Top             =   1380
            Width           =   390
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nome completo*"
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
            Left            =   2850
            TabIndex        =   33
            Top             =   780
            Width           =   1185
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Usuário* (ex: Nome.sobrenome)"
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
            Left            =   6570
            TabIndex        =   32
            Top             =   180
            Width           =   2730
         End
      End
      Begin VB.Frame frameacesso 
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
         TabIndex        =   29
         Top             =   1320
         Width           =   11355
         Begin VB.TextBox txtValor_Limite 
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
            Left            =   10020
            MaxLength       =   20
            TabIndex        =   21
            ToolTipText     =   "Valor limite."
            Top             =   390
            Visible         =   0   'False
            Width           =   1155
         End
         Begin VB.TextBox txtResponsavel2 
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
            Left            =   1350
            MaxLength       =   50
            TabIndex        =   19
            ToolTipText     =   "Número do cracha."
            Top             =   390
            Width           =   3465
         End
         Begin VB.TextBox txtData2 
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
            MaxLength       =   20
            TabIndex        =   18
            ToolTipText     =   "Usuário."
            Top             =   390
            Width           =   1155
         End
         Begin VB.ComboBox cmbFormulario 
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
            ItemData        =   "frmUsuarios.frx":ED7B
            Left            =   4830
            List            =   "frmUsuarios.frx":ED7D
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   20
            ToolTipText     =   "Módulos."
            Top             =   390
            Width           =   6345
         End
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Valor limite"
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
            Left            =   10020
            TabIndex        =   47
            Top             =   180
            Visible         =   0   'False
            Width           =   1155
         End
         Begin VB.Label Label10 
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
            Left            =   2625
            TabIndex        =   39
            Top             =   180
            Width           =   915
         End
         Begin VB.Label Label9 
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
            Left            =   585
            TabIndex        =   38
            Top             =   180
            Width           =   345
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Módulos*"
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
            Left            =   4830
            TabIndex        =   30
            Top             =   180
            Width           =   6345
         End
      End
      Begin MSComctlLib.ListView Lista 
         Height          =   7575
         Left            =   65
         TabIndex        =   26
         Top             =   2160
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   13361
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
            Text            =   "Módulo"
            Object.Width           =   17198
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Criar novo"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Object.Tag             =   "T"
            Text            =   "Alterar"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "Excluir"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Object.Tag             =   "T"
            Text            =   "Validar"
            Object.Width           =   2117
         EndProperty
      End
   End
End
Attribute VB_Name = "frmUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Novo_Usuarios As Boolean 'OK
Dim Novo_Usuarios1 As Boolean 'OK
Public Sql_usuario_Localizar As String 'OK
Dim AcessoTotal As Boolean 'OK

'Corrige formulario
Dim Width_cmbFormulario As Long
Dim Width_txtValor_Limite As Long

Sub ProcAjuda()
On Error GoTo tratar_erro

FunAbrirVideoWeb ("http://www.youtube.com/watch?v=CsUpG3Xj0vA&list=UUODGiDjQ-BCrxh0YXfJvoqg&index=59&feature=plcp")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procConfigurar_lista()
On Error GoTo tratar_erro

frmUsuarios_configurar_listas.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro
  
frmUsuarios_abrir.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procAcessosModulo()
On Error GoTo tratar_erro

ProcLimpacampos_acesso
frameacesso.Enabled = False
Novo_Usuarios1 = False
frmUsuarios_acessos.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAnterior()
On Error GoTo tratar_erro

If txtId = 0 Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Usuarios order by Usuario", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.BOF = False Then
    TBLISTA.Find ("IDUsuario = " & txtId)
    TBLISTA.MovePrevious
    If TBLISTA.BOF = False Then
        txtId = TBLISTA!IDUsuario
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from Usuarios where IDUsuario = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
        ProcLimpaCampos
        ProcLimpacampos_acesso
        ProcPuxaDados
        ProcCarregaLista_acesso
    Else
        USMsgBox ("Fim dos cadastros de usuários."), vbInformation, "CAPRIND v5.0"
    End If
End If
Novo_Usuarios1 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcStatus()
On Error GoTo tratar_erro
  
If txtId.Text = 0 Then
    USMsgBox ("Informe o usuário antes de alterar o status."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
frmUsuarios_bloq.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procNovo_acesso()
On Error GoTo tratar_erro

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
ProcLimpacampos_acesso
frameacesso.Enabled = True
Novo_Usuarios1 = True
cmbFormulario.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcProximo()
On Error GoTo tratar_erro

If txtId = 0 Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Usuarios order by Usuario", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.BOF = False Then
    TBLISTA.Find ("IDUsuario = " & txtId)
    TBLISTA.MoveNext
    If TBLISTA.EOF = False Then
        txtId = TBLISTA!IDUsuario
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from Usuarios where IDUsuario = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
        ProcLimpaCampos
        ProcLimpacampos_acesso
        ProcPuxaDados
        ProcCarregaLista_acesso
    Else
        USMsgBox ("Fim dos cadastros de usuários."), vbInformation, "CAPRIND v5.0"
    End If
End If
Novo_Usuarios1 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ActiveResize1_ResizeComplete()
On Error GoTo tratar_erro

Width_cmbFormulario = cmbFormulario.Width
Width_txtValor_Limite = txtValor_Limite.Width
If SSTab1.Tab = 1 Then cmbFormulario_Click

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbFormulario_Click()
On Error GoTo tratar_erro

If AcessoTotal = False Then
    If cmbFormulario = "Compras/Pedido/Aprovar" Or cmbFormulario = "Avisos diário/Terceiros" Then
        With Label6
            .Left = cmbFormulario.Left
            .Width = Width_cmbFormulario - Width_txtValor_Limite
            cmbFormulario.Width = .Width
        End With
        Label13.Visible = True
        Label13.Left = txtValor_Limite.Left
        txtValor_Limite.Visible = True
        If cmbFormulario = "Avisos diário/Terceiros" Then
            Label13.Caption = "Dias"
            txtValor_Limite.ToolTipText = "Dias antes do prazo."
        Else
            Label13.Caption = "Valor limite"
            txtValor_Limite.ToolTipText = "Valor limite."
        End If
    Else
        Label6.Width = Width_cmbFormulario
        cmbFormulario.Width = Width_cmbFormulario
        Label13.Visible = False
        txtValor_Limite.Visible = False
    End If
End If

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

Private Sub cmdSetor_Click()
On Error GoTo tratar_erro

CadMaquinas = False
Funcionario = False
Usuarios = True
Estoque_Local_Armazenamento = False
frmUsuarios_Setor.Show 1

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
            Case vbKeyF7: ProcStatus
            Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
    Case 1:
        Select Case KeyCode
            Case vbKeyInsert: procNovo_acesso
            Case vbKeyF3: procSalvar_acesso
            Case vbKeyF4: procExcluir_acesso
            Case vbKeyF5: ProcImprimir
            Case vbKeyF7: procConfigurar_lista
            Case vbKeyF8: procAcessoTotal
            Case vbKeyF9: procAcessosModulo
            Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
   
Sub ProcCarregaLista_acesso()
On Error GoTo tratar_erro
  
Lista.ListItems.Clear
Set TBAcessos = CreateObject("adodb.recordset")
TBAcessos.Open "Select * FROM acessos where idusuario = " & txtId & " order by Acesso", Conexao, adOpenKeyset, adLockOptimistic
If TBAcessos.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBAcessos.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBAcessos.EOF = False
        With Lista.ListItems
            .Add , , TBAcessos!IDAcesso
            .Item(.Count).SubItems(1) = IIf(IsNull(TBAcessos!Acesso), "", TBAcessos!Acesso)
            .Item(.Count).SubItems(2) = IIf(TBAcessos!Incluir = True, "S", "N")
            .Item(.Count).SubItems(3) = IIf(TBAcessos!Alterar = True, "S", "N")
            .Item(.Count).SubItems(4) = IIf(TBAcessos!Excluir = True, "S", "N")
            .Item(.Count).SubItems(5) = IIf(TBAcessos!Validacao = True, "S", "N")
        End With
        TBAcessos.MoveNext
        Contador = Contador + 1
        PBLista.Value = Contador
    Loop
End If
TBAcessos.Close
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaLista()
On Error GoTo tratar_erro

If Sql_usuario_Localizar = "" Then Exit Sub
Set TBUsuarios = CreateObject("adodb.recordset")
TBUsuarios.Open Sql_usuario_Localizar, Conexao, adOpenKeyset, adLockOptimistic
If TBUsuarios.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBUsuarios.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBUsuarios.EOF = False
        With ListView1.ListItems
            .Add , , TBUsuarios!IDUsuario
            .Item(.Count).SubItems(1) = TBUsuarios!Usuario
            .Item(.Count).SubItems(2) = TBUsuarios!Setor
            .Item(.Count).SubItems(3) = IIf(IsNull(TBUsuarios!DtExpiracao), "não expira", Format(TBUsuarios!DtExpiracao, "dd/mm/yyyy"))
            If TBUsuarios!Bloqueado = True Then .Item(.Count).SubItems(4) = "Bloqueado" Else .Item(.Count).SubItems(4) = "Liberado"
        End With
        TBUsuarios.MoveNext
        Contador = Contador + 1
        PBLista.Value = Contador
    Loop
End If
TBUsuarios.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkExpirar_Click()
On Error GoTo tratar_erro
  
If chkExpirar.Value = 1 Then
    txtexpiracao.Enabled = True
Else
    txtexpiracao.Enabled = False
    txtexpiracao.Text = "__/__/____"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procAcessoTotal()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If USMsgBox("Deseja realmente definir acesso total para este usuário?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    AcessoTotal = True
    Conexao.Execute "DELETE from acessos where IDUsuario = " & txtId
    Lista.ListItems.Clear
    For i = 0 To cmbFormulario.ListCount - 1
        cmbFormulario.ListIndex = i
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from acessos", Conexao, adOpenKeyset, adLockOptimistic
        TBAbrir.AddNew
        TBAbrir!IDUsuario = txtId
        TBAbrir!Acesso = cmbFormulario
        TBAbrir!Incluir = True
        TBAbrir!Alterar = True
        TBAbrir!Excluir = True
        TBAbrir!Validacao = True
        TBAbrir!Data = Date
        TBAbrir!Responsavel = pubUsuario
        TBAbrir.Update
    Next
    ProcLimpacampos_acesso
    frameacesso.Enabled = False
    Novo_Usuarios1 = False
    USMsgBox ("Acesso total definido com sucesso."), vbInformation, "CAPRIND v5.0"
    '==================================
    Modulo = "Configuração do sistema/Usuários"
    Evento = "Definir acesso total"
    ID_documento = 0
    Documento = "Usuario: " & txtUsuario
    Documento1 = ""
    ProcGravaEvento
    '==================================
    AcessoTotal = False
    Direitos
    ProcCarregaLista_acesso
    ProcRecarregaMenu
End If

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
Novo_Usuarios = True
Frameusuario.Enabled = True
txtUsuario.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procExcluir_acesso()
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
                If USMsgBox("Deseja realmente excluir este(s) acesso(s)?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            Conexao.Execute "DELETE from acessos where IDAcesso = " & .ListItems(InitFor)
            '==================================
            Modulo = "Configuração do sistema/Usuários"
            Evento = "Excluir"
            ID_documento = txtIdAcesso
            Documento = "Usuario: " & txtUsuario
            Documento1 = "Acesso: " & .ListItems.Item(InitFor).SubItems(1)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) acesso(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Acesso(s) excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpacampos_acesso
    frameacesso.Enabled = False
    Novo_Usuarios1 = False
    Direitos
    ProcCarregaLista_acesso
    ProcRecarregaMenu
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcSalvar()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & ", voce não está autorizado a gravar neste formulário."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Frameusuario.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
Acao = "salvar"
If txtUsuario = "" Then
    NomeCampo = "o usuário"
    ProcVerificaAcao
    txtUsuario.SetFocus
    Exit Sub
End If
If txtNome.Text = "" Then
    NomeCampo = "o nome do usuário"
    ProcVerificaAcao
    txtNome.SetFocus
    Exit Sub
End If
'If Trim(txtSenha.Text) = "" Or Len(txtSenha.Text) < 3 Then
'    USMsgBox "O campo <senha> está em branco ou é menor que três caracteres.", vbExclamation, "CAPRIND v5.0"
'    txtSenha.SetFocus
'    Exit Sub
'End If
If txtSenha.Text <> txtRepetir.Text Then
    USMsgBox "A senha digitada é inválida. Os campos <senha> e <repetir> não são iguais.", vbExclamation, "CAPRIND v5.0"
    txtSenha.SetFocus
    Exit Sub
End If
If chkExpirar.Value = 1 Then
    If txtexpiracao = "__/__/____" Then
        NomeCampo = "a data para expiração da senha"
        ProcVerificaAcao
        imgCalendario_Click
        Exit Sub
    End If
End If
If Novo_Usuarios = True Then
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from usuarios where usuario = '" & txtUsuario.Text & "' order by usuario", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        USMsgBox ("Usuário já cadastrado, favor alterar."), vbExclamation, "CAPRIND v5.0"
        TBAbrir.Close
        txtUsuario.SetFocus
        Exit Sub
    End If
    TBAbrir.Close
End If

Set TBUsuarios = CreateObject("adodb.recordset")
TBUsuarios.Open "Select * FROM Usuarios WHERE Idusuario = " & txtId.Text, Conexao, adOpenKeyset, adLockOptimistic
If TBUsuarios.EOF = True Then
    TBUsuarios.AddNew
Else
    If TBUsuarios!Usuario <> txtUsuario Then
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "SELECT T.name AS Tabela, C.name AS Coluna FROM sys.sysobjects AS T (NOLOCK) INNER JOIN sys.all_columns AS C (NOLOCK) ON T.id = C.object_id AND T.XTYPE = 'U' WHERE C.NAME LIKE 'Resp%' and C.NAME <> 'Responsavel_rel' or C.NAME = 'Operador' or C.NAME = 'Usuario' and T.Name <> 'Usuarios' and T.Name <> 'Empresa_email' ORDER BY T.name ASC", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Do While TBAbrir.EOF = False
                Conexao.Execute "UPDATE " & TBAbrir!Tabela & " Set " & TBAbrir!Coluna & " = '" & txtUsuario & "' where " & TBAbrir!Coluna & " = '" & TBUsuarios!Usuario & "'"
                TBAbrir.MoveNext
            Loop
        End If
        TBAbrir.Close
    End If
    If TBUsuarios!Setor <> Txt_setor Then
        Conexao.Execute "UPDATE Compras_requisicao Set Setorsolic = '" & Txt_setor & "' where Solicitado = '" & txtUsuario & "'"
        Conexao.Execute "UPDATE Compras_requisicao Set Setorautor = '" & Txt_setor & "' where Autorizado = '" & txtUsuario & "'"
    End If
End If
If txtData = "" Then TBUsuarios!Data = Date Else TBUsuarios!Data = txtData
If txtResponsavel = "" Then TBUsuarios!Responsavel = pubUsuario Else TBUsuarios!Responsavel = txtResponsavel
TBUsuarios!Usuario = Trim(txtUsuario.Text)
TBUsuarios!Nome = txtNome
TBUsuarios!Setor = Txt_setor
TBUsuarios!Assinatura = txt_Caminho
TBUsuarios!Senha = txtSenha.Text
If chkAtivar_AvisosDiario.Value = 1 Then TBUsuarios!Aviso_diario = True Else TBUsuarios!Aviso_diario = False
TBUsuarios!CODIGO = IIf(txtcracha = "", Null, FunTamanhoTextoZeroEsq(txtcracha, 8))
TBUsuarios!Email = IIf(Txt_email.Text = "", Null, LCase(Txt_email.Text))
If chkExpirar.Value = 1 Then
    TBUsuarios!Expira = True
    TBUsuarios!DtExpiracao = (txtexpiracao.Text)
Else
    TBUsuarios!Expira = False
    TBUsuarios!DtExpiracao = Null
End If
TBUsuarios.Update
txtId = TBUsuarios!IDUsuario
TBUsuarios.Close
ListView1.ListItems.Clear
ProcCarregaLista
If Novo_Usuarios = True Then
    USMsgBox ("Novo usuário cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar"
    If CodigoLista <> 0 And ListView1.ListItems.Count <> 0 Then
        ListView1.SelectedItem = ListView1.ListItems(CodigoLista)
        ListView1.SetFocus
    End If
End If
'====================================
Modulo = "Configuração do sistema/Usuários"
ID_documento = txtId
Documento = "Usuário: " & txtUsuario
Documento1 = ""
ProcGravaEvento
'====================================
Novo_Usuarios = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procSalvar_acesso()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If frameacesso.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
If cmbFormulario = "" Then
    USMsgBox ("Informe o módulo antes de salvar."), vbExclamation, "CAPRIND v5.0"
    cmbFormulario.SetFocus
    Exit Sub
End If
If Novo_Usuarios1 = True Then
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from Acessos where Acesso = '" & cmbFormulario & "' and IDUsuario = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        USMsgBox "Este acesso já foi liberado para este usuário.", vbExclamation, "CAPRIND v5.0"
        Exit Sub
    End If
    TBAbrir.Close
End If
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from acessos where idacesso =  " & txtIdAcesso, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then TBGravar.AddNew
If txtData2 = "" Then TBGravar!Data = Date Else TBGravar!Data = txtData2
If txtResponsavel = "" Then TBGravar!Responsavel = pubUsuario Else TBGravar!Responsavel = txtResponsavel2
TBGravar!IDUsuario = txtId
TBGravar!Acesso = cmbFormulario
If Chknovo.Value = 1 Then TBGravar!Incluir = True Else TBGravar!Incluir = False
If ChkAlterar.Value = 1 Then TBGravar!Alterar = True Else TBGravar!Alterar = False
If chkExcluir.Value = 1 Then TBGravar!Excluir = True Else TBGravar!Excluir = False
If chkValidacao.Value = 1 Then TBGravar!Validacao = True Else TBGravar!Validacao = False
If cmbFormulario = "Compras/Pedido/Aprovar" Then
    If IIf(txtValor_Limite = "", 0, txtValor_Limite) <= 0 Then TBGravar!Valor_Limite = Null Else TBGravar!Valor_Limite = txtValor_Limite
Else
    TBGravar!Valor_Limite = Null
End If

If cmbFormulario = "Avisos diário/Terceiros" Then
    If IIf(txtValor_Limite = "", 0, txtValor_Limite) <= 0 Then TBGravar!Dias_Terceiros = 0 Else TBGravar!Dias_Terceiros = Format(txtValor_Limite, "###,##0")
Else
    TBGravar!Dias_Terceiros = Null
End If

TBGravar.Update
txtIdAcesso = TBGravar!IDAcesso
TBGravar.Close
ProcCarregaLista_acesso
If Novo_Usuarios1 = True Then
    USMsgBox ("Novo acesso cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo acesso"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar acesso"
    If CodigoLista1 <> 0 And Lista.ListItems.Count <> 0 Then
        Lista.SelectedItem = Lista.ListItems(CodigoLista1)
        Lista.SetFocus
    End If
End If
'==================================
Modulo = "Configuração do sistema/Usuários"
ID_documento = txtIdAcesso
Documento = "Usuario: " & txtUsuario
Documento1 = "Acesso: " & cmbFormulario
ProcGravaEvento
'==================================
Novo_Usuarios1 = False
ProcRecarregaMenu

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCampos()
On Error GoTo tratar_erro

txtId.Text = 0
txtData = Format(Date, "dd/mm/yy")
txtResponsavel = pubUsuario
txtUsuario.Text = ""
txtcracha = ""
txtNome.Text = ""
Txt_setor.Text = ""
txt_Caminho = ""
txtSenha.Text = ""
txtRepetir.Text = ""
Txt_email.Text = ""
chkExpirar.Value = 0
txtexpiracao.Text = "__/__/____"
chkAtivar_AvisosDiario.Value = 0
CodigoLista = 0
Caption = "Configurações do sistema - Usuários"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcExcluir()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With ListView1
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) usuário(s)?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            Conexao.Execute "DELETE from Usuarios WHERE IDUsuario = " & .ListItems(InitFor)
            Conexao.Execute "DELETE from Acessos WHERE IDUsuario = " & .ListItems(InitFor)
            '====================================
            Modulo = "Configuração do sistema/Usuários"
            Evento = "Excluir"
            ID_documento = .ListItems(InitFor)
            Documento = "Usuário: " & .ListItems.Item(InitFor).SubItems(1)
            Documento1 = ""
            ProcGravaEvento
            '===================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) usuários(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Usuários(s) excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCampos
    ListView1.ListItems.Clear
    ProcCarregaLista
    Novo_Usuarios = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro
  
Formulario = "Configuração do sistema/Usuários"
Direitos
ProcLimpaVariaveisPrincipais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procAtualiza()
On Error GoTo tratar_erro

If InputBox("Informe a senha para liberar.") = "280362U" Then
    If USMsgBox("Deseja realmente atualizar a definição de acesso e o número do cracha de todos os usuários?", vbYesNo, "CAPRIND v5.0") = vbYes Then
        Set TBUsuarios = CreateObject("adodb.recordset")
        TBUsuarios.Open "select * from Usuarios", Conexao, adOpenKeyset, adLockOptimistic
        If TBUsuarios.EOF = False Then
            TBUsuarios.MoveLast
            PBLista.Min = 0
            PBLista.Max = TBUsuarios.RecordCount
            PBLista.Value = 1
            Contador = 0
            TBUsuarios.MoveFirst
            Do While TBUsuarios.EOF = False
                If IsNull(TBUsuarios!CODIGO) = False And TBUsuarios!CODIGO <> "" Then
                    Select Case Len(TBUsuarios!CODIGO)
                        Case 1: TBUsuarios!CODIGO = "0000000" & TBUsuarios!CODIGO
                        Case 2: TBUsuarios!CODIGO = "000000" & TBUsuarios!CODIGO
                        Case 3: TBUsuarios!CODIGO = "00000" & TBUsuarios!CODIGO
                        Case 4: TBUsuarios!CODIGO = "0000" & TBUsuarios!CODIGO
                        Case 5: TBUsuarios!CODIGO = "000" & TBUsuarios!CODIGO
                        Case 6: TBUsuarios!CODIGO = "00" & TBUsuarios!CODIGO
                        Case 7: TBUsuarios!CODIGO = "0" & TBUsuarios!CODIGO
                    End Select
                    TBUsuarios.Update
                End If
                TBUsuarios.MoveNext
                Contador = Contador + 1
                PBLista.Value = Contador
            Loop
        End If
        Set TBUsuarios = CreateObject("adodb.recordset")
        TBUsuarios.Open "Select * from Usuarios_monitor_trabalho", Conexao, adOpenKeyset, adLockOptimistic
        If TBUsuarios.EOF = False Then
            TBUsuarios.MoveLast
            PBLista.Min = 0
            PBLista.Max = TBUsuarios.RecordCount
            PBLista.Value = 1
            Contador = 0
            TBUsuarios.MoveFirst
            Do While TBUsuarios.EOF = False
                If IsNull(TBUsuarios!Modulo) = True Or TBUsuarios!Modulo = "" Then
                    TBUsuarios!Modulo = "PCP/Monitor de trabalho"
                    TBUsuarios.Update
                End If
                TBUsuarios.MoveNext
                Contador = Contador + 1
                PBLista.Value = Contador
            Loop
        End If
        TBUsuarios.Close
        
        USMsgBox ("Atualização efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
        '==================================
        Modulo = "Configuração do sistema/Usuários"
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

Private Sub ProcImprimir()
On Error GoTo tratar_erro

frmUsuarios_localizarsetor.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub imgCalendario_Click()
On Error GoTo tratar_erro

If txtexpiracao.Enabled = False Then Exit Sub
Usuarios = True
Faturamento = False
Compras_Pedido = False
Compras_Requisicao = False
Compras_Fallow_up = False
Vendas_Carteira = False
Vendas_PI = False
Manutencao = False
Compras_Cotacao = False
Inspecao_recebimento = False
Funcionario = False
RNC = False
SolicitacaoAcao = False
Troca_Duplicata = False
Financeiro_Contas_Recebidas = False
Engenharia_Normas = False
Qualidade_PPAP_Plano = False
Qualidade_PPAP_FMEA = False
Qualidade_sistema = False
Engenharia = False
Compras_Fornecedores = False
Vendas_Programacao = False
Qualidade_PPAP_PSW = False
Outros_solicitacaoPCP = False
Estoque_recebimento = False
FrmCalendario.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

If Novo_Usuarios = True Then
    If USMsgBox("O usuário ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvar
        If Novo_Usuarios = True Then Exit Sub Else Unload Me
    End If
End If
If Novo_Usuarios1 = True Then
    If USMsgBox("O acesso ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        procSalvar_acesso
        If Novo_Usuarios1 = True Then Exit Sub Else Unload Me
    End If
End If
Novo_Usuarios = False
Novo_Usuarios1 = False
Unload Me

frmMDI.ProcVerificaAvisoDiario

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

With Lista
    If .ListItems.Count = 0 Then Exit Sub
        txtIdAcesso = .SelectedItem
        cmbFormulario.Text = .SelectedItem.SubItems(1)
1:
        
        If .SelectedItem.SubItems(2) = "S" Then Chknovo.Value = 1 Else Chknovo.Value = 0
        If .SelectedItem.SubItems(3) = "S" Then ChkAlterar.Value = 1 Else ChkAlterar.Value = 0
        If .SelectedItem.SubItems(4) = "S" Then chkExcluir.Value = 1 Else chkExcluir.Value = 0
        If .SelectedItem.SubItems(5) = "S" Then chkValidacao.Value = 1 Else chkValidacao.Value = 0
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from acessos where IDAcesso = " & txtIdAcesso, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            txtData2 = IIf(IsNull(TBAbrir!Data), "", Format(TBAbrir!Data, "dd/mm/yy"))
            txtResponsavel2 = IIf(IsNull(TBAbrir!Responsavel), "", TBAbrir!Responsavel)
            txtValor_Limite = ""
            If cmbFormulario = "Compras/Pedido/Aprovar" Then
                txtValor_Limite = IIf(IsNull(TBAbrir!Valor_Limite), "", Format(TBAbrir!Valor_Limite, "###,##0.00"))
            ElseIf cmbFormulario = "Avisos diário/Terceiros" Then
                txtValor_Limite = IIf(IsNull(TBAbrir!Dias_Terceiros), "", TBAbrir!Dias_Terceiros)
            End If
        End If
        TBAbrir.Close
        CodigoLista1 = .SelectedItem.index
        Novo_Usuarios1 = False
        frameacesso.Enabled = True
End With

Exit Sub
tratar_erro:
    If Err.Number = "383" Then
        Conexao.Execute "DELETE from Acessos WHERE Acesso = '" & Lista.SelectedItem.SubItems(1) & "'"
        ProcCarregaLista_acesso
        GoTo 1:
    End If
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro
  
Formulario = "Configuração do sistema/Usuários"
Direitos
SSTab1.Tab = 0
ProcLimpaVariaveisPrincipais

ProcCarregaToolBar1 Me, 15195, 13, True
ProcCarregaToolBar2 Me, 15195, 13, True
cmbFormulario.Clear
ProcCarregaComboModulos cmbFormulario, False, ""

ProcRemoveObjetosResize Me

Width_cmbFormulario = cmbFormulario.Width
Width_txtValor_Limite = txtValor_Limite.Width
If SSTab1.Tab = 1 Then cmbFormulario_Click

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With ListView1
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If .ListItems.Item(InitFor).SubItems(1) = "PROCAM" Then GoTo Proximo
                Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open "SELECT * FROM SYSOBJECTS OBJ where Left(OBJ.NAME,7) = 'Mascara' and Right(OBJ.NAME,10) <> 'PrimaryKey'", Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    Do While TBAbrir.EOF = False
                        ProcVerificaRegistroUtilizadoSemMsg TBAbrir!Name, "Usuario = '" & .ListItems(InitFor).ListSubItems(1) & "'"
                        If Permitido = False Then GoTo Proximo
                        TBAbrir.MoveNext
                    Loop
                End If
                TBAbrir.Close
                ProcVerificaRegistroUtilizadoSemMsg "Usuarios_Setor_Responsavel", "Responsavel_CC = '" & .ListItems(InitFor).ListSubItems(1) & "'"
                If Permitido = False Then GoTo Proximo
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView ListView1, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With ListView1
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If .ListItems.Item(InitFor).SubItems(1) = "PROCAM" Then
                USMsgBox ("Não é permitido excluir o administrador do sistema PROCAM, fale com o administrador do sistema."), vbExclamation, "CAPRIND v5.0"
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            
            Mensagem = "Não é permitido excluir este usuário, pois o mesmo está sendo utilizado em"
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "SELECT * FROM SYSOBJECTS OBJ where Left(OBJ.NAME,7) = 'Mascara' and Right(OBJ.NAME,10) <> 'PrimaryKey'", Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False Then
                Do While TBFI.EOF = False
                    ProcVerificaRegistroUtilizado TBFI!Name, "Usuario = '" & .ListItems(InitFor).ListSubItems(1) & "'", "outro módulo"
                    If Permitido = False Then
                        .ListItems.Item(InitFor).Checked = False
                        Exit Sub
                    End If
                    TBFI.MoveNext
                Loop
            End If
            TBFI.Close
            
            Mensagem = "Não é permitido excluir este usuário, pois o mesmo está sendo utilizado no módulo"
            ProcVerificaRegistroUtilizado "Usuarios_Setor_Responsavel", "Responsavel_CC = '" & .ListItems(InitFor).ListSubItems(1) & "'", "Custos/Centro de custo"
            If Permitido = False Then .ListItems.Item(InitFor).Checked = False
        End If
    Next InitFor
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If ListView1.ListItems.Count = 0 Then Exit Sub
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * FROM Usuarios WHERE IdUsuario = " & ListView1.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    ProcLimpaCampos
    ProcPuxaDados
    CodigoLista = ListView1.SelectedItem.index
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error GoTo tratar_erro

If txtId = 0 Then
    SSTab1.Tab = 0
    Exit Sub
End If
Select Case SSTab1.Tab
    Case 0:
        ListView1.Visible = True
        If ListView1.Visible = True Then ListView1.SetFocus
    Case 1:
        ListView1.Visible = False
        If Novo_Usuarios = True Then
            USMsgBox ("Salve o usuário antes de prosseguir."), vbExclamation, "CAPRIND v5.0"
            SSTab1.Tab = 0
            Exit Sub
        End If
        Lista.SetFocus
        ProcLimpacampos_acesso
        ProcCarregaLista_acesso
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_email_LostFocus()
On Error GoTo tratar_erro

Txt_email = LCase(Txt_email)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtcracha_Change()
On Error GoTo tratar_erro

If txtcracha <> "" Then
    VerifNumero = txtcracha
    ProcVerificaNumero
    If VerifNumero = False Then
        txtcracha = ""
        txtcracha.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtcracha_LostFocus()
On Error GoTo tratar_erro

If txtcracha <> "" Then txtcracha = FunTamanhoTextoZeroEsq(txtcracha, 8)
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtexpiracao_LostFocus()
On Error GoTo tratar_erro

If txtexpiracao.Text <> "__/__/____" Then
    VerifData = txtexpiracao.Text
    ProcVerificaData
    If VerifData = False Then
        txtexpiracao.Text = "__/__/____"
        txtexpiracao.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpacampos_acesso()
On Error GoTo tratar_erro

txtIdAcesso = 0
txtData2 = Format(Date, "dd/mm/yy")
txtResponsavel2 = pubUsuario
cmbFormulario.ListIndex = -1
Chknovo.Value = 0
ChkAlterar.Value = 0
chkExcluir.Value = 0
chkValidacao.Value = 0
txtValor_Limite = ""
CodigoLista1 = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcPuxaDados()
On Error GoTo tratar_erro

txtId.Text = TBAbrir!IDUsuario
txtUsuario.Text = TBAbrir!Usuario
Caption = "Configurações do sistema - Usuários (Usuário : " & TBAbrir!Usuario & ")"
If TBAbrir!CODIGO <> "" Then txtcracha = IIf(IsNull(TBAbrir!CODIGO), "", TBAbrir!CODIGO)
txtNome.Text = IIf(IsNull(TBAbrir!Nome), "", TBAbrir!Nome)
Txt_setor.Text = IIf(IsNull(TBAbrir!Setor), "", TBAbrir!Setor)
txt_Caminho = IIf(IsNull(TBAbrir!Assinatura), "", TBAbrir!Assinatura)
txtSenha.Text = IIf(IsNull(TBAbrir!Senha), "", TBAbrir!Senha)
txtRepetir.Text = IIf(IsNull(TBAbrir!Senha), "", TBAbrir!Senha)
If TBAbrir!Aviso_diario = True Then chkAtivar_AvisosDiario.Value = 1 Else chkAtivar_AvisosDiario.Value = 0
If IsNull(TBAbrir!DtExpiracao) = False And TBAbrir!DtExpiracao <> "" Then
    chkExpirar.Value = 1
    txtexpiracao.Text = Format(TBAbrir!DtExpiracao, "dd/mm/yyyy")
Else
    chkExpirar.Value = 0
    txtexpiracao.Text = "__/__/____"
End If
Txt_email.Text = IIf(IsNull(TBAbrir!Email), "", TBAbrir!Email)
txtData = IIf(IsNull(TBAbrir!Data), "", Format(TBAbrir!Data, "dd/mm/yy"))
txtResponsavel = IIf(IsNull(TBAbrir!Responsavel), "", TBAbrir!Responsavel)
Frameusuario.Enabled = True
Novo_Usuarios = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtValor_Limite_Change()
On Error GoTo tratar_erro

If txtValor_Limite.Text <> "" Then
    VerifNumero = txtValor_Limite.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtValor_Limite.Text = ""
        txtValor_Limite.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtValor_Limite_LostFocus()
On Error GoTo tratar_erro

If cmbFormulario = "Compras/Pedido/Aprovar" Then txtValor_Limite = Format(txtValor_Limite, "###,##0.00") Else txtValor_Limite = Format(txtValor_Limite, "###,##0")

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
    Case 8: ProcStatus
    Case 9: procAtualiza
    Case 11: ProcAjuda
    Case 12: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar2_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: procNovo_acesso
    Case 2: procSalvar_acesso
    Case 3: procExcluir_acesso
    Case 4: ProcImprimir
    Case 5: ProcAnterior
    Case 6: ProcProximo
    Case 7: procConfigurar_lista
    Case 8: procAcessoTotal
    Case 9: procAcessosModulo
    Case 11: ProcAjuda
    Case 12: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
