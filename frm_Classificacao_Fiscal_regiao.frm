VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{8C1279ED-044C-4258-A3E3-0D5514B899FC}#1.44#0"; "ControlesUteis.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_Classificacao_Fiscal_regiao 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Fiscal | Classificação fiscal | Regiões e Substituição tributária"
   ClientHeight    =   8235
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   10980
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   8235
   ScaleWidth      =   10980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   44
      Top             =   7830
      Width           =   10980
      _ExtentX        =   19368
      _ExtentY        =   714
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   43
      Top             =   0
      Width           =   10980
      _ExtentX        =   19368
      _ExtentY        =   688
      DibPicture      =   "frm_Classificacao_Fiscal_regiao.frx":0000
      CaptionDelimiter=   "|"
      CaptionOnCenter =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Icon            =   "frm_Classificacao_Fiscal_regiao.frx":7180
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   195
      TabIndex        =   33
      Top             =   840
      Width           =   10515
      _ExtentX        =   18547
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
      ButtonKey5      =   "9"
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
      ButtonKey6      =   "10"
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
      ButtonKey7      =   "11"
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
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   5340
         Top             =   270
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frm_Classificacao_Fiscal_regiao.frx":749A
         Count           =   1
      End
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   225
      TabIndex        =   34
      Top             =   7380
      Width           =   10515
      _ExtentX        =   18547
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   7245
      Left            =   150
      TabIndex        =   18
      Top             =   480
      Width           =   10665
      _ExtentX        =   18812
      _ExtentY        =   12779
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
      TabCaption(0)   =   "Regiões (UF's)"
      TabPicture(0)   =   "frm_Classificacao_Fiscal_regiao.frx":A883
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Lista"
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(2)=   "Frame2"
      Tab(0).Control(3)=   "txtId"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Substituição tributária ( NCM | Regiões)"
      TabPicture(1)   =   "frm_Classificacao_Fiscal_regiao.frx":A89F
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Lista1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "txtId1"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VB.TextBox txtId1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   690
         TabIndex        =   23
         Text            =   "0"
         Top             =   3900
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.TextBox txtId 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -74310
         TabIndex        =   19
         Text            =   "0"
         Top             =   3900
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   -74925
         TabIndex        =   20
         Top             =   1320
         Width           =   10515
         Begin VB.OptionButton OptDE 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Dentro do estado"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   240
            TabIndex        =   0
            Top             =   180
            Width           =   1575
         End
         Begin VB.OptionButton OptCO 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Centro oeste"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   6340
            TabIndex        =   2
            Top             =   180
            Width           =   1275
         End
         Begin VB.OptionButton OptNN 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Norte e nordeste"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   3260
            TabIndex        =   1
            Top             =   180
            Width           =   1635
         End
         Begin VB.OptionButton OptSS 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Sul e sudeste"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   9060
            TabIndex        =   3
            Top             =   180
            Width           =   1275
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   -74925
         TabIndex        =   21
         Top             =   1870
         Width           =   10515
         Begin VB.TextBox Txt_aliquota_interna_ICMS 
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
            TabIndex        =   8
            ToolTipText     =   "Alíquota interna do ICMS (Cálculo DIFAL)."
            Top             =   405
            Width           =   1080
         End
         Begin VB.TextBox Txt_FCP 
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
            Left            =   8460
            TabIndex        =   7
            ToolTipText     =   "Percentual relativo ao FCP (fundo de combate à pobreza)."
            Top             =   405
            Width           =   765
         End
         Begin VB.TextBox txtUF 
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
            Left            =   7890
            TabIndex        =   6
            ToolTipText     =   "UF."
            Top             =   405
            Width           =   570
         End
         Begin VB.TextBox txtResponsavel 
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
            Left            =   1555
            Locked          =   -1  'True
            TabIndex        =   5
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pelo cadastro."
            Top             =   405
            Width           =   6315
         End
         Begin VB.TextBox txtData 
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
            TabIndex        =   4
            TabStop         =   0   'False
            ToolTipText     =   "Data do cadastro."
            Top             =   405
            Width           =   1365
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Alíq. int. (%)"
            BeginProperty Font 
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
            Left            =   9315
            TabIndex        =   41
            Top             =   210
            Width           =   930
         End
         Begin VB.Label Label29 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "FCP"
            BeginProperty Font 
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
            Left            =   8700
            TabIndex        =   40
            Top             =   210
            Width           =   285
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
            Index           =   9
            Left            =   4255
            TabIndex        =   30
            Top             =   210
            Width           =   915
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
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
            Left            =   690
            TabIndex        =   29
            Top             =   210
            Width           =   345
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
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
            Left            =   8085
            TabIndex        =   22
            Top             =   210
            Width           =   195
         End
      End
      Begin MSComctlLib.ListView Lista1 
         Height          =   4035
         Left            =   75
         TabIndex        =   17
         Top             =   2820
         Width           =   10515
         _ExtentX        =   18547
         _ExtentY        =   7117
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
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
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "N"
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Object.Tag             =   "N"
            Text            =   "NCM"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Grupo"
            Object.Width           =   3822
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Tag             =   "N"
            Text            =   "CST"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Object.Tag             =   "N"
            Text            =   "Margem"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Object.Tag             =   "N"
            Text            =   "Alíquota"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Object.Tag             =   "D"
            Text            =   "Data"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Object.Tag             =   "T"
            Text            =   "Responsável"
            Object.Width           =   3822
         EndProperty
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1485
         Left            =   75
         TabIndex        =   24
         Top             =   1320
         Width           =   10515
         Begin VB.TextBox txtAliqAplic 
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
            Left            =   9510
            MaxLength       =   10
            TabIndex        =   45
            ToolTipText     =   "Alíquota aplicável à operação"
            Top             =   420
            Width           =   840
         End
         Begin VB.CommandButton Cmd_CF 
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
            Left            =   6720
            Picture         =   "frm_Classificacao_Fiscal_regiao.frx":A8BB
            Style           =   1  'Graphical
            TabIndex        =   38
            ToolTipText     =   "Abrir módulo para consulta de classificação fiscal."
            Top             =   420
            Width           =   315
         End
         Begin VB.TextBox Txt_ID_CF 
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
            Left            =   5010
            Locked          =   -1  'True
            TabIndex        =   37
            TabStop         =   0   'False
            ToolTipText     =   "ID da NCM."
            Top             =   420
            Width           =   525
         End
         Begin VB.ComboBox cmbTributaria 
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
            ItemData        =   "frm_Classificacao_Fiscal_regiao.frx":A9BD
            Left            =   180
            List            =   "frm_Classificacao_Fiscal_regiao.frx":A9D6
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   16
            ToolTipText     =   "Tipo antecipação tributária."
            Top             =   1020
            Width           =   10200
         End
         Begin VB.TextBox txtResponsavel1 
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
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   11
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pelo cadastro."
            Top             =   420
            Width           =   3330
         End
         Begin VB.TextBox txtData1 
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
            ToolTipText     =   "Data do cadastro."
            Top             =   420
            Width           =   885
         End
         Begin VB.TextBox txtAliquota 
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
            Left            =   8670
            MaxLength       =   10
            TabIndex        =   15
            ToolTipText     =   "Alíquota interna do ICMS."
            Top             =   420
            Width           =   840
         End
         Begin VB.TextBox txtMargem 
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
            Left            =   7845
            MaxLength       =   10
            TabIndex        =   14
            ToolTipText     =   "Margem do valor agregado (MVA)."
            Top             =   420
            Width           =   810
         End
         Begin VB.ComboBox cmbCST 
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
            ItemData        =   "frm_Classificacao_Fiscal_regiao.frx":AC97
            Left            =   7065
            List            =   "frm_Classificacao_Fiscal_regiao.frx":AD58
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   13
            ToolTipText     =   "Código de situação tributária"
            Top             =   420
            Width           =   780
         End
         Begin ControlesUteis.txt txtUF1 
            Height          =   360
            Left            =   4425
            TabIndex        =   12
            TabStop         =   0   'False
            ToolTipText     =   "UF."
            Top             =   420
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   635
            Tamanho         =   570
            Text            =   ""
            FocusColor      =   16777215
            CaptionColor    =   0
            ShowCaption     =   0   'False
            Caption         =   ""
            Locked          =   -1  'True
            MaxLength       =   2
            BackColor       =   14737632
            Negative        =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   0
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
            Left            =   5550
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   36
            TabStop         =   0   'False
            ToolTipText     =   "Classificação fiscal."
            Top             =   420
            Width           =   1155
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Alíq. aplic."
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
            Left            =   9555
            TabIndex        =   46
            Top             =   210
            Width           =   765
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
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
            Left            =   5190
            TabIndex        =   42
            Top             =   210
            Width           =   165
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
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
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   5955
            TabIndex        =   39
            Top             =   210
            Width           =   330
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo antecipação tributária"
            BeginProperty Font 
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
            Left            =   4290
            TabIndex        =   35
            Top             =   810
            Width           =   1920
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
            Index           =   0
            Left            =   2295
            TabIndex        =   32
            Top             =   210
            Width           =   915
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
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
            Left            =   450
            TabIndex        =   31
            Top             =   210
            Width           =   345
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Alíq. int."
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
            Left            =   8850
            TabIndex        =   28
            Top             =   210
            Width           =   615
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
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
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   4620
            TabIndex        =   27
            Top             =   210
            Width           =   195
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "MVA (%)"
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
            Left            =   7965
            TabIndex        =   26
            Top             =   210
            Width           =   660
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "CST"
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
            Left            =   7320
            TabIndex        =   25
            Top             =   210
            Width           =   285
         End
      End
      Begin MSComctlLib.ListView Lista 
         Height          =   4115
         Left            =   -74925
         TabIndex        =   9
         Top             =   2745
         Width           =   10515
         _ExtentX        =   18547
         _ExtentY        =   7250
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
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
            Alignment       =   2
            SubItemIndex    =   1
            Object.Tag             =   "T"
            Text            =   "UF"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Object.Width           =   9959
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Object.Tag             =   "D"
            Text            =   "Data"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "Responsável"
            Object.Width           =   3881
         EndProperty
      End
   End
End
Attribute VB_Name = "frm_Classificacao_Fiscal_regiao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Novo_CF_Regiao  As Boolean 'OK
Dim Novo_CF_Regiao1 As Boolean 'OK

Private Sub cmbCST_Click()
On Error GoTo tratar_erro

txtMargem = ""
txtAliquota = ""
cmbTributaria.ListIndex = -1
Set TBFI = CreateObject("adodb.recordset")
TBFI.Open "Select * from cst where id_uf = " & txtId & " and ID_CF = " & IIf(Txt_ID_CF = "", 0, Txt_ID_CF) & " and cst = '" & cmbCST & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBFI.EOF = False Then
    txtMargem = IIf(IsNull(TBFI!Margem), "", TBFI!Margem)
    txtAliquota = IIf(IsNull(TBFI!Aliquota), "", TBFI!Aliquota)
    NomeCampo = "Não foi encontrado o campo tipo antecipação tributária"
    If IsNull(TBFI!Tipo) = False And TBFI!Tipo <> "" Then cmbTributaria = TBFI!Tipo
End If
TBFI.Close

Exit Sub
tratar_erro:
    If Err.Number = "383" Then
        USMsgBox (NomeCampo & ", favor alterar."), vbExclamation, "CAPRIND v5.0"
        Exit Sub
    End If
End Sub

Private Sub ProcExcluir()
On Error GoTo tratar_erro

Permitido = False
With Lista
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir esta(s) UF('s)?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select * from regioes WHERE ID = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False Then
                '==================================
                Modulo = "Fiscal/Classificação fiscal/Cadastro de regiões"
                Evento = "Excluir"
                ID_documento = .ListItems(InitFor)
                Documento = "UF: " & TBFI!UF
                Documento1 = ""
                ProcGravaEvento
                '==================================
                Conexao.Execute "DELETE from regioes where id = " & .ListItems(InitFor)
                Conexao.Execute "DELETE from CST where ID_UF = " & .ListItems(InitFor)
            End If
            TBFI.Close
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe a(s) UF('s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("UF('s) excluída(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCampos
    ProcCarregaLista
    Frame1.Enabled = False
    Novo_CF_Regiao = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluirSub()
On Error GoTo tratar_erro

Permitido = False
With Lista1
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir esta(s) substituição(ões) tributária?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select * from CST WHERE ID = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False Then
                '==================================
                Modulo = "Fiscal/Classificação fiscal/Cadastro de regiões"
                Evento = "Excluir substituição tributária"
                ID_documento = .ListItems(InitFor)
                Documento = "UF: " & Lista.SelectedItem.ListSubItems(1)
                Documento1 = "CST: " & TBFI!CST
                ProcGravaEvento
                '==================================
                Conexao.Execute "DELETE from CST where ID = " & .ListItems(InitFor)
            End If
            TBFI.Close
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe a(s) substituição(ões) tributária antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Substituição(ões) tributária excluída(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCampos1
    ProcCarregaLista1
    Frame4.Enabled = False
    Novo_CF_Regiao1 = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovo()
On Error GoTo tratar_erro

ProcLimpaCampos
Novo_CF_Regiao = True
Frame1.Enabled = True
txtuf.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovoSub()
On Error GoTo tratar_erro

ProcLimpaCampos1
Frame4.Enabled = True
Cmd_CF_Click
Novo_CF_Regiao1 = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

If Novo_CF_Regiao = True Then
    If USMsgBox("A UF ainda não foi salva, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcGravar
        If Novo_CF_Regiao = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
If Novo_CF_Regiao1 = True Then
    If USMsgBox("A substituição tributária ainda não foi salva, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcGravarSub
        If Novo_CF_Regiao1 = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
Novo_CF_Regiao = False
Novo_CF_Regiao1 = False
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcGravar()
On Error GoTo tratar_erro

If Frame1.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
Acao = "salvar"
If txtuf.Text = "" Then
    NomeCampo = "a UF"
    ProcVerificaAcao
    txtuf.SetFocus
    Exit Sub
End If
If Txt_FCP <> "" Then
    valor = Txt_FCP
    If valor > 2 Then
        USMsgBox ("O FCP não pode ser maior que 2%."), vbExclamation, "CAPRIND v5.0"
        Txt_FCP.SetFocus
        Exit Sub
    End If
End If
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from regioes where id = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then TBGravar.AddNew
If txtData <> "" Then TBGravar!Data = txtData Else TBGravar!Data = Date
If txtResponsavel <> "" Then TBGravar!Responsavel = txtResponsavel Else TBGravar!Responsavel = pubUsuario
If OptDE.Value = True Then TBGravar!regiao = "DE"
If OptSS.Value = True Then TBGravar!regiao = "SS"
If OptNN.Value = True Then TBGravar!regiao = "NN"
If OptCO.Value = True Then TBGravar!regiao = "CO"
TBGravar!UF = txtuf.Text
TBGravar!FCP = IIf(Txt_FCP = "", Null, Txt_FCP)
TBGravar!ICMS_interno = IIf(Txt_aliquota_interna_ICMS = "", Null, Txt_aliquota_interna_ICMS)
TBGravar.Update
txtId = TBGravar!ID
TBGravar.Close
ProcCarregaLista
If Novo_CF_Regiao = True Then
    USMsgBox ("Nova UF cadastrada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar"
    If CodigoLista <> 0 And Lista.ListItems.Count <> 0 Then
        Lista.SelectedItem = Lista.ListItems(CodigoLista)
1:
        Lista.SetFocus
    End If
End If
'==================================
Modulo = "Fiscal/Classificação fiscal/Cadastro de regiões"
ID_documento = txtId
Documento = "UF: " & txtuf.Text
Documento1 = ""
ProcGravaEvento
'==================================
Novo_CF_Regiao = False

Exit Sub
tratar_erro:
    If Err.Number = "35600" Then GoTo 1
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcGravarSub()
On Error GoTo tratar_erro

If Frame4.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
Acao = "salvar"
If Txt_CF = "" Then
    NomeCampo = "a classificação fiscal"
    ProcVerificaAcao
    Txt_CF.SetFocus
    Exit Sub
End If
If cmbCST = "" Then
    NomeCampo = "a CST"
    ProcVerificaAcao
    cmbCST.SetFocus
    Exit Sub
End If
If IsNumeric(txtMargem) = False And txtMargem <> "" Then
    NomeCampo = "a margem"
    ProcVerificaAcao
    txtMargem.SetFocus
    Exit Sub
End If
If IsNumeric(txtAliquota) = False And txtAliquota <> "" Then
    NomeCampo = "a alíquota"
    ProcVerificaAcao
    txtAliquota.SetFocus
    Exit Sub
End If
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from CST where id = " & txtId1, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then TBGravar.AddNew
If txtData1 <> "" Then TBGravar!Data = txtData1 Else TBGravar!Data = Date
If txtResponsavel1 <> "" Then TBGravar!Responsavel = txtResponsavel1 Else TBGravar!Responsavel = pubUsuario
TBGravar!ID_UF = txtId
TBGravar!ID_CF = Txt_ID_CF
TBGravar!CST = cmbCST
TBGravar!Margem = IIf(txtMargem = "", "0", txtMargem)
TBGravar!Aliquota = IIf(txtAliquota = "", "0", txtAliquota)
TBGravar!Aliquota_aplicavel = IIf(txtAliqAplic = "", "0", txtAliqAplic)

TBGravar!Tipo = IIf(cmbTributaria = "", Null, cmbTributaria)
TBGravar.Update
txtId1 = TBGravar!ID
TBGravar.Close
ProcCarregaLista1
If Novo_CF_Regiao1 = True Then
    USMsgBox ("Nova CST cadastrada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Nova substituição tributária"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar substituição tributária"
    If CodigoLista1 <> 0 And Lista1.ListItems.Count <> 0 Then
        Lista1.SelectedItem = Lista1.ListItems(CodigoLista1)
        Lista1.SetFocus
    End If
End If
'==================================
Modulo = "Fiscal/Classificação fiscal/Cadastro de regiões"
ID_documento = txtId1
Documento = "UF: " & txtuf.Text
Documento1 = "CST: " & cmbCST
ProcGravaEvento
'==================================
Novo_CF_Regiao1 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_CF_Click()
On Error GoTo tratar_erro

Vendas_Proposta = False
Vendas_PI = False
Faturamento = False
Clientes = False
Compras_Pedido = False
Familia_NCM = False
ClassFiscal = True
frmProj_Classificacao_Fiscal.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyInsert: If SSTab1.Tab = 0 Then ProcNovo Else ProcNovoSub
    Case vbKeyF3: If SSTab1.Tab = 0 Then ProcGravar Else ProcGravarSub
    Case vbKeyF4: If SSTab1.Tab = 0 Then ProcExcluir Else ProcExcluirSub
    Case vbKeyEscape: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 10515, 7, True

SSTab1.Tab = 0
OptDE.Value = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaLista()
On Error GoTo tratar_erro

Lista.ListItems.Clear
If OptDE.Value = True Then TextoFiltro = "DE"
If OptSS.Value = True Then TextoFiltro = "SS"
If OptNN.Value = True Then TextoFiltro = "NN"
If OptCO.Value = True Then TextoFiltro = "CO"
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from regioes where regiao = '" & TextoFiltro & "' order by UF", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    With Lista.ListItems
        Do While TBLISTA.EOF = False
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!UF), "", TBLISTA!UF)
            If OptDE.Value = True Then .Item(.Count).SubItems(2) = "Dentro do Estado"
            If OptSS.Value = True Then .Item(.Count).SubItems(2) = "Sul e Suldeste"
            If OptNN.Value = True Then .Item(.Count).SubItems(2) = "Norte e Nordeste"
            If OptCO.Value = True Then .Item(.Count).SubItems(2) = "Centro Oeste"
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Data), "", Format(TBLISTA!Data, "dd/mm/yy"))
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!Responsavel), "", TBLISTA!Responsavel)
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

Sub ProcCarregaLista1()
On Error GoTo tratar_erro

Lista1.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select CST.*, CF.IDIntClasse, CF.txt_Grupo from CST LEFT JOIN tbl_ClassificacaoFiscal CF ON CF.Idclass = CST.ID_CF where CST.ID_UF = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    With Lista1.ListItems
        Do While TBLISTA.EOF = False
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!IDIntClasse), "", TBLISTA!IDIntClasse)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Txt_grupo), "", TBLISTA!Txt_grupo)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!CST), "", TBLISTA!CST)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!Margem), "", TBLISTA!Margem)
            .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!Aliquota), "", TBLISTA!Aliquota)
            .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA!Data), "", Format(TBLISTA!Data, "dd/mm/yy"))
            .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA!Responsavel), "", TBLISTA!Responsavel)
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

Private Sub Lista_Click()
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Regioes where Id = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    ProcLimpaCampos
    txtId = TBLISTA!ID
    txtData = IIf(IsNull(TBLISTA!Data), "", Format(TBLISTA!Data, "dd/mm/yy"))
    txtResponsavel = IIf(IsNull(TBLISTA!Responsavel), "", TBLISTA!Responsavel)
    txtuf = IIf(IsNull(TBLISTA!UF), "", TBLISTA!UF)
    Txt_FCP = IIf(IsNull(TBLISTA!FCP), "", Format(TBLISTA!FCP, "###,##0.0000"))
    Txt_aliquota_interna_ICMS = IIf(IsNull(TBLISTA!ICMS_interno), "", Format(TBLISTA!ICMS_interno, "###,##0.00"))
    Frame1.Enabled = True
    Novo_CF_Regiao = False
    CodigoLista = Lista.SelectedItem.index
End If

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

Private Sub Lista1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView Lista1, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista1_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista1.ListItems.Count = 0 Then Exit Sub
ProcLimpaCampos1
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from CST where id = " & Lista1.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    ProcCarregaDadosST
    CodigoLista1 = Lista1.SelectedItem.index
    Frame4.Enabled = True
End If
TBLISTA.Close
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaDadosST()
On Error GoTo tratar_erro

txtId1.Text = TBLISTA!ID
txtData1 = IIf(IsNull(TBLISTA!Data), "", Format(TBLISTA!Data, "dd/mm/yy"))
txtResponsavel1 = IIf(IsNull(TBLISTA!Responsavel), "", TBLISTA!Responsavel)
txtAliqAplic.Text = IIf(IsNull(TBLISTA!Aliquota_aplicavel), "", TBLISTA!Aliquota_aplicavel)

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Regioes where Id = " & IIf(IsNull(TBLISTA!ID_UF), 0, TBLISTA!ID_UF), Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    txtUF1.Text = IIf(IsNull(TBAbrir!UF), "", TBAbrir!UF)
End If

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from tbl_ClassificacaoFiscal where Idclass = " & IIf(IsNull(TBLISTA!ID_CF), 0, TBLISTA!ID_CF), Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    Txt_ID_CF = TBAbrir!Idclass
    Txt_CF = IIf(IsNull(TBAbrir!IDIntClasse), "", TBAbrir!IDIntClasse)
End If
TBAbrir.Close

NomeCampo = "A CST " & TBLISTA!CST & " não é com substituição tributária"
If IsNull(TBLISTA!CST) = False And TBLISTA!CST <> "" Then cmbCST = TBLISTA!CST
1:
    Novo_CF_Regiao1 = False

Exit Sub
tratar_erro:
    If Err.Number = "383" Then
        USMsgBox (NomeCampo & ", favor alterar."), vbExclamation, "CAPRIND v5.0"
        GoTo 1
    End If
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub OptCO_Click()
On Error GoTo tratar_erro

If OptCO.Value = True Then
    Frame1.Enabled = False
    ProcLimpaCampos
    ProcCarregaLista
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub OptDE_Click()
On Error GoTo tratar_erro

If OptDE.Value = True Then
    Frame1.Enabled = False
    ProcLimpaCampos
    ProcCarregaLista
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub OptNN_Click()
On Error GoTo tratar_erro

If OptNN.Value = True Then
    Frame1.Enabled = False
    ProcLimpaCampos
    ProcCarregaLista
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub OptSS_Click()
On Error GoTo tratar_erro

If OptSS.Value = True Then
    Frame1.Enabled = False
    ProcLimpaCampos
    ProcCarregaLista
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCampos()
On Error GoTo tratar_erro

txtId = 0
txtData = Format(Date, "dd/mm/yy")
txtResponsavel = pubUsuario
txtuf.Text = ""
CodigoLista = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCampos1()
On Error GoTo tratar_erro

txtId1 = 0
txtData1 = Format(Date, "dd/mm/yy")
txtResponsavel1 = pubUsuario
txtUF1.Text = txtuf.Text
Txt_ID_CF = ""
Txt_CF = ""
cmbCST.ListIndex = -1
txtMargem = ""
txtAliquota = ""
cmbTributaria.ListIndex = -1
CodigoLista1 = 0

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
        If Lista.Visible = True Then Lista.SetFocus
        ProcCarregaLista
    Case 1:
        If Novo_CF_Regiao = True Then
            USMsgBox ("Salve a UF antes de prosseguir."), vbExclamation, "CAPRIND v5.0"
            SSTab1.Tab = 0
            Exit Sub
        End If
        Lista1.SetFocus
        ProcLimpaCampos1
        txtUF1.Text = txtuf.Text
        ProcCarregaLista1
End Select
        
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_FCP_Change()
On Error GoTo tratar_erro

If Txt_FCP <> "" Then
    VerifNumero = Txt_FCP
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_FCP = ""
        Txt_FCP.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_FCP_LostFocus()
On Error GoTo tratar_erro

Txt_FCP = Format(Txt_FCP, "###,##0.0000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_aliquota_interna_ICMS_Change()
On Error GoTo tratar_erro

If Txt_aliquota_interna_ICMS <> "" Then
    VerifNumero = Txt_aliquota_interna_ICMS
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_aliquota_interna_ICMS = ""
        Txt_aliquota_interna_ICMS.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_aliquota_interna_ICMS_LostFocus()
On Error GoTo tratar_erro

Txt_aliquota_interna_ICMS = Format(Txt_aliquota_interna_ICMS, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtAliquota_LostFocus()
On Error GoTo tratar_erro

If txtAliquota.Text <> "" Then
    VerifNumero = txtAliquota.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtAliquota.Text = ""
        txtAliquota.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtMargem_LostFocus()
    On Error GoTo tratar_erro

If txtMargem.Text <> "" Then
    VerifNumero = txtMargem.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtMargem.Text = ""
        txtMargem.SetFocus
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
    Case 1: If SSTab1.Tab = 0 Then ProcNovo Else ProcNovoSub
    Case 2: If SSTab1.Tab = 0 Then ProcGravar Else ProcGravarSub
    Case 3: If SSTab1.Tab = 0 Then ProcExcluir Else ProcExcluirSub
    'Case 5: ProcAjuda
    Case 6: ProcSair
End Select
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
