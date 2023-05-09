VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm_centro_de_custo 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Administrativo - Custos - Centro de custo"
   ClientHeight    =   10035
   ClientLeft      =   120
   ClientTop       =   525
   ClientWidth     =   15360
   ControlBox      =   0   'False
   Icon            =   "Frm_centro_de_custo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MousePointer    =   99  'Custom
   ScaleHeight     =   10035
   ScaleWidth      =   15360
   WindowState     =   2  'Maximized
   Begin DrawSuite2022.USProgressBar PBLista1 
      Height          =   255
      Left            =   90
      TabIndex        =   107
      Top             =   9720
      Visible         =   0   'False
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
      ItemData        =   "Frm_centro_de_custo.frx":000C
      Left            =   240
      List            =   "Frm_centro_de_custo.frx":000E
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Empresa."
      Top             =   1690
      Width           =   3255
   End
   Begin DrawSuite2022.USToolBar USToolBar2 
      Height          =   975
      Left            =   60
      TabIndex        =   81
      Top             =   330
      Visible         =   0   'False
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   1720
      ButtonCount     =   10
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
      ButtonEnabled7  =   0   'False
      ButtonIconSize7 =   32
      ButtonAlignment7=   2
      ButtonType7     =   1
      ButtonStyle7    =   -1
      BeginProperty ButtonFont7 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState7    =   -1
      ButtonLeft7     =   268
      ButtonTop7      =   4
      ButtonWidth7    =   2
      ButtonHeight7   =   54
      ButtonCaption8  =   "Ajuda"
      ButtonEnabled8  =   0   'False
      ButtonIconSize8 =   32
      ButtonToolTipText8=   "Ajuda (F1)"
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
      ButtonLeft8     =   272
      ButtonTop8      =   2
      ButtonWidth8    =   36
      ButtonHeight8   =   21
      ButtonUseMaskColor8=   0   'False
      ButtonCaption9  =   "Sair"
      ButtonEnabled9  =   0   'False
      ButtonIconSize9 =   32
      ButtonToolTipText9=   "Sair (Esc)"
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
      ButtonLeft9     =   310
      ButtonTop9      =   2
      ButtonWidth9    =   26
      ButtonHeight9   =   21
      ButtonUseMaskColor9=   0   'False
      ButtonEnabled10 =   0   'False
      ButtonIconSize10=   32
      ButtonKey10     =   "10"
      ButtonAlignment10=   2
      BeginProperty ButtonFont10 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState10   =   5
      ButtonLeft10    =   338
      ButtonTop10     =   2
      ButtonWidth10   =   24
      ButtonHeight10  =   24
      ButtonUseMaskColor10=   0   'False
      Begin DrawSuite2022.USImageList USImageList2 
         Left            =   10560
         Top             =   210
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "Frm_centro_de_custo.frx":0010
         Count           =   1
      End
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
      Tabs            =   5
      TabsPerRow      =   5
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
      TabCaption(0)   =   "Centro de custo"
      TabPicture(0)   =   "Frm_centro_de_custo.frx":53F4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "USToolBar1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Lista"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "TXTID"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame9"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Responsáveis pelo centro de custo"
      TabPicture(1)   =   "Frm_centro_de_custo.frx":5410
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "TxtID_resp"
      Tab(1).Control(1)=   "Frame4"
      Tab(1).Control(2)=   "Lista1"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Previsão orçamentária"
      TabPicture(2)   =   "Frm_centro_de_custo.frx":542C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame6"
      Tab(2).Control(1)=   "Frame1"
      Tab(2).Control(2)=   "Txt_ID_prev"
      Tab(2).Control(3)=   "USImageList3"
      Tab(2).Control(4)=   "Lista2"
      Tab(2).Control(5)=   "USToolBar3"
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "Depreciação"
      TabPicture(3)   =   "Frm_centro_de_custo.frx":5448
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame7"
      Tab(3).Control(1)=   "Txt_ID_depreciacao"
      Tab(3).Control(2)=   "Frame3"
      Tab(3).Control(3)=   "Lista3"
      Tab(3).ControlCount=   4
      TabCaption(4)   =   "Rateio"
      TabPicture(4)   =   "Frm_centro_de_custo.frx":5464
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Lista4"
      Tab(4).Control(1)=   "Lista5"
      Tab(4).Control(2)=   "Txt_ID_rateio"
      Tab(4).Control(3)=   "Frame5"
      Tab(4).Control(4)=   "Frame8"
      Tab(4).ControlCount=   5
      Begin VB.Frame Frame9 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   60
         TabIndex        =   104
         Top             =   9540
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
            ItemData        =   "Frm_centro_de_custo.frx":5480
            Left            =   13110
            List            =   "Frm_centro_de_custo.frx":548A
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   60
            Width           =   1965
         End
         Begin DrawSuite2022.USProgressBar PBLista 
            Height          =   255
            Left            =   30
            TabIndex        =   106
            Top             =   90
            Width           =   11505
            _ExtentX        =   20294
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
            Left            =   11730
            TabIndex        =   105
            Top             =   120
            Width           =   1260
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Filtrar rateios"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   -74925
         TabIndex        =   100
         Top             =   5610
         Width           =   15195
         Begin VB.OptionButton OptAteomes2 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Até o mês"
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
            Left            =   1020
            TabIndex        =   54
            Top             =   270
            Width           =   1035
         End
         Begin VB.OptionButton OptDomes2 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Do mês"
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
            Left            =   150
            TabIndex        =   53
            Top             =   270
            Value           =   -1  'True
            Width           =   825
         End
         Begin VB.ComboBox cmbAno2 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            ItemData        =   "Frm_centro_de_custo.frx":549F
            Left            =   10170
            List            =   "Frm_centro_de_custo.frx":54A1
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   56
            Top             =   240
            Width           =   855
         End
         Begin MSComctlLib.TabStrip TabFiltro2 
            Height          =   345
            Left            =   2160
            TabIndex        =   55
            Top             =   240
            Width           =   8115
            _ExtentX        =   14314
            _ExtentY        =   609
            TabWidthStyle   =   1
            MultiRow        =   -1  'True
            TabMinWidth     =   1177
            _Version        =   393216
            BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
               NumTabs         =   12
               BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Jan"
                  Key             =   "Jan"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Fev"
                  Key             =   "Fev"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Mar"
                  Key             =   "Mar"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Abril"
                  Key             =   "Abr"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Maio"
                  Key             =   "Maio"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Jun"
                  Key             =   "Jun"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Jul"
                  Key             =   "Jul"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Ago"
                  Key             =   "Ago"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab9 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Set"
                  Key             =   "Set"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab10 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Out"
                  Key             =   "Out"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab11 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Nov"
                  Key             =   "Nov"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab12 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Dez"
                  Key             =   "Dez"
                  ImageVarType    =   2
               EndProperty
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Filtrar depreciações"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   -74925
         TabIndex        =   99
         Top             =   2160
         Width           =   15195
         Begin VB.ComboBox cmbAno1 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            ItemData        =   "Frm_centro_de_custo.frx":54A3
            Left            =   10170
            List            =   "Frm_centro_de_custo.frx":54A5
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton OptDomes1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Do mês"
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
            Left            =   150
            TabIndex        =   37
            Top             =   270
            Value           =   -1  'True
            Width           =   825
         End
         Begin VB.OptionButton OptAteomes1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Até o mês"
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
            Left            =   1020
            TabIndex        =   38
            Top             =   270
            Width           =   1035
         End
         Begin MSComctlLib.TabStrip TabFiltro1 
            Height          =   345
            Left            =   2160
            TabIndex        =   39
            Top             =   240
            Width           =   8115
            _ExtentX        =   14314
            _ExtentY        =   609
            TabWidthStyle   =   1
            MultiRow        =   -1  'True
            TabMinWidth     =   1177
            _Version        =   393216
            BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
               NumTabs         =   12
               BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Jan"
                  Key             =   "Jan"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Fev"
                  Key             =   "Fev"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Mar"
                  Key             =   "Mar"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Abril"
                  Key             =   "Abr"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Maio"
                  Key             =   "Maio"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Jun"
                  Key             =   "Jun"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Jul"
                  Key             =   "Jul"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Ago"
                  Key             =   "Ago"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab9 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Set"
                  Key             =   "Set"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab10 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Out"
                  Key             =   "Out"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab11 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Nov"
                  Key             =   "Nov"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab12 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Dez"
                  Key             =   "Dez"
                  ImageVarType    =   2
               EndProperty
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Filtrar previsões"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   -74925
         TabIndex        =   98
         Top             =   2160
         Width           =   15195
         Begin VB.ComboBox Cmb_revisao_filtrar 
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
            ItemData        =   "Frm_centro_de_custo.frx":54A7
            Left            =   12660
            List            =   "Frm_centro_de_custo.frx":54CF
            Style           =   2  'Dropdown List
            TabIndex        =   28
            ToolTipText     =   "Revisão."
            Top             =   180
            Width           =   1065
         End
         Begin VB.OptionButton OptAteomes 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Até o mês"
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
            Left            =   1020
            TabIndex        =   25
            Top             =   270
            Width           =   1035
         End
         Begin VB.OptionButton OptDomes 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Do mês"
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
            Left            =   150
            TabIndex        =   24
            Top             =   270
            Value           =   -1  'True
            Width           =   825
         End
         Begin VB.ComboBox cmbAno 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            ItemData        =   "Frm_centro_de_custo.frx":54FC
            Left            =   10170
            List            =   "Frm_centro_de_custo.frx":54FE
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   240
            Width           =   855
         End
         Begin MSComctlLib.TabStrip TabFiltro 
            Height          =   345
            Left            =   2160
            TabIndex        =   26
            Top             =   240
            Width           =   8115
            _ExtentX        =   14314
            _ExtentY        =   609
            TabWidthStyle   =   1
            MultiRow        =   -1  'True
            TabMinWidth     =   1177
            _Version        =   393216
            BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
               NumTabs         =   12
               BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Jan"
                  Key             =   "Jan"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Fev"
                  Key             =   "Fev"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Mar"
                  Key             =   "Mar"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Abril"
                  Key             =   "Abr"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Maio"
                  Key             =   "Maio"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Jun"
                  Key             =   "Jun"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Jul"
                  Key             =   "Jul"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Ago"
                  Key             =   "Ago"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab9 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Set"
                  Key             =   "Set"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab10 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Out"
                  Key             =   "Out"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab11 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Nov"
                  Key             =   "Nov"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab12 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Dez"
                  Key             =   "Dez"
                  ImageVarType    =   2
               EndProperty
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Revisão:"
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
            Left            =   11925
            TabIndex        =   101
            Top             =   240
            Width           =   675
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   840
         Left            =   -74925
         TabIndex        =   90
         Top             =   4770
         Width           =   15195
         Begin VB.TextBox Txt_ID_PC_rateio 
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
            Left            =   5190
            MaxLength       =   255
            MouseIcon       =   "Frm_centro_de_custo.frx":5500
            MousePointer    =   99  'Custom
            TabIndex        =   44
            Text            =   "0"
            ToolTipText     =   "ID PC."
            Top             =   375
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.TextBox Txt_percentual_rateio 
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
            Left            =   13830
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   51
            TabStop         =   0   'False
            ToolTipText     =   "Percentual."
            Top             =   375
            Width           =   1155
         End
         Begin VB.CheckBox Chk_valor_rateio 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Valor"
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
            Left            =   12905
            TabIndex        =   48
            Top             =   180
            Width           =   675
         End
         Begin VB.CheckBox Chk_percentual_rateio 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Percentual"
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
            Left            =   13875
            TabIndex        =   50
            Top             =   180
            Width           =   1065
         End
         Begin VB.TextBox Txt_valor_rateio 
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
            Left            =   12665
            Locked          =   -1  'True
            TabIndex        =   49
            TabStop         =   0   'False
            ToolTipText     =   "Valor."
            Top             =   375
            Width           =   1155
         End
         Begin VB.TextBox txtResponsavel_rateio 
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
            TabIndex        =   43
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pelo cadastro."
            Top             =   375
            Width           =   3795
         End
         Begin VB.TextBox Txt_codigo_PC_rateio 
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
            Left            =   5190
            Locked          =   -1  'True
            MaxLength       =   255
            TabIndex        =   45
            TabStop         =   0   'False
            ToolTipText     =   "Código contábil."
            Top             =   375
            Width           =   1870
         End
         Begin VB.CommandButton Cmd_localizar_PC_rateio 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   12225
            Picture         =   "Frm_centro_de_custo.frx":580A
            Style           =   1  'Graphical
            TabIndex        =   47
            ToolTipText     =   "Localizar plano de contas."
            Top             =   375
            Width           =   315
         End
         Begin VB.TextBox Txt_descricao_PC_rateio 
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
            Left            =   7080
            Locked          =   -1  'True
            MaxLength       =   255
            TabIndex        =   46
            TabStop         =   0   'False
            ToolTipText     =   "Descrição"
            Top             =   375
            Width           =   5115
         End
         Begin MSComCtl2.DTPicker txtData_rateio 
            Height          =   315
            Left            =   180
            TabIndex        =   42
            ToolTipText     =   "Data do rateio."
            Top             =   375
            Width           =   1200
            _ExtentX        =   2117
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
            Format          =   197984259
            CurrentDate     =   39057
         End
         Begin VB.Label Label8 
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
            Left            =   608
            TabIndex        =   94
            Top             =   180
            Width           =   345
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
            Index           =   6
            Left            =   2820
            TabIndex        =   93
            Top             =   180
            Width           =   915
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Código contábil"
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
            Left            =   5580
            TabIndex        =   92
            Top             =   180
            Width           =   2115
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
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
            Index           =   7
            Left            =   9285
            TabIndex        =   91
            Top             =   180
            Width           =   705
            WordWrap        =   -1  'True
         End
      End
      Begin VB.TextBox Txt_ID_rateio 
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   -73455
         TabIndex        =   89
         Text            =   "0"
         Top             =   6270
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox Txt_ID_depreciacao 
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   -73440
         TabIndex        =   85
         Text            =   "0"
         Top             =   3720
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   840
         Left            =   -74925
         TabIndex        =   82
         Top             =   1320
         Width           =   15195
         Begin VB.TextBox Txt_ID_PC_depreciacao 
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
            Left            =   5190
            MaxLength       =   255
            MouseIcon       =   "Frm_centro_de_custo.frx":590C
            MousePointer    =   99  'Custom
            TabIndex        =   31
            Text            =   "0"
            ToolTipText     =   "ID PC."
            Top             =   375
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.TextBox Txt_descricao_PC_depreciacao 
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
            Left            =   7080
            Locked          =   -1  'True
            MaxLength       =   255
            TabIndex        =   33
            TabStop         =   0   'False
            ToolTipText     =   "Descrição"
            Top             =   375
            Width           =   6285
         End
         Begin VB.CommandButton Cmd_localizar_PC_depreciacao 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   13395
            Picture         =   "Frm_centro_de_custo.frx":5C16
            Style           =   1  'Graphical
            TabIndex        =   34
            ToolTipText     =   "Localizar plano de contas."
            Top             =   375
            Width           =   315
         End
         Begin VB.TextBox Txt_codigo_PC_depreciacao 
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
            Left            =   5190
            Locked          =   -1  'True
            MaxLength       =   255
            TabIndex        =   32
            TabStop         =   0   'False
            ToolTipText     =   "Código contábil."
            Top             =   375
            Width           =   1870
         End
         Begin VB.TextBox txtResponsavel_depreciacao 
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
            TabIndex        =   30
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pelo cadastro."
            Top             =   375
            Width           =   3795
         End
         Begin VB.TextBox Txt_valor_depreciacao 
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
            Left            =   13830
            MaxLength       =   50
            TabIndex        =   35
            ToolTipText     =   "Valor."
            Top             =   375
            Width           =   1155
         End
         Begin MSComCtl2.DTPicker txtData_depreciacao 
            Height          =   315
            Left            =   180
            TabIndex        =   29
            ToolTipText     =   "Data da depreciação."
            Top             =   375
            Width           =   1200
            _ExtentX        =   2117
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
            Format          =   197984259
            CurrentDate     =   39057
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
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
            Left            =   570
            TabIndex        =   95
            Top             =   0
            Width           =   60
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Valor"
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
            Left            =   14235
            TabIndex        =   88
            Top             =   180
            Width           =   345
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
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
            Left            =   9870
            TabIndex        =   87
            Top             =   180
            Width           =   705
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Código contábil"
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
            Left            =   5580
            TabIndex        =   86
            Top             =   180
            Width           =   2115
            WordWrap        =   -1  'True
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
            Index           =   5
            Left            =   2820
            TabIndex        =   84
            Top             =   180
            Width           =   915
         End
         Begin VB.Label Label7 
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
            Left            =   608
            TabIndex        =   83
            Top             =   180
            Width           =   345
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   840
         Left            =   -74925
         TabIndex        =   72
         Top             =   1320
         Width           =   15195
         Begin VB.TextBox Txt_revisao 
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
            Left            =   6020
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   17
            TabStop         =   0   'False
            ToolTipText     =   "Revisão."
            Top             =   375
            Width           =   765
         End
         Begin VB.ComboBox Cmb_mes 
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
            ItemData        =   "Frm_centro_de_custo.frx":5D18
            Left            =   4530
            List            =   "Frm_centro_de_custo.frx":5D40
            Style           =   2  'Dropdown List
            TabIndex        =   15
            ToolTipText     =   "Mês."
            Top             =   375
            Width           =   855
         End
         Begin VB.TextBox Txt_valor 
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
            Height          =   315
            Left            =   13710
            MaxLength       =   20
            TabIndex        =   22
            ToolTipText     =   "Valor."
            Top             =   375
            Width           =   1310
         End
         Begin VB.CommandButton Cmd_localizar_PC 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   13290
            Picture         =   "Frm_centro_de_custo.frx":5D81
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Localizar plano de contas."
            Top             =   375
            Width           =   315
         End
         Begin VB.TextBox Txt_ID_PC 
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
            Left            =   6795
            MaxLength       =   255
            MouseIcon       =   "Frm_centro_de_custo.frx":5E83
            MousePointer    =   99  'Custom
            TabIndex        =   18
            Text            =   "0"
            ToolTipText     =   "ID PC."
            Top             =   375
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.TextBox Txt_descricao_PC 
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
            Left            =   8685
            Locked          =   -1  'True
            MaxLength       =   255
            TabIndex        =   20
            TabStop         =   0   'False
            ToolTipText     =   "Descrição"
            Top             =   375
            Width           =   4605
         End
         Begin VB.TextBox txtResponsavel_prev 
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
            Left            =   1290
            Locked          =   -1  'True
            TabIndex        =   14
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pelo cadastro."
            Top             =   375
            Width           =   3225
         End
         Begin VB.TextBox txtData_prev 
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
            TabIndex        =   13
            TabStop         =   0   'False
            ToolTipText     =   "Data do cadastro."
            Top             =   375
            Width           =   1095
         End
         Begin MSMask.MaskEdBox Txt_ano 
            Height          =   315
            Left            =   5385
            TabIndex        =   16
            ToolTipText     =   "Ano."
            Top             =   375
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   0
            MaxLength       =   4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "####"
            PromptChar      =   "_"
         End
         Begin VB.TextBox Txt_codigo_PC 
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
            Left            =   6795
            Locked          =   -1  'True
            MaxLength       =   255
            TabIndex        =   19
            TabStop         =   0   'False
            ToolTipText     =   "Código contábil."
            Top             =   375
            Width           =   1875
         End
         Begin VB.Label Label11 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Revisão"
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
            Left            =   6125
            TabIndex        =   96
            Top             =   180
            Width           =   555
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mês"
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
            Left            =   4815
            TabIndex        =   79
            Top             =   180
            Width           =   285
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ano"
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
            Left            =   5550
            TabIndex        =   78
            Top             =   180
            Width           =   285
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Valor"
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
            Left            =   14140
            TabIndex        =   77
            Top             =   180
            Width           =   450
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Código contábil"
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
            Left            =   7185
            TabIndex        =   76
            Top             =   180
            Width           =   2115
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
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
            Index           =   4
            Left            =   10635
            TabIndex        =   75
            Top             =   180
            Width           =   705
            WordWrap        =   -1  'True
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
            Index           =   2
            Left            =   2445
            TabIndex        =   74
            Top             =   180
            Width           =   915
         End
         Begin VB.Label Label6 
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
            Left            =   555
            TabIndex        =   73
            Top             =   180
            Width           =   345
         End
      End
      Begin VB.TextBox Txt_ID_prev 
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   -73440
         TabIndex        =   71
         Text            =   "0"
         Top             =   3720
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox TxtID_resp 
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   -73440
         TabIndex        =   69
         Text            =   "0"
         Top             =   3720
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   840
         Left            =   -74925
         TabIndex        =   65
         Top             =   1320
         Width           =   15195
         Begin VB.ComboBox Cmb_resp_CC 
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
            ItemData        =   "Frm_centro_de_custo.frx":618D
            Left            =   5640
            List            =   "Frm_centro_de_custo.frx":618F
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   11
            ToolTipText     =   "Responsável pelo centro de custo."
            Top             =   375
            Width           =   9375
         End
         Begin VB.TextBox txtData_resp 
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
            TabIndex        =   9
            TabStop         =   0   'False
            ToolTipText     =   "Data do cadastro."
            Top             =   375
            Width           =   1095
         End
         Begin VB.TextBox txtResponsavel_resp 
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
            Left            =   1290
            Locked          =   -1  'True
            TabIndex        =   10
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pelo cadastro."
            Top             =   375
            Width           =   4335
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Responsável pelo centro de custo"
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
            Left            =   9112
            TabIndex        =   68
            Top             =   180
            Width           =   2430
         End
         Begin VB.Label Label3 
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
            Left            =   555
            TabIndex        =   67
            Top             =   180
            Width           =   345
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
            Index           =   1
            Left            =   3000
            TabIndex        =   66
            Top             =   180
            Width           =   915
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   1980
         Left            =   75
         TabIndex        =   59
         Top             =   1320
         Width           =   15195
         Begin VB.TextBox txtStatus 
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
            Left            =   11970
            Locked          =   -1  'True
            TabIndex        =   102
            TabStop         =   0   'False
            ToolTipText     =   "Status."
            Top             =   375
            Width           =   1665
         End
         Begin VB.CheckBox Chk_consolidacao 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Consolidação"
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
            Left            =   13710
            TabIndex        =   5
            Top             =   435
            Width           =   1305
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
            Left            =   4230
            Locked          =   -1  'True
            TabIndex        =   2
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pelo cadastro."
            Top             =   375
            Width           =   2325
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
            Left            =   3450
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   1
            TabStop         =   0   'False
            ToolTipText     =   "Data do cadastro."
            Top             =   375
            Width           =   765
         End
         Begin VB.TextBox txt_descricao 
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
            Left            =   7500
            MaxLength       =   50
            TabIndex        =   4
            ToolTipText     =   "Descrição."
            Top             =   375
            Width           =   4455
         End
         Begin VB.TextBox txt_obs 
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
            Height          =   885
            Left            =   180
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   6
            ToolTipText     =   "Observações."
            Top             =   950
            Width           =   14835
         End
         Begin MSMask.MaskEdBox Txt_codigo 
            Height          =   315
            Left            =   6570
            TabIndex        =   3
            ToolTipText     =   "Código."
            Top             =   375
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   0
            MaxLength       =   7
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "###.###"
            PromptChar      =   "_"
         End
         Begin VB.Label Label13 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
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
            Left            =   12570
            TabIndex        =   103
            Top             =   180
            Width           =   465
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
            Left            =   1470
            TabIndex        =   70
            Top             =   180
            Width           =   735
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
            Left            =   4935
            TabIndex        =   64
            Top             =   180
            Width           =   915
         End
         Begin VB.Label Label19 
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
            Left            =   3660
            TabIndex        =   63
            Top             =   180
            Width           =   345
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
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
            Index           =   0
            Left            =   7162
            TabIndex        =   62
            Top             =   750
            Width           =   870
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
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
            Left            =   9382
            TabIndex        =   61
            Top             =   180
            Width           =   690
         End
         Begin VB.Label Label9 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Código"
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
            Left            =   6780
            TabIndex        =   60
            Top             =   180
            Width           =   495
         End
      End
      Begin VB.TextBox TXTID 
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3240
         TabIndex        =   58
         Text            =   "0"
         Top             =   4740
         Visible         =   0   'False
         Width           =   735
      End
      Begin MSComctlLib.ListView Lista 
         Height          =   6210
         Left            =   60
         TabIndex        =   7
         Top             =   3315
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   10954
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
         NumItems        =   11
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "N"
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Tag             =   "T"
            Text            =   "Empresa"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Object.Tag             =   "D"
            Text            =   "Data"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Tag             =   "T"
            Text            =   "Responsável"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "Código"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Object.Tag             =   "T"
            Text            =   "Descrição"
            Object.Width           =   5299
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Object.Tag             =   "N"
            Text            =   "Orçado"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Object.Tag             =   "N"
            Text            =   "Real"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Object.Tag             =   "N"
            Text            =   "Variação"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   9
            Object.Tag             =   "N"
            Text            =   "Percentual"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   10
            Text            =   "Status"
            Object.Width           =   1940
         EndProperty
      End
      Begin MSComctlLib.ListView Lista1 
         Height          =   7530
         Left            =   -74910
         TabIndex        =   12
         Top             =   2175
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   13282
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "N"
            Object.Width           =   529
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
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Tag             =   "T"
            Text            =   "Responsável pelo centro de custo"
            Object.Width           =   19059
         EndProperty
      End
      Begin DrawSuite2022.USToolBar USToolBar1 
         Height          =   975
         Left            =   60
         TabIndex        =   80
         Top             =   330
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   1720
         ButtonCount     =   14
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
         ButtonCaption7  =   "Status"
         ButtonEnabled7  =   0   'False
         ButtonIconSize7 =   32
         ButtonToolTipText7=   "Status (F7)"
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
         ButtonCaption8  =   "Consolidar"
         ButtonEnabled8  =   0   'False
         ButtonIconSize8 =   32
         ButtonToolTipText8=   "Consolidar (F8)"
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
         ButtonWidth8    =   58
         ButtonHeight8   =   21
         ButtonUseMaskColor8=   0   'False
         ButtonCaption9  =   "Visualizar"
         ButtonEnabled9  =   0   'False
         ButtonIconSize9 =   32
         ButtonToolTipText9=   "Visualizar lançamentos realizados (F9)"
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
         ButtonLeft9     =   369
         ButtonTop9      =   2
         ButtonWidth9    =   52
         ButtonHeight9   =   21
         ButtonUseMaskColor9=   0   'False
         ButtonCaption10 =   "Atualizar"
         ButtonEnabled10 =   0   'False
         ButtonToolTipText10=   "Utilizado pelo administrador do sistema."
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
         ButtonLeft10    =   423
         ButtonTop10     =   2
         ButtonWidth10   =   50
         ButtonHeight10  =   21
         ButtonUseMaskColor10=   0   'False
         ButtonEnabled11 =   0   'False
         ButtonIconSize11=   32
         ButtonAlignment11=   2
         ButtonType11    =   1
         ButtonStyle11   =   -1
         BeginProperty ButtonFont11 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState11   =   -1
         ButtonLeft11    =   475
         ButtonTop11     =   4
         ButtonWidth11   =   2
         ButtonHeight11  =   54
         ButtonCaption12 =   "Ajuda"
         ButtonEnabled12 =   0   'False
         ButtonIconSize12=   32
         ButtonToolTipText12=   "Ajuda (F1)"
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
         ButtonLeft12    =   479
         ButtonTop12     =   2
         ButtonWidth12   =   36
         ButtonHeight12  =   21
         ButtonUseMaskColor12=   0   'False
         ButtonCaption13 =   "Sair"
         ButtonEnabled13 =   0   'False
         ButtonIconSize13=   32
         ButtonToolTipText13=   "Sair (Esc)"
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
         ButtonLeft13    =   517
         ButtonTop13     =   2
         ButtonWidth13   =   26
         ButtonHeight13  =   21
         ButtonUseMaskColor13=   0   'False
         ButtonEnabled14 =   0   'False
         ButtonIconSize14=   32
         ButtonKey14     =   "14"
         ButtonAlignment14=   2
         BeginProperty ButtonFont14 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState14   =   5
         ButtonLeft14    =   545
         ButtonTop14     =   2
         ButtonWidth14   =   24
         ButtonHeight14  =   24
         ButtonUseMaskColor14=   0   'False
         Begin DrawSuite2022.USImageList USImageList1 
            Left            =   10020
            Top             =   240
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "Frm_centro_de_custo.frx":6191
            Count           =   1
         End
      End
      Begin MSComctlLib.ListView Lista3 
         Height          =   6960
         Left            =   -74925
         TabIndex        =   36
         Top             =   2745
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   12277
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "N"
            Object.Width           =   529
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
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Tag             =   "T"
            Text            =   "Código contábil"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "Descrição"
            Object.Width           =   13063
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Object.Tag             =   "V"
            Text            =   "Valor"
            Object.Width           =   2117
         EndProperty
      End
      Begin MSComctlLib.ListView Lista5 
         Height          =   3510
         Left            =   -74925
         TabIndex        =   52
         Top             =   6195
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   6191
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
            Object.Tag             =   "D"
            Text            =   "Data"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Responsável"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Tag             =   "N"
            Text            =   "Código"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "Descrição"
            Object.Width           =   5648
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Object.Tag             =   "T"
            Text            =   "Código contábil"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Object.Tag             =   "T"
            Text            =   "Descrição"
            Object.Width           =   5648
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Object.Tag             =   "V"
            Text            =   "Valor"
            Object.Width           =   2117
         EndProperty
      End
      Begin MSComctlLib.ListView Lista4 
         Height          =   3420
         Left            =   -74925
         TabIndex        =   41
         Top             =   1335
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   6033
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
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "N"
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Tag             =   "T"
            Text            =   "Código"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Descrição"
            Object.Width           =   15707
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Object.Tag             =   "N"
            Text            =   "Orçado"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Object.Tag             =   "N"
            Text            =   "Real"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Object.Tag             =   "N"
            Text            =   "Variação"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Object.Tag             =   "N"
            Text            =   "Percentual"
            Object.Width           =   1764
         EndProperty
      End
      Begin DrawSuite2022.USImageList USImageList3 
         Left            =   -68940
         Top             =   5010
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "Frm_centro_de_custo.frx":E1FA
         Count           =   1
      End
      Begin MSComctlLib.ListView Lista2 
         Height          =   6960
         Left            =   -74925
         TabIndex        =   23
         Top             =   2745
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   12277
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
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "N"
            Object.Width           =   529
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
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Object.Tag             =   "T"
            Text            =   "Mês"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Object.Tag             =   "N"
            Text            =   "Ano"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Object.Tag             =   "N"
            Text            =   "Rev."
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Object.Tag             =   "T"
            Text            =   "Código contábil"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Object.Tag             =   "T"
            Text            =   "Descrição"
            Object.Width           =   9446
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Object.Tag             =   "V"
            Text            =   "Valor"
            Object.Width           =   2117
         EndProperty
      End
      Begin DrawSuite2022.USToolBar USToolBar3 
         Height          =   975
         Left            =   -74940
         TabIndex        =   97
         Top             =   330
         Visible         =   0   'False
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
         ButtonToolTipText7=   "Copriar (F7)"
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
         ButtonLeft9     =   355
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
         ButtonLeft10    =   359
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
         ButtonLeft11    =   397
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
         ButtonLeft12    =   425
         ButtonTop12     =   2
         ButtonWidth12   =   24
         ButtonHeight12  =   24
         ButtonUseMaskColor12=   0   'False
      End
   End
End
Attribute VB_Name = "Frm_centro_de_custo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Novo_Centro As Boolean 'OK
Dim Novo_Centro1 As Boolean 'OK
Dim Novo_Centro2 As Boolean 'OK
Dim Novo_Centro3 As Boolean 'OK
Dim Novo_Centro4 As Boolean 'OK
Dim FormulaRel_Centro As String 'OK

Sub ProcAjuda()
On Error GoTo tratar_erro

FunAbrirVideoWeb ("http://www.youtube.com/watch?v=21fQiAzBU6I&list=UUODGiDjQ-BCrxh0YXfJvoqg&index=16&feature=plcp")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 Then Exit Sub
NomeRel = "Custos_centro_de_custo.rpt"

Set TBAcessos = CreateObject("adodb.recordset")
TBAcessos.Open "Select * from Acessos where IDUsuario = " & pubIDUsuario & " and Acesso = 'Custos/Centro de custo/Visualizar todos'", Conexao, adOpenKeyset, adLockOptimistic
If TBAcessos.EOF = False Then
    FormulaRel_Centro = "{Usuarios_setor.ID_empresa} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
Else
    FormulaRel_Centro = "{Usuarios_Setor_Responsavel.Responsavel_CC} = '" & pubUsuario & "' and {Usuarios_setor.ID_empresa} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
End If
TBAcessos.Close

ProcImprimirRel FormulaRel_Centro, ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAnterior()
On Error GoTo tratar_erro

If txtId = 0 Then Exit Sub

Set TBAcessos = CreateObject("adodb.recordset")
TBAcessos.Open "Select * from Acessos where IDUsuario = " & pubIDUsuario & " and Acesso = 'Custos/Centro de custo/Visualizar todos'", Conexao, adOpenKeyset, adLockOptimistic
If TBAcessos.EOF = False Then
    TextoFiltro = "Select * from Usuarios_setor order by Codigo"
Else
    TextoFiltro = "Select Usuarios_setor.* from Usuarios_setor INNER JOIN Usuarios_Setor_Responsavel ON Usuarios_setor.ID = Usuarios_Setor_Responsavel.ID_CC where Usuarios_Setor_Responsavel.Responsavel_CC = '" & pubUsuario & "' order by Usuarios_setor.Codigo"
End If
TBAcessos.Close

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.BOF = False Then
    TBAbrir.Find ("ID = " & txtId)
    TBAbrir.MovePrevious
    If TBAbrir.BOF = False Then
        txtId = TBAbrir!ID
        Set TBLISTA = CreateObject("adodb.recordset")
        TBLISTA.Open "Select * from Usuarios_setor where ID = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
        ProcLimpaCampos
        ProcCarregaDados
        ProcCarregaListaResp
        ProcFiltrarMes
        ProcFiltrarMes1
        ProcFiltrarMes2
    Else
        USMsgBox ("Fim dos cadastros de centro de custo."), vbInformation, "CAPRIND v5.0"
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcProximo()
On Error GoTo tratar_erro

If txtId = 0 Then Exit Sub

Set TBAcessos = CreateObject("adodb.recordset")
TBAcessos.Open "Select * from Acessos where IDUsuario = " & pubIDUsuario & " and Acesso = 'Custos/Centro de custo/Visualizar todos'", Conexao, adOpenKeyset, adLockOptimistic
If TBAcessos.EOF = False Then
    TextoFiltro = "Select * from Usuarios_setor order by Codigo"
Else
    TextoFiltro = "Select Usuarios_setor.* from Usuarios_setor INNER JOIN Usuarios_Setor_Responsavel ON Usuarios_setor.ID = Usuarios_Setor_Responsavel.ID_CC where Usuarios_Setor_Responsavel.Responsavel_CC = '" & pubUsuario & "' order by Usuarios_setor.Codigo"
End If
TBAcessos.Close

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.BOF = False Then
    TBAbrir.Find ("ID = " & txtId)
    TBAbrir.MoveNext
    If TBAbrir.EOF = False Then
        txtId = TBAbrir!ID
        Set TBLISTA = CreateObject("adodb.recordset")
        TBLISTA.Open "Select * from Usuarios_setor where ID = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
        ProcLimpaCampos
        ProcCarregaDados
        ProcCarregaListaResp
        ProcFiltrarMes
        ProcFiltrarMes1
        ProcFiltrarMes2
    Else
        USMsgBox ("Fim dos cadastros de centro de custo."), vbInformation, "CAPRIND v5.0"
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcConsolidar()
On Error GoTo tratar_erro

Frm_centro_de_custo_consolidar.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcVisualizar()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Acao = "visualizar os lançamentos realizados"
If txtId = 0 Then
    NomeCampo = "o centro de custo"
    ProcVerificaAcao
    Exit Sub
End If
Frm_centro_de_custo_visualizar.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAtualizar()
On Error GoTo tratar_erro

If InputBox("Informe a senha para liberar.") = "280362C" Then Frm_centro_de_custo_atualizar.Show 1
            
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcAtualizacao()
On Error GoTo tratar_erro

If USMsgBox("Deseja realmente executar essa(s) atualização(ões)?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    With Frm_centro_de_custo_atualizar
        If .Chk1.Value = 1 Then
            'Atualiza valor do centro de custo de investimento
            Conexao.Execute "DELETE from CC_realizado where ID_CC = 53"
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select CCR.* from (CC_realizado CCR INNER JOIN Usuarios_Setor US ON CCR.ID_CC = US.ID) INNER JOIN projproduto P ON P.Codproduto = CCR.Cod_produto and P.ID_CC <> CCR.ID_CC where P.ID_CC is not null and CCR.ID_estoque is not null and CCR.ID_estoque <> 0 and US.Consolidacao = 'False' order by CCR.ID", Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False Then
                TBFI.MoveLast
                PBLista.Min = 0
                PBLista.Max = TBFI.RecordCount
                PBLista.Value = 1
                Contador = 0
                TBFI.MoveFirst
                Do While TBFI.EOF = False
                    'Verifica se tem CC vinculado ao produto/serviço do pedido de compra e cria débito e crédito
                    If IsNull(TBFI!ID_lista) = False And TBFI!ID_lista <> "0" Then
                        ProcSalvarCCRealizadoInvestAtualizacao TBFI!Data, TBFI!ID_empresa, "Débito", 53, TBFI!Cod_produto, TBFI!ID_PC, TBFI!ID_estoque, TBFI!ID_lista, TBFI!valor, True, False
                        ProcSalvarCCRealizadoInvestAtualizacao TBFI!Data, TBFI!ID_empresa, "Crédito", 53, TBFI!Cod_produto, TBFI!ID_PC, TBFI!ID_estoque, TBFI!ID_lista, TBFI!valor, True, False
                    Else
                        ProcSalvarCCRealizadoInvestAtualizacao TBFI!Data, TBFI!ID_empresa, "Débito", 53, TBFI!Cod_produto, TBFI!ID_PC, TBFI!ID_estoque, IIf(IsNull(TBFI!ID_lista), 0, TBFI!ID_lista), TBFI!valor, True, False
                    End If
                    TBFI.MoveNext
                    Contador = Contador + 1
                    PBLista.Value = Contador
                Loop
            End If
            TBFI.Close
        End If
        
        If .Chk2.Value = 1 Then
            'Atualiza valor utilizado dos centros consolidados
            Conexao.Execute "DELETE from CC_realizado from CC_realizado INNER JOIN Usuarios_Setor ON CC_realizado.ID_CC = Usuarios_Setor.ID Where Usuarios_Setor.Consolidacao = 'True'"
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select * from CC_realizado order by Data, id", Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False Then
                TBFI.MoveLast
                PBLista.Min = 0
                PBLista.Max = TBFI.RecordCount
                PBLista.Value = 1
                Contador = 0
                TBFI.MoveFirst
                Do While TBFI.EOF = False
                    'Verifica movimentação sem origem (Financeiro ou Estoque)
                    Permitido = True
                    If IsNull(TBFI!ID_financeiro) = False And TBFI!ID_financeiro <> 0 Then
                        Set TBContas = CreateObject("adodb.recordset")
                        TBContas.Open "Select * from tbl_ContasPagar where IdIntConta = " & TBFI!ID_financeiro, Conexao, adOpenKeyset, adLockOptimistic
                        If TBContas.EOF = True Then
                            Conexao.Execute "DELETE from CC_realizado where ID_financeiro = " & TBFI!ID_financeiro
                            Permitido = False
                        End If
                        TBContas.Close
                    Else
                        Set TBEstoque = CreateObject("adodb.recordset")
                        TBEstoque.Open "Select * from Estoque_movimentacao where Idoperacao = " & TBFI!ID_estoque, Conexao, adOpenKeyset, adLockOptimistic
                        If TBEstoque.EOF = True Then
                            Conexao.Execute "DELETE from CC_realizado where ID_estoque = " & TBFI!ID_estoque
                            Permitido = False
                        End If
                        TBEstoque.Close
                    End If
                    
                    If Permitido = True Then
                        'AtualizO centro de custo consolidado
                        'Grava movimentação no centro consolidado
                        Set TBAfericao = CreateObject("adodb.recordset")
                        TBAfericao.Open "Select * from Usuarios_Setor_Consolidacao where ID_CC_consolidado = " & TBFI!ID_CC, Conexao, adOpenKeyset, adLockOptimistic
                        If TBAfericao.EOF = False Then
                            Do While TBAfericao.EOF = False
                                Set TBGravar = CreateObject("adodb.recordset")
                                TBGravar.Open "Select * from CC_realizado where ID_CC = " & TBAfericao!ID_CC & " and ID_origem = " & TBFI!ID, Conexao, adOpenKeyset, adLockOptimistic
                                If TBGravar.EOF = True Then TBGravar.AddNew
                                ProcEnviaDadosAtualizacaoCCCons TBAfericao!ID_CC, TBFI!ID
                                TBGravar.Update
                                TBGravar.Close
                                
                                Set TBCiclo = CreateObject("adodb.recordset")
                                TBCiclo.Open "Select * from Usuarios_Setor_Consolidacao where ID_CC_consolidado = " & TBAfericao!ID_CC, Conexao, adOpenKeyset, adLockOptimistic
                                If TBCiclo.EOF = False Then
                                    Do While TBCiclo.EOF = False
                                        Set TBGravar = CreateObject("adodb.recordset")
                                        TBGravar.Open "Select * from CC_realizado where ID_CC = " & TBCiclo!ID_CC & " and ID_origem = " & TBFI!ID, Conexao, adOpenKeyset, adLockOptimistic
                                        If TBGravar.EOF = True Then TBGravar.AddNew
                                        ProcEnviaDadosAtualizacaoCCCons TBCiclo!ID_CC, TBFI!ID
                                        TBGravar.Update
                                        TBGravar.Close
                                        
                                        TBCiclo.MoveNext
                                    Loop
                                End If
                                TBCiclo.Close
                                
                                TBAfericao.MoveNext
                            Loop
                        End If
                        TBAfericao.Close
                    End If
                    
                    TBFI.MoveNext
                    Contador = Contador + 1
                    PBLista.Value = Contador
                Loop
            End If
            TBFI.Close
        End If
        
        If .Chk3.Value = 1 Then
            'Atualiza valor previsto dos centros consolidados
            Conexao.Execute "DELETE from Usuarios_setor_previsao from Usuarios_setor_previsao INNER JOIN Usuarios_Setor ON Usuarios_setor_previsao.ID_CC = Usuarios_Setor.ID where Usuarios_Setor.Consolidacao = 'True'"
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select Usuarios_setor_previsao.* from Usuarios_setor_previsao INNER JOIN Usuarios_Setor ON Usuarios_setor_previsao.ID_CC = Usuarios_Setor.ID where Usuarios_Setor.Consolidacao is null or Usuarios_Setor.Consolidacao = 'False' order by Usuarios_setor_previsao.ID", Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False Then
                TBFI.MoveLast
                PBLista.Min = 0
                PBLista.Max = TBFI.RecordCount
                PBLista.Value = 1
                Contador = 0
                TBFI.MoveFirst
                Do While TBFI.EOF = False
                    Set TBAfericao = CreateObject("adodb.recordset")
                    TBAfericao.Open "Select * from Usuarios_Setor_Consolidacao where ID_CC_consolidado = " & TBFI!ID_CC, Conexao, adOpenKeyset, adLockOptimistic
                    If TBAfericao.EOF = False Then
                        Do While TBAfericao.EOF = False
                            Set TBGravar = CreateObject("adodb.recordset")
                            TBGravar.Open "Select * from Usuarios_setor_previsao where ID_CC = " & TBAfericao!ID_CC & " and ID_origem = " & TBFI!ID, Conexao, adOpenKeyset, adLockOptimistic
                            If TBGravar.EOF = True Then TBGravar.AddNew
                            ProcEnviaDadosPrevisaoaAtualizacao TBAfericao!ID_CC, TBFI!ID
                            TBGravar.Update
                            TBGravar.Close
                            
                            Set TBCiclo = CreateObject("adodb.recordset")
                            TBCiclo.Open "Select * from Usuarios_Setor_Consolidacao where ID_CC_consolidado = " & TBAfericao!ID_CC, Conexao, adOpenKeyset, adLockOptimistic
                            If TBCiclo.EOF = False Then
                                Do While TBCiclo.EOF = False
                                    Set TBGravar = CreateObject("adodb.recordset")
                                    TBGravar.Open "Select * from Usuarios_setor_previsao where ID_CC = " & TBCiclo!ID_CC & " and ID_origem = " & TBFI!ID, Conexao, adOpenKeyset, adLockOptimistic
                                    If TBGravar.EOF = True Then TBGravar.AddNew
                                    ProcEnviaDadosPrevisaoaAtualizacao TBCiclo!ID_CC, TBFI!ID
                                    TBGravar.Update
                                    TBGravar.Close
                                    TBCiclo.MoveNext
                                Loop
                            End If
                            TBCiclo.Close
                            
                            TBAfericao.MoveNext
                        Loop
                    End If
                    TBFI.MoveNext
                    Contador = Contador + 1
                    PBLista.Value = Contador
                Loop
            End If
            TBFI.Close
        End If
        USMsgBox ("Atualização efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
        
        ProcCarregaLista
        '==================================
        Modulo = "Custos/Centro de custo"
        Evento = "Atualizar"
        ID_documento = 0
        Documento = ""
        Documento1 = ""
        ProcGravaEvento
        '==================================
    End With
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvarCCRealizadoInvestAtualizacao(Data1 As Date, ID_empresa As Integer, Operacao As String, ID_CC As Long, Cod_produto As Long, ID_plano_contas As Long, ID_estoque As Long, ID_lista As Long, valor As Double, CC_produto As Boolean, Bloqueado As Boolean)
On Error GoTo tratar_erro

NovoValor = Replace(valor, ",", ".")
ProcINSERTINTO "CC_realizado", "Data, Responsavel, ID_empresa, Operacao, ID_CC, Cod_produto, ID_PC, ID_estoque, ID_lista, Valor, Bloqueado", "'" & Data & "', '" & pubUsuario & "', " & ID_empresa & ", '" & Operacao & "', " & ID_CC & ", " & Cod_produto & ", " & ID_plano_contas & ", " & IIf(ID_estoque = 0, "NULL", ID_estoque) & ", " & ID_lista & ", " & NovoValor & ", " & IIf(Bloqueado = True, 1, 0) & ""

Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select ID from CC_realizado where ID_estoque = " & ID_estoque, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = False Then
    If CC_produto = True And Operacao = "Crédito" Then Conexao.Execute "Update CC_realizado Set ID_ref_debito = " & TBGravar!ID - 1 & " where ID = " & TBGravar!ID
End If
TBGravar.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviaDadosPrevisaoaAtualizacao(ID_CC As Long, ID_origem As Long)
On Error GoTo tratar_erro

TBGravar!ID_CC = ID_CC
TBGravar!Data = TBFI!Data
TBGravar!Responsavel = TBFI!Responsavel
TBGravar!Mes = TBFI!Mes
TBGravar!Ano = TBFI!Ano
TBGravar!Revisao = TBFI!Revisao
TBGravar!ID_PC = TBFI!ID_PC
TBGravar!valor = TBFI!valor
TBGravar!ID_origem = ID_origem

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviaDadosAtualizacaoCCCons(ID_CC As Long, ID_origem As Long)
On Error GoTo tratar_erro

TBGravar!Data = TBFI!Data
TBGravar!Responsavel = TBFI!Responsavel
TBGravar!ID_empresa = TBFI!ID_empresa
TBGravar!Operacao = TBFI!Operacao
TBGravar!ID_ref_debito = TBFI!ID_ref_debito
TBGravar!ID_CC = ID_CC
TBGravar!Cod_produto = TBFI!Cod_produto
TBGravar!ID_PC = TBFI!ID_PC
TBGravar!ID_estoque = TBFI!ID_estoque
TBGravar!ID_lista = TBFI!ID_lista
TBGravar!ID_financeiro = TBFI!ID_financeiro
TBGravar!valor = TBFI!valor
TBGravar!Percentual = TBFI!Percentual
TBGravar!Depreciacao = TBFI!Depreciacao
TBGravar!Rateio = TBFI!Rateio
TBGravar!ID_CC_rateio = TBFI!ID_CC_rateio
TBGravar!Bloqueado = TBFI!Bloqueado
TBGravar!ID_origem = ID_origem

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_percentual_rateio_Click()
On Error GoTo tratar_erro

If Chk_percentual_rateio.Value = 1 Then
    With Txt_percentual_rateio
        .Text = ""
        .Locked = False
        .TabStop = True
        .SetFocus
    End With
    Chk_valor_rateio.Value = 0
    With Txt_valor_rateio
        .Text = ""
        .Locked = True
        .TabStop = False
    End With
Else
    With Txt_percentual_rateio
        .Locked = True
        .TabStop = False
    End With
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_valor_rateio_Click()
On Error GoTo tratar_erro

If Chk_valor_rateio.Value = 1 Then
    With Txt_valor_rateio
        .Text = ""
        .Locked = False
        .TabStop = True
        .SetFocus
    End With
    Chk_percentual_rateio.Value = 0
    With Txt_percentual_rateio
        .Text = ""
        .Locked = True
        .TabStop = False
    End With
Else
    With Txt_valor_rateio
        .Locked = True
        .TabStop = False
    End With
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_empresa_Click()
On Error GoTo tratar_erro

ProcLimpaCampos
ProcLimpaCamposResp
ProcLimpaCamposPrev
ProcLimpaCamposDepreciacao
ProcLimpaCamposRateio
Lista1.ListItems.Clear
Lista2.ListItems.Clear
Lista3.ListItems.Clear
Lista4.ListItems.Clear
Lista5.ListItems.Clear

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

Private Sub Cmb_revisao_filtrar_Click()
On Error GoTo tratar_erro

ProcFiltrarMes

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbAno_Click()
On Error GoTo tratar_erro

ProcFiltrarMes

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbAno1_Click()
On Error GoTo tratar_erro

ProcFiltrarMes1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbAno2_Click()
On Error GoTo tratar_erro

ProcFiltrarMes2

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_localizar_PC_Click()
On Error GoTo tratar_erro

ProcLimpaVariaveisCarregaPC
Sit_REG = 1
frmproj_produto_PC.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimpaVariaveisCarregaPC()
On Error GoTo tratar_erro

Plano_contas_produtos = False
Plano_contas_familias = False
Plano_centro_de_custo = True
Plano_instituicao = False
Plano_opcoesgerais = False
Plano_Faturamento = False
Financeiro_Contas_Pagar = False
Financeiro_Contas_Pagas = False
Financeiro_Contas_Receber = False
Financeiro_Contas_Recebidas = False
Plano_PCP = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_localizar_PC_depreciacao_Click()
On Error GoTo tratar_erro

ProcLimpaVariaveisCarregaPC
Sit_REG = 2
frmproj_produto_PC.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_localizar_PC_rateio_Click()
On Error GoTo tratar_erro

ProcLimpaVariaveisCarregaPC
Sit_REG = 3
frmproj_produto_PC.Show 1

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
            Case vbKeyF3: ProcGravar
            Case vbKeyF4: If Cmb_opcao_lista = "Excluir" Then ProcExcluir
            Case vbKeyF5: ProcImprimir
            Case vbKeyF7: If Cmb_opcao_lista = "Status" Then ProcStatus
            Case vbKeyF8: ProcConsolidar
            Case vbKeyF9: ProcVisualizar
            Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
    Case 1:
        Select Case KeyCode
            Case vbKeyInsert: ProcNovoResp
            Case vbKeyF3: ProcGravarResp
            Case vbKeyF4: ProcExcluirResp
            Case vbKeyF5: ProcImprimir
            Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
    Case 2:
        Select Case KeyCode
            Case vbKeyInsert: ProcNovoPrev
            Case vbKeyF3: ProcGravarPrev
            Case vbKeyF4: ProcExcluirPrev
            Case vbKeyF5: ProcImprimir
            Case vbKeyF7: ProcCopiarPrev
            Case vbKeyF8: ProcRevisarPrev
            Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
    Case 3:
        Select Case KeyCode
            Case vbKeyInsert: ProcNovoDepreciacao
            Case vbKeyF3: ProcGravarDepreciacao
            Case vbKeyF4: ProcExcluirDepreciacao
            Case vbKeyF5: ProcImprimir
            Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
    Case 4:
        Select Case KeyCode
            Case vbKeyInsert: ProcNovoRateio
            Case vbKeyF3: ProcGravarRateio
            Case vbKeyF4: ProcExcluirRateio
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
    
ProcCarregaToolBar1 Me, 15195, 13, True
ProcCarregaToolBar2 Me, 15195, 10, False
ProcCarregaToolBar3 Me, 15195, 12, False

ProcVerifAcessoVisLancamentos
Formulario = "Custos/Centro de custo"
Direitos
ProcLimpaVariaveisPrincipais
SSTab1.Tab = 0
ProcCarregaComboEmpresa Cmb_empresa, False
Cmb_opcao_lista = "Excluir"

Cmb_resp_CC.Clear
Set TBFI = CreateObject("adodb.recordset")
TBFI.Open "Select Usuario from Usuarios where Usuario is not null order by Usuario", Conexao, adOpenKeyset, adLockOptimistic
If TBFI.EOF = False Then
    Do While TBFI.EOF = False
        Cmb_resp_CC.AddItem TBFI!Usuario
        TBFI.MoveNext
    Loop
End If
TBFI.Close

ProcCarregaLista
txtData_depreciacao.Value = Date
txtData_rateio.Value = Date
Cmb_revisao_filtrar = "Todas"

ProcCarregaComboAno cmbAno, "2011", 2
ProcCarregaComboAno cmbAno1, "2011", 1
ProcCarregaComboAno cmbAno2, "2011", 1
TabFiltro.Tabs(Month(Date)).Selected = True
TabFiltro1.Tabs(Month(Date)).Selected = True
TabFiltro2.Tabs(Month(Date)).Selected = True

ProcRemoveObjetosResize Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro
    
ProcVerifAcessoVisLancamentos
Formulario = "Custos/Centro de custo"
Direitos
ProcLimpaVariaveisPrincipais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcVerifAcessoVisLancamentos()
On Error GoTo tratar_erro

With USToolBar1
    Set TBAcessos = CreateObject("adodb.recordset")
    TBAcessos.Open "Select * from Acessos where IDUsuario = " & pubIDUsuario & " and Acesso = 'Custos/Centro de custo/Visualizar lançamentos realizados'", Conexao, adOpenKeyset, adLockOptimistic
    If TBAcessos.EOF = False Then
        .ButtonState(9) = 0
    Else
        .ButtonState(9) = 5
    End If
    TBAcessos.Close
    .Refresh
End With

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
                If USMsgBox("Deseja realmente excluir este(s) centro(s) de custo?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select * from Usuarios_Setor where ID = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False Then
                '==================================
                Modulo = "Custos/Centro de custo"
                Evento = "Excluir"
                ID_documento = .ListItems(InitFor)
                If IsNull(TBFI!CODIGO) = False And TBFI!CODIGO <> "" Then Documento = "Código: " & TBFI!CODIGO & " - Descrição: " & TBFI!Setor Else Documento = "Descrição: " & TBFI!Setor
                Documento1 = ""
                ProcGravaEvento
                '==================================
                Conexao.Execute "DELETE from usuarios_setor where ID = " & .ListItems(InitFor)
                Conexao.Execute "DELETE from Usuarios_Setor_Responsavel where ID_CC = " & .ListItems(InitFor)
            End If
            TBFI.Close
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) centro(s) de custo antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Centro(s) de custo excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCampos
    ProcCarregaLista
    Frame2.Enabled = False
    Novo_Centro = False
    ProcLimparTudo
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluirResp()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso.")
    Exit Sub
End If
Permitido = False
With Lista1
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) responsável(eis) pelo centro de custo?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select * from usuarios_setor_responsavel where ID = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False Then
                '==================================
                Modulo = "Custos/Centro de custo"
                Evento = "Excluir responsável pelo centro de custo"
                ID_documento = .ListItems(InitFor)
                If IsNumeric(txt_Codigo) = True Then Documento = "Código: " & txt_Codigo & " - Descrição: " & Txt_descricao Else Documento = "Descrição: " & Txt_descricao
                Documento1 = "Responsável pelo centro de custo: " & TBFI!Responsavel_CC
                ProcGravaEvento
                '==================================
                Conexao.Execute "DELETE from usuarios_setor_responsavel where ID = " & .ListItems(InitFor)
            End If
            TBFI.Close
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) responsável(eis) pelo centro de custo antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Responsável(eis) pelo centro de custo excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCamposResp
    ProcCarregaListaResp
    Frame4.Enabled = False
    Novo_Centro1 = False
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluirPrev()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso.")
    Exit Sub
End If
Permitido = False
With Lista2
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir esta(s) previsão(ões) orçamentária(s)?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select * from usuarios_setor_previsao where ID = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False Then
                'Excluir previsão no centro consolidado
                Set TBAfericao = CreateObject("adodb.recordset")
                TBAfericao.Open "Select * from Usuarios_Setor_Consolidacao where ID_CC_consolidado = " & TBFI!ID_CC, Conexao, adOpenKeyset, adLockOptimistic
                If TBAfericao.EOF = False Then
                    Do While TBAfericao.EOF = False
                        Conexao.Execute "DELETE from Usuarios_setor_previsao where ID_CC = " & TBAfericao!ID_CC & " and ID_origem = " & TBFI!ID
                        
                        Set TBCiclo = CreateObject("adodb.recordset")
                        TBCiclo.Open "Select * from Usuarios_Setor_Consolidacao where ID_CC_consolidado = " & TBAfericao!ID_CC, Conexao, adOpenKeyset, adLockOptimistic
                        If TBCiclo.EOF = False Then
                            Do While TBCiclo.EOF = False
                                Conexao.Execute "DELETE from Usuarios_setor_previsao where ID_CC = " & TBCiclo!ID_CC & " and ID_origem = " & TBFI!ID
                                TBCiclo.MoveNext
                            Loop
                        End If
                        TBCiclo.Close
                        
                        TBAfericao.MoveNext
                    Loop
                End If
                TBAfericao.Close
                
                '==================================
                Modulo = "Custos/Centro de custo"
                Evento = "Excluir previsão orçamentária"
                ID_documento = .ListItems(InitFor)
                If IsNumeric(txt_Codigo) = True Then Documento = "Código: " & txt_Codigo & " - Descrição: " & Txt_descricao Else Documento = "Descrição: " & Txt_descricao
                Documento1 = "Ano: " & TBFI!Ano & " - Revisão : " & TBFI!Revisao & " - Código contábil: " & .ListItems(InitFor).ListSubItems(6) & " - Conta contábil: " & .ListItems(InitFor).ListSubItems(7) & " - Valor: " & .ListItems(InitFor).ListSubItems(8)
                ProcGravaEvento
                '==================================
                Conexao.Execute "DELETE from usuarios_setor_previsao where ID = " & .ListItems(InitFor)
            End If
            TBFI.Close
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe a(s) previsão(ões) orçamentária(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Previsão(ões) orçamentária(s) excluída(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    
    ProcCarregaLista
    If CodigoLista <> 0 And Lista.ListItems.Count <> 0 Then
        Lista.SelectedItem = Lista.ListItems(CodigoLista)
        Lista.SetFocus
    End If
    
    ProcLimpaCamposPrev
    ProcFiltrarMes
    Frame1.Enabled = False
    Novo_Centro2 = False
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluirDepreciacao()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso.")
    Exit Sub
End If
Permitido = False
With Lista3
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir esta(s) depreciação(ões)?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select * from CC_realizado where ID = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False Then
                Conexao.Execute "DELETE from CC_realizado where ID_origem = " & .ListItems(InitFor) & " and Depreciacao = 'True'"
                
                '==================================
                Modulo = "Custos/Centro de custo"
                Evento = "Excluir depreciação"
                ID_documento = .ListItems(InitFor)
                If IsNumeric(txt_Codigo) = True Then Documento = "Código: " & txt_Codigo & " - Descrição: " & Txt_descricao Else Documento = "Descrição: " & Txt_descricao
                Documento1 = "Código contábil: " & .ListItems(InitFor).ListSubItems(3) & " - Conta contábil: " & .ListItems(InitFor).ListSubItems(4) & " - Valor: " & .ListItems(InitFor).ListSubItems(5)
                ProcGravaEvento
                '==================================
                
                TBFI.Delete
            End If
            TBFI.Close
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe a(s) depreciação(ões) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Depreciação(ões) excluída(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    
    ProcCarregaLista
    If CodigoLista <> 0 And Lista.ListItems.Count <> 0 Then
        Lista.SelectedItem = Lista.ListItems(CodigoLista)
        Lista.SetFocus
    End If
    
    ProcLimpaCamposDepreciacao
    ProcFiltrarMes1
    Frame3.Enabled = False
    Novo_Centro3 = False
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluirRateio()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso.")
    Exit Sub
End If
Permitido = False
With Lista5
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) rateios(s)?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            
            'Débito
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select * from CC_realizado where ID = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False Then
                Conexao.Execute "DELETE from CC_realizado where ID_origem = " & TBFI!ID & " and Operacao = 'Débito' and Rateio = 'True'"
                
                TBFI.Delete
            End If
            
            'Crédito
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select * from CC_realizado where ID_ref_debito = " & .ListItems(InitFor) & " and Operacao = 'Crédito' and Rateio = 'True'", Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False Then
                Conexao.Execute "DELETE from CC_realizado where ID_origem = " & TBFI!ID & " and Operacao = 'Crédito' and Rateio = 'True'"
                
                TBFI.Delete
            End If
            TBFI.Close
            
            '==================================
            Modulo = "Custos/Centro de custo"
            Evento = "Excluir rateio"
            ID_documento = .ListItems(InitFor)
            If IsNumeric(txt_Codigo) = True Then Documento = "Código: " & txt_Codigo & " - Descrição: " & Txt_descricao Else Documento = "Descrição: " & Txt_descricao
            If IsNumeric(.ListItems(InitFor).ListSubItems(3)) = True Then Documento1 = "Código: " & .ListItems(InitFor).ListSubItems(3) & " - Descrição: " & .ListItems(InitFor).ListSubItems(4) & " - Código contábil: " & .ListItems(InitFor).ListSubItems(5) & " - Conta contábil: " & .ListItems(InitFor).ListSubItems(6) & " - Valor: " & .ListItems(InitFor).ListSubItems(7)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) rateio(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Rateio(s) excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    
    ProcCarregaLista
    If CodigoLista <> 0 And Lista.ListItems.Count <> 0 Then
        Lista.SelectedItem = Lista.ListItems(CodigoLista)
        Lista.SetFocus
    End If
    ProcCarregaListaCCRateio
    
    ProcLimpaCamposRateio
    ProcFiltrarMes2
    Frame5.Enabled = False
    Novo_Centro4 = False
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
Novo_Centro = True
Frame2.Enabled = True
Cmb_empresa.SetFocus
ProcLimparTudo
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovoResp()
On Error GoTo tratar_erro

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
ProcLimpaCamposResp
Novo_Centro1 = True
Frame4.Enabled = True
Cmb_resp_CC.SetFocus
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovoPrev()
On Error GoTo tratar_erro

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
ProcLimpaCamposPrev
Novo_Centro2 = True
Frame1.Enabled = True
Cmb_mes.SetFocus
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovoDepreciacao()
On Error GoTo tratar_erro

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
ProcLimpaCamposDepreciacao
Novo_Centro3 = True
Frame3.Enabled = True
txtData_depreciacao.SetFocus
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovoRateio()
On Error GoTo tratar_erro

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
ProcLimpaCamposRateio
Novo_Centro4 = True
With USToolBar2
    .ButtonState(2) = 0
    .Refresh
End With
Frame5.Enabled = True
txtData_rateio.SetFocus
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

If Novo_Centro = True Then
    If USMsgBox("O centro de custo ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcGravar
        If Novo_Centro = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
If Novo_Centro1 = True Then
    If USMsgBox("O responsável pelo centro de custo ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcGravarResp
        If Novo_Centro1 = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
If Novo_Centro2 = True Then
    If USMsgBox("A previsão orçamentária ainda não foi salva, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcGravarPrev
        If Novo_Centro2 = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
If Novo_Centro3 = True Then
    If USMsgBox("A depreciação ainda não foi salva, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcGravarDepreciacao
        If Novo_Centro3 = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
If Novo_Centro4 = True Then
    If USMsgBox("O rateio ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcGravarRateio
        If Novo_Centro4 = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
Novo_Centro = False
Novo_Centro1 = False
Novo_Centro2 = False
Novo_Centro3 = False
Novo_Centro4 = False
Unload Me

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
                    ProcVerificaRegistroUtilizadoSemMsg "usuarios", "Setor = '" & .ListItems.Item(InitFor).ListSubItems(5) & "'"
                    If Permitido = False Then
                        .ListItems.Item(InitFor).Checked = False
                        GoTo Proximo
                    End If
                    ProcVerificaRegistroUtilizadoSemMsg "Funcionarios", "setor = '" & .ListItems.Item(InitFor).ListSubItems(5) & "'"
                    If Permitido = False Then
                        .ListItems.Item(InitFor).Checked = False
                        GoTo Proximo
                    End If
                    ProcVerificaRegistroUtilizadoSemMsg "CadMaquinas", "setor = '" & .ListItems.Item(InitFor).ListSubItems(5) & "'"
                    If Permitido = False Then
                        .ListItems.Item(InitFor).Checked = False
                        GoTo Proximo
                    End If
                    ProcVerificaRegistroUtilizadoSemMsg "Compras_pedido_lista_custo", "ID_CC = " & .ListItems.Item(InitFor)
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

Private Sub Lista_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With Lista
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True And Cmb_opcao_lista = "Excluir" Then
            Mensagem = "Não é permitido excluir este centro de custo, pois o mesmo está sendo utilizado no módulo"
            ProcVerificaRegistroUtilizado "usuarios", "Setor = '" & .ListItems.Item(InitFor).ListSubItems(5) & "'", "Configurações do sistema/Usuários"
            If Permitido = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            ProcVerificaRegistroUtilizado "Funcionarios", "setor = '" & .ListItems.Item(InitFor).ListSubItems(5) & "'", "RH/Funcionários"
            If Permitido = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            ProcVerificaRegistroUtilizado "CadMaquinas", "setor = '" & .ListItems.Item(InitFor).ListSubItems(5) & "'", "PCP/Postos de trabalho"
            If Permitido = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            ProcVerificaRegistroUtilizado "Compras_pedido_lista_custo", "ID_CC = " & .ListItems.Item(InitFor), "Compras/Pedido"
            If Permitido = False Then
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

Private Sub Lista1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Lista1
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then .ListItems.Item(InitFor).Checked = False Else .ListItems.Item(InitFor).Checked = True
        Next InitFor
    End With
Else
    ProcOrdenaListView Lista1, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Lista2
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then .ListItems.Item(InitFor).Checked = False Else .ListItems.Item(InitFor).Checked = True
        Next InitFor
    End With
Else
    ProcOrdenaListView Lista2, ColumnHeader
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
TBLISTA.Open "Select * from Usuarios_setor where ID = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    ProcLimpaCampos
    ProcCarregaDados
    CodigoLista = Lista.SelectedItem.index
End If
TBLISTA.Close
Frame2.Enabled = True
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista1_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro
 
If Lista1.ListItems.Count = 0 Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Usuarios_setor_responsavel where ID = " & Lista1.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    ProcLimpaCamposResp
    ProcCarregaDadosResp
    CodigoLista1 = Lista1.SelectedItem.index
End If
TBLISTA.Close
Frame4.Enabled = True
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista2_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro
 
If Lista2.ListItems.Count = 0 Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select Usuarios_setor_previsao.*, tbl_familia.Codigo, tbl_familia.txt_descricao from Usuarios_setor_previsao INNER JOIN tbl_familia ON Usuarios_setor_previsao.ID_PC = tbl_familia.int_codfamilia where Usuarios_setor_previsao.ID = " & Lista2.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    ProcLimpaCamposPrev
    ProcCarregaDadosPrev
    CodigoLista2 = Lista2.SelectedItem.index
End If
TBLISTA.Close
Frame1.Enabled = True
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcCarregaDados()
On Error GoTo tratar_erro

If IsNull(TBLISTA!ID_empresa) = False And TBLISTA!ID_empresa <> "" Then ProcPuxaDadosComboEmpresa Cmb_empresa, TBLISTA!ID_empresa
txtId.Text = TBLISTA!ID
txtData = IIf(IsNull(TBLISTA!Data), "", Format(TBLISTA!Data, "dd/mm/yy"))
txtResponsavel = IIf(IsNull(TBLISTA!Responsavel), "", TBLISTA!Responsavel)
txt_Codigo = IIf(IsNull(TBLISTA!CODIGO), "___.___", TBLISTA!CODIGO)
Txt_descricao = IIf(IsNull(TBLISTA!Setor), "", TBLISTA!Setor)
txtStatus = IIf(IsNull(TBLISTA!DtBloq), "Liberado", "Bloqueado")

With SSTab1
    If TBLISTA!Consolidacao = True Then
        Chk_consolidacao.Value = 1
        .TabVisible(1) = False
        .TabVisible(2) = False
        .TabVisible(3) = False
        .TabVisible(4) = False
        .TabsPerRow = 1
    Else
        Chk_consolidacao.Value = 0
        .TabVisible(1) = True
        .TabVisible(2) = True
        .TabVisible(3) = True
        .TabVisible(4) = True
        .TabsPerRow = 5
        .Tabs = 5
    End If
End With
Chk_consolidacao.Enabled = True
ProcVerificaRegistroUtilizadoSemMsg "usuarios", "Setor = '" & Txt_descricao & "'"
If Permitido = False Then
    Chk_consolidacao.Enabled = False
Else
    ProcVerificaRegistroUtilizadoSemMsg "Funcionarios", "setor = '" & Txt_descricao & "'"
    If Permitido = False Then
        Chk_consolidacao.Enabled = False
    Else
        ProcVerificaRegistroUtilizadoSemMsg "CadMaquinas", "setor = '" & Txt_descricao & "'"
        If Permitido = False Then
            Chk_consolidacao.Enabled = False
        Else
            ProcVerificaRegistroUtilizadoSemMsg "Compras_pedido_lista_custo", "ID_CC = " & txtId
            If Permitido = False Then Chk_consolidacao.Enabled = False
        End If
    End If
End If

ProcMostrarEsconderBotaoCons

Caption = "Administrativo - Custos - Centro de custo (Código : " & TBLISTA!CODIGO & " - Descrição : " & TBLISTA!Setor & ")"
Txt_obs = IIf(IsNull(TBLISTA!Obs), "", TBLISTA!Obs)
ProcLimparTudo
Novo_Centro = False
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaDadosResp()
On Error GoTo tratar_erro
 
TxtID_resp = TBLISTA!ID
txtData = IIf(IsNull(TBLISTA!Data), "", Format(TBLISTA!Data, "dd/mm/yy"))
txtResponsavel = IIf(IsNull(TBLISTA!Responsavel), "", TBLISTA!Responsavel)
If IsNull(TBLISTA!Responsavel_CC) = False And TBLISTA!Responsavel_CC <> "" Then Cmb_resp_CC = TBLISTA!Responsavel_CC
Novo_Centro1 = False
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaDadosPrev()
On Error GoTo tratar_erro
 
Txt_ID_prev = TBLISTA!ID
txtData = IIf(IsNull(TBLISTA!Data), "", Format(TBLISTA!Data, "dd/mm/yy"))
txtResponsavel = IIf(IsNull(TBLISTA!Responsavel), "", TBLISTA!Responsavel)
If IsNull(TBLISTA!Mes) = False And TBLISTA!Mes <> "" Then Cmb_mes = FunVerificaNumeroMes(TBLISTA!Mes)
Txt_ano = IIf(IsNull(TBLISTA!Ano), "____", TBLISTA!Ano)
Txt_revisao = IIf(IsNull(TBLISTA!Revisao), "", TBLISTA!Revisao)
Txt_ID_PC = IIf(IsNull(TBLISTA!ID_PC), "", TBLISTA!ID_PC)
Txt_codigo_PC = IIf(IsNull(TBLISTA!CODIGO), "", TBLISTA!CODIGO)
Txt_descricao_PC = IIf(IsNull(TBLISTA!Txt_descricao), "", TBLISTA!Txt_descricao)
Txt_valor = IIf(IsNull(TBLISTA!valor), "", Format(TBLISTA!valor, "###,##0.00"))
Novo_Centro2 = False
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaDadosDepreciacao()
On Error GoTo tratar_erro
 
Txt_ID_depreciacao = TBLISTA!ID
txtData_depreciacao = IIf(IsNull(TBLISTA!Data), "", Format(TBLISTA!Data, "dd/mm/yy"))
txtResponsavel_depreciacao = IIf(IsNull(TBLISTA!Responsavel), "", TBLISTA!Responsavel)
Txt_ID_PC_depreciacao = IIf(IsNull(TBLISTA!ID_PC), "", TBLISTA!ID_PC)
Txt_codigo_PC_depreciacao = IIf(IsNull(TBLISTA!CODIGO), "", TBLISTA!CODIGO)
Txt_descricao_PC_depreciacao = IIf(IsNull(TBLISTA!Txt_descricao), "", TBLISTA!Txt_descricao)
Txt_valor_depreciacao = IIf(IsNull(TBLISTA!valor), "", Format(TBLISTA!valor, "###,##0.00"))
Novo_Centro3 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaDadosRateio()
On Error GoTo tratar_erro
 
Txt_ID_rateio = TBLISTA!ID
txtData_rateio = IIf(IsNull(TBLISTA!Data), "", Format(TBLISTA!Data, "dd/mm/yy"))
txtResponsavel_rateio = IIf(IsNull(TBLISTA!Responsavel), "", TBLISTA!Responsavel)
Txt_ID_PC_rateio = IIf(IsNull(TBLISTA!ID_PC), "", TBLISTA!ID_PC)
Txt_codigo_PC_rateio = IIf(IsNull(TBLISTA!CODIGO), "", TBLISTA!CODIGO)
Txt_descricao_PC_rateio = IIf(IsNull(TBLISTA!Txt_descricao), "", TBLISTA!Txt_descricao)
Txt_valor_rateio = IIf(IsNull(TBLISTA!valor), "", Format(TBLISTA!valor, "###,##0.00"))
Txt_percentual_rateio = IIf(IsNull(TBLISTA!Percentual), "", Format(TBLISTA!Percentual, "###,##0.0000000000"))
Novo_Centro4 = False

With USToolBar2
    .ButtonState(2) = 5
    .Refresh
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaLista()
On Error GoTo tratar_erro

Set TBAcessos = CreateObject("adodb.recordset")
TBAcessos.Open "Select * from Acessos where IDUsuario = " & pubIDUsuario & " and Acesso = 'Custos/Centro de custo/Visualizar todos'", Conexao, adOpenKeyset, adLockOptimistic
If TBAcessos.EOF = False Then
    TextoFiltro = "Select * from Usuarios_setor where ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " order by Codigo"
Else
    TextoFiltro = "Select Usuarios_setor.* from Usuarios_setor INNER JOIN Usuarios_Setor_Responsavel ON Usuarios_setor.ID = Usuarios_Setor_Responsavel.ID_CC where Usuarios_Setor_Responsavel.Responsavel_CC = '" & pubUsuario & "' and Usuarios_setor.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " order by Usuarios_setor.Codigo"
End If
TBAcessos.Close

Lista.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    TBLISTA.MoveLast
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    TBLISTA.MoveFirst
    Do While TBLISTA.EOF = False
        With Lista.ListItems.Add(, , TBLISTA!ID)
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select * from Empresa where Codigo = " & IIf(IsNull(TBLISTA!ID_empresa), 0, TBLISTA!ID_empresa), Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                .SubItems(1) = IIf(IsNull(TBAbrir!Empresa), "", TBAbrir!Empresa)
            End If
            
            .SubItems(2) = IIf(IsNull(TBLISTA!Data), "", Format(TBLISTA!Data, "dd/mm/yy"))
            .SubItems(3) = IIf(IsNull(TBLISTA!Responsavel), "", TBLISTA!Responsavel)
            .SubItems(4) = IIf(IsNull(TBLISTA!CODIGO), "", TBLISTA!CODIGO)
            .SubItems(5) = IIf(IsNull(TBLISTA!Setor), "", TBLISTA!Setor)
            
            'Valor orçado do ano atual
            valor = 0
            If TBLISTA!Consolidacao = True Then
                Set TBFI = CreateObject("adodb.recordset")
                TBFI.Open "Select ID_CC_consolidado from Usuarios_Setor_Consolidacao where ID_CC = " & TBLISTA!ID & " order by ID_CC_consolidado", Conexao, adOpenKeyset, adLockReadOnly
                If TBFI.EOF = False Then
                    Do While TBFI.EOF = False
                        Set TBAbrir = CreateObject("adodb.recordset")
                        TBAbrir.Open "Select Sum(Valor) as Valor1 from Usuarios_Setor_Previsao where ID_CC = " & TBFI!ID_CC_consolidado & " and Ano = " & Year(Date) & " group by Revisao", Conexao, adOpenKeyset, adLockReadOnly
                        If TBAbrir.EOF = False Then
                            TBAbrir.MoveLast
                            valor = valor + IIf(IsNull(TBAbrir!Valor1), 0, TBAbrir!Valor1)
                        End If
                        TBAbrir.Close
                        TBFI.MoveNext
                    Loop
                End If
                TBFI.Close
            Else
                Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open "Select Valor from Centro_de_custo_previsto_anual where ID_CC = " & TBLISTA!ID & " and Ano = " & Year(Date) & " order by Revisao desc", Conexao, adOpenKeyset, adLockReadOnly
                If TBAbrir.EOF = False Then
                    valor = IIf(IsNull(TBAbrir!valor), 0, TBAbrir!valor)
                End If
                TBAbrir.Close
            End If
            
            .SubItems(6) = Format(valor, "###,##0.00")
            
            'Valor real do ano atual
            'Débito
            Valor1 = 0
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select Sum(Valor_debito) as Valor1 from Centro_de_custo_real_anual where ID_CC = " & TBLISTA!ID & " and Ano = " & Year(Date), Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                Valor1 = IIf(IsNull(TBAbrir!Valor1), 0, Format(TBAbrir!Valor1, "###,##0.00"))
            End If
            
            'Crédito
            Valor2 = 0
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select Sum(Valor_credito) as Valor2 from Centro_de_custo_real_anual where ID_CC = " & TBLISTA!ID & " and Ano = " & Year(Date), Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                Valor2 = IIf(IsNull(TBAbrir!Valor2), 0, Format(TBAbrir!Valor2, "###,##0.00"))
            End If
            TBAbrir.Close
            
            'Real
            Valor3 = Valor1 - Valor2
            .SubItems(7) = Format(Valor3, "###,##0.00")
            
            'Variação
            Valor_Cofins_Prod = valor - Valor3
            .SubItems(8) = Format(Valor_Cofins_Prod, "###,##0.00")
            
            If valor = 0 Then
                Valor_total = -100
            ElseIf valor <> 0 And Valor_Cofins_Prod <> 0 Then
                    Valor_total = (Valor_Cofins_Prod / valor) * 100
                Else
                    Valor_total = 0
            End If
            .SubItems(9) = Format(Valor_total, "###,##0.00") & "%"
            .SubItems(10) = IIf(IsNull(TBLISTA!DtBloq), "Liberado", "Bloqueado")
            
            If (valor - Valor3) < 0 Then
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
            End If
            
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

Sub ProcCarregaListaResp()
On Error GoTo tratar_erro

Lista1.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Usuarios_setor_responsavel where ID_cc = " & txtId & " order by Responsavel_CC", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    TBLISTA.MoveLast
    PBLista1.Min = 0
    PBLista1.Max = TBLISTA.RecordCount
    PBLista1.Value = 1
    Contador = 0
    TBLISTA.MoveFirst
    Do While TBLISTA.EOF = False
        With Lista1.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Data), "", Format(TBLISTA!Data, "dd/mm/yy"))
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Responsavel), "", TBLISTA!Responsavel)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Responsavel_CC), "", TBLISTA!Responsavel_CC)
        End With
        TBLISTA.MoveNext
        Contador = Contador + 1
        PBLista1.Value = Contador
    Loop
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaListaPrev(StrSQL_carrega_previsao As String)
On Error GoTo tratar_erro

Lista2.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open StrSQL_carrega_previsao, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    TBLISTA.MoveLast
    PBLista1.Min = 0
    PBLista1.Max = TBLISTA.RecordCount
    PBLista1.Value = 1
    Contador = 0
    TBLISTA.MoveFirst
    Do While TBLISTA.EOF = False
        With Lista2.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Data), "", Format(TBLISTA!Data, "dd/mm/yy"))
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Responsavel), "", TBLISTA!Responsavel)
            If IsNull(TBLISTA!Mes) = False And TBLISTA!Mes <> "" Then .Item(.Count).SubItems(3) = FunVerificaNumeroMes(TBLISTA!Mes)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!Ano), "", TBLISTA!Ano)
            .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!Revisao), "", TBLISTA!Revisao)
            .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA!CODIGO), "", TBLISTA!CODIGO)
            .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA!Txt_descricao), "", TBLISTA!Txt_descricao)
            .Item(.Count).SubItems(8) = IIf(IsNull(TBLISTA!valor), "", Format(TBLISTA!valor, "###,##0.00"))
        End With
        TBLISTA.MoveNext
        Contador = Contador + 1
        PBLista1.Value = Contador
    Loop
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaListaDepreciacao(StrSQL_carrega_depreciacao As String)
On Error GoTo tratar_erro

Lista3.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open StrSQL_carrega_depreciacao, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    TBLISTA.MoveLast
    PBLista1.Min = 0
    PBLista1.Max = TBLISTA.RecordCount
    PBLista1.Value = 1
    Contador = 0
    TBLISTA.MoveFirst
    Do While TBLISTA.EOF = False
        With Lista3.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Data), "", Format(TBLISTA!Data, "dd/mm/yy"))
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Responsavel), "", TBLISTA!Responsavel)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!CODIGO), "", TBLISTA!CODIGO)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!Txt_descricao), "", TBLISTA!Txt_descricao)
            .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!valor), "", Format(TBLISTA!valor, "###,##0.00"))
        End With
        TBLISTA.MoveNext
        Contador = Contador + 1
        PBLista1.Value = Contador
    Loop
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaListaCCRateio()
On Error GoTo tratar_erro

Set TBAcessos = CreateObject("adodb.recordset")
TBAcessos.Open "Select * from Acessos where IDUsuario = " & pubIDUsuario & " and Acesso = 'Custos/Centro de custo/Visualizar todos'", Conexao, adOpenKeyset, adLockOptimistic
If TBAcessos.EOF = False Then
    TextoFiltro = "Select * from Usuarios_setor where ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and (Consolidacao = 'False' or Consolidacao is null) and ID <> " & txtId & " order by Codigo"
Else
    TextoFiltro = "Select Usuarios_setor.* from Usuarios_setor INNER JOIN Usuarios_Setor_Responsavel ON Usuarios_setor.ID = Usuarios_Setor_Responsavel.ID_CC where Usuarios_Setor_Responsavel.Responsavel_CC = '" & pubUsuario & "' and Usuarios_setor.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and (Usuarios_setor.Consolidacao = 'False' or Usuarios_setor.Consolidacao is null) and Usuarios_setor.ID <> " & txtId & " order by Usuarios_setor.Codigo"
End If
TBAcessos.Close

Lista4.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    TBLISTA.MoveLast
    PBLista1.Min = 0
    PBLista1.Max = TBLISTA.RecordCount
    PBLista1.Value = 1
    Contador = 0
    TBLISTA.MoveFirst
    Do While TBLISTA.EOF = False
        With Lista4.ListItems.Add(, , TBLISTA!ID)
            .SubItems(1) = IIf(IsNull(TBLISTA!CODIGO), "", TBLISTA!CODIGO)
            .SubItems(2) = IIf(IsNull(TBLISTA!Setor), "", TBLISTA!Setor)
            
            'Valor orçado do ano atual
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select Sum(Valor) as Valor1 from Usuarios_Setor_Previsao where ID_CC = " & TBLISTA!ID & " and Ano = " & Year(Date), Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                valor = IIf(IsNull(TBAbrir!Valor1), 0, TBAbrir!Valor1)
                .SubItems(3) = Format(valor, "###,##0.00")
            End If
            TBAbrir.Close
            
            'Valor real do ano atual
            'Débito
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select Sum(Valor) as Valor1 from Centro_de_custo_real_resumido where ID_CC = " & TBLISTA!ID & " and Ano = " & Year(Date) & " and Operacao = 'Débito'", Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                Valor1 = IIf(IsNull(TBAbrir!Valor1), 0, Format(TBAbrir!Valor1, "###,##0.00"))
            End If
            
            'Crédito
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select Sum(Valor) as Valor2 from Centro_de_custo_real_resumido where ID_CC = " & TBLISTA!ID & " and Ano = " & Year(Date) & " and Operacao = 'Crédito'", Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                Valor2 = IIf(IsNull(TBAbrir!Valor2), 0, Format(TBAbrir!Valor2, "###,##0.00"))
            End If
            TBAbrir.Close
            
            Valor3 = Valor1 - Valor2
            .SubItems(4) = Format(Valor3, "###,##0.00")
            .SubItems(5) = Format(valor - Valor3, "###,##0.00")
            
            If valor = 0 And Valor3 > 0 Then
                Valor_total = 100
            ElseIf valor <> Valor3 And Valor3 > 0 Then
                    Valor_total = (valor / Valor3) * 100
                Else
                    Valor_total = 0
            End If
            .SubItems(6) = Format(Valor_total, "###,##0.00") & "%"
            
            If (valor - Valor3) < 0 Then
                .ForeColor = vbRed
                .ListSubItems(1).ForeColor = vbRed
                .ListSubItems(2).ForeColor = vbRed
                .ListSubItems(3).ForeColor = vbRed
                .ListSubItems(4).ForeColor = vbRed
                .ListSubItems(5).ForeColor = vbRed
                .ListSubItems(6).ForeColor = vbRed
            End If
            
        End With
        TBLISTA.MoveNext
        Contador = Contador + 1
        PBLista1.Value = Contador
    Loop
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaListaRateio(StrSQL_carrega_rateio As String)
On Error GoTo tratar_erro

Lista5.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open StrSQL_carrega_rateio, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    TBLISTA.MoveLast
    PBLista1.Min = 0
    PBLista1.Max = TBLISTA.RecordCount
    PBLista1.Value = 1
    Contador = 0
    TBLISTA.MoveFirst
    Do While TBLISTA.EOF = False
        With Lista5.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Data), "", Format(TBLISTA!Data, "dd/mm/yy"))
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Responsavel), "", TBLISTA!Responsavel)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Codigo_CC), "", TBLISTA!Codigo_CC)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!Setor), "", TBLISTA!Setor)
            .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!CODIGO), "", TBLISTA!CODIGO)
            .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA!Txt_descricao), "", TBLISTA!Txt_descricao)
            .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA!valor), "", Format(TBLISTA!valor, "###,##0.00"))
        End With
        TBLISTA.MoveNext
        Contador = Contador + 1
        PBLista1.Value = Contador
    Loop
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimparTudo()
On Error GoTo tratar_erro

Frame4.Enabled = False
Frame1.Enabled = False
ProcLimpaCamposResp
ProcLimpaCamposPrev
Novo_Centro1 = False
Novo_Centro2 = False

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
txtStatus = "Liberado"
txt_Codigo = "___.___"
Txt_descricao = ""
Chk_consolidacao.Value = 0
ProcMostrarEsconderBotaoCons
Txt_obs = ""
CodigoLista = 0
Caption = "Administrativo - Custos - Centro de custo"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCamposResp()
On Error GoTo tratar_erro

TxtID_resp = 0
txtData_resp = Format(Date, "dd/mm/yy")
txtResponsavel_resp = pubUsuario
Cmb_resp_CC.ListIndex = -1
CodigoLista1 = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCamposPrev()
On Error GoTo tratar_erro

Txt_ID_prev = 0
txtData_prev = Format(Date, "dd/mm/yy")
txtResponsavel_prev = pubUsuario
Cmb_mes.ListIndex = -1
Txt_ano = "____"
Txt_revisao = 0
Txt_ID_PC = 0
Txt_codigo_PC = ""
Txt_descricao_PC = ""
Txt_valor = ""
CodigoLista2 = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCamposDepreciacao()
On Error GoTo tratar_erro

Txt_ID_depreciacao = 0
txtData_depreciacao.Value = Date
txtResponsavel_depreciacao = pubUsuario
Txt_ID_PC_depreciacao = 0
Txt_codigo_PC_depreciacao = ""
Txt_descricao_PC_depreciacao = ""
Txt_valor_depreciacao = ""
CodigoLista3 = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCamposRateio()
On Error GoTo tratar_erro

Txt_ID_rateio = 0
txtData_rateio.Value = Date
txtResponsavel_rateio = pubUsuario
Txt_ID_PC_rateio = 0
Txt_codigo_PC_rateio = ""
Txt_descricao_PC_rateio = ""
Chk_valor_rateio.Value = 0
Txt_valor_rateio = ""
Chk_percentual_rateio.Value = 0
Txt_percentual_rateio = ""
CodigoLista4 = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcGravar()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Frame2.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
Acao = "salvar"
If IsNumeric(txt_Codigo) = False Then
    NomeCampo = "o código"
    ProcVerificaAcao
    txt_Codigo.SetFocus
    Exit Sub
End If
If Txt_descricao = "" Then
    NomeCampo = "a descricao"
    ProcVerificaAcao
    Txt_descricao.SetFocus
    Exit Sub
End If
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from Usuarios_setor where ID = " & txtId.Text, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then
    TBGravar.AddNew
Else
    If Txt_descricao <> TBGravar!Setor Then
        Conexao.Execute "Update usuarios set Setor = '" & Txt_descricao & "' where Setor = '" & TBGravar!Setor & "'"
        Conexao.Execute "Update Funcionarios set setor = '" & Txt_descricao & "' where setor = '" & TBGravar!Setor & "'"
        Conexao.Execute "Update CadMaquinas set setor = '" & Txt_descricao & "' where setor = '" & TBGravar!Setor & "'"
    End If
End If
TBGravar!ID_empresa = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
If txtData <> "" Then TBGravar!Data = txtData Else TBGravar!Data = Date
If txtResponsavel <> "" Then TBGravar!Responsavel = txtResponsavel Else TBGravar!Responsavel = pubUsuario
TBGravar!CODIGO = txt_Codigo
TBGravar!Setor = Txt_descricao
If Chk_consolidacao.Value = 1 Then
    TBGravar!Consolidacao = True
Else
    TBGravar!Consolidacao = False
    Conexao.Execute "DELETE from Usuarios_Setor_Consolidacao where ID_CC = " & txtId
End If
TBGravar!Obs = Txt_obs
TBGravar.Update
txtId = TBGravar!ID
TBGravar.Close
ProcCarregaLista
If Novo_Centro = True Then
    USMsgBox ("Novo centro de custo cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar"
    If CodigoLista <> 0 And Lista.ListItems.Count <> 0 Then
        Lista.SelectedItem = Lista.ListItems(CodigoLista)
        Lista.SetFocus
    End If
End If
'==================================
Modulo = "Custos/Centro de custo"
ID_documento = txtId
Documento = "Código: " & txt_Codigo & " - Descrição: " & Txt_descricao
Documento1 = ""
ProcGravaEvento
'==================================
Novo_Centro = False
ProcMostrarEsconderBotaoCons

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcMostrarEsconderBotaoCons()
On Error GoTo tratar_erro

With USToolBar1
    If Chk_consolidacao.Value = 1 Then
        .ButtonState(8) = 0
        .ButtonState(9) = 5
    Else
        .ButtonState(8) = 5
        ProcVerifAcessoVisLancamentos
    End If
    .Refresh
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcGravarResp()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Frame4.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
Acao = "salvar"
If Cmb_resp_CC = "" Then
    NomeCampo = "o responsável pelo centro de custo"
    ProcVerificaAcao
    Cmb_resp_CC.SetFocus
    Exit Sub
End If
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from Usuarios_setor_responsavel where ID = " & TxtID_resp.Text, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then TBGravar.AddNew
TBGravar!ID_CC = txtId
If txtData <> "" Then TBGravar!Data = txtData_resp Else TBGravar!Data = Date
If txtResponsavel <> "" Then TBGravar!Responsavel = txtResponsavel_resp Else TBGravar!Responsavel = pubUsuario
TBGravar!Responsavel_CC = Cmb_resp_CC
TBGravar.Update
TxtID_resp = TBGravar!ID
TBGravar.Close
ProcCarregaListaResp
If Novo_Centro1 = True Then
    USMsgBox ("Novo responsável pelo centro de custo cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo responsável pelo centro de custo"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar responsável pelo centro de custo"
    If CodigoLista1 <> 0 And Lista1.ListItems.Count <> 0 Then
        Lista1.SelectedItem = Lista1.ListItems(CodigoLista1)
        Lista1.SetFocus
    End If
End If
'==================================
Modulo = "Custos/Centro de custo"
ID_documento = TxtID_resp
Documento = "Código: " & txt_Codigo & " - Descrição: " & Txt_descricao
Documento1 = "Responsável pelo centro de custo: " & Cmb_resp_CC
ProcGravaEvento
'==================================
Novo_Centro1 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcGravarPrev()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Frame1.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
Acao = "salvar"
If Cmb_mes = "" Then
    NomeCampo = "o mês"
    ProcVerificaAcao
    Cmb_mes.SetFocus
    Exit Sub
End If
If IsNumeric(Txt_ano) = False Then
    NomeCampo = "o ano"
    ProcVerificaAcao
    Txt_ano.SetFocus
    Exit Sub
End If
If Txt_ID_PC = 0 Then
    NomeCampo = "a conta contábil"
    ProcVerificaAcao
    Cmd_localizar_PC.SetFocus
    Exit Sub
End If
If Txt_valor = "" Then
    NomeCampo = "o valor"
    ProcVerificaAcao
    Txt_valor.SetFocus
    Exit Sub
End If

qt = FunVerificaMes(Cmb_mes)
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from Usuarios_setor_previsao where ID = " & Txt_ID_prev.Text, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from Usuarios_setor_previsao where ID_CC = " & txtId & " and Mes = " & qt & " and Ano = " & Txt_ano & " and revisao = " & Txt_revisao & " and ID_PC = " & Txt_ID_PC, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        USMsgBox ("Já existe uma previsão orçamentária desta conta contábil no ano de " & Cmb_ano & ", revisão " & Txt_revisao & " para este centro de custo."), vbExclamation, "CAPRIND v5.0"
        TBAbrir.Close
        Exit Sub
    End If
    TBAbrir.Close
    
    TBGravar.AddNew
End If
ProcEnviaDadosPrevisao txtId, 0
TBGravar.Update
Txt_ID_prev = TBGravar!ID
TBGravar.Close

'Grava previsão no centro consolidado
Set TBAfericao = CreateObject("adodb.recordset")
TBAfericao.Open "Select * from Usuarios_Setor_Consolidacao where ID_CC_consolidado = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
If TBAfericao.EOF = False Then
    Do While TBAfericao.EOF = False
        Set TBGravar = CreateObject("adodb.recordset")
        TBGravar.Open "Select * from Usuarios_setor_previsao where ID_CC = " & TBAfericao!ID_CC & " and ID_origem = " & Txt_ID_prev, Conexao, adOpenKeyset, adLockOptimistic
        If TBGravar.EOF = True Then TBGravar.AddNew
        ProcEnviaDadosPrevisao TBAfericao!ID_CC, Txt_ID_prev
        TBGravar.Update
        TBGravar.Close
        
        Set TBCiclo = CreateObject("adodb.recordset")
        TBCiclo.Open "Select * from Usuarios_Setor_Consolidacao where ID_CC_consolidado = " & TBAfericao!ID_CC, Conexao, adOpenKeyset, adLockOptimistic
        If TBCiclo.EOF = False Then
            Do While TBCiclo.EOF = False
                Set TBGravar = CreateObject("adodb.recordset")
                TBGravar.Open "Select * from Usuarios_setor_previsao where ID_CC = " & TBCiclo!ID_CC & " and ID_origem = " & Txt_ID_prev, Conexao, adOpenKeyset, adLockOptimistic
                If TBGravar.EOF = True Then TBGravar.AddNew
                ProcEnviaDadosPrevisao TBCiclo!ID_CC, Txt_ID_prev
                TBGravar.Update
                TBGravar.Close
                TBCiclo.MoveNext
            Loop
        End If
        TBCiclo.Close
        
        TBAfericao.MoveNext
    Loop
End If
TBAfericao.Close

ProcCarregaLista
If CodigoLista <> 0 And Lista.ListItems.Count <> 0 Then
    Lista.SelectedItem = Lista.ListItems(CodigoLista)
    Lista.SetFocus
End If

If Novo_Centro2 = True Then
    ProcCarregaListaPrev "Select Usuarios_setor_previsao.*, tbl_familia.Codigo, tbl_familia.txt_descricao from Usuarios_setor_previsao INNER JOIN tbl_familia ON Usuarios_setor_previsao.ID_PC = tbl_familia.int_codfamilia where Usuarios_setor_previsao.ID = " & Txt_ID_prev
    
    USMsgBox ("Nova previsão orçamentária cadastrada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Nova previsão orçamentária"
Else
    ProcFiltrarMes
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar responsável pelo centro de custo"
    If CodigoLista2 <> 0 And Lista2.ListItems.Count <> 0 Then
        Lista2.SelectedItem = Lista2.ListItems(CodigoLista2)
        Lista2.SetFocus
    End If
End If
'==================================
Modulo = "Custos/Centro de custo"
ID_documento = Txt_ID_prev
Documento = "Código: " & txt_Codigo & " - Descrição: " & Txt_descricao
Documento1 = "Ano: " & Cmb_ano & " - Revisão : " & Txt_revisao & " - Código contábil: " & Txt_codigo_PC & " - Conta contábil: " & Txt_descricao_PC & " - Valor: " & Format(Txt_valor, "###,##0.00")
ProcGravaEvento
'==================================
Novo_Centro2 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviaDadosPrevisao(ID_CC As Long, ID_origem As Long)
On Error GoTo tratar_erro

TBGravar!ID_CC = ID_CC
If txtData <> "" Then TBGravar!Data = txtData_prev Else TBGravar!Data = Date
If txtResponsavel <> "" Then TBGravar!Responsavel = txtResponsavel_prev Else TBGravar!Responsavel = pubUsuario
TBGravar!Mes = qt
TBGravar!Ano = Txt_ano
TBGravar!Revisao = Txt_revisao
TBGravar!ID_PC = Txt_ID_PC
TBGravar!valor = Txt_valor
TBGravar!ID_origem = ID_origem

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviaDadosRevisarPrev(ID_CC As Long, ID_origem As Long)
On Error GoTo tratar_erro

TBGravar!ID_CC = ID_CC
TBGravar!Data = Date
TBGravar!Responsavel = pubUsuario
TBGravar!Mes = TBFI!Mes
TBGravar!Ano = TBFI!Ano
TBGravar!Revisao = TBFI!Revisao + 1
TBGravar!ID_PC = TBFI!ID_PC
TBGravar!valor = TBFI!valor
TBGravar!ID_origem = ID_origem

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcGravarDepreciacao()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Frame3.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
Acao = "salvar"
If Txt_ID_PC_depreciacao = 0 Then
    NomeCampo = "a conta contábil"
    ProcVerificaAcao
    Cmd_localizar_PC_depreciacao.SetFocus
    Exit Sub
End If
Valor1 = IIf(Txt_valor_depreciacao = "", 0, Txt_valor_depreciacao)
If Valor1 = 0 Then
    NomeCampo = "o valor"
    ProcVerificaAcao
    Txt_valor_depreciacao.SetFocus
    Exit Sub
End If

Set TBFI = CreateObject("adodb.recordset")
TBFI.Open "Select * from CC_realizado where ID = " & Txt_ID_depreciacao, Conexao, adOpenKeyset, adLockOptimistic
If TBFI.EOF = True Then TBFI.AddNew
ProcEnviaDadosDepreciacao txtId, 0
TBFI.Update
Txt_ID_depreciacao = TBFI!ID
TBFI.Close

'Grava depreciação no centro consolidado
Set TBAfericao = CreateObject("adodb.recordset")
TBAfericao.Open "Select * from Usuarios_Setor_Consolidacao where ID_CC_consolidado = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
If TBAfericao.EOF = False Then
    Do While TBAfericao.EOF = False
        Set TBFI = CreateObject("adodb.recordset")
        TBFI.Open "Select * from CC_realizado where ID_CC = " & TBAfericao!ID_CC & " and ID_origem = " & Txt_ID_depreciacao, Conexao, adOpenKeyset, adLockOptimistic
        If TBFI.EOF = True Then TBFI.AddNew
        ProcEnviaDadosDepreciacao TBAfericao!ID_CC, Txt_ID_depreciacao
        TBFI.Update
        TBFI.Close
        
        Set TBCiclo = CreateObject("adodb.recordset")
        TBCiclo.Open "Select * from Usuarios_Setor_Consolidacao where ID_CC_consolidado = " & TBAfericao!ID_CC, Conexao, adOpenKeyset, adLockOptimistic
        If TBCiclo.EOF = False Then
            Do While TBCiclo.EOF = False
                Set TBFI = CreateObject("adodb.recordset")
                TBFI.Open "Select * from CC_realizado where ID_CC = " & TBCiclo!ID_CC & " and ID_origem = " & Txt_ID_depreciacao, Conexao, adOpenKeyset, adLockOptimistic
                If TBFI.EOF = True Then TBFI.AddNew
                ProcEnviaDadosDepreciacao TBCiclo!ID_CC, Txt_ID_depreciacao
                TBFI.Update
                TBFI.Close
                
                TBCiclo.MoveNext
            Loop
        End If
        TBCiclo.Close
        
        TBAfericao.MoveNext
    Loop
End If
TBAfericao.Close

ProcCarregaLista
If CodigoLista <> 0 And Lista.ListItems.Count <> 0 Then
    Lista.SelectedItem = Lista.ListItems(CodigoLista)
    Lista.SetFocus
End If

If Novo_Centro3 = True Then
    ProcCarregaListaDepreciacao "Select CC_realizado.*, tbl_familia.Codigo, tbl_familia.txt_descricao from CC_realizado INNER JOIN tbl_familia ON CC_realizado.ID_PC = tbl_familia.int_codfamilia where CC_realizado.ID = " & Txt_ID_depreciacao & " and CC_realizado.Depreciacao = 'True'"
    
    USMsgBox ("Nova depreciação cadastrada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Nova depreciação"
Else
    ProcFiltrarMes1
    
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar depreciação"
    If CodigoLista3 <> 0 And Lista3.ListItems.Count <> 0 Then
        Lista3.SelectedItem = Lista3.ListItems(CodigoLista3)
        Lista3.SetFocus
    End If
End If
'==================================
Modulo = "Custos/Centro de custo"
ID_documento = Txt_ID_depreciacao
Documento = "Código: " & txt_Codigo & " - Descrição: " & Txt_descricao
Documento1 = "Código contábil: " & Txt_codigo_PC_depreciacao & " - Conta contábil: " & Txt_descricao_PC_depreciacao & " - Valor: " & Format(Txt_valor_depreciacao, "###,##0.00")
ProcGravaEvento
'==================================
Novo_Centro3 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviaDadosDepreciacao(ID_CC As Long, ID_origem As Long)
On Error GoTo tratar_erro

TBFI!ID_CC = ID_CC
If txtData <> "" Then TBFI!Data = txtData_depreciacao.Value Else TBFI!Data = Date
If txtResponsavel <> "" Then TBFI!Responsavel = txtResponsavel_depreciacao Else TBFI!Responsavel = pubUsuario
TBFI!ID_empresa = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
TBFI!Operacao = "Débito"
TBFI!ID_PC = Txt_ID_PC_depreciacao
TBFI!valor = Txt_valor_depreciacao
TBFI!Depreciacao = True
TBFI!ID_origem = ID_origem

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcGravarRateio()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Frame5.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
Acao = "salvar"
If Txt_ID_PC_rateio = 0 Then
    NomeCampo = "a conta contábil"
    ProcVerificaAcao
    Cmd_localizar_PC_rateio.SetFocus
    Exit Sub
End If
Valor1 = IIf(Txt_valor_rateio = "", 0, Txt_valor_rateio)
If Valor1 = 0 Then
    NomeCampo = "o valor"
    ProcVerificaAcao
    Chk_valor_rateio.Value = 1
    Txt_valor_rateio.SetFocus
    Exit Sub
End If
Valor2 = Lista.SelectedItem.ListSubItems(7)
If Valor1 > Valor2 Then
    USMsgBox ("O valor não pode ser maior do que o valor do centro de custo."), vbExclamation, "CAPRIND v5.0"
    Txt_valor_rateio.SetFocus
    Exit Sub
End If

'Verifica o número de CC selecionado e calcula o valor por CC
Contador = 0
With Lista4
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then Contador = Contador + 1
    Next InitFor
End With
If Contador > 0 Then
    valor = Valor1 / Contador
    
    Valor1 = Txt_percentual_rateio
    Valor2 = Valor1 / Contador
End If

Permitido = False
With Lista4
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            Permitido = True
            
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select * from CC_realizado where ID = " & Txt_ID_rateio, Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = True Then TBFI.AddNew
            ProcEnviaDadosRateio "Débito", 0, .ListItems(InitFor), 0
            
            TBFI.Update
            Txt_ID_rateio = TBFI!ID
            IDAntigo = TBFI!ID
            
            'Grava rateio no centro consolidado
            Set TBAfericao = CreateObject("adodb.recordset")
            TBAfericao.Open "Select * from Usuarios_Setor_Consolidacao where ID_CC_consolidado = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
            If TBAfericao.EOF = False Then
                Do While TBAfericao.EOF = False
                    Set TBFI = CreateObject("adodb.recordset")
                    TBFI.Open "Select * from CC_realizado where ID_CC = " & TBAfericao!ID_CC & " and ID_origem = " & IDAntigo & " and Operacao = 'Débito'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBFI.EOF = True Then TBFI.AddNew
                    ProcEnviaDadosRateio "Débito", 0, TBAfericao!ID_CC, IDAntigo
                    TBFI.Update
                    TBFI.Close
                    
                    Set TBCiclo = CreateObject("adodb.recordset")
                    TBCiclo.Open "Select * from Usuarios_Setor_Consolidacao where ID_CC_consolidado = " & TBAfericao!ID_CC, Conexao, adOpenKeyset, adLockOptimistic
                    If TBCiclo.EOF = False Then
                        Do While TBCiclo.EOF = False
                            Set TBFI = CreateObject("adodb.recordset")
                            TBFI.Open "Select * from CC_realizado where ID_CC = " & TBCiclo!ID_CC & " and ID_origem = " & IDAntigo & " and Operacao = 'Débito'", Conexao, adOpenKeyset, adLockOptimistic
                            If TBFI.EOF = True Then TBFI.AddNew
                            ProcEnviaDadosRateio "Débito", 0, TBCiclo!ID_CC, IDAntigo
                            TBFI.Update
                            TBFI.Close
                            
                            TBCiclo.MoveNext
                        Loop
                    End If
                    TBCiclo.Close
                    
                    TBAfericao.MoveNext
                Loop
            End If
            
            'Cria crédito no CC principal vinculado ao débito do CC selecionado
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select * from CC_realizado where ID_ref_debito = " & IDAntigo, Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = True Then TBFI.AddNew
            ProcEnviaDadosRateio "Crédito", IDAntigo, Lista.SelectedItem, 0
            
            TBFI.Update
            TBFI.Close
            
            'Grava rateio no centro consolidado
            Set TBAfericao = CreateObject("adodb.recordset")
            TBAfericao.Open "Select * from Usuarios_Setor_Consolidacao where ID_CC_consolidado = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
            If TBAfericao.EOF = False Then
                Do While TBAfericao.EOF = False
                    Set TBFI = CreateObject("adodb.recordset")
                    TBFI.Open "Select * from CC_realizado where ID_CC = " & TBAfericao!ID_CC & " and ID_origem = " & IDAntigo & " and Operacao = 'Crédito' and ID_ref_debito = " & IDAntigo, Conexao, adOpenKeyset, adLockOptimistic
                    If TBFI.EOF = True Then TBFI.AddNew
                    ProcEnviaDadosRateio "Crédito", IDAntigo, TBAfericao!ID_CC, IDAntigo
                    TBFI.Update
                    TBFI.Close
                    
                    Set TBCiclo = CreateObject("adodb.recordset")
                    TBCiclo.Open "Select * from Usuarios_Setor_Consolidacao where ID_CC_consolidado = " & TBAfericao!ID_CC, Conexao, adOpenKeyset, adLockOptimistic
                    If TBCiclo.EOF = False Then
                        Do While TBCiclo.EOF = False
                            Set TBFI = CreateObject("adodb.recordset")
                            TBFI.Open "Select * from CC_realizado where ID_CC = " & TBCiclo!ID_CC & " and ID_origem = " & IDAntigo & " and Operacao = 'Crédito' and ID_ref_debito = " & IDAntigo, Conexao, adOpenKeyset, adLockOptimistic
                            If TBFI.EOF = True Then TBFI.AddNew
                            ProcEnviaDadosRateio "Crédito", IDAntigo, TBCiclo!ID_CC, IDAntigo
                            TBFI.Update
                            TBFI.Close
                            
                            TBCiclo.MoveNext
                        Loop
                    End If
                    TBCiclo.Close
                    
                    TBAfericao.MoveNext
                Loop
            End If
            TBAfericao.Close
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) centro(s) de custo antes de salvar."), vbExclamation, "CAPRIND v5.0"
Else
    ProcCarregaLista
    If CodigoLista <> 0 And Lista.ListItems.Count <> 0 Then
        Lista.SelectedItem = Lista.ListItems(CodigoLista)
        Lista.SetFocus
    End If
    ProcCarregaListaCCRateio
    
    If Novo_Centro4 = True Then
        ProcCarregaListaRateio "Select CC_realizado.*, Usuarios_Setor.Codigo as Codigo_CC, Usuarios_Setor.Setor, tbl_familia.Codigo, tbl_familia.Txt_descricao from (CC_realizado INNER JOIN tbl_familia ON CC_realizado.ID_PC = tbl_familia.int_codfamilia) INNER JOIN Usuarios_Setor ON CC_realizado.ID_CC = Usuarios_Setor.Id where CC_realizado.ID = " & Txt_ID_rateio & " and CC_realizado.Rateio = 'True' and CC_realizado.Operacao = 'Débito'"
        
        USMsgBox ("Novo rateio cadastrada com sucesso."), vbInformation, "CAPRIND v5.0"
        Evento = "Novo rateio"
    Else
        ProcFiltrarMes2
        
        USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
        Evento = "Alterar rateio"
        If CodigoLista4 <> 0 And Lista5.ListItems.Count <> 0 Then
            Lista5.SelectedItem = Lista5.ListItems(CodigoLista4)
            Lista5.SetFocus
        End If
    End If
    '==================================
    Modulo = "Custos/Centro de custo"
    ID_documento = Txt_ID_depreciacao
    Documento = "Código: " & txt_Codigo & " - Descrição: " & Txt_descricao
    Documento1 = "Código contábil: " & Txt_codigo_PC_depreciacao & " - Conta contábil: " & Txt_descricao_PC_depreciacao & " - Valor: " & Format(Txt_valor_depreciacao, "###,##0.00")
    ProcGravaEvento
    '==================================
    Novo_Centro3 = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviaDadosRateio(Operacao As String, ID_ref_debito As Long, ID_CC As Long, ID_origem As Long)
On Error GoTo tratar_erro

TBFI!Data = txtData_rateio
TBFI!Responsavel = txtResponsavel_rateio
TBFI!ID_empresa = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
TBFI!Operacao = Operacao
TBFI!ID_ref_debito = IDAntigo
TBFI!ID_CC = ID_CC
TBFI!ID_PC = Txt_ID_PC_rateio
TBFI!valor = valor
TBFI!Percentual = Valor2
TBFI!Rateio = True
TBFI!ID_CC_rateio = txtId
TBFI!ID_origem = ID_origem
TBFI!ID_ref_debito = ID_ref_debito

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista3_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Lista3
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then .ListItems.Item(InitFor).Checked = False Else .ListItems.Item(InitFor).Checked = True
        Next InitFor
    End With
Else
    ProcOrdenaListView Lista3, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista4_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView Lista4, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista5_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Lista5
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then .ListItems.Item(InitFor).Checked = False Else .ListItems.Item(InitFor).Checked = True
        Next InitFor
    End With
Else
    ProcOrdenaListView Lista5, ColumnHeader
End If
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista3_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro
 
If Lista3.ListItems.Count = 0 Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select CC_realizado.*, tbl_familia.Codigo, tbl_familia.txt_descricao from CC_realizado INNER JOIN tbl_familia ON CC_realizado.ID_PC = tbl_familia.int_codfamilia where CC_realizado.ID = " & Lista3.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    ProcLimpaCamposDepreciacao
    ProcCarregaDadosDepreciacao
    CodigoLista3 = Lista3.SelectedItem.index
End If
TBLISTA.Close
Frame3.Enabled = True
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista5_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro
 
If Lista5.ListItems.Count = 0 Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select CC_realizado.*, Usuarios_Setor.Codigo as Codigo_CC, Usuarios_Setor.Setor, tbl_familia.Codigo, tbl_familia.txt_descricao from (CC_realizado INNER JOIN tbl_familia ON CC_realizado.ID_PC = tbl_familia.int_codfamilia) INNER JOIN Usuarios_Setor ON CC_realizado.ID_CC = Usuarios_Setor.Id where CC_realizado.ID = " & Lista5.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    ProcLimpaCamposRateio
    ProcCarregaDadosRateio
    CodigoLista4 = Lista5.SelectedItem.index
End If
TBLISTA.Close
Frame5.Enabled = True
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub OptAteomes_Click()
On Error GoTo tratar_erro

ProcFiltrarMes

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub OptAteomes1_Click()
On Error GoTo tratar_erro

ProcFiltrarMes1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub OptAteomes2_Click()
On Error GoTo tratar_erro

ProcFiltrarMes2

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub OptDomes_Click()
On Error GoTo tratar_erro

ProcFiltrarMes

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub OptDomes1_Click()
On Error GoTo tratar_erro

ProcFiltrarMes1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub OptDomes2_Click()
On Error GoTo tratar_erro

ProcFiltrarMes2

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
Cmb_empresa.Visible = True
USToolBar1.Visible = False
USToolBar2.Visible = False
USToolBar3.Visible = False
With USToolBar2
    .ButtonState(2) = 0
    .Refresh
End With
If SSTab1.Tab = 0 Then
    PBLista.Visible = True
    PBLista1.Visible = False
Else
    PBLista.Visible = False
    PBLista1.Visible = True
End If

Select Case SSTab1.Tab
    Case 0:
        USToolBar1.Visible = True
        If Lista.Visible = True Then Lista.SetFocus
    Case 1:
        USToolBar2.Visible = True
        Cmb_empresa.Visible = False
        Lista1.SetFocus
        ProcCarregaListaResp
    Case 2:
        USToolBar3.Visible = True
        Cmb_empresa.Visible = False
        Lista2.SetFocus
        ProcFiltrarMes
    Case 3:
        USToolBar2.Visible = True
        Cmb_empresa.Visible = False
        Lista3.SetFocus
        ProcFiltrarMes1
    Case 4:
        USToolBar2.Visible = True
        Cmb_empresa.Visible = False
        Lista4.SetFocus
        
        With USToolBar2
            If Novo_Centro4 = True Then .ButtonState(2) = 0 Else .ButtonState(2) = 5
            .Refresh
        End With
        
        ProcCarregaListaCCRateio
        ProcFiltrarMes2
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcCopiarPrev()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With Lista2
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente copiar esta(s) previsão(ões) orçamentária(s)?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
        End If
    Next InitFor
End With
If Permitido = False Then USMsgBox ("Informe a(s) previsão(ões) orçamentária(s) antes de copiar."), vbExclamation, "CAPRIND v5.0" Else Frm_centro_de_custo_copiar_prev.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcRevisarPrev()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With Lista2
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente revisar esta(s) previsão(ões) orçamentária(s)?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select * from usuarios_setor_previsao where ID = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False Then
                '==================================
                Modulo = "Custos/Centro de custo"
                Evento = "Revisar previsão orçamentária"
                ID_documento = .ListItems(InitFor)
                If IsNumeric(txt_Codigo) = True Then Documento = "Código: " & txt_Codigo & " - Descrição: " & Txt_descricao Else Documento = "Descrição: " & Txt_descricao
                Documento1 = "Ano: " & TBFI!Ano & " - Revisão : " & TBFI!Revisao & " - Código contábil: " & .ListItems(InitFor).ListSubItems(6) & " - Conta contábil: " & .ListItems(InitFor).ListSubItems(7) & " - Valor: " & .ListItems(InitFor).ListSubItems(8)
                ProcGravaEvento
                '==================================
                
                Set TBGravar = CreateObject("adodb.recordset")
                TBGravar.Open "Select * from usuarios_setor_previsao", Conexao, adOpenKeyset, adLockOptimistic
                TBGravar.AddNew
                ProcEnviaDadosRevisarPrev TBFI!ID_CC, 0
                TBGravar.Update
                IDlista = TBGravar!ID
                TBGravar.Close
                
                'Grava previsão no centro consolidado
                Set TBAfericao = CreateObject("adodb.recordset")
                TBAfericao.Open "Select * from Usuarios_Setor_Consolidacao where ID_CC_consolidado = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
                If TBAfericao.EOF = False Then
                    Do While TBAfericao.EOF = False
                        Set TBGravar = CreateObject("adodb.recordset")
                        TBGravar.Open "Select * from Usuarios_setor_previsao where ID_CC = " & TBAfericao!ID_CC & " and ID_origem = " & IDlista, Conexao, adOpenKeyset, adLockOptimistic
                        If TBGravar.EOF = True Then TBGravar.AddNew
                        ProcEnviaDadosRevisarPrev TBAfericao!ID_CC, IDlista
                        TBGravar.Update
                        TBGravar.Close
                        
                        Set TBCiclo = CreateObject("adodb.recordset")
                        TBCiclo.Open "Select * from Usuarios_Setor_Consolidacao where ID_CC_consolidado = " & TBAfericao!ID_CC, Conexao, adOpenKeyset, adLockOptimistic
                        If TBCiclo.EOF = False Then
                            Do While TBCiclo.EOF = False
                                Set TBGravar = CreateObject("adodb.recordset")
                                TBGravar.Open "Select * from Usuarios_setor_previsao where ID_CC = " & TBCiclo!ID_CC & " and ID_origem = " & IDlista, Conexao, adOpenKeyset, adLockOptimistic
                                If TBGravar.EOF = True Then TBGravar.AddNew
                                ProcEnviaDadosRevisarPrev TBCiclo!ID_CC, IDlista
                                TBGravar.Update
                                TBGravar.Close
                                TBCiclo.MoveNext
                            Loop
                        End If
                        TBCiclo.Close
                        
                        TBAfericao.MoveNext
                    Loop
                End If
                TBAfericao.Close
                
            End If
            TBFI.Close
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe a(s) previsão(ões) orçamentária(s) antes de revisar."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Previsão(ões) orçamentária(s) revisada(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    
    ProcCarregaLista
    If CodigoLista <> 0 And Lista.ListItems.Count <> 0 Then
        Lista.SelectedItem = Lista.ListItems(CodigoLista)
        Lista.SetFocus
    End If
    
    ProcFiltrarMes
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub TabFiltro_Click()
On Error GoTo tratar_erro

ProcFiltrarMes

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrarMes()
On Error GoTo tratar_erro

M = FunVerificaMes(TabFiltro.SelectedItem.key)
If Cmb_revisao_filtrar = "Todas" Then TextoFiltro = "" Else TextoFiltro = "and Usuarios_setor_previsao.Revisao = " & Cmb_revisao_filtrar
If OptDomes.Value = True Then TextoFiltro1 = "Usuarios_setor_previsao.Mes = '" & M & "'" Else TextoFiltro1 = "Usuarios_setor_previsao.Mes <= '" & M & "'"
Familiatext = "Select Usuarios_setor_previsao.*, tbl_familia.Codigo, tbl_familia.txt_descricao from Usuarios_setor_previsao INNER JOIN tbl_familia ON Usuarios_setor_previsao.ID_PC = tbl_familia.int_codfamilia where Usuarios_setor_previsao.ID_CC = " & txtId & " and " & TextoFiltro1 & " and Usuarios_setor_previsao.Ano = '" & cmbAno & "' " & TextoFiltro & " order by Usuarios_setor_previsao.mes, tbl_familia.Codigo desc"
ProcCarregaListaPrev Familiatext

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub TabFiltro1_Click()
On Error GoTo tratar_erro

ProcFiltrarMes1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrarMes1()
On Error GoTo tratar_erro

M = FunVerificaMes(TabFiltro1.SelectedItem.key)
If OptDomes1.Value = True Then
    Familiatext = "Select CC_realizado.*, tbl_familia.Codigo, tbl_familia.txt_descricao from CC_realizado INNER JOIN tbl_familia ON CC_realizado.ID_PC = tbl_familia.int_codfamilia where CC_realizado.ID_CC = " & txtId & " and CC_realizado.Depreciacao = 'True' and Month(CC_realizado.Data) = '" & M & "' and Year(CC_realizado.Data) = '" & cmbAno & "' order by CC_realizado.Data desc, tbl_familia.Codigo"
Else
    Familiatext = "Select CC_realizado.*, tbl_familia.Codigo, tbl_familia.txt_descricao from CC_realizado INNER JOIN tbl_familia ON CC_realizado.ID_PC = tbl_familia.int_codfamilia where CC_realizado.ID_CC = " & txtId & " and CC_realizado.Depreciacao = 'True' and Month(CC_realizado.Data) <= '" & M & "' and Year(CC_realizado.Data) = '" & cmbAno & "' order by CC_realizado.Data desc, tbl_familia.Codigo"
End If
ProcCarregaListaDepreciacao Familiatext

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub TabFiltro2_Click()
On Error GoTo tratar_erro

ProcFiltrarMes2

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrarMes2()
On Error GoTo tratar_erro

M = FunVerificaMes(TabFiltro2.SelectedItem.key)
If OptDomes2.Value = True Then
    Familiatext = "Select CC_realizado.*, Usuarios_Setor.Codigo as Codigo_CC, Usuarios_Setor.Setor, tbl_familia.Codigo, tbl_familia.Txt_descricao from (CC_realizado INNER JOIN tbl_familia ON CC_realizado.ID_PC = tbl_familia.int_codfamilia) INNER JOIN Usuarios_Setor ON CC_realizado.ID_CC = Usuarios_Setor.Id where CC_realizado.ID_CC_rateio = " & txtId & " and CC_realizado.Rateio = 'True' and CC_realizado.Operacao = 'Débito' and Month(CC_realizado.Data) = '" & M & "' and Year(CC_realizado.Data) = '" & cmbAno & "' and Usuarios_Setor.Consolidacao = 'False' order by CC_realizado.Data desc"
Else
    Familiatext = "Select CC_realizado.*, Usuarios_Setor.Codigo as Codigo_CC, Usuarios_Setor.Setor, tbl_familia.Codigo, tbl_familia.Txt_descricao from (CC_realizado INNER JOIN tbl_familia ON CC_realizado.ID_PC = tbl_familia.int_codfamilia) INNER JOIN Usuarios_Setor ON CC_realizado.ID_CC = Usuarios_Setor.Id where CC_realizado.ID_CC_rateio = " & txtId & " and CC_realizado.Rateio = 'True' and CC_realizado.Operacao = 'Débito' and Month(CC_realizado.Data) <= '" & M & "' and Year(CC_realizado.Data) = '" & cmbAno & "' and Usuarios_Setor.Consolidacao = 'False' order by CC_realizado.Data desc"
End If
ProcCarregaListaRateio Familiatext

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_percentual_rateio_Change()
On Error GoTo tratar_erro

If Chk_percentual_rateio.Value = 1 Then
    Txt_valor_rateio = ""
    If Txt_percentual_rateio <> "" Then
        VerifNumero = Txt_percentual_rateio
        ProcVerificaNumero
        If VerifNumero = False Then
            Txt_percentual_rateio = ""
            Txt_percentual_rateio.SetFocus
            Exit Sub
        End If
        Valor1 = Txt_percentual_rateio
        Valor2 = Lista.SelectedItem.ListSubItems(7)
        Valor3 = (Valor2 * Valor1) / 100
        Txt_valor_rateio = Format(Valor3, "###,##0.00")
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_percentual_rateio_LostFocus()
On Error GoTo tratar_erro

Txt_percentual_rateio = Format(Txt_percentual_rateio, "###,##0.0000000000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_Valor_Change()
On Error GoTo tratar_erro

If Txt_valor <> "" Then
    VerifNumero = Txt_valor
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_valor = ""
        Txt_valor.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_valor_depreciacao_Change()
On Error GoTo tratar_erro

If Txt_valor_depreciacao <> "" Then
    VerifNumero = Txt_valor_depreciacao
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_valor_depreciacao = ""
        Txt_valor_depreciacao.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_valor_depreciacao_LostFocus()
On Error GoTo tratar_erro

Txt_valor_depreciacao = Format(Txt_valor_depreciacao, "###,##0.00")
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_Valor_LostFocus()
On Error GoTo tratar_erro

Txt_valor = Format(Txt_valor, "###,##0.00")
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_valor_rateio_Change()
On Error GoTo tratar_erro

If Chk_valor_rateio.Value = 1 Then
    Txt_percentual_rateio = ""
    If Txt_valor_rateio <> "" Then
        VerifNumero = Txt_valor_rateio
        ProcVerificaNumero
        If VerifNumero = False Then
            Txt_valor_rateio = ""
            Txt_valor_rateio.SetFocus
            Exit Sub
        End If
        Valor1 = Txt_valor_rateio
        Valor2 = Lista.SelectedItem.ListSubItems(7)
        If Valor2 <> 0 Then Valor3 = (Valor1 * 100) / Valor2 Else Valor3 = 0
        Txt_percentual_rateio = Format(Valor3, "###,##0.0000000000")
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_valor_rateio_LostFocus()
On Error GoTo tratar_erro

Txt_valor_rateio = Format(Txt_valor_rateio, "###,##0.00")
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcNovo
    Case 2: ProcGravar
    Case 3: ProcExcluir
    Case 4: ProcImprimir
    Case 5: ProcAnterior
    Case 6: ProcProximo
    Case 7: ProcStatus
    Case 8: ProcConsolidar
    Case 9: ProcVisualizar
    Case 10: ProcAtualizar
    Case 12: ProcAjuda
    Case 13: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar2_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1:
        Select Case SSTab1.Tab
            Case 1: ProcNovoResp
            Case 3: ProcNovoDepreciacao
            Case 4: ProcNovoRateio
        End Select
    Case 2:
        Select Case SSTab1.Tab
            Case 1: ProcGravarResp
            Case 3: ProcGravarDepreciacao
            Case 4: ProcGravarRateio
        End Select
    Case 3:
        Select Case SSTab1.Tab
            Case 1: ProcExcluirResp
            Case 3: ProcExcluirDepreciacao
            Case 4: ProcExcluirRateio
        End Select
    Case 4: ProcImprimir
    Case 5: ProcAnterior
    Case 6: ProcProximo
    Case 8: ProcAjuda
    Case 9: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar3_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcNovoPrev
    Case 2: ProcGravarPrev
    Case 3: ProcExcluirPrev
    Case 4: ProcImprimir
    Case 5: ProcAnterior
    Case 6: ProcProximo
    Case 7: ProcCopiarPrev
    Case 8: ProcRevisarPrev
    Case 10: ProcAjuda
    Case 11: ProcSair
End Select

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
    USMsgBox ("Informe o(s) centro(s) de custo antes de alterar o status."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Compras_Pedido = False
Vendas_PI = False
Vendas_Proposta = False
Plano_centro_de_custo = True

frmCompras_pedido_cancelar.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
