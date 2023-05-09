VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVendas_analise 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Outros - Análise crítica"
   ClientHeight    =   10035
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15360
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
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
   Begin DrawSuite2022.USToolBar USToolBar2 
      Height          =   975
      Left            =   75
      TabIndex        =   285
      Top             =   330
      Visible         =   0   'False
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
      ButtonCaption7  =   "Validação"
      ButtonEnabled7  =   0   'False
      ButtonIconSize7 =   32
      ButtonToolTipText7=   "Validar/Cancelar validação."
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
      ButtonLeft7     =   309
      ButtonTop7      =   2
      ButtonWidth7    =   53
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
      ButtonLeft8     =   364
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
      ButtonLeft9     =   368
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
      ButtonLeft10    =   411
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
      ButtonLeft11    =   443
      ButtonTop11     =   2
      ButtonWidth11   =   24
      ButtonHeight11  =   24
      Begin DrawSuite2022.USImageList USImageList2 
         Left            =   13770
         Top             =   150
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmVendas_analise.frx":0000
         Count           =   1
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar3 
      Height          =   975
      Left            =   75
      TabIndex        =   274
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
      ButtonCaption1  =   "Salvar"
      ButtonEnabled1  =   0   'False
      ButtonIconSize1 =   32
      ButtonToolTipText1=   "Salvar."
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
      ButtonToolTipText2=   "Excluir."
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
      ButtonCaption6  =   "Validação"
      ButtonEnabled6  =   0   'False
      ButtonIconSize6 =   32
      ButtonToolTipText6=   "Validar/Cancelar validação."
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
      ButtonLeft6     =   271
      ButtonTop6      =   2
      ButtonWidth6    =   53
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
      ButtonLeft7     =   326
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft8     =   330
      ButtonTop8      =   2
      ButtonWidth8    =   41
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft9     =   373
      ButtonTop9      =   2
      ButtonWidth9    =   30
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
      ButtonLeft10    =   405
      ButtonTop10     =   2
      ButtonWidth10   =   24
      ButtonHeight10  =   24
      ButtonUseMaskColor10=   0   'False
      Begin DrawSuite2022.USImageList USImageList3 
         Left            =   12240
         Top             =   120
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmVendas_analise.frx":606F
         Count           =   1
      End
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   75
      TabIndex        =   205
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   10065
      Left            =   0
      TabIndex        =   178
      Top             =   0
      Width           =   15390
      _ExtentX        =   27146
      _ExtentY        =   17754
      _Version        =   393216
      Tabs            =   7
      Tab             =   2
      TabsPerRow      =   7
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
      TabCaption(0)   =   "Análise crítica"
      TabPicture(0)   =   "frmVendas_analise.frx":B798
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "USToolBar1"
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(2)=   "Lista"
      Tab(0).Control(3)=   "txtIDproduto"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtID"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame3"
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Engenharia"
      TabPicture(1)   =   "frmVendas_analise.frx":B7B4
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(1)=   "SSTab_engenharia"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Processo"
      TabPicture(2)   =   "frmVendas_analise.frx":B7D0
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "USToolBar7"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "SSTab_processo"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Frame5"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "PCP"
      TabPicture(3)   =   "frmVendas_analise.frx":B7EC
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame20"
      Tab(3).Control(1)=   "SSTab4"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "Qualidade"
      TabPicture(4)   =   "frmVendas_analise.frx":B808
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame10"
      Tab(4).Control(1)=   "SSTab_qualidade"
      Tab(4).ControlCount=   2
      TabCaption(5)   =   "Compras"
      TabPicture(5)   =   "frmVendas_analise.frx":B824
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame9"
      Tab(5).Control(1)=   "PBLista_compras"
      Tab(5).Control(2)=   "SSTab_compras"
      Tab(5).ControlCount=   3
      TabCaption(6)   =   "Documentos"
      TabPicture(6)   =   "frmVendas_analise.frx":B840
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "CommonDialog1"
      Tab(6).Control(1)=   "txtID_doc"
      Tab(6).Control(1).Enabled=   0   'False
      Tab(6).Control(2)=   "Frame14"
      Tab(6).Control(3)=   "Lista_doc"
      Tab(6).ControlCount=   4
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
         Height          =   825
         Left            =   -74925
         TabIndex        =   262
         Top             =   1320
         Width           =   15195
         Begin VB.TextBox txtRespValidacao_Compras 
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
            Left            =   6390
            Locked          =   -1  'True
            TabIndex        =   145
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pela validação."
            Top             =   390
            Width           =   5565
         End
         Begin VB.TextBox txtDtValidacao_Compras 
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
            Left            =   4620
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   144
            TabStop         =   0   'False
            ToolTipText     =   "Data do cadastro."
            Top             =   390
            Width           =   1755
         End
         Begin VB.TextBox txtPrazo_Compras 
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
            Left            =   3060
            Locked          =   -1  'True
            TabIndex        =   143
            TabStop         =   0   'False
            ToolTipText     =   "Prazo."
            Top             =   390
            Width           =   1545
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
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
            Index           =   103
            Left            =   8182
            TabIndex        =   362
            Top             =   180
            Width           =   1980
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Data/hora validação"
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
            Index           =   95
            Left            =   4770
            TabIndex        =   361
            Top             =   180
            Width           =   1455
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Prazo"
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
            Left            =   3630
            TabIndex        =   263
            Top             =   180
            Width           =   405
         End
      End
      Begin VB.Frame Frame4 
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
         Height          =   1755
         Left            =   -74925
         TabIndex        =   252
         Top             =   1320
         Width           =   15195
         Begin VB.TextBox txtResponsavel_engenharia 
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
            TabIndex        =   271
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pelo cadastro."
            Top             =   390
            Width           =   5385
         End
         Begin VB.TextBox txtData_engenharia 
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
            TabIndex        =   270
            TabStop         =   0   'False
            ToolTipText     =   "Data do cadastro."
            Top             =   390
            Width           =   1185
         End
         Begin VB.TextBox Txt_obs_engenharia 
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
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   264
            ToolTipText     =   "Observações."
            Top             =   1002
            Width           =   14835
         End
         Begin VB.TextBox txtDtValidacao_Engenharia 
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
            Left            =   7980
            Locked          =   -1  'True
            TabIndex        =   261
            TabStop         =   0   'False
            ToolTipText     =   "Data e hora da validação."
            Top             =   390
            Width           =   1665
         End
         Begin VB.TextBox txtPrazo_Engenharia 
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
            Left            =   6780
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   254
            TabStop         =   0   'False
            ToolTipText     =   "Prazo."
            Top             =   390
            Width           =   1185
         End
         Begin VB.TextBox txtRespValidacao_Engenharia 
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
            Left            =   9660
            Locked          =   -1  'True
            TabIndex        =   253
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pela validação."
            Top             =   390
            Width           =   5355
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
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
            Index           =   63
            Left            =   11347
            TabIndex        =   322
            Top             =   180
            Width           =   1980
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Data/hora validação"
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
            Index           =   62
            Left            =   8085
            TabIndex        =   321
            Top             =   180
            Width           =   1455
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Prazo"
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
            Index           =   61
            Left            =   7170
            TabIndex        =   320
            Top             =   180
            Width           =   405
         End
         Begin VB.Label Label1 
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
            Index           =   60
            Left            =   3615
            TabIndex        =   319
            Top             =   180
            Width           =   915
         End
         Begin VB.Label Label1 
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
            Height          =   225
            Index           =   17
            Left            =   7140
            TabIndex        =   265
            Top             =   810
            Width           =   945
         End
         Begin VB.Label Label1 
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
            Index           =   40
            Left            =   600
            TabIndex        =   255
            Top             =   180
            Width           =   345
         End
      End
      Begin VB.Frame Frame5 
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
         Height          =   825
         Left            =   75
         TabIndex        =   256
         Top             =   1320
         Width           =   15195
         Begin VB.TextBox txtRespValidacao_processo 
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
            Left            =   6390
            Locked          =   -1  'True
            TabIndex        =   259
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pela validação."
            Top             =   390
            Width           =   5565
         End
         Begin VB.TextBox txtDtValidacao_processo 
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
            Left            =   4620
            Locked          =   -1  'True
            TabIndex        =   258
            TabStop         =   0   'False
            ToolTipText     =   "Data e hora da validação."
            Top             =   390
            Width           =   1755
         End
         Begin VB.TextBox txtPrazo_Processo 
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
            Left            =   3060
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   257
            TabStop         =   0   'False
            ToolTipText     =   "Prazo."
            Top             =   390
            Width           =   1545
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
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
            Index           =   74
            Left            =   8182
            TabIndex        =   333
            Top             =   180
            Width           =   1980
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Data/hora validação"
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
            Index           =   73
            Left            =   4770
            TabIndex        =   332
            Top             =   180
            Width           =   1455
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Prazo"
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
            Left            =   3630
            TabIndex        =   260
            Top             =   180
            Width           =   405
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   -74925
         TabIndex        =   223
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
            TabIndex        =   35
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
            TabIndex        =   34
            Text            =   "30"
            ToolTipText     =   "Número de registros por página."
            Top             =   180
            Width           =   555
         End
         Begin DrawSuite2022.USButton cmdPagProx 
            Height          =   315
            Left            =   11760
            TabIndex        =   39
            ToolTipText     =   "Próxima página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmVendas_analise.frx":B85C
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
            TabIndex        =   38
            ToolTipText     =   "Página anterior."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmVendas_analise.frx":F000
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
            TabIndex        =   36
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
            TabIndex        =   37
            ToolTipText     =   "Primeira página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmVendas_analise.frx":12B09
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
            TabIndex        =   40
            ToolTipText     =   "Última página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmVendas_analise.frx":16BF8
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
            Left            =   4410
            TabIndex        =   302
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
            TabIndex        =   226
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
            TabIndex        =   225
            Top             =   240
            Width           =   1275
         End
         Begin VB.Label Label1 
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
            Index           =   43
            Left            =   3090
            TabIndex        =   224
            Top             =   240
            Width           =   645
         End
      End
      Begin VB.TextBox txtID 
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
         Left            =   -73200
         Locked          =   -1  'True
         MouseIcon       =   "frmVendas_analise.frx":1A484
         TabIndex        =   201
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   6900
         Visible         =   0   'False
         Width           =   495
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   -69600
         Top             =   3930
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.TextBox txtID_doc 
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
         Height          =   315
         Left            =   -70350
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   200
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   4110
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.TextBox txtIDproduto 
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
         Left            =   -72660
         Locked          =   -1  'True
         TabIndex        =   179
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   6900
         Visible         =   0   'False
         Width           =   615
      End
      Begin MSComctlLib.ListView Lista 
         Height          =   2925
         Left            =   -74925
         TabIndex        =   33
         Top             =   6150
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   5159
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
         NumItems        =   14
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "N"
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Tag             =   "T"
            Text            =   "Nº análise"
            Object.Width           =   2028
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Object.Tag             =   "N"
            Text            =   "Rev."
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Object.Tag             =   "D"
            Text            =   "Data"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "Responsável"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Object.Tag             =   "T"
            Text            =   "Cód. interno"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Object.Tag             =   "T"
            Text            =   "Descrição"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Object.Tag             =   "T"
            Text            =   "Cliente"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Object.Tag             =   "N"
            Text            =   "Valor total"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   9
            Object.Tag             =   "T"
            Text            =   "Engen."
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   10
            Object.Tag             =   "T"
            Text            =   "Processo"
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   11
            Object.Tag             =   "T"
            Text            =   "PCP"
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   12
            Object.Tag             =   "T"
            Text            =   "Qualid."
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   13
            Object.Tag             =   "T"
            Text            =   "Compras"
            Object.Width           =   1499
         EndProperty
      End
      Begin VB.Frame Frame1 
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
         Height          =   4815
         Left            =   -74925
         TabIndex        =   180
         Top             =   1320
         Width           =   15195
         Begin VB.TextBox txtID_entrega 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   180
            TabIndex        =   250
            Text            =   "0"
            Top             =   3420
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.TextBox txtID_cobranca 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   7710
            TabIndex        =   249
            Text            =   "0"
            Top             =   3420
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.TextBox txtdescricao 
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
            MaxLength       =   255
            TabIndex        =   248
            ToolTipText     =   "Descrição."
            Top             =   1614
            Width           =   5270
         End
         Begin VB.TextBox txtdesenho 
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
            TabIndex        =   247
            ToolTipText     =   "Código interno."
            Top             =   1002
            Width           =   2055
         End
         Begin VB.TextBox Txt_analise 
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
            Left            =   180
            Locked          =   -1  'True
            TabIndex        =   246
            TabStop         =   0   'False
            ToolTipText     =   "Número da análise."
            Top             =   390
            Width           =   2055
         End
         Begin VB.TextBox txtContato 
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
            Left            =   180
            MaxLength       =   100
            TabIndex        =   20
            ToolTipText     =   "Contato do cliente."
            Top             =   2226
            Width           =   3255
         End
         Begin VB.TextBox txtNRef 
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
            Left            =   165
            MaxLength       =   100
            TabIndex        =   26
            ToolTipText     =   "Número da referência."
            Top             =   2838
            Width           =   5940
         End
         Begin VB.ComboBox cmbLocal_entrega 
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
            ItemData        =   "frmVendas_analise.frx":1A78E
            Left            =   180
            List            =   "frmVendas_analise.frx":1A790
            Sorted          =   -1  'True
            TabIndex        =   28
            ToolTipText     =   "Local de entrega."
            Top             =   3420
            Width           =   7095
         End
         Begin VB.TextBox txtDepartamento 
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
            Left            =   3870
            MaxLength       =   100
            TabIndex        =   22
            ToolTipText     =   "Departamento do contato."
            Top             =   2226
            Width           =   3870
         End
         Begin VB.TextBox txtTelefone 
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
            Left            =   7755
            MaxLength       =   14
            TabIndex        =   23
            ToolTipText     =   "Número do telefone."
            Top             =   2226
            Width           =   1840
         End
         Begin VB.TextBox txtFax 
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
            Left            =   9615
            MaxLength       =   14
            TabIndex        =   24
            ToolTipText     =   "Número do fax."
            Top             =   2226
            Width           =   1840
         End
         Begin VB.TextBox txtEmail 
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
            Left            =   11475
            MaxLength       =   60
            TabIndex        =   25
            ToolTipText     =   "E-mail."
            Top             =   2226
            Width           =   3540
         End
         Begin VB.ComboBox cmbLocal_cobranca 
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
            ItemData        =   "frmVendas_analise.frx":1A792
            Left            =   7710
            List            =   "frmVendas_analise.frx":1A794
            Sorted          =   -1  'True
            TabIndex        =   30
            ToolTipText     =   "Local de cobrança."
            Top             =   3420
            Width           =   6975
         End
         Begin VB.CommandButton cmdLocalentrega 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   7290
            Picture         =   "frmVendas_analise.frx":1A796
            Style           =   1  'Graphical
            TabIndex        =   29
            ToolTipText     =   "Filtrar local de entrega do cliente."
            Top             =   3420
            Width           =   315
         End
         Begin VB.CommandButton cmdLocalcobranca 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   14700
            Picture         =   "frmVendas_analise.frx":1ABB1
            Style           =   1  'Graphical
            TabIndex        =   31
            ToolTipText     =   "Filtrar local de cobrança do cliente."
            Top             =   3420
            Width           =   315
         End
         Begin VB.TextBox txtReferencia 
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
            Left            =   6120
            MaxLength       =   180
            TabIndex        =   27
            ToolTipText     =   "Descrição da referência."
            Top             =   2838
            Width           =   8895
         End
         Begin VB.CommandButton cmdContato 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   3450
            Picture         =   "frmVendas_analise.frx":1AFCC
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Localizar contatos do cliente."
            Top             =   2226
            Width           =   315
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
            ItemData        =   "frmVendas_analise.frx":1B0CE
            Left            =   12194
            List            =   "frmVendas_analise.frx":1B0D0
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   12
            ToolTipText     =   "Unidade comercial."
            Top             =   1002
            Width           =   855
         End
         Begin VB.TextBox Txt_rev_analise 
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
            TabIndex        =   0
            TabStop         =   0   'False
            ToolTipText     =   "Revisão da análise."
            Top             =   390
            Width           =   495
         End
         Begin VB.TextBox Txt_tipo 
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
            Left            =   5460
            Locked          =   -1  'True
            MaxLength       =   255
            TabIndex        =   15
            TabStop         =   0   'False
            ToolTipText     =   "Tipo."
            Top             =   1614
            Width           =   2745
         End
         Begin VB.CommandButton Cmd_tipo 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   8220
            Picture         =   "frmVendas_analise.frx":1B0D2
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Localizar tipos."
            Top             =   1614
            Width           =   315
         End
         Begin VB.TextBox Txt_qtde_sol 
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
            Left            =   13056
            TabIndex        =   13
            ToolTipText     =   "Quantidade solicitada."
            Top             =   1002
            Width           =   975
         End
         Begin VB.CommandButton Cmd_filtrar_produto 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   2760
            Picture         =   "frmVendas_analise.frx":1B1D4
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Filtrar por código interno."
            Top             =   1002
            Width           =   315
         End
         Begin VB.TextBox txtOBS 
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
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   32
            ToolTipText     =   "Observação."
            Top             =   4050
            Width           =   14835
         End
         Begin VB.TextBox Txt_data_status 
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
            Left            =   13860
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   4
            TabStop         =   0   'False
            ToolTipText     =   "Data da aprovação/cancelamento/perda."
            Top             =   390
            Visible         =   0   'False
            Width           =   825
         End
         Begin VB.TextBox Txt_status 
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
            Height          =   315
            Left            =   10425
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   3
            TabStop         =   0   'False
            ToolTipText     =   "Status."
            Top             =   390
            Width           =   4260
         End
         Begin VB.CommandButton Cmd_status 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   14700
            Picture         =   "frmVendas_analise.frx":1B5EF
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Informar novo status."
            Top             =   390
            Width           =   315
         End
         Begin VB.TextBox txtIDCliente 
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
            Left            =   8640
            TabIndex        =   17
            ToolTipText     =   "Código do cliente."
            Top             =   1614
            Width           =   705
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
            Left            =   6210
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   10
            ToolTipText     =   "Família."
            Top             =   1002
            Width           =   5130
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
            ItemData        =   "frmVendas_analise.frx":1B6F1
            Left            =   11345
            List            =   "frmVendas_analise.frx":1B6F3
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   11
            ToolTipText     =   "Unidade de estoque."
            Top             =   1002
            Width           =   855
         End
         Begin VB.CommandButton cmdCliente 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   14700
            Picture         =   "frmVendas_analise.frx":1B6F5
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Localizar clientes."
            Top             =   1614
            Width           =   315
         End
         Begin VB.TextBox txtCliente 
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
            Left            =   9360
            Locked          =   -1  'True
            MaxLength       =   255
            TabIndex        =   18
            TabStop         =   0   'False
            ToolTipText     =   "Cliente."
            Top             =   1614
            Width           =   5325
         End
         Begin VB.ComboBox cmbReferencia 
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
            Left            =   3495
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   9
            ToolTipText     =   "Código de referência."
            Top             =   1002
            Width           =   2715
         End
         Begin VB.CommandButton cmdProduto 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   3090
            Picture         =   "frmVendas_analise.frx":1B7F7
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Localizar produtos/serviços."
            Top             =   1002
            Width           =   315
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
            Left            =   3630
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   2
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pelo cadastro."
            Top             =   390
            Width           =   6785
         End
         Begin VB.TextBox txtRev_desenho 
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
            Left            =   2250
            TabIndex        =   6
            ToolTipText     =   "Revisão."
            Top             =   1002
            Width           =   495
         End
         Begin VB.TextBox txtqtde 
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
            Left            =   14040
            TabIndex        =   14
            Text            =   "1,000"
            ToolTipText     =   "Quantidade do lote."
            Top             =   1002
            Width           =   975
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
            Left            =   2760
            Locked          =   -1  'True
            TabIndex        =   1
            TabStop         =   0   'False
            ToolTipText     =   "Data do cadastro."
            Top             =   390
            Width           =   855
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Local de cobrança"
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
            Index           =   59
            Left            =   10552
            TabIndex        =   318
            Top             =   3210
            Width           =   1290
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Descrição da referência"
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
            Left            =   9720
            TabIndex        =   317
            Top             =   2640
            Width           =   1695
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Email"
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
            Left            =   13065
            TabIndex        =   316
            Top             =   2010
            Width           =   360
         End
         Begin VB.Label Label1 
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
            Index           =   56
            Left            =   10400
            TabIndex        =   315
            Top             =   2010
            Width           =   270
         End
         Begin VB.Label Label1 
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
            Index           =   55
            Left            =   8360
            TabIndex        =   314
            Top             =   2010
            Width           =   630
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
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
            Index           =   54
            Left            =   5288
            TabIndex        =   313
            Top             =   2010
            Width           =   1035
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
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
            Index           =   53
            Left            =   11775
            TabIndex        =   312
            Top             =   1410
            Width           =   495
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
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
            Index           =   52
            Left            =   6682
            TabIndex        =   311
            Top             =   1410
            Width           =   300
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Qtde. solic."
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
            Left            =   13131
            TabIndex        =   310
            Top             =   810
            Width           =   825
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Un. com."
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
            Left            =   12299
            TabIndex        =   309
            Top             =   810
            Width           =   645
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Un. est."
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
            Left            =   11480
            TabIndex        =   308
            Top             =   810
            Width           =   585
         End
         Begin VB.Label Label1 
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
            Index           =   48
            Left            =   8535
            TabIndex        =   307
            Top             =   810
            Width           =   480
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Código de referência"
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
            Left            =   4102
            TabIndex        =   306
            Top             =   810
            Width           =   1500
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
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
            Index           =   42
            Left            =   2325
            TabIndex        =   305
            Top             =   810
            Width           =   345
         End
         Begin VB.Label Label1 
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
            Index           =   30
            Left            =   6565
            TabIndex        =   304
            Top             =   180
            Width           =   915
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
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
            Index           =   24
            Left            =   2310
            TabIndex        =   303
            Top             =   180
            Width           =   375
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Nº análise"
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
            Left            =   787
            TabIndex        =   251
            Top             =   180
            Width           =   840
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Número da referência"
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
            Left            =   2355
            TabIndex        =   238
            Top             =   2640
            Width           =   1560
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Local de entrega"
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
            Left            =   3127
            TabIndex        =   237
            Top             =   3210
            Width           =   1200
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Contato"
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
            Left            =   1515
            TabIndex        =   236
            Top             =   2010
            Width           =   585
         End
         Begin VB.Label Label1 
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
            Index           =   9
            Left            =   7125
            TabIndex        =   231
            Top             =   3840
            Width           =   945
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Qtde. lote"
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
            Left            =   14160
            TabIndex        =   230
            Top             =   810
            Width           =   735
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Cód. interno"
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
            Left            =   757
            TabIndex        =   229
            Top             =   810
            Width           =   900
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
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
            Index           =   5
            Left            =   2470
            TabIndex        =   228
            Top             =   1410
            Width           =   690
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
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
            Index           =   2
            Left            =   12278
            TabIndex        =   204
            Top             =   180
            Width           =   555
         End
         Begin VB.Label Label1 
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
            Index           =   0
            Left            =   3015
            TabIndex        =   187
            Top             =   180
            Width           =   345
         End
      End
      Begin TabDlg.SSTab SSTab_engenharia 
         Height          =   8760
         Left            =   -74940
         TabIndex        =   181
         Top             =   3090
         Width           =   15255
         _ExtentX        =   26908
         _ExtentY        =   15452
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
         TabCaption(0)   =   "Materiais / Terceiros / Outros"
         TabPicture(0)   =   "frmVendas_analise.frx":1B8F9
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "USToolBar4"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Lista_engenharia"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Frame7"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "txtID_engenharia"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "txtIDproduto_engenharia"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).ControlCount=   5
         TabCaption(1)   =   "Check-list"
         TabPicture(1)   =   "frmVendas_analise.frx":1B915
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Txt_ID_check(0)"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Frame_check(0)"
         Tab(1).Control(2)=   "USToolBar5"
         Tab(1).Control(3)=   "Lista_check(0)"
         Tab(1).ControlCount=   4
         TabCaption(2)   =   "Normas"
         TabPicture(2)   =   "frmVendas_analise.frx":1B931
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Txt_ID_norma"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).Control(1)=   "Frame15"
         Tab(2).Control(2)=   "Lista_normas"
         Tab(2).Control(3)=   "USToolBar6"
         Tab(2).ControlCount=   4
         Begin VB.TextBox Txt_ID_check 
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
            Height          =   315
            Index           =   0
            Left            =   -71565
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   269
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   4800
            Visible         =   0   'False
            Width           =   675
         End
         Begin VB.TextBox Txt_ID_norma 
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
            Height          =   315
            Left            =   -71520
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   203
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   4800
            Visible         =   0   'False
            Width           =   675
         End
         Begin VB.Frame Frame15 
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
            Height          =   1965
            Left            =   -74925
            TabIndex        =   202
            Top             =   1320
            Width           =   15105
            Begin VB.TextBox Txt_obs_norma 
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
               Height          =   825
               Left            =   150
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   71
               ToolTipText     =   "Observações."
               Top             =   1005
               Width           =   14745
            End
            Begin VB.TextBox Txt_norma 
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
               Left            =   13380
               MaxLength       =   50
               TabIndex        =   70
               ToolTipText     =   "Norma."
               Top             =   390
               Width           =   1515
            End
            Begin VB.TextBox txtResponsavel_norma 
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
               Left            =   1710
               Locked          =   -1  'True
               TabIndex        =   69
               TabStop         =   0   'False
               ToolTipText     =   "Responsável pelo cadastro."
               Top             =   390
               Width           =   11655
            End
            Begin VB.TextBox txtData_norma 
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
               TabIndex        =   68
               TabStop         =   0   'False
               ToolTipText     =   "Data do cadastro."
               Top             =   390
               Width           =   1515
            End
            Begin VB.Label Label1 
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
               Index           =   72
               Left            =   7080
               TabIndex        =   331
               Top             =   180
               Width           =   915
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Norma"
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
               Left            =   13965
               TabIndex        =   273
               Top             =   180
               Width           =   465
            End
            Begin VB.Label Label1 
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
               Index           =   19
               Left            =   7110
               TabIndex        =   245
               Top             =   795
               Width           =   945
            End
            Begin VB.Label Label1 
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
               Index           =   18
               Left            =   765
               TabIndex        =   244
               Top             =   180
               Width           =   345
            End
         End
         Begin VB.TextBox txtIDproduto_engenharia 
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
            Left            =   2430
            Locked          =   -1  'True
            TabIndex        =   194
            TabStop         =   0   'False
            ToolTipText     =   "Data do cadastro."
            Top             =   4260
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Frame Frame_check 
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
            Height          =   1965
            Index           =   0
            Left            =   -74925
            TabIndex        =   184
            Top             =   1320
            Width           =   15105
            Begin VB.TextBox Txt_ID_descricao_check 
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
               Height          =   315
               Index           =   0
               Left            =   4860
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   287
               TabStop         =   0   'False
               Text            =   "0"
               Top             =   390
               Visible         =   0   'False
               Width           =   675
            End
            Begin VB.CommandButton Cmd_localizar_desc_check 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Index           =   0
               Left            =   14580
               Picture         =   "frmVendas_analise.frx":1B94D
               Style           =   1  'Graphical
               TabIndex        =   63
               ToolTipText     =   "Localizar descrição padrão do check-list."
               Top             =   390
               Width           =   315
            End
            Begin VB.TextBox Txt_descricao_chek 
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
               Index           =   0
               Left            =   4860
               Locked          =   -1  'True
               TabIndex        =   62
               TabStop         =   0   'False
               ToolTipText     =   "Descrição."
               Top             =   390
               Width           =   9705
            End
            Begin VB.Frame frameNumeros_engenharia 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
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
               Height          =   405
               Left            =   5640
               TabIndex        =   186
               Top             =   4650
               Width           =   2985
            End
            Begin VB.TextBox Txt_data_check 
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
               Index           =   0
               Left            =   150
               Locked          =   -1  'True
               MaxLength       =   25
               TabIndex        =   60
               TabStop         =   0   'False
               ToolTipText     =   "Data do cadastro."
               Top             =   390
               Width           =   855
            End
            Begin VB.TextBox Txt_responsavel_check 
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
               Index           =   0
               Left            =   1020
               Locked          =   -1  'True
               TabIndex        =   61
               TabStop         =   0   'False
               ToolTipText     =   "Responsável pelo cadastro."
               Top             =   390
               Width           =   3825
            End
            Begin VB.TextBox Txt_texto_check 
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
               Height          =   825
               Index           =   0
               Left            =   150
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   66
               TabStop         =   0   'False
               Top             =   1005
               Width           =   14745
            End
            Begin VB.CheckBox Chk_nao_chek 
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
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   0
               Left            =   240
               TabIndex        =   64
               Top             =   795
               Width           =   585
            End
            Begin VB.CheckBox Chk_sim_chek 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Sim - Quais?"
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
               Left            =   900
               TabIndex        =   65
               Top             =   795
               Width           =   1185
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
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
               Index           =   71
               Left            =   9367
               TabIndex        =   330
               Top             =   180
               Width           =   690
            End
            Begin VB.Label Label1 
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
               Index           =   70
               Left            =   2475
               TabIndex        =   329
               Top             =   180
               Width           =   915
            End
            Begin VB.Label Label1 
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
               Index           =   14
               Left            =   405
               TabIndex        =   243
               Top             =   180
               Width           =   345
            End
         End
         Begin VB.TextBox txtID_engenharia 
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
            Height          =   315
            Left            =   3060
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   183
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   4260
            Visible         =   0   'False
            Width           =   675
         End
         Begin VB.Frame Frame7 
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
            Height          =   2325
            Left            =   75
            TabIndex        =   182
            Top             =   1320
            Width           =   15105
            Begin VB.ComboBox Cmb_un_com_engenharia 
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
               ItemData        =   "frmVendas_analise.frx":1BA4F
               Left            =   13020
               List            =   "frmVendas_analise.frx":1BA51
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   56
               ToolTipText     =   "Unidade comercial."
               Top             =   1005
               Width           =   855
            End
            Begin VB.OptionButton Opt_terceiros 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Terceiros"
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
               Left            =   1260
               TabIndex        =   42
               Top             =   450
               Width           =   1095
            End
            Begin VB.OptionButton Opt_outros 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Outros"
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
               Left            =   2400
               TabIndex        =   43
               Top             =   450
               Width           =   885
            End
            Begin VB.OptionButton Opt_material 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Material"
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
               TabIndex        =   41
               Top             =   450
               Width           =   1005
            End
            Begin VB.ComboBox cmbReferencia_engenharia 
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
               Left            =   180
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   51
               ToolTipText     =   "Código de referência."
               Top             =   1005
               Width           =   2115
            End
            Begin VB.Frame Frame12 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Criar novo produto"
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
               Height          =   525
               Left            =   11580
               TabIndex        =   192
               Top             =   180
               Width           =   3345
               Begin VB.CheckBox chkManual_engenharia 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Cód. manual ?"
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
                  Height          =   225
                  Left            =   120
                  TabIndex        =   49
                  Top             =   270
                  Width           =   1335
               End
               Begin VB.CheckBox chkAuto_engenharia 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Cód. automático ?"
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
                  Height          =   225
                  Left            =   1620
                  TabIndex        =   50
                  Top             =   270
                  Width           =   1605
               End
            End
            Begin VB.ComboBox cmbun_engenharia 
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
               ItemData        =   "frmVendas_analise.frx":1BA53
               Left            =   12170
               List            =   "frmVendas_analise.frx":1BA55
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   55
               ToolTipText     =   "Unidade de estoque."
               Top             =   1005
               Width           =   855
            End
            Begin VB.ComboBox cmbfamilia_engenharia 
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
               Left            =   7605
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   54
               ToolTipText     =   "Família."
               Top             =   1005
               Width           =   4560
            End
            Begin VB.CommandButton Cmd_filtrar_produto_engenharia 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   10830
               Picture         =   "frmVendas_analise.frx":1BA57
               Style           =   1  'Graphical
               TabIndex        =   47
               ToolTipText     =   "Filtrar por código interno."
               Top             =   390
               Width           =   315
            End
            Begin VB.CommandButton cmdProduto_engenharia 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   11160
               Picture         =   "frmVendas_analise.frx":1BE72
               Style           =   1  'Graphical
               TabIndex        =   48
               ToolTipText     =   "Localizar produtos/serviços."
               Top             =   390
               Width           =   315
            End
            Begin VB.TextBox txtdesenho_engenharia 
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
               Left            =   8520
               MaxLength       =   50
               TabIndex        =   46
               ToolTipText     =   "Código interno."
               Top             =   390
               Width           =   2295
            End
            Begin VB.TextBox txtQtde_engenharia 
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
               Left            =   13890
               MaxLength       =   20
               TabIndex        =   57
               ToolTipText     =   "Quantidade."
               Top             =   1005
               Width           =   1035
            End
            Begin VB.TextBox txtdescricao_engenharia 
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
               Left            =   2310
               MaxLength       =   255
               TabIndex        =   53
               ToolTipText     =   "Descrição."
               Top             =   1005
               Width           =   5280
            End
            Begin VB.TextBox txtData_engenharia_prod 
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
               Left            =   3420
               Locked          =   -1  'True
               MaxLength       =   25
               TabIndex        =   44
               TabStop         =   0   'False
               ToolTipText     =   "Data do cadastro."
               Top             =   390
               Width           =   1125
            End
            Begin VB.TextBox txtResponsavel_engenharia_prod 
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
               Left            =   4560
               Locked          =   -1  'True
               TabIndex        =   45
               TabStop         =   0   'False
               ToolTipText     =   "Responsável pelo cadastro."
               Top             =   390
               Width           =   3945
            End
            Begin VB.TextBox txtAnalise_engenharia 
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
               Height          =   555
               Left            =   180
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   58
               ToolTipText     =   "Análise crítica."
               Top             =   1620
               Width           =   14745
            End
            Begin VB.TextBox txtReferencia_engenharia 
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
               Left            =   180
               MaxLength       =   50
               TabIndex        =   52
               ToolTipText     =   "Código de referência."
               Top             =   1005
               Visible         =   0   'False
               Width           =   2115
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Qtde."
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
               Index           =   69
               Left            =   14197
               TabIndex        =   328
               Top             =   810
               Width           =   420
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Un. est."
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
               Index           =   68
               Left            =   12305
               TabIndex        =   327
               Top             =   810
               Width           =   585
            End
            Begin VB.Label Label1 
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
               Index           =   67
               Left            =   9645
               TabIndex        =   326
               Top             =   810
               Width           =   480
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
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
               Index           =   66
               Left            =   4605
               TabIndex        =   325
               Top             =   810
               Width           =   690
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Cód. interno"
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
               Index           =   65
               Left            =   9217
               TabIndex        =   324
               Top             =   180
               Width           =   900
            End
            Begin VB.Label Label1 
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
               Index           =   64
               Left            =   6075
               TabIndex        =   323
               Top             =   180
               Width           =   915
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Análise crítica"
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
               Left            =   7065
               TabIndex        =   242
               Top             =   1410
               Width           =   975
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Código de referência"
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
               Left            =   487
               TabIndex        =   241
               Top             =   810
               Width           =   1500
            End
            Begin VB.Label Label1 
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
               Index           =   10
               Left            =   3810
               TabIndex        =   240
               Top             =   180
               Width           =   345
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Un. com."
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
               Left            =   13125
               TabIndex        =   239
               Top             =   810
               Width           =   645
            End
         End
         Begin MSComctlLib.ListView Lista_engenharia 
            Height          =   2955
            Left            =   75
            TabIndex        =   59
            Top             =   3660
            Width           =   15105
            _ExtentX        =   26644
            _ExtentY        =   5212
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
            NumItems        =   8
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
               Object.Width           =   13961
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
               Alignment       =   1
               SubItemIndex    =   5
               Object.Tag             =   "N"
               Text            =   "Qtde."
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   6
               Object.Tag             =   "N"
               Text            =   "Valor unit."
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   7
               Object.Tag             =   "N"
               Text            =   "Valor total"
               Object.Width           =   2117
            EndProperty
         End
         Begin MSComctlLib.ListView Lista_normas 
            Height          =   3315
            Left            =   -74925
            TabIndex        =   72
            Top             =   3300
            Width           =   15105
            _ExtentX        =   26644
            _ExtentY        =   5847
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
               Alignment       =   2
               SubItemIndex    =   1
               Object.Tag             =   "D"
               Text            =   "Data"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Object.Tag             =   "T"
               Text            =   "Responsável"
               Object.Width           =   3158
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Object.Tag             =   "T"
               Text            =   "Norma"
               Object.Width           =   20152
            EndProperty
         End
         Begin DrawSuite2022.USToolBar USToolBar4 
            Height          =   975
            Left            =   75
            TabIndex        =   266
            Top             =   330
            Width           =   15105
            _ExtentX        =   26644
            _ExtentY        =   1720
            ButtonCount     =   5
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
            ButtonCaption4  =   "Copiar"
            ButtonEnabled4  =   0   'False
            ButtonIconSize4 =   32
            ButtonToolTipText4=   "Copiar material/terceiro da engenharia."
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
            ButtonWidth4    =   44
            ButtonHeight4   =   21
            ButtonUseMaskColor4=   0   'False
            ButtonEnabled5  =   0   'False
            ButtonIconSize5 =   32
            ButtonKey5      =   "5"
            ButtonAlignment5=   2
            BeginProperty ButtonFont5 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonState5    =   5
            ButtonLeft5     =   179
            ButtonTop5      =   2
            ButtonWidth5    =   24
            ButtonHeight5   =   24
            ButtonUseMaskColor5=   0   'False
            Begin DrawSuite2022.USImageList USImageList4 
               Left            =   10980
               Top             =   270
               _ExtentX        =   900
               _ExtentY        =   767
               Img1            =   "frmVendas_analise.frx":1BF74
               Count           =   1
            End
         End
         Begin DrawSuite2022.USToolBar USToolBar6 
            Height          =   975
            Left            =   -74925
            TabIndex        =   267
            Top             =   330
            Width           =   15105
            _ExtentX        =   26644
            _ExtentY        =   1720
            ButtonCount     =   4
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
            ButtonKey4      =   "4"
            ButtonAlignment4=   2
            BeginProperty ButtonFont4 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonState4    =   5
            ButtonLeft4     =   133
            ButtonTop4      =   2
            ButtonWidth4    =   24
            ButtonHeight4   =   24
            ButtonUseMaskColor4=   0   'False
            Begin DrawSuite2022.USImageList USImageList6 
               Left            =   13170
               Top             =   210
               _ExtentX        =   900
               _ExtentY        =   767
               Img1            =   "frmVendas_analise.frx":1E90B
               Count           =   1
            End
         End
         Begin DrawSuite2022.USToolBar USToolBar5 
            Height          =   975
            Left            =   -74925
            TabIndex        =   268
            Top             =   330
            Width           =   15105
            _ExtentX        =   26644
            _ExtentY        =   1720
            ButtonCount     =   4
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
            ButtonKey4      =   "4"
            ButtonAlignment4=   2
            BeginProperty ButtonFont4 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonState4    =   5
            ButtonLeft4     =   133
            ButtonTop4      =   2
            ButtonWidth4    =   24
            ButtonHeight4   =   24
            ButtonUseMaskColor4=   0   'False
            Begin DrawSuite2022.USImageList USImageList5 
               Left            =   13170
               Top             =   210
               _ExtentX        =   900
               _ExtentY        =   767
               Img1            =   "frmVendas_analise.frx":206C6
               Count           =   1
            End
         End
         Begin MSComctlLib.ListView Lista_check 
            Height          =   3315
            Index           =   0
            Left            =   -74925
            TabIndex        =   67
            Top             =   3300
            Width           =   15105
            _ExtentX        =   26644
            _ExtentY        =   5847
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
               Alignment       =   2
               SubItemIndex    =   1
               Object.Tag             =   "D"
               Text            =   "Data"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Object.Tag             =   "T"
               Text            =   "Responsável"
               Object.Width           =   4410
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Object.Tag             =   "T"
               Text            =   "Descrição"
               Object.Width           =   16783
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   4
               Object.Tag             =   "T"
               Text            =   "Executar"
               Object.Width           =   2117
            EndProperty
         End
      End
      Begin VB.Frame Frame20 
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
         Height          =   1755
         Left            =   -74925
         TabIndex        =   185
         Top             =   1320
         Width           =   15195
         Begin VB.TextBox txtRespValidacao_PCP 
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
            Left            =   9660
            Locked          =   -1  'True
            TabIndex        =   107
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pela validação."
            Top             =   390
            Width           =   5355
         End
         Begin VB.TextBox txtPrazo_PCP 
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
            Left            =   6780
            Locked          =   -1  'True
            TabIndex        =   105
            TabStop         =   0   'False
            ToolTipText     =   "Prazo."
            Top             =   390
            Width           =   1185
         End
         Begin VB.TextBox txtDtValidacao_PCP 
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
            Left            =   7980
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   106
            TabStop         =   0   'False
            ToolTipText     =   "Data do cadastro."
            Top             =   390
            Width           =   1665
         End
         Begin VB.TextBox txtData_PCP 
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
            TabIndex        =   103
            TabStop         =   0   'False
            ToolTipText     =   "Data do cadastro."
            Top             =   390
            Width           =   1185
         End
         Begin VB.TextBox txtResponsavel_PCP 
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
            TabIndex        =   104
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pelo cadastro."
            Top             =   390
            Width           =   5385
         End
         Begin VB.TextBox txtAnalise_PCP 
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
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   108
            ToolTipText     =   "Análise crítica."
            Top             =   1002
            Width           =   14835
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
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
            Index           =   90
            Left            =   11347
            TabIndex        =   349
            Top             =   180
            Width           =   1980
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Data/hora validação"
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
            Index           =   89
            Left            =   8085
            TabIndex        =   348
            Top             =   180
            Width           =   1455
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Prazo"
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
            Index           =   88
            Left            =   7170
            TabIndex        =   347
            Top             =   180
            Width           =   405
         End
         Begin VB.Label Label1 
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
            Index           =   87
            Left            =   3615
            TabIndex        =   346
            Top             =   180
            Width           =   915
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Análise crítica"
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
            Left            =   7095
            TabIndex        =   233
            Top             =   810
            Width           =   975
         End
         Begin VB.Label Label1 
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
            Index           =   27
            Left            =   600
            TabIndex        =   232
            Top             =   180
            Width           =   345
         End
      End
      Begin VB.Frame Frame14 
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
         Height          =   2265
         Left            =   -74925
         TabIndex        =   197
         Top             =   1320
         Width           =   15195
         Begin VB.CommandButton Cmd_limpar_caminho 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   14370
            Picture         =   "frmVendas_analise.frx":22481
            Style           =   1  'Graphical
            TabIndex        =   174
            ToolTipText     =   "Limpar caminho."
            Top             =   390
            Width           =   315
         End
         Begin VB.CommandButton Cmd_visualizar_arquivo 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   14700
            Picture         =   "frmVendas_analise.frx":225BF
            Style           =   1  'Graphical
            TabIndex        =   175
            ToolTipText     =   "Visualizar arquivo."
            Top             =   390
            Width           =   315
         End
         Begin VB.TextBox Txt_obs_doc 
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
            Height          =   1095
            Left            =   180
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   176
            ToolTipText     =   "Observação."
            Top             =   1020
            Width           =   14835
         End
         Begin VB.TextBox Txt_caminho 
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
            Left            =   3000
            Locked          =   -1  'True
            TabIndex        =   172
            TabStop         =   0   'False
            ToolTipText     =   "Caminho."
            Top             =   390
            Width           =   11025
         End
         Begin VB.CommandButton cmdImportar 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   14040
            Picture         =   "frmVendas_analise.frx":22B81
            Style           =   1  'Graphical
            TabIndex        =   173
            ToolTipText     =   "Localizar arquivo (F2)"
            Top             =   390
            Width           =   315
         End
         Begin VB.TextBox txtData_doc 
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
            TabIndex        =   170
            TabStop         =   0   'False
            ToolTipText     =   "Data do cadastro."
            Top             =   390
            Width           =   855
         End
         Begin VB.TextBox txtResponsavel_doc 
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
            Left            =   1050
            Locked          =   -1  'True
            TabIndex        =   171
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pelo cadastro."
            Top             =   390
            Width           =   1935
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Index           =   119
            Left            =   7800
            TabIndex        =   378
            Top             =   180
            Width           =   1425
         End
         Begin VB.Label Label1 
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
            Index           =   118
            Left            =   1560
            TabIndex        =   377
            Top             =   180
            Width           =   915
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
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
            Index           =   39
            Left            =   7155
            TabIndex        =   199
            Top             =   810
            Width           =   945
         End
         Begin VB.Label Label1 
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
            Index           =   38
            Left            =   450
            TabIndex        =   198
            Top             =   180
            Width           =   345
         End
      End
      Begin MSComctlLib.ListView Lista_doc 
         Height          =   6105
         Left            =   -74925
         TabIndex        =   177
         Top             =   3600
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   10769
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
            Text            =   "Caminho"
            Object.Width           =   25576
         EndProperty
      End
      Begin VB.Frame Frame10 
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
         Height          =   825
         Left            =   -74925
         TabIndex        =   189
         Top             =   1320
         Width           =   15195
         Begin VB.TextBox txtRespValidacao_Qualidade 
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
            Left            =   6390
            Locked          =   -1  'True
            TabIndex        =   119
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pela validação."
            Top             =   390
            Width           =   5565
         End
         Begin VB.TextBox txtPrazo_Qualidade 
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
            Left            =   3060
            Locked          =   -1  'True
            TabIndex        =   117
            TabStop         =   0   'False
            ToolTipText     =   "Prazo."
            Top             =   390
            Width           =   1545
         End
         Begin VB.TextBox txtDtValidacao_Qualidade 
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
            Left            =   4620
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   118
            TabStop         =   0   'False
            ToolTipText     =   "Data do cadastro."
            Top             =   390
            Width           =   1755
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
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
            Index           =   94
            Left            =   8182
            TabIndex        =   353
            Top             =   180
            Width           =   1980
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Data/hora validação"
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
            Index           =   93
            Left            =   4770
            TabIndex        =   352
            Top             =   180
            Width           =   1455
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Prazo"
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
            Left            =   3630
            TabIndex        =   286
            Top             =   180
            Width           =   405
         End
      End
      Begin DrawSuite2022.USProgressBar PBLista_compras 
         Height          =   255
         Left            =   -74940
         TabIndex        =   219
         Top             =   9720
         Width           =   6795
         _ExtentX        =   11986
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
      Begin TabDlg.SSTab SSTab_processo 
         Height          =   8745
         Left            =   60
         TabIndex        =   207
         Top             =   2160
         Width           =   15255
         _ExtentX        =   26908
         _ExtentY        =   15425
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
         TabCaption(0)   =   "Produtos"
         TabPicture(0)   =   "frmVendas_analise.frx":22C83
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "optProcessos_item_analise"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "optProcessos_item"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "txtID_processos_item"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "txtIDproduto_processos"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Frame2"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Lista_processos_item"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).ControlCount=   6
         TabCaption(1)   =   "Fases"
         TabPicture(1)   =   "frmVendas_analise.frx":22C9F
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Label1(45)"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "PBLista_Fases"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "lista_Processos"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "Frame6"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "txtID_processos"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).Control(5)=   "Txt_valor_total_processo"
         Tab(1).Control(5).Enabled=   0   'False
         Tab(1).ControlCount=   6
         Begin VB.OptionButton optProcessos_item_analise 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Produto da análise"
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
            Height          =   210
            Left            =   -74760
            TabIndex        =   222
            Top             =   360
            Width           =   1845
         End
         Begin VB.OptionButton optProcessos_item 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Produto da estrutura"
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
            Height          =   210
            Left            =   -72780
            TabIndex        =   221
            Top             =   360
            Width           =   2025
         End
         Begin VB.TextBox Txt_valor_total_processo 
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
            Left            =   13560
            Locked          =   -1  'True
            MaxLength       =   20
            MouseIcon       =   "frmVendas_analise.frx":22CBB
            TabIndex        =   102
            TabStop         =   0   'False
            ToolTipText     =   "Valor total do processo."
            Top             =   7530
            Width           =   1620
         End
         Begin VB.TextBox txtID_processos_item 
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
            Left            =   -68010
            Locked          =   -1  'True
            MouseIcon       =   "frmVendas_analise.frx":22FC5
            MousePointer    =   99  'Custom
            TabIndex        =   217
            TabStop         =   0   'False
            Text            =   "0"
            ToolTipText     =   "Data do cadastro."
            Top             =   4050
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox txtIDproduto_processos 
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
            Left            =   -68820
            Locked          =   -1  'True
            TabIndex        =   216
            TabStop         =   0   'False
            ToolTipText     =   "Data do cadastro."
            Top             =   4050
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Frame Frame2 
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
            Height          =   1515
            Left            =   -74925
            TabIndex        =   214
            Top             =   330
            Width           =   15105
            Begin VB.ComboBox Cmb_un_com_processos_item 
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
               ItemData        =   "frmVendas_analise.frx":232CF
               Left            =   13110
               List            =   "frmVendas_analise.frx":232D1
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   79
               ToolTipText     =   "Unidade comercial."
               Top             =   460
               Width           =   855
            End
            Begin VB.TextBox txtQtde_processos_item 
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
               Left            =   13980
               TabIndex        =   80
               Text            =   "1,000"
               ToolTipText     =   "Quantidade."
               Top             =   460
               Width           =   945
            End
            Begin VB.TextBox txtDescricao_processos_item 
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
               MaxLength       =   255
               TabIndex        =   81
               ToolTipText     =   "Descrição."
               Top             =   1065
               Width           =   14730
            End
            Begin VB.TextBox txtCodInterno_processos_item 
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
               TabIndex        =   73
               ToolTipText     =   "Código interno."
               Top             =   460
               Width           =   1905
            End
            Begin VB.CommandButton cmdProduto_processos 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   2430
               Picture         =   "frmVendas_analise.frx":232D3
               Style           =   1  'Graphical
               TabIndex        =   75
               ToolTipText     =   "Localizar produtos."
               Top             =   460
               Width           =   315
            End
            Begin VB.CommandButton cmdPesquisar_preocessos_item 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   2100
               Picture         =   "frmVendas_analise.frx":233D5
               Style           =   1  'Graphical
               TabIndex        =   74
               ToolTipText     =   "Filtrar por código interno."
               Top             =   460
               Width           =   315
            End
            Begin VB.ComboBox cmbFamilia_processos_item 
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
               Left            =   4995
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   77
               ToolTipText     =   "Família."
               Top             =   460
               Width           =   7250
            End
            Begin VB.ComboBox cmbUn_processos_item 
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
               ItemData        =   "frmVendas_analise.frx":237F0
               Left            =   12240
               List            =   "frmVendas_analise.frx":237F2
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   78
               ToolTipText     =   "Unidade de estoque."
               Top             =   460
               Width           =   855
            End
            Begin VB.ComboBox cmbReferencia_processos_item 
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
               Left            =   2850
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   76
               ToolTipText     =   "Código de referência."
               Top             =   460
               Width           =   2145
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Un. com."
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
               Index           =   123
               Left            =   13215
               TabIndex        =   382
               Top             =   270
               Width           =   645
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Código de referência"
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
               Index           =   122
               Left            =   3172
               TabIndex        =   381
               Top             =   270
               Width           =   1500
            End
            Begin VB.Label Label1 
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
               Index           =   121
               Left            =   8380
               TabIndex        =   380
               Top             =   270
               Width           =   480
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Un. est."
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
               Index           =   120
               Left            =   12375
               TabIndex        =   379
               Top             =   270
               Width           =   585
            End
            Begin VB.Label Label1 
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
               Index           =   22
               Left            =   7200
               TabIndex        =   227
               Top             =   870
               Width           =   690
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Qtde."
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
               Left            =   14287
               TabIndex        =   220
               Top             =   270
               Width           =   420
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Cód. interno"
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
               Left            =   682
               TabIndex        =   215
               Top             =   270
               Width           =   900
            End
         End
         Begin VB.TextBox txtID_processos 
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
            Height          =   315
            Left            =   1500
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   213
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   3870
            Visible         =   0   'False
            Width           =   675
         End
         Begin VB.Frame Frame6 
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
            Height          =   2235
            Left            =   75
            TabIndex        =   208
            Top             =   330
            Width           =   15105
            Begin VB.TextBox txtValorHoraPrep_Processos 
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
               Left            =   11580
               MaxLength       =   255
               TabIndex        =   97
               ToolTipText     =   "Valor por hora de preparação."
               Top             =   975
               Width           =   990
            End
            Begin VB.TextBox txtErro 
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
               MaxLength       =   20
               TabIndex        =   94
               ToolTipText     =   "Porcentagem de erro."
               Top             =   975
               Width           =   495
            End
            Begin VB.TextBox txtFase 
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
               Left            =   10050
               TabIndex        =   85
               ToolTipText     =   "Fase."
               Top             =   390
               Width           =   465
            End
            Begin VB.TextBox txtgrupo_op 
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
               Left            =   10530
               TabIndex        =   86
               ToolTipText     =   "Grupo/operação."
               Top             =   390
               Width           =   1545
            End
            Begin VB.TextBox txtValorTotal_processos 
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
               Left            =   13590
               Locked          =   -1  'True
               MaxLength       =   20
               TabIndex        =   99
               TabStop         =   0   'False
               ToolTipText     =   "Valor total."
               Top             =   975
               Width           =   1335
            End
            Begin VB.TextBox txtValorHora_processos 
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
               Left            =   12585
               MaxLength       =   255
               TabIndex        =   98
               ToolTipText     =   "Custo por hora de execução."
               Top             =   975
               Width           =   990
            End
            Begin VB.Frame Frame23 
               BackColor       =   &H00E0E0E0&
               Height          =   435
               Left            =   6330
               TabIndex        =   209
               Top             =   870
               Width           =   1350
               Begin VB.CheckBox chkPchora 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Pçs x hora?"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   210
                  Left            =   120
                  TabIndex        =   91
                  Top             =   180
                  Width           =   1185
               End
            End
            Begin VB.TextBox txtData_processos 
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
               TabIndex        =   83
               TabStop         =   0   'False
               ToolTipText     =   "Data do cadastro."
               Top             =   390
               Width           =   1395
            End
            Begin VB.TextBox txtResponsavel_processos 
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
               Left            =   1590
               Locked          =   -1  'True
               TabIndex        =   84
               TabStop         =   0   'False
               ToolTipText     =   "Responsável pelo cadastro."
               Top             =   390
               Width           =   8445
            End
            Begin VB.CommandButton cmdFiltrar_processos 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   14280
               Picture         =   "frmVendas_analise.frx":237F4
               Style           =   1  'Graphical
               TabIndex        =   88
               ToolTipText     =   "Filtrar por posto de trabalho."
               Top             =   390
               Width           =   315
            End
            Begin VB.TextBox txtTotalHora_processos 
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
               Left            =   10560
               Locked          =   -1  'True
               MaxLength       =   20
               TabIndex        =   96
               TabStop         =   0   'False
               ToolTipText     =   "Tempo de execução por peça."
               Top             =   975
               Width           =   1005
            End
            Begin VB.CommandButton cmdMaquina_processos 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   14610
               Picture         =   "frmVendas_analise.frx":23C0F
               Style           =   1  'Graphical
               TabIndex        =   89
               ToolTipText     =   "Localizar postos de trabalho."
               Top             =   390
               Width           =   315
            End
            Begin VB.TextBox txtDescricao_processos 
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
               MaxLength       =   255
               TabIndex        =   90
               ToolTipText     =   "Descrição."
               Top             =   975
               Width           =   6105
            End
            Begin VB.TextBox txtMaquina_processos 
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
               Left            =   12090
               MaxLength       =   50
               TabIndex        =   87
               ToolTipText     =   "Posto de trabalho."
               Top             =   390
               Width           =   2205
            End
            Begin VB.TextBox txtPecaHora_processos 
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
               Locked          =   -1  'True
               MaxLength       =   20
               TabIndex        =   95
               TabStop         =   0   'False
               Text            =   "1"
               ToolTipText     =   "Peça por hora."
               Top             =   975
               Width           =   525
            End
            Begin MSMask.MaskEdBox txtPreparacao_processos 
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
               Left            =   7740
               TabIndex        =   92
               ToolTipText     =   "Tempo de preparação previsto."
               Top             =   975
               Width           =   885
               _ExtentX        =   1561
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
            Begin RichTextLib.RichTextBox txtTrabalho 
               Height          =   525
               Left            =   180
               TabIndex        =   100
               ToolTipText     =   "Instruções de trabalho."
               Top             =   1560
               Width           =   14745
               _ExtentX        =   26009
               _ExtentY        =   926
               _Version        =   393217
               BorderStyle     =   0
               Enabled         =   -1  'True
               ScrollBars      =   2
               TextRTF         =   $"frmVendas_analise.frx":23D11
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
            Begin MSMask.MaskEdBox txtExecucao_processos 
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
               Left            =   8640
               TabIndex        =   93
               ToolTipText     =   "Tempo de execução previsto."
               Top             =   975
               Width           =   855
               _ExtentX        =   1508
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
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Valor total"
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
               Index           =   86
               Left            =   13890
               TabIndex        =   345
               Top             =   780
               Width           =   735
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Vlr. hr. exe."
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
               Index           =   85
               Left            =   12645
               TabIndex        =   344
               Top             =   780
               Width           =   870
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Vlr. hr. prep."
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
               Index           =   84
               Left            =   11610
               TabIndex        =   343
               Top             =   780
               Width           =   930
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Exec. x peça"
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
               Index           =   83
               Left            =   10597
               TabIndex        =   342
               Top             =   780
               Width           =   930
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Peça"
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
               Index           =   82
               Left            =   10110
               TabIndex        =   341
               Top             =   780
               Width           =   345
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Erro%"
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
               Index           =   81
               Left            =   9525
               TabIndex        =   340
               Top             =   780
               Width           =   465
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Execução"
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
               Index           =   80
               Left            =   8707
               TabIndex        =   339
               Top             =   780
               Width           =   690
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Preparação"
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
               Index           =   79
               Left            =   7770
               TabIndex        =   338
               Top             =   780
               Width           =   825
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Posto de trabalho"
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
               Index           =   78
               Left            =   12555
               TabIndex        =   337
               Top             =   180
               Width           =   1275
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Grupo/op"
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
               Index           =   77
               Left            =   10965
               TabIndex        =   336
               Top             =   180
               Width           =   675
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Fase"
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
               Index           =   76
               Left            =   10110
               TabIndex        =   335
               Top             =   180
               Width           =   345
            End
            Begin VB.Label Label1 
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
               Index           =   75
               Left            =   5355
               TabIndex        =   334
               Top             =   180
               Width           =   915
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "Instruções de trabalho"
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
               Left            =   6780
               TabIndex        =   212
               Top             =   1350
               Width           =   1635
            End
            Begin VB.Label Label1 
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
               Index           =   23
               Left            =   705
               TabIndex        =   211
               Top             =   180
               Width           =   345
            End
            Begin VB.Label Label1 
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
               Index           =   25
               Left            =   2887
               TabIndex        =   210
               Top             =   780
               Width           =   690
            End
         End
         Begin MSComctlLib.ListView Lista_processos_item 
            Height          =   5685
            Left            =   -74925
            TabIndex        =   82
            Top             =   1860
            Width           =   15105
            _ExtentX        =   26644
            _ExtentY        =   10028
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
            NumItems        =   7
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
               Object.Width           =   13432
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Object.Tag             =   "T"
               Text            =   "Família"
               Object.Width           =   4762
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   4
               Object.Tag             =   "T"
               Text            =   "Un. est."
               Object.Width           =   1499
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   5
               Object.Tag             =   "T"
               Text            =   "Un. com."
               Object.Width           =   1499
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   6
               Object.Tag             =   "N"
               Text            =   "Qtde."
               Object.Width           =   2117
            EndProperty
         End
         Begin MSComctlLib.ListView lista_Processos 
            Height          =   4935
            Left            =   75
            TabIndex        =   101
            Top             =   2580
            Width           =   15105
            _ExtentX        =   26644
            _ExtentY        =   8705
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
            NumItems        =   8
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Tag             =   "N"
               Object.Width           =   529
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Object.Tag             =   "T"
               Text            =   "Fase"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Object.Tag             =   "T"
               Text            =   "Posto de trab."
               Object.Width           =   2813
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Object.Tag             =   "T"
               Text            =   "Descrição do posto"
               Object.Width           =   10266
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Object.Tag             =   "D"
               Text            =   "Exec. x peça"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   5
               Object.Tag             =   "N"
               Text            =   "Valor hora prep."
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   6
               Object.Tag             =   "N"
               Text            =   "Valor hora exec."
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   7
               Object.Tag             =   "N"
               Text            =   "Valor total"
               Object.Width           =   2646
            EndProperty
         End
         Begin DrawSuite2022.USProgressBar PBLista_Fases 
            Height          =   255
            Left            =   30
            TabIndex        =   301
            Top             =   7560
            Width           =   12255
            _ExtentX        =   21616
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
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Valor total :"
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
            Index           =   45
            Left            =   12480
            TabIndex        =   218
            Top             =   7530
            Width           =   1935
            WordWrap        =   -1  'True
         End
      End
      Begin TabDlg.SSTab SSTab4 
         Height          =   7995
         Left            =   -74940
         TabIndex        =   275
         Top             =   3090
         Width           =   15255
         _ExtentX        =   26908
         _ExtentY        =   14102
         _Version        =   393216
         Tabs            =   1
         TabsPerRow      =   1
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
         TabCaption(0)   =   "Check-list"
         TabPicture(0)   =   "frmVendas_analise.frx":23D8F
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Lista_check(1)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "USToolBar8"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Txt_ID_check(1)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Frame_check(1)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).ControlCount=   4
         Begin VB.Frame Frame_check 
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
            Height          =   1965
            Index           =   1
            Left            =   75
            TabIndex        =   278
            Top             =   1320
            Width           =   15105
            Begin VB.TextBox Txt_ID_descricao_check 
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
               Height          =   315
               Index           =   1
               Left            =   4860
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   288
               TabStop         =   0   'False
               Text            =   "0"
               Top             =   390
               Visible         =   0   'False
               Width           =   675
            End
            Begin VB.CheckBox Chk_sim_chek 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Sim - Quais?"
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
               Left            =   900
               TabIndex        =   114
               Top             =   795
               Width           =   1185
            End
            Begin VB.CheckBox Chk_nao_chek 
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
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   1
               Left            =   240
               TabIndex        =   113
               Top             =   795
               Width           =   585
            End
            Begin VB.TextBox Txt_texto_check 
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
               Height          =   825
               Index           =   1
               Left            =   150
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   115
               TabStop         =   0   'False
               Top             =   1005
               Width           =   14745
            End
            Begin VB.TextBox Txt_responsavel_check 
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
               Index           =   1
               Left            =   1020
               Locked          =   -1  'True
               TabIndex        =   110
               TabStop         =   0   'False
               ToolTipText     =   "Responsável pelo cadastro."
               Top             =   390
               Width           =   3825
            End
            Begin VB.TextBox Txt_data_check 
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
               Index           =   1
               Left            =   150
               Locked          =   -1  'True
               MaxLength       =   25
               TabIndex        =   109
               TabStop         =   0   'False
               ToolTipText     =   "Data do cadastro."
               Top             =   390
               Width           =   855
            End
            Begin VB.Frame Frame16 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
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
               Height          =   405
               Left            =   5640
               TabIndex        =   279
               Top             =   4650
               Width           =   2985
            End
            Begin VB.TextBox Txt_descricao_chek 
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
               Index           =   1
               Left            =   4860
               Locked          =   -1  'True
               TabIndex        =   111
               TabStop         =   0   'False
               ToolTipText     =   "Descrição."
               Top             =   390
               Width           =   9705
            End
            Begin VB.CommandButton Cmd_localizar_desc_check 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Index           =   1
               Left            =   14580
               Picture         =   "frmVendas_analise.frx":23DAB
               Style           =   1  'Graphical
               TabIndex        =   112
               ToolTipText     =   "Localizar descrição padrão do check-list."
               Top             =   390
               Width           =   315
            End
            Begin VB.Label Label1 
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
               Index           =   92
               Left            =   9367
               TabIndex        =   351
               Top             =   180
               Width           =   690
            End
            Begin VB.Label Label1 
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
               Index           =   91
               Left            =   2475
               TabIndex        =   350
               Top             =   180
               Width           =   915
            End
            Begin VB.Label Label1 
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
               Index           =   16
               Left            =   405
               TabIndex        =   280
               Top             =   180
               Width           =   345
            End
         End
         Begin VB.TextBox Txt_ID_check 
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
            Height          =   315
            Index           =   1
            Left            =   3420
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   276
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   4830
            Visible         =   0   'False
            Width           =   675
         End
         Begin DrawSuite2022.USToolBar USToolBar8 
            Height          =   975
            Left            =   75
            TabIndex        =   277
            Top             =   330
            Width           =   15105
            _ExtentX        =   26644
            _ExtentY        =   1720
            ButtonCount     =   4
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
            ButtonKey4      =   "4"
            ButtonAlignment4=   2
            BeginProperty ButtonFont4 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonState4    =   5
            ButtonLeft4     =   133
            ButtonTop4      =   2
            ButtonWidth4    =   24
            ButtonHeight4   =   24
            ButtonUseMaskColor4=   0   'False
            Begin DrawSuite2022.USImageList USImageList8 
               Left            =   13170
               Top             =   210
               _ExtentX        =   900
               _ExtentY        =   767
               Img1            =   "frmVendas_analise.frx":23EAD
               Count           =   1
            End
         End
         Begin MSComctlLib.ListView Lista_check 
            Height          =   3315
            Index           =   1
            Left            =   75
            TabIndex        =   116
            Top             =   3300
            Width           =   15105
            _ExtentX        =   26644
            _ExtentY        =   5847
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
               Alignment       =   2
               SubItemIndex    =   1
               Object.Tag             =   "D"
               Text            =   "Data"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Object.Tag             =   "D"
               Text            =   "Responsável"
               Object.Width           =   4410
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Object.Tag             =   "T"
               Text            =   "Descrição"
               Object.Width           =   16783
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   4
               Object.Tag             =   "T"
               Text            =   "Executar"
               Object.Width           =   2117
            EndProperty
         End
      End
      Begin TabDlg.SSTab SSTab_qualidade 
         Height          =   8025
         Left            =   -74940
         TabIndex        =   188
         Top             =   2160
         Width           =   15255
         _ExtentX        =   26908
         _ExtentY        =   14155
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
         TabCaption(0)   =   "Instrumentos"
         TabPicture(0)   =   "frmVendas_analise.frx":25C68
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Lista_Qualidade"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Frame8"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "txtIDproduto_qualidade"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "txtID_qualidade"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).ControlCount=   4
         TabCaption(1)   =   "Check-list"
         TabPicture(1)   =   "frmVendas_analise.frx":25C84
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Lista_check(2)"
         Tab(1).Control(1)=   "Txt_ID_check(2)"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "Frame_check(2)"
         Tab(1).ControlCount=   3
         Begin VB.Frame Frame_check 
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
            Height          =   1965
            Index           =   2
            Left            =   -74925
            TabIndex        =   282
            Top             =   330
            Width           =   15105
            Begin VB.TextBox Txt_ID_descricao_check 
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
               Height          =   315
               Index           =   2
               Left            =   4860
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   289
               TabStop         =   0   'False
               Text            =   "0"
               Top             =   390
               Visible         =   0   'False
               Width           =   675
            End
            Begin VB.CommandButton Cmd_localizar_desc_check 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Index           =   2
               Left            =   14580
               Picture         =   "frmVendas_analise.frx":25CA0
               Style           =   1  'Graphical
               TabIndex        =   138
               ToolTipText     =   "Localizar descrição padrão do check-list."
               Top             =   390
               Width           =   315
            End
            Begin VB.TextBox Txt_descricao_chek 
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
               Index           =   2
               Left            =   4860
               Locked          =   -1  'True
               TabIndex        =   137
               TabStop         =   0   'False
               ToolTipText     =   "Descrição."
               Top             =   390
               Width           =   9705
            End
            Begin VB.Frame Frame11 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
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
               Height          =   405
               Left            =   5640
               TabIndex        =   283
               Top             =   4650
               Width           =   2985
            End
            Begin VB.TextBox Txt_data_check 
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
               Index           =   2
               Left            =   150
               Locked          =   -1  'True
               MaxLength       =   25
               TabIndex        =   135
               TabStop         =   0   'False
               ToolTipText     =   "Data do cadastro."
               Top             =   390
               Width           =   855
            End
            Begin VB.TextBox Txt_responsavel_check 
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
               Index           =   2
               Left            =   1020
               Locked          =   -1  'True
               TabIndex        =   136
               TabStop         =   0   'False
               ToolTipText     =   "Responsável pelo cadastro."
               Top             =   390
               Width           =   3825
            End
            Begin VB.TextBox Txt_texto_check 
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
               Height          =   825
               Index           =   2
               Left            =   150
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   141
               TabStop         =   0   'False
               Top             =   1005
               Width           =   14745
            End
            Begin VB.CheckBox Chk_nao_chek 
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
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   2
               Left            =   240
               TabIndex        =   139
               Top             =   795
               Width           =   585
            End
            Begin VB.CheckBox Chk_sim_chek 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Sim - Quais?"
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
               Left            =   900
               TabIndex        =   140
               Top             =   795
               Width           =   1185
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   $"frmVendas_analise.frx":25DA2
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
               Left            =   435
               TabIndex        =   284
               Top             =   180
               Width           =   9780
            End
         End
         Begin VB.TextBox Txt_ID_check 
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
            Height          =   315
            Index           =   2
            Left            =   -71565
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   281
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   4830
            Visible         =   0   'False
            Width           =   675
         End
         Begin VB.TextBox txtID_qualidade 
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
            Height          =   315
            Left            =   3840
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   196
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   4350
            Visible         =   0   'False
            Width           =   675
         End
         Begin VB.TextBox txtIDproduto_qualidade 
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
            Left            =   4530
            Locked          =   -1  'True
            TabIndex        =   195
            TabStop         =   0   'False
            ToolTipText     =   "Data do cadastro."
            Top             =   4350
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Frame Frame8 
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
            Height          =   2205
            Left            =   75
            TabIndex        =   190
            Top             =   330
            Width           =   15105
            Begin VB.ComboBox Cmb_un_com_qualidade 
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
               ItemData        =   "frmVendas_analise.frx":25E6C
               Left            =   12750
               List            =   "frmVendas_analise.frx":25E6E
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   131
               ToolTipText     =   "Unidade comercial."
               Top             =   945
               Width           =   855
            End
            Begin VB.ComboBox cmbReferencia_qualidade 
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
               Left            =   7830
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   125
               ToolTipText     =   "Código de referência."
               Top             =   390
               Width           =   3645
            End
            Begin VB.Frame Frame13 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Criar novo produto"
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
               Height          =   525
               Left            =   11565
               TabIndex        =   193
               Top             =   180
               Width           =   3345
               Begin VB.CheckBox chkAuto_qualidade 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Cód. automático ?"
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
                  Height          =   225
                  Left            =   1620
                  TabIndex        =   127
                  Top             =   270
                  Width           =   1605
               End
               Begin VB.CheckBox chkManual_qualidade 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Cód. manual ?"
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
                  Height          =   225
                  Left            =   120
                  TabIndex        =   126
                  Top             =   270
                  Width           =   1335
               End
            End
            Begin VB.ComboBox cmbun_qualidade 
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
               ItemData        =   "frmVendas_analise.frx":25E70
               Left            =   11895
               List            =   "frmVendas_analise.frx":25E72
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   130
               ToolTipText     =   "Unidade de estoque."
               Top             =   945
               Width           =   855
            End
            Begin VB.ComboBox cmbfamilia_qualidade 
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
               Left            =   6890
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   129
               ToolTipText     =   "Família."
               Top             =   945
               Width           =   5010
            End
            Begin VB.TextBox txtdesenho_qualidade 
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
               Left            =   5400
               MaxLength       =   50
               TabIndex        =   122
               ToolTipText     =   "Código interno."
               Top             =   390
               Width           =   1665
            End
            Begin VB.CommandButton cmdProduto_qualidade 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   7410
               Picture         =   "frmVendas_analise.frx":25E74
               Style           =   1  'Graphical
               TabIndex        =   124
               ToolTipText     =   "Localizar produtos."
               Top             =   390
               Width           =   315
            End
            Begin VB.CommandButton Cmd_filtrar_produto_qualidade 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   7080
               Picture         =   "frmVendas_analise.frx":25F76
               Style           =   1  'Graphical
               TabIndex        =   123
               ToolTipText     =   "Filtrar por código interno."
               Top             =   390
               Width           =   315
            End
            Begin VB.TextBox txtAnalise_qualidade 
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
               Height          =   495
               Left            =   180
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   133
               ToolTipText     =   "Análise crítica"
               Top             =   1560
               Width           =   14730
            End
            Begin VB.TextBox txtData_Qualidade1 
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
               TabIndex        =   120
               TabStop         =   0   'False
               ToolTipText     =   "Data do cadastro."
               Top             =   390
               Width           =   1365
            End
            Begin VB.TextBox txtResponsavel_Qualidade1 
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
               Left            =   1560
               Locked          =   -1  'True
               TabIndex        =   121
               TabStop         =   0   'False
               ToolTipText     =   "Responsável pelo cadastro."
               Top             =   390
               Width           =   3825
            End
            Begin VB.TextBox txtQtde_Qualidade 
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
               Left            =   13605
               MaxLength       =   20
               TabIndex        =   132
               ToolTipText     =   "Quantidade."
               Top             =   960
               Width           =   1305
            End
            Begin VB.TextBox txtdescricao_Qualidade 
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
               MaxLength       =   255
               TabIndex        =   128
               ToolTipText     =   "Descrição."
               Top             =   945
               Width           =   6695
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Qtde."
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
               Index           =   102
               Left            =   14047
               TabIndex        =   360
               Top             =   750
               Width           =   420
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Un. est."
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
               Index           =   101
               Left            =   12030
               TabIndex        =   359
               Top             =   750
               Width           =   585
            End
            Begin VB.Label Label1 
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
               Index           =   100
               Left            =   9155
               TabIndex        =   358
               Top             =   750
               Width           =   480
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Un. com."
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
               Index           =   99
               Left            =   12855
               TabIndex        =   357
               Top             =   750
               Width           =   645
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Código de referência"
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
               Index           =   98
               Left            =   8902
               TabIndex        =   356
               Top             =   180
               Width           =   1500
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Cód. interno"
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
               Index           =   97
               Left            =   5782
               TabIndex        =   355
               Top             =   180
               Width           =   900
            End
            Begin VB.Label Label1 
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
               Index           =   96
               Left            =   3015
               TabIndex        =   354
               Top             =   180
               Width           =   915
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Análise crítica"
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
               Left            =   7058
               TabIndex        =   235
               Top             =   1350
               Width           =   975
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
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
               Index           =   34
               Left            =   3182
               TabIndex        =   234
               Top             =   750
               Width           =   690
            End
            Begin VB.Label Label1 
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
               Index           =   33
               Left            =   690
               TabIndex        =   191
               Top             =   180
               Width           =   345
            End
         End
         Begin MSComctlLib.ListView Lista_Qualidade 
            Height          =   4995
            Left            =   75
            TabIndex        =   134
            Top             =   2550
            Width           =   15105
            _ExtentX        =   26644
            _ExtentY        =   8811
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
            NumItems        =   8
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
               Object.Width           =   13961
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
               Alignment       =   1
               SubItemIndex    =   5
               Object.Tag             =   "N"
               Text            =   "Qtde."
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   6
               Object.Tag             =   "N"
               Text            =   "Valor unit."
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   7
               Object.Tag             =   "N"
               Text            =   "Valor total"
               Object.Width           =   2117
            EndProperty
         End
         Begin MSComctlLib.ListView Lista_check 
            Height          =   5235
            Index           =   2
            Left            =   -74925
            TabIndex        =   142
            Top             =   2310
            Width           =   15105
            _ExtentX        =   26644
            _ExtentY        =   9234
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
               Text            =   "Descrição"
               Object.Width           =   16783
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   4
               Object.Tag             =   "T"
               Text            =   "Executar"
               Object.Width           =   2117
            EndProperty
         End
      End
      Begin DrawSuite2022.USToolBar USToolBar7 
         Height          =   975
         Left            =   75
         TabIndex        =   272
         Top             =   330
         Visible         =   0   'False
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
         ButtonCaption5  =   "Copiar"
         ButtonEnabled5  =   0   'False
         ButtonIconSize5 =   32
         ButtonToolTipText5=   "Copiar fase(s) do processo."
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
         ButtonWidth5    =   44
         ButtonHeight5   =   21
         ButtonUseMaskColor5=   0   'False
         ButtonCaption6  =   "Utensílios"
         ButtonEnabled6  =   0   'False
         ButtonIconSize6 =   32
         ButtonToolTipText6=   "Abrir utensílios da fase."
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
         ButtonLeft6     =   241
         ButtonTop6      =   2
         ButtonWidth6    =   63
         ButtonHeight6   =   21
         ButtonUseMaskColor6=   0   'False
         ButtonCaption7  =   "Validação"
         ButtonEnabled7  =   0   'False
         ButtonIconSize7 =   32
         ButtonToolTipText7=   "Validar/Cancelar validação."
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
         ButtonLeft7     =   306
         ButtonTop7      =   2
         ButtonWidth7    =   53
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
         ButtonLeft8     =   361
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
         ButtonLeft9     =   365
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
         ButtonLeft10    =   408
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
         ButtonLeft11    =   440
         ButtonTop11     =   2
         ButtonWidth11   =   24
         ButtonHeight11  =   24
         ButtonUseMaskColor11=   0   'False
         Begin DrawSuite2022.USImageList USImageList7 
            Left            =   11010
            Top             =   300
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmVendas_analise.frx":26391
            Count           =   1
         End
      End
      Begin TabDlg.SSTab SSTab_compras 
         Height          =   8025
         Left            =   -74940
         TabIndex        =   290
         Top             =   2160
         Width           =   15255
         _ExtentX        =   26908
         _ExtentY        =   14155
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
         TabCaption(0)   =   "Valores"
         TabPicture(0)   =   "frmVendas_analise.frx":2C3E1
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label1(44)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label1(114)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label1(115)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Label1(116)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Label1(117)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Lista_Compras"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "Txt_total_outros"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "Txt_total_ferramentas"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "Txt_total_compras"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "Txt_total_terceiros"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "Txt_total_materiais"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "txtID_Compras"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "Frame22"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).ControlCount=   13
         TabCaption(1)   =   "Check-list"
         TabPicture(1)   =   "frmVendas_analise.frx":2C3FD
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Txt_ID_check(3)"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Frame_check(3)"
         Tab(1).Control(2)=   "Lista_check(3)"
         Tab(1).ControlCount=   3
         Begin VB.TextBox Txt_ID_check 
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
            Height          =   315
            Index           =   3
            Left            =   -71610
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   300
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   4830
            Visible         =   0   'False
            Width           =   675
         End
         Begin VB.Frame Frame_check 
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
            Height          =   1965
            Index           =   3
            Left            =   -74925
            TabIndex        =   296
            Top             =   330
            Width           =   15105
            Begin VB.CheckBox Chk_sim_chek 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Sim - Quais?"
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
               Left            =   900
               TabIndex        =   167
               Top             =   795
               Width           =   1185
            End
            Begin VB.CheckBox Chk_nao_chek 
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
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   3
               Left            =   240
               TabIndex        =   166
               Top             =   795
               Width           =   585
            End
            Begin VB.TextBox Txt_texto_check 
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
               Height          =   825
               Index           =   3
               Left            =   150
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   168
               TabStop         =   0   'False
               Top             =   1005
               Width           =   14745
            End
            Begin VB.TextBox Txt_responsavel_check 
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
               Index           =   3
               Left            =   1020
               Locked          =   -1  'True
               TabIndex        =   163
               TabStop         =   0   'False
               ToolTipText     =   "Responsável pelo cadastro."
               Top             =   390
               Width           =   3825
            End
            Begin VB.TextBox Txt_data_check 
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
               Index           =   3
               Left            =   150
               Locked          =   -1  'True
               MaxLength       =   25
               TabIndex        =   162
               TabStop         =   0   'False
               ToolTipText     =   "Data do cadastro."
               Top             =   390
               Width           =   855
            End
            Begin VB.Frame Frame18 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
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
               Height          =   405
               Left            =   5640
               TabIndex        =   298
               Top             =   4650
               Width           =   2985
            End
            Begin VB.CommandButton Cmd_localizar_desc_check 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Index           =   3
               Left            =   14580
               Picture         =   "frmVendas_analise.frx":2C419
               Style           =   1  'Graphical
               TabIndex        =   165
               ToolTipText     =   "Localizar descrição padrão do check-list."
               Top             =   390
               Width           =   315
            End
            Begin VB.TextBox Txt_ID_descricao_check 
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
               Height          =   315
               Index           =   3
               Left            =   4860
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   297
               TabStop         =   0   'False
               Text            =   "0"
               Top             =   390
               Visible         =   0   'False
               Width           =   675
            End
            Begin VB.TextBox Txt_descricao_chek 
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
               Index           =   3
               Left            =   4860
               Locked          =   -1  'True
               TabIndex        =   164
               TabStop         =   0   'False
               ToolTipText     =   "Descrição."
               Top             =   390
               Width           =   9705
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
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
               Index           =   113
               Left            =   9367
               TabIndex        =   372
               Top             =   180
               Width           =   690
            End
            Begin VB.Label Label1 
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
               Index           =   112
               Left            =   2475
               TabIndex        =   371
               Top             =   180
               Width           =   915
            End
            Begin VB.Label Label1 
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
               Index           =   31
               Left            =   405
               TabIndex        =   299
               Top             =   180
               Width           =   345
            End
         End
         Begin VB.Frame Frame22 
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
            Height          =   1455
            Left            =   75
            TabIndex        =   292
            Top             =   330
            Width           =   15105
            Begin VB.TextBox txtData_Compras 
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
               TabIndex        =   147
               TabStop         =   0   'False
               ToolTipText     =   "Data do cadastro."
               Top             =   390
               Width           =   1155
            End
            Begin VB.TextBox txtResponsavel_Compras 
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
               Locked          =   -1  'True
               TabIndex        =   148
               TabStop         =   0   'False
               ToolTipText     =   "Responsável pelo cadastro."
               Top             =   390
               Width           =   3375
            End
            Begin VB.TextBox txtTexto_Compras 
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
               TabIndex        =   153
               TabStop         =   0   'False
               ToolTipText     =   "Descrição."
               Top             =   1005
               Width           =   10395
            End
            Begin VB.TextBox txtValor_Compras 
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
               Left            =   12060
               MaxLength       =   20
               TabIndex        =   155
               Text            =   "0,0000"
               ToolTipText     =   "Valor unitário."
               Top             =   1005
               Width           =   1455
            End
            Begin VB.TextBox txtSetor_compras 
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
               Left            =   12690
               Locked          =   -1  'True
               MaxLength       =   255
               TabIndex        =   152
               TabStop         =   0   'False
               ToolTipText     =   "Setor."
               Top             =   390
               Width           =   2235
            End
            Begin VB.TextBox txtQtde_compras 
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
               Left            =   10590
               Locked          =   -1  'True
               MaxLength       =   255
               TabIndex        =   154
               TabStop         =   0   'False
               ToolTipText     =   "Quantidade."
               Top             =   1005
               Width           =   1455
            End
            Begin VB.TextBox Txtdesenho_compras 
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
               Left            =   4740
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   149
               TabStop         =   0   'False
               ToolTipText     =   "Código interno."
               Top             =   390
               Width           =   1665
            End
            Begin VB.TextBox Txt_referencia_compras 
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
               Left            =   6420
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   150
               TabStop         =   0   'False
               ToolTipText     =   "Código de referência."
               Top             =   390
               Width           =   2115
            End
            Begin VB.TextBox Txt_familia_compras 
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
               Left            =   8550
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   151
               TabStop         =   0   'False
               ToolTipText     =   "Família."
               Top             =   390
               Width           =   4125
            End
            Begin VB.TextBox Txt_valor_total 
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
               Left            =   13530
               Locked          =   -1  'True
               MaxLength       =   255
               TabIndex        =   156
               TabStop         =   0   'False
               ToolTipText     =   "Quantidade."
               Top             =   1005
               Width           =   1395
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Valor total"
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
               Index           =   111
               Left            =   13860
               TabIndex        =   370
               Top             =   810
               Width           =   735
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Valor unitário"
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
               Index           =   110
               Left            =   12315
               TabIndex        =   369
               Top             =   810
               Width           =   945
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Quantidade"
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
               Index           =   109
               Left            =   10897
               TabIndex        =   368
               Top             =   810
               Width           =   840
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
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
               Index           =   108
               Left            =   13612
               TabIndex        =   367
               Top             =   180
               Width           =   390
            End
            Begin VB.Label Label1 
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
               Index           =   107
               Left            =   10372
               TabIndex        =   366
               Top             =   180
               Width           =   480
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Código de referência"
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
               Index           =   106
               Left            =   6727
               TabIndex        =   365
               Top             =   180
               Width           =   1500
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Cód. interno"
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
               Index           =   105
               Left            =   5122
               TabIndex        =   364
               Top             =   180
               Width           =   900
            End
            Begin VB.Label Label1 
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
               Index           =   104
               Left            =   2580
               TabIndex        =   363
               Top             =   180
               Width           =   915
            End
            Begin VB.Label Label1 
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
               Index           =   37
               Left            =   5032
               TabIndex        =   294
               Top             =   810
               Width           =   690
            End
            Begin VB.Label Label1 
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
               Index           =   36
               Left            =   585
               TabIndex        =   293
               Top             =   180
               Width           =   345
            End
         End
         Begin VB.TextBox txtID_Compras 
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
            Height          =   315
            Left            =   2985
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   291
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   3240
            Visible         =   0   'False
            Width           =   675
         End
         Begin VB.TextBox Txt_total_materiais 
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
            Left            =   6915
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   157
            TabStop         =   0   'False
            ToolTipText     =   "Valor total de materiais."
            Top             =   7470
            Width           =   1620
         End
         Begin VB.TextBox Txt_total_terceiros 
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
            Left            =   8550
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   158
            TabStop         =   0   'False
            ToolTipText     =   "Valor total de terceiros."
            Top             =   7470
            Width           =   1620
         End
         Begin VB.TextBox Txt_total_compras 
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
            Left            =   13455
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   161
            TabStop         =   0   'False
            ToolTipText     =   "Valor total de compras."
            Top             =   7470
            Width           =   1620
         End
         Begin VB.TextBox Txt_total_ferramentas 
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
            Left            =   11820
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   160
            TabStop         =   0   'False
            ToolTipText     =   "Valor total de ferramentas."
            Top             =   7470
            Width           =   1620
         End
         Begin VB.TextBox Txt_total_outros 
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
            Left            =   10185
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   159
            TabStop         =   0   'False
            ToolTipText     =   "Valor total outros."
            Top             =   7470
            Width           =   1620
         End
         Begin MSComctlLib.ListView Lista_Compras 
            Height          =   5385
            Left            =   75
            TabIndex        =   146
            Top             =   1800
            Width           =   15105
            _ExtentX        =   26644
            _ExtentY        =   9499
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
            NumItems        =   8
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
               Object.Width           =   13961
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
               Alignment       =   1
               SubItemIndex    =   5
               Object.Tag             =   "N"
               Text            =   "Qtde."
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   6
               Object.Tag             =   "N"
               Text            =   "Valor unit."
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   7
               Object.Tag             =   "N"
               Text            =   "Valor total"
               Object.Width           =   2117
            EndProperty
         End
         Begin MSComctlLib.ListView Lista_check 
            Height          =   5235
            Index           =   3
            Left            =   -74925
            TabIndex        =   169
            Top             =   2310
            Width           =   15105
            _ExtentX        =   26644
            _ExtentY        =   9234
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
               Text            =   "Descrição"
               Object.Width           =   16783
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   4
               Object.Tag             =   "T"
               Text            =   "Executar"
               Object.Width           =   2117
            EndProperty
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "(=) Total compras"
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
            Index           =   117
            Left            =   13485
            TabIndex        =   376
            Top             =   7260
            Width           =   2460
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "(+) Total ferram."
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
            Index           =   116
            Left            =   11910
            TabIndex        =   375
            Top             =   7260
            Width           =   2340
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "(+) Total outros"
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
            Index           =   115
            Left            =   10290
            TabIndex        =   374
            Top             =   7260
            Width           =   2310
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "(+) Total terceiros"
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
            Index           =   114
            Left            =   8550
            TabIndex        =   373
            Top             =   7260
            Width           =   2520
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Total materiais"
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
            Index           =   44
            Left            =   7080
            TabIndex        =   295
            Top             =   7260
            Width           =   2190
            WordWrap        =   -1  'True
         End
      End
      Begin DrawSuite2022.USToolBar USToolBar1 
         Height          =   975
         Left            =   -74925
         TabIndex        =   206
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
         ButtonLeft4     =   130
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
         ButtonLeft5     =   192
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
         ButtonLeft6     =   249
         ButtonTop6      =   2
         ButtonWidth6    =   55
         ButtonHeight6   =   21
         ButtonUseMaskColor6=   0   'False
         ButtonCaption7  =   "Filtrar todos"
         ButtonEnabled7  =   0   'False
         ButtonIconSize7 =   32
         ButtonToolTipText7=   "Filtrar todos os registros."
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
         ButtonLeft7     =   306
         ButtonTop7      =   2
         ButtonWidth7    =   77
         ButtonHeight7   =   21
         ButtonUseMaskColor7=   0   'False
         ButtonCaption8  =   "Copiar"
         ButtonEnabled8  =   0   'False
         ButtonIconSize8 =   32
         ButtonToolTipText8=   "Copiar análise crítica."
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
         ButtonLeft8     =   385
         ButtonTop8      =   2
         ButtonWidth8    =   44
         ButtonHeight8   =   21
         ButtonUseMaskColor8=   0   'False
         ButtonCaption9  =   "Revisão"
         ButtonEnabled9  =   0   'False
         ButtonIconSize9 =   32
         ButtonToolTipText9=   "Revisar análise crítica."
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
         ButtonLeft9     =   431
         ButtonTop9      =   2
         ButtonWidth9    =   53
         ButtonHeight9   =   21
         ButtonUseMaskColor9=   0   'False
         ButtonCaption10 =   "Fechar"
         ButtonEnabled10 =   0   'False
         ButtonIconSize10=   32
         ButtonToolTipText10=   "Fechar análise crítica."
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
         ButtonLeft10    =   486
         ButtonTop10     =   2
         ButtonWidth10   =   46
         ButtonHeight10  =   21
         ButtonUseMaskColor10=   0   'False
         ButtonCaption11 =   "Produto/Processo"
         ButtonEnabled11 =   0   'False
         ButtonIconSize11=   32
         ButtonToolTipText11=   "Cadastrar produto e processo da análise crítica."
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
         ButtonLeft11    =   534
         ButtonTop11     =   2
         ButtonWidth11   =   110
         ButtonHeight11  =   21
         ButtonUseMaskColor11=   0   'False
         ButtonCaption12 =   "Prazos"
         ButtonEnabled12 =   0   'False
         ButtonIconSize12=   32
         ButtonToolTipText12=   "Definir prazos."
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
         ButtonLeft12    =   646
         ButtonTop12     =   2
         ButtonWidth12   =   46
         ButtonHeight12  =   21
         ButtonUseMaskColor12=   0   'False
         ButtonCaption13 =   "Atualizar"
         ButtonEnabled13 =   0   'False
         ButtonIconSize13=   32
         ButtonToolTipText13=   "Utilizado pelo administrador do sistema."
         ButtonKey13     =   "13"
         ButtonAlignment13=   2
         BeginProperty ButtonFont13 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft13    =   694
         ButtonTop13     =   2
         ButtonWidth13   =   59
         ButtonHeight13  =   21
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
         ButtonLeft14    =   755
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft15    =   759
         ButtonTop15     =   2
         ButtonWidth15   =   41
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft16    =   802
         ButtonTop16     =   2
         ButtonWidth16   =   30
         ButtonHeight16  =   21
         ButtonUseMaskColor16=   0   'False
         ButtonEnabled17 =   0   'False
         ButtonIconSize17=   32
         ButtonKey17     =   "17"
         ButtonAlignment17=   2
         BeginProperty ButtonFont17 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState17   =   5
         ButtonLeft17    =   834
         ButtonTop17     =   2
         ButtonWidth17   =   24
         ButtonHeight17  =   24
         Begin DrawSuite2022.USImageList USImageList1 
            Left            =   13500
            Top             =   270
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmVendas_analise.frx":2C51B
            Count           =   1
         End
      End
   End
End
Attribute VB_Name = "frmVendas_analise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Novo_Analise                  As Boolean 'OK
Dim Novo_analise1                 As Boolean 'OK
Dim Novo_analise2                 As Boolean 'OK
Dim Novo_analise3                 As Boolean 'OK
Dim Novo_analise4                 As Boolean 'OK
Dim Novo_analise5                 As Boolean 'OK
Dim Novo_analise6                 As Boolean 'OK
Dim Novo_analise7                 As Boolean 'OK
Dim Novo_analise8                 As Boolean 'OK
Dim Novo_analise9                 As Boolean 'OK
Dim Novo_analise10                As Boolean 'OK
Public StrSql_AnaliseCritica      As String  'OK
Dim TBLISTA_AnaliseCritica        As ADODB.Recordset 'OK
Dim Acesso                        As Boolean 'OK

Private Sub Chk_nao_chek_Click(index As Integer)
On Error GoTo tratar_erro

If Chk_nao_chek(index).Value = 1 Then
    Chk_sim_chek(index).Value = 0
    With Txt_texto_check(index)
        .Text = ""
        .Locked = True
        .TabStop = False
    End With
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_sim_chek_Click(index As Integer)
On Error GoTo tratar_erro

With Txt_texto_check(index)
    If Chk_sim_chek(index).Value = 1 Then
        Chk_nao_chek(index).Value = 0
        .Locked = False
        .TabStop = True
    Else
        .Text = ""
        .Locked = True
        .TabStop = False
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkAuto_engenharia_Click()
On Error GoTo tratar_erro

With txtdesenho_engenharia
    If chkAuto_engenharia.Value = 1 Then
        chkManual_engenharia.Value = 0
        ProcLiberaCamposEng
        .Locked = True
        .TabStop = False
    Else
        ProcBloqueiaCamposEng
        .Locked = False
        .TabStop = True
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkAuto_qualidade_Click()
On Error GoTo tratar_erro

With txtdesenho_qualidade
    If chkAuto_qualidade.Value = 1 Then
        chkManual_qualidade.Value = 0
        ProcLiberaCamposQualidade
        .Locked = True
        .TabStop = False
    Else
        ProcBloqueiaCamposQualidade
        .Locked = False
        .TabStop = True
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkManual_engenharia_Click()
On Error GoTo tratar_erro

With txtdesenho_engenharia
    If chkManual_engenharia.Value = 1 Then
        ProcLiberaCamposEng
        chkAuto_engenharia.Value = 0
        USMsgBox ("Informe o código interno do registro."), vbInformation, "CAPRIND v5.0"
        .Text = ""
        .SetFocus
    Else
        ProcBloqueiaCamposEng
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkManual_qualidade_Click()
On Error GoTo tratar_erro

With txtdesenho_qualidade
    If chkManual_qualidade.Value = 1 Then
        ProcLiberaCamposQualidade
        chkAuto_qualidade.Value = 0
        USMsgBox ("Informe o código interno do instrumento."), vbInformation, "CAPRIND v5.0"
        .Text = ""
        .SetFocus
    Else
        ProcBloqueiaCamposQualidade
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkPchora_Click()
On Error GoTo tratar_erro

With txtPecaHora_processos
    If chkPchora.Value = 1 Then
        .Locked = False
        .TabStop = True
    Else
        .Locked = True
        .TabStop = False
        .Text = 1
    End If
End With
ProcCalculaExecucao
ProcCalculamaquina

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcProduto_processo()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If txtId = "" Then
    USMsgBox ("Informe a análise crítica antes de criar o produto/processo."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Txt_status <> "APROVADA" Then
    USMsgBox ("Só é permitido criar o produto/processo de análise crítica com o status APROVADA."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If USMsgBox("Deseja realmente cadastrar o(s) produto(s) e o(s) processo(s) desta análise crítica " & Txt_analise & "?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    'Verifica se o produto já está cadastrado
    Set TBClientes = CreateObject("adodb.recordset")
    TBClientes.Open "Select * from Vendas_analise_ProdutosProcessos where id_analise = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
    If TBClientes.EOF = False Then
        Do While TBClientes.EOF = False
            Set TBCiclo = CreateObject("adodb.recordset")
            TBCiclo.Open "Select * from projproduto where Desenho = '" & TBClientes!Codinterno & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBCiclo.EOF = True Then
                If TBClientes!Produto_analise = True Then Permitido = True Else Permitido = False
                frmVendas_analise_cadastro.Show 1
                
                If Permitido1 = True Then
                    Set TBItem = CreateObject("adodb.recordset")
                    TBItem.Open "Select Codproduto, Desenho, descricaotecnica, classe, unidade from projproduto where codproduto = " & Codproduto, Conexao, adOpenKeyset, adLockOptimistic
                    If TBItem.EOF = False Then
                        Conexao.Execute "Update Vendas_analise_ProdutosProcessos Set codProduto = " & TBItem!Codproduto & ", Codinterno = '" & TBItem!Desenho & "', Descricao = '" & TBItem!descricaotecnica & "', Familia = '" & TBItem!Classe & "', un = '" & TBItem!Unidade & "' where id = " & TBClientes!ID
                    End If
                    TBItem.Close
                End If
            Else
                If USMsgBox("Já existe produto cadastrado com este código interno " & TBClientes!Codinterno & ", deseja prosseguir com o cadastro/alteração do processo assim mesmo?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
                Codproduto = TBClientes!Codproduto
                Permitido1 = True
            End If
            TBCiclo.Close
            If Permitido1 = True Then ProcCadastrarProcesso
            TBClientes.MoveNext
        Loop
        
        If Permitido1 = True And Codproduto <> "" And Codproduto <> "0" Then
            USMsgBox ("Produto(s) e processo(s) cadastrados com sucesso."), vbInformation, "CAPRIND v5.0"
            '==================================
            Modulo = "Outros/Análise crítica"
            Evento = "Cadastrar produto(s) e processo(s)"
            ID_documento = txtId.Text
            Documento = "Nº análise: " & Txt_analise & " - Rev.: " & Txt_rev_analise
            Documento1 = ""
            ProcGravaEvento
            '==================================
            
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select * from Vendas_analise where ID = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                ProcLimpaCampos
                ProcPuxaDados
            End If
            Lista.ListItems.Clear
            ProcCarregaLista (1)
            
            'Atualiza código interno na proposta
            Conexao.Execute "Update vendas_carteira Set Desenho = '" & txtdesenho & "' where IDAnalise = " & txtId
        End If
        
    Else
        USMsgBox ("Não é possivel cadastrar o(s) produto(s) e o(s) processo(s) desta análise crítica, pois o(s) processo(s) não está(ão) cadastrado(s)."), vbExclamation, "CAPRIND v5.0"
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCadastrarProcesso()
On Error GoTo tratar_erro

'Cadastra processo
If Codproduto = "" Then Exit Sub
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from Processos where codproduto = " & Codproduto, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then
    TBGravar.AddNew
    TBGravar!cronometrado = "NÃO"
    TBGravar!Bloqueado = False
    TBGravar!Nprocesso = FunCriarNovoNumeroProcesso
    TBGravar!Revisao = 0

    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select SubTipoItem, Codproduto, RevDesenho, Processo from projproduto where codproduto = " & Codproduto, Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
        Select Case TBProduto!SubTipoItem
            Case 1: TBGravar!Tipo = "E"
            Case 2: TBGravar!Tipo = "M"
            Case 3: TBGravar!Tipo = "F"
        End Select
        Codproduto = TBProduto!Codproduto
        TBGravar!Contador = TBProduto!RevDesenho
        TBGravar!Codproduto = TBProduto!Codproduto
    
        TBProduto!Processo = True
        TBProduto.Update
    End If
    
    TBGravar!DtImplantacao = Date
    TBGravar!elaborado = pubUsuario
    TBGravar.Update
End If
IDPROCESSO = TBGravar!IDPROCESSO
Conexao.Execute "Update Processos set ordenarprocesso = " & TBGravar!IDPROCESSO & " where IDprocesso = " & TBGravar!IDPROCESSO

'Cadatra as fases
TotalGeral = 0
Set TBFIltro = CreateObject("adodb.recordset")
TBFIltro.Open "Select * from Vendas_analise_setores where IDanalise = " & txtId & " and Setor = 'PROCESSOS' and ID_processo_item = " & TBClientes!ID & " order by Fase", Conexao, adOpenKeyset, adLockOptimistic
If TBFIltro.EOF = False Then
    Do While TBFIltro.EOF = False
        Set TBFases = CreateObject("adodb.recordset")
        TBFases.Open "Select * from Fases where IDprocesso = " & IDPROCESSO & " and Fase = " & TBFIltro!Fase, Conexao, adOpenKeyset, adLockOptimistic
        If TBFases.EOF = True Then
            TBFases.AddNew
            TBFases!Revisao = 0
            TBFases!versao = "A"
            TBFases!cronometrado = False
            TBFases!Plano_inspecao = False
        End If
        TBFases!IDPROCESSO = TBGravar!IDPROCESSO
        TBFases!Fase = TBFIltro!Fase
        TBFases!Grupo_op = TBFIltro!Grupo_op
        TBFases!maquina = TBFIltro!Texto
        TBFases!Descricao = TBFIltro!Trabalho

        Procsaida
        TBFases!Preparacao = Preparacao
        ElapsedTime (Preparacao)
        TBFases!TPSegundos = s
        TBFases!TempoPreparacao = TempoPreparacao

        TBFases!Execucao = Execucao
        TBFases!TESegundos = FunCalculaSegPC(TBFases!Execucao, TBFIltro!Peca)
        TBFases!TempoExecucao = FormataTempo(TBFases!TESegundos)
        
        TBFases!pc_te = TBFIltro!Peca
        TBFases!pecahora = TBFIltro!pecahora

        TotalGeral = TotalGeral + IIf(IsNull(TBFases!TESegundos), 0, TBFases!TESegundos)
        TBFases.Update
        
        'Cadatrar utensilios da fase
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select Codinterno, Qtde, ID_acessorio from Vendas_analise_setores where IDanalise = " & txtId & " and Fase = '" & TBFIltro!Fase & "' and Setor = 'FERRAMENTAS'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Do While TBAbrir.EOF = False
                Set TBFerramenta = CreateObject("adodb.recordset")
                TBFerramenta.Open "Select * from Ferramentas where IDprocesso = " & IDPROCESSO & " and IDFase = " & TBFases!IDFase, Conexao, adOpenKeyset, adLockOptimistic
                If TBFerramenta.EOF = True Then TBFerramenta.AddNew
                TBFerramenta!IDPROCESSO = IDPROCESSO
                TBFerramenta!IDFase = TBFases!IDFase
                TBFerramenta!Numero = TBAbrir!Codinterno
                TBFerramenta!quantidade = TBAbrir!Qtde
                TBFerramenta!ID_acessorio = TBAbrir!ID_acessorio
                TBFerramenta.Update
                TBFerramenta.Close
                TBAbrir.MoveNext
            Loop
        End If
        TBAbrir.Close
        TBFases.Close
        TBFIltro.MoveNext
    Loop
End If
TBFIltro.Close

Tempoprocesso = FormataTempo(TotalGeral)
ProcFormataHora (Tempoprocesso)
Qtd = s + DecimoSegundos
Data_Prog = FormataTempo(TBGravar!TTotalSEG)

Conexao.Execute "Update Processos Set TTotalSEG = " & Qtd & ", TTotal = '" & Data_Prog & "' where IDProcesso = " & IDPROCESSO

TBGravar.Close
ProcAtualizaCustoProcesso

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
If TBFIltro!Preparacao > "023:59:59" Then ProcFormataHora (TBFIltro!Preparacao) Else DataResultado = TBFIltro!Preparacao
ElapsedTime (DataResultado)
Preparacao = DataResultado
TempoPreparacao = HoraTotal
'=====================================================
' CALCULO DE TEMPO DE EXECUÇÃO MAIOR QUE 23:59:59 HS =
'=====================================================
DataResultado = 0
If TBFIltro!Execucao > "023:59:59" Then ProcFormataHora (TBFIltro!Execucao) Else DataResultado = TBFIltro!Execucao
ElapsedTime (DataResultado)
Execucao = DataResultado
TempoExecucao = HoraTotal

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
TBProcessos.Open "Select * from processos where idprocesso = " & IDPROCESSO, Conexao, adOpenKeyset, adLockOptimistic
If TBProcessos.EOF = False Then
    'Localiza fases do processo para atualizar custo
    Set TBFases = CreateObject("adodb.recordset")
    TBFases.Open "Select * FROM Fases where IDProcesso = " & IDPROCESSO & " order by fase", Conexao, adOpenKeyset, adLockOptimistic
    If Not (TBFases.BOF And TBFases.EOF) Then
        TBFases.MoveFirst
        Do Until TBFases.EOF
            'Busca custo hora da maquina
            Set TBMaquinas = CreateObject("adodb.recordset")
            TBMaquinas.Open "Select * FROM CadMaquinas where Maquina = '" & TBFases("maquina") & "'", Conexao, adOpenKeyset, adLockOptimistic
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
    TBProcessos.Update
End If
TBProcessos.Close
                
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFerramentas()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Acao = "abrir o módulo de ferramentas da fase"
If txtFase.Text = "" Then
    NomeCampo = "a fase"
    ProcVerificaAcao
    Exit Sub
End If
If Novo_analise2 = True Then
    USMsgBox ("Salve a fase antes de abrir o módulo de ferramentas da fase."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
frmVendas_analise_processos_ferramentas.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbFamilia_Click()
On Error GoTo tratar_erro

txtdescricao = FunBuscaDescPadraoFamilia(cmbfamilia, txtdesenho, txtdescricao)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbfamilia_engenharia_Click()
On Error GoTo tratar_erro

txtdescricao_engenharia = FunBuscaDescPadraoFamilia(cmbfamilia_engenharia, txtdesenho_engenharia, txtdescricao_engenharia)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbFamilia_processos_item_Click()
On Error GoTo tratar_erro

txtDescricao_processos_item = FunBuscaDescPadraoFamilia(cmbFamilia_processos_item, txtCodInterno_processos_item, txtDescricao_processos_item)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbfamilia_qualidade_Click()
On Error GoTo tratar_erro

txtdescricao_Qualidade = FunBuscaDescPadraoFamilia(cmbfamilia_qualidade, txtdesenho_qualidade, txtdescricao_Qualidade)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbLocal_cobranca_Click()
On Error GoTo tratar_erro

If cmbLocal_cobranca <> "" Then txtID_cobranca = cmbLocal_cobranca.ItemData(cmbLocal_cobranca.ListIndex) Else txtID_cobranca = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbLocal_entrega_Click()
On Error GoTo tratar_erro

If cmbLocal_entrega <> "" Then txtID_entrega = cmbLocal_entrega.ItemData(cmbLocal_entrega.ListIndex) Else txtID_entrega = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_filtrar_produto_Click()
On Error GoTo tratar_erro

ProcCarregaDadosProduto

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaDadosProduto()
On Error GoTo tratar_erro

Procliberacampos
If txtdesenho <> "" Then
    ProcLimpaCamposProd
    
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "SELECT * from projproduto where desenho = '" & txtdesenho & "' and Vendas = 'True' and Bloqueado = 'False'", Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
        txtidproduto = TBProduto!Codproduto
        txtRev_desenho = IIf(IsNull(TBProduto!RevDesenho), "", TBProduto!RevDesenho)
        ProcCarregaComboCodRef cmbReferencia, "P.codproduto = " & TBProduto!Codproduto, 0, "", False, False
        txtdescricao = IIf(IsNull(TBProduto!Descricao), "", TBProduto!Descricao)
        
        NomeCampo = "a família"
        If IsNull(TBProduto!Classe) = False And TBProduto!Classe <> "" Then cmbfamilia = TBProduto!Classe
        NomeCampo = "a unidade de estoque"
        If IsNull(TBProduto!Unidade) = False And TBProduto!Unidade <> "" Then cmbun = TBProduto!Unidade
        NomeCampo = "a unidade comercial"
        If IsNull(TBProduto!Unidade_com) = False And TBProduto!Unidade_com <> "" Then Cmb_un_com = TBProduto!Unidade_com
2:
        ProcBloqueiaCampos
        TBProduto.Close
    Else
        Procliberacampos
    End If
Else
    Procliberacampos
End If

Exit Sub
tratar_erro:
    If Err.Number = "383" Then
        USMsgBox ("Não foi encontrado " & NomeCampo & " deste produto."), vbExclamation, "CAPRIND v5.0"
        GoTo 2
    End If
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimpaCamposProd()
On Error GoTo tratar_erro

txtidproduto = ""
txtRev_desenho = ""
cmbReferencia.Clear
cmbfamilia.ListIndex = -1
cmbun.ListIndex = -1
Cmb_un_com.ListIndex = -1
txtdescricao = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub Procliberacampos()
On Error GoTo tratar_erro

With txtRev_desenho
    .Locked = False
    .TabStop = True
End With
With cmbfamilia
    .Locked = False
    .TabStop = True
End With
With cmbun
    .Locked = False
    .TabStop = True
End With
With Cmb_un_com
    .Locked = False
    .TabStop = True
End With
With txtdescricao
    .Locked = False
    .TabStop = True
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcBloqueiaCampos()
On Error GoTo tratar_erro

With txtRev_desenho
    .Locked = True
    .TabStop = False
End With
With cmbfamilia
    .Locked = True
    .TabStop = False
End With
With cmbun
    .Locked = True
    .TabStop = False
End With
With Cmb_un_com
    .Locked = True
    .TabStop = False
End With
With txtdescricao
    .Locked = True
    .TabStop = False
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_filtrar_produto_engenharia_Click()
On Error GoTo tratar_erro

ProcCarregaDadosProdutoEng

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaDadosProdutoEng()
On Error GoTo tratar_erro

ProcLiberaCamposEng
If chkAuto_engenharia = 1 Or chkManual_engenharia = 1 Then Exit Sub
If txtdesenho_engenharia <> "" Then
    ProcLimpaCamposProdEng
    
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "SELECT * from projproduto where desenho = '" & txtdesenho_engenharia & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
        txtIDproduto_engenharia = TBProduto!Codproduto
        If TBProduto!Tipo = "S" Then
            Opt_terceiros.Value = True
        Else
            If TBProduto!SubTipoItem = 0 Then Opt_material.Value = True Else Opt_outros.Value = True
        End If
        txtRev_desenho_engenharia = IIf(IsNull(TBProduto!RevDesenho), "", TBProduto!RevDesenho)
        ProcCarregaComboCodRef cmbReferencia_engenharia, "P.codproduto = " & TBProduto!Codproduto, 0, "", False, False
        txtdescricao_engenharia = IIf(IsNull(TBProduto!Descricao), "", TBProduto!Descricao)
        NomeCampo = "família"
        If IsNull(TBProduto!Classe) = False And TBProduto!Classe <> "" Then cmbfamilia_engenharia = TBProduto!Classe
        NomeCampo = "unidade"
        If IsNull(TBProduto!Unidade) = False And TBProduto!Unidade <> "" Then cmbun_engenharia = TBProduto!Unidade
        NomeCampo = "unidade comercial"
        If IsNull(TBProduto!Unidade_com) = False And TBProduto!Unidade_com <> "" Then Cmb_un_com_engenharia = TBProduto!Unidade
2:
        ProcBloqueiaCamposEng
        TBProduto.Close
    Else
        ProcLiberaCamposEng
    End If
Else
    ProcLiberaCamposEng
End If

Exit Sub
tratar_erro:
    If Err.Number = "383" Then
        USMsgBox ("Não foi encontrado a " & NomeCampo & " deste registro."), vbExclamation, "CAPRIND v5.0"
        GoTo 2
    End If
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimpaCamposProdEng()
On Error GoTo tratar_erro

txtIDproduto_engenharia = ""
txtdescricao_engenharia = ""
cmbfamilia_engenharia.ListIndex = -1
cmbun_engenharia.ListIndex = -1
Cmb_un_com_engenharia.ListIndex = -1
cmbReferencia_engenharia.Clear
txtReferencia_engenharia = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLiberaCamposEng()
On Error GoTo tratar_erro

If chkAuto_engenharia.Value = 1 Or chkManual_engenharia.Value = 1 Then
    cmbReferencia_engenharia.Visible = False
    txtReferencia_engenharia.Visible = True
End If
With txtdescricao_engenharia
    .Locked = False
    .TabStop = True
End With
With cmbfamilia_engenharia
    .Locked = False
    .TabStop = True
End With
With cmbun_engenharia
    .Locked = False
    .TabStop = True
End With
With Cmb_un_com_engenharia
    .Locked = False
    .TabStop = True
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLiberaCamposProcessos()
On Error GoTo tratar_erro

With txtDescricao_processos_item
    .Locked = False
    .TabStop = True
End With
With cmbFamilia_processos_item
    .Locked = False
    .TabStop = True
End With
With cmbUn_processos_item
    .Locked = False
    .TabStop = True
End With
With Cmb_un_com_processos_item
    .Locked = False
    .TabStop = True
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLiberaCamposQualidade()
On Error GoTo tratar_erro

With txtdescricao_Qualidade
    .Locked = False
    .TabStop = True
End With
With cmbfamilia_qualidade
    .Locked = False
    .TabStop = True
End With
With cmbun_qualidade
    .Locked = False
    .TabStop = True
End With
With Cmb_un_com_qualidade
    .Locked = False
    .TabStop = True
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcBloqueiaCamposEng()
On Error GoTo tratar_erro

cmbReferencia_engenharia.Visible = True
txtReferencia_engenharia.Visible = False
With txtdescricao_engenharia
    .Locked = True
    .TabStop = False
End With
With cmbfamilia_engenharia
    .Locked = True
    .TabStop = False
End With
With cmbun_engenharia
    .Locked = True
    .TabStop = False
End With
With Cmb_un_com_engenharia
    .Locked = True
    .TabStop = False
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcBloqueiaCamposProcessos()
On Error GoTo tratar_erro

With txtDescricao_processos_item
    .Locked = True
    .TabStop = False
End With
With cmbFamilia_processos_item
    .Locked = True
    .TabStop = False
End With
With cmbUn_processos_item
    .Locked = True
    .TabStop = False
End With
With Cmb_un_com_processos_item
    .Locked = True
    .TabStop = False
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcBloqueiaCamposQualidade()
On Error GoTo tratar_erro

With txtdescricao_Qualidade
    .Locked = True
    .TabStop = False
End With
With cmbfamilia
    .Locked = True
    .TabStop = False
End With
With cmbun_qualidade
    .Locked = True
    .TabStop = False
End With
With Cmb_un_com_qualidade
    .Locked = True
    .TabStop = False
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_filtrar_produto_qualidade_Click()
On Error GoTo tratar_erro

ProcCarregaDadosProdutoQualidade

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaDadosProdutoQualidade()
On Error GoTo tratar_erro

ProcLiberaCamposQualidade
If chkAuto_qualidade = 1 Or chkManual_qualidade = 1 Then Exit Sub
If txtdesenho_qualidade <> "" Then
    ProcLimpaCamposProdQualidade
    
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "SELECT * from projproduto where desenho = '" & txtdesenho_qualidade & "' and Instrumento = 'True'", Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
        txtIDproduto_qualidade = TBProduto!Codproduto
        txtRev_desenho_qualidade = IIf(IsNull(TBProduto!RevDesenho), "", TBProduto!RevDesenho)
        txtdescricao_Qualidade = IIf(IsNull(TBProduto!Descricao), "", TBProduto!Descricao)
        ProcCarregaComboCodRef cmbReferencia_qualidade, "P.codproduto = " & TBProduto!Codproduto, 0, "", False, False
        NomeCampo = "família"
        If IsNull(TBProduto!Classe) = False And TBProduto!Classe <> "" Then cmbfamilia_qualidade = TBProduto!Classe
        NomeCampo = "unidade"
        If IsNull(TBProduto!Unidade) = False And TBProduto!Unidade <> "" Then cmbun_qualidade = TBProduto!Unidade
        NomeCampo = "unidade comercial"
        If IsNull(TBProduto!Unidade_com) = False And TBProduto!Unidade_com <> "" Then Cmb_un_com_qualidade = TBProduto!Unidade_com
2:
        ProcBloqueiaCamposQualidade
        TBProduto.Close
    Else
        ProcLiberaCamposQualidade
    End If
Else
    ProcLiberaCamposQualidade
End If

Exit Sub
tratar_erro:
    If Err.Number = "383" Then
        USMsgBox ("Não foi encontrado a " & NomeCampo & " deste instrumento."), vbExclamation, "CAPRIND v5.0"
        GoTo 2
    End If
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimpaCamposProdQualidade()
On Error GoTo tratar_erro

txtIDproduto_qualidade = ""
cmbReferencia_qualidade.Clear
txtdescricao_Qualidade = ""
cmbfamilia_qualidade.ListIndex = -1
cmbun_qualidade.ListIndex = -1
Cmb_un_com_qualidade.ListIndex = -1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaDadosProdutoProcessos()
On Error GoTo tratar_erro

If optProcessos_item_analise.Value = True Then Exit Sub
ProcLiberaCamposProcessos
If txtCodInterno_processos_item <> "" Then
    ProcLimpaCamposProdProcessos
    
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "SELECT * from projproduto where desenho = '" & txtCodInterno_processos_item & "' and (SubTipoItem = 2 or SubTipoItem = 3)", Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
        txtIDproduto_processos = TBProduto!Codproduto
        txtDescricao_processos_item = IIf(IsNull(TBProduto!Descricao), "", TBProduto!Descricao)
        ProcCarregaComboCodRef cmbReferencia_processos_item, "P.codproduto = " & TBProduto!Codproduto, 0, "", False, False
        NomeCampo = "família"
        If IsNull(TBProduto!Classe) = False And TBProduto!Classe <> "" Then cmbFamilia_processos_item = TBProduto!Classe
        NomeCampo = "unidade"
        If IsNull(TBProduto!Unidade) = False And TBProduto!Unidade <> "" Then cmbUn_processos_item = TBProduto!Unidade
        NomeCampo = "unidade comercial"
        If IsNull(TBProduto!Unidade_com) = False And TBProduto!Unidade_com <> "" Then Cmb_un_com_processos_item = TBProduto!Unidade_com
2:
        ProcBloqueiaCamposProcessos
        TBProduto.Close
    Else
        ProcLiberaCamposProcessos
    End If
Else
    ProcLiberaCamposProcessos
End If

Exit Sub
tratar_erro:
    If Err.Number = "383" Then
        USMsgBox ("Não foi encontrado a " & NomeCampo & " deste produto."), vbExclamation, "CAPRIND v5.0"
        GoTo 2
    End If
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimpaCamposProdProcessos()
On Error GoTo tratar_erro

txtIDproduto_processos = ""
cmbReferencia_processos_item.Clear
txtDescricao_processos_item = ""
cmbFamilia_processos_item.ListIndex = -1
cmbUn_processos_item.ListIndex = -1
Cmb_un_com_processos_item.ListIndex = -1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImpostos()
On Error GoTo tratar_erro

Formulario = "Outros/Análise crítica/Vendas"
Direitos
ProcVerificaAcessos
Acao = "fazer o fechamento"
If txtId = "" Then
    NomeCampo = "a análise crítica"
    ProcVerificaAcao
    ProcAbrir
    Exit Sub
End If
If FunVerificaProsseguir = True Then
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "select ID from Vendas_analise where ID = " & txtId & " and Fechada = 'False' and (DtValidacao_Engenharia IS NULL or DtValidacao_Processo IS NULL or DtValidacao_PCP IS NULL or DtValidacao_Qualidade IS NULL or DtValidacao_Compras IS NULL)", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        USMsgBox ("Não é permitido fechar esta análise crítica, pois não foram validados todos os setores."), vbExclamation, "CAPRIND v5.0"
        TBAbrir.Close
        Exit Sub
    End If
    frmVendas_analise_impostos.Show 1
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

Private Sub Cmd_localizar_desc_check_Click(index As Integer)
On Error GoTo tratar_erro

ProcAbrirDescCheckList index

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAbrirDescCheckList(index As Integer)
On Error GoTo tratar_erro

Sit_REG = index
frmVendas_analise_descricao_checklist.Show 1

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

Private Sub Cmd_status_Click()
On Error GoTo tratar_erro

If txtId = "" Then Exit Sub
If Txt_status = "REVISADA" Then Exit Sub
frmVendas_analise_status.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_tipo_Click()
On Error GoTo tratar_erro

frmVendas_analise_tipo.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAnterior()
On Error GoTo tratar_erro

If txtId = 0 Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from vendas_analise order by id", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.BOF = False Then
    TBLISTA.Find ("id = " & txtId)
    TBLISTA.MovePrevious
    If TBLISTA.BOF = False Then
        ProcLimpaCampos
        ProcLimparTudo
        txtId = TBLISTA!ID
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from vendas_analise where id  = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            ProcPuxaDados
        End If
        TBAbrir.Close
        ProcCarregaLista_Engenharia_Prod
        ProcCarregaLista_Engenharia_Checklist
        ProcCarregaLista_Engenharia_Normas
        ProcCarregaLista_processos_item
        ProcCarregaLista_PCP_Checklist
        ProcCarregaLista_Instrumentos
        ProcCarregaLista_Qualidade_Checklist
        ProcCarregalista_Compras
        ProcCarregaLista_Compras_Checklist
        ProcCarregaLista_Doc
    Else
        USMsgBox ("Fim dos cadastros da análise."), vbInformation, "CAPRIND v5.0"
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procFiltrarTodos()
On Error GoTo tratar_erro

StrSql_AnaliseCritica = "Select * from vendas_analise order by ordenaranalise desc, ID desc"
Lista.ListItems.Clear
ProcCarregaLista (1)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdContato_Click()
On Error GoTo tratar_erro

If txtIDcliente.Text <> "" And txtIDcliente.Text <> "0" Then
    Analise_critica = True
    Vendas_Proposta = False
    Vendas_PI = False
    Telemarketing = False
    Qualidade_PPAP_PSW = False
    Financeiro_Contas_Pagar = False
    Financeiro_Contas_Pagas = False
    Financeiro_Contas_Receber = False
    Financeiro_Contas_Recebidas = False
    frmVendas_propostaII_contato.Show 1
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdLocalcobranca_Click()
On Error GoTo tratar_erro

With cmbLocal_cobranca
    .Clear
    Set TBEstoque = CreateObject("adodb.recordset")
    TBEstoque.Open "Select * from clientes_cobranca where idcliente = " & txtIDcliente.Text & " and Tipo = 'C'", Conexao, adOpenKeyset, adLockOptimistic
    If TBEstoque.EOF = False Then
        Do While TBEstoque.EOF = False
            If IsNull(TBEstoque!Tipo_endereco) = False And TBEstoque!Tipo_endereco <> "" Then
                Endereco = TBEstoque!Tipo_endereco & ": " & IIf(IsNull(TBEstoque!endereco_Cobranca), "", TBEstoque!endereco_Cobranca)
            Else
                Endereco = IIf(IsNull(TBEstoque!endereco_Cobranca), "", TBEstoque!endereco_Cobranca)
            End If
            If IsNull(TBEstoque!Tipo_bairro) = False And TBEstoque!Tipo_bairro <> "" Then
                Bairro = TBEstoque!Tipo_bairro & ": " & IIf(IsNull(TBEstoque!bairro_Cobranca), "", TBEstoque!bairro_Cobranca)
            Else
                Bairro = IIf(IsNull(TBEstoque!bairro_Cobranca), "", TBEstoque!bairro_Cobranca)
            End If
            Endereco1 = Endereco & " - " & IIf(IsNull(TBEstoque!Numero), "", TBEstoque!Numero) & " - " & Bairro & " - " & IIf(IsNull(TBEstoque!cidade_Cobranca), "", TBEstoque!cidade_Cobranca) & " - " & IIf(IsNull(TBEstoque!uf_Cobranca), "", TBEstoque!uf_Cobranca) & " - " & IIf(IsNull(TBEstoque!cep_Cobranca), "", TBEstoque!cep_Cobranca)
            ID_Cobranca = TBEstoque!idCobranca
            
            .AddItem Endereco1
            .ItemData(.NewIndex) = ID_Cobranca
            TBEstoque.MoveNext
        Loop
        .Text = Endereco1
        txtID_cobranca = ID_Cobranca
    End If
    TBEstoque.Close
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdLocalentrega_Click()
On Error GoTo tratar_erro

With cmbLocal_entrega
    .Clear
    Set TBEstoque = CreateObject("adodb.recordset")
    TBEstoque.Open "Select * from clientes_entrega where idcliente = " & txtIDcliente.Text & " and Tipo = 'C'", Conexao, adOpenKeyset, adLockOptimistic
    If TBEstoque.EOF = False Then
        Do While TBEstoque.EOF = False
            If IsNull(TBEstoque!Tipo_endereco) = False And TBEstoque!Tipo_endereco <> "" Then
                Endereco = TBEstoque!Tipo_endereco & ": " & IIf(IsNull(TBEstoque!endereco_entrega), "", TBEstoque!endereco_entrega)
            Else
                Endereco = IIf(IsNull(TBEstoque!endereco_entrega), "", TBEstoque!endereco_entrega)
            End If
            If IsNull(TBEstoque!Tipo_bairro) = False And TBEstoque!Tipo_bairro <> "" Then
                Bairro = TBEstoque!Tipo_bairro & ": " & IIf(IsNull(TBEstoque!bairro_entrega), "", TBEstoque!bairro_entrega)
            Else
                Bairro = IIf(IsNull(TBEstoque!bairro_entrega), "", TBEstoque!bairro_entrega)
            End If
            Endereco1 = Endereco & " - " & IIf(IsNull(TBEstoque!Numero), "", TBEstoque!Numero) & " - " & Bairro & " - " & IIf(IsNull(TBEstoque!cidade_entrega), "", TBEstoque!cidade_entrega) & " - " & IIf(IsNull(TBEstoque!uf_entrega), "", TBEstoque!uf_entrega) & " - " & IIf(IsNull(TBEstoque!cep_entrega), "", TBEstoque!cep_entrega)
            ID_entrega = TBEstoque!identrega
            
            .AddItem Endereco1
            .ItemData(.NewIndex) = ID_entrega
            TBEstoque.MoveNext
        Loop
        .Text = Endereco1
        txtID_entrega = ID_entrega
    End If
    TBEstoque.Close
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_AnaliseCritica.AbsolutePage <> 2 Then
    If TBLISTA_AnaliseCritica.AbsolutePage = -3 Then
        ProcExibePagina (TBLISTA_AnaliseCritica.PageCount - 1)
    Else
        TBLISTA_AnaliseCritica.AbsolutePage = TBLISTA_AnaliseCritica.AbsolutePage - 2
        ProcExibePagina (TBLISTA_AnaliseCritica.AbsolutePage)
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
    TBLISTA_AnaliseCritica.AbsolutePage = txtPagIr.Text
    ProcExibePagina (TBLISTA_AnaliseCritica.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_AnaliseCritica.AbsolutePage = 1
ProcExibePagina (TBLISTA_AnaliseCritica.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_AnaliseCritica.AbsolutePage <> -3 Then
    If TBLISTA_AnaliseCritica.AbsolutePage = 1 Then
        ProcExibePagina (2)
    Else
        ProcExibePagina (TBLISTA_AnaliseCritica.AbsolutePage)
    End If
Else
    ProcExibePagina (TBLISTA_AnaliseCritica.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_AnaliseCritica.AbsolutePage = TBLISTA_AnaliseCritica.PageCount
ProcExibePagina (TBLISTA_AnaliseCritica.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdProduto_processos_Click()
On Error GoTo tratar_erro

If optProcessos_item_analise.Value = True Then Exit Sub
Sit_REG = 4
frmVendas_analise_item.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdcliente_Click()
On Error GoTo tratar_erro

ProcConfVariaveisLocCliente False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, True, False, False, False, False, False, False
frmVendas_LocalizarCliente.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCopiar()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If txtId = "" Then
    USMsgBox ("Informe a análise crítica antes de copiar."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If USMsgBox("Deseja realmente copiar a análise crítica " & Txt_analise & "?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    Revisar = False
    Set TBCotacao = CreateObject("adodb.recordset")
    TBCotacao.Open "Select Nanalise from Vendas_analise where Year(data) = '" & Year(Date) & "' order by Ordenaranalise desc", Conexao, adOpenKeyset, adLockOptimistic
    If TBCotacao.EOF = False Then Cotacao = Left(TBCotacao!Nanalise, Len(TBCotacao!Nanalise) - 3) + 1 Else Cotacao = 1
    Ano = Right(Year(Date), 2)
    Select Case Len(Cotacao)
        Case 1: NumeroAnalise = "000" & Cotacao & "/" & Ano
        Case 2: NumeroAnalise = "00" & Cotacao & "/" & Ano
        Case 3: NumeroAnalise = "0" & Cotacao & "/" & Ano
        Case 4: NumeroAnalise = Cotacao & "/" & Ano
        Case 5: NumeroAnalise = Cotacao & "/" & Ano
    End Select
    Txt_analise = NumeroAnalise
    
    ProcCopiarRevisar
    '==================================
    Modulo = "Outros/Análise crítica"
    Evento = "Novo"
    ID_documento = txtId.Text
    Documento = "Nº análise: " & Txt_analise & " - Rev.: " & Txt_rev_analise
    Documento1 = ""
    ProcGravaEvento
    '==================================
    USMsgBox ("Análise crítica copiada com sucesso."), vbInformation, "CAPRIND v5.0"
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from vendas_analise where id = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        ProcLimpaCampos
        ProcLimpaCampos_Engenharia_Prod
        ProcLimpaCampos_Engenharia_Checklist
        ProcLimpaCampos_Processos
        ProcLimpaCampos_PCP
        ProcLimpaCampos_Instrumentos
        ProcLimpaCampos_Compras
        ProcPuxaDados
        ProcCarregaLista_Engenharia_Prod
        ProcCarregaLista_Engenharia_Checklist
        ProcCarregaLista_Engenharia_Normas
        ProcCarregaLista_processos_item
        ProcCarregaLista_processos
        ProcCarregaLista_PCP_Checklist
        ProcCarregaLista_Instrumentos
        ProcCarregaLista_Qualidade_Checklist
        ProcCarregalista_Compras
    End If
    Frame1.Enabled = True
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCopiarRevisar()
On Error GoTo tratar_erro

Set TBFI = CreateObject("adodb.recordset")
TBFI.Open "Select * from Vendas_analise where ID = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
If TBFI.EOF = False Then
    Set TBGravar = CreateObject("adodb.recordset")
    TBGravar.Open "Select * from Vendas_analise", Conexao, adOpenKeyset, adLockOptimistic
    TBGravar.AddNew
    IDAntigo = txtId
    TBGravar!Nanalise = Txt_analise
    If Revisar = False Then
        TBGravar!Revisao = 0
    Else
        TBFI!status = "REVISADA"
        TBFI!Data_status = Date
        TBFI.Update
        
        TBGravar!Ordenaranalise = TBFI!Ordenaranalise
        TBGravar!Revisao = IIf(IsNull(TBFI!Revisao), 0, TBFI!Revisao) + 1
    End If
    TBGravar!Data = Date
    TBGravar!Responsavel = pubUsuario
    TBGravar!status = "ABERTA EM ANALISE"
    TBGravar!Data_status = Null
    TBGravar!IDProduto = TBFI!IDProduto
    TBGravar!Codinterno = TBFI!Codinterno
    TBGravar!RevDesenho = TBFI!RevDesenho
    TBGravar!N_referencia = TBFI!N_referencia
    TBGravar!Unidade = TBFI!Unidade
    TBGravar!Unidade_com = TBFI!Unidade_com
    TBGravar!qtde_solicitada = TBFI!qtde_solicitada
    TBGravar!Qtde = TBFI!Qtde
    TBGravar!Descricao = TBFI!Descricao
    TBGravar!Familia = TBFI!Familia
    TBGravar!Tipo = TBFI!Tipo
    TBGravar!IDCliente = TBFI!IDCliente
    TBGravar!Cliente = TBFI!Cliente
    TBGravar!Obs = TBFI!Obs
    TBGravar!NRef = TBFI!NRef
    TBGravar!Referencia = TBFI!Referencia
    TBGravar!contato = TBFI!contato
    TBGravar!Departamento = TBFI!Departamento
    TBGravar!telefone = TBFI!telefone
    TBGravar!Fax = TBFI!Fax
    TBGravar!Email = TBFI!Email
    TBGravar!ID_entrega = TBFI!ID_entrega
    TBGravar!Local_entrega = TBFI!Local_entrega
    TBGravar!ID_Cobranca = TBFI!ID_Cobranca
    TBGravar!Local_Cobranca = TBFI!Local_Cobranca
    
    TBGravar!Data_Engenharia = TBFI!Data_Engenharia
    TBGravar!Responsavel_Engenharia = TBFI!Responsavel_Engenharia
    
    TBGravar!data_PCP = TBFI!data_PCP
    TBGravar!responsavel_PCP = TBFI!responsavel_PCP
    TBGravar!Analise_PCP = TBFI!Analise_PCP
    
    TBGravar!Opt_ICMS_processo = TBFI!Opt_ICMS_processo
    TBGravar!Opt_PIS_processo = TBFI!Opt_PIS_processo
    TBGravar!Opt_cofins_processo = TBFI!Opt_cofins_processo
    TBGravar!Opt_CSLL_processo = TBFI!Opt_CSLL_processo
    TBGravar!Opt_ISSQN_processo = TBFI!Opt_ISSQN_processo
    TBGravar!Opt_IRPJ_processo = TBFI!Opt_IRPJ_processo
    TBGravar!Opt_simples_processo = TBFI!Opt_simples_processo
    TBGravar!Opt_comissao_processo = TBFI!Opt_comissao_processo
    TBGravar!Opt_despesas_financeiras_processo = TBFI!Opt_despesas_financeiras_processo
    TBGravar!Opt_frete_processo = TBFI!Opt_frete_processo
    TBGravar!Opt_margem_processo = TBFI!Opt_margem_processo
    TBGravar!Opt_despesas_comerciais_processo = TBFI!Opt_despesas_comerciais_processo
    TBGravar!Opt_despesas_administrativas_processo = TBFI!Opt_despesas_administrativas_processo
    TBGravar!ICMS_processo = TBFI!ICMS_processo
    TBGravar!PIS_processo = TBFI!PIS_processo
    TBGravar!Cofins_processo = TBFI!Cofins_processo
    TBGravar!CSLL_processo = TBFI!CSLL_processo
    TBGravar!ISSQN_processo = TBFI!ISSQN_processo
    TBGravar!IRPJ_processo = TBFI!IRPJ_processo
    TBGravar!Simples_processo = TBFI!Simples_processo
    TBGravar!Comissao_processo = TBFI!Comissao_processo
    TBGravar!Despesas_financeiras_processo = TBFI!Despesas_financeiras_processo
    TBGravar!Frete_processo = TBFI!Frete_processo
    TBGravar!Margem_processo = TBFI!Margem_processo
    TBGravar!Despesas_comerciais_processo = TBFI!Despesas_comerciais_processo
    TBGravar!Despesas_administrativas_processo = TBFI!Despesas_administrativas_processo
    
    TBGravar!Opt_ICMS_materiais = TBFI!Opt_ICMS_materiais
    TBGravar!Opt_PIS_materiais = TBFI!Opt_PIS_materiais
    TBGravar!Opt_cofins_materiais = TBFI!Opt_cofins_materiais
    TBGravar!Opt_CSLL_materiais = TBFI!Opt_CSLL_materiais
    TBGravar!Opt_ISSQN_materiais = TBFI!Opt_ISSQN_materiais
    TBGravar!Opt_IRPJ_materiais = TBFI!Opt_IRPJ_materiais
    TBGravar!Opt_simples_materiais = TBFI!Opt_simples_materiais
    TBGravar!Opt_comissao_materiais = TBFI!Opt_comissao_materiais
    TBGravar!Opt_despesas_financeiras_materiais = TBFI!Opt_despesas_financeiras_materiais
    TBGravar!Opt_frete_materiais = TBFI!Opt_frete_materiais
    TBGravar!Opt_margem_materiais = TBFI!Opt_margem_materiais
    TBGravar!Opt_despesas_comerciais_materiais = TBFI!Opt_despesas_comerciais_materiais
    TBGravar!Opt_despesas_administrativas_materiais = TBFI!Opt_despesas_administrativas_materiais
    TBGravar!ICMS_materiais = TBFI!ICMS_materiais
    TBGravar!PIS_materiais = TBFI!PIS_materiais
    TBGravar!Cofins_materiais = TBFI!Cofins_materiais
    TBGravar!CSLL_materiais = TBFI!CSLL_materiais
    TBGravar!ISSQN_materiais = TBFI!ISSQN_materiais
    TBGravar!IRPJ_materiais = TBFI!IRPJ_materiais
    TBGravar!Simples_materiais = TBFI!Simples_materiais
    TBGravar!Comissao_materiais = TBFI!Comissao_materiais
    TBGravar!Despesas_financeiras_materiais = TBFI!Despesas_financeiras_materiais
    TBGravar!Frete_materiais = TBFI!Frete_materiais
    TBGravar!Margem_materiais = TBFI!Margem_materiais
    TBGravar!Despesas_comerciais_materiais = TBFI!Despesas_comerciais_materiais
    TBGravar!Despesas_administrativas_materiais = TBFI!Despesas_administrativas_materiais
        
    TBGravar!Opt_ICMS_terceiros = TBFI!Opt_ICMS_terceiros
    TBGravar!Opt_PIS_terceiros = TBFI!Opt_PIS_terceiros
    TBGravar!Opt_cofins_terceiros = TBFI!Opt_cofins_terceiros
    TBGravar!Opt_CSLL_terceiros = TBFI!Opt_CSLL_terceiros
    TBGravar!Opt_ISSQN_terceiros = TBFI!Opt_ISSQN_terceiros
    TBGravar!Opt_IRPJ_terceiros = TBFI!Opt_IRPJ_terceiros
    TBGravar!Opt_simples_terceiros = TBFI!Opt_simples_terceiros
    TBGravar!Opt_comissao_terceiros = TBFI!Opt_comissao_terceiros
    TBGravar!Opt_despesas_financeiras_terceiros = TBFI!Opt_despesas_financeiras_terceiros
    TBGravar!Opt_frete_terceiros = TBFI!Opt_frete_terceiros
    TBGravar!Opt_margem_terceiros = TBFI!Opt_margem_terceiros
    TBGravar!Opt_despesas_comerciais_terceiros = TBFI!Opt_despesas_comerciais_terceiros
    TBGravar!Opt_despesas_administrativas_terceiros = TBFI!Opt_despesas_administrativas_terceiros
    TBGravar!ICMS_terceiros = TBFI!ICMS_terceiros
    TBGravar!PIS_terceiros = TBFI!PIS_terceiros
    TBGravar!Cofins_terceiros = TBFI!Cofins_terceiros
    TBGravar!CSLL_terceiros = TBFI!CSLL_terceiros
    TBGravar!ISSQN_terceiros = TBFI!ISSQN_terceiros
    TBGravar!IRPJ_terceiros = TBFI!IRPJ_terceiros
    TBGravar!Simples_terceiros = TBFI!Simples_terceiros
    TBGravar!Comissao_terceiros = TBFI!Comissao_terceiros
    TBGravar!Despesas_financeiras_terceiros = TBFI!Despesas_financeiras_terceiros
    TBGravar!Frete_terceiros = TBFI!Frete_terceiros
    TBGravar!Margem_terceiros = TBFI!Margem_terceiros
    TBGravar!Despesas_comerciais_terceiros = TBFI!Despesas_comerciais_terceiros
    TBGravar!Despesas_administrativas_terceiros = TBFI!Despesas_administrativas_terceiros
        
    TBGravar!Opt_ICMS_outros = TBFI!Opt_ICMS_outros
    TBGravar!Opt_PIS_outros = TBFI!Opt_PIS_outros
    TBGravar!Opt_cofins_outros = TBFI!Opt_cofins_outros
    TBGravar!Opt_CSLL_outros = TBFI!Opt_CSLL_outros
    TBGravar!Opt_ISSQN_outros = TBFI!Opt_ISSQN_outros
    TBGravar!Opt_IRPJ_outros = TBFI!Opt_IRPJ_outros
    TBGravar!Opt_simples_outros = TBFI!Opt_simples_outros
    TBGravar!Opt_comissao_outros = TBFI!Opt_comissao_outros
    TBGravar!Opt_despesas_financeiras_outros = TBFI!Opt_despesas_financeiras_outros
    TBGravar!Opt_frete_outros = TBFI!Opt_frete_outros
    TBGravar!Opt_margem_outros = TBFI!Opt_margem_outros
    TBGravar!Opt_despesas_comerciais_outros = TBFI!Opt_despesas_comerciais_outros
    TBGravar!Opt_despesas_administrativas_outros = TBFI!Opt_despesas_administrativas_outros
    TBGravar!ICMS_outros = TBFI!ICMS_outros
    TBGravar!PIS_outros = TBFI!PIS_outros
    TBGravar!Cofins_outros = TBFI!Cofins_outros
    TBGravar!CSLL_outros = TBFI!CSLL_outros
    TBGravar!ISSQN_outros = TBFI!ISSQN_outros
    TBGravar!IRPJ_outros = TBFI!IRPJ_outros
    TBGravar!Simples_outros = TBFI!Simples_outros
    TBGravar!Comissao_outros = TBFI!Comissao_outros
    TBGravar!Despesas_financeiras_outros = TBFI!Despesas_financeiras_outros
    TBGravar!Frete_outros = TBFI!Frete_outros
    TBGravar!Margem_outros = TBFI!Margem_outros
    TBGravar!Despesas_comerciais_outros = TBFI!Despesas_comerciais_outros
    TBGravar!Despesas_administrativas_outros = TBFI!Despesas_administrativas_outros
        
    TBGravar!Opt_ICMS_total = TBFI!Opt_ICMS_total
    TBGravar!Opt_PIS_total = TBFI!Opt_PIS_total
    TBGravar!Opt_cofins_total = TBFI!Opt_cofins_total
    TBGravar!Opt_CSLL_total = TBFI!Opt_CSLL_total
    TBGravar!Opt_ISSQN_total = TBFI!Opt_ISSQN_total
    TBGravar!Opt_IRPJ_total = TBFI!Opt_IRPJ_total
    TBGravar!Opt_simples_total = TBFI!Opt_simples_total
    TBGravar!Opt_comissao_total = TBFI!Opt_comissao_total
    TBGravar!Opt_despesas_financeiras_total = TBFI!Opt_despesas_financeiras_total
    TBGravar!Opt_frete_total = TBFI!Opt_frete_total
    TBGravar!Opt_margem_total = TBFI!Opt_margem_total
    TBGravar!Opt_despesas_comerciais_total = TBFI!Opt_despesas_comerciais_total
    TBGravar!Opt_despesas_administrativas_total = TBFI!Opt_despesas_administrativas_total
    TBGravar!ICMS_total = TBFI!ICMS_total
    TBGravar!PIS_total = TBFI!PIS_total
    TBGravar!Cofins_total = TBFI!Cofins_total
    TBGravar!CSLL_total = TBFI!CSLL_total
    TBGravar!ISSQN_total = TBFI!ISSQN_total
    TBGravar!IRPJ_total = TBFI!IRPJ_total
    TBGravar!Simples_total = TBFI!Simples_total
    TBGravar!Comissao_total = TBFI!Comissao_total
    TBGravar!Despesas_financeiras_total = TBFI!Despesas_financeiras_total
    TBGravar!Frete_total = TBFI!Frete_total
    TBGravar!Margem_total = TBFI!Margem_total
    TBGravar!Despesas_comerciais_total = TBFI!Despesas_comerciais_total
    TBGravar!Despesas_administrativas_total = TBFI!Despesas_administrativas_total
    
    TBGravar!chkTotal = TBFI!chkTotal
    TBGravar!Valor_total_processo = TBFI!Valor_total_processo
    TBGravar!Valor_total_materiais = TBFI!Valor_total_materiais
    TBGravar!Valor_total_terceiros = TBFI!Valor_total_terceiros
    TBGravar!Valor_total_outros = TBFI!Valor_total_outros
    TBGravar!Valor_total = TBFI!Valor_total
    TBGravar!obs_engenharia = TBFI!obs_engenharia
    TBGravar!Fechada = TBFI!Fechada
    
    TBGravar!Prazo_engenharia = TBFI!Prazo_engenharia
    TBGravar!Prazo_processos = TBFI!Prazo_engenharia
    TBGravar!Prazo_pcp = TBFI!Prazo_engenharia
    TBGravar!Prazo_qualidade = TBFI!Prazo_engenharia
    TBGravar!Prazo_compras = TBFI!Prazo_engenharia
    
    TBGravar.Update
    txtId = TBGravar!ID
    If Revisar = False Then Conexao.Execute "Update Vendas_analise set ordenaranalise = " & TBGravar!ID & " where ID = " & TBGravar!ID
    TBGravar.Close
    ProcCopiarRevisarSetores
    Lista.ListItems.Clear
    ProcCarregaLista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
End If
TBFI.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCopiarRevisarSetores()
On Error GoTo tratar_erro

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Vendas_analise_setores where IDanalise = " & IDAntigo & " and setor <> 'PROCESSOS'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    Do While TBAbrir.EOF = False
        Set TBGravar = CreateObject("adodb.recordset")
        TBGravar.Open "Select * from Vendas_analise_setores", Conexao, adOpenKeyset, adLockOptimistic
        TBGravar.AddNew
        TBGravar!IDAnalise = txtId
        TBGravar!Responsavel = TBAbrir!Responsavel
        TBGravar!Data = TBAbrir!Data
        TBGravar!Texto = TBAbrir!Texto
        TBGravar!Referencia = TBAbrir!Referencia
        TBGravar!Descricao = TBAbrir!Descricao
        TBGravar!Familia = TBAbrir!Familia
        TBGravar!Un = TBAbrir!Un
        TBGravar!Unidade_com = TBAbrir!Unidade_com
        TBGravar!Qtde = TBAbrir!Qtde
        TBGravar!Peca = TBAbrir!Peca
        TBGravar!Execucao = TBAbrir!Execucao
        TBGravar!Produtividade = TBAbrir!Produtividade
        TBGravar!Preparacao = TBAbrir!Preparacao
        TBGravar!VlrUnit = TBAbrir!VlrUnit
        TBGravar!PrecoHora_Setup = TBAbrir!PrecoHora_Setup
        TBGravar!vlrTotal = TBAbrir!vlrTotal
        TBGravar!Analise = TBAbrir!Analise
        TBGravar!Setor = TBAbrir!Setor
        TBGravar!Data_compras = TBAbrir!Data_compras
        TBGravar!Responsavel_compras = TBAbrir!Responsavel_compras
        TBGravar!pecahora = TBAbrir!pecahora
        TBGravar!TotalHora = TBAbrir!TotalHora
        TBGravar!Trabalho = TBAbrir!Trabalho
        TBGravar!Fase = TBAbrir!Fase
        TBGravar!Grupo_op = TBAbrir!Grupo_op
        TBGravar!Erro_processos = TBAbrir!Erro_processos
        TBGravar!IDProduto = TBAbrir!IDProduto
        TBGravar!Codinterno = TBAbrir!Codinterno
        TBGravar!N_referencia = TBAbrir!N_referencia
        TBGravar!Tipo = TBAbrir!Tipo
        TBGravar!Nao_considerar = TBAbrir!Nao_considerar
        TBGravar!ID_acessorio = TBAbrir!ID_acessorio
        TBGravar.Update
        TBGravar.Close
        TBAbrir.MoveNext
    Loop
End If
TBAbrir.Close
ProcCopiarRevisarCheckList
ProcCopiarRevisarProcessos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCopiarRevisarCheckList()
On Error GoTo tratar_erro

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Vendas_analise_setores_checklist where IDanalise = " & IDAntigo, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    Do While TBAbrir.EOF = False
        Set TBGravar = CreateObject("adodb.recordset")
        TBGravar.Open "Select * from Vendas_analise_setores_checklist", Conexao, adOpenKeyset, adLockOptimistic
        TBGravar.AddNew
        TBGravar!IDAnalise = txtId
        TBGravar!IDchecklist = TBAbrir!IDchecklist
        TBGravar!Setor = TBAbrir!Setor
        TBGravar!Data = TBAbrir!Data
        TBGravar!Responsavel = TBAbrir!Responsavel
        TBGravar!Sim = TBAbrir!Sim
        TBGravar!Quais = TBAbrir!Quais
        TBGravar.Update
        TBGravar.Close
        TBAbrir.MoveNext
    Loop
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCopiarRevisarProcessos()
On Error GoTo tratar_erro

Set TBProcessos = CreateObject("adodb.recordset")
TBProcessos.Open "Select * from Vendas_analise_ProdutosProcessos where ID_analise = " & IDAntigo, Conexao, adOpenKeyset, adLockOptimistic
If TBProcessos.EOF = False Then
    Do While TBProcessos.EOF = False
        Set TBItem = CreateObject("adodb.recordset")
        TBItem.Open "Select * from Vendas_analise_ProdutosProcessos", Conexao, adOpenKeyset, adLockOptimistic
        TBItem.AddNew
        TBItem!id_analise = txtId
        TBItem!Codproduto = TBProcessos!Codproduto
        TBItem!Codinterno = TBProcessos!Codinterno
        TBItem!Referencia = TBProcessos!Referencia
        TBItem!Descricao = TBProcessos!Descricao
        TBItem!Familia = TBProcessos!Familia
        TBItem!Un = TBProcessos!Un
        TBItem!Unidade_com = TBProcessos!Unidade_com
        TBItem!Produto_analise = TBProcessos!Produto_analise
        TBItem!Qtde = TBProcessos!Qtde
        TBItem.Update
        
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from Vendas_analise_setores where IDanalise = " & IDAntigo & " and ID_processo_item = " & TBProcessos!ID & " and setor = 'PROCESSOS'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Do While TBAbrir.EOF = False
                Set TBGravar = CreateObject("adodb.recordset")
                TBGravar.Open "Select * from Vendas_analise_setores", Conexao, adOpenKeyset, adLockOptimistic
                TBGravar.AddNew
                TBGravar!IDAnalise = txtId
                TBGravar!Responsavel = TBAbrir!Responsavel
                TBGravar!Data = TBAbrir!Data
                TBGravar!Texto = TBAbrir!Texto
                TBGravar!Referencia = TBAbrir!Referencia
                TBGravar!Descricao = TBAbrir!Descricao
                TBGravar!Familia = TBAbrir!Familia
                TBGravar!Un = TBAbrir!Un
                TBGravar!Unidade_com = TBAbrir!Unidade_com
                TBGravar!Qtde = TBAbrir!Qtde
                TBGravar!Peca = TBAbrir!Peca
                TBGravar!Execucao = TBAbrir!Execucao
                TBGravar!Produtividade = TBAbrir!Produtividade
                TBGravar!Preparacao = TBAbrir!Preparacao
                TBGravar!VlrUnit = TBAbrir!VlrUnit
                TBGravar!PrecoHora_Setup = TBAbrir!PrecoHora_Setup
                TBGravar!vlrTotal = TBAbrir!vlrTotal
                TBGravar!Analise = TBAbrir!Analise
                TBGravar!Setor = TBAbrir!Setor
                TBGravar!Data_compras = TBAbrir!Data_compras
                TBGravar!Responsavel_compras = TBAbrir!Responsavel_compras
                TBGravar!pecahora = TBAbrir!pecahora
                TBGravar!TotalHora = TBAbrir!TotalHora
                TBGravar!Trabalho = TBAbrir!Trabalho
                TBGravar!Fase = TBAbrir!Fase
                TBGravar!Grupo_op = TBAbrir!Grupo_op
                TBGravar!Erro_processos = TBAbrir!Erro_processos
                TBGravar!IDProduto = TBAbrir!IDProduto
                TBGravar!Codinterno = TBAbrir!Codinterno
                TBGravar!N_referencia = TBAbrir!N_referencia
                TBGravar!Tipo = TBAbrir!Tipo
                TBGravar!ID_processo_item = TBItem!ID
                TBGravar.Update
                TBGravar.Close
                TBAbrir.MoveNext
            Loop
        End If
        TBAbrir.Close
        TBItem.Close
        TBProcessos.MoveNext
    Loop
End If
TBProcessos.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCopiar_engenharia()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If USMsgBox("Deseja realmente copiar o(s) registro(s) da engenharia?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    If FunVerifStatusAnalise("copiar o(s) registro(s) da engenharia", True) = False Then Exit Sub
    If FunVerifValidSetorAnalise("copiar o(s) registro(s) da engenharia", SSTab1.Tab, True) = False Then Exit Sub
    frmVendas_analise_CopiaEngenharia.Show 1
    ProcLimpaCampos_Engenharia_Prod
    ProcCarregaLista_Engenharia_Prod
    Frame7.Enabled = False
    Novo_analise1 = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCopiar_processo()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If USMsgBox("Deseja realmente copiar a(s) fase(s) do processo?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    If FunVerifStatusAnalise("copiar a(s) fase(s) do processo", True) = False Then Exit Sub
    If FunVerifValidSetorAnalise("copiar a(s) fase(s) do processo", SSTab1.Tab, True) = False Then Exit Sub
    frmVendas_analise_CopiaProcesso.Show 1
    ProcLimpaCampos_Processos
    ProcCarregaLista_processos
    Frame6.Enabled = False
    Novo_analise2 = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdFiltrar_processos_Click()
On Error GoTo tratar_erro

If txtMaquina_processos = "" Then
    USMsgBox ("Informe a máquina antes de pesquisar."), vbExclamation, "CAPRIND v5.0"
    txtMaquina_processos.SetFocus
    Exit Sub
End If
Set TBMaquinas = CreateObject("adodb.recordset")
TBMaquinas.Open "select * from cadmaquinas where maquina = '" & txtMaquina_processos & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBMaquinas.EOF = False Then
    txtDescricao_processos = IIf(IsNull(TBMaquinas!Descricao), "", TBMaquinas!Descricao)
    txtValorHora_processos = IIf(IsNull(TBMaquinas!PrecoHora), "", Format(TBMaquinas!PrecoHora, "###,##0.00"))
    txtValorHoraPrep_Processos = IIf(IsNull(TBMaquinas!PrecoHora_Setup), "", Format(TBMaquinas!PrecoHora_Setup, "###,##0.00"))
End If
TBMaquinas.Close

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

Private Sub cmdMaquina_processos_Click()
On Error GoTo tratar_erro

frmVendas_analise_maquina.Show 1

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
Frame1.Enabled = True
ProcLimpaCampos
ProcLimparTudo
Novo_Analise = True
ProcEscondeDataStatus
txtdesenho.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovo_Engenharia_Prod()
On Error GoTo tratar_erro

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If FunVerifStatusAnalise("criar novo registro", True) = False Then Exit Sub
If FunVerifValidSetorAnalise("criar novo registro", SSTab1.Tab, True) = False Then Exit Sub
ProcLimpaCampos_Engenharia_Prod
Novo_analise1 = True
Frame7.Enabled = True
txtdesenho_engenharia.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovo_Engenharia_Checklist()
On Error GoTo tratar_erro

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If FunVerifStatusAnalise("criar novo check-list da engenharia", True) = False Then Exit Sub
If FunVerifValidSetorAnalise("criar novo check-list da engenharia", SSTab1.Tab, True) = False Then Exit Sub
ProcLimpaCampos_Engenharia_Checklist
Novo_analise2 = True
Frame_check(0).Enabled = True
ProcAbrirDescCheckList 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovo_Engenharia_Norma()
On Error GoTo tratar_erro

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If FunVerifStatusAnalise("criar nova norma", True) = False Then Exit Sub
If FunVerifValidSetorAnalise("criar nova norma", SSTab1.Tab, True) = False Then Exit Sub
ProcLimpaCampos_Normas
Novo_analise3 = True
Frame15.Enabled = True
Txt_norma.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovo_Processos_Item()
On Error GoTo tratar_erro

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If FunVerifStatusAnalise("criar novo produto do processo", True) = False Then Exit Sub
If FunVerifValidSetorAnalise("criar novo produto do processo", SSTab1.Tab, True) = False Then Exit Sub
optProcessos_item_analise.Enabled = True
optProcessos_item.Enabled = True
ProcLimpaCampos_Processos_item True
Novo_analise4 = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovo_Processos()
On Error GoTo tratar_erro

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If FunVerifStatusAnalise("criar nova fase no processo", True) = False Then Exit Sub
If FunVerifValidSetorAnalise("criar nova fase do processo", SSTab1.Tab, True) = False Then Exit Sub
ProcLimpaCampos_Processos
Set TBFases = CreateObject("adodb.recordset")
TBFases.Open "Select Fase from Vendas_analise_setores where IDanalise = " & txtId & " and Setor = 'PROCESSOS' and ID_processo_item = " & txtID_processos_item & " order by fase desc", Conexao, adOpenKeyset, adLockOptimistic
If TBFases.EOF = False Then Fase = Int(TBFases!Fase) + 10 Else Fase = "10"
TBFases.Close
Novo_analise5 = True
Frame6.Enabled = True
txtFase = Fase
txtFase.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovo_PCP_Checklist()
On Error GoTo tratar_erro

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If FunVerifStatusAnalise("criar novo check-list do PCP", True) = False Then Exit Sub
If FunVerifValidSetorAnalise("criar novo check-list do PCP", SSTab1.Tab, True) = False Then Exit Sub
ProcLimpaCampos_PCP_Checklist
Novo_analise6 = True
Frame_check(1).Enabled = True
ProcAbrirDescCheckList 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovo_Instrumento()
On Error GoTo tratar_erro

If SSTab_qualidade.Enabled = False Then Exit Sub
If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If FunVerifStatusAnalise("criar novo instrumento", True) = False Then Exit Sub
If FunVerifValidSetorAnalise("criar novo instrumento", SSTab1.Tab, True) = False Then Exit Sub
ProcLimpaCampos_Instrumentos
Novo_analise7 = True
Frame8.Enabled = True
txtdesenho_qualidade.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovo_Qualidade_Checklist()
On Error GoTo tratar_erro

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If FunVerifStatusAnalise("criar novo check-list da qualidade", True) = False Then Exit Sub
If FunVerifValidSetorAnalise("criar novo check-list da qualidade", SSTab1.Tab, True) = False Then Exit Sub
ProcLimpaCampos_Qualidade_Checklist
Novo_analise8 = True
Frame_check(2).Enabled = True
ProcAbrirDescCheckList 2

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovo_Compras_Checklist()
On Error GoTo tratar_erro

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If FunVerifStatusAnalise("criar novo check-list de compras", True) = False Then Exit Sub
If FunVerifValidSetorAnalise("criar novo check-list de compras", SSTab1.Tab, True) = False Then Exit Sub
ProcLimpaCampos_Compras_Checklist
Novo_analise9 = True
Frame_check(3).Enabled = True
ProcAbrirDescCheckList 3

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procNovo_doc()
On Error GoTo tratar_erro

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If FunVerifStatusAnalise("criar novo documento", True) = False Then Exit Sub
Proclimpacampos_doc
Novo_analise10 = True
Frame14.Enabled = True
cmdImportar_Click

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPesquisar_preocessos_item_Click()
On Error GoTo tratar_erro

ProcCarregaDadosProdutoProcessos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdProduto_Click()
On Error GoTo tratar_erro

Sit_REG = 1
frmVendas_analise_item.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdProduto_engenharia_Click()
On Error GoTo tratar_erro

Sit_REG = 2
frmVendas_analise_item.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdProduto_qualidade_Click()
On Error GoTo tratar_erro

Sit_REG = 3
frmVendas_analise_item.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcProximo()
On Error GoTo tratar_erro

If txtId = 0 Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from vendas_analise order by id", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.BOF = False Then
    TBLISTA.Find ("id = " & txtId)
    TBLISTA.MoveNext
    If TBLISTA.EOF = False Then
        ProcLimpaCampos
        ProcLimpaCampos_Engenharia_Prod
        ProcLimpaCampos_Engenharia_Checklist
        ProcLimpaCampos_Normas
        ProcLimpaCampos_Processos
        ProcLimpaCampos_Processos_item True
        ProcLimpaCampos_PCP
        ProcLimpaCampos_Compras
        txtId = TBLISTA!ID
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from vendas_analise where id  = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            ProcPuxaDados
            ProcPuxadados_Engenharia_Checklist
        End If
        TBAbrir.Close
        ProcCarregaLista_Engenharia_Prod
        ProcCarregaLista_Engenharia_Checklist
        ProcCarregaLista_Engenharia_Normas
        ProcCarregaLista_processos_item
        ProcCarregaLista_PCP_Checklist
        ProcCarregaLista_Instrumentos
        ProcCarregaLista_Qualidade_Checklist
        ProcCarregalista_Compras
        ProcCarregaLista_Compras_Checklist
        ProcCarregaLista_Doc
    Else
        USMsgBox ("Fim dos cadastros da análise."), vbInformation, "CAPRIND v5.0"
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcRevisao()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If txtId = "" Then
    USMsgBox ("Informe a análise crítica antes de revisar."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If USMsgBox("Deseja realmente revisar a análise crítica " & Txt_analise & "?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    If FunVerifStatusAnalise("revisar análise crítica", True) = False Then Exit Sub
    Revisar = True
    '==================================
    Modulo = "Outros/Análise crítica"
    Evento = "Revisar"
    ID_documento = txtId.Text
    Documento = "Nº análise: " & Txt_analise & " - Rev.: " & Txt_rev_analise
    Documento1 = ""
    ProcGravaEvento
    '==================================
    ProcCopiarRevisar
    USMsgBox ("Análise crítica revisada com sucesso."), vbInformation, "CAPRIND v5.0"
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from vendas_analise where id = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        ProcLimpaCampos
        ProcLimpaCampos_Engenharia_Prod
        ProcLimpaCampos_Engenharia_Checklist
        ProcLimpaCampos_Processos
        ProcLimpaCampos_PCP
        ProcLimpaCampos_Instrumentos
        ProcLimpaCampos_Compras
        ProcPuxaDados
        ProcCarregaLista_Engenharia_Prod
        ProcCarregaLista_Engenharia_Checklist
        ProcCarregaLista_Engenharia_Normas
        ProcCarregaLista_processos_item
        ProcCarregaLista_processos
        ProcCarregaLista_PCP_Checklist
        ProcCarregaLista_Instrumentos
        ProcCarregaLista_Qualidade_Checklist
        ProcCarregalista_Compras
    End If
    Frame1.Enabled = True
End If

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
            Case vbKeyF2: ProcAbrir
            Case vbKeyF3: ProcSalvar
            Case vbKeyF5: ProcImprimir
            Case vbKeyF7: ProcCopiar
            Case vbKeyF8: ProcRevisao
            'Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
    Case 1:
        Select Case SSTab_engenharia.Tab
            Case 0:
                Select Case KeyCode
                    Case vbKeyInsert: ProcNovo_Engenharia_Prod
                    Case vbKeyF3: ProcSalvar_Engenharia_Prod
                    Case vbKeyF4: ProcExcluir_engenharia_prod
                    Case vbKeyF8: ProcCopiar_engenharia
                    Case vbKeyF5: ProcImprimir
                    'Case vbKeyF1: ProcAjuda
                    Case vbKeyEscape: ProcSair
                End Select
            Case 1:
                Select Case KeyCode
                    Case vbKeyInsert: ProcNovo_Engenharia_Checklist
                    Case vbKeyF3: ProcSalvar_Engenharia_CheckList
                    Case vbKeyF4: ProcExcluir_engenharia_CheckList
                    Case vbKeyF5: ProcImprimir
                    'Case vbKeyF1: ProcAjuda
                    Case vbKeyEscape: ProcSair
                End Select
            Case 2:
                Select Case KeyCode
                    Case vbKeyInsert: ProcNovo_Engenharia_Norma
                    Case vbKeyF3: ProcSalvar_Engenharia_Norma
                    Case vbKeyF4: ProcExcluir_engenharia_Norma
                    Case vbKeyF5: ProcImprimir
                    'Case vbKeyF1: ProcAjuda
                    Case vbKeyEscape: ProcSair
                End Select
        End Select
    Case 2:
        Select Case SSTab_processo.Tab
            Case 0:
                Select Case KeyCode
                    Case vbKeyInsert: ProcNovo_Processos_Item
                    Case vbKeyF3: ProcSalvar_Processos_item
                    Case vbKeyF4: ProcExcluir_Processos_item
                    Case vbKeyF5: ProcImprimir
                    'Case vbKeyF1: ProcAjuda
                    Case vbKeyEscape: ProcSair
                End Select
            Case 1:
                Select Case KeyCode
                    Case vbKeyInsert: ProcNovo_Processos
                    Case vbKeyF3: ProcSalvar_processos
                    Case vbKeyF4: ProcExcluir_processos
                    Case vbKeyF5: ProcImprimir
                    Case vbKeyF7: ProcFerramentas
                    Case vbKeyF8: ProcCopiar_processo
                    'Case vbKeyF1: ProcAjuda
                    Case vbKeyEscape: ProcSair
                End Select
        End Select
    Case 3:
        Select Case KeyCode
            Case vbKeyInsert: ProcNovo_PCP_Checklist
            Case vbKeyF3: ProcSalvar_PCP_CheckList
            Case vbKeyF4: ProcExcluir_PCP_CheckList
            Case vbKeyF5: ProcImprimir
            'Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
    Case 4:
        Select Case SSTab_qualidade.Tab
            Case 0:
                Select Case KeyCode
                    Case vbKeyInsert: ProcNovo_Instrumento
                    Case vbKeyF3: ProcSalvar_Instrumento
                    Case vbKeyF4: ProcExcluir_Instrumento
                    Case vbKeyF5: ProcImprimir
                    Case vbKeyF7: ProcSalvar_Instrumento
                    'Case vbKeyF1: ProcAjuda
                    Case vbKeyEscape: ProcSair
                End Select
            Case 1:
                Select Case KeyCode
                    Case vbKeyInsert: ProcNovo_Qualidade_Checklist
                    Case vbKeyF3: ProcSalvar_Qualidade_CheckList
                    Case vbKeyF4: ProcExcluir_Qualidade_CheckList
                    Case vbKeyF5: ProcImprimir
                    'Case vbKeyF1: ProcAjuda
                    Case vbKeyEscape: ProcSair
                End Select
        End Select
    Case 5:
        Select Case SSTab_compras.Tab
            Case 0:
                Select Case KeyCode
                    Case vbKeyF3: ProcSalvar_Compras
                    Case vbKeyF4: ProcExcluir_Compras
                    Case vbKeyF5: ProcImprimir
                    'Case vbKeyF1: ProcAjuda
                    Case vbKeyEscape: ProcSair
                End Select
            Case 1:
                Select Case KeyCode
                    Case vbKeyInsert: ProcNovo_Compras_Checklist
                    Case vbKeyF3: ProcSalvar_Compras_CheckList
                    Case vbKeyF4: ProcExcluir_Compras_CheckList
                    Case vbKeyF5: ProcImprimir
                    'Case vbKeyF1: ProcAjuda
                    Case vbKeyEscape: ProcSair
                End Select
        End Select
    Case 6:
        Select Case KeyCode
            Case vbKeyInsert: procNovo_doc
            Case vbKeyF2: cmdImportar_Click
            Case vbKeyF3: procSalvar_doc
            Case vbKeyF4: procExcluir_doc
            Case vbKeyF5: ProcImprimir
            'Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcPuxaDados()
On Error GoTo tratar_erro

Caption = "Outros - Análise crítica - (Análise : " & TBAbrir!Nanalise & " - Cód. interno : " & IIf(IsNull(TBAbrir!Codinterno), "", TBAbrir!Codinterno) & ")"
txtdesenho.Text = IIf(IsNull(TBAbrir!Codinterno), "", TBAbrir!Codinterno)

txtId = TBAbrir!ID
Txt_analise = IIf(IsNull(TBAbrir!Nanalise), "", TBAbrir!Nanalise)
Txt_rev_analise = IIf(IsNull(TBAbrir!Revisao), 0, TBAbrir!Revisao)
txtData = IIf(IsNull(TBAbrir!Data), "", Format(TBAbrir!Data, "dd/mm/yy"))
txtResponsavel.Text = IIf(IsNull(TBAbrir!Responsavel), "", TBAbrir!Responsavel)
Txt_status = IIf(IsNull(TBAbrir!status), "", TBAbrir!status)
If Txt_status = "ABERTA EM ANALISE" Then ProcEscondeDataStatus Else ProcMostraDataStatus
Txt_data_status = IIf(IsNull(TBAbrir!Data_status), "", Format(TBAbrir!Data_status, "dd/mm/yy"))
txtidproduto = IIf(IsNull(TBAbrir!IDProduto), 0, TBAbrir!IDProduto)
txtRev_desenho = IIf(IsNull(TBAbrir!RevDesenho), "", TBAbrir!RevDesenho)
If IsNull(TBAbrir!N_referencia) = False And TBAbrir!N_referencia <> "" Then
    With cmbReferencia
        .Clear
        .AddItem TBAbrir!N_referencia
        .Text = TBAbrir!N_referencia
    End With
End If
If IsNull(TBAbrir!Unidade) = False And TBAbrir!Unidade <> "" Then cmbun = TBAbrir!Unidade
If IsNull(TBAbrir!Unidade_com) = False And TBAbrir!Unidade_com <> "" Then Cmb_un_com = TBAbrir!Unidade_com
Txt_qtde_sol = IIf(IsNull(TBAbrir!qtde_solicitada), "", Format(TBAbrir!qtde_solicitada, "###,##0.0000"))
txtQtde = IIf(IsNull(TBAbrir!Qtde), "1,00", Format(TBAbrir!Qtde, "###,##0.0000"))
If IsNull(TBAbrir!Familia) = False And TBAbrir!Familia <> "" Then
    VerifDadosPadraoFamilia = False
    cmbfamilia = TBAbrir!Familia
    VerifDadosPadraoFamilia = True
End If
txtdescricao = IIf(IsNull(TBAbrir!Descricao), "", TBAbrir!Descricao)
Txt_tipo = IIf(IsNull(TBAbrir!Tipo), "", TBAbrir!Tipo)
txtIDcliente = IIf(IsNull(TBAbrir!IDCliente), 0, TBAbrir!IDCliente)
txtCliente = IIf(IsNull(TBAbrir!Cliente), "", TBAbrir!Cliente)
txtObs = IIf(IsNull(TBAbrir!Obs), "", TBAbrir!Obs)
txtContato = IIf(IsNull(TBAbrir!contato), "", TBAbrir!contato)
txtdepartamento = IIf(IsNull(TBAbrir!Departamento), "", TBAbrir!Departamento)
txttelefone = IIf(IsNull(TBAbrir!telefone), "", TBAbrir!telefone)
txtFax = IIf(IsNull(TBAbrir!Fax), "", TBAbrir!Fax)
txtEmail = IIf(IsNull(TBAbrir!Email), "", TBAbrir!Email)
txtNRef = IIf(IsNull(TBAbrir!NRef), "", TBAbrir!NRef)
txtreferencia = IIf(IsNull(TBAbrir!Referencia), "", TBAbrir!Referencia)
cmbLocal_entrega = IIf(IsNull(TBAbrir!Local_entrega), "", TBAbrir!Local_entrega)
txtID_entrega = IIf(IsNull(TBAbrir!ID_entrega), "", TBAbrir!ID_entrega)
cmbLocal_cobranca = IIf(IsNull(TBAbrir!Local_Cobranca), "", TBAbrir!Local_Cobranca)
txtID_cobranca = IIf(IsNull(TBAbrir!ID_Cobranca), "", TBAbrir!ID_Cobranca)

'Engenharia
txtData_engenharia_prod = IIf(IsNull(TBAbrir!Data_Engenharia), "", Format(TBAbrir!Data_Engenharia, "dd/mm/yy"))
txtResponsavel_engenharia_prod = IIf(IsNull(TBAbrir!Responsavel_Engenharia), "", TBAbrir!Responsavel_Engenharia)
txtPrazo_Engenharia = IIf(IsNull(TBAbrir!Prazo_engenharia), "", Format(TBAbrir!Prazo_engenharia, "DD/mm/YY"))
txtDtValidacao_Engenharia = IIf(IsNull(TBAbrir!DtValidacao_Engenharia), "", TBAbrir!DtValidacao_Engenharia)
txtRespValidacao_Engenharia = IIf(IsNull(TBAbrir!RespValidacao_Engenharia), "", TBAbrir!RespValidacao_Engenharia)
Txt_obs_engenharia = IIf(IsNull(TBAbrir!obs_engenharia), "", TBAbrir!obs_engenharia)

'Processo
txtPrazo_Processo = IIf(IsNull(TBAbrir!Prazo_processos), "", Format(TBAbrir!Prazo_processos, "DD/mm/YY"))
txtDtValidacao_processo = IIf(IsNull(TBAbrir!DtValidacao_Processo), "", TBAbrir!DtValidacao_Processo)
txtRespValidacao_processo = IIf(IsNull(TBAbrir!RespValidacao_Processo), "", TBAbrir!RespValidacao_Processo)

'PCP
txtData_PCP = IIf(IsNull(TBAbrir!data_PCP), "", Format(TBAbrir!data_PCP, "dd/mm/yy"))
txtResponsavel_PCP = IIf(IsNull(TBAbrir!responsavel_PCP), "", TBAbrir!responsavel_PCP)
txtPrazo_PCP = IIf(IsNull(TBAbrir!Prazo_pcp), "", Format(TBAbrir!Prazo_pcp, "DD/mm/YY"))
txtDtValidacao_PCP = IIf(IsNull(TBAbrir!DtValidacao_Pcp), "", TBAbrir!DtValidacao_Pcp)
txtRespValidacao_PCP = IIf(IsNull(TBAbrir!RespValidacao_Pcp), "", TBAbrir!RespValidacao_Pcp)
txtAnalise_PCP = IIf(IsNull(TBAbrir!Analise_PCP), "", TBAbrir!Analise_PCP)

'Qualidade
txtPrazo_Qualidade = IIf(IsNull(TBAbrir!Prazo_qualidade), "", Format(TBAbrir!Prazo_qualidade, "DD/mm/YY"))
txtDtValidacao_Qualidade = IIf(IsNull(TBAbrir!DtValidacao_Qualidade), "", TBAbrir!DtValidacao_Qualidade)
txtRespValidacao_Qualidade = IIf(IsNull(TBAbrir!RespValidacao_Qualidade), "", TBAbrir!RespValidacao_Qualidade)

'Compras
txtPrazo_Compras = IIf(IsNull(TBAbrir!Prazo_compras), "", Format(TBAbrir!Prazo_compras, "DD/mm/YY"))
txtDtValidacao_Compras = IIf(IsNull(TBAbrir!DtValidacao_Compras), "", TBAbrir!DtValidacao_Compras)
txtRespValidacao_Compras = IIf(IsNull(TBAbrir!RespValidacao_Compras), "", TBAbrir!RespValidacao_Compras)

Frame1.Enabled = True
Novo_Analise = False
ProcLimparTudo

'Verifica se o produto já esta cadastrado
Procliberacampos
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select * from projproduto where Desenho = '" & txtdesenho & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    ProcBloqueiaCampos
End If
TBProduto.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcPuxadados_Engenharia_Prod()
On Error GoTo tratar_erro

txtdesenho_engenharia = IIf(IsNull(TBMaterial!Codinterno), "", TBMaterial!Codinterno)
txtID_engenharia = TBMaterial!ID
txtData_engenharia_prod = IIf(IsNull(TBMaterial!Data), "", Format(TBMaterial!Data, "dd/mm/yy"))
txtResponsavel_engenharia_prod = IIf(IsNull(TBMaterial!Responsavel), "", TBMaterial!Responsavel)
txtIDproduto_engenharia = IIf(IsNull(TBMaterial!IDProduto), 0, TBMaterial!IDProduto)
If TBMaterial!Tipo = "M" Then Opt_material.Value = True
If TBMaterial!Tipo = "O" Then Opt_outros.Value = True
If TBMaterial!Tipo = "T" Then Opt_terceiros.Value = True
If IsNull(TBMaterial!N_referencia) = False And TBMaterial!N_referencia <> "" Then
    cmbReferencia_engenharia.AddItem TBMaterial!N_referencia
    cmbReferencia_engenharia = TBMaterial!N_referencia
End If
If IsNull(TBMaterial!Un) = False And TBMaterial!Un <> "" Then cmbun_engenharia = TBMaterial!Un
If IsNull(TBMaterial!Unidade_com) = False And TBMaterial!Unidade_com <> "" Then Cmb_un_com_engenharia = TBMaterial!Unidade_com
txtQtde_engenharia = IIf(IsNull(TBMaterial!Qtde), "", Format(TBMaterial!Qtde, "###,##0.0000"))
If IsNull(TBMaterial!Familia) = False And TBMaterial!Familia <> "" Then
    VerifDadosPadraoFamilia = False
    cmbfamilia_engenharia = TBMaterial!Familia
    VerifDadosPadraoFamilia = True
End If
txtdescricao_engenharia = IIf(IsNull(TBMaterial!Texto), "", TBMaterial!Texto)
txtAnalise_engenharia = IIf(IsNull(TBMaterial!Analise), "", TBMaterial!Analise)
Novo_analise1 = False
Frame7.Enabled = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcPuxadados_Engenharia_Checklist()
On Error GoTo tratar_erro

Txt_ID_check(0) = TBMaterial!ID
Txt_ID_descricao_check(0) = TBMaterial!IDchecklist
Txt_descricao_chek(0) = IIf(IsNull(TBMaterial!Descricao), "", TBMaterial!Descricao)
Txt_data_check(0) = IIf(IsNull(TBMaterial!Data), "", Format(TBMaterial!Data, "dd/mm/yy"))
Txt_responsavel_check(0) = IIf(IsNull(TBMaterial!Responsavel), "", TBMaterial!Responsavel)
If TBMaterial!Sim = True Then
    Chk_sim_chek(0).Value = 1
    Txt_texto_check(0) = IIf(IsNull(TBMaterial!Quais), "", TBMaterial!Quais)
Else
    Chk_nao_chek(0).Value = 1
End If
Novo_analise2 = False
Frame_check(0).Enabled = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcPuxadados_Normas()
On Error GoTo tratar_erro

Txt_ID_norma = TBMaterial!ID
txtData_norma = IIf(IsNull(TBMaterial!Data), "", Format(TBMaterial!Data, "dd/mm/yy"))
txtResponsavel_norma = IIf(IsNull(TBMaterial!Responsavel), "", TBMaterial!Responsavel)
Txt_norma = IIf(IsNull(TBMaterial!Texto), "", TBMaterial!Texto)
Txt_obs_norma = IIf(IsNull(TBMaterial!Analise), "", TBMaterial!Analise)
Novo_analise3 = False
Frame15.Enabled = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcPuxadados_Processos_item()
On Error GoTo tratar_erro

If TBMaterial!Produto_analise = True Then optProcessos_item_analise.Value = True Else optProcessos_item.Value = True
txtCodInterno_processos_item = IIf(IsNull(TBMaterial!Codinterno), "", TBMaterial!Codinterno)
ProcCarregaDadosProdutoProcessos
txtID_processos_item = TBMaterial!ID
txtIDproduto_processos = IIf(IsNull(TBMaterial!Codproduto), 0, TBMaterial!Codproduto)
If IsNull(TBMaterial!Referencia) = False And TBMaterial!Referencia <> "" Then cmbReferencia_processos_item = TBMaterial!Referencia
If IsNull(TBMaterial!Un) = False And TBMaterial!Un <> "" Then cmbUn_processos_item = TBMaterial!Un
If IsNull(TBMaterial!Unidade_com) = False And TBMaterial!Unidade_com <> "" Then Cmb_un_com_processos_item = TBMaterial!Unidade_com
If IsNull(TBMaterial!Familia) = False And TBMaterial!Familia <> "" Then
    VerifDadosPadraoFamilia = False
    cmbFamilia_processos_item = TBMaterial!Familia
    VerifDadosPadraoFamilia = True
End If
txtDescricao_processos_item = IIf(IsNull(TBMaterial!Descricao), "", TBMaterial!Descricao)
txtQtde_processos_item = IIf(IsNull(TBMaterial!Qtde), "1,000", Format(TBMaterial!Qtde, "###,##0.0000"))
Novo_analise4 = False
If optProcessos_item.Value = True Then Frame2.Enabled = True Else Frame2.Enabled = False
optProcessos_item.Enabled = True
optProcessos_item_analise.Enabled = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcPuxadados_processos()
On Error GoTo tratar_erro

txtID_processos = TBAbrir!ID
txtMaquina_processos = IIf(IsNull(TBAbrir!Texto), "", TBAbrir!Texto)
txtDescricao_processos = IIf(IsNull(TBAbrir!Descricao), "", TBAbrir!Descricao)
txtTotalHora_processos = IIf(IsNull(TBAbrir!TotalHora), "", Format(TBAbrir!TotalHora, "hh:mm:ss"))
txtPecaHora_processos = IIf(IsNull(TBAbrir!Peca), "1", TBAbrir!Peca)
Procentrada
If TempoPreparacao = "" Then txtPreparacao_processos.Text = "___:__:__" Else txtPreparacao_processos.Text = TempoPreparacao
If TempoExecucao = "" Then txtExecucao_processos.Text = "___:__:__" Else txtExecucao_processos.Text = TempoExecucao
txtData_processos = IIf(IsNull(TBAbrir!Data), "", Format(TBAbrir!Data, "dd/mm/yy"))
txtResponsavel_processos = IIf(IsNull(TBAbrir!Responsavel), "", TBAbrir!Responsavel)
txtValorHora_processos = IIf(IsNull(TBAbrir!VlrUnit), "", Format(TBAbrir!VlrUnit, "###,##0.00"))
txtValorHoraPrep_Processos = IIf(IsNull(TBAbrir!PrecoHora_Setup), "", Format(TBAbrir!PrecoHora_Setup, "###,##0.00"))
txtValorTotal_processos = IIf(IsNull(TBAbrir!vlrTotal), "", Format(TBAbrir!vlrTotal, "###,##0.00"))
txtTrabalho.TextRTF = IIf(IsNull(TBAbrir!Trabalho), "", TBAbrir!Trabalho)
txtgrupo_op = IIf(IsNull(TBAbrir!Grupo_op), "", TBAbrir!Grupo_op)
txtFase = IIf(IsNull(TBAbrir!Fase), "", TBAbrir!Fase)
txtErro = IIf(IsNull(TBAbrir!Erro_processos), "", TBAbrir!Erro_processos)
If TBAbrir!pecahora = True Then chkPchora.Value = 1 Else chkPchora.Value = 0
Novo_analise5 = False
Frame6.Enabled = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcPuxadados_PCP_Checklist()
On Error GoTo tratar_erro

Txt_ID_check(1) = TBMaterial!ID
Txt_ID_descricao_check(1) = TBMaterial!IDchecklist
Txt_descricao_chek(1) = IIf(IsNull(TBMaterial!Descricao), "", TBMaterial!Descricao)
Txt_data_check(1) = IIf(IsNull(TBMaterial!Data), "", Format(TBMaterial!Data, "dd/mm/yy"))
Txt_responsavel_check(1) = IIf(IsNull(TBMaterial!Responsavel), "", TBMaterial!Responsavel)
If TBMaterial!Sim = True Then
    Chk_sim_chek(1).Value = 1
    Txt_texto_check(1) = IIf(IsNull(TBMaterial!Quais), "", TBMaterial!Quais)
Else
    Chk_nao_chek(1).Value = 1
End If
Novo_analise6 = False
Frame_check(1).Enabled = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcPuxadados_Instrumentos()
On Error GoTo tratar_erro

txtdesenho_qualidade = IIf(IsNull(TBAbrir!Codinterno), "", TBAbrir!Codinterno)
txtID_qualidade = TBAbrir!ID
txtData_Qualidade1 = IIf(IsNull(TBAbrir!Data), "", Format(TBAbrir!Data, "dd/mm/yy"))
txtResponsavel_Qualidade1 = IIf(IsNull(TBAbrir!Responsavel), "", TBAbrir!Responsavel)
txtIDproduto_qualidade = IIf(IsNull(TBAbrir!IDProduto), 0, TBAbrir!IDProduto)
If IsNull(TBAbrir!N_referencia) = False And TBAbrir!N_referencia <> "" Then cmbReferencia_qualidade = TBAbrir!N_referencia
If IsNull(TBAbrir!Un) = False And TBAbrir!Un <> "" Then cmbun_qualidade = TBAbrir!Un
If IsNull(TBAbrir!Unidade_com) = False And TBAbrir!Unidade_com <> "" Then Cmb_un_com_qualidade = TBAbrir!Unidade_com
txtQtde_Qualidade = IIf(IsNull(TBAbrir!Qtde), "", Format(TBAbrir!Qtde, "###,##0.0000"))
If IsNull(TBAbrir!Familia) = False And TBAbrir!Familia <> "" Then
    VerifDadosPadraoFamilia = False
    cmbfamilia_qualidade = TBAbrir!Familia
    VerifDadosPadraoFamilia = True
End If
txtdescricao_Qualidade = IIf(IsNull(TBAbrir!Texto), "", TBAbrir!Texto)
txtAnalise_qualidade = IIf(IsNull(TBAbrir!Analise), "", TBAbrir!Analise)
Novo_analise7 = False
Frame8.Enabled = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcPuxadados_Qualidade_Checklist()
On Error GoTo tratar_erro

Txt_ID_check(2) = TBMaterial!ID
Txt_ID_descricao_check(2) = TBMaterial!IDchecklist
Txt_descricao_chek(2) = IIf(IsNull(TBMaterial!Descricao), "", TBMaterial!Descricao)
Txt_data_check(2) = IIf(IsNull(TBMaterial!Data), "", Format(TBMaterial!Data, "dd/mm/yy"))
Txt_responsavel_check(2) = IIf(IsNull(TBMaterial!Responsavel), "", TBMaterial!Responsavel)
If TBMaterial!Sim = True Then
    Chk_sim_chek(2).Value = 1
    Txt_texto_check(2) = IIf(IsNull(TBMaterial!Quais), "", TBMaterial!Quais)
Else
    Chk_nao_chek(2).Value = 1
End If
Novo_analise8 = False
Frame_check(2).Enabled = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcPuxadados_compras()
On Error GoTo tratar_erro

txtID_Compras = TBAbrir!ID
txtData_Compras = IIf(IsNull(TBAbrir!Data_compras), "", Format(TBAbrir!Data_compras, "dd/mm/yy"))
txtResponsavel_Compras = IIf(IsNull(TBAbrir!Responsavel_compras), "", TBAbrir!Responsavel_compras)
Txtdesenho_compras = IIf(IsNull(TBAbrir!Codinterno), "", TBAbrir!Codinterno)
Txt_referencia_compras = IIf(IsNull(TBAbrir!N_referencia), "", TBAbrir!N_referencia)
Txt_familia_compras = IIf(IsNull(TBAbrir!Familia), "", TBAbrir!Familia)
txtSetor_compras = IIf(IsNull(TBAbrir!Setor), "", TBAbrir!Setor)
txtTexto_Compras = IIf(IsNull(TBAbrir!Texto), "", TBAbrir!Texto)
txtQtde_compras = IIf(IsNull(TBAbrir!Qtde), 0, Format(TBAbrir!Qtde, "###,##0.0000"))
txtValor_Compras = IIf(IsNull(TBAbrir!VlrUnit), 0, Format(TBAbrir!VlrUnit, "###,##0.0000"))
Txt_valor_total = IIf(IsNull(TBAbrir!vlrTotal), 0, Format(TBAbrir!vlrTotal, "###,##0.00"))
Frame22.Enabled = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcPuxadados_Compras_Checklist()
On Error GoTo tratar_erro

Txt_ID_check(3) = TBMaterial!ID
Txt_ID_descricao_check(3) = TBMaterial!IDchecklist
Txt_descricao_chek(3) = IIf(IsNull(TBMaterial!Descricao), "", TBMaterial!Descricao)
Txt_data_check(3) = IIf(IsNull(TBMaterial!Data), "", Format(TBMaterial!Data, "dd/mm/yy"))
Txt_responsavel_check(3) = IIf(IsNull(TBMaterial!Responsavel), "", TBMaterial!Responsavel)
If TBMaterial!Sim = True Then
    Chk_sim_chek(3).Value = 1
    Txt_texto_check(3) = IIf(IsNull(TBMaterial!Quais), "", TBMaterial!Quais)
Else
    Chk_nao_chek(3).Value = 1
End If
Novo_analise9 = False
Frame_check(3).Enabled = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcPuxadados_Doc()
On Error GoTo tratar_erro

txtID_doc = TBMaterial!ID
txtData_doc = IIf(IsNull(TBMaterial!Data), "", Format(TBMaterial!Data, "dd/mm/yy"))
txtResponsavel_doc = IIf(IsNull(TBMaterial!Responsavel), "", TBMaterial!Responsavel)
txt_Caminho = IIf(IsNull(TBMaterial!Texto), "", TBMaterial!Texto)
Txt_obs_doc = IIf(IsNull(TBMaterial!Analise), "", TBMaterial!Analise)
Novo_analise10 = False
Frame14.Enabled = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCampos()
On Error GoTo tratar_erro

txtId = 0
Txt_rev_analise = 0
txtidproduto = 0
Txt_analise = ""
Txt_rev_analise = 0
txtData = Format(Date, "dd/mm/yy")
txtResponsavel = pubUsuario
Txt_status = "ABERTA EM ANALISE"
Txt_data_status = ""
txtdesenho = ""
txtRev_desenho = ""
cmbReferencia.Clear
cmbun.ListIndex = -1
Cmb_un_com.ListIndex = -1
Txt_qtde_sol = ""
txtQtde = "1,00"
txtdescricao = ""
cmbfamilia.ListIndex = -1
Txt_tipo = ""
txtIDcliente = 0
txtCliente = ""
txtObs = ""
txtContato = ""
txtdepartamento = ""
txttelefone = ""
txtFax = ""
txtEmail = ""
txtNRef = ""
txtreferencia = ""
txtID_cobranca = 0
cmbLocal_cobranca = ""
txtID_entrega = 0
cmbLocal_entrega = ""
optProcessos_item.Enabled = False
optProcessos_item_analise.Enabled = False
optProcessos_item.Value = 0
optProcessos_item_analise.Value = 0

'Engenharia
ProcLimpaCampos_Engenharia

'Processo
txtPrazo_Processo = ""
txtDtValidacao_processo = ""
txtRespValidacao_processo = ""

'PCP
ProcLimpaCampos_PCP

'Qualidade
txtPrazo_Qualidade = ""
txtDtValidacao_Qualidade = ""
txtRespValidacao_Qualidade = ""

'Compras
txtPrazo_Compras = ""
txtDtValidacao_Compras = ""
txtRespValidacao_Compras = ""

Caption = "Outros - Análise crítica"
CodigoLista = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

'Sub ProcLimpaCampos2()
'On Error GoTo tratar_erro
'
'txtIDProduto = 0
'txtRev_desenho = ""
'cmbreferencia.Clear
'cmbun.ListIndex = -1
'Cmb_un_com.ListIndex = -1
'txtDescricao = ""
'cmbfamilia.ListIndex = -1
'
'Exit Sub
'tratar_erro:
'    usMsgbox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
'    Exit Sub
'End Sub

Sub ProcLimpaCampos_Engenharia()
On Error GoTo tratar_erro

txtData_engenharia = Format(Date, "dd/mm/yy")
txtResponsavel_engenharia = pubUsuario
txtPrazo_Engenharia = ""
txtDtValidacao_Engenharia = ""
txtRespValidacao_Engenharia = ""
Txt_obs_engenharia = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCampos_Engenharia_Prod()
On Error GoTo tratar_erro

txtID_engenharia = 0
txtIDproduto_engenharia = 0
txtData_engenharia_prod = Format(Date, "dd/mm/yy")
txtResponsavel_engenharia_prod = pubUsuario
Opt_material.Value = False
Opt_terceiros.Value = False
Opt_outros.Value = False
txtdesenho_engenharia = ""
cmbReferencia_engenharia.Clear
txtReferencia_engenharia = ""
chkAuto_engenharia.Value = 0
chkManual_engenharia.Value = 0
cmbun_engenharia.ListIndex = -1
Cmb_un_com_engenharia.ListIndex = -1
txtQtde_engenharia = ""
txtdescricao_engenharia = ""
cmbfamilia_engenharia.ListIndex = -1
txtAnalise_engenharia = ""
CodigoLista1 = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCampos_Engenharia_Checklist()
On Error GoTo tratar_erro

Txt_ID_check(0) = 0
Txt_data_check(0) = Format(Date, "dd/mm/yy")
Txt_responsavel_check(0) = pubUsuario
Txt_ID_descricao_check(0) = 0
Txt_descricao_chek(0) = ""
Chk_nao_chek(0).Value = 0
Chk_sim_chek(0).Value = 0
Txt_texto_check(0) = ""
CodigoLista2 = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCampos_Normas()
On Error GoTo tratar_erro

Txt_ID_norma = 0
txtResponsavel_norma = pubUsuario
txtData_norma = Format(Date, "dd/mm/yy")
Txt_norma = ""
Txt_obs_norma = ""
CodigoLista3 = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCampos_Processos_item(Principal As Boolean)
On Error GoTo tratar_erro

If Principal = True Then
    txtID_processos_item = 0
    CodigoLista4 = 0
    optProcessos_item_analise.Value = False
    optProcessos_item.Value = False
End If
txtIDproduto_processos = ""
txtCodInterno_processos_item = ""
cmbReferencia_processos_item.Clear
txtDescricao_processos_item = ""
cmbFamilia_processos_item.ListIndex = -1
cmbUn_processos_item.ListIndex = -1
Cmb_un_com_processos_item.ListIndex = -1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCampos_Processos()
On Error GoTo tratar_erro

txtID_processos = 0
txtResponsavel_processos = pubUsuario
txtData_processos = Format(Date, "dd/mm/yy")
txtTotalHora_processos = ""
txtMaquina_processos = ""
txtDescricao_processos = ""
txtPecaHora_processos = "1"
txtValorHora_processos = ""
txtValorHoraPrep_Processos = ""
txtValorTotal_processos = ""
txtErro = ""
txtFase = ""
txtgrupo_op = ""
txtTrabalho.Text = ""
txtExecucao_processos = "___:__:__"
txtPreparacao_processos = "___:__:__"
chkPchora.Value = 0
CodigoLista5 = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCampos_PCP()
On Error GoTo tratar_erro

txtData_PCP = Format(Date, "dd/mm/yy")
txtResponsavel_PCP = pubUsuario
txtPrazo_PCP = ""
txtDtValidacao_PCP = ""
txtRespValidacao_PCP = ""
txtAnalise_PCP = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCampos_PCP_Checklist()
On Error GoTo tratar_erro

Txt_ID_check(1) = 0
Txt_data_check(1) = Format(Date, "dd/mm/yy")
Txt_responsavel_check(1) = pubUsuario
Txt_ID_descricao_check(1) = 0
Txt_descricao_chek(1) = ""
Chk_nao_chek(1).Value = 0
Chk_sim_chek(1).Value = 0
Txt_texto_check(1) = ""
CodigoLista6 = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCampos_Instrumentos()
On Error GoTo tratar_erro

txtID_qualidade = 0
txtIDproduto_qualidade = 0
txtResponsavel_Qualidade1 = pubUsuario
txtData_Qualidade1 = Format(Date, "dd/mm/yy")
txtdesenho_qualidade = ""
cmbReferencia_qualidade.Clear
chkAuto_qualidade.Value = 0
chkManual_qualidade.Value = 0
cmbun_qualidade.ListIndex = -1
Cmb_un_com_qualidade.ListIndex = -1
txtQtde_Qualidade = ""
txtdescricao_Qualidade = ""
cmbfamilia_engenharia.ListIndex = -1
txtAnalise_qualidade = ""
CodigoLista7 = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCampos_Qualidade_Checklist()
On Error GoTo tratar_erro

Txt_ID_check(2) = 0
Txt_data_check(2) = Format(Date, "dd/mm/yy")
Txt_responsavel_check(2) = pubUsuario
Txt_ID_descricao_check(2) = 0
Txt_descricao_chek(2) = ""
Chk_nao_chek(2).Value = 0
Chk_sim_chek(2).Value = 0
Txt_texto_check(2) = ""
CodigoLista8 = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCampos_Compras()
On Error GoTo tratar_erro

txtID_Compras = 0
txtResponsavel_Compras = pubUsuario
txtData_Compras = Format(Date, "dd/mm/yy")
Txtdesenho_compras = ""
Txt_referencia_compras = ""
Txt_familia_compras = ""
txtSetor_compras = ""
txtTexto_Compras = ""
txtQtde_compras = ""
txtValor_Compras = ""
Txt_valor_total = ""
CodigoLista9 = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCampos_Compras_Checklist()
On Error GoTo tratar_erro

Txt_ID_check(3) = 0
Txt_data_check(3) = Format(Date, "dd/mm/yy")
Txt_responsavel_check(3) = pubUsuario
Txt_ID_descricao_check(3) = 0
Txt_descricao_chek(3) = ""
Chk_nao_chek(3).Value = 0
Chk_sim_chek(3).Value = 0
Txt_texto_check(3) = ""
CodigoLista10 = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub Proclimpacampos_doc()
On Error GoTo tratar_erro

txtID_doc = 0
txtData_doc = Format(Date, "dd/mm/yy")
txtResponsavel_doc = pubUsuario
txt_Caminho = ""
Txt_obs_doc = ""
CodigoLista11 = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluir_Engenharia()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso.")
    Exit Sub
End If
If USMsgBox("Deseja realmente excluir o(s) cadastro(s) da engenharia?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    If FunVerifStatusAnalise("excluir dados da engenharia", True) = False Then Exit Sub
    If FunVerifValidSetorAnalise("excluir dados da engenharia", SSTab1.Tab, True) = False Then Exit Sub
    Conexao.Execute "Update vendas_analise Set data_engenharia = Null, responsavel_engenharia = 'Null', Obs_engenharia = Null where id = " & txtId
    USMsgBox ("Dados da engenharia excluídos com sucesso."), vbInformation, "CAPRIND v5.0"
    '==================================
    Modulo = "Outros/Análise crítica"
    Evento = "Excluir dados da engenharia"
    ID_documento = txtId
    Documento = "Nº análise: " & Txt_analise & " - Rev.: " & Txt_rev_analise
    Documento1 = ""
    ProcGravaEvento
    '==================================
    ProcLimpaCampos_Engenharia
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluir_engenharia_prod()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso.")
    Exit Sub
End If
Permitido = False
With Lista_engenharia
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) registro(s)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            Conexao.Execute "DELETE from vendas_analise_setores where id  = " & .ListItems(InitFor)
            '==================================
            Modulo = "Outros/Análise crítica"
            Evento = "Excluir registro da engenharia"
            ID_documento = .ListItems.Item(InitFor)
            Documento = "Nº análise: " & Txt_analise & " - Rev.: " & Txt_rev_analise
            Documento1 = "Cód. interno: " & .ListItems(InitFor).SubItems(1) & " - Descrição: " & .ListItems(InitFor).SubItems(2)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) registro(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Registro(s) excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCampos_Engenharia_Prod
    ProcCarregaLista_Engenharia_Prod
    Novo_analise1 = False
    Frame7.Enabled = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluir_engenharia_CheckList()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso.")
    Exit Sub
End If
Permitido = False
With Lista_check(0)
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) check-list da engenharia?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            Conexao.Execute "DELETE from Vendas_analise_setores_checklist where id  = " & .ListItems(InitFor)
            '==================================
            Modulo = "Outros/Análise crítica"
            Evento = "Excluir check-list da engenharia"
            ID_documento = .ListItems.Item(InitFor)
            Documento = "Nº análise: " & Txt_analise & " - Rev.: " & Txt_rev_analise
            Documento1 = "Descrição: " & .ListItems(InitFor).SubItems(3)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) check-list da engenharia antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Check-list da engenharia excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCampos_Engenharia_Checklist
    ProcCarregaLista_Engenharia_Checklist
    Novo_analise2 = False
    Frame_check(0).Enabled = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluir_engenharia_Norma()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso.")
    Exit Sub
End If
Permitido = False
With Lista_normas
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir esta(s) norma(s)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            Conexao.Execute "DELETE from vendas_analise_setores where id  = " & .ListItems(InitFor)
            '==================================
            Modulo = "Outros/Análise crítica"
            Evento = "Excluir norma"
            ID_documento = .ListItems.Item(InitFor)
            Documento = "Nº análise: " & Txt_analise & " - Rev.: " & Txt_rev_analise
            Documento1 = "Norma: " & .ListItems(InitFor).SubItems(1)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe a(s) norma(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Norma(s) excluída(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCampos_Normas
    ProcCarregaLista_Engenharia_Normas
    Novo_analise3 = False
    Frame15.Enabled = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluir_Processos_item()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso.")
    Exit Sub
End If

Permitido = False
With Lista_processos_item
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) produto(s) do processo?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            Conexao.Execute "DELETE from vendas_analise_ProdutosProcessos where id  = " & .ListItems(InitFor)
            Conexao.Execute "DELETE from vendas_analise_setores where id_processo_item = " & .ListItems(InitFor)
            '==================================
            Modulo = "Outros/Análise crítica"
            Evento = "Excluir produto do processo"
            ID_documento = .ListItems.Item(InitFor)
            Documento = "Nº análise: " & Txt_analise & " - Rev.: " & Txt_rev_analise
            Documento1 = "Cód. interno: " & .ListItems(InitFor).SubItems(1) & " - Descrição: " & .ListItems(InitFor).SubItems(2)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) produto(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Produto(s) excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCampos_Processos_item True
    ProcCarregaLista_processos_item
    Novo_analise4 = False
    Frame2.Enabled = False
    With optProcessos_item
        .Value = False
        .Enabled = False
    End With
    With optProcessos_item_analise
        .Value = False
        .Enabled = False
    End With
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluir_processos()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso.")
    Exit Sub
End If
Permitido = False
With lista_Processos
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir esta(s) fase(s)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            Conexao.Execute "DELETE from vendas_analise_setores where id  = " & .ListItems(InitFor)
            '==================================
            Modulo = "Outros/Análise crítica"
            Evento = "Excluir fase"
            ID_documento = .ListItems.Item(InitFor)
            Documento = "Nº análise: " & Txt_analise & " - Rev.: " & Txt_rev_analise
            Documento1 = "Fase: " & .ListItems(InitFor).SubItems(1) & " - Máquina: " & .ListItems(InitFor).SubItems(2)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe a(s) fase(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Fase(s) excluída(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCampos_Processos
    ProcCarregaLista_processos
    Novo_analise5 = False
    Frame6.Enabled = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluir_PCP()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso.")
    Exit Sub
End If
If USMsgBox("Deseja realmente excluir o(s) cadastro(s) do PCP?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    If FunVerifStatusAnalise("excluir dados do PCP", True) = False Then Exit Sub
    If FunVerifValidSetorAnalise("excluir dados do PCP", SSTab1.Tab, True) = False Then Exit Sub
    Conexao.Execute "Update vendas_analise Set data_PCP = Null, responsavel_PCP = 'Null', Analise_PCP = 'Null' where id = " & txtId
    USMsgBox ("Dados do PCP excluídos com sucesso."), vbInformation, "CAPRIND v5.0"
    '==================================
    Modulo = "Outros/Análise crítica"
    Evento = "Excluir dados do PCP"
    ID_documento = txtId
    Documento = "Nº análise: " & Txt_analise & " - Rev.: " & Txt_rev_analise
    Documento1 = ""
    ProcGravaEvento
    '==================================
    ProcLimpaCampos_PCP
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluir_PCP_CheckList()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso.")
    Exit Sub
End If
Permitido = False
With Lista_check(1)
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) check-list do PCP?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            Conexao.Execute "DELETE from Vendas_analise_setores_checklist where id  = " & .ListItems(InitFor)
            '==================================
            Modulo = "Outros/Análise crítica"
            Evento = "Excluir check-list do PCP"
            ID_documento = .ListItems.Item(InitFor)
            Documento = "Nº análise: " & Txt_analise & " - Rev.: " & Txt_rev_analise
            Documento1 = "Descrição: " & .ListItems(InitFor).SubItems(3)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) check-list do PCP antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Check-list do PCP excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCampos_PCP_Checklist
    ProcCarregaLista_PCP_Checklist
    Novo_analise6 = False
    Frame_check(1).Enabled = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluir_Instrumento()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso.")
    Exit Sub
End If
Permitido = False
With Lista_Qualidade
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) instrumento(s)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            Conexao.Execute "DELETE from vendas_analise_setores where id  = " & .ListItems(InitFor)
            '==================================
            Modulo = "Outros/Análise crítica"
            Evento = "Excluir instrumento"
            ID_documento = txtId
            Documento1 = "Cód. interno: " & .ListItems(InitFor).SubItems(1) & " - Descrição: " & .ListItems(InitFor).SubItems(2)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) instrumento(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Instrumento(s) excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCampos_Instrumentos
    ProcCarregaLista_Instrumentos
    Novo_analise7 = False
    Frame8.Enabled = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluir_Qualidade_CheckList()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso.")
    Exit Sub
End If
Permitido = False
With Lista_check(2)
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) check-list da qualidade?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            Conexao.Execute "DELETE from Vendas_analise_setores_checklist where id  = " & .ListItems(InitFor)
            '==================================
            Modulo = "Outros/Análise crítica"
            Evento = "Excluir check-list da qualidade"
            ID_documento = txtId
            Documento = "Nº análise: " & Txt_analise & " - Rev.: " & Txt_rev_analise
            Documento1 = "Descrição: " & .ListItems(InitFor).SubItems(3)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) check-list da qualidade antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Check-list da qualidade excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCampos_Qualidade_Checklist
    ProcCarregaLista_Qualidade_Checklist
    Novo_analise8 = False
    Frame_check(2).Enabled = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluir_Compras()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso.")
    Exit Sub
End If
Permitido = False
With Lista_compras
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) valor(es)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            Conexao.Execute "Update vendas_analise_setores Set VlrUnit = NULL, VlrTotal = NULL, Responsavel_compras = NULL, Data_compras = NULL where id = " & .ListItems(InitFor)
            '==================================
            Modulo = "Outros/Análise crítica"
            Evento = "Excluir valor"
            ID_documento = txtId
            Documento = "Nº análise: " & Txt_analise & " - Rev.: " & Txt_rev_analise
            Documento1 = "Cód. interno: " & .ListItems(InitFor).SubItems(1) & " - Descrição: " & .ListItems(InitFor).SubItems(2)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) valor(es) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Valor(es) excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCampos_Compras
    ProcCarregalista_Compras
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluir_Compras_CheckList()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso.")
    Exit Sub
End If
Permitido = False
With Lista_check(3)
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) check-list de compras?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            Conexao.Execute "DELETE from Vendas_analise_setores_checklist where id  = " & .ListItems(InitFor)
            '==================================
            Modulo = "Outros/Análise crítica"
            Evento = "Excluir check-list de compras"
            ID_documento = txtId
            Documento = "Nº análise: " & Txt_analise & " - Rev.: " & Txt_rev_analise
            Documento1 = "Descrição: " & .ListItems(InitFor).SubItems(3)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) check-list de compras antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Check-list de compras excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCampos_Compras_Checklist
    ProcCarregaLista_Compras_Checklist
    Novo_analise9 = False
    Frame_check(3).Enabled = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procExcluir_doc()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso.")
    Exit Sub
End If
Permitido = False
With Lista_doc
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) documento(s)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            Conexao.Execute "DELETE from vendas_analise_setores where id  = " & .ListItems(InitFor)
            '==================================
            Modulo = "Outros/Análise crítica"
            Evento = "Excluir documento"
            ID_documento = txtId
            Documento = "Nº análise: " & Txt_analise & " - Rev.: " & Txt_rev_analise
            Documento1 = "Caminho: " & .ListItems(InitFor).SubItems(1)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) documento(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Documento(s) excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    Proclimpacampos_doc
    ProcCarregaLista_Doc
    Novo_analise10 = False
    Frame14.Enabled = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15195, 17, True
ProcCarregaToolBar2 Me, 15195, 11, False
ProcCarregaToolBar3 Me, 15195, 10, False
ProcCarregaToolBar4 Me, 15105, 5, True
ProcCarregaToolBar5 Me, 15105, 4, True

ProcCarregaToolBar6 Me, 15105, 4, True
ProcCarregaToolBar7 Me, 15195, 11, True

ProcCarregaToolBar8 Me, 15105, 4, True

ProcVerifMostrarEsconderTab "Outros/Análise crítica/Engenharia", 1
ProcVerifMostrarEsconderTab "Outros/Análise crítica/Processos", 2
ProcVerifMostrarEsconderTab "Outros/Análise crítica/Pcp", 3
ProcVerifMostrarEsconderTab "Outros/Análise crítica/Qualidade", 4
ProcVerifMostrarEsconderTab "Outros/Análise crítica/Compras", 5
ProcVerifMostrarEsconderTab "Outros/Análise crítica/Documentos", 6

Formulario = "Outros/Análise crítica"
Direitos
SSTab1.Tab = 0
SSTab_engenharia.Tab = 0
SSTab_processo.Tab = 0
SSTab_qualidade.Tab = 0
SSTab_compras.Tab = 0
ProcCarregaFamiliaUN

ProcRemoveObjetosResize Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcVerifMostrarEsconderTab(Formulario As String, NTab As Integer)
On Error GoTo tratar_erro

If NTab = 1 Then Contador = 7
Set TBAcessos = CreateObject("adodb.recordset")
TBAcessos.Open "Select Acesso from Acessos where IDUsuario = " & pubIDUsuario & " and Acesso = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBAcessos.EOF = True Then
    SSTab1.TabVisible(NTab) = False
    Contador = Contador - 1
Else
    SSTab1.TabVisible(NTab) = True
End If
TBAcessos.Close
SSTab1.TabsPerRow = Contador

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

Select Case SSTab1.Tab
    Case 0: Formulario = "Outros/Análise crítica"
    Case 1: Formulario = "Outros/Análise crítica/Engenharia"
    Case 2: Formulario = "Outros/Análise crítica/Processos"
    Case 3: Formulario = "Outros/Análise crítica/Pcp"
    Case 4: Formulario = "Outros/Análise crítica/Qualidade"
    Case 5: Formulario = "Outros/Análise crítica/Compras"
    Case 6: Formulario = "Outros/Análise crítica/Documentos"
End Select
Direitos
ProcLimpaVariaveisPrincipais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro

If txtId = "" Then
    USMsgBox ("Informe a análise crítica antes de visualizar impressão."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
NomeRel = "Vendas_analise.rpt"
ProcImprimirRel "{Vendas_analise.ID} = " & txtId, ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimparTudo()
On Error GoTo tratar_erro

Frame7.Enabled = False
Frame15.Enabled = False
Frame6.Enabled = False
Frame8.Enabled = False
Frame22.Enabled = False
Frame14.Enabled = False
ProcLimpaCampos_Engenharia_Prod
ProcLimpaCampos_Engenharia_Checklist
ProcLimpaCampos_Normas
ProcLimpaCampos_Processos_item True
ProcLimpaCampos_Processos
ProcLimpaCampos_PCP_Checklist
ProcLimpaCampos_Instrumentos
ProcLimpaCampos_Qualidade_Checklist
ProcLimpaCampos_Compras
Txt_total_materiais = "0,00"
Txt_total_terceiros = "0,00"
Txt_total_outros = "0,00"
Txt_total_ferramentas = "0,00"
Txt_total_compras = "0,00"
Proclimpacampos_doc
Lista_engenharia.ListItems.Clear
Lista_check(0).ListItems.Clear
Lista_normas.ListItems.Clear
Lista_processos_item.ListItems.Clear
lista_Processos.ListItems.Clear
Lista_check(1).ListItems.Clear
Lista_Qualidade.ListItems.Clear
Lista_check(2).ListItems.Clear
Lista_compras.ListItems.Clear
Lista_doc.ListItems.Clear
Lista_check(3).ListItems.Clear
Novo_analise1 = False
Novo_analise2 = False
Novo_analise3 = False
Novo_analise4 = False
Novo_analise5 = False
Novo_analise6 = False
Novo_analise7 = False
Novo_analise8 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcEscondeDataStatus()
On Error GoTo tratar_erro

Txt_status.Width = (Cmd_status.Left - txtResponsavel.Left) - txtResponsavel.Width
Label1(2).Left = Txt_status.Left + (Txt_status.Width / 2) - (Label1(2).Width / 2)
Txt_data_status = ""
Txt_data_status.Visible = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcMostraDataStatus()
On Error GoTo tratar_erro

Txt_status.Width = (Txt_data_status.Left - txtResponsavel.Left) - txtResponsavel.Width
Label1(2).Left = Txt_status.Left + (Txt_status.Width / 2) - (Label1(2).Width / 2)
Txt_data_status = Format(Date, "dd/mm/yy")
Txt_data_status.Visible = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

If Novo_Analise = True Then
    If USMsgBox("A análise ainda não foi salva, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        SSTab1.Tab = 0
        ProcSalvar
        If Novo_Analise = True Then Exit Sub Else Unload Me
    End If
End If
If Novo_analise1 = True Then
    If USMsgBox("O registro da engenharia ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        SSTab1.Tab = 1
        SSTab_engenharia.Tab = 0
        ProcSalvar_Engenharia_Prod
        If Novo_analise1 = True Then Exit Sub Else Unload Me
    End If
End If
If Novo_analise2 = True Then
    If USMsgBox("O check-list da engenharia ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        SSTab1.Tab = 1
        SSTab_engenharia.Tab = 1
        ProcSalvar_Engenharia_CheckList
        If Novo_analise2 = True Then Exit Sub Else Unload Me
    End If
End If
If Novo_analise3 = True Then
    If USMsgBox("A norma da engenharia ainda não foi salva, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        SSTab1.Tab = 1
        SSTab_engenharia.Tab = 2
        ProcSalvar_Engenharia_Norma
        If Novo_analise3 = True Then Exit Sub Else Unload Me
    End If
End If
If Novo_analise4 = True Then
    If USMsgBox("O produto do processo ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        SSTab1.Tab = 2
        SSTab_processo.Tab = 0
        ProcSalvar_Processos_item
        If Novo_analise4 = True Then Exit Sub Else Unload Me
    End If
End If
If Novo_analise5 = True Then
    If USMsgBox("A fase ainda não foi salva, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        SSTab1.Tab = 2
        SSTab_processo.Tab = 1
        ProcSalvar_processos
        If Novo_analise5 = True Then Exit Sub Else Unload Me
    End If
End If
If Novo_analise6 = True Then
    If USMsgBox("O check-list do PCP ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        SSTab1.Tab = 3
        ProcSalvar_PCP_CheckList
        If Novo_analise6 = True Then Exit Sub Else Unload Me
    End If
End If
If Novo_analise7 = True Then
    If USMsgBox("O instrumento ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        SSTab1.Tab = 4
        SSTab_qualidade.Tab = 0
        ProcSalvar_Instrumento
        If Novo_analise7 = True Then Exit Sub Else Unload Me
    End If
End If
If Novo_analise8 = True Then
    If USMsgBox("O check-list da qualidade ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        SSTab1.Tab = 4
        SSTab_qualidade.Tab = 1
        ProcSalvar_Qualidade_CheckList
        If Novo_analise8 = True Then Exit Sub Else Unload Me
    End If
End If
If Novo_analise9 = True Then
    If USMsgBox("O check-list de compras ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        SSTab1.Tab = 5
        SSTab_compras.Tab = 1
        ProcSalvar_Compras_CheckList
        If Novo_analise9 = True Then Exit Sub Else Unload Me
    End If
End If
If Novo_analise10 = True Then
    If USMsgBox("O documento ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        SSTab1.Tab = 6
        procSalvar_doc
        If Novo_analise10 = True Then Exit Sub Else Unload Me
    End If
End If
Novo_Analise = False
Novo_analise1 = False
Novo_analise2 = False
Novo_analise3 = False
Novo_analise4 = False
Novo_analise5 = False
Novo_analise6 = False
Novo_analise7 = False
Novo_analise8 = False
Novo_analise9 = False
Novo_analise10 = False
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
If Txt_status = "REVISADA" Then
    USMsgBox ("Não é permitida a alteração de análise crítica revisada."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Frame1.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
Acao = "salvar"
If txtdesenho = "" Then
    NomeCampo = "o código interno"
    ProcVerificaAcao
    cmdProduto_Click
    Exit Sub
End If
If txtRev_desenho = "" Then
    NomeCampo = "a revisão"
    ProcVerificaAcao
    txtRev_desenho.SetFocus
    Exit Sub
End If
If txtdescricao = "" Then
    NomeCampo = "a descrição"
    ProcVerificaAcao
    txtdescricao.SetFocus
    Exit Sub
End If
If cmbfamilia = "" Then
    NomeCampo = "a família"
    ProcVerificaAcao
    cmbfamilia.SetFocus
    Exit Sub
End If
If cmbun = "" Then
    NomeCampo = "a unidade de estoque"
    ProcVerificaAcao
    cmbun.SetFocus
    Exit Sub
End If
If Cmb_un_com = "" Then
    NomeCampo = "a unidade comercial"
    ProcVerificaAcao
    Cmb_un_com.SetFocus
    Exit Sub
End If
Qtde = IIf(Txt_qtde_sol = "", 0, Txt_qtde_sol)
If Qtde <= 0 Then
    NomeCampo = "a quantidade solicitada"
    ProcVerificaAcao
    Txt_qtde_sol.SetFocus
    Exit Sub
End If
Qtde = IIf(txtQtde = "", 0, txtQtde)
If Qtde <= 0 Then
    NomeCampo = "a quantidade do lote"
    ProcVerificaAcao
    txtQtde.SetFocus
    Exit Sub
End If
If txtCliente = "" Then
    NomeCampo = "o cliente"
    ProcVerificaAcao
    cmdcliente_Click
    Exit Sub
End If
If Txt_status = "APROVADA" Then
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from vendas_analise_setores where IDanalise = " & txtId & " and Setor = 'NORMAS'", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Do While TBAbrir.EOF = False
            Set TBFIltro = CreateObject("adodb.recordset")
            TBFIltro.Open "Select * from Norma where norma = '" & TBAbrir!Texto & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBFIltro.EOF = True Then
                USMsgBox ("Não é permitido aprovar, pois a empresa não possui todas as normas."), vbExclamation, "CAPRIND v5.0"
                Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open "Select * from vendas_analise where id = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    ProcLimpaCampos
                    ProcPuxaDados
                End If
                TBAbrir.Close
                TBFIltro.Close
                frmVendas_analise_normas.Show 1
                Exit Sub
            End If
            TBFIltro.Close
            TBAbrir.MoveNext
        Loop
    End If
    TBAbrir.Close
End If
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from vendas_analise where ID = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = False Then
    If TBGravar!status = "APROVADA" And Txt_status = "APROVADA" Then
        USMsgBox ("Não é permitido alterar esta análise, pois a mesma já está aprovada."), vbExclamation, "CAPRIND v5.0"
        TBGravar.Close
        Exit Sub
    End If
    If TBGravar!status = "APROVADA" And Txt_status <> "APROVADA" Then
        Set TBFIltro = CreateObject("adodb.recordset")
        TBFIltro.Open "Select Codigo from vendas_carteira where IDAnalise = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
        If TBFIltro.EOF = False Then
            USMsgBox ("Não é permitido alterar esta análise, pois a mesma está vinculada a uma proposta."), vbExclamation, "CAPRIND v5.0"
            TBFIltro.Close
            Exit Sub
        End If
        TBFIltro.Close
    End If
    
    If TBGravar!Qtde <> txtQtde Then
        'Altera cadastro item do preocesso
        Qtde = txtQtde
        If txtidproduto <> "" Then
            Set TBItem = CreateObject("adodb.recordset")
            TBItem.Open "Select * from projproduto where codproduto = " & txtidproduto, Conexao, adOpenKeyset, adLockOptimistic
            If TBItem.EOF = False Then
                Conexao.Execute "Update Vendas_analise_ProdutosProcessos Set codProduto = " & TBItem!Codproduto & ", Codinterno = '" & TBItem!Desenho & "', Descricao = '" & TBItem!Descricao & "', Familia = '" & TBItem!Classe & "', un = '" & TBItem!Unidade & "', Qtde = '" & Qtde & "' where id_analise = " & txtId & " and Produto_analise = 'True'"
            Else
                Conexao.Execute "Update Vendas_analise_ProdutosProcessos Set codProduto = " & txtidproduto & ", Codinterno = '" & txtdesenho & "', Descricao = '" & txtdescricao & "', Familia = '" & cmbfamilia & "', un = '" & cmbun & "', Qtde = '" & Qtde & "' where id_analise = " & txtId & " and Produto_analise = 'True'"
            End If
            TBItem.Close
            
            txtQtde_processos_item = Format(txtQtde, "###,##0.0000")
        End If
        
        'Calcula custo de preparação diluido na quantidade
        Set TBProcessos = CreateObject("adodb.recordset")
        TBProcessos.Open "select * from vendas_analise_ProdutosProcessos where codinterno = '" & txtdesenho & "' and Produto_analise = 'True'", Conexao, adOpenKeyset, adLockOptimistic
        If TBProcessos.EOF = False Then
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "select * from vendas_analise_setores where idanalise = " & txtId & " and id_processo_item = " & TBProcessos!ID & " and setor = 'PROCESSOS'", Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                Do While TBAbrir.EOF = False
                    Qtde = 0
                    valor = 0
                    qt = 0
                    Qtd = 0
                    ValorTotal = 0
                    quantidade = 0
                    
                    'Calcula valor de execução
                    Qtde = IIf(IsNull(TBAbrir!Qtde), 0, TBAbrir!Qtde) 'Qtde. execução em seg
                    valor = IIf(IsNull(TBAbrir!VlrUnit), 0, TBAbrir!VlrUnit)
                    ValorTotal = Qtde * valor
                                        
                    'Calcula valor de execução
                    'Converte tempo de preparação em seg
                    dataCalculo = IIf(IsNull(TBAbrir!Preparacao), 0, TBAbrir!Preparacao)
                    ProcFormataHora (dataCalculo)
                    qt = s / 3600
                    
                    If IsNull(TBAbrir!PrecoHora_Setup) = False And TBAbrir!PrecoHora_Setup <> "" Then valor = TBAbrir!PrecoHora_Setup Else valor = IIf(IsNull(TBAbrir!VlrUnit), 0, TBAbrir!VlrUnit)  'Preço por hr preparação
                    Qtd = IIf(IsNull(TBProcessos!Qtde), 0, TBProcessos!Qtde)
                    'Dilui o valor hora de preparação pela quantidade do lote
                    valor = valor / Qtd
                    
                    If qt > 0 Then
                        quantidade = qt * valor
                        ValorTotal = ValorTotal + quantidade
                    End If
                    TBAbrir!vlrTotal = Format(ValorTotal, "###,##0.00")
                    TBAbrir.Update
                    TBAbrir.MoveNext
                Loop
            End If
            TBAbrir.Close
        End If
        TBProcessos.Close
    End If
Else
    TBGravar.AddNew
    Set TBCotacao = CreateObject("adodb.recordset")
    TBCotacao.Open "Select Nanalise from Vendas_analise where Year(data) = '" & Year(Date) & "' order by Ordenaranalise desc", Conexao, adOpenKeyset, adLockOptimistic
    If TBCotacao.EOF = False Then Cotacao = Left(TBCotacao!Nanalise, Len(TBCotacao!Nanalise) - 3) + 1 Else Cotacao = 1
    Ano = Right(Year(Date), 2)
    Select Case Len(Cotacao)
        Case 1: NumeroAnalise = "000" & Cotacao & "/" & Ano
        Case 2: NumeroAnalise = "00" & Cotacao & "/" & Ano
        Case 3: NumeroAnalise = "0" & Cotacao & "/" & Ano
        Case 4: NumeroAnalise = Cotacao & "/" & Ano
        Case 5: NumeroAnalise = Cotacao & "/" & Ano
    End Select
    Txt_analise = NumeroAnalise
    TBGravar!Nanalise = NumeroAnalise
    TBGravar!Data = Date
    TBGravar!Responsavel = pubUsuario
    TBGravar!status = Txt_status
    TBGravar!Revisao = 0
End If
ProcEnviaDados
TBGravar.Update
txtId = TBGravar!ID
If Novo_Analise = True Then
    Conexao.Execute "Update Vendas_analise set ordenaranalise = " & TBGravar!ID & " where ID = " & TBGravar!ID
    
    'Grava os check-list padrões
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from Vendas_analise_descricao_checklist where Padrao = 'True'", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Do While TBAbrir.EOF = False
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select * from Vendas_analise_setores_checklist", Conexao, adOpenKeyset, adLockOptimistic
            TBFI.AddNew
            TBFI!IDAnalise = txtId
            TBFI!IDchecklist = TBAbrir!ID
            TBFI!Setor = TBAbrir!Setor
            TBFI!Data = Date
            TBFI!Responsavel = pubUsuario
            TBFI!Sim = True
            TBFI.Update
            TBFI.Close
            TBAbrir.MoveNext
        Loop
    End If
End If
TBGravar.Close
Lista.ListItems.Clear

If Novo_Analise = True Then
    USMsgBox ("Nova análise crítica cadastrada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo"
    StrSql_AnaliseCritica = "Select * from vendas_analise where ID = " & txtId
    ProcCarregaLista (1)
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar"
    ProcCarregaLista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
    If CodigoLista <> 0 And Lista.ListItems.Count <> 0 Then
        Lista.SelectedItem = Lista.ListItems(CodigoLista)
        Lista.SetFocus
    End If
End If
'==================================
Modulo = "Outros/Análise crítica"
ID_documento = txtId.Text
Documento = "Nº análise: " & Txt_analise & " - Rev.: " & Txt_rev_analise
Documento1 = ""
ProcGravaEvento
'==================================
Novo_Analise = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvar_Engenharia()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If FunVerifStatusAnalise("alterar os dados da engenharia", True) = False Then Exit Sub
If FunVerifValidSetorAnalise("alterar os dados da engenharia", SSTab1.Tab, True) = False Then Exit Sub
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from vendas_analise where ID = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = False Then
    TBGravar!Data_Engenharia = Date
    TBGravar!Responsavel_Engenharia = pubUsuario
    TBGravar!obs_engenharia = Txt_obs_engenharia
    TBGravar.Update
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
End If
TBGravar.Close
'==================================
Modulo = "Outros/Análise crítica"
Evento = "Alterar dados da engenharia"
ID_documento = txtId
Documento = "Nº análise: " & Txt_analise & " - Rev.: " & Txt_rev_analise
Documento1 = ""
ProcGravaEvento
'==================================
txtData_engenharia = Format(Date, "dd/mm/yy")
txtResponsavel_engenharia = pubUsuario

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvar_Engenharia_Prod()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Frame7.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
Acao = "salvar"
If Opt_material.Value = False And Opt_terceiros.Value = False And Opt_outros.Value = False Then
    NomeCampo = "o tipo"
    ProcVerificaAcao
    Exit Sub
End If
If chkAuto_engenharia.Value = 0 And txtdesenho_engenharia = "" Then
    NomeCampo = "o código interno"
    ProcVerificaAcao
    cmdProduto_engenharia_Click
    Exit Sub
End If
If cmbun_engenharia = "" Then
    NomeCampo = "a unidade de estoque"
    ProcVerificaAcao
    cmbun_engenharia.SetFocus
    Exit Sub
End If
If Cmb_un_com_engenharia = "" Then
    NomeCampo = "a unidade comercial"
    ProcVerificaAcao
    Cmb_un_com_engenharia.SetFocus
    Exit Sub
End If
Qtde = IIf(txtQtde_engenharia = "", 0, txtQtde_engenharia)
If Qtde <= 0 Then
    NomeCampo = "a quantidade"
    ProcVerificaAcao
    txtQtde_engenharia.SetFocus
    Exit Sub
End If
If txtdescricao_engenharia = "" Then
    NomeCampo = "a descrição"
    ProcVerificaAcao
    txtdescricao_engenharia.SetFocus
    Exit Sub
End If
If cmbfamilia_engenharia = "" Then
    NomeCampo = "a família"
    ProcVerificaAcao
    cmbfamilia_engenharia.SetFocus
    Exit Sub
End If
If chkAuto_engenharia.Value = 1 Then
    ProcNovoProdutoAutoEngQualidade
    If txtReferencia_engenharia <> "" Then
        cmbReferencia_engenharia.AddItem txtReferencia_engenharia
        cmbReferencia_engenharia = txtReferencia_engenharia
    End If
    chkAuto_engenharia.Value = 0
End If
If chkManual_engenharia.Value = 1 Then
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select Codproduto from projproduto where desenho = '" & txtdesenho_engenharia & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
        USMsgBox ("Já existe um produto cadastrado com este código interno, favor alterar."), vbExclamation, "CAPRIND v5.0"
        txtdesenho_engenharia.SetFocus
        Exit Sub
    End If
    TBProduto.Close
    ProcNovoProdutoManualEngQualidade
    If txtReferencia_engenharia <> "" Then
        cmbReferencia_engenharia.AddItem txtReferencia_engenharia
        cmbReferencia_engenharia = txtReferencia_engenharia
    End If
    chkManual_engenharia.Value = 0
End If

Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from vendas_analise_setores where ID = " & txtID_engenharia, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then
    TBGravar.AddNew
    TBGravar!IDAnalise = txtId
    TBGravar!Data = Date
    TBGravar!Responsavel = pubUsuario
    USMsgBox ("Novo registro cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo registro"
Else
    If FunVerifStatusAnalise("alterar os dados desse registro", True) = False Then Exit Sub
    If FunVerifValidSetorAnalise("alterar os dados desse registro", SSTab1.Tab, True) = False Then Exit Sub
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar registro"
End If
ProcEnviadados_Engenharia_Prod
TBGravar.Update
txtID_engenharia = TBGravar!ID
TBGravar.Close
ProcCarregaLista_Engenharia_Prod
If Novo_analise1 = False Then
    If CodigoLista1 <> 0 And Lista_engenharia.ListItems.Count <> 0 Then
        Lista_engenharia.SelectedItem = Lista_engenharia.ListItems(CodigoLista1)
        Lista_engenharia.SetFocus
    End If
End If
'==================================
Modulo = "Outros/Análise crítica"
ID_documento = txtID_engenharia
Documento = "Nº análise: " & Txt_analise & " - Rev.: " & Txt_rev_analise
Documento1 = "Cód. interno: " & txtdesenho_engenharia & " - Descrição: " & txtdescricao_engenharia
ProcGravaEvento
'==================================
Novo_analise1 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvar_Engenharia_CheckList()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Frame_check(0).Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
Acao = "salvar"
If Txt_descricao_chek(0) = "" Then
    NomeCampo = "a descrição do check-list"
    ProcVerificaAcao
    ProcAbrirDescCheckList 0
    Exit Sub
End If
If Chk_nao_chek(0).Value = 0 And Chk_sim_chek(0).Value = 0 Then
    NomeCampo = "uma das opções"
    ProcVerificaAcao
    Exit Sub
End If
If Chk_sim_chek(0).Value = 1 And Txt_texto_check(0) = "" Then
    NomeCampo = "quais"
    ProcVerificaAcao
    Txt_texto_check(0).SetFocus
    Exit Sub
End If
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from Vendas_analise_setores_checklist where ID = " & Txt_ID_check(0), Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then
    TBGravar.AddNew
    TBGravar!IDAnalise = txtId
    TBGravar!Data = Date
    TBGravar!Responsavel = pubUsuario
    TBGravar!Setor = "ENGENHARIA"
    USMsgBox ("Novo check-list da engenharia cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo check-list da engenharia"
Else
    If FunVerifStatusAnalise("alterar o check-list da engenharia", True) = False Then Exit Sub
    If FunVerifValidSetorAnalise("alterar o check-list da engenharia", SSTab1.Tab, True) = False Then Exit Sub
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar check-list da engenharia"
End If
ProcEnviadados_Engenharia_Checklist
TBGravar.Update
Txt_ID_check(0) = TBGravar!ID
TBGravar.Close
ProcCarregaLista_Engenharia_Checklist
If Novo_analise2 = False Then
    If CodigoLista2 <> 0 And Lista_check(0).ListItems.Count <> 0 Then
        Lista_check(0).SelectedItem = Lista_check(0).ListItems(CodigoLista2)
        Lista_check(0).SetFocus
    End If
End If
'==================================
Modulo = "Outros/Análise crítica"
ID_documento = Txt_ID_check(0)
Documento = "Nº análise: " & Txt_analise & " - Rev.: " & Txt_rev_analise
Documento1 = "Descrição: " & Txt_descricao_chek(0)
ProcGravaEvento
'==================================
Novo_analise2 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvar_Engenharia_Norma()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Frame15.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
Acao = "salvar"
If Txt_norma = "" Then
    NomeCampo = "a norma"
    ProcVerificaAcao
    Txt_norma.SetFocus
    Exit Sub
End If
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from vendas_analise_setores where ID = " & Txt_ID_norma, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then
    TBGravar.AddNew
    TBGravar!IDAnalise = txtId
    TBGravar!Responsavel = pubUsuario
    TBGravar!Data = Date
    TBGravar!Setor = "NORMAS"
    USMsgBox ("Nova norma cadastrada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Nova norma"
Else
    If FunVerifStatusAnalise("alterar a norma", True) = False Then Exit Sub
    If FunVerifValidSetorAnalise("alterar a norma", SSTab1.Tab, True) = False Then Exit Sub
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar norma"
End If
TBGravar!Texto = Txt_norma
TBGravar!Analise = Txt_obs_norma
TBGravar.Update
Txt_ID_norma = TBGravar!ID
TBGravar.Close
ProcCarregaLista_Engenharia_Normas
If Novo_analise3 = False Then
    If CodigoLista3 <> 0 And Lista_normas.ListItems.Count <> 0 Then
        Lista_normas.SelectedItem = Lista_normas.ListItems(CodigoLista3)
        Lista_normas.SetFocus
    End If
End If
'==================================
Modulo = "Outros/Análise crítica"
ID_documento = Txt_ID_norma
Documento = "Nº análise: " & Txt_analise & " - Rev.: " & Txt_rev_analise
Documento1 = "Norma: " & Txt_norma
ProcGravaEvento
'==================================
Novo_analise3 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvar_Processos_item()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Frame2.Enabled = False And optProcessos_item.Value = True Then
    ProcVerificaSalvar
    Exit Sub
End If
Acao = "salvar"
If optProcessos_item_analise.Value = False And optProcessos_item.Value = False Then
    NomeCampo = "a origem do produto do processo"
    ProcVerificaAcao
    Exit Sub
End If
If txtCodInterno_processos_item = "" Then
    NomeCampo = "o código interno"
    ProcVerificaAcao
    txtCodInterno_processos_item.SetFocus
    Exit Sub
End If
If txtDescricao_processos_item = "" Then
    NomeCampo = "a descrição"
    ProcVerificaAcao
    txtDescricao_processos_item.SetFocus
    Exit Sub
End If
If cmbFamilia_processos_item = "" Then
    NomeCampo = "a familía"
    ProcVerificaAcao
    cmbFamilia_processos_item.SetFocus
    Exit Sub
End If
If cmbUn_processos_item = "" Then
    NomeCampo = "a unidade de estoque"
    ProcVerificaAcao
    cmbUn_processos_item.SetFocus
    Exit Sub
End If
If Cmb_un_com_processos_item = "" Then
    NomeCampo = "a unidade comercial"
    ProcVerificaAcao
    Cmb_un_com_processos_item.SetFocus
    Exit Sub
End If
Qtde = IIf(txtQtde_processos_item = "", 0, txtQtde_processos_item)
If Qtde <= 0 Then
    NomeCampo = "a quantidade"
    ProcVerificaAcao
    txtQtde_processos_item.SetFocus
    Exit Sub
End If

'Verifica se já foi adicionado o produto principal da analise
If optProcessos_item_analise.Value = True Then
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from vendas_analise_ProdutosProcessos where Produto_analise = 'True' and id_analise = " & txtId & " and ID <> " & txtID_processos_item, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        USMsgBox ("Não é permitido salvar, pois já existe um produto da análise cadastrado."), vbExclamation, "CAPRIND v5.0"
        TBAbrir.Close
        Exit Sub
    End If
End If

'Verifica se já foi adicionado produto com este cod. interno
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from vendas_analise_ProdutosProcessos where codinterno = '" & txtCodInterno_processos_item & "' and id_analise = " & txtId & " and id <> " & txtID_processos_item, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    USMsgBox ("Não é permitido salvar, pois já existe este produto cadastrado para esta análise crítica."), vbExclamation, "CAPRIND v5.0"
    TBAbrir.Close
    Exit Sub
End If
TBAbrir.Close

Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from vendas_analise_ProdutosProcessos where ID = " & txtID_processos_item, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then
    TBGravar.AddNew
    USMsgBox ("Novo produto do processo cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo produto do processo"
Else
    If FunVerifStatusAnalise("alterar o produto do processo", True) = False Then Exit Sub
    If FunVerifValidSetorAnalise("alterar o produto do processo", SSTab1.Tab, True) = False Then Exit Sub
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar produto do processo"
End If

ProcEnviadados_Processos_item
TBGravar.Update
txtID_processos_item = TBGravar!ID
Codinterno = TBGravar!Codinterno

If txtIDproduto_processos <> "" And txtIDproduto_processos <> "0" And (Novo_analise4 = True Or Novo_analise4 = False And txtCodInterno_processos_item <> Codinterno) Then
    'Verifica se existe processo para o produto
    Permitido = True
    Permitido1 = True
    Set TBProcessos = CreateObject("adodb.recordset")
    TBProcessos.Open "Select Fases.* from Fases INNER JOIN Processos ON Fases.IDProcesso = Processos.IDProcesso where Processos.codproduto = " & IIf(txtIDproduto_processos = "", 0, txtIDproduto_processos), Conexao, adOpenKeyset, adLockOptimistic
    If TBProcessos.EOF = False Then
        Do While TBProcessos.EOF = False
            If Permitido = True Then
                If USMsgBox("Já existe processo cadastrado para este produto, deseja aproveitar esse processo?", vbYesNo, "CAPRIND v5.0") = vbNo Then GoTo Pula
            End If
            If Permitido1 = True Then Conexao.Execute "DELETE from vendas_analise_setores where id_processo_item = " & txtID_processos_item & " and Setor = 'PROCESSOS'"
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select * from vendas_analise_setores", Conexao, adOpenKeyset, adLockOptimistic
            TBFI.AddNew
            TBFI!IDAnalise = txtId
            TBFI!Responsavel = pubUsuario
            TBFI!Data = Date
            TBFI!Texto = IIf(IsNull(TBProcessos!maquina), "", TBProcessos!maquina)
            TBFI!ID_processo_item = txtID_processos_item
            
            'Verifica custo hora do posto
            Set TBMaquinas = CreateObject("adodb.recordset")
            TBMaquinas.Open "select PrecoHora, PrecoHora_Setup, Descricao from cadmaquinas where maquina = '" & TBProcessos!maquina & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBMaquinas.EOF = False Then
                TBFI!VlrUnit = IIf(IsNull(TBMaquinas!PrecoHora), "", Format(TBMaquinas!PrecoHora, "###,##0.00"))
                TBFI!PrecoHora_Setup = IIf(IsNull(TBMaquinas!PrecoHora_Setup), "", Format(TBMaquinas!PrecoHora_Setup, "###,##0.00"))
                TBFI!Descricao = IIf(IsNull(TBMaquinas!Descricao), "", TBMaquinas!Descricao)
            End If
            TBMaquinas.Close
            
            If TBProcessos!Execucao > "23:59:59" Then
                ProcFormataHora (TBProcessos!Execucao)
                Familiatext = DataResultado
                TBFI!TotalHora = FunCalculaSegPC(Familiatext, IIf(IsNull(TBProcessos!pc_te), 0, TBProcessos!pc_te))
            Else
                TotalGeral = FunCalculaSegPC(TBProcessos!Execucao, IIf(IsNull(TBProcessos!pc_te), 0, TBProcessos!pc_te))
                Texto = FormataTempo(TotalGeral)
                TBFI!TotalHora = Texto
            End If
            ProcFormataHora (TBFI!TotalHora)
            TBFI!Qtde = (s + DecimoSegundos) / 3600
            
'            Qtde = 0
'            Valor = 0
'            qt = 0
'            Qtd = 0
'            ValorTotal = 0
'            Quantidade = 0
'            quantnovo = 0
'
'            NovoValor = Replace(TBFI!TotalHora, ",", ".")
'            If TBFI!TotalHora <> "" Then
'                ProcFormataHora (TBFI!TotalHora)
'                HoraResultado = DataResultado
'                ElapsedTime (HoraResultado)
'                Qtde = (S + DecimoSegundos) / 3600
'            Else
'                Qtde = 0
'            End If
'            Valor = TBFI!VlrUnit
'
'            If Len(TBProcessos!Execucao) < 7 Then
'                qt = 0
'            Else
'                HoraResultado = TBProcessos!Execucao
'                ElapsedTime (HoraResultado)
'                qt = S / 3600
'            End If
'            quantnovo = 0
'
'            Qtd = IIf(txtQtde_processos_item = "", 0, txtQtde_processos_item)
'            ValorTotal = Qtde * Valor
'            If qt > 0 Then
'                Quantidade = (qt / Qtd) * Valor
'                ValorTotal = ValorTotal + Quantidade
'            End If
'            TBFI!VlrTotal = Format(ValorTotal, "###,##0.00")
            
            Qtd = IIf(txtQtde_processos_item = "", 0, txtQtde_processos_item)
            TBFI!vlrTotal = Format(IIf(IsNull(TBProcessos!Custo), 0, TBProcessos!Custo) + (IIf(IsNull(TBProcessos!Custoprep), 0, TBProcessos!Custoprep) / Qtd), "###,##0.00")
            
            TBFI!Peca = IIf(IsNull(TBProcessos!pc_te), 0, TBProcessos!pc_te)
            TBFI!Execucao = IIf(IsNull(TBProcessos!Execucao), Null, TBProcessos!Execucao)
            TBFI!Preparacao = IIf(IsNull(TBProcessos!Preparacao), Null, TBProcessos!Preparacao)
            TBFI!Trabalho = IIf(IsNull(TBProcessos!Descricao), Null, TBProcessos!Descricao)
            TBFI!Fase = IIf(IsNull(TBProcessos!Fase), Null, TBProcessos!Fase)
            TBFI!Grupo_op = IIf(IsNull(TBProcessos!Grupo_op), Null, TBProcessos!Grupo_op)
            TBFI!pecahora = TBProcessos!pecahora
            TBFI!Erro_processos = 0
            TBFI!Setor = "PROCESSOS"
            TBFI.Update
            TBFI.Close
            Permitido = False
            Permitido1 = False
            TBProcessos.MoveNext
        Loop
    End If
    TBProcessos.Close
End If
Pula:

TBGravar.Close
ProcCarregaLista_processos_item
If Novo_analise4 = False Then
    If CodigoLista4 <> 0 And Lista_processos_item.ListItems.Count <> 0 Then
        Lista_processos_item.SelectedItem = Lista_processos_item.ListItems(CodigoLista4)
        Lista_processos_item.SetFocus
    End If
End If
'==================================
Modulo = "Outros/Análise crítica"
ID_documento = txtID_processos_item
Documento = "Nº análise: " & Txt_analise & " - Rev.: " & Txt_rev_analise
Documento1 = "Cód. interno: " & txtCodInterno_processos_item & " - Descrição: " & txtDescricao_processos_item
ProcGravaEvento
'==================================
Novo_analise4 = False
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvar_processos()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Frame6.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
Acao = "salvar"
If txtMaquina_processos = "" Then
    NomeCampo = "a máquina"
    ProcVerificaAcao
    txtMaquina_processos.SetFocus
    Exit Sub
End If
If txtPecaHora_processos = "" Then
    NomeCampo = "a quantidade de peças por hora"
    ProcVerificaAcao
    txtPecaHora_processos.SetFocus
    Exit Sub
End If
txtPreparacao_processos.PromptInclude = False
If Len(txtPreparacao_processos.Text) < 7 Then
    txtPreparacao_processos.PromptInclude = True
    USMsgBox "Verifique se faltam dados no campo preparação ( " & txtPreparacao_processos.Text & " ) á serem preenchidos.", vbExclamation, "CAPRIND v5.0"
    txtPreparacao_processos.SetFocus
    Exit Sub
End If
txtPreparacao_processos.PromptInclude = True
txtExecucao_processos.PromptInclude = False
If Len(txtExecucao_processos.Text) < 7 Then
    txtExecucao_processos.PromptInclude = True
    USMsgBox "Verifique se faltam dados no campo execução ( " & txtExecucao_processos.Text & " ) á serem preenchidos.", vbExclamation, "CAPRIND v5.0"
    txtExecucao_processos.SetFocus
    Exit Sub
End If
txtExecucao_processos.PromptInclude = True
If txtFase = "" Then
    NomeCampo = "a fase"
    ProcVerificaAcao
    txtFase.SetFocus
    Exit Sub
End If
If txtTrabalho = "" Then
    NomeCampo = "a instruções de trabalho"
    ProcVerificaAcao
    txtTrabalho.SetFocus
    Exit Sub
End If
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from vendas_analise_setores where ID = " & txtID_processos, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then
    TBGravar.AddNew
    TBGravar!IDAnalise = txtId
    TBGravar!Data = Date
    TBGravar!Responsavel = pubUsuario
    USMsgBox ("Nova fase cadastrada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo fase do processo"
Else
    If FunVerifStatusAnalise("alterar a fase do processo", True) = False Then Exit Sub
    If FunVerifValidSetorAnalise("alterar a fase do processo", SSTab1.Tab, True) = False Then Exit Sub
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar processo"
End If
ProcEnviadados_processos
TBGravar.Update
txtID_processos = TBGravar!ID
TBGravar.Close
ProcCarregaLista_processos
If Novo_analise5 = False Then
    If CodigoLista5 <> 0 Then
        If lista_Processos.ListItems.Count = 0 Then Exit Sub
        lista_Processos.SelectedItem = lista_Processos.ListItems(CodigoLista5)
        lista_Processos.SetFocus
    End If
End If
'==================================
Modulo = "Outros/Análise crítica"
ID_documento = txtID_processos
Documento = "Nº análise: " & Txt_analise & " - Rev.: " & Txt_rev_analise
Documento1 = "Fase: " & txtFase & " - Máquina: " & txtMaquina_processos
ProcGravaEvento
'==================================
Novo_analise5 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvar_PCP()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If FunVerifStatusAnalise("alterar os dados do PCP", True) = False Then Exit Sub
If FunVerifValidSetorAnalise("alterar os dados do PCP", SSTab1.Tab, True) = False Then Exit Sub
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from vendas_analise where ID = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = False Then
    TBGravar!data_PCP = Date
    TBGravar!responsavel_PCP = pubUsuario
    TBGravar!Analise_PCP = txtAnalise_PCP
    TBGravar.Update
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
End If
TBGravar.Close
'==================================
Modulo = "Outros/Análise crítica"
Evento = "Alterar dados do PCP"
ID_documento = txtId
Documento = "Nº análise: " & Txt_analise & " - Rev.: " & Txt_rev_analise
Documento1 = ""
ProcGravaEvento
'==================================
txtData_PCP = Format(Date, "dd/mm/yy")
txtResponsavel_PCP = pubUsuario

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvar_PCP_CheckList()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Frame_check(1).Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
Acao = "salvar"
If Txt_descricao_chek(1) = "" Then
    NomeCampo = "a descrição do check-list"
    ProcVerificaAcao
    ProcAbrirDescCheckList 1
    Exit Sub
End If
If Chk_nao_chek(1).Value = 0 And Chk_sim_chek(1).Value = 0 Then
    NomeCampo = "uma das opções"
    ProcVerificaAcao
    Exit Sub
End If
If Chk_sim_chek(1).Value = 1 And Txt_texto_check(1) = "" Then
    NomeCampo = "quais"
    ProcVerificaAcao
    Txt_texto_check(1).SetFocus
    Exit Sub
End If
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from Vendas_analise_setores_checklist where ID = " & Txt_ID_check(1), Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then
    TBGravar.AddNew
    TBGravar!IDAnalise = txtId
    TBGravar!Data = Date
    TBGravar!Responsavel = pubUsuario
    TBGravar!Setor = "PCP"
    USMsgBox ("Novo check-list do PCP cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo check-list do PCP"
Else
    If FunVerifStatusAnalise("alterar o check-list do PCP", True) = False Then Exit Sub
    If FunVerifValidSetorAnalise("alterar o check-list do PCP", SSTab1.Tab, True) = False Then Exit Sub
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar check-list do PCP"
End If
ProcEnviadados_PCP_Checklist
TBGravar.Update
Txt_ID_check(1) = TBGravar!ID
TBGravar.Close
ProcCarregaLista_PCP_Checklist
If Novo_analise6 = False Then
    If CodigoLista6 <> 0 And Lista_check(1).ListItems.Count <> 0 Then
        Lista_check(1).SelectedItem = Lista_check(1).ListItems(CodigoLista6)
        Lista_check(1).SetFocus
    End If
End If
'==================================
Modulo = "Outros/Análise crítica"
ID_documento = Txt_ID_check(1)
Documento = "Nº análise: " & Txt_analise & " - Rev.: " & Txt_rev_analise
Documento1 = "Descrição: " & Txt_descricao_chek(1)
ProcGravaEvento
'==================================
Novo_analise6 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvar_Instrumento()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Frame8.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
Acao = "salvar"
If chkAuto_qualidade.Value = 0 And txtdesenho_qualidade = "" Then
    NomeCampo = "o código interno"
    ProcVerificaAcao
    cmdProduto_qualidade_Click
    Exit Sub
End If
If cmbun_qualidade = "" Then
    NomeCampo = "a unidade de estoque"
    ProcVerificaAcao
    cmbun_qualidade.SetFocus
    Exit Sub
End If
If Cmb_un_com_qualidade = "" Then
    NomeCampo = "a unidade comercial"
    ProcVerificaAcao
    Cmb_un_com_qualidade.SetFocus
    Exit Sub
End If
Qtde = IIf(txtQtde_Qualidade = "", 0, txtQtde_Qualidade)
If Qtde <= 0 Then
    NomeCampo = "a quantidade"
    ProcVerificaAcao
    txtQtde_Qualidade.SetFocus
    Exit Sub
End If
If txtdescricao_Qualidade = "" Then
    NomeCampo = "a descrição"
    ProcVerificaAcao
    txtdescricao_Qualidade.SetFocus
    Exit Sub
End If
If cmbfamilia_qualidade = "" Then
    NomeCampo = "a família"
    ProcVerificaAcao
    cmbfamilia_qualidade.SetFocus
    Exit Sub
End If
If chkAuto_engenharia.Value = 1 Then ProcNovoProdutoAutoEngQualidade
If chkManual_engenharia.Value = 1 Then
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select * from projproduto where desenho = '" & txtdesenho_qualidade & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
        USMsgBox ("Já existe um produto cadastrado com este código interno, favor alterar."), vbExclamation, "CAPRIND v5.0"
        txtdesenho_qualidade.SetFocus
        Exit Sub
    End If
    TBProduto.Close
    ProcNovoProdutoManualEngQualidade
End If
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from vendas_analise_setores where ID = " & txtID_qualidade, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then
    TBGravar.AddNew
    TBGravar!IDAnalise = txtId
    TBGravar!Responsavel = pubUsuario
    TBGravar!Data = Date
    USMsgBox ("Novo instrumento cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo instrumento"
Else
    If FunVerifStatusAnalise("alterar o instrumento", True) = False Then Exit Sub
    If FunVerifValidSetorAnalise("alterar o instrumento", SSTab1.Tab, True) = False Then Exit Sub
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar instrumento"
End If
ProcEnviadados_Instrumento
TBGravar.Update
txtID_qualidade = TBGravar!ID
TBGravar.Close
ProcCarregaLista_Instrumentos
If Novo_analise7 = False Then
    If CodigoLista7 <> 0 And Lista_Qualidade.ListItems.Count <> 0 Then
        Lista_Qualidade.SelectedItem = Lista_Qualidade.ListItems(CodigoLista7)
        Lista_Qualidade.SetFocus
    End If
End If
'==================================
Modulo = "Outros/Análise crítica"
ID_documento = txtID_qualidade
Documento = "Nº análise: " & Txt_analise & " - Rev.: " & Txt_rev_analise
Documento1 = "Cód. interno: " & txtdesenho_qualidade & " - Descrição: " & txtdescricao_Qualidade
ProcGravaEvento
'==================================
Novo_analise7 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvar_Qualidade_CheckList()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Frame_check(2).Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
Acao = "salvar"
If Txt_descricao_chek(2) = "" Then
    NomeCampo = "a descrição do check-list"
    ProcVerificaAcao
    ProcAbrirDescCheckList 2
    Exit Sub
End If
If Chk_nao_chek(2).Value = 0 And Chk_sim_chek(2).Value = 0 Then
    NomeCampo = "uma das opções"
    ProcVerificaAcao
    Exit Sub
End If
If Chk_sim_chek(2).Value = 1 And Txt_texto_check(2) = "" Then
    NomeCampo = "quais"
    ProcVerificaAcao
    Txt_texto_check(2).SetFocus
    Exit Sub
End If
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from Vendas_analise_setores_checklist where ID = " & Txt_ID_check(2), Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then
    TBGravar.AddNew
    TBGravar!IDAnalise = txtId
    TBGravar!Data = Date
    TBGravar!Responsavel = pubUsuario
    TBGravar!Setor = "QUALIDADE"
    USMsgBox ("Novo check-list da qualidade cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo check-list da qualidade"
Else
    If FunVerifStatusAnalise("alterar o check-list da qualidade", True) = False Then Exit Sub
    If FunVerifValidSetorAnalise("alterar o check-list da qualidade", SSTab1.Tab, True) = False Then Exit Sub
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar check-list da qualidade"
End If
ProcEnviadados_Qualidade_Checklist
TBGravar.Update
Txt_ID_check(2) = TBGravar!ID
TBGravar.Close
ProcCarregaLista_Qualidade_Checklist
If Novo_analise8 = False Then
    If CodigoLista8 <> 0 And Lista_check(2).ListItems.Count <> 0 Then
        Lista_check(2).SelectedItem = Lista_check(2).ListItems(CodigoLista8)
        Lista_check(2).SetFocus
    End If
End If
'==================================
Modulo = "Outros/Análise crítica"
ID_documento = Txt_ID_check(2)
Documento = "Nº análise: " & Txt_analise & " - Rev.: " & Txt_rev_analise
Documento1 = "Descrição: " & Txt_descricao_chek(2)
ProcGravaEvento
'==================================
Novo_analise8 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvar_Compras()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If txtValor_Compras = "" Then
    USMsgBox ("Informe o valor antes de salvar."), vbExclamation, "CAPRIND v5.0"
    txtValor_Compras.SetFocus
    Exit Sub
End If
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from vendas_analise_setores where ID = " & txtID_Compras, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = False Then
    If FunVerifStatusAnalise("alterar os dados de compras", True) = False Then Exit Sub
    If FunVerifValidSetorAnalise("alterar os dados de compras", SSTab1.Tab, True) = False Then Exit Sub
    TBGravar!Data_compras = Date
    TBGravar!Responsavel_compras = pubUsuario
    TBGravar!VlrUnit = txtValor_Compras
    TBGravar!vlrTotal = Format(TBGravar!VlrUnit * TBGravar!Qtde, "###,##0.00")
    TBGravar.Update
End If
TBGravar.Close
ProcCarregalista_Compras
USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
If CodigoLista9 <> 0 Then
    If Lista_compras.ListItems.Count = 0 Then Exit Sub
    Lista_compras.SelectedItem = Lista_compras.ListItems(CodigoLista9)
    Lista_compras.SetFocus
End If
'==================================
Modulo = "Outros/Análise crítica"
Evento = "Alterar dados de compras"
ID_documento = txtID_Compras
Documento = "Nº análise: " & Txt_analise & " - Rev.: " & Txt_rev_analise
Documento1 = "Cód. interno: " & Txtdesenho_compras & " - Descrição: " & txtTexto_Compras
ProcGravaEvento
'==================================
txtData_Compras = Format(Date, "dd/mm/yy")
txtResponsavel_Compras = pubUsuario

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvar_Compras_CheckList()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Frame_check(3).Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
Acao = "salvar"
If Txt_descricao_chek(3) = "" Then
    NomeCampo = "a descrição do check-list"
    ProcVerificaAcao
    ProcAbrirDescCheckList 3
    Exit Sub
End If
If Chk_nao_chek(3).Value = 0 And Chk_sim_chek(3).Value = 0 Then
    NomeCampo = "uma das opções"
    ProcVerificaAcao
    Exit Sub
End If
If Chk_sim_chek(3).Value = 1 And Txt_texto_check(3) = "" Then
    NomeCampo = "quais"
    ProcVerificaAcao
    Txt_texto_check(3).SetFocus
    Exit Sub
End If
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from Vendas_analise_setores_checklist where ID = " & Txt_ID_check(3), Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then
    TBGravar.AddNew
    TBGravar!IDAnalise = txtId
    TBGravar!Data = Date
    TBGravar!Responsavel = pubUsuario
    TBGravar!Setor = "COMPRAS"
    USMsgBox ("Novo check-list de compras cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo check-list de compras"
Else
    If FunVerifStatusAnalise("alterar o check-list de compras", True) = False Then Exit Sub
    If FunVerifValidSetorAnalise("alterar o check-list de compras", SSTab1.Tab, True) = False Then Exit Sub
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar check-list de compras"
End If
ProcEnviadados_Compras_Checklist
TBGravar.Update
Txt_ID_check(3) = TBGravar!ID
TBGravar.Close
ProcCarregaLista_Compras_Checklist
If Novo_analise9 = False Then
    If CodigoLista10 <> 0 And Lista_check(3).ListItems.Count <> 0 Then
        Lista_check(3).SelectedItem = Lista_check(3).ListItems(CodigoLista10)
        Lista_check(3).SetFocus
    End If
End If
'==================================
Modulo = "Outros/Análise crítica"
ID_documento = Txt_ID_check(3)
Documento = "Nº análise: " & Txt_analise & " - Rev.: " & Txt_rev_analise
Documento1 = "Descrição: " & Txt_descricao_chek(3)
ProcGravaEvento
'==================================
Novo_analise9 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procSalvar_doc()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Frame14.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
Acao = "salvar"
If txt_Caminho = "" Then
    NomeCampo = "o caminho"
    ProcVerificaAcao
    cmdImportar.SetFocus
    Exit Sub
End If
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from vendas_analise_setores where ID = " & txtID_doc, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then
    TBGravar.AddNew
    USMsgBox ("Novo documento cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo documento"
Else
    If FunVerifStatusAnalise("alterar o documento", True) = False Then Exit Sub
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar documento"
End If
ProcEnviadados_doc
TBGravar.Update
txtID_doc = TBGravar!ID
TBGravar.Close
ProcCarregaLista_Doc
If Novo_analise10 = False Then
    If CodigoLista11 <> 0 And Lista_doc.ListItems.Count <> 0 Then
        Lista_doc.SelectedItem = Lista_doc.ListItems(CodigoLista11)
        Lista_doc.SetFocus
    End If
End If
'==================================
Modulo = "Outros/Análise crítica"
ID_documento = txtID_doc
Documento = "Nº análise: " & Txt_analise & " - Rev.: " & Txt_rev_analise
Documento1 = "Caminho: " & txt_Caminho
ProcGravaEvento
'==================================
Novo_analise10 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAbrir()
On Error GoTo tratar_erro

frmVendas_analise_abrir.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_check_ColumnClick(index As Integer, ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Lista_check(index)
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If FunVerifStatusAnalise("", False) = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                If FunVerifValidSetorAnalise("", SSTab1.Tab, False) = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView Lista_check(index), ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_check_ItemCheck(index As Integer, ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

Select Case index
    Case 0: MsgTexto = "da engenharia"
    Case 1: MsgTexto = "do PCP"
    Case 2: MsgTexto = "da qualidade"
    Case 3: MsgTexto = "de compras"
End Select
With Lista_check(index)
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True And .ListItems.Item(InitFor) = Item Then
            If FunVerifStatusAnalise("excluir este check-list " & MsgTexto, True) = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            If FunVerifValidSetorAnalise("excluir este check-list" & MsgTexto, SSTab1.Tab, True) = False Then .ListItems.Item(InitFor).Checked = False
        End If
    Next InitFor
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_check_ItemClick(index As Integer, ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With Lista_check(index)
    If .ListItems.Count = 0 Then Exit Sub
    Set TBMaterial = CreateObject("adodb.recordset")
    TBMaterial.Open "Select VASC.*, VADC.Descricao from Vendas_analise_setores_checklist VASC INNER JOIN Vendas_analise_descricao_checklist VADC ON VADC.ID = VASC.IDchecklist where VASC.ID = " & Lista_check(index).SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
    If TBMaterial.EOF = False Then
        Select Case index
            Case 0:
                ProcLimpaCampos_Engenharia_Checklist
                ProcPuxadados_Engenharia_Checklist
                CodigoLista2 = .SelectedItem.index
            Case 1:
                ProcLimpaCampos_PCP_Checklist
                ProcPuxadados_PCP_Checklist
                CodigoLista6 = .SelectedItem.index
            Case 2:
                ProcLimpaCampos_Qualidade_Checklist
                ProcPuxadados_Qualidade_Checklist
                CodigoLista8 = .SelectedItem.index
            Case 3:
                ProcLimpaCampos_Compras_Checklist
                ProcPuxadados_Compras_Checklist
                CodigoLista10 = .SelectedItem.index
        End Select
    End If
    TBMaterial.Close
End With

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

Private Sub Lista_Compras_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Lista_compras
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If FunVerifStatusAnalise("", False) = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                If FunVerifValidSetorAnalise("", SSTab1.Tab, False) = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView Lista_compras, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_Compras_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With Lista_compras
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True And .ListItems.Item(InitFor) = Item Then
            If FunVerifStatusAnalise("excluir este valor", True) = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            If FunVerifValidSetorAnalise("excluir este valor", SSTab1.Tab, True) = False Then .ListItems.Item(InitFor).Checked = False
        End If
    Next InitFor
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_Compras_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista_compras.ListItems.Count = 0 Then Exit Sub
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "select * from vendas_analise_setores where id = " & Lista_compras.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    ProcLimpaCampos_Compras
    ProcPuxadados_compras
    CodigoLista9 = Lista_compras.SelectedItem.index
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_doc_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Lista_doc
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If FunVerifStatusAnalise("", False) = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView Lista_doc, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_doc_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With Lista_doc
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True And .ListItems.Item(InitFor) = Item Then
            If FunVerifStatusAnalise("excluir este documento", True) = False Then .ListItems.Item(InitFor).Checked = False
        End If
    Next InitFor
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_doc_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista_doc.ListItems.Count = 0 Then Exit Sub
Set TBMaterial = CreateObject("adodb.recordset")
TBMaterial.Open "Select * from vendas_analise_setores where id = " & Lista_doc.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBMaterial.EOF = False Then
    Proclimpacampos_doc
    ProcPuxadados_Doc
    CodigoLista11 = Lista_doc.SelectedItem.index
End If
TBMaterial.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_engenharia_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Lista_engenharia
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If FunVerifStatusAnalise("", False) = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                If FunVerifValidSetorAnalise("", SSTab1.Tab, False) = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView Lista_engenharia, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_engenharia_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With Lista_engenharia
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True And .ListItems.Item(InitFor) = Item Then
            If FunVerifStatusAnalise("excluir este registro", True) = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            If FunVerifValidSetorAnalise("excluir este registro", SSTab1.Tab, True) = False Then .ListItems.Item(InitFor).Checked = False
        End If
    Next InitFor
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_engenharia_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista_engenharia.ListItems.Count = 0 Then Exit Sub
Set TBMaterial = CreateObject("adodb.recordset")
TBMaterial.Open "select * from vendas_analise_setores where id = " & Lista_engenharia.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBMaterial.EOF = False Then
    ProcLimpaCampos_Engenharia_Prod
    ProcPuxadados_Engenharia_Prod
    CodigoLista1 = Lista_engenharia.SelectedItem.index
End If
TBMaterial.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 Then Exit Sub
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from vendas_analise where id = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    ProcLimpaCampos
    ProcPuxaDados
    ProcCarregaLista_Engenharia_Prod
    ProcCarregaLista_Engenharia_Checklist
    ProcCarregaLista_Engenharia_Normas
    ProcCarregaLista_processos_item
    ProcCarregaLista_PCP_Checklist
    ProcCarregaLista_Instrumentos
    ProcCarregaLista_Qualidade_Checklist
    ProcCarregalista_Compras
    ProcCarregaLista_Compras_Checklist
    ProcCarregaLista_Doc
End If
CodigoLista = Lista.SelectedItem.index

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_normas_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Lista_normas
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If FunVerifStatusAnalise("", False) = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                If FunVerifValidSetorAnalise("", SSTab1.Tab, False) = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView Lista_normas, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_normas_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With Lista_normas
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True And .ListItems.Item(InitFor) = Item Then
            If FunVerifStatusAnalise("excluir esta norma", True) = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            If FunVerifValidSetorAnalise("excluir esta norma", SSTab1.Tab, True) = False Then .ListItems.Item(InitFor).Checked = False
        End If
    Next InitFor
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_normas_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista_normas.ListItems.Count = 0 Then Exit Sub
Set TBMaterial = CreateObject("adodb.recordset")
TBMaterial.Open "select * from vendas_analise_setores where id = " & Lista_normas.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBMaterial.EOF = False Then
    ProcLimpaCampos_Normas
    ProcPuxadados_Normas
    CodigoLista3 = Lista_normas.SelectedItem.index
End If
TBMaterial.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub lista_Processos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With lista_Processos
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If FunVerifStatusAnalise("", False) = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                If FunVerifValidSetorAnalise("", SSTab1.Tab, False) = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView lista_Processos, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_processos_item_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Lista_processos_item
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If FunVerifStatusAnalise("", False) = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                If FunVerifValidSetorAnalise("", SSTab1.Tab, False) = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView Lista_processos_item, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_processos_item_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With Lista_processos_item
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True And .ListItems.Item(InitFor) = Item Then
            If FunVerifStatusAnalise("excluir este produto do processo", True) = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            If FunVerifValidSetorAnalise("excluir este produto do processo", SSTab1.Tab, True) = False Then .ListItems.Item(InitFor).Checked = False
        End If
    Next InitFor
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_processos_item_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista_processos_item.ListItems.Count = 0 Then Exit Sub
Set TBMaterial = CreateObject("adodb.recordset")
TBMaterial.Open "select * from vendas_analise_ProdutosProcessos where id = " & Lista_processos_item.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBMaterial.EOF = False Then
    ProcLimpaCampos_Processos_item True
    ProcPuxadados_Processos_item
    CodigoLista4 = Lista_processos_item.SelectedItem.index
End If
TBMaterial.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub lista_Processos_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With lista_Processos
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True And .ListItems.Item(InitFor) = Item Then
            If FunVerifStatusAnalise("excluir esta fase do processo", True) = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            If FunVerifValidSetorAnalise("excluir esta fase do processo", SSTab1.Tab, True) = False Then .ListItems.Item(InitFor).Checked = False
        End If
    Next InitFor
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub lista_Processos_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If lista_Processos.ListItems.Count = 0 Then Exit Sub
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "select * from vendas_analise_setores where id = " & lista_Processos.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    ProcLimpaCampos_Processos
    ProcPuxadados_processos
    CodigoLista5 = lista_Processos.SelectedItem.index
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_Qualidade_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Lista_Qualidade
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If FunVerifStatusAnalise("", False) = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                If FunVerifValidSetorAnalise("", SSTab1.Tab, False) = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView Lista_Qualidade, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_Qualidade_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With Lista_Qualidade
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True And .ListItems.Item(InitFor) = Item Then
            If FunVerifStatusAnalise("excluir este instrumento", True) = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            If FunVerifValidSetorAnalise("excluir este instrumento", SSTab1.Tab, True) = False Then .ListItems.Item(InitFor).Checked = False
        End If
    Next InitFor
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_Qualidade_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista_Qualidade.ListItems.Count = 0 Then Exit Sub
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "select * from vendas_analise_setores where id = " & Lista_Qualidade.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    ProcLimpaCampos_Instrumentos
    ProcPuxadados_Instrumentos
    CodigoLista7 = Lista_Qualidade.SelectedItem.index
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optProcessos_item_analise_Click()
On Error GoTo tratar_erro

ProcCarregaComboFamilia cmbFamilia_processos_item, "familia <> 'Null' and vendas = 'True'", False
ProcCarregaDadosItemAnalise

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optProcessos_item_Click()
On Error GoTo tratar_erro

ProcCarregaComboFamilia cmbFamilia_processos_item, "familia <> 'Null'", False
ProcLimpaCampos_Processos_item False
If optProcessos_item.Value = True Then Frame2.Enabled = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaDadosItemAnalise()
On Error GoTo tratar_erro

ProcLimpaCampos_Processos_item False
If optProcessos_item_analise.Value = True Then
    Frame2.Enabled = False
    txtCodInterno_processos_item = txtdesenho
    txtIDproduto_processos = txtidproduto
    If cmbReferencia <> "" Then
        cmbReferencia_processos_item.AddItem cmbReferencia
        cmbReferencia_processos_item = cmbReferencia
    End If
    txtDescricao_processos_item = txtdescricao
    If cmbfamilia <> "" Then cmbFamilia_processos_item = cmbfamilia
    If cmbun <> "" Then cmbUn_processos_item = cmbun
    If Cmb_un_com <> "" Then Cmb_un_com_processos_item = Cmb_un_com
    If txtQtde <> "" Then txtQtde_processos_item = txtQtde
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub SSTab_compras_Click(PreviousTab As Integer)
On Error GoTo tratar_erro

If SSTab1.Tab = 5 Then
    Select Case SSTab_compras.Tab
        Case 0:
            USToolBar3.Visible = True
            USToolBar2.Visible = False
            PBLista.Visible = False
            If Lista_compras.Visible = True Then Lista_compras.SetFocus
            ProcCarregalista_Compras
        Case 1:
            USToolBar3.Visible = False
            USToolBar2.Visible = True
            PBLista.Visible = True
            Lista_check(3).SetFocus
            ProcCarregaLista_Compras_Checklist
    End Select
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub SSTab_engenharia_Click(PreviousTab As Integer)
On Error GoTo tratar_erro

If SSTab1.Tab = 1 Then
    Select Case SSTab_engenharia.Tab
        Case 0:
            If Lista_engenharia.Visible = True Then Lista_engenharia.SetFocus
            ProcCarregaLista_Engenharia_Prod
        Case 1:
            Lista_check(0).SetFocus
            ProcCarregaLista_Engenharia_Checklist
        Case 2:
            Lista_normas.SetFocus
            ProcCarregaLista_Engenharia_Normas
    End Select
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub SSTab_qualidade_Click(PreviousTab As Integer)
On Error GoTo tratar_erro

If SSTab1.Tab = 4 Then
    Select Case SSTab_compras.Tab
        Case 0:
            If Lista_Qualidade.Visible = True Then Lista_Qualidade.SetFocus
            ProcCarregaLista_Instrumentos
        Case 1:
            Lista_check(2).SetFocus
            ProcCarregaLista_Qualidade_Checklist
    End Select
End If

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

If SSTab1.Tab = 2 Then
    If SSTab_processo.Tab = 0 Then PBLista.Visible = True
ElseIf SSTab1.Tab = 5 Then
        If SSTab_compras.Tab = 1 Then PBLista.Visible = True
    Else
        PBLista.Visible = True
End If

Select Case SSTab1.Tab
    Case 0:
        USToolBar1.Visible = True
        USToolBar3.Visible = False
        USToolBar2.Visible = False
        
        If Lista.Visible = True Then Lista.SetFocus
        Formulario = "Outros/Análise crítica"
        Direitos
    Case 1:
        USToolBar1.Visible = False
        With USToolBar3
            .Visible = True
            .ButtonToolTipText(1) = "Salvar."
            .ButtonToolTipText(2) = "Excluir."
        End With
        USToolBar2.Visible = False
        
        If Lista_engenharia.Visible = True Then Lista_engenharia.SetFocus
        Formulario = "Outros/Análise crítica/Engenharia"
        Direitos
        ProcVerificaAcessos
        If FunVerificaProsseguir = False Then Exit Sub
        ProcCarregaLista_Engenharia_Prod
    Case 2:
        USToolBar1.Visible = False
        USToolBar3.Visible = False
        SSTab_processo.Tab = 0
        If SSTab_processo.Tab = 0 Then
            USToolBar2.Visible = True
            USToolBar7.Visible = False
            If Lista_processos_item.Visible = True Then Lista_processos_item.SetFocus
            Formulario = "Outros/Análise crítica/Processos"
            Direitos
            If FunVerificaProsseguir = False Then Exit Sub
            ProcCarregaLista_processos_item
        Else
            USToolBar7.Visible = True
            USToolBar2.Visible = False
            PBLista.Visible = False
            If lista_Processos.Visible = True Then lista_Processos.SetFocus
            Formulario = "Outros/Análise crítica/Processos"
            Direitos
            ProcVerificaAcessos
            If FunVerificaProsseguir = False Then Exit Sub
            ProcLimpaCampos_Processos
            ProcCarregaLista_processos
        End If
    Case 3:
        USToolBar1.Visible = False
        With USToolBar3
            .Visible = True
            .ButtonToolTipText(1) = "Salvar."
            .ButtonToolTipText(2) = "Excluir."
            .Refresh
        End With
        USToolBar2.Visible = False
        
        If Lista_check(1).Visible = True Then Lista_check(1).SetFocus
        Formulario = "Outros/Análise crítica/Pcp"
        Direitos
        ProcVerificaAcessos
        If FunVerificaProsseguir = False Then Exit Sub
        ProcCarregaLista_PCP_Checklist
    Case 4:
        USToolBar1.Visible = False
        USToolBar3.Visible = False
        USToolBar2.Visible = True
        
        If Lista_Qualidade.Visible = True Then Lista_Qualidade.SetFocus
        Formulario = "Outros/Análise crítica/Qualidade"
        Direitos
        ProcVerificaAcessos
        If FunVerificaProsseguir = False Then Exit Sub
        ProcCarregaLista_Instrumentos
    Case 5:
        USToolBar1.Visible = False
        If SSTab_compras.Tab = 0 Then
            With USToolBar3
                .Visible = True
                .ButtonToolTipText(1) = "Salvar (F3)"
                .ButtonToolTipText(2) = "Excluir (F4)"
                .Refresh
            End With
            USToolBar2.Visible = False
        Else
            USToolBar3.Visible = False
            USToolBar2.Visible = True
        End If
        If SSTab_compras.Tab = 1 Then PBLista.Visible = True Else PBLista.Visible = False
        If Lista_compras.Visible Then Lista_compras.SetFocus
        Formulario = "Outros/Análise crítica/Compras"
        Direitos
        ProcVerificaAcessos
        If FunVerificaProsseguir = False Then Exit Sub
        ProcCarregalista_Compras
    Case 6:
        USToolBar1.Visible = False
        USToolBar3.Visible = False
        USToolBar2.Visible = True
        
        If Lista_doc.Visible Then Lista_doc.SetFocus
        Formulario = "Outros/Análise crítica/Documentos"
        Direitos
        ProcVerificaAcessos
        If FunVerificaProsseguir = False Then Exit Sub
        ProcCarregaLista_Doc
End Select

With USToolBar2
    If SSTab1.Tab = 6 Then .ButtonState(7) = 5 Else .ButtonState(7) = 0
    .Refresh
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Function FunVerificaProsseguir() As Boolean
On Error GoTo tratar_erro

FunVerificaProsseguir = True
If Acesso = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & ", você não tem acesso para este módulo " & Formulario & ", fale com o administrador do sistema."), vbExclamation, "CAPRIND v5.0"
    SSTab1.Tab = 0
    FunVerificaProsseguir = False
    Exit Function
End If
If Novo_Analise = True Then
    USMsgBox ("Salve a análise crítica antes de prosseguir."), vbExclamation, "CAPRIND v5.0"
    SSTab1.Tab = 0
    FunVerificaProsseguir = False
    Exit Function
End If

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Sub ProcCarregaLista(Pagina As Integer)
On Error GoTo tratar_erro

Lista.ListItems.Clear
lblRegistros.Caption = "Nº de registros: 0"
lblPaginas.Caption = "Página: 0 de: 0"
If StrSql_AnaliseCritica = "" Then Exit Sub
Set TBLISTA_AnaliseCritica = CreateObject("adodb.recordset")
TBLISTA_AnaliseCritica.Open StrSql_AnaliseCritica, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA_AnaliseCritica.EOF = False Then ProcExibePagina (Pagina)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina(Pagina)
On Error GoTo tratar_erro

Lista.ListItems.Clear
TBLISTA_AnaliseCritica.PageSize = IIf(txtNreg = "", 30, txtNreg)
TBLISTA_AnaliseCritica.AbsolutePage = Pagina
TamanhoPagina = TBLISTA_AnaliseCritica.PageSize
ContadorReg = 1
PBLista.Min = 0
PBLista.Max = FunVerifMaxPBListaPaginacao(TBLISTA_AnaliseCritica.RecordCount - IIf(Pagina > 1, (TBLISTA_AnaliseCritica.PageSize * (Pagina - 1)), 0), TBLISTA_AnaliseCritica.PageSize)
PBLista.Value = 1
Contador = 0
Do While TBLISTA_AnaliseCritica.EOF = False And (ContadorReg <= TamanhoPagina)
    With Lista.ListItems
        .Add , , TBLISTA_AnaliseCritica!ID
        .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA_AnaliseCritica!Nanalise), "", TBLISTA_AnaliseCritica!Nanalise)
        .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA_AnaliseCritica!Revisao), 0, TBLISTA_AnaliseCritica!Revisao)
        .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA_AnaliseCritica!Data), "", Format(TBLISTA_AnaliseCritica!Data, "dd/mm/yy"))
        .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA_AnaliseCritica!Responsavel), "", TBLISTA_AnaliseCritica!Responsavel)
        .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA_AnaliseCritica!Codinterno), "", TBLISTA_AnaliseCritica!Codinterno)
        .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA_AnaliseCritica!Descricao), "", TBLISTA_AnaliseCritica!Descricao)
        .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA_AnaliseCritica!Cliente), "", TBLISTA_AnaliseCritica!Cliente)
        Formulario = "Outros/Análise crítica/Vendas"
        ProcVerificaAcessos
        If Acesso = True Then
            Lista.ColumnHeaders(8).Width = "2000"
            Lista.ColumnHeaders(9).Width = "1200"
            .Item(.Count).SubItems(8) = IIf(IsNull(TBLISTA_AnaliseCritica!Valor_total), "0,00", Format(TBLISTA_AnaliseCritica!Valor_total, "###,##0.00"))
        Else
            Lista.ColumnHeaders(8).Width = "3200"
            Lista.ColumnHeaders(9).Width = "0"
            .Item(.Count).SubItems(8) = ""
        End If
        .Item(.Count).SubItems(9) = IIf(IsNull(TBLISTA_AnaliseCritica!DtValidacao_Engenharia) = False, "Sim", "Não")
        .Item(.Count).SubItems(10) = IIf(IsNull(TBLISTA_AnaliseCritica!DtValidacao_Processo) = False, "Sim", "Não")
        .Item(.Count).SubItems(11) = IIf(IsNull(TBLISTA_AnaliseCritica!DtValidacao_Pcp) = False, "Sim", "Não")
        .Item(.Count).SubItems(12) = IIf(IsNull(TBLISTA_AnaliseCritica!DtValidacao_Qualidade) = False, "Sim", "Não")
        .Item(.Count).SubItems(13) = IIf(IsNull(TBLISTA_AnaliseCritica!DtValidacao_Compras) = False, "Sim", "Não")
    End With
    TBLISTA_AnaliseCritica.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista.Value = Contador
Loop
lblRegistros.Caption = "Nº de registros: " & TBLISTA_AnaliseCritica.RecordCount
If TBLISTA_AnaliseCritica.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "Página: 1 de: " & TBLISTA_AnaliseCritica.PageCount
ElseIf TBLISTA_AnaliseCritica.AbsolutePage = adPosEOF Then
        lblPaginas.Caption = "Página: " & TBLISTA_AnaliseCritica.PageCount & " de: " & TBLISTA_AnaliseCritica.PageCount
    Else
        lblPaginas.Caption = "Página: " & TBLISTA_AnaliseCritica.AbsolutePage - 1 & " de: " & TBLISTA_AnaliseCritica.PageCount
End If


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaLista_Engenharia_Prod()
On Error GoTo tratar_erro

Lista_engenharia.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "select * from vendas_analise_setores where idanalise = " & txtId & " and setor = 'ENGENHARIA' order by Codinterno", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBLISTA.EOF = False
        With Lista_engenharia.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Codinterno), "", TBLISTA!Codinterno)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Texto), "", TBLISTA!Texto)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Un), "", TBLISTA!Un)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!Unidade_com), "", TBLISTA!Unidade_com)
            .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!Qtde), "", Format(TBLISTA!Qtde, "###,##0.0000"))
            .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA!VlrUnit), "", Format(TBLISTA!VlrUnit, "###,##0.0000000000"))
            .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA!vlrTotal), "", Format(TBLISTA!vlrTotal, "###,##0.00"))
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

Sub ProcCarregaLista_Engenharia_Checklist()
On Error GoTo tratar_erro

Lista_check(0).ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select VASC.*, VADC.Descricao from Vendas_analise_setores_checklist VASC INNER JOIN Vendas_analise_descricao_checklist VADC ON VADC.ID = VASC.IDchecklist where VASC.idanalise = " & txtId & " and VASC.setor = 'ENGENHARIA'", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBLISTA.EOF = False
        With Lista_check(0).ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Data), "", Format(TBLISTA!Data, "dd/mm/yy"))
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Responsavel), "", TBLISTA!Responsavel)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Descricao), "", TBLISTA!Descricao)
            .Item(.Count).SubItems(4) = IIf(TBLISTA!Sim = True, "Sim", "Não")
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

Sub ProcCarregaLista_Engenharia_Normas()
On Error GoTo tratar_erro

Lista_normas.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "select * from vendas_analise_setores where idanalise = " & txtId & " and setor = 'NORMAS' order by ID desc", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBLISTA.EOF = False
        With Lista_normas.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Data), "", Format(TBLISTA!Data, "dd/mm/yy"))
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Responsavel), "", TBLISTA!Responsavel)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Texto), "", TBLISTA!Texto)
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

Sub ProcCarregaLista_processos_item()
On Error GoTo tratar_erro

Lista_processos_item.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "select * from vendas_analise_ProdutosProcessos where id_analise = " & txtId & " order by Codinterno", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBLISTA.EOF = False
        With Lista_processos_item.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Codinterno), "", TBLISTA!Codinterno)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Descricao), "", TBLISTA!Descricao)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Familia), "", TBLISTA!Familia)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!Un), "", TBLISTA!Un)
            .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!Unidade_com), "", TBLISTA!Unidade_com)
            .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA!Qtde), "1,000", Format(TBLISTA!Qtde, "###,##0.0000"))
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

Sub ProcCarregaLista_processos()
On Error GoTo tratar_erro

valor = 0
lista_Processos.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
StrSql = "select * from vendas_analise_setores where idanalise = " & txtId & " and Setor = 'PROCESSOS' and ID_processo_item = " & txtID_processos_item & " order by fase, texto"
'Debug.print StrSql


TBLISTA.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista_Fases.Min = 0
    PBLista_Fases.Max = TBLISTA.RecordCount
    PBLista_Fases.Value = 1
    Contador = 0
    Do While TBLISTA.EOF = False
        With lista_Processos.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Fase), "", TBLISTA!Fase)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Texto), "", TBLISTA!Texto)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Descricao), "", TBLISTA!Descricao)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!TotalHora), "", Format(TBLISTA!TotalHora, "HH:MM:SS"))
            .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!PrecoHora_Setup), "", Format(TBLISTA!PrecoHora_Setup, "###,##0.00"))
            .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA!VlrUnit), "", Format(TBLISTA!VlrUnit, "###,##0.00"))
            .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA!vlrTotal), "", Format(TBLISTA!vlrTotal, "###,##0.00"))
            valor = valor + IIf(IsNull(TBLISTA!vlrTotal), 0, TBLISTA!vlrTotal)
        End With
        TBLISTA.MoveNext
        Contador = Contador + 1
        PBLista_Fases.Value = Contador
    Loop
End If
TBLISTA.Close
Txt_valor_total_processo = Format(valor, "###,##0.00")
        
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaLista_PCP_Checklist()
On Error GoTo tratar_erro

Lista_check(1).ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select VASC.*, VADC.Descricao from Vendas_analise_setores_checklist VASC INNER JOIN Vendas_analise_descricao_checklist VADC ON VADC.ID = VASC.IDchecklist where VASC.idanalise = " & txtId & " and VASC.setor = 'PCP'", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBLISTA.EOF = False
        With Lista_check(1).ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Data), "", Format(TBLISTA!Data, "dd/mm/yy"))
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Responsavel), "", TBLISTA!Responsavel)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Descricao), "", TBLISTA!Descricao)
            .Item(.Count).SubItems(4) = IIf(TBLISTA!Sim = True, "Sim", "Não")
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

Sub ProcCarregaLista_Instrumentos()
On Error GoTo tratar_erro

Lista_Qualidade.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "select * from vendas_analise_setores where idanalise = " & txtId & " and setor = 'QUALIDADE' order by Codinterno", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBLISTA.EOF = False
        With Lista_Qualidade.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Codinterno), "", TBLISTA!Codinterno)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Texto), "", TBLISTA!Texto)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Un), "", TBLISTA!Un)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!Unidade_com), "", TBLISTA!Unidade_com)
            .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!Qtde), "", Format(TBLISTA!Qtde, "###,##0.0000"))
            .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA!VlrUnit), "", Format(TBLISTA!VlrUnit, "###,##0.0000000000"))
            .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA!vlrTotal), "", Format(TBLISTA!vlrTotal, "###,##0.00"))
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

Sub ProcCarregaLista_Qualidade_Checklist()
On Error GoTo tratar_erro

Lista_check(2).ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select VASC.*, VADC.Descricao from Vendas_analise_setores_checklist VASC INNER JOIN Vendas_analise_descricao_checklist VADC ON VADC.ID = VASC.IDchecklist where VASC.idanalise = " & txtId & " and VASC.setor = 'QUALIDADE'", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBLISTA.EOF = False
        With Lista_check(2).ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Data), "", Format(TBLISTA!Data, "dd/mm/yy"))
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Responsavel), "", TBLISTA!Responsavel)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Descricao), "", TBLISTA!Descricao)
            .Item(.Count).SubItems(4) = IIf(TBLISTA!Sim = True, "Sim", "Não")
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

Sub ProcCarregalista_Compras()
On Error GoTo tratar_erro

ValorPagar = 0 'Material
ValorPago = 0 'Terceiros
Valor_Cofins_Prod = 0 'Outros
Valor_Cofins_Serv = 0 'Ferramentas
Lista_compras.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select VAS.* from vendas_analise_setores VAS LEFT JOIN Projproduto P ON P.Codproduto = VAS.IDproduto where VAS.idanalise = " & txtId & " and VAS.Setor <> 'PROCESSOS' and VAS.Setor <> 'DOCUMENTO' and VAS.Setor <> 'NORMAS' and (P.Compras = 1 or P.Desenho IS NULL)", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista_compras.Min = 0
    PBLista_compras.Max = TBLISTA.RecordCount
    PBLista_compras.Value = 1
    Contador = 0
    Do While TBLISTA.EOF = False
        With Lista_compras.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Codinterno), "", TBLISTA!Codinterno)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Texto), "", TBLISTA!Texto)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Un), "", TBLISTA!Un)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!Unidade_com), "", TBLISTA!Unidade_com)
            .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!Qtde), "", Format(TBLISTA!Qtde, "###,##0.0000"))
            .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA!VlrUnit), "", Format(TBLISTA!VlrUnit, "###,##0.0000"))
            .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA!vlrTotal), "", Format(TBLISTA!vlrTotal, "###,##0.00"))
            If TBLISTA!Setor = "FERRAMENTAS" Then
                Valor_Cofins_Serv = Valor_Cofins_Serv + IIf(IsNull(TBLISTA!vlrTotal), 0, TBLISTA!vlrTotal)
            Else
                If TBLISTA!Tipo = "M" Then
                    ValorPagar = ValorPagar + IIf(IsNull(TBLISTA!vlrTotal), 0, TBLISTA!vlrTotal)
                ElseIf TBLISTA!Tipo = "T" Then
                        ValorPago = ValorPago + IIf(IsNull(TBLISTA!vlrTotal), 0, TBLISTA!vlrTotal)
                    Else
                        Valor_Cofins_Prod = Valor_Cofins_Prod + IIf(IsNull(TBLISTA!vlrTotal), 0, TBLISTA!vlrTotal)
                End If
            End If
        End With
        TBLISTA.MoveNext
        Contador = Contador + 1
        PBLista_compras.Value = Contador
    Loop
End If
Txt_total_materiais = Format(ValorPagar, "###,##0.00")
Txt_total_terceiros = Format(ValorPago, "###,##0.00")
Txt_total_outros = Format(Valor_Cofins_Prod, "###,##0.00")
Txt_total_ferramentas = Format(Valor_Cofins_Serv, "###,##0.00")
Txt_total_compras = Format(ValorPagar + ValorPago + Valor_Cofins_Prod + Valor_Cofins_Serv, "###,##0.00")


ValorTotalMateriais = Replace(Txt_total_materiais, ".", "")
ValorTotalMateriais = Replace(ValorTotalMateriais, ",", ".")

If Txt_total_materiais <> "" And Txt_total_materiais <> "0,00" Then
StrSql = "Update Vendas_analise set Valor_total_materiais  = '" & ValorTotalMateriais & "' Where ID = " & txtId.Text & ""
'Debug.print StrSql

Conexao.Execute (StrSql)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaLista_Compras_Checklist()
On Error GoTo tratar_erro

Lista_check(3).ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select VASC.*, VADC.Descricao from Vendas_analise_setores_checklist VASC INNER JOIN Vendas_analise_descricao_checklist VADC ON VADC.ID = VASC.IDchecklist where VASC.idanalise = " & txtId & " and VASC.setor = 'COMPRAS'", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBLISTA.EOF = False
        With Lista_check(3).ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Data), "", Format(TBLISTA!Data, "dd/mm/yy"))
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Responsavel), "", TBLISTA!Responsavel)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Descricao), "", TBLISTA!Descricao)
            .Item(.Count).SubItems(4) = IIf(TBLISTA!Sim = True, "Sim", "Não")
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

Sub ProcCarregaLista_Doc()
On Error GoTo tratar_erro

Lista_doc.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "select * from vendas_analise_setores where idanalise = " & txtId & " and setor = 'DOCUMENTO' order by ID", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBLISTA.EOF = False
        With Lista_doc.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Texto), "", TBLISTA!Texto)
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

Sub ProcEnviaDados()
On Error GoTo tratar_erro

TBGravar!Revisao = IIf(Txt_rev_analise = "", Null, Txt_rev_analise)
TBGravar!status = Txt_status
TBGravar!Data_status = IIf(Txt_data_status = "", Null, Txt_data_status)
TBGravar!IDProduto = IIf(txtidproduto = "", Null, txtidproduto)
TBGravar!Codinterno = txtdesenho
TBGravar!RevDesenho = txtRev_desenho
TBGravar!N_referencia = cmbReferencia
TBGravar!Unidade = cmbun
TBGravar!Unidade_com = Cmb_un_com
TBGravar!qtde_solicitada = Txt_qtde_sol
TBGravar!Qtde = IIf(txtQtde = "", 1, txtQtde)
TBGravar!Descricao = txtdescricao
TBGravar!Familia = cmbfamilia
TBGravar!Tipo = Txt_tipo
TBGravar!IDCliente = txtIDcliente
TBGravar!Cliente = txtCliente
TBGravar!contato = txtContato
TBGravar!Departamento = txtdepartamento
TBGravar!telefone = txttelefone
TBGravar!Fax = txtFax
TBGravar!Email = txtEmail
TBGravar!NRef = txtNRef
TBGravar!Referencia = txtreferencia

If cmbLocal_cobranca <> "" Then
    TBGravar!Local_Cobranca = cmbLocal_cobranca
    TBGravar!ID_Cobranca = txtID_cobranca
Else
    TBGravar!Local_Cobranca = ""
    TBGravar!ID_Cobranca = 0
End If

If cmbLocal_entrega <> "" Then
    TBGravar!Local_entrega = cmbLocal_entrega
    TBGravar!ID_entrega = txtID_entrega
Else
    TBGravar!Local_entrega = ""
    TBGravar!ID_entrega = 0
End If

TBGravar!Obs = txtObs

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviadados_Engenharia_Prod()
On Error GoTo tratar_erro

TBGravar!IDProduto = IIf(txtIDproduto_engenharia = "", Null, txtIDproduto_engenharia)
If Opt_material.Value = True Then TBGravar!Tipo = "M"
If Opt_terceiros.Value = True Then TBGravar!Tipo = "T"
If Opt_outros.Value = True Then TBGravar!Tipo = "O"
TBGravar!Codinterno = txtdesenho_engenharia
TBGravar!N_referencia = cmbReferencia_engenharia
TBGravar!Un = cmbun_engenharia
TBGravar!Unidade_com = Cmb_un_com_engenharia
TBGravar!Qtde = txtQtde_engenharia
TBGravar!Texto = txtdescricao_engenharia
TBGravar!Familia = cmbfamilia_engenharia
TBGravar!Analise = Trim(txtAnalise_engenharia)
TBGravar!Setor = "ENGENHARIA"

If Novo_analise1 = True Then
    If Opt_outros.Value = True Then
        Permitido = True
        Set TBProduto = CreateObject("adodb.recordset")
        TBProduto.Open "Select Valor_Total from Vendas_analise where Codinterno = '" & txtdesenho_engenharia & "' and Fechada = 'True' order by ID desc", Conexao, adOpenKeyset, adLockOptimistic
        If TBProduto.EOF = True Then
            Permitido = False
        End If
    Else
        Permitido = False
    End If
    If Permitido = True Then
        TBGravar!VlrUnit = IIf(IsNull(TBProduto!Valor_total), 0, TBProduto!Valor_total)
    Else
        Set TBProduto = CreateObject("adodb.recordset")
        TBProduto.Open "Select Pcusto from projproduto where codproduto = " & IIf(IsNull(TBGravar!IDProduto), 0, TBGravar!IDProduto), Conexao, adOpenKeyset, adLockOptimistic
        If TBProduto.EOF = False Then
            TBGravar!VlrUnit = IIf(IsNull(TBProduto!PCusto), 0, TBProduto!PCusto)
        End If
        TBProduto.Close
    End If
End If
TBGravar!vlrTotal = Format(TBGravar!VlrUnit * TBGravar!Qtde, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviadados_Engenharia_Checklist()
On Error GoTo tratar_erro

TBGravar!IDchecklist = Txt_ID_descricao_check(0)
If Chk_sim_chek(0).Value = 1 Then TBGravar!Sim = True Else TBGravar!Sim = False
TBGravar!Quais = Txt_texto_check(0)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviadados_Processos_item()
On Error GoTo tratar_erro

TBGravar!id_analise = txtId
TBGravar!Codproduto = IIf(txtIDproduto_processos = "", Null, txtIDproduto_processos)
TBGravar!Codinterno = txtCodInterno_processos_item
TBGravar!Referencia = IIf(cmbReferencia_processos_item = "", Null, cmbReferencia_processos_item)
TBGravar!Un = cmbUn_processos_item
TBGravar!Unidade_com = Cmb_un_com_processos_item
TBGravar!Descricao = txtDescricao_processos_item
TBGravar!Familia = cmbFamilia_processos_item
If optProcessos_item_analise.Value = True Then TBGravar!Produto_analise = True Else TBGravar!Produto_analise = False
TBGravar!Qtde = txtQtde_processos_item

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "select * from vendas_analise_setores where idanalise = " & txtId & " and id_processo_item = " & txtID_processos_item & " and setor = 'PROCESSOS'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    Do While TBAbrir.EOF = False
        Qtde = 0
        valor = 0
        qt = 0
        Qtd = 0
        ValorTotal = 0
        quantidade = 0
        Qtde = IIf(IsNull(TBAbrir!Qtde), 0, TBAbrir!Qtde)
        valor = IIf(IsNull(TBAbrir!VlrUnit), 0, TBAbrir!VlrUnit)
            
        dataCalculo = IIf(IsNull(TBAbrir!Preparacao), 0, TBAbrir!Preparacao)
        ProcFormataHora (dataCalculo)
        qt = s / 3600
        Qtd = txtQtde_processos_item
        ValorTotal = Qtde * valor
        If qt > 0 Then
            quantidade = (qt / Qtd) * valor
            ValorTotal = ValorTotal + quantidade
        End If
        TBAbrir!vlrTotal = Format(ValorTotal, "###,##0.00")
        TBAbrir.Update
        TBAbrir.MoveNext
    Loop
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviadados_processos()
On Error GoTo tratar_erro

TBGravar!Texto = txtMaquina_processos
TBGravar!Descricao = txtDescricao_processos
TBGravar!TotalHora = IIf(txtTotalHora_processos = "", Null, txtTotalHora_processos)
ProcFormataHora (TBGravar!TotalHora)
TBGravar!Qtde = (s + DecimoSegundos) / 3600
TBGravar!Peca = IIf(txtPecaHora_processos = "", Null, txtPecaHora_processos)
TBGravar!Execucao = IIf(txtExecucao_processos = "", Null, txtExecucao_processos)
TBGravar!Preparacao = IIf(txtPreparacao_processos = "", Null, txtPreparacao_processos)
TBGravar!VlrUnit = IIf(txtValorHora_processos = "", Null, txtValorHora_processos)
TBGravar!PrecoHora_Setup = IIf(txtValorHoraPrep_Processos = "", Null, txtValorHoraPrep_Processos)
TBGravar!vlrTotal = IIf(txtValorTotal_processos = "", Null, txtValorTotal_processos)
TBGravar!Trabalho = IIf(txtTrabalho.TextRTF = "", Null, txtTrabalho.TextRTF)
TBGravar!Fase = IIf(txtFase = "", Null, txtFase)
TBGravar!Grupo_op = IIf(txtgrupo_op = "", Null, txtgrupo_op)
If chkPchora.Value = 1 Then TBGravar!pecahora = True Else TBGravar!pecahora = False
TBGravar!Erro_processos = IIf(txtErro = "", Null, txtErro)
TBGravar!Setor = "PROCESSOS"
TBGravar!ID_processo_item = txtID_processos_item

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviadados_PCP_Checklist()
On Error GoTo tratar_erro

TBGravar!IDchecklist = Txt_ID_descricao_check(1)
If Chk_sim_chek(1).Value = 1 Then TBGravar!Sim = True Else TBGravar!Sim = False
TBGravar!Quais = Txt_texto_check(1)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviadados_Instrumento()
On Error GoTo tratar_erro

TBGravar!Tipo = "M"
TBGravar!IDProduto = IIf(txtIDproduto_qualidade = "", Null, txtIDproduto_qualidade)
TBGravar!Codinterno = txtdesenho_qualidade
TBGravar!N_referencia = cmbReferencia_qualidade
TBGravar!Un = cmbun_qualidade
TBGravar!Unidade_com = Cmb_un_com_qualidade
TBGravar!Qtde = txtQtde_Qualidade
TBGravar!Texto = txtdescricao_Qualidade
TBGravar!Familia = cmbfamilia_qualidade
TBGravar!Analise = Trim(txtAnalise_qualidade)
TBGravar!Setor = "QUALIDADE"

Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select Pcusto from projproduto where codproduto = " & IIf(IsNull(TBGravar!IDProduto), 0, TBGravar!IDProduto), Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    TBGravar!VlrUnit = IIf(IsNull(TBProduto!PCusto), 0, TBProduto!PCusto)
End If
TBProduto.Close
TBGravar!vlrTotal = Format(TBGravar!VlrUnit * TBGravar!Qtde, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviadados_Qualidade_Checklist()
On Error GoTo tratar_erro

TBGravar!IDchecklist = Txt_ID_descricao_check(2)
If Chk_sim_chek(2).Value = 1 Then TBGravar!Sim = True Else TBGravar!Sim = False
TBGravar!Quais = Txt_texto_check(2)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviadados_Compras_Checklist()
On Error GoTo tratar_erro

TBGravar!IDchecklist = Txt_ID_descricao_check(3)
If Chk_sim_chek(3).Value = 1 Then TBGravar!Sim = True Else TBGravar!Sim = False
TBGravar!Quais = Txt_texto_check(3)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviadados_doc()
On Error GoTo tratar_erro

TBGravar!IDAnalise = txtId
TBGravar!Data = IIf(txtData_doc = "", Date, txtData_doc)
TBGravar!Responsavel = IIf(txtResponsavel_doc = "", pubUsuario, txtResponsavel_doc)
TBGravar!Texto = txt_Caminho
TBGravar!Analise = Trim(Txt_obs_doc)
TBGravar!Setor = "DOCUMENTO"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub SSTab_processo_Click(PreviousTab As Integer)
On Error GoTo tratar_erro

If txtID_processos_item = 0 Then
    SSTab_processo.Tab = 0
    Exit Sub
End If
Select Case SSTab_processo.Tab
    Case 0:
        USToolBar7.Visible = False
        USToolBar2.Visible = True
        PBLista.Visible = True
        If Lista_processos_item.Visible = True Then Lista_processos_item.SetFocus
        Formulario = "Outros/Análise crítica/Processos"
        Direitos
    Case 1:
        USToolBar7.Visible = True
        USToolBar2.Visible = False
        PBLista.Visible = False
        If lista_Processos.Visible = True Then lista_Processos.SetFocus
        Formulario = "Outros/Análise crítica/Processos"
        Direitos
        ProcVerificaAcessos
        If FunVerificaProsseguir = False Then Exit Sub
        ProcLimpaCampos_Processos
        ProcCarregaLista_processos
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_analise_Change()
On Error GoTo tratar_erro

If Novo_Analise = True Then
VerifCodigo:
    Set TBFI = CreateObject("adodb.recordset")
    TBFI.Open "Select ID from Vendas_analise where Nanalise = '" & Txt_analise & "' and ID <> " & txtId, Conexao, adOpenKeyset, adLockOptimistic
    If TBFI.EOF = False Then
        Cotacao = Left(Txt_analise, Len(Txt_analise) - 3) + 1
        Ano = Right(Year(Date), 2)
        Select Case Len(Cotacao)
            Case 1: NumeroAnalise = "000" & Cotacao & "/" & Ano
            Case 2: NumeroAnalise = "00" & Cotacao & "/" & Ano
            Case 3: NumeroAnalise = "0" & Cotacao & "/" & Ano
            Case 4: NumeroAnalise = Cotacao & "/" & Ano
            Case 5: NumeroAnalise = Cotacao & "/" & Ano
        End Select
        Txt_analise = NumeroAnalise
        GoTo VerifCodigo
    End If
    TBFI.Close
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtCodInterno_processos_item_Change()
On Error GoTo tratar_erro

ProcLimpaCamposProdProcessos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtDesenho_Change()
On Error GoTo tratar_erro

ProcLimpaCamposProd

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtdesenho_engenharia_Change()
On Error GoTo tratar_erro

If chkManual_engenharia.Value = 0 And chkAuto_engenharia.Value = 0 Then ProcLimpaCamposProdEng

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtdesenho_qualidade_Change()
On Error GoTo tratar_erro

If chkManual_qualidade.Value = 0 Then ProcLimpaCamposProdQualidade

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtExecucao_processos_Change()
On Error GoTo tratar_erro

ProcCalculaExecucao

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

Private Sub txtQtde_processos_item_Change()
On Error GoTo tratar_erro

If txtQtde_processos_item <> "" Then
    VerifNumero = txtQtde_processos_item
    ProcVerificaNumero
    If VerifNumero = False Then
        txtQtde_processos_item = ""
        txtQtde_processos_item.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtQtde_processos_item_GotFocus()
On Error GoTo tratar_erro

FunGotFocus txtQtde_processos_item

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtQtde_processos_item_LostFocus()
On Error GoTo tratar_erro

txtQtde_processos_item = Format(txtQtde_processos_item, "###,##0.0000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_qtde_sol_Change()
On Error GoTo tratar_erro

If Txt_qtde_sol <> "" Then
    VerifNumero = Txt_qtde_sol
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_qtde_sol = ""
        Txt_qtde_sol.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_qtde_sol_GotFocus()
On Error GoTo tratar_erro

FunGotFocus Txt_qtde_sol

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_qtde_sol_LostFocus()
On Error GoTo tratar_erro

Txt_qtde_sol = Format(Txt_qtde_sol, "###,##0.0000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtErro_Change()
On Error GoTo tratar_erro

If txtErro.Text <> "" Then
    VerifNumero = txtErro.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtErro.Text = ""
        txtErro.SetFocus
        Exit Sub
    End If
End If
ProcCalculaExecucao

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtFase_LostFocus()
On Error GoTo tratar_erro

If txtFase.Text <> "" Then
    VerifNumero = txtFase.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtFase.Text = ""
        txtFase.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtIDcliente_Change()
On Error GoTo tratar_erro

If txtIDcliente <> "" Then
    VerifNumero = txtIDcliente
    ProcVerificaNumero
    If VerifNumero = False Then
        txtIDcliente = ""
        txtIDcliente.SetFocus
        Exit Sub
    End If
    ProcPuxadadosCliente
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcPuxadadosCliente()
On Error GoTo tratar_erro

IDFornecedor = txtIDcliente
txtCliente = ""
txtContato = ""
cmbLocal_entrega.ListIndex = -1
cmbLocal_cobranca.ListIndex = -1
Set TBClientes = CreateObject("adodb.recordset")
TBClientes.Open "Select * from clientes where IDCliente = " & IIf(IDFornecedor = "", 0, IDFornecedor) & " and status <> 'Bloqueado'", Conexao, adOpenKeyset, adLockOptimistic
If TBClientes.EOF = False Then
    txtCliente = IIf(IsNull(TBClientes!NomeRazao), "", TBClientes!NomeRazao)
    cmdLocalcobranca_Click
    cmdLocalentrega_Click
End If
TBClientes.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtPreparacao_processos_Change()
On Error GoTo tratar_erro

ProcCalculaExecucao

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

txtQtde = Format(txtQtde, "###,##0.0000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtMaquina_processos_Change()
On Error GoTo tratar_erro

txtDescricao_maquina = ""
If txtmaquina = "" Then Exit Sub
Set TBProposta = CreateObject("adodb.recordset")
TBProposta.Open "select * from cadmaquinas where maquina = '" & txtmaquina & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBProposta.EOF = False Then
    txtDescricao_maquina.Locked = True
    txtDescricao_maquina.TabStop = False
Else
    txtDescricao_maquina.Locked = False
    txtDescricao_maquina.TabStop = True
End If
TBProposta.Close
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtMaquina_processos_LostFocus()
On Error GoTo tratar_erro

If txtmaquina = "" Then Exit Sub
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "select * from cadmaquinas where maquina = '" & txtmaquina & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    txtmaquina = IIf(IsNull(TBAbrir!maquina), "", TBAbrir!maquina)
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtPecaHora_processos_Change()
On Error GoTo tratar_erro

If txtPecaHora_processos.Text <> "" Then
    VerifNumero = txtPecaHora_processos.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtPecaHora_processos.Text = ""
        txtPecaHora_processos.SetFocus
        Exit Sub
    End If
End If
ProcCalculaExecucao

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtQtde_engenharia_Change()
On Error GoTo tratar_erro

If txtQtde_engenharia.Text <> "" Then
    VerifNumero = txtQtde_engenharia.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtQtde_engenharia.Text = ""
        txtQtde_engenharia.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtQtde_engenharia_GotFocus()
On Error GoTo tratar_erro

FunGotFocus txtQtde_engenharia

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtQtde_engenharia_LostFocus()
On Error GoTo tratar_erro

txtQtde_engenharia = Format(txtQtde_engenharia, "###,##0.0000")
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtQtde_Qualidade_Change()
On Error GoTo tratar_erro

If txtQtde_Qualidade.Text <> "" Then
    VerifNumero = txtQtde_Qualidade.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtQtde_Qualidade.Text = ""
        txtQtde_Qualidade.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtQtde_Qualidade_GotFocus()
On Error GoTo tratar_erro

FunGotFocus txtQtde_Qualidade

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtQtde_Qualidade_LostFocus()
On Error GoTo tratar_erro

txtQtde_Qualidade = Format(txtQtde_Qualidade, "###,##0.0000")
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtTotalHora_processos_Change()
On Error GoTo tratar_erro

ProcCalculamaquina

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtValor_Compras_Change()
On Error GoTo tratar_erro

If txtValor_Compras.Text <> "" Then
    VerifNumero = txtValor_Compras.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtValor_Compras.Text = ""
        txtValor_Compras.SetFocus
        Exit Sub
    End If
    valor = txtValor_Compras
    Qtde = IIf(txtQtde_compras = "", 0, txtQtde_compras)
    Txt_valor_total = Format(valor * Qtde, "###,##0.00")
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtValor_Compras_GotFocus()
On Error GoTo tratar_erro

If txtValor_Compras = "0,0000" Then txtValor_Compras = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtValor_Compras_LostFocus()
On Error GoTo tratar_erro

txtValor_Compras.Text = IIf(txtValor_Compras.Text = "", "0,0000", Format(txtValor_Compras.Text, "###,##0.0000"))
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtValorHora_processos_Change()
On Error GoTo tratar_erro

If txtValorHora_processos <> "" Then
    VerifNumero = txtValorHora_processos
    ProcVerificaNumero
    If VerifNumero = False Then
        txtValorHora_processos = ""
        txtValorHora_processos.SetFocus
        Exit Sub
    End If
    ProcCalculamaquina
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtValorHora_processos_LostFocus()
On Error GoTo tratar_erro

txtValorHora_processos = Format(txtValorHora_processos, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCalculamaquina()
On Error GoTo tratar_erro

Qtde = 0
valor = 0
qt = 0
Qtd = 0
ValorTotal = 0
quantidade = 0
quantnovo = 0

'Calcula valor de execução
NovoValor = Replace(txtTotalHora_processos, ",", ".")
If txtTotalHora_processos <> "" Then
    ProcFormataHora (txtTotalHora_processos)
    HoraResultado = DataResultado
    ElapsedTime (HoraResultado)
    Qtde = (s + DecimoSegundos) / 3600
End If
valor = IIf(txtValorHora_processos = "", 0, txtValorHora_processos)
ValorTotal = Qtde * valor

'Calcula valor de preparação
Qtd = IIf(txtQtde_processos_item = "", 0, txtQtde_processos_item)
Valor1 = IIf(txtValorHoraPrep_Processos = "", 0, txtValorHoraPrep_Processos) / Qtd
txtPreparacao_processos.PromptInclude = False
If Len(txtPreparacao_processos.Text) = 7 Then
    txtPreparacao_processos.PromptInclude = True
    ProcFormataHora (txtPreparacao_processos)
    HoraResultado = DataResultado
    ElapsedTime (HoraResultado)
    qt = s / 3600
End If
quantnovo = IIf(txtErro = "", 0, txtErro)

If qt > 0 Then
    quantidade = qt * Valor1
    ValorTotal = ValorTotal + quantidade
End If
txtValorTotal_processos = Format(ValorTotal, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaFamiliaUN()
On Error GoTo tratar_erro

'Carrega combo família
ProcCarregaComboFamilia cmbfamilia, "familia <> 'Null' and vendas = 'True'", False
ProcCarregaComboFamilia cmbfamilia_engenharia, "familia <> 'Null'", False
ProcCarregaComboFamilia cmbfamilia_qualidade, "familia <> 'Null' and (compras = 'True' or qualidade = 'True')", False

'Carrega combo unidade
ProcCarregaComboUnidade cmbun, False
ProcCarregaComboUnidade Cmb_un_com, False
ProcCarregaComboUnidade cmbun_engenharia, False
ProcCarregaComboUnidade Cmb_un_com_engenharia, False
ProcCarregaComboUnidade cmbUn_processos_item, False
ProcCarregaComboUnidade Cmb_un_com_processos_item, False
ProcCarregaComboUnidade cmbun_qualidade, False
ProcCarregaComboUnidade Cmb_un_com_qualidade, False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovoProdutoAutoEngQualidade()
On Error GoTo tratar_erro

If SSTab1.Tab = 1 Then
    txtdesenho_engenharia = FunCriaNovoProdServ(False, "codmanual = 'False' and (subtipoitem = 0 or subtipoitem = 1 or subtipoitem = 4 or subtipoitem = 5)", txtdesenho_engenharia, txtReferencia_engenharia, 0, txtdescricao_engenharia, txtdescricao_engenharia, cmbfamilia_engenharia, 0, 0, 0, cmbun_engenharia, Cmb_un_com_engenharia, 0, True, False, True, False, IIf(Opt_outros.Value = True, 4, IIf(Opt_terceiros.Value = True, 5, 0)), IIf(Opt_terceiros.Value = True, "S", "P"), "", 0, 0, 0, "", 0, "", "")
Else
    txtdesenho_qualidade = FunCriaNovoProdServ(False, "codmanual = 'False' and (subtipoitem = 0 or subtipoitem = 1 or subtipoitem = 4 or subtipoitem = 5)", txtdesenho_qualidade, cmbReferencia_qualidade, 0, txtdescricao_Qualidade, txtdescricao_Qualidade, cmbfamilia_qualidade, 0, 0, 0, cmbun_qualidade, Cmb_un_com_qualidade, 0, True, False, True, True, 4, "P", "", 0, 0, 0, "", 0, "", "")
End If
If SSTab1.Tab = 1 Then txtIDproduto_engenharia = Codproduto Else txtIDproduto_qualidade = Codproduto

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovoProdutoManualEngQualidade()
On Error GoTo tratar_erro

If SSTab1.Tab = 1 Then
    txtdesenho_engenharia = FunCriaNovoProdServ(True, "", txtdesenho_engenharia, txtReferencia_engenharia, 0, txtdescricao_engenharia, txtdescricao_engenharia, cmbfamilia_engenharia, 0, 0, 0, cmbun_engenharia, Cmb_un_com_engenharia, 0, True, False, True, False, IIf(Opt_outros.Value = True, 4, IIf(Opt_terceiros.Value = True, 5, 0)), IIf(Opt_terceiros.Value = True, "S", "P"), "", 0, 0, 0, "", 0, "", "")
Else
    txtdesenho_qualidade = FunCriaNovoProdServ(True, "", txtdesenho_qualidade, cmbReferencia_qualidade, 0, txtdescricao_Qualidade, txtdescricao_Qualidade, cmbfamilia_qualidade, 0, 0, 0, cmbun_qualidade, Cmb_un_com_qualidade, 0, True, False, True, True, 4, "P", "", 0, 0, 0, "", 0, "", "")
End If
If SSTab1.Tab = 1 Then txtIDproduto_engenharia = Codproduto Else txtIDproduto_qualidade = Codproduto

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcVerificaAcessos()
On Error GoTo tratar_erro

Acesso = False
Set TBAcessos = CreateObject("adodb.recordset")
TBAcessos.Open "select * from Acessos where IDUsuario = " & pubIDUsuario & " and Acesso = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBAcessos.EOF = False Then Acesso = True
TBAcessos.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCalculaExecucao()
On Error GoTo tratar_erro

txtPreparacao_processos.PromptInclude = True
txtExecucao_processos.PromptInclude = False
If Len(txtExecucao_processos.Text) < 7 Then
    txtExecucao_processos.PromptInclude = True
    txtTotalHora_processos = ""
    Exit Sub
End If
txtExecucao_processos.PromptInclude = True
If txtExecucao_processos > "023:59:59" Then
    ProcFormataHora (txtExecucao_processos)
    Familiatext = DataResultado
    TotalGeral = FunCalculaSegPC(Familiatext, txtPecaHora_processos)
Else
    If txtPecaHora_processos <> "" Then
        TotalGeral = FunCalculaSegPC(txtExecucao_processos, txtPecaHora_processos)
    End If
End If
If txtErro <> "" And txtErro <> "0" Then
    quantnovo = (TotalGeral * txtErro) / 100
    TotalGeral = TotalGeral + quantnovo
End If
Texto = FormataTempo(TotalGeral)
txtTotalHora_processos = Texto
ProcCalculamaquina
txtExecucao_processos.PromptInclude = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub Procentrada()
On Error GoTo tratar_erro

TempoPreparacao = ""
TempoExecucao = ""

'Rotina de tranformacao de preparacao
If IsNull(TBAbrir!Preparacao) = False And TBAbrir!Preparacao <> "___:__:__" Then
    ProcFormataHora (TBAbrir!Preparacao)
    TempoPreparacao = HoraTotal
    If Len(TempoPreparacao) = 8 Then TempoPreparacao = "0" & TempoPreparacao
End If

'Rotina de tranformacao de execucao
If IsNull(TBAbrir!Execucao) = False And TBAbrir!Execucao <> "___:__:__" Then
    ProcFormataHora (TBAbrir!Execucao)
    TempoExecucao = HoraTotal
    If Len(TempoExecucao) = 8 Then TempoExecucao = "0" & TempoExecucao
End If
        
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtValorHoraPrep_Processos_Change()
On Error GoTo tratar_erro

If txtValorHoraPrep_Processos <> "" Then
    VerifNumero = txtValorHoraPrep_Processos
    ProcVerificaNumero
    If VerifNumero = False Then
        txtValorHoraPrep_Processos = ""
        txtValorHoraPrep_Processos.SetFocus
        Exit Sub
    End If
    ProcCalculamaquina
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtValorHoraPrep_Processos_LostFocus()
On Error GoTo tratar_erro

txtValorHoraPrep_Processos = Format(txtValorHoraPrep_Processos, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar6_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcNovo_Engenharia_Norma
    Case 2: ProcSalvar_Engenharia_Norma
    Case 3: ProcExcluir_engenharia_Norma
    'Case 5: ProcAjuda
    Case 6: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar4_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcNovo_Engenharia_Prod
    Case 2: ProcSalvar_Engenharia_Prod
    Case 3: ProcExcluir_engenharia_prod
    Case 4: ProcCopiar_engenharia
    'Case 6: ProcAjuda
    Case 7: ProcSair
End Select

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
    Case 3: ProcSalvar
    Case 4: ProcImprimir
    Case 5: ProcAnterior
    Case 6: ProcProximo
    Case 7: procFiltrarTodos
    Case 8: ProcCopiar
    Case 9: ProcRevisao
    Case 10: ProcImpostos
    Case 11: ProcProduto_processo
    Case 12: ProcPrazos
    Case 13: ProcAtualizar
    'Case 15: ProcAjuda
    Case 16: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar5_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcNovo_Engenharia_Checklist
    Case 2: ProcSalvar_Engenharia_CheckList
    Case 3: ProcExcluir_engenharia_CheckList
    'Case 5: ProcAjuda
    Case 6: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar7_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcNovo_Processos
    Case 2: ProcSalvar_processos
    Case 3: ProcExcluir_processos
    Case 4: ProcImprimir
    Case 5: ProcCopiar_processo
    Case 6: ProcFerramentas
    Case 7: procValidar
    'Case 9: ProcAjuda
    Case 10: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar8_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcNovo_PCP_Checklist
    Case 2: ProcSalvar_PCP_CheckList
    Case 3: ProcExcluir_PCP_CheckList
    'Case 5: ProcAjuda
    Case 6: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar3_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case SSTab1.Tab
    Case 1: 'Engenharia
        Select Case ButtonIndex
            Case 1: ProcSalvar_Engenharia
            Case 2: ProcExcluir_Engenharia
            Case 3: ProcImprimir
            Case 4: ProcAnterior
            Case 5: ProcProximo
            Case 6: procValidar
            'Case 8: ProcAjuda
            Case 9: ProcSair
        End Select
    Case 3: 'PCP
        Select Case ButtonIndex
            Case 1: ProcSalvar_PCP
            Case 2: ProcExcluir_PCP
            Case 3: ProcImprimir
            Case 4: ProcAnterior
            Case 5: ProcProximo
            Case 6: procValidar
            'Case 8: ProcAjuda
            Case 9: ProcSair
        End Select
    Case 5: 'Compras
        Select Case ButtonIndex
            Case 1: ProcSalvar_Compras
            Case 2: ProcExcluir_Compras
            Case 3: ProcImprimir
            Case 4: ProcAnterior
            Case 5: ProcProximo
            Case 6: procValidar
            'Case 8: ProcAjuda
            Case 9: ProcSair
        End Select
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar2_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case SSTab1.Tab
    Case 2: 'Processo
        Select Case ButtonIndex
            Case 1: ProcNovo_Processos_Item
            Case 2: ProcSalvar_Processos_item
            Case 3: ProcExcluir_Processos_item
            Case 4: ProcImprimir
            Case 5: ProcAnterior
            Case 6: ProcProximo
            Case 7: procValidar
            'Case 9: ProcAjuda
            Case 10: ProcSair
        End Select
    Case 4: 'Qualidade
        Select Case SSTab_qualidade.Tab
            Case 0: 'Instrumentos
                Select Case ButtonIndex
                    Case 1: ProcNovo_Instrumento
                    Case 2: ProcSalvar_Instrumento
                    Case 3: ProcExcluir_Instrumento
                    Case 4: ProcImprimir
                    Case 5: ProcAnterior
                    Case 6: ProcProximo
                    Case 7: procValidar
                    'Case 9: ProcAjuda
                    Case 10: ProcSair
                End Select
            Case 1: 'Check-list
                Select Case ButtonIndex
                    Case 1: ProcNovo_Qualidade_Checklist
                    Case 2: ProcSalvar_Qualidade_CheckList
                    Case 3: ProcExcluir_Qualidade_CheckList
                    Case 4: ProcImprimir
                    Case 5: ProcAnterior
                    Case 6: ProcProximo
                    Case 7: procValidar
                    'Case 9: ProcAjuda
                    Case 10: ProcSair
                End Select
        End Select
    Case 5: 'Compras - Check-list
        Select Case ButtonIndex
            Case 1: ProcNovo_Compras_Checklist
            Case 2: ProcSalvar_Compras_CheckList
            Case 3: ProcExcluir_Compras_CheckList
            Case 4: ProcImprimir
            Case 5: ProcAnterior
            Case 6: ProcProximo
            Case 7: procValidar
            'Case 9: ProcAjuda
            Case 10: ProcSair
        End Select
    Case 6: 'Documentos
        Select Case ButtonIndex
            Case 1: procNovo_doc
            Case 2: procSalvar_doc
            Case 3: procExcluir_doc
            Case 4: ProcImprimir
            Case 5: ProcAnterior
            Case 6: ProcProximo
            Case 7: procValidar
            'Case 9: ProcAjuda
            Case 10: ProcSair
        End Select
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAtualizar()
On Error GoTo tratar_erro

If InputBox("Informe a senha para liberar.") = "280362A" Then
    If USMsgBox("Deseja realmente atualizar os produtos do processo?", vbYesNo, "CAPRIND v5.0") = vbYes Then
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from vendas_analise order by ID", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            PBLista.Min = 0
            PBLista.Max = TBAbrir.RecordCount
            PBLista.Value = 1
            Contador = 0
            Do While TBAbrir.EOF = False
                Set TBGravar = CreateObject("adodb.recordset")
                TBGravar.Open "Select * from Vendas_analise_ProdutosProcessos where ID_analise = " & TBAbrir!ID, Conexao, adOpenKeyset, adLockOptimistic
                If TBGravar.EOF = True Then
                    TBGravar.AddNew
                    TBGravar!id_analise = TBAbrir!ID
                    TBGravar!Codproduto = TBAbrir!IDProduto
                    TBGravar!Codinterno = IIf(IsNull(TBAbrir!Codinterno), "", TBAbrir!Codinterno)
                    TBGravar!Referencia = IIf(IsNull(TBAbrir!N_referencia), "", TBAbrir!N_referencia)
                    TBGravar!Un = IIf(IsNull(TBAbrir!Unidade), "", TBAbrir!Unidade)
                    TBGravar!Descricao = IIf(IsNull(TBAbrir!Descricao), "", TBAbrir!Descricao)
                    TBGravar!Familia = IIf(IsNull(TBAbrir!Familia), "", TBAbrir!Familia)
                    TBGravar!Produto_analise = True
                    TBGravar!Qtde = IIf(IsNull(TBAbrir!qtde_solicitada), 0, TBAbrir!qtde_solicitada)
                    TBGravar.Update
                    
                    Set TBFI = CreateObject("adodb.recordset")
                    TBFI.Open "select * from vendas_analise_setores where idanalise = " & TBAbrir!ID & " and id_processo_item is null and setor = 'PROCESSOS'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBFI.EOF = False Then
                        Do While TBFI.EOF = False
                            TBFI!ID_processo_item = TBGravar!ID
                            TBFI.Update
                            TBFI.MoveNext
                        Loop
                    End If
                    TBFI.Close
                End If
                TBGravar.Close
                TBAbrir.MoveNext
                Contador = Contador + 1
                PBLista.Value = Contador
            Loop
        End If
        
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "select * from vendas_analise_setores where Setor = 'PROCESSOS' and (Preparacao is not null or Execucao is not null) order by IDAnalise", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            PBLista.Min = 0
            PBLista.Max = TBAbrir.RecordCount
            PBLista.Value = 1
            Contador = 0
            Do While TBAbrir.EOF = False
                ProcFormataHora (TBAbrir!Preparacao)
                TBAbrir!Preparacao1 = HoraTotal
                
                ProcFormataHora (TBAbrir!Execucao)
                TBAbrir!Execucao1 = HoraTotal
                
                TBAbrir.Update
                TBAbrir.MoveNext
                Contador = Contador + 1
                PBLista.Value = Contador
            Loop
        End If
        TBAbrir.Close
        
        USMsgBox ("Atualização efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
        '==================================
        Modulo = "Outros/Análise crítica"
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

Public Sub ProcPrazos()
On Error GoTo tratar_erro

frmVendas_analise_prazos.Show 1
            
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub procValidar()
On Error GoTo tratar_erro

Select Case SSTab1.Tab
    Case 1: Formulario = "Outros/Análise crítica/Engenharia"
    Case 2: Formulario = "Outros/Análise crítica/Processos"
    Case 3: Formulario = "Outros/Análise crítica/Pcp"
    Case 4: Formulario = "Outros/Análise crítica/Qualidade"
    Case 5: Formulario = "Outros/Análise crítica/Compras"
End Select
frmValidar.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Function FunVerifStatusAnalise(Acao As String, MostrarMsg As Boolean) As Boolean
On Error GoTo tratar_erro

FunVerifStatusAnalise = True
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "select ID from Vendas_analise where ID = " & txtId & " and Status <> 'ABERTA EM ANALISE' and Status <> 'APROVADA'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    If MostrarMsg = True Then USMsgBox ("Só é permitido " & Acao & ", se a análise estiver com o status aberta em análise ou aprovada."), vbExclamation, "CAPRIND v5.0"
    FunVerifStatusAnalise = False
End If

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Function FunVerifValidSetorAnalise(Acao As String, NTab As Integer, MostrarMsg As Boolean) As Boolean
On Error GoTo tratar_erro

FunVerifValidSetorAnalise = True
Select Case NTab
    Case 1:
        TextoFiltro = "DtValidacao_Engenharia"
        MsgTexto = "a engenharia já foi validada"
    Case 2:
        TextoFiltro = "DtValidacao_Processo"
        MsgTexto = "o processo já foi validado"
    Case 3:
        TextoFiltro = "DtValidacao_PCP"
        MsgTexto = "o PCP já foi validado"
    Case 4:
        TextoFiltro = "DtValidacao_Qualidade"
        MsgTexto = "a qualidade já foi validada"
    Case 5:
        TextoFiltro = "DtValidacao_Compras"
        MsgTexto = "compras já foi validada"
End Select
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "select ID from Vendas_analise where ID = " & txtId & " and " & TextoFiltro & " IS NOT NULL", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    If MostrarMsg = True Then USMsgBox ("Não é permitido " & Acao & ", pois " & MsgTexto & "."), vbExclamation, "CAPRIND v5.0"
    FunVerifValidSetorAnalise = False
End If

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function
