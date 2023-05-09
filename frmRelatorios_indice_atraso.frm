VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form frmRelatorios_indice_atraso 
   BackColor       =   &H00E0E0E0&
   Caption         =   "PCP - Relatórios - Índice de atraso"
   ClientHeight    =   10035
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15360
   ClipControls    =   0   'False
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
      ScreenHeight    =   768
      ScreenWidth     =   1366
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
      Height          =   825
      Left            =   55
      TabIndex        =   29
      Top             =   9180
      Width           =   15195
      Begin VB.TextBox Txt_qtde_total_antecipada 
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
         Left            =   6322
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade total antecipada."
         Top             =   390
         Width           =   2550
      End
      Begin VB.TextBox Txt_qtde_total_atraso 
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
         Left            =   9393
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade total atraso."
         Top             =   390
         Width           =   2550
      End
      Begin VB.TextBox Txt_qtde_total_concluida 
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
         Left            =   3251
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade total concluída."
         Top             =   390
         Width           =   2550
      End
      Begin VB.TextBox Txt_indice 
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
         Left            =   12465
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   20
         TabStop         =   0   'False
         ToolTipText     =   "Índice."
         Top             =   390
         Width           =   2550
      End
      Begin VB.TextBox Txt_qtde_total_emitida 
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
         MaxLength       =   20
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade total emitida."
         Top             =   390
         Width           =   2550
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qtde. total antecipada"
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
         Left            =   6652
         TabIndex        =   34
         Top             =   180
         Width           =   1890
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qtde. total atraso"
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
         Left            =   9918
         TabIndex        =   33
         Top             =   180
         Width           =   1500
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Índice"
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
         Left            =   13478
         TabIndex        =   32
         Top             =   180
         Width           =   525
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qtde. total concluída"
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
         Left            =   3656
         TabIndex        =   31
         Top             =   180
         Width           =   1740
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qtde. total emitida"
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
         Left            =   660
         TabIndex        =   30
         Top             =   180
         Width           =   1590
      End
   End
   Begin VB.CheckBox Chk_componente 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Componente"
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
      Left            =   3750
      TabIndex        =   8
      Top             =   1110
      Value           =   1  'Checked
      Width           =   1275
   End
   Begin VB.CheckBox Chk_produto_final 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Produto final"
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
      Left            =   240
      TabIndex        =   6
      Top             =   1110
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CheckBox Chk_subconjunto 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Subconjunto"
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
      Left            =   1980
      TabIndex        =   7
      Top             =   1110
      Value           =   1  'Checked
      Width           =   1245
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   5910
      Left            =   60
      TabIndex        =   14
      Top             =   2970
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   10425
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
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "N"
         Text            =   "Ordem"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Tipo"
         Object.Width           =   970
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "Cód. interno"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Cod. de ref."
         Object.Width           =   2381
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Tag             =   "T"
         Text            =   "Descrição"
         Object.Width           =   6265
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Object.Tag             =   "T"
         Text            =   "Cliente"
         Object.Width           =   6265
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Object.Tag             =   "D"
         Text            =   "Prazo final"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   8
         Object.Tag             =   "D"
         Text            =   "Dt. concl."
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   9
         Object.Tag             =   "N"
         Text            =   "Índice"
         Object.Width           =   1764
      EndProperty
   End
   Begin MSComctlLib.ListView Lista1 
      Height          =   5915
      Left            =   60
      TabIndex        =   15
      Top             =   2965
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   10425
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
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "T"
         Object.Width           =   14296
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Object.Tag             =   "N"
         Text            =   "Qtde. emitida"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Object.Tag             =   "N"
         Text            =   "Qtde. concluída"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Object.Tag             =   "N"
         Text            =   "Qtde. antecipada"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Object.Tag             =   "N"
         Text            =   "Qtde. atrasada"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Object.Tag             =   "N"
         Text            =   "Índice"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.Frame Frame7 
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
      Height          =   1035
      Left            =   55
      TabIndex        =   28
      Top             =   1350
      Width           =   1695
      Begin VB.OptionButton Opt_individual 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Individual"
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
         ToolTipText     =   "0"
         Top             =   330
         Value           =   -1  'True
         Width           =   1185
      End
      Begin VB.OptionButton Opt_comparativo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Comparativo"
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
         TabIndex        =   1
         Top             =   600
         Width           =   1425
      End
   End
   Begin VB.Frame Frame5 
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
      Height          =   1035
      Left            =   1770
      TabIndex        =   27
      Top             =   1350
      Width           =   1455
      Begin VB.OptionButton optDetalhado 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Detalhado"
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
         TabIndex        =   2
         Top             =   330
         Value           =   -1  'True
         Width           =   1185
      End
      Begin VB.OptionButton optResumido 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Resumido"
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
         TabIndex        =   3
         Top             =   600
         Width           =   1155
      End
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   55
      TabIndex        =   36
      Top             =   8910
      Width           =   11775
      _ExtentX        =   20770
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
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   35
      Top             =   0
      Width           =   15195
      _ExtentX        =   26802
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft2     =   40
      ButtonTop2      =   2
      ButtonWidth2    =   51
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
      ButtonLeft3     =   93
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft4     =   97
      ButtonTop4      =   2
      ButtonWidth4    =   36
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft5     =   135
      ButtonTop5      =   2
      ButtonWidth5    =   26
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
      ButtonLeft6     =   163
      ButtonTop6      =   2
      ButtonWidth6    =   24
      ButtonHeight6   =   24
      ButtonUseMaskColor6=   0   'False
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   11610
         Top             =   90
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmRelatorios_indice_atraso.frx":0000
         Count           =   1
      End
   End
   Begin VB.Frame Frame4 
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
      Height          =   1035
      Left            =   3240
      TabIndex        =   24
      Top             =   1350
      Width           =   12015
      Begin VB.ComboBox cmbTexto 
         BackColor       =   &H00FFFFFF&
         Height          =   330
         ItemData        =   "frmRelatorios_indice_atraso.frx":2DFA
         Left            =   4080
         List            =   "frmRelatorios_indice_atraso.frx":2DFC
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         ToolTipText     =   "Texto para pesquisa."
         Top             =   480
         Width           =   7755
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
         ItemData        =   "frmRelatorios_indice_atraso.frx":2DFE
         Left            =   180
         List            =   "frmRelatorios_indice_atraso.frx":2E14
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         ToolTipText     =   "Opções para filtro."
         Top             =   480
         Width           =   3885
      End
      Begin VB.Label Label8 
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
         Left            =   1702
         TabIndex        =   26
         Top             =   270
         Width           =   840
      End
      Begin VB.Label Label9 
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
         Left            =   7222
         TabIndex        =   25
         Top             =   270
         Width           =   1470
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Height          =   585
      Left            =   55
      TabIndex        =   21
      Top             =   2370
      Width           =   15195
      Begin MSComCtl2.DTPicker msk_fltFim 
         Height          =   315
         Left            =   13710
         TabIndex        =   13
         ToolTipText     =   "Data final."
         Top             =   180
         Width           =   1305
         _ExtentX        =   2302
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
         Format          =   489422849
         CurrentDate     =   39057
      End
      Begin MSComCtl2.DTPicker msk_fltInicio 
         Height          =   315
         Left            =   11610
         TabIndex        =   12
         ToolTipText     =   "Data inicio."
         Top             =   180
         Width           =   1305
         _ExtentX        =   2302
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
         Format          =   489422849
         CurrentDate     =   39057
      End
      Begin VB.OptionButton optData 
         BackColor       =   &H00E0E0E0&
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
         Height          =   195
         Left            =   180
         TabIndex        =   9
         Top             =   240
         Value           =   -1  'True
         Width           =   1185
      End
      Begin VB.OptionButton optData2 
         BackColor       =   &H00E0E0E0&
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
         Height          =   195
         Left            =   1800
         TabIndex        =   10
         Top             =   240
         Width           =   1425
      End
      Begin VB.OptionButton optData3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Dt. conclusão"
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
         Left            =   3600
         TabIndex        =   11
         Top             =   240
         Width           =   1425
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "De :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   11220
         TabIndex        =   23
         Top             =   180
         Width           =   300
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Até :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   13260
         TabIndex        =   22
         Top             =   180
         Width           =   360
      End
   End
   Begin VB.Label Lbl_relatorio 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Registros encontrados: 0000 - 00:00:00"
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
      Left            =   11940
      TabIndex        =   37
      Top             =   8940
      Width           =   3315
   End
End
Attribute VB_Name = "frmRelatorios_indice_atraso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DataBanco As String 'OK
Dim DataTexto As String 'OK

Private Sub Chk_produto_final_Click()
On Error GoTo tratar_erro

ProcLimparListaeCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_componente_Click()
On Error GoTo tratar_erro

ProcLimparListaeCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_subconjunto_Click()
On Error GoTo tratar_erro

ProcLimparListaeCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbTexto_Click()
On Error GoTo tratar_erro

ProcLimparListaeCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro

If Opt_individual.Value = True Then
    If optDetalhado.Value = True Then
        If Lista.ListItems.Count = 0 Then Exit Sub
    Else
        If Lista1.ListItems.Count = 0 Then Exit Sub
    End If
Else
    If Lista1.ListItems.Count = 0 Then Exit Sub
End If
Vendas_Relatorio_Historico = False
Vendas_Relatorio_IndiceAtraso = False
Vendas_Relatorio_Comissao = False
Compras_Relatorio_IndiceAtraso = False
PCP_relatorios_indice_atraso = True
Manutencao_Relatorio_Historico = False
FrmMenu_impressao_padrao.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

ProcExcluirDadosProducaoRelatorios
ProcExcluirDadosProducaoRelatoriosTotal
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyF2: ProcFiltrar
    Case vbKeyF5: ProcImprimir
    'Case vbKeyF1: ProcAjuda
    Case vbKeyEscape: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaLista()
On Error GoTo tratar_erro

Posicao = 0
Lista.ListItems.Clear
Lista1.ListItems.Clear
If TBLISTA.EOF = False Then
    Posicao = TBLISTA.RecordCount
    TBLISTA.MoveLast
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    contador = 0
    TBLISTA.MoveFirst
    Do While TBLISTA.EOF = False
        If optDetalhado.Value = True Then
            With Lista.ListItems
                .Add , , TBLISTA!ID
                .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Ordem), "", TBLISTA!Ordem)
                .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Data4), "", TBLISTA!Data4)
                .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Totalhsutil), "", TBLISTA!Totalhsutil)
                .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!Data1), "", TBLISTA!Data1)
                .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!Data2), "", TBLISTA!Data2)
                '.Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA!Execucaoprev), "", TBLISTA!Execucaoprev)
                '.Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA!Totalhsprev), "", TBLISTA!Totalhsprev)
                .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA!Data3), "", TBLISTA!Data3)
                .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA!Data), "", Format(TBLISTA!Data, "dd/mm/yy"))
                .Item(.Count).SubItems(8) = IIf(IsNull(TBLISTA!Data5), "", Format(TBLISTA!Data5, "dd/mm/yy"))
                .Item(.Count).SubItems(9) = IIf(IsNull(TBLISTA!Eficiencia), "", Format(TBLISTA!Eficiencia, "###,##0.00") & "%")
            End With
        Else
            With Lista1.ListItems
                .Add , , TBLISTA!ID
                .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!maquina), "", TBLISTA!maquina)
                .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!QtdePrev), "", TBLISTA!QtdePrev)
                .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!qtdeOK), "", TBLISTA!qtdeOK)
                .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!qtdeNC), "", TBLISTA!qtdeNC)
                .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!Qtdetotalprod), "", TBLISTA!Qtdetotalprod)
                .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA!Eficiencia), "", Format(TBLISTA!Eficiencia, "###,##0.00") & "%")
            End With
        End If
        TBLISTA.MoveNext
        contador = contador + 1
        PBLista.Value = contador
    Loop
End If

Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Producao_Relatorios_Total where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    Txt_qtde_total_emitida = TBLISTA!QtdePrevista 'emitida
    Txt_qtde_total_concluida = TBLISTA!QtdeProduzida 'concluida
    Txt_qtde_total_antecipada = TBLISTA!qtdeNC 'antecipado
    Txt_qtde_total_atraso = TBLISTA!QtdeOrdem 'atraso
    Txt_indice = Format(TBLISTA!TotalEficiencia, "###,##0.00") & "%"
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCamposTotais()
On Error GoTo tratar_erro

Lbl_relatorio.Caption = "Registros encontrados: 0000 - 00:00:00"
Txt_qtde_total_emitida = ""
Txt_qtde_total_concluida = ""
Txt_qtde_total_antecipada = ""
Txt_qtde_total_atraso = ""
Txt_indice = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15195, 5, True

Formulario = "PCP/Relatórios/Índice de atraso"
Direitos
ProcLimpaVariaveisPrincipais
msk_fltInicio.Value = Date
msk_fltFim.Value = Date
cmbfiltrarpor.Text = "Código interno"

ProcRemoveObjetosResize Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

Formulario = "PCP/Relatórios/Índice de atraso"
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

Private Sub Lista1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView Lista1, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub msk_fltFim_Click()
On Error GoTo tratar_erro

ProcLimparListaeCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbfiltrarpor_Click()
On Error GoTo tratar_erro

ProcCarregaComboTexto

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaComboTexto()
On Error GoTo tratar_erro

ProcLimparListaeCampos

If Opt_individual.Value = True Then
    cmbTexto.Clear
    Texto = ""
    If cmbfiltrarpor = "Família" Then
        ProcCarregaComboFamilia cmbTexto, "familia <> 'Null' and vendas = 'True'", False
    Else
        Select Case cmbfiltrarpor
            Case "Cliente": Texto = "Cliente"
            Case "Código de referência": Texto = "N_Referencia"
            Case "Código interno": Texto = "desenho"
            Case "Descrição": Texto = "Produto"
            Case "Ordem": Texto = "Ordem"
        End Select
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select " & Texto & " as NomeCampo1 from Producao where " & Texto & " is not null group by " & Texto, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Do While TBAbrir.EOF = False
                If TBAbrir!NomeCampo1 <> "" Then cmbTexto.AddItem TBAbrir!NomeCampo1
                TBAbrir.MoveNext
            Loop
        End If
    End If
    TBAbrir.Close
End If
Lista1.ColumnHeaders(2).Text = cmbfiltrarpor

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

Acao = "filtrar"
If Chk_produto_final.Value = 0 And Chk_subconjunto.Value = 0 And Chk_componente.Value = 0 Then
    NomeCampo = "uma das opções"
    ProcVerificaAcao
    Exit Sub
End If
If Opt_individual.Value = True And cmbTexto = "" Then
    NomeCampo = "o texto para pesquisa"
    ProcVerificaAcao
    cmbTexto.SetFocus
    Exit Sub
End If
With msk_fltFim
    If FunVerificaDataFinal(msk_fltInicio.Value, .Value) = False Then
        .Value = Date
        .SetFocus
        Exit Sub
    End If
End With
Inicio = Time
ProcAbrirTabelas
If Permitido = True Then ProcGravarTotalizacoes
Set TBLISTA = CreateObject("adodb.recordset")
If Opt_individual.Value = True And optDetalhado.Value = True Then
    TBLISTA.Open "Select * from Producao_Relatorios where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "' order by Data, Maquina", Conexao, adOpenKeyset, adLockOptimistic
Else
    TBLISTA.Open "Select * from Producao_Relatorios where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "' order by Maquina", Conexao, adOpenKeyset, adLockOptimistic
End If
ProcCarregaLista

intervalo = Time
ElapsedTime (intervalo - Inicio)
Lbl_relatorio.Caption = "Registros encontrados: " & FunTamanhoTextoZeroEsq(Posicao, 4) & " - " & HoraTotal

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimparListaeCampos()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista1.ListItems.Clear
ProcLimpaCamposTotais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAbrirTabelas()
On Error GoTo tratar_erro

'Deleta registros e adiciona novos
ProcExcluirDadosProducaoRelatorios
ProcExcluirDadosProducaoRelatoriosTotal

If Chk_produto_final.Value = 1 And Chk_subconjunto.Value = 1 And Chk_componente.Value = 1 Then 'Prod + Sub + Com
    TipoFiltro = "and (Producao.tipo = 'E' or Producao.tipo = 'M' or Producao.tipo = 'F')"
ElseIf Chk_produto_final.Value = 1 And Chk_subconjunto.Value = 1 And Chk_componente.Value = 0 Then 'Prod + Sub
        TipoFiltro = "and (Producao.tipo = 'E' or Producao.tipo = 'M')"
    ElseIf Chk_produto_final.Value = 1 And Chk_subconjunto.Value = 0 And Chk_componente.Value = 1 Then 'Prod + Com
            TipoFiltro = "and (Producao.tipo = 'E' or Producao.tipo = 'F')"
    ElseIf Chk_produto_final.Value = 0 And Chk_subconjunto.Value = 1 And Chk_componente.Value = 1 Then 'Sub + Com
            TipoFiltro = "and (Producao.tipo = 'M' or Producao.tipo = 'F')"
        ElseIf Chk_produto_final.Value = 1 And Chk_subconjunto.Value = 0 And Chk_componente.Value = 0 Then 'Prod
                TipoFiltro = "and Producao.tipo = 'E'"
            ElseIf Chk_produto_final.Value = 0 And Chk_subconjunto.Value = 1 And Chk_componente.Value = 0 Then 'Sub
                    TipoFiltro = "and Producao.tipo = 'M'"
                Else 'Com
                    TipoFiltro = "and Producao.tipo = 'F'"
End If

If optData.Value = True Then DataFiltro = "producao.data"
If optData2.Value = True Then DataFiltro = "producao.PrazoEntrega"
If optData3.Value = True Then DataFiltro = "producao.dataentrega"

If cmbfiltrarpor = "Família" Then
    INNERJOINTEXTO = "O.*, P.classe from producao O INNER JOIN projproduto P ON O.desenho = P.Desenho"
Else
    INNERJOINTEXTO = "* from producao"
End If
Select Case cmbfiltrarpor
    Case "Ordem": TextoFiltro = "Ordem"
    Case "Código interno": TextoFiltro = "desenho"
    Case "Código de referência": TextoFiltro = "N_Referencia"
    Case "Descrição": TextoFiltro = "Produto"
    Case "Família": TextoFiltro = "projproduto.Classe"
    Case "Cliente": TextoFiltro = "Cliente"
End Select
If Opt_individual.Value = True Then
    TextoFiltro1 = TextoFiltro & " = '" & cmbTexto & "' and"
    Ordenar = DataFiltro
Else
    TextoFiltro1 = ""
    Ordenar = TextoFiltro
End If
Set TBCarteira = CreateObject("adodb.recordset")
TBCarteira.Open "Select " & INNERJOINTEXTO & " where " & TextoFiltro1 & " " & DataFiltro & " Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "' " & TipoFiltro & " order by " & Ordenar, Conexao, adOpenKeyset, adLockOptimistic
ProcFiltrar1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar1()
On Error GoTo tratar_erro

If TBCarteira.EOF = False Then
    Permitido = True
    TBCarteira.MoveLast
    PBLista.Min = 0
    PBLista.Max = TBCarteira.RecordCount
    PBLista.Value = 1
    contador = 0
    TBCarteira.MoveFirst
    Do While TBCarteira.EOF = False
        Set TBProdutividade = CreateObject("adodb.recordset")
        If Opt_individual.Value = True And optDetalhado.Value = True Then
            TBProdutividade.Open "Select * from Producao_Relatorios", Conexao, adOpenKeyset, adLockOptimistic
            ProcEnviaDadosDetalhado
        Else
            Select Case cmbfiltrarpor
                Case "Código interno": TBProdutividade.Open "Select * from Producao_Relatorios where maquina = '" & TBCarteira!Desenho & "' and Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
                Case "Código de referência": TBProdutividade.Open "Select * from Producao_Relatorios where maquina = '" & TBCarteira!N_referencia & "' and Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
                Case "Descrição": TBProdutividade.Open "Select * from Producao_Relatorios where maquina = '" & TBCarteira!Produto & "' and Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
                Case "Família": TBProdutividade.Open "Select * from Producao_Relatorios where maquina = '" & TBCarteira!Classe & "' and Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
                Case "Cliente": TBProdutividade.Open "Select * from Producao_Relatorios where maquina = '" & TBCarteira!Cliente & "' and Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
                Case "Ordem": TBProdutividade.Open "Select * from Producao_Relatorios where maquina = '" & TBCarteira!Ordem & "' and Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
            End Select
            ProcEnviaDadosResumido
        End If
        TBCarteira.MoveNext
        contador = contador + 1
        PBLista.Value = contador
    Loop
End If
TBCarteira.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviaDadosDetalhado()
On Error GoTo tratar_erro

TBProdutividade.AddNew
TBProdutividade!Ordem = TBCarteira!Ordem 'Ordem

Select Case TBCarteira!Tipo
    Case "E": Tipo = "P"
    Case "M": Tipo = "S"
    Case "F": Tipo = "C"
End Select
TBProdutividade!Data4 = Tipo 'Tipo

TBProdutividade!Data = TBCarteira!PrazoEntrega 'Prazo final
'TBProdutividade!Execucaoprev = TBCarteira!Lista 'pedido interno
'TBProdutividade!Totalhsprev = TBCarteira!Revisao 'revisão
TBProdutividade!Totalhsutil = TBCarteira!Desenho 'Cod interno
TBProdutividade!Data1 = TBCarteira!N_referencia 'cod referencia
TBProdutividade!Data2 = TBCarteira!Produto 'descrição
TBProdutividade!Data3 = TBCarteira!Cliente
TBProdutividade!Data5 = TBCarteira!DataEntrega 'Data conclusão
TBProdutividade!QtdePrev = 1 'emitida
If TBCarteira!Concluida = True Then TBProdutividade!qtdeOK = 1 'concluida
If TBCarteira!PrazoEntrega > TBCarteira!DataEntrega Then TBProdutividade!qtdeNC = 1 'Antecipada
If TBCarteira!PrazoEntrega < TBCarteira!DataEntrega Then TBProdutividade!Qtdetotalprod = 1 'Antecipada
If TBProdutividade!qtdeOK <> 0 Then TBProdutividade!Eficiencia = (TBProdutividade!Qtdetotalprod / TBProdutividade!qtdeOK) * 100
TBProdutividade!Responsavel = pubUsuario
TBProdutividade!Modulo = Formulario
TBProdutividade!maquina = cmbTexto
TBProdutividade.Update
TBProdutividade.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviaDadosResumido()
On Error GoTo tratar_erro

If TBProdutividade.EOF = True Then TBProdutividade.AddNew
If Opt_individual.Value = True Then
    Texto = cmbTexto
Else
    Select Case cmbfiltrarpor
        Case "Código interno": Texto = TBCarteira!Desenho
        Case "Código de referência": Texto = TBCarteira!N_referencia
        Case "Descrição": Texto = TBCarteira!Produto
        Case "Família": Texto = TBCarteira!Classe
        Case "Cliente": Texto = TBCarteira!Cliente
        Case "Ordem": Texto = TBCarteira!Ordem
    End Select
End If
TBProdutividade!maquina = Texto
TBProdutividade!Responsavel = pubUsuario
TBProdutividade!Modulo = Formulario
TBProdutividade!QtdePrev = TBProdutividade!QtdePrev + 1 'emitida
If TBCarteira!Concluida = True Then TBProdutividade!qtdeOK = TBProdutividade!qtdeOK + 1 'concluida
If TBCarteira!PrazoEntrega > TBCarteira!DataEntrega Then TBProdutividade!qtdeNC = TBProdutividade!qtdeNC + 1 'Antecipada
If TBCarteira!PrazoEntrega < TBCarteira!DataEntrega Then TBProdutividade!Qtdetotalprod = TBProdutividade!Qtdetotalprod + 1 'Antecipada
If TBProdutividade!qtdeOK <> 0 Then TBProdutividade!Eficiencia = (TBProdutividade!Qtdetotalprod / TBProdutividade!qtdeOK) * 100
TBProdutividade.Update
TBProdutividade.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcGravarTotalizacoes()
On Error GoTo tratar_erro

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Producao_Relatorios_Total where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = True Then TBAbrir.AddNew
TBAbrir!Data_inicial = msk_fltInicio.Value
TBAbrir!Data_final = msk_fltFim.Value
If Opt_individual.Value = True Then TBAbrir!Texto = cmbfiltrarpor & ") : " & cmbTexto Else TBAbrir!Texto = cmbfiltrarpor
TBAbrir!Responsavel = pubUsuario
TBAbrir!Modulo = Formulario

quantidade = 0
QTLOTE = 0
Quant = 0
Qtde = 0
Set TBproducao = CreateObject("adodb.recordset")
TBproducao.Open "Select Sum(QtdePrev) as quantidade, Sum(QtdeOK) as QTLOTE, Sum(QtdeNC) as Quant, Sum(Qtdetotalprod) as Qtde from Producao_Relatorios where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBproducao.EOF = False Then
    TBAbrir!QtdePrevista = IIf(IsNull(TBproducao!quantidade), 0, TBproducao!quantidade) 'emitida
    TBAbrir!QtdeProduzida = IIf(IsNull(TBproducao!QTLOTE), 0, TBproducao!QTLOTE) 'concluida
    TBAbrir!qtdeNC = IIf(IsNull(TBproducao!Quant), 0, TBproducao!Quant) 'antecipado
    TBAbrir!QtdeOrdem = IIf(IsNull(TBproducao!Qtde), 0, TBproducao!Qtde) 'atraso
End If
TBproducao.Close

If TBAbrir!QtdeProduzida <> 0 Then TBAbrir!TotalEficiencia = (TBAbrir!QtdeOrdem / TBAbrir!QtdeProduzida) * 100 Else TBAbrir!TotalEficiencia = 0
TBAbrir.Update
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub msk_fltInicio_Click()
On Error GoTo tratar_erro

ProcLimparListaeCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt_comparativo_Click()
On Error GoTo tratar_erro

If Opt_comparativo.Value = True Then
    optDetalhado.Enabled = False
    optResumido.Value = True
    cmbTexto.ListIndex = -1
    cmbTexto.Enabled = False
End If
ProcLimparListaeCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt_individual_Click()
On Error GoTo tratar_erro

If Opt_individual.Value = True Then
    optDetalhado.Value = True
    optDetalhado.Enabled = True
    cmbTexto.Enabled = True
    ProcCarregaComboTexto
End If
ProcLimparListaeCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optData_Click()
On Error GoTo tratar_erro

If optData.Value = 1 Then msk_fltInicio.SetFocus
ProcLimparListaeCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optData2_Click()
On Error GoTo tratar_erro

If optData2.Value = 1 Then msk_fltInicio.SetFocus
ProcLimparListaeCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optData3_Click()
On Error GoTo tratar_erro

If optData3.Value = 1 Then msk_fltInicio.SetFocus
ProcLimparListaeCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optDetalhado_Click()
On Error GoTo tratar_erro

If optDetalhado.Value = True Then
    Lista.Visible = True
    Lista1.Visible = False
End If
ProcLimparListaeCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optResumido_Click()
On Error GoTo tratar_erro

If optResumido.Value = True Then
    Lista.Visible = False
    Lista1.Visible = True
End If
ProcLimparListaeCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcFiltrar
    Case 2: ProcImprimir
    'Case 4: ProcAjuda
    Case 5: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
