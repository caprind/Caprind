VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmcqnc_abrir 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Qualidade - Não conformidade - Localizar"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10755
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   10755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame7 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ordem de retrabalho"
      Height          =   855
      Left            =   8635
      TabIndex        =   34
      Top             =   990
      Width           =   2085
      Begin VB.CheckBox chkApenas 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Apenas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   330
         TabIndex        =   36
         ToolTipText     =   "Trazer apenas ordens de retrabalho no filtro"
         Top             =   540
         Width           =   1725
      End
      Begin VB.CheckBox chkNaoIncluir 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Não incluir"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   330
         TabIndex        =   35
         ToolTipText     =   "Não incluir ordens de retrabalho no filtro"
         Top             =   300
         Width           =   1725
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Retrabalho"
      Height          =   855
      Left            =   6135
      TabIndex        =   31
      Top             =   990
      Width           =   2475
      Begin VB.CheckBox chkSemOrdem 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Sem ordens emitidas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   330
         TabIndex        =   33
         ToolTipText     =   "Filtrar não conformidades de retrabalho que ainda não foram criadas ordens de retrabalho"
         Top             =   300
         Width           =   1935
      End
      Begin VB.CheckBox chkComOrdem 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Com ordens emitidas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   330
         TabIndex        =   32
         ToolTipText     =   "Filtrar não conformidades de retrabalho que já foram criadas ordens de retrabalho"
         Top             =   540
         Width           =   1935
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Não conformidade"
      Height          =   855
      Left            =   55
      TabIndex        =   29
      Top             =   990
      Width           =   1965
      Begin VB.CheckBox Chk_analisada 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Analisada"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   330
         TabIndex        =   14
         Top             =   300
         Value           =   1  'Checked
         Width           =   1035
      End
      Begin VB.CheckBox Chk_nao_analisada 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Não analisada"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   330
         TabIndex        =   15
         Top             =   540
         Value           =   1  'Checked
         Width           =   1335
      End
   End
   Begin VB.CheckBox chkDataConclusao 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Conclusão ordem"
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
      Left            =   3180
      TabIndex        =   11
      Top             =   3570
      Width           =   1755
   End
   Begin VB.CheckBox chkEmissao 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Emissão ordem"
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
      Left            =   1530
      TabIndex        =   10
      Top             =   3570
      Width           =   1605
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ordens"
      Height          =   855
      Left            =   2045
      TabIndex        =   27
      Top             =   990
      Width           =   4065
      Begin VB.CheckBox chkEscopo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Escopo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   330
         TabIndex        =   18
         Top             =   540
         Value           =   1  'Checked
         Width           =   825
      End
      Begin VB.CheckBox chkSemEscopo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fora do escopo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2280
         TabIndex        =   19
         Top             =   540
         Value           =   1  'Checked
         Width           =   1425
      End
      Begin VB.CheckBox chkConcluida 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Concluidas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   330
         TabIndex        =   16
         Top             =   300
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkNaoConcluida 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Não concluidas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2280
         TabIndex        =   17
         Top             =   300
         Value           =   1  'Checked
         Width           =   1395
      End
   End
   Begin VB.CheckBox Chk_periodo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Data da NC"
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
      Left            =   240
      TabIndex        =   9
      Top             =   3570
      Width           =   1245
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   1515
      Left            =   55
      TabIndex        =   20
      Top             =   1800
      Width           =   10665
      Begin VB.CommandButton Cmd_salvar 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   9840
         Picture         =   "frmcqnc_abrir.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Salvar filtro para pesquisa (F3)."
         Top             =   1050
         Width           =   315
      End
      Begin VB.CommandButton Cmd_excluir 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   10170
         Picture         =   "frmcqnc_abrir.frx":0053
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Excluir filtro para pesquisa (F4)."
         Top             =   1050
         Width           =   315
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0E0&
         Height          =   510
         Left            =   5700
         TabIndex        =   28
         Top             =   210
         Width           =   4785
         Begin VB.OptionButton Optfim 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Fim frase"
            Height          =   255
            Left            =   2760
            TabIndex        =   7
            Top             =   180
            Width           =   1155
         End
         Begin VB.OptionButton Optinicio 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Início frase"
            Height          =   255
            Left            =   180
            TabIndex        =   5
            Top             =   180
            Value           =   -1  'True
            Width           =   1275
         End
         Begin VB.OptionButton Optmeio 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Meio frase"
            Height          =   255
            Left            =   1470
            TabIndex        =   6
            Top             =   180
            Width           =   1275
         End
         Begin VB.OptionButton optIgual 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Igual"
            Height          =   255
            Left            =   3930
            TabIndex        =   8
            Top             =   180
            Width           =   705
         End
      End
      Begin VB.ComboBox cmbfiltrarpor 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         ItemData        =   "frmcqnc_abrir.frx":0191
         Left            =   180
         List            =   "frmcqnc_abrir.frx":01B6
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Opções para filtro."
         Top             =   390
         Width           =   5385
      End
      Begin VB.TextBox txtTexto 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
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
         TabIndex        =   1
         ToolTipText     =   "Texto para pesquisa."
         Top             =   1050
         Width           =   9645
      End
      Begin VB.ComboBox cmbTexto 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
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
         ItemData        =   "frmcqnc_abrir.frx":0237
         Left            =   180
         List            =   "frmcqnc_abrir.frx":0259
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "Opções para filtro."
         Top             =   1050
         Visible         =   0   'False
         Width           =   9645
      End
      Begin VB.Label Label1 
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
         Left            =   4267
         TabIndex        =   22
         Top             =   840
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
         Left            =   2452
         TabIndex        =   21
         Top             =   180
         Width           =   840
      End
   End
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   7320
      Top             =   240
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frmcqnc_abrir.frx":02D3
      Count           =   1
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   26
      Top             =   0
      Width           =   10665
      _ExtentX        =   18812
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
      ButtonEnabled2  =   0   'False
      ButtonIconSize2 =   32
      ButtonAlignment2=   2
      ButtonType2     =   1
      ButtonStyle2    =   -1
      BeginProperty ButtonFont2 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState2    =   -1
      ButtonLeft2     =   40
      ButtonTop2      =   4
      ButtonWidth2    =   2
      ButtonHeight2   =   54
      ButtonUseMaskColor2=   0   'False
      ButtonCaption3  =   "Ajuda"
      ButtonEnabled3  =   0   'False
      ButtonIconSize3 =   32
      ButtonToolTipText3=   "Ajuda (F1)"
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
      ButtonLeft3     =   44
      ButtonTop3      =   2
      ButtonWidth3    =   36
      ButtonHeight3   =   21
      ButtonCaption4  =   "Sair"
      ButtonEnabled4  =   0   'False
      ButtonIconSize4 =   32
      ButtonToolTipText4=   "Sair (Esc)"
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
      ButtonLeft4     =   82
      ButtonTop4      =   2
      ButtonWidth4    =   26
      ButtonHeight4   =   21
      ButtonEnabled5  =   0   'False
      ButtonIconSize5 =   32
      BeginProperty ButtonFont5 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState5    =   5
      ButtonLeft5     =   110
      ButtonTop5      =   2
      ButtonWidth5    =   24
      ButtonHeight5   =   24
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   675
      Left            =   55
      TabIndex        =   23
      Top             =   3270
      Width           =   10665
      Begin MSComCtl2.DTPicker msk_fltFim 
         Height          =   315
         Left            =   9210
         TabIndex        =   13
         ToolTipText     =   "Data final."
         Top             =   210
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
         Format          =   197459969
         CurrentDate     =   39057
      End
      Begin MSComCtl2.DTPicker msk_fltInicio 
         Height          =   315
         Left            =   7320
         TabIndex        =   12
         ToolTipText     =   "Data inicio."
         Top             =   210
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
         Format          =   197853185
         CurrentDate     =   39057
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
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
         Height          =   285
         Left            =   6960
         TabIndex        =   25
         Top             =   240
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
         Height          =   285
         Left            =   8805
         TabIndex        =   24
         Top             =   240
         Width           =   360
      End
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   2115
      Left            =   60
      TabIndex        =   30
      Top             =   3960
      Visible         =   0   'False
      Width           =   10665
      _ExtentX        =   18812
      _ExtentY        =   3731
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
         Size            =   6.75
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
         Text            =   "Filtrar por"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Local da frase"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "Texto para pesquisa"
         Object.Width           =   11509
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Tag             =   "N"
         Text            =   "IDTexto"
         Object.Width           =   0
      EndProperty
   End
End
Attribute VB_Name = "frmcqnc_abrir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Chk_periodo_Click()
On Error GoTo tratar_erro

If Chk_periodo.Value = 1 Then
    chkEmissao.Value = 0
    chkDataConclusao.Value = 0
    Frame2.Enabled = True
    msk_fltInicio.SetFocus
Else
    Frame2.Enabled = False
    msk_fltInicio.Value = Date
    msk_fltFim.Value = Date
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub chkApenas_Click()
On Error GoTo tratar_erro

If chkApenas.Value = 1 Then chkNaoIncluir.Value = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub chkDataConclusao_Click()
On Error GoTo tratar_erro

If chkDataConclusao.Value = 1 Then
    Chk_periodo.Value = 0
    chkEmissao.Value = 0
    Frame2.Enabled = True
    msk_fltInicio.SetFocus
Else
    Frame2.Enabled = False
    msk_fltInicio.Value = Date
    msk_fltFim.Value = Date
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub chkEmissao_Click()
On Error GoTo tratar_erro

If chkEmissao.Value = 1 Then
    chkDataConclusao.Value = 0
    Chk_periodo.Value = 0
    Frame2.Enabled = True
    msk_fltInicio.SetFocus
Else
    Frame2.Enabled = False
    msk_fltInicio.Value = Date
    msk_fltFim.Value = Date
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub chkNaoIncluir_Click()
On Error GoTo tratar_erro

If chkNaoIncluir.Value = 1 Then chkApenas.Value = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub cmbfiltrarpor_Click()
On Error GoTo tratar_erro

txtTexto = ""
If cmbfiltrarpor = "Disposição" Or cmbfiltrarpor = "Grupo" Or cmbfiltrarpor = "Família" Then
    txtTexto.Visible = False
    With cmbTexto
        .Visible = True
        .Clear
        .AddItem ""
        If cmbfiltrarpor = "Família" Then
            ProcCarregaComboFamilia cmbTexto, "familia <> 'Null'", True
        ElseIf cmbfiltrarpor = "Grupo" Then
            ProcCarregaComboGrupoFamilia cmbTexto, "Grupo <> 'Null'", True
        Else
            .AddItem "Aprovado"
            .AddItem "Aprovado c / desvio"
            .AddItem "Devolver"
            .AddItem "Nada consta"
            .AddItem "Outros"
            .AddItem "Reaproveitar"
            .AddItem "Rejeitar"
            .AddItem "Retrabalhar"
            .AddItem "Selecionar"
        End If
    End With
Else
    txtTexto.Visible = True
    cmbTexto.Visible = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

Acao = "filtrar"
If Chk_analisada.Value = 0 And Chk_nao_analisada.Value = 0 Then
    NomeCampo = "uma das opções"
    ProcVerificaAcao
    Exit Sub
End If
If chkConcluida.Value = 0 And chkNaoConcluida.Value = 0 Then
    NomeCampo = "uma das opções de ordens concluidas ou não concluidas"
    ProcVerificaAcao
    Exit Sub
End If
If chkEscopo.Value = 0 And chkSemEscopo.Value = 0 Then
    NomeCampo = "uma das opções de escopo"
    ProcVerificaAcao
    Exit Sub
End If

With msk_fltFim
    If FunVerificaDataFinal(msk_fltInicio.Value, .Value) = False Then
        .Value = Date
        .SetFocus
        Exit Sub
    End If
End With

With frmcqnc
    .ListaFases.ListItems.Clear
    Filtro = ""
    FiltroRel = ""
    
    If Chk_analisada.Value = 1 And Chk_nao_analisada.Value = 1 Then
        TextoFiltroAnalisada = "(CQ.Analizada = 'True' or CQ.Analizada = 'False')"
        TextoFiltroAnalisadaRel = "({CQ_NC_FABRICA.Analizada} = True or {CQ_NC_FABRICA.Analizada} = False)"
    ElseIf Chk_analisada.Value = 1 And Chk_nao_analisada.Value = 0 Then
        TextoFiltroAnalisada = "CQ.Analizada = 'True'"
        TextoFiltroAnalisadaRel = "{CQ_NC_FABRICA.Analizada} = True"
    Else
        TextoFiltroAnalisada = "CQ.Analizada = 'False'"
        TextoFiltroAnalisadaRel = "{CQ_NC_FABRICA.Analizada} = False"
    End If
    
    TextoFiltroRetrabalho = ""
    TextoFiltroRetrabalhoRel = ""
    If chkNaoIncluir.Value = 1 Then
        TextoFiltroRetrabalho = " and (P.Retrabalho IS NULL or P.Retrabalho = 'False')"
        'TextoFiltroRetrabalhoRel = " and ({Producao.Retrabalho} = false or ISNULL({Producao.Retrabalho}))"
        TextoFiltroRetrabalhoRel = " and ISNULL({Producao.Retrabalho})"
    ElseIf chkApenas.Value = 1 Then
        TextoFiltroRetrabalho = " and P.Retrabalho = 'True'"
        TextoFiltroRetrabalhoRel = " and {Producao.Retrabalho} = true"
    End If
    
    TextoFiltroRetrabalho1 = ""
    TextoFiltroRetrabalhoRel1 = ""
    If chkSemOrdem.Value = 1 And chkComOrdem.Value = 0 Then
        TextoFiltroRetrabalho1 = " and CQ.PARECERCQ = 'Retrabalhar' and CQ.OrdemRetrabalho IS NULL"
        TextoFiltroRetrabalhoRel1 = " and {CQ_NC_FABRICA.PARECERCQ} = 'Retrabalhar' and ISNULL({CQ_NC_FABRICA.OrdemRetrabalho})"
    ElseIf chkSemOrdem.Value = 0 And chkComOrdem.Value = 1 Then
        TextoFiltroRetrabalho1 = " and CQ.PARECERCQ = 'Retrabalhar' and CQ.OrdemRetrabalho IS NOT NULL"
        TextoFiltroRetrabalhoRel1 = " and {CQ_NC_FABRICA.PARECERCQ} = 'Retrabalhar' and NOT(ISNULL({CQ_NC_FABRICA.OrdemRetrabalho}))"
    ElseIf chkSemOrdem.Value = 1 And chkComOrdem.Value = 1 Then
        TextoFiltroRetrabalho1 = " and CQ.PARECERCQ = 'Retrabalhar'"
        TextoFiltroRetrabalhoRel1 = " and {CQ_NC_FABRICA.PARECERCQ} = 'Retrabalhar'"
    End If
    TextoFiltroRetrabalho = TextoFiltroRetrabalho & TextoFiltroRetrabalho1
    TextoFiltroRetrabalhoRel = TextoFiltroRetrabalhoRel & TextoFiltroRetrabalhoRel1
    
    DataFiltro = ""
    DataFiltroRel = ""
    If Chk_periodo.Value = 1 Then
        DataFiltro = " and CQ.Data Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "'"
        DataFiltroRel = " and {CQ_NC_FABRICA.Data} >= Date(" & Year(msk_fltInicio.Value) & "," & Month(msk_fltInicio.Value) & "," & Day(msk_fltInicio.Value) & ") and {CQ_NC_FABRICA.Data} <= Date(" & _
                                    Year(msk_fltFim.Value) & "," & Month(msk_fltFim.Value) & "," & Day(msk_fltFim.Value) & ")"
    ElseIf chkDataConclusao.Value = 1 Then
        DataFiltro = " and P.Dataentrega Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "'"
        DataFiltroRel = " and {Producao.Dataentrega} >= Date(" & Year(msk_fltInicio.Value) & "," & Month(msk_fltInicio.Value) & "," & Day(msk_fltInicio.Value) & ") and {Producao.Dataentrega} <= Date(" & _
                                Year(msk_fltFim.Value) & "," & Month(msk_fltFim.Value) & "," & Day(msk_fltFim.Value) & ")"
    ElseIf chkEmissao.Value = 1 Then
        DataFiltro = " and P.Data Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "'"
        DataFiltroRel = " and {Producao.Data} >= Date(" & Year(msk_fltInicio.Value) & "," & Month(msk_fltInicio.Value) & "," & Day(msk_fltInicio.Value) & ") and {Producao.Data} <= Date(" & _
                                Year(msk_fltFim.Value) & "," & Month(msk_fltFim.Value) & "," & Day(msk_fltFim.Value) & ")"
    End If
    
    .dataTipo_CQNC = 0
    If Chk_periodo.Value = 1 Then
        .dataTipo_CQNC = 1
    ElseIf chkEmissao.Value = 1 Then
        .dataTipo_CQNC = 2
    ElseIf chkDataConclusao.Value = 1 Then
        .dataTipo_CQNC = 3
    End If
    If Chk_periodo.Value = 1 Or chkEmissao.Value = 1 Or chkDataConclusao.Value = 1 Then
        .dataTipo_CQNC_DE = msk_fltInicio.Value
        .dataTipo_CQNC_Ate = msk_fltFim.Value
    Else
        .dataTipo_CQNC_DE = ""
        .dataTipo_CQNC_Ate = ""
    End If
    
    FiltroEscopo = ""
    FiltroEscopoRel = ""
    If chkEscopo.Value = 0 And chkSemEscopo.Value = 1 Then
        FiltroEscopo = " and (P.Escopo = 'False' or P.Escopo IS NULL)"
        FiltroEscopoRel = " and ({Producao.Escopo} = False or ISNULL({Producao.Escopo}))"
    ElseIf chkEscopo.Value = 1 And chkSemEscopo.Value = 0 Then
        FiltroEscopo = " and P.Escopo = 'True'"
        FiltroEscopoRel = " and {Producao.Escopo} = True"
    End If
    
    FiltroConcluida = ""
    FiltroConcluidaRel = ""
    If chkNaoConcluida.Value = 0 And chkConcluida.Value = 1 Then
        FiltroConcluida = " and P.Concluida = 'True'"
        FiltroConcluidaRel = " and {Producao.Concluida} = True"
    ElseIf chkNaoConcluida.Value = 1 And chkConcluida.Value = 0 Then
        FiltroConcluida = " and P.Concluida = 'False'"
        FiltroConcluidaRel = " and {Producao.Concluida} = False"
    End If
End With

If Lista.ListItems.Count = 0 Then
    INNERJOIN = ""
    If txtTexto <> "" Or cmbTexto <> "" Then
        If cmbfiltrarpor = "Disposição" Then
            Filtro = "CQ.PARECERCQ = '" & cmbTexto & "' and "
            FiltroRel = "{CQ_NC_FABRICA.PARECERCQ} = '" & cmbTexto & "' and "
        ElseIf cmbfiltrarpor = "ID" Or cmbfiltrarpor = "Ordem" Or cmbfiltrarpor = "OS" Or cmbfiltrarpor = "Numero de série" Then
            Select Case cmbfiltrarpor
                Case "ID": TextoFiltro = "Codigo"
                Case "Ordem": TextoFiltro = "Ordem"
                Case "OS": TextoFiltro = "OS"
                Case "Numero de série": TextoFiltro = "NumeroSerie"
            End Select
            Filtro = "CQ." & TextoFiltro & " = " & txtTexto & " and "
            FiltroRel = "{CQ_NC_FABRICA." & TextoFiltro & "} = " & txtTexto & " and "
        ElseIf cmbfiltrarpor = "Grupo" Or cmbfiltrarpor = "Família" Then
            Select Case cmbfiltrarpor
                Case "Grupo": TextoFiltro = "Grupo"
                Case "Família": TextoFiltro = "Familia"
            End Select
            Filtro = "F." & TextoFiltro & " = '" & cmbTexto & "' and "
            FiltroRel = "{Projfamilia." & TextoFiltro & "}" & " = '" & cmbTexto & "' and "
            INNERJOIN = "INNER JOIN projproduto PP ON PP.desenho = P.desenho INNER JOIN projfamilia F ON PP.Classe = F.Familia"
        ElseIf cmbfiltrarpor = "Pedido interno" Then
            Filtro = "VP.Ncotacao" & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and "
            FiltroRel = "{Vendas_proposta.Ncotacao}" & FunVerifTipoFiltroIMFRel(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " And "
            INNERJOIN = "LEFT JOIN Producao_pedidos PPE ON PPE.Ordem = P.Ordem LEFT JOIN Vendas_carteira VC ON VC.codigo = PPE.IDCarteira LEFT JOIN Vendas_proposta VP ON VP.Cotacao = VC.Cotacao"
        Else
            Select Case cmbfiltrarpor
                Case "Código interno": TextoFiltro = "Desenho"
                Case "Descrição": TextoFiltro = "Produto"
                Case "Código de referência": TextoFiltro = "N_Referencia"
            End Select
            Filtro = "P." & TextoFiltro & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and "
            FiltroRel = "{Producao." & TextoFiltro & "}" & FunVerifTipoFiltroIMFRel(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " And "
        End If
    End If
    
    'Usado para o filtro do relatorio personalizado da Esplendor
    If (cmbfiltrarpor = "ID" Or cmbfiltrarpor = "Ordem" Or cmbfiltrarpor = "OS" Or cmbfiltrarpor = "Pedido interno") And txtTexto <> "" Then
        CamposFiltro = "P.desenho, P.Ordem "
        frmcqnc.pesquisaPorOrdem = True
    Else
        CamposFiltro = "P.desenho "
        frmcqnc.pesquisaPorOrdem = False
    End If
Else
    With Lista
        CamposFiltro = "P.desenho "
        frmcqnc.pesquisaPorOrdem = False
        For InitFor = 1 To .ListItems.Count
            If .ListItems(InitFor).ListSubItems(1) = "Disposição" Then
                If Filtro = "" Then
                    Filtro = "CQ.PARECERCQ = '" & .ListItems(InitFor).ListSubItems(3) & "' and "
                    FiltroRel = "{CQ_NC_FABRICA.PARECERCQ} = '" & .ListItems(InitFor).ListSubItems(3) & "' and "
                Else
                    Filtro = Filtro & "CQ.PARECERCQ = '" & .ListItems(InitFor).ListSubItems(3) & "' and "
                    FiltroRel = FiltroRel & "{CQ_NC_FABRICA.PARECERCQ} = '" & .ListItems(InitFor).ListSubItems(3) & "' and "
                End If
            ElseIf .ListItems(InitFor).ListSubItems(1) = "ID" Or .ListItems(InitFor).ListSubItems(1) = "Ordem" Or .ListItems(InitFor).ListSubItems(1) = "OS" Then
                Select Case .ListItems(InitFor).ListSubItems(1)
                    Case "ID": TextoFiltro = "Codigo"
                    Case "Ordem": TextoFiltro = "Ordem"
                    Case "OS": TextoFiltro = "OS"
                    Case "Disposição": TextoFiltro = "PARECERCQ"
                End Select
                If Filtro = "" Then
                    Filtro = "CQ." & TextoFiltro & " = " & .ListItems(InitFor).ListSubItems(3) & " and "
                    FiltroRel = "{CQ_NC_FABRICA." & TextoFiltro & "} = " & .ListItems(InitFor).ListSubItems(3) & " and "
                Else
                    Filtro = Filtro & "CQ." & TextoFiltro & " = " & .ListItems(InitFor).ListSubItems(3) & " and "
                    FiltroRel = FiltroRel & "{CQ_NC_FABRICA." & TextoFiltro & "} = " & .ListItems(InitFor).ListSubItems(3) & " and "
                End If
                                
                CamposFiltro = "P.desenho, P.Ordem "
                frmcqnc.pesquisaPorOrdem = True
                
            ElseIf .ListItems(InitFor).ListSubItems(1) = "Grupo" Or .ListItems(InitFor).ListSubItems(1) = "Família" Then
                Select Case .ListItems(InitFor).ListSubItems(1)
                    Case "Grupo": TextoFiltro = "Grupo"
                    Case "Família": TextoFiltro = "Familia"
                End Select
                If Filtro = "" Then
                    Filtro = "F." & TextoFiltro & " = '" & .ListItems(InitFor).ListSubItems(3) & "' and "
                    FiltroRel = "{Projfamilia." & TextoFiltro & "}" & " = '" & .ListItems(InitFor).ListSubItems(3) & "' and "
                Else
                    Filtro = Filtro & "F." & TextoFiltro & " = '" & .ListItems(InitFor).ListSubItems(3) & "' and "
                    FiltroRel = FiltroRel & "{Projfamilia." & TextoFiltro & "}" & " = '" & .ListItems(InitFor).ListSubItems(3) & "' and "
                End If
            ElseIf .ListItems(InitFor).ListSubItems(1) = "Pedido interno" Then
                If Filtro = "" Then
                    Filtro = "VP.Ncotacao" & FunVerifTipoFiltroIMFLista(.ListItems(InitFor).ListSubItems(2), .ListItems(InitFor).ListSubItems(3)) & " and "
                    FiltroRel = "{Vendas_proposta.Ncotacao}" & FunVerifTipoFiltroIMFListaRel(.ListItems(InitFor).ListSubItems(2), .ListItems(InitFor).ListSubItems(3)) & " And "
                Else
                    Filtro = Filtro & "VP.Ncotacao" & FunVerifTipoFiltroIMFLista(.ListItems(InitFor).ListSubItems(2), .ListItems(InitFor).ListSubItems(3)) & " and "
                    FiltroRel = FiltroRel & "{Vendas_proposta.Ncotacao}" & FunVerifTipoFiltroIMFListaRel(.ListItems(InitFor).ListSubItems(2), .ListItems(InitFor).ListSubItems(3)) & " And "
                End If
                
                CamposFiltro = "P.desenho, P.Ordem "
                frmcqnc.pesquisaPorOrdem = True
            Else
                Select Case .ListItems(InitFor).ListSubItems(1)
                    Case "Código interno": TextoFiltro = "Desenho"
                    Case "Descrição": TextoFiltro = "Produto"
                    Case "Código de referência": TextoFiltro = "N_Referencia"
                End Select
                If Filtro = "" Then
                    Filtro = "P." & TextoFiltro & FunVerifTipoFiltroIMFLista(.ListItems(InitFor).ListSubItems(2), .ListItems(InitFor).ListSubItems(3)) & " and "
                    FiltroRel = "{Producao." & TextoFiltro & "}" & FunVerifTipoFiltroIMFListaRel(.ListItems(InitFor).ListSubItems(2), .ListItems(InitFor).ListSubItems(3)) & " And "
                Else
                    Filtro = Filtro & "P." & TextoFiltro & FunVerifTipoFiltroIMFLista(.ListItems(InitFor).ListSubItems(2), .ListItems(InitFor).ListSubItems(3)) & " and "
                    FiltroRel = FiltroRel & "{Producao." & TextoFiltro & "}" & FunVerifTipoFiltroIMFListaRel(.ListItems(InitFor).ListSubItems(2), .ListItems(InitFor).ListSubItems(3)) & " And "
                End If
            End If
        Next InitFor
    End With
End If
With frmcqnc
    selectfiltro = "from producao P INNER JOIN CQ_NC_FABRICA CQ on P.Ordem = CQ.Ordem LEFT JOIN Producao_pedidos PPE ON PPE.Ordem = P.Ordem LEFT JOIN Vendas_carteira VC ON VC.codigo = PPE.IDCarteira LEFT JOIN Vendas_proposta VP ON VP.Cotacao = VC.Cotacao LEFT JOIN projproduto PP ON PP.desenho = P.desenho LEFT JOIN projfamilia F ON PP.Classe = F.Familia where " & Filtro & TextoFiltroAnalisada & TextoFiltroRetrabalho & DataFiltro & FiltroEscopo & FiltroConcluida
    Campos = "P.Quant, CQ.QTCD, CQ.codigo, CQ.OS, CQ.operador, CQ.Lote, CQ.TTNC, CQ.Data, CQ.ParecerCQ, CQ.Ordem,CQ.OrdemRetrabalho, CQ.IDProducao, CQ.Analizada "
    
    .StrSql_CQ_NC = "Select " & Campos & selectfiltro & " GROUP BY " & Campos & " order by CQ.Codigo desc"
    .StrSql_CQ_NC_FIltro = "Select " & CamposFiltro & selectfiltro & " GROUP BY " & CamposFiltro 'Usado para o relatorio da Esplendor
    .FormulaRel_CQ_NC = FiltroRel & TextoFiltroAnalisadaRel & TextoFiltroRetrabalhoRel & DataFiltroRel & FiltroEscopoRel & FiltroConcluidaRel
    .ProcCarregaLista (1)
    'Debug.print .StrSql_CQ_NC
End With
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Cmd_excluir_Click()
On Error GoTo tratar_erro

Permitido = False
Inicio:
    With Lista
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                If Permitido = False Then
                    If USMsgBox("Deseja realmente excluir este(s) filtro(s) para pesquisa?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
                End If
                Permitido = True
                .ListItems.Remove (InitFor)
                GoTo Inicio
            End If
        Next InitFor
    End With
    If Permitido = False Then
        USMsgBox ("Informe o(s) filtro(s) para pesquisa antes de excluir."), vbExclamation, "CAPRIND v5.0"
    Else
        USMsgBox ("Filtro(s) para pesquisa excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
        If Lista.ListItems.Count = 0 Then
            Lista.Visible = False
            Me.Height = 4395
        End If
    End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Cmd_salvar_Click()
On Error GoTo tratar_erro

If txtTexto.Visible = True And txtTexto = "" Or cmbTexto.Visible = True And cmbTexto = "" Then
    USMsgBox ("Informe o texto para pesquisa antes de adicionar o filtro na lista."), vbExclamation, "CAPRIND v5.0"
    If txtTexto.Visible = True Then txtTexto.SetFocus Else cmbTexto.SetFocus
    Exit Sub
End If

With Lista.ListItems
    .Add , , ""
    .Item(.Count).SubItems(1) = cmbfiltrarpor.Text
    If Optinicio.Value = True Then .Item(.Count).SubItems(2) = "Início"
    If Optmeio.Value = True Then .Item(.Count).SubItems(2) = "Meio"
    If Optfim.Value = True Then .Item(.Count).SubItems(2) = "Fim"
    If optIgual.Value = True Then .Item(.Count).SubItems(2) = "Igual"
    If txtTexto.Visible = True Then
        .Item(.Count).SubItems(3) = txtTexto
    Else
        .Item(.Count).SubItems(3) = cmbTexto.Text
        .Item(.Count).SubItems(4) = cmbTexto.ItemData(cmbTexto.ListIndex)
    End If
End With
Lista.Visible = True
Me.Height = 6525

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyEscape: Unload Me
    Case vbKeyF2: ProcFiltrar
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 10665, 5, True

cmbfiltrarpor = "OS"
msk_fltFim.Value = Date
msk_fltInicio.Value = Date

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub txtTexto_Change()
On Error GoTo tratar_erro

If txtTexto <> "" Then
    If cmbfiltrarpor = "Ordem" Or cmbfiltrarpor = "OS" Or cmbfiltrarpor = "ID" Then
        VerifNumero = txtTexto
        ProcVerificaNumero
        If VerifNumero = False Then
            txtTexto = ""
            txtTexto.SetFocus
            Exit Sub
        End If
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcFiltrar
    'Case 3: ProcAjuda
    Case 4: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

