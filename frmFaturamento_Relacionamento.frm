VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFaturamento_Relacionamento 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Administrativo - Faturamento - Nota fiscal - Relacionamento de nota fiscal"
   ClientHeight    =   8595
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   11820
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
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
      Left            =   180
      TabIndex        =   26
      Top             =   7230
      Width           =   11475
      Begin VB.TextBox txtQtde1 
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
         Left            =   9840
         TabIndex        =   20
         Top             =   375
         Width           =   1455
      End
      Begin VB.TextBox txtQtde 
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
         Left            =   8700
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   375
         Width           =   1125
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
         Height          =   315
         Left            =   2250
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "Descri��o."
         Top             =   375
         Width           =   6435
      End
      Begin VB.TextBox txtCodinterno 
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
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "C�digo interno."
         Top             =   375
         Width           =   2055
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Qtde. � relacionar"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   9900
         TabIndex        =   30
         Top             =   180
         Width           =   1335
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Descri��o"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2250
         TabIndex        =   29
         Top             =   180
         Width           =   6435
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Qtde. entr."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   8850
         TabIndex        =   28
         Top             =   180
         Width           =   825
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "C�d. interno"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   750
         TabIndex        =   27
         Top             =   180
         Width           =   900
      End
   End
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   53
      Top             =   8190
      Width           =   11820
      _ExtentX        =   20849
      _ExtentY        =   714
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   52
      Top             =   0
      Width           =   11820
      _ExtentX        =   20849
      _ExtentY        =   688
      DibPicture      =   "frmFaturamento_Relacionamento.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Icon            =   "frmFaturamento_Relacionamento.frx":1B63
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7755
      Left            =   90
      TabIndex        =   25
      Top             =   390
      Width           =   11625
      _ExtentX        =   20505
      _ExtentY        =   13679
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   14215660
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Notas fiscais � relacionar"
      TabPicture(0)   =   "frmFaturamento_Relacionamento.frx":1E7D
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "USToolBar1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame9"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "ListView1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "PBLista"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame6"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Notas fiscais relacionadas"
      TabPicture(1)   =   "frmFaturamento_Relacionamento.frx":1E99
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txt_ID"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "USToolBar2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "PBLista1"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "ListView2"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Frame3"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      Begin VB.Frame Frame6 
         BackColor       =   &H00E0E0E0&
         Height          =   510
         Left            =   60
         TabIndex        =   50
         Top             =   1290
         Width           =   2175
         Begin VB.OptionButton Opt_saida 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Saida"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   180
            TabIndex        =   10
            Top             =   210
            Value           =   -1  'True
            Width           =   705
         End
         Begin VB.OptionButton Opt_entrada 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Entrada"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   1080
            TabIndex        =   11
            Top             =   210
            Width           =   885
         End
      End
      Begin DrawSuite2022.USProgressBar PBLista 
         Height          =   255
         Left            =   60
         TabIndex        =   39
         Top             =   6555
         Width           =   11475
         _ExtentX        =   20241
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
      Begin MSComctlLib.ListView ListView1 
         Height          =   3210
         Left            =   30
         TabIndex        =   2
         Top             =   2730
         Width           =   11475
         _ExtentX        =   20241
         _ExtentY        =   5662
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483641
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
         NumItems        =   14
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "N"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Tag             =   "N"
            Text            =   "IDprod"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "C�d. interno"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Tag             =   "T"
            Text            =   "C�d. ref."
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "Descri��o"
            Object.Width           =   2259
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Object.Tag             =   "D"
            Text            =   "Dt. emiss�o"
            Object.Width           =   1676
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Object.Tag             =   "T"
            Text            =   "Nota fiscal"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   7
            Object.Tag             =   "T"
            Text            =   "Tipo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Object.Tag             =   "T"
            Text            =   "Destinat�rio"
            Object.Width           =   2259
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   9
            Object.Tag             =   "N"
            Text            =   "Qtde."
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   10
            Object.Tag             =   "N"
            Text            =   "Qtde."
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   11
            Object.Tag             =   "N"
            Text            =   "Saldo"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   12
            Object.Tag             =   "N"
            Text            =   "Vlr. unit."
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Object.Tag             =   "T"
            Text            =   "Un. com"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   60
         TabIndex        =   46
         Top             =   5940
         Width           =   11475
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
            Left            =   5760
            TabIndex        =   4
            ToolTipText     =   "N�mero da p�gina."
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
            Left            =   3090
            TabIndex        =   3
            Text            =   "30"
            ToolTipText     =   "N�mero de registros por p�gina."
            Top             =   180
            Width           =   555
         End
         Begin DrawSuite2022.USButton cmdPagProx 
            Height          =   315
            Left            =   7980
            TabIndex        =   8
            ToolTipText     =   "Pr�xima p�gina."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmFaturamento_Relacionamento.frx":1EB5
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
            Left            =   7440
            TabIndex        =   7
            ToolTipText     =   "P�gina anterior."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmFaturamento_Relacionamento.frx":565C
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
            Left            =   6330
            TabIndex        =   5
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
            Left            =   6900
            TabIndex        =   6
            ToolTipText     =   "Primeira p�gina."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmFaturamento_Relacionamento.frx":916B
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
            Left            =   8520
            TabIndex        =   9
            ToolTipText     =   "�ltima p�gina."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmFaturamento_Relacionamento.frx":D25F
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
         Begin VB.Label lblPaginas 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "P�gina: 0 de: 0"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   9840
            TabIndex        =   49
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lblRegistros 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "N� de registros: 0"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   180
            TabIndex        =   48
            Top             =   240
            Width           =   1275
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Carregar               registros por p�gina"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   2400
            TabIndex        =   47
            Top             =   240
            Width           =   2760
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0E0&
         Height          =   915
         Left            =   60
         TabIndex        =   43
         Top             =   1800
         Width           =   11475
         Begin VB.Frame Frame11 
            BackColor       =   &H00E0E0E0&
            Height          =   510
            Left            =   2070
            TabIndex        =   51
            Top             =   300
            WhatsThisHelpID =   210
            Width           =   4785
            Begin VB.OptionButton optIgual 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Igual"
               Height          =   255
               Left            =   3930
               TabIndex        =   16
               Top             =   180
               Width           =   705
            End
            Begin VB.OptionButton Optmeio 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Meio frase"
               Height          =   255
               Left            =   1470
               TabIndex        =   14
               Top             =   180
               Width           =   1275
            End
            Begin VB.OptionButton Optinicio 
               BackColor       =   &H00E0E0E0&
               Caption         =   "In�cio frase"
               Height          =   255
               Left            =   180
               TabIndex        =   13
               Top             =   180
               Value           =   -1  'True
               Width           =   1275
            End
            Begin VB.OptionButton Optfim 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Fim frase"
               Height          =   255
               Left            =   2760
               TabIndex        =   15
               Top             =   180
               Width           =   1155
            End
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
            Left            =   6930
            TabIndex        =   1
            ToolTipText     =   "Texto para pesquisa."
            Top             =   450
            Width           =   4365
         End
         Begin VB.ComboBox cmbfiltrarpor 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   315
            ItemData        =   "frmFaturamento_Relacionamento.frx":10AF9
            Left            =   180
            List            =   "frmFaturamento_Relacionamento.frx":10B15
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   0
            ToolTipText     =   "Op��es para filtro."
            Top             =   450
            Width           =   1815
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Filtrar por"
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   660
            TabIndex        =   45
            Top             =   240
            Width           =   705
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Texto para pesquisa"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   8377
            TabIndex        =   44
            Top             =   240
            Width           =   1470
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Height          =   510
         Left            =   2250
         TabIndex        =   42
         Top             =   1290
         Width           =   9285
         Begin VB.CheckBox chkEmitente 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Filtrar notas de qualquer emitente"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   270
            TabIndex        =   57
            ToolTipText     =   "Se marcado filtra notas de qualquer emitente com o mesmo c�digo do item."
            Top             =   210
            Width           =   2835
         End
         Begin VB.CheckBox chkSaldoEstoque 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Filtrar notas com saldo em estoque"
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   3600
            TabIndex        =   54
            ToolTipText     =   "Se marcado filtra somente itens com saldo em estoque"
            Top             =   210
            Width           =   3675
         End
         Begin VB.OptionButton optProduto 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Produtos"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   330
            TabIndex        =   12
            Top             =   540
            Value           =   -1  'True
            Width           =   945
         End
      End
      Begin VB.TextBox txt_ID 
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
         Left            =   -74220
         Locked          =   -1  'True
         TabIndex        =   37
         TabStop         =   0   'False
         ToolTipText     =   "C�digo interno."
         Top             =   3060
         Visible         =   0   'False
         Width           =   645
      End
      Begin DrawSuite2022.USToolBar USToolBar1 
         Height          =   975
         Left            =   60
         TabIndex        =   38
         Top             =   330
         Width           =   11475
         _ExtentX        =   20241
         _ExtentY        =   1720
         ButtonCount     =   6
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
         ButtonCaption1  =   "Filtrar"
         ButtonEnabled1  =   0   'False
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
         ButtonCaption2  =   "Relacionar"
         ButtonEnabled2  =   0   'False
         ButtonIconSize2 =   32
         ButtonToolTipText2=   "Relacionar (F3)"
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
         ButtonWidth2    =   58
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
         ButtonLeft3     =   100
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
         ButtonLeft4     =   104
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
         ButtonLeft5     =   142
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
         ButtonLeft6     =   170
         ButtonTop6      =   2
         ButtonWidth6    =   24
         ButtonHeight6   =   24
         ButtonUseMaskColor6=   0   'False
         Begin VB.TextBox txtIDProduto 
            Height          =   285
            Left            =   8730
            TabIndex        =   56
            Text            =   "Text1"
            Top             =   1140
            Visible         =   0   'False
            Width           =   585
         End
         Begin VB.TextBox txtIDNota 
            Height          =   285
            Left            =   7710
            TabIndex        =   55
            Text            =   "Text1"
            Top             =   1110
            Visible         =   0   'False
            Width           =   585
         End
         Begin DrawSuite2022.USImageList USImageList1 
            Left            =   3720
            Top             =   330
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmFaturamento_Relacionamento.frx":10BAD
            Count           =   1
         End
      End
      Begin DrawSuite2022.USToolBar USToolBar2 
         Height          =   975
         Left            =   -74970
         TabIndex        =   40
         Top             =   330
         Width           =   11475
         _ExtentX        =   20241
         _ExtentY        =   1720
         ButtonCount     =   6
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft1     =   2
         ButtonTop1      =   2
         ButtonWidth1    =   38
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft2     =   42
         ButtonTop2      =   2
         ButtonWidth2    =   39
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
         ButtonLeft3     =   83
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
         ButtonLeft4     =   87
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
         ButtonLeft5     =   125
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
         ButtonLeft6     =   153
         ButtonTop6      =   2
         ButtonWidth6    =   24
         ButtonHeight6   =   24
         ButtonUseMaskColor6=   0   'False
         Begin DrawSuite2022.USImageList USImageList2 
            Left            =   4170
            Top             =   90
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmFaturamento_Relacionamento.frx":139CF
            Count           =   1
         End
      End
      Begin DrawSuite2022.USProgressBar PBLista1 
         Height          =   255
         Left            =   -74940
         TabIndex        =   41
         Top             =   6585
         Width           =   11475
         _ExtentX        =   20241
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
      Begin MSComctlLib.ListView ListView2 
         Height          =   4410
         Left            =   -74940
         TabIndex        =   21
         Top             =   2160
         Width           =   11475
         _ExtentX        =   20241
         _ExtentY        =   7779
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483641
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
            Text            =   "Dt. emiss�o"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "N"
            Text            =   "Nota fiscal"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Object.Tag             =   "T"
            Text            =   "Tipo"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "Destinat�rio/Emitente"
            Object.Width           =   8036
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Object.Tag             =   "N"
            Text            =   "Qtde. relac."
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Object.Tag             =   "N"
            Text            =   "Vlr. unit."
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   7
            Object.Tag             =   "T"
            Text            =   "Un. com."
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Frame Frame3 
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
         Left            =   -74945
         TabIndex        =   31
         Top             =   6840
         Width           =   11475
         Begin VB.TextBox txtQtde3 
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
            Left            =   2730
            Locked          =   -1  'True
            TabIndex        =   22
            TabStop         =   0   'False
            Text            =   "0,000"
            Top             =   390
            Width           =   1635
         End
         Begin VB.TextBox txtQtdeRel 
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
            Left            =   4815
            Locked          =   -1  'True
            TabIndex        =   23
            TabStop         =   0   'False
            Text            =   "0,000"
            ToolTipText     =   "Quantidade relacionada."
            Top             =   390
            Width           =   1635
         End
         Begin VB.TextBox txtSaldo 
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
            Left            =   6900
            Locked          =   -1  'True
            TabIndex        =   24
            TabStop         =   0   'False
            Text            =   "0,000"
            ToolTipText     =   "Saldo"
            Top             =   390
            Width           =   1635
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Qtde. sa�da"
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
            Left            =   3075
            TabIndex        =   36
            Top             =   180
            Width           =   945
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-"
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
            Left            =   4545
            TabIndex        =   35
            Top             =   450
            Width           =   75
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Qtde. relacionada"
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
            Left            =   4890
            TabIndex        =   34
            Top             =   180
            Width           =   1485
         End
         Begin VB.Label Label8 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "="
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
            Left            =   6630
            TabIndex        =   33
            Top             =   450
            Width           =   135
         End
         Begin VB.Label Label9 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Saldo"
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
            Left            =   7545
            TabIndex        =   32
            Top             =   180
            Width           =   465
         End
      End
   End
End
Attribute VB_Name = "frmFaturamento_Relacionamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RelacionamentoSimultaneo As Boolean 'OK

Private Sub ProcExcluir()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Aten��o usu�rio " & pubUsuario & " voce n�o tem acesso a este recurso.")
    Exit Sub
End If
Permitido = False
With ListView2
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir esta(s) nota(s) relacionada(s)?", vbYesNo) = vbNo Then Exit Sub
            End If
            Permitido = True
            Id_Item = .ListItems(InitFor)
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select * from Faturamento_Relacionamento WHERE ID = " & Id_Item, Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False Then
                '==================================
                Modulo = Formulario
                Evento = "Excluir"
                ID_documento = .ListItems(InitFor)
                
                Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open "select int_NotaFiscal, ID, TipoNF, Serie from tbl_dados_nota_fiscal where ID = " & frmFaturamento_Prod_Serv.txtId, Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    If IsNull(TBAbrir!int_NotaFiscal) = True Or TBAbrir!int_NotaFiscal = "" Then NomeCampo = "N� ordem: " & TBAbrir!ID Else NomeCampo = "N� nota: " & TBAbrir!int_NotaFiscal
                    Documento = NomeCampo & " - Tipo: " & TBAbrir!TipoNF & " - S�rie: " & TBAbrir!Serie
                End If
                
                Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open "select int_NotaFiscal, ID, TipoNF, Serie from tbl_dados_nota_fiscal where ID = " & TBFI!ID_nota_relacionada, Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    Documento1 = "N� nota relacionada: " & TBAbrir!int_NotaFiscal & " - Tipo: " & TBAbrir!TipoNF & " - S�rie: " & TBAbrir!Serie
                End If
                TBAbrir.Close
                
                ProcGravaEvento
                '==================================
                
                If Left(frmFaturamento_Prod_Serv.cmbFinalidade_emissao, 1) = 2 Then
                    Set TBProduto = CreateObject("adodb.recordset")
                    TBProduto.Open "Select Complemento_descricao from tbl_Detalhes_Nota where Int_codigo = " & frmFaturamento_Prod_Serv.txtidproduto, Conexao, adOpenKeyset, adLockOptimistic
                    If TBProduto.EOF = False Then
                        complemento = ""
                        Set TBAbrir = CreateObject("adodb.recordset")
                        TBAbrir.Open "Select NF.dt_DataEmissao, NF.int_NotaFiscal from tbl_Dados_Nota_Fiscal NF INNER JOIN Faturamento_Relacionamento FR ON FR.Id_nota_relacionada = NF.ID where FR.Id_nota = " & TBFI!ID_nota & " and FR.ID <> " & Id_Item, Conexao, adOpenKeyset, adLockOptimistic
                        If TBAbrir.EOF = False Then
                            Do While TBAbrir.EOF = False
                                If IsNull(TBAbrir!int_NotaFiscal) = False Then OF = TBAbrir!int_NotaFiscal
                                If complemento = "" Then complemento = "S/NF " & OF & " - " & Format(TBAbrir!dt_DataEmissao, "dd/mm/yy") Else complemento = complemento & " ; S/NF " & OF & " - " & Format(TBAbrir!dt_DataEmissao, "dd/mm/yy")
                                TBAbrir.MoveNext
                            Loop
                        End If
                        TBAbrir.Close
                        TBProduto!Complemento_descricao = IIf(complemento = "", Null, complemento)
                        TBProduto.Update
                    End If
                Else
                    'Atualiza o saldo no produto da NF relacionada
                    Set TBProduto = CreateObject("adodb.recordset")
                    TBProduto.Open "Select Saldo, Complemento_descricao from tbl_Detalhes_Nota where Int_codigo = " & TBFI!id_produto_relacionada, Conexao, adOpenKeyset, adLockOptimistic
                    If TBProduto.EOF = False Then
                        TBProduto!Saldo = Format(TBProduto!Saldo + TBFI!Qtde, "###,##0.0000")
                        
                        complemento = ""
                        Set TBAbrir = CreateObject("adodb.recordset")
                        TBAbrir.Open "Select NF.dt_DataEmissao, NF.int_NotaFiscal from tbl_Dados_Nota_Fiscal NF INNER JOIN Faturamento_Relacionamento FR ON FR.Id_nota = NF.ID where FR.Id_produto_relacionada = " & TBFI!id_produto_relacionada & " and FR.ID <> " & Id_Item, Conexao, adOpenKeyset, adLockOptimistic
                        If TBAbrir.EOF = False Then
                            Do While TBAbrir.EOF = False
                                If IsNull(TBAbrir!int_NotaFiscal) = False Then OF = TBAbrir!int_NotaFiscal
                                If complemento = "" Then complemento = "S/NF " & OF & " - " & Format(TBAbrir!dt_DataEmissao, "dd/mm/yy") Else complemento = complemento & " ; S/NF " & OF & " - " & Format(TBAbrir!dt_DataEmissao, "dd/mm/yy")
                                TBAbrir.MoveNext
                            Loop
                        End If
                        TBAbrir.Close
                        TBProduto!Complemento_descricao = complemento 'IIf(Complemento = "", Null, Complemento)
                        TBProduto.Update
                    End If
                    
                    'Atualiza o complemento da descri��o e o saldo no produto da NF
                    Set TBProduto = CreateObject("adodb.recordset")
                    TBProduto.Open "Select Saldo,  Complemento_descricao from tbl_Detalhes_Nota where Int_codigo = " & TBFI!ID_Produto, Conexao, adOpenKeyset, adLockOptimistic
                    If TBProduto.EOF = False Then
                        TBProduto!Saldo = Format(TBProduto!Saldo + TBFI!Qtde, "###,##0.0000")
                        
                        complemento = ""
                        Set TBAbrir = CreateObject("adodb.recordset")
                        TBAbrir.Open "Select NF.dt_DataEmissao, NF.int_NotaFiscal from tbl_Dados_Nota_Fiscal NF INNER JOIN Faturamento_Relacionamento FR ON FR.Id_nota_relacionada = NF.ID where FR.Id_produto = " & TBFI!ID_Produto & " and FR.ID <> " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
                        If TBAbrir.EOF = False Then
                            Do While TBAbrir.EOF = False
                                If IsNull(TBAbrir!int_NotaFiscal) = False Then OF = TBAbrir!int_NotaFiscal
                                If complemento = "" Then complemento = "S/NF " & OF & " - " & Format(TBAbrir!dt_DataEmissao, "dd/mm/yy") Else complemento = complemento & " ; S/NF " & OF & " - " & Format(TBAbrir!dt_DataEmissao, "dd/mm/yy")
                                TBAbrir.MoveNext
                            Loop
                        End If
                        TBAbrir.Close
                        TBProduto!Complemento_descricao = IIf(complemento = "", Null, complemento)
                        TBProduto.Update
                    End If
                    TBProduto.Close
                End If
                ProcCarregaCampoComplemento

                Conexao.Execute "DELETE from Faturamento_Relacionamento where ID = " & .ListItems(InitFor)
                
            End If
            TBFI.Close
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe a(s) nota(s) relacionada(s) antes de excluir."), vbExclamation
Else
    USMsgBox ("Nota(s) relacionada(s) exclu�da(s) com sucesso."), vbInformation
    With frmFaturamento_Prod_Serv
        If Left(.cmbFinalidade_emissao, 1) <> 1 Then
            Set TBProduto = CreateObject("adodb.recordset")
            TBProduto.Open "Select * from tbl_Detalhes_Nota where Int_codigo = " & .txtidproduto, Conexao, adOpenKeyset, adLockOptimistic
            If TBProduto.EOF = False Then
                .ProcAtualizaCST .txtIDEmpresa.Text, .txtIDcliente, .txt_Razao, .cbo_UF, Left(.Cmb_consumidor, 1), Left(.cmbFinalidade_emissao, 1)
            End If
            TBProduto.Close
        End If
    End With
    Txt_ID = ""
    txtQtde1 = ""
    ProcCarregaLista
    ProcCarregaListaRelacionada
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descri��o do erro : " + Error()), vbCritical, "CAPRIND V5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descri��o do erro : " + Error()), vbCritical, "CAPRIND V5.0"
    Exit Sub
End Sub

Private Sub ProcSalvarFaturamento(NovoRelacionamento As Boolean)
On Error GoTo tratar_erro

If SSTab1.Tab = 0 Then
    If ListView1.ListItems.Count = 0 Then Exit Sub
Else
    If ListView2.ListItems.Count = 0 Then Exit Sub
End If
PagarParcial = False

With frmFaturamento_Prod_Serv
    'If FunVerificaRegistroValidado("tbl_Dados_Nota_Fiscal", "ID = " & .txtID, "nota fiscal", IIf(SSTab1.Tab = 0, "", "este relacionamento"), IIf(SSTab1.Tab = 0, "relacionar", "alterar"), False, True) = False Then Exit Sub
    
    .Produto_Relacionado = True
    
    If SSTab1.Tab = 0 And .txtidproduto = 0 Then
        ProcAdicionarNovo
    Else
        If SSTab1.Tab = 0 Then
            If Left(.cmbFinalidade_emissao, 1) = 2 Then
                TextoFiltro = "ID_nota = " & .txtId & " and ID_nota_relacionada = " & ListView1.SelectedItem & " Or ID_nota = " & ListView1.SelectedItem & " And ID_nota_relacionada = " & .txtId
            Else
                TextoFiltro = "ID_nota = " & .txtId & " and ID_produto = " & .txtidproduto & " and ID_nota_relacionada = " & ListView1.SelectedItem & " and ID_produto_relacionada = " & ListView1.SelectedItem.ListSubItems(1) & " or ID_nota = " & ListView1.SelectedItem & " and ID_produto = " & ListView1.SelectedItem.ListSubItems(1) & " and ID_nota_relacionada = " & .txtId & " and ID_produto_relacionada = " & .txtidproduto
            End If
        Else
            TextoFiltro = "ID = " & ListView2.SelectedItem
        End If
        If Left(.cmbFinalidade_emissao, 1) <> 2 Then
'            Set TBGravar = CreateObject("adodb.recordset")
'            TBGravar.Open "Select * from Faturamento_Relacionamento where " & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
'            If TBGravar.EOF = False Then
'                If FunVerificaCamposSalvar(False) = False Then Exit Sub
'            Else
'                If FunVerificaCamposSalvar(True) = False Then Exit Sub
'            End If
            
            If FunVerificaCamposSalvar(NovoRelacionamento) = False Then Exit Sub
            If SSTab1.Tab = 0 Then
               ' .txtVLUnit = ListView1.SelectedItem.ListSubItems(12)
                '.Cmb_un_com = ListView1.SelectedItem.ListSubItems(13)
                TextoFiltro = ""
            Else
                .txtVLUnit = ListView2.SelectedItem.ListSubItems(6)
                .Cmb_un_com = ListView2.SelectedItem.ListSubItems(7)
                TextoFiltro = " and ID <> " & ListView2.SelectedItem
            End If
            ProcEnviaDadosRelacionamento ListView1.SelectedItem, ListView1.SelectedItem.ListSubItems(1), ListView1.SelectedItem.ListSubItems(5), ListView1.SelectedItem.ListSubItems(6), False, True, quantidade
        Else
            ProcEnviaDadosRelacionamento ListView1.SelectedItem, IIf(ListView1.SelectedItem.ListSubItems(1) = "", 0, ListView1.SelectedItem.ListSubItems(1)), ListView1.SelectedItem.ListSubItems(5), ListView1.SelectedItem.ListSubItems(6), False, True, quantidade
        End If
        PagarParcial = True
    End If
    
    If RelacionamentoSimultaneo = True And PagarParcial = True Then
        Unload Me
    Else
        If PagarParcial = True Then
            ProcCarregaLista
            ProcCarregaListaRelacionada
            .Produto_Relacionado = False
            '.ProcCarregaLista
            
            If .NF_alterada = True Then .ProcCarregaTotaisNota IIf(.txtId = "", 0, .txtId)
            '.ProcGravarTotaisNota
            '.ProcCarregaTotaisNota IIf(.txtid = "", 0, .txtid)
            '.ProcCarregaListaNota (IIf(ReturnNumbersOnly(Left(.lblPaginas(1).Caption, Len(.lblPaginas(1).Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(.lblPaginas(1).Caption, Len(.lblPaginas(1).Caption) - 5))))
        End If
    End If
    
'    If Left(.cmbFinalidade_emissao, 1) <> 1 Then
'        Set TBProduto = CreateObject("adodb.recordset")
'        TBProduto.Open "Select * from tbl_Detalhes_Nota where Int_codigo = " & .txtidproduto, Conexao, adOpenKeyset, adLockOptimistic
'        If TBProduto.EOF = False Then
'            .ProcAtualizaCST .txtIDEmpresa.Text, .txtIDcliente, .txt_Razao, .cbo_UF, Left(.Cmb_consumidor, 1), Left(.cmbFinalidade_emissao, 1)
'        End If
'        TBProduto.Close
'    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descri��o do erro : " + Error()), vbCritical, "CAPRIND V5.0"
    Exit Sub
End Sub

Private Sub ProcSalvarOrdem(NovoRelacionamento As Boolean)
On Error GoTo tratar_erro

If SSTab1.Tab = 0 Then
    If ListView1.ListItems.Count = 0 Then Exit Sub
Else
    If ListView2.ListItems.Count = 0 Then Exit Sub
End If
PagarParcial = False

With frmEstoque_Ordem_Faturamento
    'If FunVerificaRegistroValidado("tbl_Dados_Nota_Fiscal", "ID = " & .txtID, "nota fiscal", IIf(SSTab1.Tab = 0, "", "este relacionamento"), IIf(SSTab1.Tab = 0, "relacionar", "alterar"), False, True) = False Then Exit Sub
    
    .Produto_Relacionado = True
    
    If SSTab1.Tab = 0 And .txtidproduto = 0 Then
        ProcAdicionarNovo
    Else
        If SSTab1.Tab = 0 Then
            If Left(.cmbFinalidade_emissao, 1) = 2 Then
                TextoFiltro = "ID_nota = " & .txtId & " and ID_nota_relacionada = " & ListView1.SelectedItem & " Or ID_nota = " & ListView1.SelectedItem & " And ID_nota_relacionada = " & .txtId
            Else
                TextoFiltro = "ID_nota = " & .txtId & " and ID_produto = " & .txtidproduto & " and ID_nota_relacionada = " & ListView1.SelectedItem & " and ID_produto_relacionada = " & ListView1.SelectedItem.ListSubItems(1) & " or ID_nota = " & ListView1.SelectedItem & " and ID_produto = " & ListView1.SelectedItem.ListSubItems(1) & " and ID_nota_relacionada = " & .txtId & " and ID_produto_relacionada = " & .txtidproduto
            End If
        Else
            TextoFiltro = "ID = " & ListView2.SelectedItem
        End If
        If Left(.cmbFinalidade_emissao, 1) <> 2 Then
'            Set TBGravar = CreateObject("adodb.recordset")
'            TBGravar.Open "Select * from Faturamento_Relacionamento where " & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
'            If TBGravar.EOF = False Then
'                If FunVerificaCamposSalvar(False) = False Then Exit Sub
'            Else
'                If FunVerificaCamposSalvar(True) = False Then Exit Sub
'            End If
            
            If FunVerificaCamposSalvar(NovoRelacionamento) = False Then Exit Sub
            If SSTab1.Tab = 0 Then
               ' .txtVLUnit = ListView1.SelectedItem.ListSubItems(12)
                '.Cmb_un_com = ListView1.SelectedItem.ListSubItems(13)
                TextoFiltro = ""
            Else
                .txtVLUnit = ListView2.SelectedItem.ListSubItems(6)
                .Cmb_un_com = ListView2.SelectedItem.ListSubItems(7)
                TextoFiltro = " and ID <> " & ListView2.SelectedItem
            End If
            ProcEnviaDadosRelacionamento ListView1.SelectedItem, ListView1.SelectedItem.ListSubItems(1), ListView1.SelectedItem.ListSubItems(5), ListView1.SelectedItem.ListSubItems(6), False, True, quantidade
        Else
            ProcEnviaDadosRelacionamento ListView1.SelectedItem, IIf(ListView1.SelectedItem.ListSubItems(1) = "", 0, ListView1.SelectedItem.ListSubItems(1)), ListView1.SelectedItem.ListSubItems(5), ListView1.SelectedItem.ListSubItems(6), False, True, quantidade
        End If
        PagarParcial = True
    End If
    
    If RelacionamentoSimultaneo = True And PagarParcial = True Then
        Unload Me
    Else
        If PagarParcial = True Then
            ProcCarregaLista
            ProcCarregaListaRelacionada
            .Produto_Relacionado = False
            '.ProcCarregaLista
            
            If .NF_alterada = True Then .ProcCarregaTotaisNota IIf(.txtId = "", 0, .txtId)
            '.ProcGravarTotaisNota
            '.ProcCarregaTotaisNota IIf(.txtid = "", 0, .txtid)
            '.ProcCarregaListaNota (IIf(ReturnNumbersOnly(Left(.lblPaginas(1).Caption, Len(.lblPaginas(1).Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(.lblPaginas(1).Caption, Len(.lblPaginas(1).Caption) - 5))))
        End If
    End If
    
'    If Left(.cmbFinalidade_emissao, 1) <> 1 Then
'        Set TBProduto = CreateObject("adodb.recordset")
'        TBProduto.Open "Select * from tbl_Detalhes_Nota where Int_codigo = " & .txtidproduto, Conexao, adOpenKeyset, adLockOptimistic
'        If TBProduto.EOF = False Then
'            .ProcAtualizaCST .txtIDEmpresa.Text, .txtIDcliente, .txt_Razao, .cbo_UF, Left(.Cmb_consumidor, 1), Left(.cmbFinalidade_emissao, 1)
'        End If
'        TBProduto.Close
'    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descri��o do erro : " + Error()), vbCritical, "CAPRIND V5.0"
    Exit Sub
End Sub

Private Sub ProcAdicionarNovo()
On Error GoTo tratar_erro

PagarParcial = False
Encontrou = False
With ListView1
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If PagarParcial = False Then
                If USMsgBox("Deseja realmente relacionar essa(s) nota(s)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
                If USMsgBox("Alguma nota selecionada ser� relacionada com quantidade parcial?", vbYesNo, "CAPRIND v5.0") = vbYes Then Encontrou = True
            End If
            PagarParcial = True
            
            IDlista = .ListItems(InitFor).SubItems(1)
            Desenho = .ListItems(InitFor).SubItems(2)
            If Encontrou = True Then
                Compras_Pedido = False
                Vendas_PI = False
                Compras_Cotacao = False
                Faturamento = True
                Qtde = .ListItems(InitFor).SubItems(11)
                Permitido2 = True
                frmVendas_PI_liberaritem.Show 1
                If Permitido2 = False Then Exit Sub
            Else
                ValorNC = .ListItems(InitFor).SubItems(11)
            End If
            
            With frmFaturamento_Prod_Serv
                Set TBExecucao = CreateObject("adodb.recordset")
'================================================================================
'Busca dados do item na nota de entrada referente a lista
'================================================================================
                TBExecucao.Open "Select * from tbl_Detalhes_Nota where Int_codigo = " & ListView1.ListItems(InitFor).ListSubItems(1), Conexao, adOpenKeyset, adLockOptimistic
                If TBExecucao.EOF = False Then
                    .txtidproduto = 0
                    .txtCod_Produto = TBExecucao!int_Cod_Produto
                    
                    'Verifica se existe cadastro da CFOP de devolu��o industrializa��o 5.902
                    If .Txt_ID_CFOP_prod = "" Then
                        Set TBFI = CreateObject("adodb.recordset")
                        TBFI.Open "Select IDCountCfop from tbl_NaturezaOperacao where id_CFOP = '5.902' or id_CFOP = '5902'", Conexao, adOpenKeyset, adLockOptimistic
                        If TBFI.EOF = False Then
                            .ProcCarregaDadosCFOPProdServ TBFI!IDCountCfop, True
                        End If
                    End If
                    
                    
                    .Txt_ID_CF = TBExecucao!ID_CF
'                    .ProcCarregaDadosCSTProd
                    
                    .txtDescricao_Produto = TBExecucao!Txt_descricao
                    .txtUN = TBExecucao!txt_Unid
                    .Cmb_un_com = TBExecucao!Unidade_com
                    .txtQTD = ValorNC
                    .cmbReferencia.Text = ListView1.ListItems(InitFor).ListSubItems(3)
                    .txtVLUnit = ListView1.ListItems(InitFor).ListSubItems(12)
                    .chkRetorno.Value = 1
                End If
                TBExecucao.Close
                ProcEnviaDadosRelacionamento ListView1.ListItems(InitFor), ListView1.ListItems(InitFor).ListSubItems(1), ListView1.ListItems(InitFor).ListSubItems(5), ListView1.ListItems(InitFor).ListSubItems(6), True, False, ValorNC
            End With
        End If
    Next InitFor
End With
If PagarParcial = False Then
    USMsgBox ("Informe a(s) nota(s) antes de relacionar."), vbExclamation, "CAPRIND V5.0"
Else
    USMsgBox ("Nota(s) relacionada(s) com sucesso."), vbInformation, "CAPRIND V5.0"
    With frmFaturamento_Prod_Serv
        .ProcCarregaLista
        If .NF_alterada = True Then .ProcCarregaTotaisNota IIf(.txtId = "", 0, .txtId)
        .ProcGravarTotaisNota
        .ProcCarregaTotaisNota IIf(.txtId = "", 0, .txtId)
        .ProcCarregaListaNota (IIf(ReturnNumbersOnly(Left(.lblPaginas(1).Caption, Len(.lblPaginas(1).Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(.lblPaginas(1).Caption, Len(.lblPaginas(1).Caption) - 5))))
    End With
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descri��o do erro : " + Error()), vbCritical, "CAPRIND V5.0"
    Exit Sub
End Sub

Private Sub ProcEnviaDadosRelacionamento(ID_nota_relacionada As Long, ID_produto_relacionado As Long, DtEmissaoNFrelacionada As Date, NFrelacionada As Long, SalvarSimultaneamente As Boolean, CarregarComplemento As Boolean, Qtde As Double)
On Error GoTo tratar_erro

TextoFiltro3 = ""
With frmFaturamento_Prod_Serv

If .txtidproduto = 0 Then
    .ProcSalvarProduto
End If

    If SSTab1.Tab = 0 Then
        If Left(.cmbFinalidade_emissao, 1) = 2 Then
            TextoFiltro = "ID_nota = " & .txtId & " and ID_nota_relacionada = " & ID_nota_relacionada & " Or ID_nota = " & ID_nota_relacionada & " And ID_nota_relacionada = " & .txtId
        Else
            TextoFiltro = "ID_nota = " & .txtId & " and ID_produto = " & .txtidproduto & " and ID_nota_relacionada = " & ID_nota_relacionada & " and ID_produto_relacionada = " & ID_produto_relacionado & " or ID_nota = " & ListView1.SelectedItem & " and ID_produto = " & ID_produto_relacionado & " and ID_nota_relacionada = " & .txtId & " and ID_produto_relacionada = " & .txtidproduto
        End If
    Else
        TextoFiltro = "ID = " & ListView2.SelectedItem
    End If
    Set TBGravar = CreateObject("adodb.recordset")
    TBGravar.Open "Select * from Faturamento_Relacionamento where " & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
    If TBGravar.EOF = False Then
        If SalvarSimultaneamente = False Then USMsgBox ("Altera��o efetuada com sucesso."), vbInformation, "CAPRIND V5.0"
        Evento = "Alterar"
        ID_documento = TBGravar!ID_nota
        TextoFiltro1 = "ID = " & TBGravar!ID_nota_relacionada
        TextoFiltro2 = "id_nota = " & TBGravar!ID_nota_relacionada
        If Left(.cmbFinalidade_emissao, 1) <> 2 Then TextoFiltro3 = " and Int_codigo = " & TBGravar!id_produto_relacionada
        complemento = ""
        Permitido = True
    Else
        TBGravar.AddNew
        If SalvarSimultaneamente = False Then USMsgBox "Novo relacionamento cadastrado com sucesso.", vbInformation, "CAPRIND v5.0"
        Evento = "Novo"
        ID_documento = ID_nota_relacionada
        TextoFiltro1 = "ID = " & ID_nota_relacionada
        TextoFiltro2 = "id_nota = " & ID_nota_relacionada
        If Left(.cmbFinalidade_emissao, 1) <> 2 Then TextoFiltro3 = " and Int_codigo = " & ID_produto_relacionado
        TBGravar!ID_nota_relacionada = ID_nota_relacionada
        complemento = "S/NF " & NFrelacionada & " - " & Format(DtEmissaoNFrelacionada, "dd/mm/yy")
        Permitido = False
    End If
    
    '==================================
    Modulo = Formulario & "/Relacionamento de nota fiscal"
    
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "select * from tbl_dados_nota_fiscal where ID = " & .txtId, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        If IsNull(TBAbrir!int_NotaFiscal) = True Or TBAbrir!int_NotaFiscal = "" Then NomeCampo = "N� ordem: " & TBAbrir!ID Else NomeCampo = "N� nota: " & TBAbrir!int_NotaFiscal
        Documento = NomeCampo & " - Tipo: " & TBAbrir!TipoNF & " - S�rie: " & TBAbrir!Serie
    End If
    
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "select * from tbl_dados_nota_fiscal where " & TextoFiltro1, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Documento1 = "N� nota: " & TBAbrir!int_NotaFiscal & " - Tipo: " & TBAbrir!TipoNF & " - S�rie: " & TBAbrir!Serie
    End If
    TBAbrir.Close
    
    ProcGravaEvento
    '==================================
    
    If Left(.cmbFinalidade_emissao, 1) <> 2 Then
        Set TBProduto = CreateObject("adodb.recordset")
        TBProduto.Open "Select Int_codigo, int_Qtd, Saldo from tbl_Detalhes_Nota where " & TextoFiltro2 & TextoFiltro3, Conexao, adOpenKeyset, adLockOptimistic
        If TBProduto.EOF = False Then
            
            'Verfifica saldo do produto da NF relacionada
            If Permitido = True Then
                ProcVerifSaldo "(ID_produto = " & TBProduto!Int_codigo & " or ID_produto_relacionada = " & TBProduto!Int_codigo & ") and ID <> " & TBGravar!ID
            Else
                ProcVerifSaldo "(ID_produto_relacionada = " & TBProduto!Int_codigo & " or ID_produto_relacionada = " & TBProduto!Int_codigo & ")"
            End If
            TBProduto!Saldo = Format(TBProduto!int_Qtd - (qt + Qtde), "###,##0.0000")
            TBProduto.Update
            
            TBGravar!id_produto_relacionada = TBProduto!Int_codigo
        End If
        'Debug.print .txtidproduto
        TBGravar!ID_Produto = .txtidproduto
    End If
    TBGravar!Qtde = Qtde
    TBGravar!ID_nota = .txtId
    
    'Salva saldo e o complemento da descri��o no produto da NF
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select int_Qtd, Saldo, Complemento_descricao from tbl_Detalhes_Nota where Int_codigo = " & .txtidproduto, Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
        
        If Left(.cmbFinalidade_emissao, 1) <> 2 Then
            'Verfifica saldo do produto da NF
            If Permitido = True Then
                ProcVerifSaldo "(Id_produto = " & .txtidproduto & " or ID_produto_relacionada = " & .txtidproduto & ") and ID <> " & TBGravar!ID
            Else
                ProcVerifSaldo "(Id_produto = " & .txtidproduto & " or ID_produto_relacionada = " & .txtidproduto & ")"
            End If
            TBProduto!Saldo = Format(TBProduto!int_Qtd - (qt + Qtde), "###,##0.0000")
        End If
        If complemento <> "" Then
            If IsNull(TBProduto!Complemento_descricao) = True Or TBProduto!Complemento_descricao = "" Then
                TBProduto!Complemento_descricao = complemento
            Else
                If TBProduto!Complemento_descricao <> complemento Then TBProduto!Complemento_descricao = TBProduto!Complemento_descricao & " ; " & complemento
            End If
        End If
        
        TBProduto.Update
    End If
    TBProduto.Close
    If CarregarComplemento = True Then ProcCarregaCampoComplemento
    
    TBGravar.Update
    TBGravar.Close
    .txtidproduto = 0
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descri��o do erro : " + Error()), vbCritical, "CAPRIND V5.0"
    Exit Sub
End Sub

Private Sub ProcVerifSaldo(Filtro As String)
On Error GoTo tratar_erro

qt = 0
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select Sum(Qtde) as qt from Faturamento_Relacionamento where " & Filtro, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    qt = IIf(IsNull(TBAbrir!qt), 0, TBAbrir!qt)
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descri��o do erro : " + Error()), vbCritical, "CAPRIND V5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaCampoComplemento()
On Error GoTo tratar_erro

With frmFaturamento_Prod_Serv
    If .txtidproduto = 0 Then Exit Sub
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select Complemento_descricao from tbl_Detalhes_Nota where Int_codigo = " & .txtidproduto, Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
        frmFaturamento_Prod_Serv.Txt_complemento_descricao = IIf(IsNull(TBProduto!Complemento_descricao), "", TBProduto!Complemento_descricao)
    End If
    TBProduto.Close
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descri��o do erro : " + Error()), vbCritical, "CAPRIND V5.0"
    Exit Sub
End Sub

Function FunVerificaCamposSalvar(Novo As Boolean) As Boolean
On Error GoTo tratar_erro

With frmFaturamento_Prod_Serv
    FunVerificaCamposSalvar = True
    
    Acao = "salvar"
    If Frame2.Enabled = False Then
        NomeCampo = "a nota fiscal na lista"
        ProcVerificaAcao
        FunVerificaCamposSalvar = False
        Exit Function
    End If
    
    'verifica quantidade digitada
    quantidade = IIf(txtQtde1 = "", 0, txtQtde1)
    If quantidade <= 0 Then
        NomeCampo = "a quantidade � relacionar"
        ProcVerificaAcao
        txtQtde1.SetFocus
        FunVerificaCamposSalvar = False
        Exit Function
    End If
    
    If Novo = True Then
    
        'verifica se a nota seleciona j� foi relacionada
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from Faturamento_Relacionamento where ID_nota = " & .txtId & " And ID_produto = " & .txtidproduto & " And ID_nota_relacionada = " & ListView1.SelectedItem & " And ID_produto_relacionada = " & ListView1.SelectedItem.ListSubItems(1) & " Or ID_nota = " & ListView1.SelectedItem & " And ID_produto = " & ListView1.SelectedItem.ListSubItems(1) & " And ID_nota_relacionada = " & .txtId & " And ID_produto_relacionada = " & .txtidproduto, Conexao, adOpenKeyset, adLockReadOnly
        If TBAbrir.EOF = False Then
            USMsgBox ("N�o � permitido relacionar esta nota " & ListView1.SelectedItem.ListSubItems(6) & ", pois a mesma j� esta relacionada."), vbExclamation, "CAPRIND V5.0"
            FunVerificaCamposSalvar = False
            TBAbrir.Close
            Exit Function
        End If
        TBAbrir.Close
    
        'verifica se o saldo da nota selecionada na tela de faturamento
        Qtde = Format(txtQtde, "###,##0.0000")
'        'Debug.print quantidade
'        'Debug.print Qtde
        If quantidade > Qtde Then
            USMsgBox ("N�o � permitido relacionar esta quantidade, pois a mesma � maior que o saldo dispon�vel na nota fiscal " & .txtNFiscal & "."), vbExclamation, "CAPRIND V5.0"
            FunVerificaCamposSalvar = False
            Exit Function
        End If
    Else
        'verifica se o saldo da nota selecionada na tela de faturamento, no bot�o alterar
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select FR.Qtde, NFP.Saldo from Faturamento_Relacionamento FR INNER JOIN tbl_Detalhes_Nota NFP ON NFP.Int_codigo = FR.ID_produto where ID = " & ListView2.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Qtde = TBAbrir!Saldo + TBAbrir!Qtde
            If quantidade > Qtde Then
                USMsgBox ("N�o � permitido relacionar esta quantidade, pois a mesma � maior que o saldo dispon�vel na nota fiscal " & .txtNFiscal & "."), vbExclamation, "CAPRIND V5.0"
                txtQtde1.SetFocus
                FunVerificaCamposSalvar = False
                TBAbrir.Close
                Exit Function
            End If
        End If
        TBAbrir.Close
    End If
    
    If txtQtde <> "" Then
        Qtde = Format(txtQtde, "###,##0.0000")
        If quantidade > Qtde Then
            If .opt_Saida.Value = True Then NomeCampo = "sa�da" Else NomeCampo = "entrada"
            USMsgBox ("N�o � permitido relacionar esta quantidade, pois a mesma � maior que a quantidade de " & NomeCampo & "."), vbExclamation, "CAPRIND V5.0"
            txtQtde1.SetFocus
            FunVerificaCamposSalvar = False
            Exit Function
        End If
    End If
    
    If Novo = True Then
        'verifica se o saldo da nota selecionada na lista
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select Saldo from tbl_Detalhes_Nota where Int_codigo = " & ListView1.SelectedItem.ListSubItems(1), Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Qtde = IIf(IsNull(TBAbrir!Saldo), 0, TBAbrir!Saldo)
            If quantidade > Qtde Then
                USMsgBox ("N�o � permitido relacionar esta quantidade, pois a mesma � maior que o saldo dispon�vel na nota fiscal " & ListView1.SelectedItem.ListSubItems(6) & "."), vbExclamation, "CAPRIND V5.0"
                txtQtde1.SetFocus
                FunVerificaCamposSalvar = False
                TBAbrir.Close
                Exit Function
            End If
        End If
        TBAbrir.Close
    Else
        'verifica se o saldo da nota selecionada na lista no bot�o alterar
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select FR.Qtde, NFP.Saldo from Faturamento_Relacionamento FR INNER JOIN tbl_Detalhes_Nota NFP ON NFP.Int_codigo = FR.ID_produto_relacionada where ID = " & ListView2.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Qtde = TBAbrir!Saldo + TBAbrir!Qtde
            If quantidade > Qtde Then
                USMsgBox ("N�o � permitido relacionar esta quantidade, pois a mesma � maior que o saldo dispon�vel na nota fiscal " & ListView2.SelectedItem.ListSubItems(2) & "."), vbExclamation, "CAPRIND V5.0"
                txtQtde1.SetFocus
                FunVerificaCamposSalvar = False
                TBAbrir.Close
                Exit Function
            End If
        End If
        TBAbrir.Close
    End If
End With

Exit Function
tratar_erro:
    USMsgBox ("Descri��o do erro : " + Error()), vbCritical, "CAPRIND V5.0"
    Exit Function
End Function

Private Sub cmbfiltrarpor_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
If cmbfiltrarpor = "Nota fiscal" And txtTexto <> "" Then
    VerifNumero = txtTexto
    ProcVerificaNumero
    If VerifNumero = False Then
        txtTexto = ""
        txtTexto.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descri��o do erro : " + Error()), vbCritical, "CAPRIND V5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro
Dim Strsql2 As Long

ListView1.ListItems.Clear
With frmFaturamento_Prod_Serv
'===========================================
' Se for nota de saida
'===========================================
    If opt_Saida.Value = True Then
     TextoFiltroTipo = "and NF.int_tiponota = 1 and NF.ID <> " & .txtId
    End If
    
'===========================================
' Se for nota de entrada
'===========================================
    If opt_Entrada.Value = True Then
     TextoFiltroTipo = "and NF.int_tiponota = 2 and NF.ID <> " & .txtId
    End If
    
'===========================================
' Se for nota de produtos
'===========================================
    TipoNF = "M1"
    TipoFiltro = " "
    TipoFiltro = " and NFP.Tipo = 'P'"

        TextoFiltroTriangulacao = ""
        TextoFiltroCod = ""
        TextoFiltroSaldo = ""
        
    If Left(.cmbFinalidade_emissao, 1) = 2 Then  'Se for nota Complementar
        TextoFiltroTriangulacao = ""
        TextoFiltroCod = ""
        TextoFiltroSaldo = ""
        CamposFiltro = "NF.ID, NF.dt_DataEmissao, NF.int_NotaFiscal, NF.TipoNF, NF.txt_Razao_Nome"
    Else
    If opt_Entrada.Value = True Then
        If Len(.txttipocliente) = 1 Then TextoFiltroTriangulacao = " or Len(NF.txt_tipocliente) = 2" Else TextoFiltroTriangulacao = " or Len(NF.txt_tipocliente) = 1"
    End If
    
        If .txtCod_Produto <> "" Then TextoFiltroCod = " and NFP.int_Cod_Produto = '" & .txtCod_Produto & "'" Else TextoFiltroCod = ""
        TextoFiltroSaldo = " and NFP.Saldo > 0"
        CamposFiltro = "NF.ID, NF.dt_DataEmissao, NF.int_NotaFiscal, NF.TipoNF, NF.txt_Razao_Nome, NFP.Int_codigo, NFP.int_Cod_Produto, NFP.txt_Descricao, NFP.int_Qtd, NFP.dbl_ValorUnitario, NFP.Unidade_com, NFP.Saldo, NFP.codproduto, NF.Id_Int_Cliente"
    End If
    
'==============================================================================
' Se for nota de saida relaciona com notas de terceiros (Venda com devolu��o de MP)
'==============================================================================
    'If .Opt_saida = True And Formulario = "Faturamento/Nota fiscal/Pr�pria" Or .Opt_saida = True And Formulario = "Estoque/Ordem de faturamento" Then
        If Left(.cmbFinalidade_emissao, 2) = 2 Or Left(.cmbFinalidade_emissao, 2) = 4 And Formulario = "Faturamento/Nota fiscal/Terceiros" Then
            If chkEmitente.Value = 0 Then
                TextoFiltroPadrao = "NF.tiponf = '" & TipoNF & "' and NF.DtValidacao IS NOT NULL and NF.ID_empresa = " & .txtIDEmpresa.Text & " and NF.Int_status = 1 " & TextoFiltroTipo & TextoFiltroCod & TipoFiltro & TextoFiltroSaldo & " and (NF.id_int_cliente = " & .txtIDcliente & " and NF.txt_Razao_Nome = '" & .txt_Razao & "'" & TextoFiltroTriangulacao & ") group by " & CamposFiltro & " order by NF.dt_DataEmissao, NF.int_NotaFiscal"
            Else
                TextoFiltroPadrao = "NF.tiponf = '" & TipoNF & "' and NF.DtValidacao IS NOT NULL and NF.ID_empresa = " & .txtIDEmpresa.Text & " and NF.Int_status = 1 " & TextoFiltroTipo & TextoFiltroCod & TipoFiltro & TextoFiltroSaldo & " group by " & CamposFiltro & " order by NF.dt_DataEmissao, NF.int_NotaFiscal"
            End If
            
        Else
            If chkEmitente.Value = 0 Then
              ' TextoFiltroPadrao = "NF.Aplicacao='T' and NF.tiponf = '" & TipoNF & "' and NF.DtValidacao IS NOT NULL and NF.ID_empresa = " & .txtIDEmpresa.Text & " and NF.Int_status = 1 " & TextoFiltroTipo & TextoFiltroCod & TipoFiltro & TextoFiltroSaldo & " and NF.id_int_cliente = " & .txtIDcliente & " and NF.txt_Razao_Nome = '" & .txt_Razao & "' group by " & CamposFiltro & " order by NF.dt_DataEmissao, NF.int_NotaFiscal"
               TextoFiltroPadrao = "NF.Aplicacao='T' and NF.tiponf = '" & TipoNF & "' and NF.DtValidacao IS NOT NULL and NF.ID_empresa = " & .txtIDEmpresa.Text & " and NF.Int_status = 1 " & TextoFiltroTipo & TextoFiltroCod & TipoFiltro & TextoFiltroSaldo & " group by " & CamposFiltro & " order by NF.dt_DataEmissao, NF.int_NotaFiscal"
            Else
               TextoFiltroPadrao = "NF.Aplicacao='T' and NF.tiponf = '" & TipoNF & "' and NF.DtValidacao IS NOT NULL and NF.ID_empresa = " & .txtIDEmpresa.Text & " and NF.Int_status = 1 " & TextoFiltroTipo & TextoFiltroCod & TipoFiltro & TextoFiltroSaldo & " group by " & CamposFiltro & " order by NF.dt_DataEmissao, NF.int_NotaFiscal"
            End If
        End If
    'End If
'==============================================================================
' Se for nota de entrada relaciona com notas de propria (Devolu��o)
'==============================================================================
    If .opt_Entrada = True And Formulario = "Faturamento/Nota fiscal/Pr�pria" Then
    TextoFiltroTipo = "and NF.int_tiponota = 1 and NF.ID <> " & .txtId
    
    If chkEmitente.Value = 0 Then
        TextoFiltroPadrao = "NF.Aplicacao='P' and NF.tiponf = '" & TipoNF & "' and NF.DtValidacao IS NOT NULL and NF.ID_empresa = " & .txtIDEmpresa.Text & " and NF.Int_status = 1 " & TextoFiltroTipo & TextoFiltroCod & TipoFiltro & TextoFiltroSaldo & " and (NF.id_int_cliente = " & .txtIDcliente & " and NF.txt_Razao_Nome = '" & .txt_Razao & "'" & TextoFiltroTriangulacao & ") group by " & CamposFiltro & " order by NF.dt_DataEmissao, NF.int_NotaFiscal"
    Else
        TextoFiltroPadrao = "NF.Aplicacao='P' and NF.tiponf = '" & TipoNF & "' and NF.DtValidacao IS NOT NULL and NF.ID_empresa = " & .txtIDEmpresa.Text & " and NF.Int_status = 1 " & TextoFiltroTipo & TextoFiltroCod & TipoFiltro & TextoFiltroSaldo & " group by " & CamposFiltro & " order by NF.dt_DataEmissao, NF.int_NotaFiscal"
    End If
   
   'TextoFiltroPadrao = "NF.Aplicacao='P' and NF.tiponf = '" & TipoNF & "' and NF.DtValidacao IS NOT NULL and NF.ID_empresa = " & .txtIDEmpresa.Text & " and NF.Int_status = 1 " & TextoFiltroTipo & TextoFiltroCod & TipoFiltro & TextoFiltroSaldo & " and (NF.id_int_cliente = " & .txtIDcliente & TextoFiltroTriangulacao & ") group by " & CamposFiltro & " order by NF.dt_DataEmissao, NF.int_NotaFiscal"
    End If
    
End With
INNERJOINTEXTO = "Select " & CamposFiltro & " FROM (((tbl_Dados_Nota_Fiscal NF LEFT JOIN tbl_detalhes_nota NFP ON NF.ID = NFP.ID_Nota) LEFT JOIN tbl_Totais_Nota TN ON NF.ID = TN.ID_Nota) LEFT JOIN tbl_Detalhes_Recebimento DR ON NF.ID = DR.ID_Nota) LEFT JOIN tbl_proposta_nota PN ON NF.ID = PN.ID_Nota"
'INNERJOINTEXTO = "Select " & CamposFiltro & " FROM (tbl_Dados_Nota_Fiscal NF LEFT JOIN tbl_detalhes_nota NFP ON NF.ID = NFP.ID_Nota) LEFT JOIN tbl_proposta_nota PN ON NF.ID = PN.ID_Nota"
Select Case cmbfiltrarpor
    Case "Nota fiscal":
        TextoFiltro = "NF.int_NotaFiscal"
        If txtTexto <> "" Then txtTexto = FunTamanhoTextoZeroEsq(txtTexto, 9)
    Case "Destinat�rio": TextoFiltro = "NF.txt_Razao_Nome"
    Case "Emitente": TextoFiltro = "NF.txt_Razao_Nome"
    Case "C�digo interno": TextoFiltro = "NFP.int_cod_produto"
    Case "C�digo de refer�ncia": TextoFiltro = "NFP.N_Referencia"
    Case "Descri��o": TextoFiltro = "NFP.txt_descricao"
    Case "Pedido cliente": TextoFiltro = "NFP.pccliente"
    Case "Nosso n�mero": TextoFiltro = "DR.Nosso_numero"
    Case "Pedido interno/Pedido de compra": TextoFiltro = "DR.Nosso_numero"
    TextoFiltro = TextoFiltro & "  AND NF.Aplicacao='T'"
End Select

If txtTexto <> "" Then
    StrSqlLocProdPadrao = INNERJOINTEXTO & " where " & TextoFiltro & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and " & TextoFiltroPadrao
Else
    StrSqlLocProdPadrao = INNERJOINTEXTO & " where " & TextoFiltroPadrao
End If

'Debug.print StrSqlLocProdPadrao

ProcCarregaLista

Exit Sub
tratar_erro:
    USMsgBox ("Descri��o do erro : " + Error()), vbCritical, "CAPRIND V5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLocalizar_produto_padrao.AbsolutePage <> 2 Then
    If TBLocalizar_produto_padrao.AbsolutePage = -3 Then
        ProcExibePagina (TBLocalizar_produto_padrao.PageCount - 1)
    Else
        TBLocalizar_produto_padrao.AbsolutePage = TBLocalizar_produto_padrao.AbsolutePage - 2
        ProcExibePagina (TBLocalizar_produto_padrao.AbsolutePage)
    End If
Else
    ProcExibePagina (1)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descri��o do erro : " + Error()), vbCritical, "CAPRIND V5.0"
    Exit Sub
End Sub

Private Sub cmdPagIr_Click()
On Error GoTo tratar_erro

If txtPagIr = "" Then Exit Sub
Quant = ReturnNumbersOnly(Right(lblPaginas.Caption, 4))
If Quant <= 1 Or txtPagIr > Quant Then Exit Sub
If txtPagIr.Text >= 1 And txtPagIr.Text <= Quant Then
    TBLocalizar_produto_padrao.AbsolutePage = txtPagIr.Text
    ProcExibePagina (TBLocalizar_produto_padrao.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descri��o do erro : " + Error()), vbCritical, "CAPRIND V5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLocalizar_produto_padrao.AbsolutePage = 1
ProcExibePagina (TBLocalizar_produto_padrao.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descri��o do erro : " + Error()), vbCritical, "CAPRIND V5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLocalizar_produto_padrao.AbsolutePage <> -3 Then
    If TBLocalizar_produto_padrao.AbsolutePage = 1 Then
        ProcExibePagina (2)
    Else
        ProcExibePagina (TBLocalizar_produto_padrao.AbsolutePage)
    End If
Else
    ProcExibePagina (TBLocalizar_produto_padrao.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descri��o do erro : " + Error()), vbCritical, "CAPRIND V5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLocalizar_produto_padrao.AbsolutePage = TBLocalizar_produto_padrao.PageCount
ProcExibePagina (TBLocalizar_produto_padrao.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descri��o do erro : " + Error()), vbCritical, "CAPRIND V5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case SSTab1.Tab
    Case 0:
        Select Case KeyCode
            Case vbKeyF2: ProcFiltrar
            Case vbKeyF3: ProcSalvarFaturamento True
            Case vbKeyEscape: ProcSair
        End Select
    Case 1:
        Select Case KeyCode
            Case vbKeyF3: If USToolBar2.ButtonState(1) = 0 Then ProcSalvarFaturamento False
            Case vbKeyF4: ProcExcluir
            Case vbKeyEscape: ProcSair
        End Select
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descri��o do erro : " + Error()), vbCritical, "CAPRIND V5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 11475, 6, True
ProcCarregaToolBar2 Me, 11475, 6, True

SSTab1.Tab = 0
ListView1.CheckBoxes = True
RelacionamentoSimultaneo = False
With frmFaturamento_Prod_Serv

txtIDNota = .txtId
txtidproduto = .txtidproduto

    If Left(.cmbFinalidade_emissao, 1) = 2 Then
        With ListView1
        .CheckBoxes = True
            .Height = 4080
            .ColumnHeaders(3).Width = 0
            .ColumnHeaders(4).Width = 0
            .ColumnHeaders(5).Width = 0
            .ColumnHeaders(9).Width = 9262
            .ColumnHeaders(10).Width = 0
            .ColumnHeaders(11).Width = 0
            .ColumnHeaders(12).Width = 0
            .ColumnHeaders(13).Width = 0
        End With
        Frame9.Top = 6810
        PBLista.Top = 7425
        Frame2.Visible = False
        
        USToolBar2.ButtonState(1) = 5
        With ListView2
            .Top = 1320
            .Height = 6090
            .ColumnHeaders(5).Width = 6955
            .ColumnHeaders(6).Width = 0
            .ColumnHeaders(7).Width = 0
            .ColumnHeaders(8).Width = 0
        End With
        PBLista1.Top = 7425
        Frame3.Visible = False
    End If
    If .opt_Saida = True Then
        If Left(.cmbFinalidade_emissao, 2) = 2 Then
            opt_Saida.Value = True
            opt_Entrada.Enabled = False
        Else
            opt_Entrada.Value = True
        End If
        
        Texto = "Sa�da"
        Label5.Caption = "Qtde. sa�da"
        txtQtde.ToolTipText = "Quantidade de sa�da."
        ListView1.ColumnHeaders(10).Text = "Qtde. entr."
        ListView1.ColumnHeaders(11).Text = "Qtde. sa�da"
        Label1.Caption = "Qtde. sa�da"
        Label1.Left = 3075
        txtQtde3.ToolTipText = "Quantidade de sa�da."
        If .txtidproduto = 0 Then
            With ListView1
                .CheckBoxes = True
                .ColumnHeaders(1).Width = 300
                .ColumnHeaders(5).Width = 1590
                .ColumnHeaders(9).Width = 1590
                .Height = 4050
            End With
            PBLista.Top = 7425
            Frame9.Top = 6810
            Frame2.Visible = False
            RelacionamentoSimultaneo = True
        End If
    Else
        If Left(.cmbFinalidade_emissao, 2) = 2 Or Left(.cmbFinalidade_emissao, 2) = 4 And Formulario = "Faturamento/Nota fiscal/Terceiros" Then
            opt_Saida.Enabled = False
            opt_Entrada.Value = True
        Else
            With ListView1
                .CheckBoxes = True
                .ColumnHeaders(1).Width = 300
                .ColumnHeaders(5).Width = 1590
                .ColumnHeaders(9).Width = 1590
                .Height = 4050
            End With
            opt_Saida.Value = True
        End If
        Texto = "Entrada"
        Label5.Caption = "Qtde. entr."
        txtQtde.ToolTipText = "Quantidade de entrada."
        txtQtde.Locked = False
        ListView1.ColumnHeaders(10).Text = "Qtde. sa�da"
        ListView1.ColumnHeaders(11).Text = "Qtde. entr."
        Label1.Caption = "Qtde. entrada"
        Label1.Left = 2955
        txtQtde3.ToolTipText = "Quantidade de entrada."
    End If
    If Formulario = "Faturamento/Nota fiscal/Pr�pria" Then
        Caption = "Nota fiscal - Pr�pria - Relacionamento de nota fiscal (Nota fiscal : " & .txtNFiscal & " - " & Texto & ")"
    ElseIf Formulario = "Faturamento/Nota fiscal/Terceiros" Then
            Caption = "Nota fiscal - Terceiros - Relacionamento de nota fiscal (Nota fiscal : " & .txtNFiscal & " - " & Texto & ")"
        ElseIf Formulario = "Estoque/Ordem de faturamento" Then
                Caption = "Ordem de faturamento - Relacionamento de nota fiscal (Ordem : " & .txtId & " - " & Texto & ")"
            Else
                Caption = "Nota fiscal - Relacionamento de nota fiscal (Nota fiscal : " & .txtNFiscal & " - " & Texto & ")"
                
    End If
    txtCodinterno = .txtCod_Produto
    txtdescricao = .txtDescricao_Produto
    txtQtde = .txtQTD
End With
cmbfiltrarpor = "Nota fiscal"
ProcFiltrar

ProcCarregaListaRelacionada

StrSql = "UPDATE Estoque_Controle SET  NF = EM.Documento from  Estoque_Controle EC INNER JOIN Estoque_movimentacao EM ON EC.IdEstoque = EM.IdEstoque Where EM.Operacao = 'ENTRADA_NOTA_FISCAL'"

'Debug.print StrSql

Conexao.Execute (StrSql)

Exit Sub
tratar_erro:
    USMsgBox ("Descri��o do erro : " + Error()), vbCritical, "CAPRIND V5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaListaRelacionada()
On Error GoTo tratar_erro

ListView2.ListItems.Clear
With frmFaturamento_Prod_Serv
    Qtde = IIf(txtQtde = "", 0, txtQtde)
    quantidade = 0
    
    If Left(.cmbFinalidade_emissao, 1) = 2 Then
        TextoFiltro = "ID_nota = " & .txtId & " or ID_nota_relacionada = " & .txtId
    Else
        TextoFiltro = "ID_nota = " & txtIDNota & " and ID_produto = " & txtidproduto & " or ID_nota_relacionada = " & .txtId & " and ID_produto_relacionada = " & .txtidproduto
    End If
    
    StrSql = "Select * from Faturamento_Relacionamento where " & TextoFiltro
    'Debug.print StrSql
    
    
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        PBLista1.Min = 0
        PBLista1.Max = TBAbrir.RecordCount
        PBLista1.Value = 1
        Contador = 0
        Do While TBAbrir.EOF = False
            With ListView2.ListItems
                .Add , , TBAbrir!ID
                
                With frmFaturamento_Prod_Serv
                    If TBAbrir!ID_nota = .txtId Then
                        If Left(.cmbFinalidade_emissao, 1) = 2 Then TextoFiltro = "NF.ID = " & TBAbrir!ID_nota_relacionada Else TextoFiltro = "NFP.Int_codigo = " & TBAbrir!id_produto_relacionada
                    Else
                        If Left(.cmbFinalidade_emissao, 1) = 2 Then TextoFiltro = "NF.ID = " & TBAbrir!ID_nota Else TextoFiltro = "NFP.Int_codigo = " & TBAbrir!ID_Produto
                    End If
                End With
                Set TBFI = CreateObject("adodb.recordset")
                TBFI.Open "Select NF.dt_DataEmissao, NF.int_NotaFiscal, NF.TipoNF, NF.txt_Razao_Nome, NFP.dbl_ValorUnitario, NFP.Unidade_com from tbl_Dados_Nota_Fiscal NF LEFT JOIN tbl_Detalhes_Nota NFP ON NFP.ID_nota = NF.ID where " & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
                If TBFI.EOF = False Then
                    .Item(.Count).SubItems(1) = IIf(IsNull(TBFI!dt_DataEmissao), "", (Format(TBFI!dt_DataEmissao, "dd/mm/yy")))
                    .Item(.Count).SubItems(2) = IIf(IsNull(TBFI!int_NotaFiscal), "", TBFI!int_NotaFiscal)
                    If IsNull(TBFI!TipoNF) = False Then
                        If TBFI!TipoNF = "M1" Then TipoNF2 = "Produto(s)"
                        If TBFI!TipoNF = "SA" Then TipoNF2 = "Servi�o(s)"
                        If TBFI!TipoNF = "M1SA" Then TipoNF2 = "Prod./Serv."
                    End If
                    .Item(.Count).SubItems(3) = TipoNF2
                    .Item(.Count).SubItems(4) = IIf(IsNull(TBFI!txt_Razao_Nome), "", TBFI!txt_Razao_Nome)
                    
                    If Left(frmFaturamento_Prod_Serv.cmbFinalidade_emissao, 1) <> 2 Then
                        .Item(.Count).SubItems(5) = IIf(IsNull(TBAbrir!Qtde), 0, Format(TBAbrir!Qtde, "###,##0.0000"))
                        .Item(.Count).SubItems(6) = IIf(IsNull(TBFI!dbl_ValorUnitario), 0, Format(TBFI!dbl_ValorUnitario, "###,##0.000000"))
                        .Item(.Count).SubItems(7) = IIf(IsNull(TBFI!Unidade_com), 0, TBFI!Unidade_com)
                    End If
                End If
                TBFI.Close
                
                quantidade = quantidade + TBAbrir!Qtde
            End With
            TBAbrir.MoveNext
            Contador = Contador + 1
            PBLista1.Value = Contador
        Loop
    End If
    
    txtQtde3 = Format(Qtde, "###,##0.0000")
    txtQtdeRel = Format(quantidade, "###,##0.0000")
    txtSaldo = Format(Qtde - quantidade, "###,##0.0000")
End With
    
Exit Sub
tratar_erro:
    USMsgBox ("Descri��o do erro : " + Error()), vbCritical, "CAPRIND V5.0"
    Exit Sub
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With ListView1
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then .ListItems.Item(InitFor).Checked = False Else .ListItems.Item(InitFor).Checked = True
        Next InitFor
    End With
Else
    ProcOrdenaListView ListView1, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descri��o do erro : " + Error()), vbCritical, "CAPRIND V5.0"
    Exit Sub
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Frame2.Visible = True Then
    With ListView1
        If .ListItems.Count = 0 Then Exit Sub
        txtCodinterno = .SelectedItem.ListSubItems(2)
        txtdescricao = .SelectedItem.ListSubItems(4)
        txtQtde = .SelectedItem.ListSubItems(11)
        txtQtde1 = .SelectedItem.ListSubItems(11)
    End With
    Frame2.Enabled = True
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descri��o do erro : " + Error()), vbCritical, "CAPRIND V5.0"
    Exit Sub
End Sub

Private Sub ListView2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With ListView2
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If frmFaturamento_Prod_Serv.txtNFiscal <> "" Then
                    If FunVerificaRegistroValidadoSemMsg("tbl_Dados_Nota_Fiscal", "ID = " & frmFaturamento_Prod_Serv.txtId, True) = False Then
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
    ProcOrdenaListView ListView2, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descri��o do erro : " + Error()), vbCritical, "CAPRIND V5.0"
    Exit Sub
End Sub

Private Sub ListView2_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With ListView2
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If frmFaturamento_Prod_Serv.txtNFiscal <> "" Then
'                If FunVerificaRegistroValidado("tbl_Dados_Nota_Fiscal", "ID = " & frmFaturamento_Prod_Serv.txtID, "nota fiscal", "este relacionamento", "excluir", False, True) = False Then
'                    .ListItems.Item(InitFor).Checked = False
'                    Exit Sub
'                End If
            End If
        End If
    Next InitFor
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descri��o do erro : " + Error()), vbCritical, "CAPRIND V5.0"
    Exit Sub
End Sub

Private Sub ListView2_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Left(frmFaturamento_Prod_Serv.cmbFinalidade_emissao, 1) = 2 Or ListView2.ListItems.Count = 0 Then Exit Sub
Txt_ID = ListView2.SelectedItem
txtQtde1 = ListView2.SelectedItem.ListSubItems(5)
Frame2.Enabled = True
txtQtde1.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descri��o do erro : " + Error()), vbCritical, "CAPRIND V5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaLista()
On Error GoTo tratar_erro

lblRegistros.Caption = "N� de reg.: 0"
lblPaginas.Caption = "P�gina: 0 de: 0"
ListView1.ListItems.Clear
If StrSqlLocProdPadrao = "" Then Exit Sub
Set TBLocalizar_produto_padrao = CreateObject("adodb.recordset")
'Debug.print StrSqlLocProdPadrao
TBLocalizar_produto_padrao.Open StrSqlLocProdPadrao, Conexao, adOpenKeyset, adLockReadOnly
If TBLocalizar_produto_padrao.EOF = False Then ProcExibePagina (1)

Exit Sub
tratar_erro:
    USMsgBox ("Descri��o do erro : " + Error()), vbCritical, "CAPRIND V5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina(Pagina)
On Error GoTo tratar_erro

ListView1.ListItems.Clear
TBLocalizar_produto_padrao.PageSize = IIf(txtNreg = "", 30, txtNreg)
TBLocalizar_produto_padrao.AbsolutePage = Pagina
TamanhoPagina = TBLocalizar_produto_padrao.PageSize
ContadorReg = 1

PBLista.Min = 0
PBLista.Max = FunVerifMaxPBListaPaginacao(TBLocalizar_produto_padrao.RecordCount - IIf(Pagina > 1, (TBLocalizar_produto_padrao.PageSize * (Pagina - 1)), 0), TBLocalizar_produto_padrao.PageSize)
PBLista.Value = 1
Contador = 0
Do While TBLocalizar_produto_padrao.EOF = False And (ContadorReg <= TamanhoPagina)
Prosseguir = True
If chkSaldoEstoque.Value = 1 Then
'=================================================================
' Vefifica Saldo do item no estoque
'=================================================================
Set TBSaldo = CreateObject("adodb.recordset")
StrSql = "Select EC.Desenho, EC.NF, sum(entrada-saida) as Saldo from estoque_movimentacao EM Inner Join Estoque_Controle EC on EC.IdEstoque = EM.IdEstoque where EC.Desenho = '" & TBLocalizar_produto_padrao!int_Cod_Produto & "' and EC.NF = '" & TBLocalizar_produto_padrao!int_NotaFiscal & "'  group by EC.Desenho, EC.NF"
'Debug.print StrSql

TBSaldo.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
If TBSaldo.EOF = False Then
    If TBSaldo!Saldo > 0 Then
        Prosseguir = True
    Else
        Prosseguir = False
    End If
Else
    Prosseguir = False
End If
TBSaldo.Close
End If


If Prosseguir = True Then

    With ListView1.ListItems
        .Add , , TBLocalizar_produto_padrao!ID
        If Left(frmFaturamento_Prod_Serv.cmbFinalidade_emissao, 1) <> 2 Then
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLocalizar_produto_padrao!Int_codigo), "", TBLocalizar_produto_padrao!Int_codigo)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLocalizar_produto_padrao!int_Cod_Produto), "", TBLocalizar_produto_padrao!int_Cod_Produto)
                If TBLocalizar_produto_padrao!Codproduto <> "" And TBLocalizar_produto_padrao!Id_Int_Cliente <> "" Then
                   Set TBFI = CreateObject("adodb.recordset")
                   TBFI.Open "Select IA.N_Referencia from item_aplicacoes IA INNER JOIN projproduto P ON IA.codproduto = P.codproduto where P.codproduto = " & TBLocalizar_produto_padrao!Codproduto & " and IA.ID_cliente_forn = " & TBLocalizar_produto_padrao!Id_Int_Cliente & " and IA.N_Referencia is not null group by IA.n_referencia", Conexao, adOpenKeyset, adLockOptimistic
                   If TBFI.EOF = True Then
                       Set TBFI = CreateObject("adodb.recordset")
                       TBFI.Open "Select IA.N_Referencia from item_aplicacoes IA INNER JOIN projproduto P ON IA.codproduto = P.codproduto where P.codproduto = " & TBLocalizar_produto_padrao!Codproduto & " and (IA.ID_cliente_forn = 0 or IA.ID_cliente_forn IS NULL) and IA.N_Referencia is not null group by IA.n_referencia", Conexao, adOpenKeyset, adLockOptimistic
                   End If
                   If TBFI.EOF = False Then
                       .Item(.Count).SubItems(3) = TBFI!N_referencia
                   End If
                   TBFI.Close
                 End If
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLocalizar_produto_padrao!Txt_descricao), "", TBLocalizar_produto_padrao!Txt_descricao)
            .Item(.Count).SubItems(9) = IIf(IsNull(TBLocalizar_produto_padrao!int_Qtd), 0, Format(TBLocalizar_produto_padrao!int_Qtd, "###,##0.0000"))
            .Item(.Count).SubItems(10) = Format(IIf(IsNull(TBLocalizar_produto_padrao!int_Qtd), 0, TBLocalizar_produto_padrao!int_Qtd) - IIf(IsNull(TBLocalizar_produto_padrao!Saldo), 0, TBLocalizar_produto_padrao!Saldo), "###,##0.0000")
            .Item(.Count).SubItems(11) = Format(IIf(IsNull(TBLocalizar_produto_padrao!Saldo), 0, TBLocalizar_produto_padrao!Saldo), "###,##0.0000")
            .Item(.Count).SubItems(12) = Format(IIf(IsNull(TBLocalizar_produto_padrao!dbl_ValorUnitario), 0, TBLocalizar_produto_padrao!dbl_ValorUnitario), "###,##0.000000")
            .Item(.Count).SubItems(13) = IIf(IsNull(TBLocalizar_produto_padrao!Unidade_com), "", TBLocalizar_produto_padrao!Unidade_com)
        End If
        .Item(.Count).SubItems(5) = IIf(IsNull(TBLocalizar_produto_padrao!dt_DataEmissao), "", Format(TBLocalizar_produto_padrao!dt_DataEmissao, "dd/mm/yy"))
        .Item(.Count).SubItems(6) = IIf(IsNull(TBLocalizar_produto_padrao!int_NotaFiscal), "", TBLocalizar_produto_padrao!int_NotaFiscal)
        If IsNull(TBLocalizar_produto_padrao!TipoNF) = False Then
            If TBLocalizar_produto_padrao!TipoNF = "M1" Then TipoNF2 = "Produto(s)"
            If TBLocalizar_produto_padrao!TipoNF = "SA" Then TipoNF2 = "Servi�o(s)"
            If TBLocalizar_produto_padrao!TipoNF = "M1SA" Then TipoNF2 = "Prod./Serv."
        End If
        .Item(.Count).SubItems(7) = TipoNF2
        .Item(.Count).SubItems(8) = IIf(IsNull(TBLocalizar_produto_padrao!txt_Razao_Nome), "", TBLocalizar_produto_padrao!txt_Razao_Nome)
        
    End With
    End If

    
    TBLocalizar_produto_padrao.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista.Value = Contador
Loop
lblRegistros.Caption = "N� de reg.: " & ContadorReg
If TBLocalizar_produto_padrao.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "P�g.: 1 de: " & TBLocalizar_produto_padrao.PageCount
ElseIf TBLocalizar_produto_padrao.AbsolutePage = adPosEOF Then
        lblPaginas.Caption = "P�g.: " & TBLocalizar_produto_padrao.PageCount & " de: " & TBLocalizar_produto_padrao.PageCount
    Else
        lblPaginas.Caption = "P�g.: " & TBLocalizar_produto_padrao.AbsolutePage - 1 & " de: " & TBLocalizar_produto_padrao.PageCount
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descri��o do erro : " + Error()), vbCritical, "CAPRIND V5.0"
    Exit Sub
End Sub

Private Sub Opt_entrada_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
ProcCarregaComboFiltrarPor (opt_Entrada.Value)
ProcCorrigeNomeColuna (opt_Entrada.Value)

Exit Sub
tratar_erro:
    USMsgBox ("Descri��o do erro : " + Error()), vbCritical, "CAPRIND V5.0"
    Exit Sub
End Sub

Private Sub opt_Saida_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
ProcCarregaComboFiltrarPor (opt_Entrada.Value)
ProcCorrigeNomeColuna (opt_Entrada.Value)

Exit Sub
tratar_erro:
    USMsgBox ("Descri��o do erro : " + Error()), vbCritical, "CAPRIND V5.0"
    Exit Sub
End Sub

Private Sub Optfim_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descri��o do erro : " + Error()), vbCritical, "CAPRIND V5.0"
    Exit Sub
End Sub

Private Sub Optinicio_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descri��o do erro : " + Error()), vbCritical, "CAPRIND V5.0"
    Exit Sub
End Sub

Private Sub Optmeio_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descri��o do erro : " + Error()), vbCritical, "CAPRIND V5.0"
    Exit Sub
End Sub

Private Sub optProduto_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descri��o do erro : " + Error()), vbCritical, "CAPRIND V5.0"
    Exit Sub
End Sub


Private Sub SSTab1_Click(PreviousTab As Integer)
On Error GoTo tratar_erro

Txt_ID = ""
Select Case SSTab1.Tab
    Case 0:
        If ListView1.Visible = True Then ListView1.SetFocus
        If Left(frmFaturamento_Prod_Serv.cmbFinalidade_emissao, 1) <> 2 Then
            With Frame2
                If RelacionamentoSimultaneo = True Then .Visible = False Else .Visible = True
                If ListView1.ListItems.Count = 0 Then
                    .Enabled = False
                    txtQtde1 = ""
                Else
                    .Enabled = True
                End If
                .Top = 7200
            End With
            Label7.Caption = "Qtde. � relacionar"
            txtQtde1.ToolTipText = "Quantidade � relacionar."
        End If
        ProcCarregaLista
    Case 1:
        ListView2.SetFocus
        If Left(frmFaturamento_Prod_Serv.cmbFinalidade_emissao, 1) <> 2 Then
            With Frame2
                .Visible = True
                .Enabled = False
                .Top = 1650
            End With
            Label7.Caption = "Qtde. relacionada"
            With txtQtde1
                .Text = ""
                .ToolTipText = "Quantidade relacionada."
            End With
        End If
        ProcCarregaListaRelacionada
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descri��o do erro : " + Error()), vbCritical, "CAPRIND V5.0"
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
    USMsgBox ("Descri��o do erro : " + Error()), vbCritical, "CAPRIND V5.0"
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
    USMsgBox ("Descri��o do erro : " + Error()), vbCritical, "CAPRIND V5.0"
    Exit Sub
End Sub

Private Sub txtQtde1_Change()
On Error GoTo tratar_erro

If txtQtde1.Text <> "" Then
    VerifNumero = txtQtde1.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtQtde1.Text = ""
        txtQtde1.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descri��o do erro : " + Error()), vbCritical, "CAPRIND V5.0"
    Exit Sub
End Sub

Private Sub txtQtde1_LostFocus()
On Error GoTo tratar_erro

txtQtde1 = Format(txtQtde1, "###,##0.0000")
    
Exit Sub
tratar_erro:
    USMsgBox ("Descri��o do erro : " + Error()), vbCritical, "CAPRIND V5.0"
    Exit Sub
End Sub

Private Sub txtTexto_Change()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
If cmbfiltrarpor = "Nota fiscal" And txtTexto <> "" Then
    VerifNumero = txtTexto
    ProcVerificaNumero
    If VerifNumero = False Then
        txtTexto = ""
        txtTexto.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descri��o do erro : " + Error()), vbCritical, "CAPRIND V5.0"
    Exit Sub
End Sub

Private Sub txtTexto_LostFocus()
On Error GoTo tratar_erro

If cmbfiltrarpor = "Nota fiscal" And txtTexto <> "" Then txtTexto = FunTamanhoTextoZeroEsq(txtTexto, 9)

Exit Sub
tratar_erro:
    USMsgBox ("Descri��o do erro : " + Error()), vbCritical, "CAPRIND V5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcFiltrar
    Case 2:
        If Formulario <> "Estoque/Ordem de faturamento" Then
            ProcSalvarFaturamento True
        Else
            ProcSalvarOrdem True
        End If
    'Case 4: ProcAjuda
    Case 5: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descri��o do erro : " + Error()), vbCritical, "CAPRIND V5.0"
    Exit Sub
End Sub

Private Sub USToolBar2_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcSalvarFaturamento False
    Case 2: ProcExcluir
    'Case 4: ProcAjuda
    Case 5: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descri��o do erro : " + Error()), vbCritical, "CAPRIND V5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaComboFiltrarPor(Entrada As Boolean)
On Error GoTo tratar_erro

With cmbfiltrarpor
    .Clear
    .AddItem "C�digo de refer�ncia"
    .AddItem "C�digo interno"
    .AddItem "Descri��o"
    If Entrada = True Then .AddItem "Emitente" Else .AddItem "Destinat�rio"
    .AddItem "Nosso n�mero"
    .AddItem "Nota fiscal"
    .AddItem "Pedido Cliente"
    .AddItem "Pedido interno/Pedido de compra"
    .Text = "Nota fiscal"
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descri��o do erro : " + Error()), vbCritical, "CAPRIND V5.0"
    Exit Sub
End Sub

Private Sub ProcCorrigeNomeColuna(Entrada As Boolean)
On Error GoTo tratar_erro

With ListView1.ColumnHeaders
    If Entrada = True Then .Item(9) = "Emitente" Else .Item(9) = "Destinat�rio"
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descri��o do erro : " + Error()), vbCritical, "CAPRIND V5.0"
    Exit Sub
End Sub
