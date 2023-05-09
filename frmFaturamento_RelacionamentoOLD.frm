VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2014.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form frmFaturamento_Relacionamento 
   BackColor       =   &H00F9F9F9&
   BorderStyle     =   0  'Nenhum
   Caption         =   "Administrativo - Faturamento - Nota fiscal - Relacionamento de nota fiscal"
   ClientHeight    =   8310
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   14760
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
   ScaleHeight     =   8310
   ScaleWidth      =   14760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Centralziar na Tela
   Begin DrawSuite2014.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   465
      Left            =   0
      TabIndex        =   51
      Top             =   0
      Width           =   14760
      _ExtentX        =   26035
      _ExtentY        =   820
      DibPicture      =   "frmFaturamento_Relacionamento.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Icon            =   "frmFaturamento_Relacionamento.frx":1B63
      ShowMaximize    =   0   'False
      ShowMinimize    =   0   'False
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7755
      Left            =   60
      TabIndex        =   17
      Top             =   480
      Width           =   14625
      _ExtentX        =   25797
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
      TabCaption(0)   =   "Notas fiscais à relacionar"
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
      Tab(0).Control(7)=   "Frame11"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Frame5"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Frame2"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Notas fiscais relacionadas"
      TabPicture(1)   =   "frmFaturamento_Relacionamento.frx":1E99
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txt_ID"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "USToolBar2"
      Tab(1).Control(2)=   "PBLista1"
      Tab(1).Control(3)=   "ListView2"
      Tab(1).Control(4)=   "Frame3"
      Tab(1).ControlCount=   5
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
         Left            =   60
         TabIndex        =   44
         Top             =   5940
         Width           =   14505
         Begin VB.TextBox txtCodinterno 
            Alignment       =   2  'Centralizar
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
            TabIndex        =   53
            TabStop         =   0   'False
            ToolTipText     =   "Código interno."
            Top             =   375
            Width           =   1695
         End
         Begin VB.TextBox txtDescricao 
            Alignment       =   2  'Centralizar
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
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   47
            TabStop         =   0   'False
            ToolTipText     =   "Descrição."
            Top             =   375
            Width           =   7635
         End
         Begin VB.TextBox txtQtde 
            Alignment       =   2  'Centralizar
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
            Left            =   11040
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   375
            Width           =   1665
         End
         Begin VB.TextBox txtQtde1 
            Alignment       =   2  'Centralizar
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
            Left            =   9540
            TabIndex        =   45
            Top             =   375
            Width           =   1215
         End
         Begin DrawSuite2014.USButton Btnrelacionar 
            Height          =   555
            Left            =   12930
            TabIndex        =   52
            ToolTipText     =   "Relacionar item com o mesmo item de outra nota"
            Top             =   210
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   979
            DibPicture      =   "frmFaturamento_Relacionamento.frx":1EB5
            BorderColor     =   5263559
            BorderColorDisabled=   13160660
            BorderColorDown =   4013465
            BorderColorOver =   4408288
            Caption         =   "Relacionar"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   16777215
            ForeColorDown   =   16777215
            ForeColorOver   =   16777215
            GradientColor1  =   5263559
            GradientColor2  =   5263559
            GradientColor3  =   5263559
            GradientColor4  =   5263559
            GradientColorDisabled1=   13160660
            GradientColorDisabled2=   13160660
            GradientColorDisabled3=   13160660
            GradientColorDisabled4=   13160660
            GradientColorDown1=   4013465
            GradientColorDown2=   4013465
            GradientColorDown3=   4013465
            GradientColorDown4=   4013465
            GradientColorOver1=   4408288
            GradientColorOver2=   4408288
            GradientColorOver3=   4408288
            GradientColorOver4=   4408288
            PicSize         =   2
            PicSizeH        =   24
            PicSizeW        =   24
            Theme           =   4
            ToolTipTitle    =   "CAPRIND v5.0"
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparente
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
            Left            =   577
            TabIndex        =   54
            Top             =   180
            Width           =   900
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Centralizar
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparente
            Caption         =   "Qtde á Relacionar"
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
            Left            =   11227
            TabIndex        =   50
            Top             =   180
            Width           =   1290
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Centralizar
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparente
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
            Left            =   2490
            TabIndex        =   49
            Top             =   180
            Width           =   6435
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Centralizar
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparente
            Caption         =   "Saldo"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   9945
            TabIndex        =   48
            Top             =   180
            Width           =   405
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Digite o texto para pesquisa"
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
         Height          =   750
         Left            =   11040
         TabIndex        =   42
         Top             =   1320
         Width           =   3525
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
            Left            =   60
            TabIndex        =   43
            ToolTipText     =   "Texto para pesquisa."
            Top             =   300
            Width           =   3375
         End
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Escolha uma opção abaixo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   3960
         TabIndex        =   37
         Top             =   1320
         WhatsThisHelpID =   210
         Width           =   4185
         Begin VB.OptionButton Optfim 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Fim frase"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2400
            TabIndex        =   41
            Top             =   330
            Width           =   975
         End
         Begin VB.OptionButton Optinicio 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Início frase"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   330
            Value           =   -1  'True
            Width           =   1125
         End
         Begin VB.OptionButton Optmeio 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Meio frase"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1290
            TabIndex        =   39
            Top             =   330
            Width           =   1125
         End
         Begin VB.OptionButton optIgual 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Igual"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3450
            TabIndex        =   38
            Top             =   330
            Width           =   705
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Finalidade"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   60
         TabIndex        =   35
         Top             =   1320
         Width           =   1815
         Begin VB.OptionButton Opt_saida 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Saida"
            BeginProperty Font 
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
            Left            =   150
            TabIndex        =   9
            Top             =   360
            Value           =   -1  'True
            Width           =   705
         End
         Begin VB.OptionButton Opt_entrada 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Entrada"
            BeginProperty Font 
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
            Left            =   900
            TabIndex        =   10
            Top             =   360
            Width           =   885
         End
      End
      Begin DrawSuite2014.USProgressBar PBLista 
         Height          =   255
         Left            =   60
         TabIndex        =   26
         Top             =   7395
         Width           =   14485
         _ExtentX        =   25559
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
         Height          =   3840
         Left            =   60
         TabIndex        =   1
         Top             =   2070
         Width           =   14505
         _ExtentX        =   25585
         _ExtentY        =   6773
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483641
         BackColor       =   16777215
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   15
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
            Text            =   "Cód. interno"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Tag             =   "T"
            Text            =   "Cód. ref."
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "Descrição"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Object.Tag             =   "D"
            Text            =   "Dt. emissão"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Object.Tag             =   "T"
            Text            =   "Nota fiscal"
            Object.Width           =   1764
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
            Text            =   "Destinatário"
            Object.Width           =   2743
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   9
            Object.Tag             =   "N"
            Text            =   "Qtde."
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   10
            Object.Tag             =   "N"
            Text            =   "Qtde."
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   11
            Object.Tag             =   "N"
            Text            =   "Saldo"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   12
            Object.Tag             =   "N"
            Text            =   "Vlr. unit."
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   13
            Object.Tag             =   "T"
            Text            =   "Un. com"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   14
            Text            =   "Chave de acesso"
            Object.Width           =   6244
         EndProperty
      End
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
         Height          =   615
         Left            =   60
         TabIndex        =   31
         Top             =   6780
         Width           =   14505
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
            Left            =   7410
            TabIndex        =   3
            ToolTipText     =   "Número da página."
            Top             =   180
            Width           =   555
         End
         Begin VB.TextBox txtNreg 
            Alignment       =   2  'Centralizar
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
            Left            =   3390
            TabIndex        =   2
            Text            =   "30"
            ToolTipText     =   "Número de registros por página."
            Top             =   180
            Width           =   555
         End
         Begin DrawSuite2014.USButton cmdPagProx 
            Height          =   315
            Left            =   9630
            TabIndex        =   7
            ToolTipText     =   "Próxima página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmFaturamento_Relacionamento.frx":3A18
            BorderColor     =   14404026
            BorderColorDown =   11632444
            BorderColorOver =   11632444
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
            GradientColor2  =   16777215
            GradientColor3  =   16777215
            GradientColorDown2=   16246986
            GradientColorDown3=   15189380
            GradientColorDown4=   14596208
            GradientColorOver1=   16643560
            GradientColorOver2=   16576988
            GradientColorOver3=   16441780
            GradientColorOver4=   16178091
            PicSizeH        =   19
            PicSizeW        =   19
         End
         Begin DrawSuite2014.USButton cmdPagAnt 
            Height          =   315
            Left            =   9090
            TabIndex        =   6
            ToolTipText     =   "Página anterior."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmFaturamento_Relacionamento.frx":71BF
            BorderColor     =   14404026
            BorderColorDown =   11632444
            BorderColorOver =   11632444
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
            GradientColor2  =   16777215
            GradientColor3  =   16777215
            GradientColorDown2=   16246986
            GradientColorDown3=   15189380
            GradientColorDown4=   14596208
            GradientColorOver1=   16643560
            GradientColorOver2=   16576988
            GradientColorOver3=   16441780
            GradientColorOver4=   16178091
            PicSizeH        =   19
            PicSizeW        =   19
         End
         Begin DrawSuite2014.USButton cmdPagIr 
            Height          =   315
            Left            =   7980
            TabIndex        =   4
            Top             =   180
            Width           =   465
            _ExtentX        =   820
            _ExtentY        =   556
            BorderColor     =   14404026
            BorderColorDown =   11632444
            BorderColorOver =   11632444
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
            GradientColor2  =   16777215
            GradientColor3  =   16777215
            GradientColorDown2=   16246986
            GradientColorDown3=   15189380
            GradientColorDown4=   14596208
            GradientColorOver1=   16643560
            GradientColorOver2=   16576988
            GradientColorOver3=   16441780
            GradientColorOver4=   16178091
            PicSizeH        =   19
            PicSizeW        =   19
         End
         Begin DrawSuite2014.USButton cmdPagPrim 
            Height          =   315
            Left            =   8550
            TabIndex        =   5
            ToolTipText     =   "Primeira página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmFaturamento_Relacionamento.frx":ACCE
            BorderColor     =   14404026
            BorderColorDown =   11632444
            BorderColorOver =   11632444
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
            GradientColor2  =   16777215
            GradientColor3  =   16777215
            GradientColorDown2=   16246986
            GradientColorDown3=   15189380
            GradientColorDown4=   14596208
            GradientColorOver1=   16643560
            GradientColorOver2=   16576988
            GradientColorOver3=   16441780
            GradientColorOver4=   16178091
            PicSizeH        =   19
            PicSizeW        =   19
         End
         Begin DrawSuite2014.USButton cmdPagUlt 
            Height          =   315
            Left            =   10170
            TabIndex        =   8
            ToolTipText     =   "Última página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmFaturamento_Relacionamento.frx":EDC2
            BorderColor     =   14404026
            BorderColorDown =   11632444
            BorderColorOver =   11632444
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
            GradientColor2  =   16777215
            GradientColor3  =   16777215
            GradientColorDown2=   16246986
            GradientColorDown3=   15189380
            GradientColorDown4=   14596208
            GradientColorOver1=   16643560
            GradientColorOver2=   16576988
            GradientColorOver3=   16441780
            GradientColorOver4=   16178091
            PicSizeH        =   19
            PicSizeW        =   19
         End
         Begin VB.Label lblPaginas 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparente
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
            Left            =   11220
            TabIndex        =   34
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lblRegistros 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparente
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
            TabIndex        =   33
            Top             =   240
            Width           =   1275
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparente
            Caption         =   "Carregar               registros por página"
            BeginProperty Font 
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
            Left            =   2700
            TabIndex        =   32
            Top             =   240
            Width           =   2760
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Filtrar por"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   8160
         TabIndex        =   30
         Top             =   1320
         Width           =   2865
         Begin VB.ComboBox cmbfiltrarpor 
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
            ItemData        =   "frmFaturamento_Relacionamento.frx":1265C
            Left            =   60
            List            =   "frmFaturamento_Relacionamento.frx":12678
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   0
            ToolTipText     =   "Opções para filtro."
            Top             =   300
            Width           =   2745
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Tipo da nota fiscal"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   1890
         TabIndex        =   29
         Top             =   1320
         Width           =   2055
         Begin VB.OptionButton optProduto_servico 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Produtos/Serviços"
            BeginProperty Font 
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
            Left            =   3240
            TabIndex        =   36
            Top             =   210
            Width           =   1605
         End
         Begin VB.OptionButton OptServico 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Serviços"
            BeginProperty Font 
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
            Left            =   1080
            TabIndex        =   12
            Top             =   360
            Width           =   915
         End
         Begin VB.OptionButton optProduto 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Produtos"
            BeginProperty Font 
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
            Left            =   90
            TabIndex        =   11
            Top             =   360
            Value           =   -1  'True
            Width           =   945
         End
      End
      Begin VB.TextBox txt_ID 
         Alignment       =   2  'Centralizar
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
         TabIndex        =   24
         TabStop         =   0   'False
         ToolTipText     =   "Código interno."
         Top             =   3060
         Visible         =   0   'False
         Width           =   645
      End
      Begin DrawSuite2014.USToolBar USToolBar1 
         Height          =   975
         Left            =   60
         TabIndex        =   25
         Top             =   330
         Width           =   14505
         _ExtentX        =   25585
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
         Begin DrawSuite2014.USImageList USImageList1 
            Left            =   3720
            Top             =   330
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmFaturamento_Relacionamento.frx":12710
            Count           =   1
         End
      End
      Begin DrawSuite2014.USToolBar USToolBar2 
         Height          =   975
         Left            =   -74940
         TabIndex        =   27
         Top             =   330
         Width           =   14505
         _ExtentX        =   25585
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
         Begin DrawSuite2014.USImageList USImageList2 
            Left            =   4170
            Top             =   90
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmFaturamento_Relacionamento.frx":15532
            Count           =   1
         End
      End
      Begin DrawSuite2014.USProgressBar PBLista1 
         Height          =   255
         Left            =   -74940
         TabIndex        =   28
         Top             =   6585
         Width           =   14505
         _ExtentX        =   25585
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
         Height          =   5220
         Left            =   -74940
         TabIndex        =   13
         Top             =   1320
         Width           =   14505
         _ExtentX        =   25585
         _ExtentY        =   9208
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
            Text            =   "Dt. emissão"
            Object.Width           =   1940
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
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "Destinatário/Emitente"
            Object.Width           =   11915
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Object.Tag             =   "N"
            Text            =   "Qtde. relac."
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Object.Tag             =   "N"
            Text            =   "Vlr. unit."
            Object.Width           =   2646
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
         TabIndex        =   18
         Top             =   6840
         Width           =   14505
         Begin VB.TextBox txtQtde3 
            Alignment       =   1  'Alinhar à Direita
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
            Left            =   4290
            Locked          =   -1  'True
            TabIndex        =   14
            TabStop         =   0   'False
            Text            =   "0,000"
            Top             =   390
            Width           =   1635
         End
         Begin VB.TextBox txtQtdeRel 
            Alignment       =   1  'Alinhar à Direita
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
            Left            =   6375
            Locked          =   -1  'True
            TabIndex        =   15
            TabStop         =   0   'False
            Text            =   "0,000"
            ToolTipText     =   "Quantidade relacionada."
            Top             =   390
            Width           =   1635
         End
         Begin VB.TextBox txtSaldo 
            Alignment       =   1  'Alinhar à Direita
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
            Left            =   8520
            Locked          =   -1  'True
            TabIndex        =   16
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
            BackStyle       =   0  'Transparente
            Caption         =   "Qtde. saída"
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
            Left            =   4635
            TabIndex        =   23
            Top             =   180
            Width           =   945
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparente
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
            Left            =   6105
            TabIndex        =   22
            Top             =   450
            Width           =   75
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparente
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
            Left            =   6450
            TabIndex        =   21
            Top             =   180
            Width           =   1485
         End
         Begin VB.Label Label8 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparente
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
            Left            =   8190
            TabIndex        =   20
            Top             =   450
            Width           =   135
         End
         Begin VB.Label Label9 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparente
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
            Left            =   9105
            TabIndex        =   19
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
Dim id_produto_entrada As Long

Private Sub ProcExcluir()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso.")
    Exit Sub
End If
Permitido = False
With ListView2
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir esta(s) nota(s) relacionada(s)?", vbQuestion + vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select * from Faturamento_Relacionamento WHERE ID = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False Then
                '==================================
                Modulo = Formulario
                Evento = "Excluir"
                ID_documento = .ListItems(InitFor)
                
                Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open "select int_NotaFiscal, ID, TipoNF, Serie from tbl_dados_nota_fiscal where ID = " & frmFaturamento_Prod_Serv.TxtID, Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    If IsNull(TBAbrir!int_NotaFiscal) = True Or TBAbrir!int_NotaFiscal = "" Then NomeCampo = "N° ordem: " & TBAbrir!ID Else NomeCampo = "N° nota: " & TBAbrir!int_NotaFiscal
                    Documento = NomeCampo & " - Tipo: " & TBAbrir!TipoNF & " - Série: " & TBAbrir!Serie
                End If
                
                Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open "select int_NotaFiscal, ID, TipoNF, Serie from tbl_dados_nota_fiscal where ID = " & TBFI!ID_nota_relacionada, Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    Documento1 = "Nº nota relacionada: " & TBAbrir!int_NotaFiscal & " - Tipo: " & TBAbrir!TipoNF & " - Série: " & TBAbrir!Serie
                End If
                TBAbrir.Close
                
                ProcGravaEvento
                '==================================
                
                If Left(frmFaturamento_Prod_Serv.cmbFinalidade_emissao, 1) = 2 Or (frmFaturamento_Prod_Serv.Opt_saida = False And Faturamento_NF_Saida = True) Then
                
                    If frmFaturamento_Prod_Serv.Opt_saida = False And Faturamento_NF_Saida = True Then
                        If TBFI!ID_nota = frmFaturamento_Prod_Serv.TxtID Then
                            procExcluirDevolucaoNF frmFaturamento_Prod_Serv.TxtID, TBFI!ID_nota_relacionada
                        Else
                            procExcluirDevolucaoNF frmFaturamento_Prod_Serv.TxtID, TBFI!ID_nota
                        End If
                    End If
                
                    Set TBProduto = CreateObject("adodb.recordset")
                    TBProduto.Open "Select Complemento_descricao from tbl_Detalhes_Nota where Int_codigo = " & frmFaturamento_Prod_Serv.txtIDProduto, Conexao, adOpenKeyset, adLockOptimistic
                    If TBProduto.EOF = False Then
                        Complemento = ""
                        Set TBAbrir = CreateObject("adodb.recordset")
                        TBAbrir.Open "Select NF.dt_DataEmissao, NF.int_NotaFiscal from tbl_Dados_Nota_Fiscal NF INNER JOIN Faturamento_Relacionamento FR ON FR.Id_nota_relacionada = NF.ID where FR.Id_nota = " & TBFI!ID_nota & " and FR.ID <> " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
                        If TBAbrir.EOF = False Then
                            Do While TBAbrir.EOF = False
                                If IsNull(TBAbrir!int_NotaFiscal) = False Then OF = TBAbrir!int_NotaFiscal
                                If Complemento = "" Then Complemento = "S/NF " & OF & " - " & Format(TBAbrir!dt_DataEmissao, "dd/mm/yy") Else Complemento = Complemento & " ; S/NF " & OF & " - " & Format(TBAbrir!dt_DataEmissao, "dd/mm/yy")
                                TBAbrir.MoveNext
                            Loop
                        End If
                        TBAbrir.Close
                        TBProduto!Complemento_descricao = IIf(Complemento = "", Null, Complemento)
                        TBProduto.Update
                    End If
                Else
                    'Atualiza o saldo no produto da NF relacionada
                    Set TBProduto = CreateObject("adodb.recordset")
                    TBProduto.Open "Select Saldo, Complemento_descricao from tbl_Detalhes_Nota where Int_codigo = " & TBFI!id_produto_relacionada, Conexao, adOpenKeyset, adLockOptimistic
                    If TBProduto.EOF = False Then
                        TBProduto!Saldo = Format(TBProduto!Saldo + TBFI!Qtde, "0.0000")
                        
                        Complemento = ""
                        Set TBAbrir = CreateObject("adodb.recordset")
                        TBAbrir.Open "Select NF.dt_DataEmissao, NF.int_NotaFiscal from tbl_Dados_Nota_Fiscal NF INNER JOIN Faturamento_Relacionamento FR ON FR.Id_nota = NF.ID where FR.Id_produto_relacionada = " & TBFI!id_produto_relacionada & " and FR.ID <> " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
                        If TBAbrir.EOF = False Then
                            Do While TBAbrir.EOF = False
                                If IsNull(TBAbrir!int_NotaFiscal) = False Then OF = TBAbrir!int_NotaFiscal
                                If Complemento = "" Then Complemento = "S/NF " & OF & " - " & Format(TBAbrir!dt_DataEmissao, "dd/mm/yy") Else Complemento = Complemento & " ; S/NF " & OF & " - " & Format(TBAbrir!dt_DataEmissao, "dd/mm/yy")
                                TBAbrir.MoveNext
                            Loop
                        End If
                        TBAbrir.Close
                        TBProduto!Complemento_descricao = IIf(Complemento = "", Null, Complemento)
                        TBProduto.Update
                    End If
                    
                    'Atualiza o complemento da descrição e o saldo no produto da NF
                    Set TBProduto = CreateObject("adodb.recordset")
                    TBProduto.Open "Select Saldo,  Complemento_descricao from tbl_Detalhes_Nota where Int_codigo = " & TBFI!ID_produto, Conexao, adOpenKeyset, adLockOptimistic
                    If TBProduto.EOF = False Then
                        TBProduto!Saldo = Format(TBProduto!Saldo + TBFI!Qtde, "0.0000")
                        
                        Complemento = ""
                        Set TBAbrir = CreateObject("adodb.recordset")
                        TBAbrir.Open "Select NF.dt_DataEmissao, NF.int_NotaFiscal from tbl_Dados_Nota_Fiscal NF INNER JOIN Faturamento_Relacionamento FR ON FR.Id_nota_relacionada = NF.ID where FR.Id_produto = " & TBFI!ID_produto & " and FR.ID <> " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
                        If TBAbrir.EOF = False Then
                            Do While TBAbrir.EOF = False
                                If IsNull(TBAbrir!int_NotaFiscal) = False Then OF = TBAbrir!int_NotaFiscal
                                If Complemento = "" Then Complemento = "S/NF " & OF & " - " & Format(TBAbrir!dt_DataEmissao, "dd/mm/yy") Else Complemento = Complemento & " ; S/NF " & OF & " - " & Format(TBAbrir!dt_DataEmissao, "dd/mm/yy")
                                TBAbrir.MoveNext
                            Loop
                        End If
                        TBAbrir.Close
                        TBProduto!Complemento_descricao = IIf(Complemento = "", Null, Complemento)
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
    USMsgBox ("Informe a(s) nota(s) relacionada(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Nota(s) relacionada(s) excluída(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    With frmFaturamento_Prod_Serv
        If Left(.cmbFinalidade_emissao, 1) <> 1 Then
            Set TBProduto = CreateObject("adodb.recordset")
            TBProduto.Open "Select * from tbl_Detalhes_Nota where Int_codigo = " & .txtIDProduto, Conexao, adOpenKeyset, adLockOptimistic
            If TBProduto.EOF = False Then
                .ProcAtualizaCST .Cmb_empresa.ItemData(.Cmb_empresa.ListIndex), .txtIDCliente, .txt_Razao, .cbo_UF, Left(.Cmb_consumidor, 1), Left(.cmbFinalidade_emissao, 1)
            End If
            TBProduto.Close
        End If
    End With
    txt_ID = ""
    txtQtde1 = ""
    ProcCarregaLista
    ProcCarregaListaRelacionada
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcSalvar(NovoRelacionamento As Boolean)
On Error GoTo tratar_erro

If SSTab1.Tab = 0 Then
    If ListView1.ListItems.Count = 0 Then Exit Sub
Else
    If ListView2.ListItems.Count = 0 Then Exit Sub
End If
PagarParcial = False
With frmFaturamento_Prod_Serv
    If FunVerificaRegistroValidado("tbl_Dados_Nota_Fiscal", "ID = " & .TxtID, "nota fiscal", IIf(SSTab1.Tab = 0, "", "este relacionamento"), IIf(SSTab1.Tab = 0, "relacionar", "alterar"), False, True) = False Then Exit Sub
    
    .Produto_Relacionado = True
    
    If SSTab1.Tab = 0 And .txtIDProduto = 0 Then
        ProcAdicionarNovo
    Else
        If SSTab1.Tab = 0 Then
'============================================================
'            Nota fiscal de complemento
'============================================================
            If Left(.cmbFinalidade_emissao, 1) = 2 Then
                TextoFiltro = "ID_nota = " & .TxtID & " and ID_nota_relacionada = " & ListView1.SelectedItem & " Or ID_nota = " & ListView1.SelectedItem & " And ID_nota_relacionada = " & .TxtID
            Else
                TextoFiltro = "ID_nota = " & .TxtID & " and ID_produto = " & .txtIDProduto & " and ID_nota_relacionada = " & ListView1.SelectedItem & " and ID_produto_relacionada = " & ListView1.SelectedItem.ListSubItems(1) & " or ID_nota = " & ListView1.SelectedItem & " and ID_produto = " & ListView1.SelectedItem.ListSubItems(1) & " and ID_nota_relacionada = " & .TxtID & " and ID_produto_relacionada = " & .txtIDProduto
            End If
        Else
            TextoFiltro = "ID = " & ListView2.SelectedItem
        End If
'============================================================
'        Nota fiscal de complemento ou se for uma nota fiscal de entrada de devolução
'============================================================
        If Left(.cmbFinalidade_emissao, 1) = 2 Or (.Opt_saida = False And Faturamento_NF_Saida = True) Then
            ProcEnviaDadosRelacionamento ListView1.SelectedItem, IIf(ListView1.SelectedItem.ListSubItems(1) = "", 0, ListView1.SelectedItem.ListSubItems(1)), ListView1.SelectedItem.ListSubItems(5), ListView1.SelectedItem.ListSubItems(6), False, True, quantidade
        Else
            '            Set TBGravar = CreateObject("adodb.recordset")
'            TBGravar.Open "Select * from Faturamento_Relacionamento where " & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
'            If TBGravar.EOF = False Then
'                If FunVerificaCamposSalvar(False) = False Then Exit Sub
'            Else
'                If FunVerificaCamposSalvar(True) = False Then Exit Sub
'            End If
            
            If FunVerificaCamposSalvar(NovoRelacionamento) = False Then Exit Sub
             If SSTab1.Tab = 0 Then
                .txtVLUnit = ListView1.SelectedItem.ListSubItems(12)
                .Cmb_un_com = ListView1.SelectedItem.ListSubItems(13)
                TextoFiltro = ""
             Else
                .txtVLUnit = ListView2.SelectedItem.ListSubItems(6)
                .Cmb_un_com = ListView2.SelectedItem.ListSubItems(7)
                TextoFiltro = " and ID <> " & ListView2.SelectedItem
             End If
'===========================================================================
' Relaciona e devolve para o estoque
'===========================================================================
            ProcEnviaDadosRelacionamento ListView1.SelectedItem, ListView1.SelectedItem.ListSubItems(1), ListView1.SelectedItem.ListSubItems(5), ListView1.SelectedItem.ListSubItems(6), False, True, quantidade
          End If
        PagarParcial = True
       End If
    
    If RelacionamentoSimultaneo = True And PagarParcial = True Then
        Unload Me
    Else
        If PagarParcial = True Then
            ProcCarregaLista
            ProcCarregaListaRelacionada
            frmFaturamento_Prod_Serv.Produto_Relacionado = False
            frmFaturamento_Prod_Serv.ProcCarregaLista
            
            If .NF_alterada = True Then frmFaturamento_Prod_Serv.ProcCarregaTotaisNota IIf(.TxtID = "", 0, .TxtID)
            frmFaturamento_Prod_Serv.ProcGravarTotaisNota
            frmFaturamento_Prod_Serv.ProcCarregaTotaisNota IIf(.TxtID = "", 0, frmFaturamento_Prod_Serv.TxtID)
            frmFaturamento_Prod_Serv.ProcCarregaListaNota (IIf(DS_RetornarNumeros(Left(.lblPaginas(1).Caption, Len(.lblPaginas(1).Caption) - 5)) <= 1, 1, DS_RetornarNumeros(Left(.lblPaginas(1).Caption, Len(.lblPaginas(1).Caption) - 5))))
            End If
        End If
    
    If Left(frmFaturamento_Prod_Serv.cmbFinalidade_emissao, 1) <> 1 Then
        Set TBProduto = CreateObject("adodb.recordset")
        TBProduto.Open "Select * from tbl_Detalhes_Nota where Int_codigo = " & .txtIDProduto, Conexao, adOpenKeyset, adLockOptimistic
        If TBProduto.EOF = False Then
            .ProcAtualizaCST .Cmb_empresa.ItemData(.Cmb_empresa.ListIndex), .txtIDCliente, .txt_Razao, .cbo_UF, Left(.Cmb_consumidor, 1), Left(.cmbFinalidade_emissao, 1)
        End If
        TBProduto.Close
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcAdicionarNovo()
On Error GoTo tratar_erro

PagarParcial = False
Encontrou = False
With ListView1
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If PagarParcial = False Then
                If USMsgBox("Deseja realmente relacionar essa(s) nota(s)?", vbQuestion + vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
                If USMsgBox("Alguma nota selecionada será relacionada com quantidade parcial?", vbQuestion + vbYesNo, "CAPRIND v5.0") = vbYes Then Encontrou = True
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
                TBExecucao.Open "Select * from tbl_Detalhes_Nota where Int_codigo = " & ListView1.ListItems(InitFor).ListSubItems(1), Conexao, adOpenKeyset, adLockOptimistic
                If TBExecucao.EOF = False Then
                    .txtIDProduto = 0
                    .txtCod_Produto = TBExecucao!int_Cod_Produto
                    
                    'Verifica se existe cadastro da CFOP de devolução industrialização 5.902
                    If .Txt_ID_CFOP_prod = "" Then
                        Set TBFI = CreateObject("adodb.recordset")
                        TBFI.Open "Select IDCountCfop from tbl_NaturezaOperacao where id_CFOP = '5.902' or id_CFOP = '5902'", Conexao, adOpenKeyset, adLockOptimistic
                        If TBFI.EOF = False Then
                            .ProcCarregaDadosCFOPProdServ TBFI!IDCountCfop, True
                        End If
                    End If
                    
                    Set TBFI = CreateObject("adodb.recordset")
                    TBFI.Open "Select Codproduto from Projproduto where Desenho = '" & TBExecucao!int_Cod_Produto & "'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBFI.EOF = False Then
                        ProcCarregaComboCodRef .cmbreferencia, "P.codproduto = " & TBFI!Codproduto, .txtIDCliente, IIf(Len(.txttipocliente) = 2, "C", "F"), True, True
                    End If
                    TBFI.Close
                    
                    .Txt_ID_CF = TBExecucao!ID_CF
                    .ProcCarregaDadosCSTProd
                    
                    .txtDescricao_Produto = TBExecucao!txt_Descricao
                    .txtun = TBExecucao!txt_Unid
                    .Cmb_un_com = TBExecucao!Unidade_com
                    .txtQtd = ValorNC
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
    USMsgBox ("Informe a(s) nota(s) antes de relacionar."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Nota(s) relacionada(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    With frmFaturamento_Prod_Serv
        .ProcCarregaLista
        If .NF_alterada = True Then .ProcCarregaTotaisNota IIf(.TxtID = "", 0, .TxtID)
        .ProcGravarTotaisNota
        .ProcCarregaTotaisNota IIf(.TxtID = "", 0, .TxtID)
        .ProcCarregaListaNota (IIf(DS_RetornarNumeros(Left(.lblPaginas(1).Caption, Len(.lblPaginas(1).Caption) - 5)) <= 1, 1, DS_RetornarNumeros(Left(.lblPaginas(1).Caption, Len(.lblPaginas(1).Caption) - 5))))
    End With
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcEnviaDadosRelacionamento(ID_nota_relacionada As Long, ID_produto_relacionado As Long, DtEmissaoNFrelacionada As Date, NFrelacionada As Long, SalvarSimultaneamente As Boolean, CarregarComplemento As Boolean, Qtde As Double)
On Error GoTo tratar_erro

TextoFiltro3 = ""
With frmFaturamento_Prod_Serv

'.ProcSalvarProduto
'==========================================================================
' Verifica o a aba se é de a relacionar ou relacionadas
'==========================================================================
If SSTab1.Tab = 0 Then
'==========================================================================
' Verifica o tipo de nota fiscal e prepara o texto para filtro
'==========================================================================
' Se for uma nota fiscal complementar ou de entrada propria devolução
'==========================================================================
  If Left(.cmbFinalidade_emissao, 1) = 2 Or (.Opt_saida = False And Faturamento_NF_Saida = True) Then
    TextoFiltro = "ID_nota = " & .TxtID & " and ID_nota_relacionada = " & ID_nota_relacionada & " Or ID_nota = " & ID_nota_relacionada & " And ID_nota_relacionada = " & .TxtID
  End If
  '=========================================================================
  ' Se for uma nota fiscal de Saida propria devolução
  '==========================================================================
  If Left(.cmbFinalidade_emissao, 1) = 4 And (.Opt_saida = True And Faturamento_NF_Saida = True) Then
    TextoFiltro = "ID_nota = " & .TxtID & " and ID_produto = " & .txtIDProduto & " and ID_nota_relacionada = " & ID_nota_relacionada & " and ID_produto_relacionada = " & ID_produto_relacionado & " or ID_nota = " & ListView1.SelectedItem & " and ID_produto = " & ID_produto_relacionado & " and ID_nota_relacionada = " & .TxtID & " and ID_produto_relacionada = " & .txtIDProduto
  End If
End If

'=========================================================================
' Somente busca as relacionadas
'=========================================================================
If SSTab1.Tab = 1 Then
  TextoFiltro = "ID = " & ListView2.SelectedItem
End If
'==========================================================================
' Verifica se já existe relacionamento
'==========================================================================
    Set TBGravar = CreateObject("adodb.recordset")
    TBGravar.Open "Select * from Faturamento_Relacionamento where " & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
    If TBGravar.EOF = False Then
        If SalvarSimultaneamente = False Then USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
        Evento = "Alterar"
        ID_documento = TBGravar!ID_nota
        TextoFiltro1 = "ID = " & TBGravar!ID_nota_relacionada
        TextoFiltro2 = "id_nota = " & TBGravar!ID_nota_relacionada
        If Left(.cmbFinalidade_emissao, 1) <> 2 Then TextoFiltro3 = " and Int_codigo = " & TBGravar!id_produto_relacionada
        Complemento = ""
        Permitido = True
'==========================================================================
' Verifica se não existe relacionamentocria um novo
'==========================================================================
    Else
        TBGravar.AddNew
        If SalvarSimultaneamente = False Then USMsgBox ("Novo relacionamento cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
        Evento = "Novo"
        ID_documento = ID_nota_relacionada
        TextoFiltro1 = "ID = " & ID_nota_relacionada
        TextoFiltro2 = "id_nota = " & ID_nota_relacionada
        If Left(.cmbFinalidade_emissao, 1) <> 2 Then TextoFiltro3 = " and Int_codigo = " & ID_produto_relacionado
        TBGravar!ID_nota_relacionada = ID_nota_relacionada
        Complemento = "S/NF " & NFrelacionada & " - " & Format(DtEmissaoNFrelacionada, "dd/mm/yy")
        Permitido = False
    End If
    
    '==================================
    Modulo = Formulario & "/Relacionamento de nota fiscal"
    
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "select * from tbl_dados_nota_fiscal where ID = " & .TxtID, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        If IsNull(TBAbrir!int_NotaFiscal) = True Or TBAbrir!int_NotaFiscal = "" Then NomeCampo = "N° ordem: " & TBAbrir!ID Else NomeCampo = "N° nota: " & TBAbrir!int_NotaFiscal
        Documento = NomeCampo & " - Tipo: " & TBAbrir!TipoNF & " - Série: " & TBAbrir!Serie
    End If
    
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "select * from tbl_dados_nota_fiscal where " & TextoFiltro1, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Documento1 = "Nº nota: " & TBAbrir!int_NotaFiscal & " - Tipo: " & TBAbrir!TipoNF & " - Série: " & TBAbrir!Serie
    End If
    TBAbrir.Close
    
    ProcGravaEvento
 '==========================================================================
' Verifica novamente se não é nota complementar, e se não é de entrada própria
'==========================================================================
   
    If Left(.cmbFinalidade_emissao, 1) <> 2 And Not (.Opt_saida = False And Faturamento_NF_Saida = True) Then
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
        TBGravar!ID_produto = .txtIDProduto
    End If
    TBGravar!Qtde = Qtde
    TBGravar!ID_nota = .TxtID
'============================================================================
'Salva saldo e o complemento da descrição no produto da NF
'============================================================================
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select int_Qtd, Saldo, Complemento_descricao from tbl_Detalhes_Nota where Int_codigo = " & .txtIDProduto, Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
        
        If Left(.cmbFinalidade_emissao, 1) <> 2 And Not (.Opt_saida = False And Faturamento_NF_Saida = True) Then
            'Verfifica saldo do produto da NF
            If Permitido = True Then
                ProcVerifSaldo "(Id_produto = " & .txtIDProduto & " or ID_produto_relacionada = " & .txtIDProduto & ") and ID <> " & TBGravar!ID
            Else
                ProcVerifSaldo "(Id_produto = " & .txtIDProduto & " or ID_produto_relacionada = " & .txtIDProduto & ")"
            End If
            TBProduto!Saldo = Format(TBProduto!int_Qtd - (qt + Qtde), "###,##0.0000")
        End If
        If Complemento <> "" Then
            If IsNull(TBProduto!Complemento_descricao) = True Or TBProduto!Complemento_descricao = "" Then
                TBProduto!Complemento_descricao = Complemento
            Else
                If TBProduto!Complemento_descricao <> Complemento Then TBProduto!Complemento_descricao = TBProduto!Complemento_descricao & " ; " & Complemento
            End If
        End If
        
        TBProduto.Update
    End If
    TBProduto.Close
    If CarregarComplemento = True Then ProcCarregaCampoComplemento
    
    TBGravar.Update
    TBGravar.Close
'=======================================================================================================================
' Se for nota fiscal de entrada, própria faz o relacionamento
'=======================================================================================================================
If .Opt_saida.Value = False And Faturamento_NF_Saida = True And Permitido = False Then
 ProcDevolucao ID_nota_relacionada
End If
'=======================================================================================================================
' Aqui começa a devolução do produto
'=======================================================================================================================
' Começa a retirar do financeiro o contas a receber dos produtos devolvidos
'=======================================================================================================================

'If Left(.cmbFinalidade_emissao, 1) = 4 Then 'Se for nota de devolução
 If Left(frmFaturamento_Prod_Serv.cmbFinalidade_emissao, 1) = 4 And Not (frmFaturamento_Prod_Serv.Opt_saida = True And Faturamento_NF_Saida = True) Then

'Abrir a tabela de contas a receber
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from tbl_contas_receber where NFiscal = '" & ListView1.SelectedItem.SubItems(3) & "'", Conexao, adOpenKeyset, adLockOptimistic
 If TBAbrir.EOF = False Then
  Valor1 = ListView1.SelectedItem.SubItems(12)
  Valor2 = TBAbrir!valor
  Valor3 = txtQtde1
  ValorTotal = Valor1 * Valor3
  
   If Valor2 > ValorTotal Then
    TBAbrir!observacoes = "Valor alterado de R$ " & TBAbrir!valor & " para R$ " & Format(TBAbrir!valor - ValorTotal, "###,##0.0000") & " devido a devolução da Nota fiscal n°" & ListView1.SelectedItem.SubItems(6)
    TBAbrir!valor = TBAbrir!valor - ValorTotal
    TBAbrir.Update
   Else
   ' Conexao.Execute "Delete tbl_contas_receber where Nfiscal = " & Format(ListView1.SelectedItem.SubItems(6), "###,##0.0000")
    Conexao.Execute "Delete tbl_contas_receber where Nfiscal = '" & ListView1.SelectedItem.SubItems(3) & "'"
   End If
   TBAbrir.Close
 End If
 
'=======================================================================================================================
' Aqui começa a devolução do produto ao estoque movimentação e estoque controle
'=======================================================================================================================
Set TBEstoque = CreateObject("adodb.recordset")
TBEstoque.Open "Select * from estoque_controle", Conexao, adOpenKeyset, adLockOptimistic
TBEstoque.AddNew
TBEstoque!LOTE = ListView1.SelectedItem.SubItems(6)

Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select * from Estoque_movimentacao", Conexao, adOpenKeyset, adLockOptimistic
TBProduto.AddNew

TBProduto!Destino = "Interno"
TBProduto!Documento = frmFaturamento_Prod_Serv.txtNFiscal.Text
TBProduto!Terceiros = False
TBEstoque!ID_empresa = frmFaturamento_Prod_Serv.Cmb_empresa.ItemData(frmFaturamento_Prod_Serv.Cmb_empresa.ListIndex)

TBProduto!LOTE = ListView1.SelectedItem.SubItems(6)
    
TBEstoque!Desenho = ListView1.SelectedItem.SubItems(2)
TBProduto!Desenho = ListView1.SelectedItem.SubItems(2)

'Atualiza valor do produto no estoque
Set TBItem = CreateObject("adodb.recordset")
TBItem.Open "Select Codproduto, Estoque, classe from projproduto where desenho = '" & ListView1.SelectedItem.SubItems(2) & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBItem.EOF = False Then
    If TBItem!estoque = True Then ControlaEstoque = True Else ControlaEstoque = False
    TBEstoque!Classe = TBItem!Classe
    TBProduto!Familia = TBItem!Classe
    
    'Grava código de referência no produto
    If ListView1.SelectedItem.SubItems(3) <> "" Then
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from item_aplicacoes where Codproduto = " & TBItem!Codproduto & " and n_referencia = '" & ListView1.SelectedItem.SubItems(3) & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = True Then TBAbrir.AddNew
        TBAbrir!Codproduto = TBItem!Codproduto
        TBAbrir!N_referencia = ListView1.SelectedItem.SubItems(3)
        TBAbrir!Descricao = ListView1.SelectedItem.SubItems(4)
        TBAbrir.Update
        TBAbrir.Close
    End If
    
    TBEstoque!Ref = ListView1.SelectedItem.SubItems(3)
End If
TBItem.Close

ValorTotal = Valor1
quantestoque = Valor3
TBProduto!VlrUnit = Format(ValorTotal, "###,##0.0000000000")
TBProduto!VlrTotal = Format(ValorTotal * quantestoque, "###,##0.00")
TBEstoque!Descricao = ListView1.SelectedItem.SubItems(4)
TBProduto!Descricao = ListView1.SelectedItem.SubItems(4)
TBEstoque!data = data
TBEstoque!Responsavel = pubUsuario
TBProduto!Responsavel = pubUsuario
TBEstoque!Certificado = 0
TBEstoque!Numero_serie = 0
TBEstoque!Corrida = 0
TBEstoque!local_armaz = "ESTOQUE PADRÃO"
TBProduto!Entrada = Valor3
TBProduto!Entrada_PC = Valor3
TBEstoque!Qtde = Valor3
Qtde = Valor3
Entrada = Valor3
TBProduto!Operacao = "ENTRADA_DEVOLUÇÃO"
TBEstoque!status = "ENTRADA_DEVOLUÇÃO"
TBEstoque!Cliente = ListView1.SelectedItem.SubItems(8)
If ControlaEstoque = True Then QtdeEstoque = Valor3 Else QtdeEstoque = 0
TBProduto!estoque_venda = Valor3
TBProduto!Obs = ""
TBProduto.Update


TBEstoque!estoque_real = Valor3
TBEstoque!estoque_real_PC = Valor3
TBEstoque!estoque_venda = Valor3
TBEstoque!valor_unitario = Format(Valor1, "###,##0.0000000000")
TBEstoque!Valor_total = Format(Valor1 * Valor3, "###,##0.00")
TBEstoque.Update

Conexao.Execute "UPDATE Estoque_movimentacao Set IDEstoque = " & TBEstoque!IDestoque & " where IDoperacao = " & TBProduto!IDoperacao
IDestoque = TBEstoque!IDestoque
Modulo = "Estoque/Movimentação/Entrada"
Evento = "Devolução"
ID_documento = IDestoque
Documento = "Cód. interno: " & ListView1.SelectedItem.SubItems(2) & " - Lote: " & ListView1.SelectedItem.SubItems(6) & "Qtde.: " & Format(Valor3, "###,##0.0000")
Documento1 = ""
ProcGravaEvento

End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
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
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcCarregaCampoComplemento()
On Error GoTo tratar_erro

With frmFaturamento_Prod_Serv
    If .txtIDProduto = 0 Then Exit Sub
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select Complemento_descricao from tbl_Detalhes_Nota where Int_codigo = " & .txtIDProduto, Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
        frmFaturamento_Prod_Serv.Txt_complemento_descricao = IIf(IsNull(TBProduto!Complemento_descricao), "", TBProduto!Complemento_descricao)
    End If
    TBProduto.Close
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
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
        NomeCampo = "a quantidade à relacionar"
        ProcVerificaAcao
        txtQtde1.SetFocus
        FunVerificaCamposSalvar = False
        Exit Function
    End If
    
    If Novo = True Then
    
        'verifica se a nota seleciona já foi relacionada
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from Faturamento_Relacionamento where ID_nota = " & .TxtID & " And ID_produto = " & .txtIDProduto & " And ID_nota_relacionada = " & ListView1.SelectedItem & " And ID_produto_relacionada = " & ListView1.SelectedItem.ListSubItems(1) & " Or ID_nota = " & ListView1.SelectedItem & " And ID_produto = " & ListView1.SelectedItem.ListSubItems(1) & " And ID_nota_relacionada = " & .TxtID & " And ID_produto_relacionada = " & .txtIDProduto, Conexao, adOpenKeyset, adLockReadOnly
        If TBAbrir.EOF = False Then
            USMsgBox ("Não é permitido relacionar esta nota " & ListView1.SelectedItem.ListSubItems(6) & ", pois a mesma já esta relacionada."), vbExclamation, "CAPRIND v5.0"
            FunVerificaCamposSalvar = False
            TBAbrir.Close
            Exit Function
        End If
        TBAbrir.Close
    
        'verifica se o saldo da nota selecionada na tela de faturamento
        Qtde = IIf(txtSaldo = "", 0, txtSaldo)
        If quantidade > Qtde Then
            USMsgBox ("Não é permitido relacionar esta quantidade, pois a mesma é maior que o saldo disponível na nota fiscal " & .txtNFiscal & "."), vbExclamation, "CAPRIND v5.0"
            FunVerificaCamposSalvar = False
            Exit Function
        End If
    Else
        'verifica se o saldo da nota selecionada na tela de faturamento, no botão alterar
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select FR.Qtde, NFP.Saldo from Faturamento_Relacionamento FR INNER JOIN tbl_Detalhes_Nota NFP ON NFP.Int_codigo = FR.ID_produto where ID = " & ListView2.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Qtde = TBAbrir!Saldo + TBAbrir!Qtde
            If quantidade > Qtde Then
                USMsgBox ("Não é permitido relacionar esta quantidade, pois a mesma é maior que o saldo disponível na nota fiscal " & .txtNFiscal & "."), vbExclamation, "CAPRIND v5.0"
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
            If .Opt_saida.Value = True Then NomeCampo = "saída" Else NomeCampo = "entrada"
            USMsgBox ("Não é permitido relacionar esta quantidade, pois a mesma é maior que a quantidade de " & NomeCampo & "."), vbExclamation, "CAPRIND v5.0"
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
                USMsgBox ("Não é permitido relacionar esta quantidade, pois a mesma é maior que o saldo disponível na nota fiscal " & ListView1.SelectedItem.ListSubItems(6) & "."), vbExclamation, "CAPRIND v5.0"
                txtQtde1.SetFocus
                FunVerificaCamposSalvar = False
                TBAbrir.Close
                Exit Function
            End If
        End If
        TBAbrir.Close
    Else
        'verifica se o saldo da nota selecionada na lista no botão alterar
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select FR.Qtde, NFP.Saldo from Faturamento_Relacionamento FR INNER JOIN tbl_Detalhes_Nota NFP ON NFP.Int_codigo = FR.ID_produto_relacionada where ID = " & ListView2.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Qtde = TBAbrir!Saldo + TBAbrir!Qtde
            If quantidade > Qtde Then
                USMsgBox ("Não é permitido relacionar esta quantidade, pois a mesma é maior que o saldo disponível na nota fiscal " & ListView2.SelectedItem.ListSubItems(2) & "."), vbExclamation, "CAPRIND v5.0"
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
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Private Sub BtnCorrigirSaldo_Click()
On Error GoTo tratar_erro
PBLista.Min = 0
PBLista.Max = TBAbrir.RecordCount
PBLista.Value = 1
Contador = 0

Contador = ListView1.ListItems.Count
Do While Contador > 0
Var = ListView1.ListItems.Item(Contador).SubItems(1)

If USMsgBox("Deseja realmente corrigir todos o saldo dos itens da lista?", vbYesNo, "CAPRIND v5.0") = vbYes Then
'    Set TBAbrir = CreateObject("adodb.recordset")
'    TBAbrir.Open "Select * from Faturamento_Relacionamento order by ID", Conexao, adOpenKeyset, adLockOptimistic
'    Do While TBAbrir.EOF = False
        Set TBLISTA = CreateObject("adodb.recordset")
'            TBLISTA.Open "Select * from tbl_detalhes_nota where Int_codigo = '" & TBAbrir!id_produto_relacionada & "' and ID_Nota = '" & TBAbrir!ID_nota_relacionada & "'", Conexao, adOpenKeyset, adLockOptimistic
           TBLISTA.Open "Select * from tbl_detalhes_nota where Int_codigo = '" & Var & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBLISTA.EOF = False Then
           Set Resultado = Conexao.Execute("Select sum(qtde) as Total from Faturamento_Relacionamento where ID_Produto_Relacionada = '" & TBLISTA!Int_codigo & "' and ID_Nota_Relacionada = '" & TBLISTA!ID_nota & "'")
            If IIf(IsNull(Resultado!Total), 0, Resultado!Total) >= TBLISTA!int_Qtd Then
            TBLISTA!Saldo = 0
            Else
            TBLISTA!Saldo = TBLISTA!int_Qtd - IIf(IsNull(Resultado!Total), 0, Resultado!Total)
            End If
            TBLISTA.Update
            End If
 '   TBAbrir.MoveNext
 '   PBLista.Value = Contador
 '   Contador = Contador + 1
 '   Loop
 '   TBAbrir.Close
    USMsgBox "Saldos corrigidos com sucesso"
    
End If
Contador = Contador - 1
Loop
ProcFiltrar

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Btnrelacionar_Click()
On Error GoTo tratar_erro

If USMsgBox("Deseja realmente relacionar esse item?", vbYesNo, "CAPRIND v5.0") = vbYes Then
ProcSalvar True
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

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
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
With frmFaturamento_Prod_Serv
    If Opt_saida.Value = True Then TextoFiltroTipo = "and NF.int_tiponota = 1 and NF.ID <> " & .TxtID Else TextoFiltroTipo = "and NF.int_tiponota = 2 and NF.ID <> " & .TxtID & " and NF.Aplicacao = 'T'"
    If optProduto.Value = True Then
        TipoNF = "M1"
    ElseIf OptServico.Value = True Then
            TipoNF = "SA"
        Else
            TipoNF = "M1SA"
    End If
    TipoFiltro = " "
    If optProduto.Value = True Then
        TipoFiltro = " and NFP.Tipo = 'P'"
    ElseIf OptServico.Value = True Then
            TipoFiltro = " and NFP.Tipo = 'S'"
    End If
    If Left(.cmbFinalidade_emissao, 1) = 2 Then
        TextoFiltroTriangulacao = ""
        TextoFiltroCod = ""
        TextoFiltroSaldo = ""
        CamposFiltro = "NF.ID, NF.dt_DataEmissao, NF.int_NotaFiscal, NF.TipoNF, NFNFE.Chave_acesso, NF.txt_Razao_Nome"
    Else
        If Len(.txttipocliente) = 1 Then TextoFiltroTriangulacao = " or Len(NF.txt_tipocliente) = 2" Else TextoFiltroTriangulacao = " or Len(NF.txt_tipocliente) = 1"
        If .txtCod_Produto <> "" Then TextoFiltroCod = " and NFP.int_Cod_Produto = '" & .txtCod_Produto & "'" Else TextoFiltroCod = ""
        If .Opt_saida = False And Faturamento_NF_Saida = True Then TextoFiltroSaldo = "" Else TextoFiltroSaldo = " and NFP.Saldo > 0" 'Se for nota propria de entrada não precisa ver saldo
        CamposFiltro = "NF.ID, NF.dt_DataEmissao, NF.int_NotaFiscal, NF.TipoNF, NF.txt_Razao_Nome,NFNFE.Chave_acesso, NF.txt_Razao_Nome, NFP.Int_codigo, NFP.int_Cod_Produto, NFP.txt_Descricao, NFP.int_Qtd, NFP.dbl_ValorUnitario, NFP.Unidade_com, NFP.Saldo, NFP.codproduto, NF.Id_Int_Cliente"
    End If
    TextoFiltroPadrao = " NF.tiponf = '" & TipoNF & "' and NF.DtValidacao IS NOT NULL and NF.ID_empresa = " & .Cmb_empresa.ItemData(.Cmb_empresa.ListIndex) & " and NF.Int_status = 1 " & TextoFiltroTipo & TextoFiltroCod & TipoFiltro & TextoFiltroSaldo & " and (NF.id_int_cliente = " & .txtIDCliente & " and NF.txt_Razao_Nome = '" & .txt_Razao & "'" & TextoFiltroTriangulacao & ") group by " & CamposFiltro & " order by NF.dt_DataEmissao " & IIf(.Opt_saida = False And Faturamento_NF_Saida = True, "DESC", "") & ", NF.int_NotaFiscal"
End With
INNERJOINTEXTO = "Select " & CamposFiltro & " FROM ((((tbl_Dados_Nota_Fiscal NF LEFT JOIN tbl_detalhes_nota NFP ON NF.ID = NFP.ID_Nota)LEFT JOIN tbl_Dados_Nota_Fiscal_NFe NFNFE ON NF.ID = NFNFE.ID_nota) LEFT JOIN tbl_Totais_Nota TN ON NF.ID = TN.ID_Nota) LEFT JOIN tbl_Detalhes_Recebimento DR ON NF.ID = DR.ID_Nota) LEFT JOIN tbl_proposta_nota PN ON NF.ID = PN.ID_Nota"
Select Case cmbfiltrarpor
    Case "Nota fiscal":
        TextoFiltro = "NF.int_NotaFiscal"
        If txtTexto <> "" Then txtTexto = FunTamanhoTextoZeroEsq(txtTexto, 9)
    Case "Destinatário": TextoFiltro = "NF.txt_Razao_Nome"
    Case "Emitente": TextoFiltro = "NF.txt_Razao_Nome"
    Case "Código interno": TextoFiltro = "NFP.int_cod_produto"
    Case "Código de referência": TextoFiltro = "NFP.N_Referencia"
    Case "Descrição": TextoFiltro = "NFP.txt_descricao"
    Case "Pedido cliente": TextoFiltro = "NFP.pccliente"
    Case "Nosso número": TextoFiltro = "DR.Nosso_numero"
    Case "Pedido interno/Pedido de compra": TextoFiltro = "DR.Nosso_numero"
End Select
If txtTexto <> "" Then
    StrSqlLocProdPadrao = INNERJOINTEXTO & " where " & TextoFiltro & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and " & TextoFiltroPadrao
Else
    StrSqlLocProdPadrao = INNERJOINTEXTO & " where " & TextoFiltroPadrao
End If

Debug.Print StrSqlLocProdPadrao

ProcCarregaLista
'Select NF.ID, NF.dt_DataEmissao, NF.int_NotaFiscal, NF.TipoNF, NF.txt_Razao_Nome,NFNFE.Chave_acesso FROM ((((tbl_Dados_Nota_Fiscal NF LEFT JOIN tbl_detalhes_nota NFP ON NF.ID = NFP.ID_Nota)LEFT JOIN tbl_Dados_Nota_Fiscal_NFe NFNFE ON NF.ID = NFNFE.ID_nota) LEFT JOIN tbl_Totais_Nota TN ON NF.ID = TN.ID_Nota) LEFT JOIN tbl_Detalhes_Recebimento DR ON NF.ID = DR.ID_Nota) LEFT JOIN tbl_proposta_nota PN ON NF.ID = PN.ID_Nota where  NF.tiponf = 'M1' and NF.DtValidacao IS NOT NULL and NF.ID_empresa = 1 and NF.Int_status = 1 and NF.int_tiponota = 1 and NF.ID <> 10276 and NFP.Tipo = 'P' and (NF.id_int_cliente = 109 and NF.txt_Razao_Nome = 'KROMI LOGISTICA DO BRASIL LTDA') group by NF.ID, NF.dt_DataEmissao, NF.int_NotaFiscal, NF.TipoNF, NF.txt_Razao_Nome,NFNFE.Chave_acesso order by NF.dt_DataEmissao , NF.int_NotaFiscal
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub cmdPagAnt_Click()
On Error GoTo tratar_erro

If DS_RetornarNumeros(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
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
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub cmdPagIr_Click()
On Error GoTo tratar_erro

If txtPagIr = "" Then Exit Sub
Quant = DS_RetornarNumeros(Right(lblPaginas.Caption, 4))
If Quant <= 1 Or txtPagIr > Quant Then Exit Sub
If txtPagIr.Text >= 1 And txtPagIr.Text <= Quant Then
    TBLocalizar_produto_padrao.AbsolutePage = txtPagIr.Text
    ProcExibePagina (TBLocalizar_produto_padrao.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub cmdPagPrim_Click()
On Error GoTo tratar_erro

If DS_RetornarNumeros(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLocalizar_produto_padrao.AbsolutePage = 1
ProcExibePagina (TBLocalizar_produto_padrao.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub cmdPagProx_Click()
On Error GoTo tratar_erro

If DS_RetornarNumeros(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
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
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub cmdPagUlt_Click()
On Error GoTo tratar_erro

If DS_RetornarNumeros(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLocalizar_produto_padrao.AbsolutePage = TBLocalizar_produto_padrao.PageCount
ProcExibePagina (TBLocalizar_produto_padrao.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Command1_Click()
On Error GoTo tratar_erro
PBLista.Min = 0
PBLista.Max = TBAbrir.RecordCount
PBLista.Value = 1
Contador = 0

    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from Faturamento_Relacionamento order by ID", Conexao, adOpenKeyset, adLockOptimistic
    Do While TBAbrir.EOF = False
        Set TBLISTA = CreateObject("adodb.recordset")
            TBLISTA.Open "Select * from tbl_detalhes_nota where Int_codigo = '" & TBAbrir!id_produto_relacionada & "' and ID_Nota = '" & TBAbrir!ID_nota_relacionada & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBLISTA.EOF = False Then
           Set Resultado = Conexao.Execute("Select sum(qtde) as Total from Faturamento_Relacionamento where ID_Produto_Relacionada = '" & TBLISTA!Int_codigo & "' and ID_Nota_Relacionada = '" & TBLISTA!ID_nota & "'")
            If IIf(IsNull(Resultado!Total), 0, Resultado!Total) >= TBLISTA!int_Qtd Then
            TBLISTA!Saldo = 0
            Else
            TBLISTA!Saldo = TBLISTA!int_Qtd - IIf(IsNull(Resultado!Total), 0, Resultado!Total)
            End If
            TBLISTA.Update
            End If
    TBAbrir.MoveNext
    PBLista.Value = Contador
    Contador = Contador + 1
    Loop
    TBAbrir.Close
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case SSTab1.Tab
    Case 0:
        Select Case KeyCode
            Case vbKeyF2: ProcFiltrar
            Case vbKeyF3: ProcSalvar True
            Case vbKeyEscape: Unload Me
        End Select
    Case 1:
        Select Case KeyCode
            Case vbKeyF3: If USToolBar2.ButtonState(1) = 0 Then ProcSalvar False
            Case vbKeyF4: ProcExcluir
            Case vbKeyEscape: Unload Me
        End Select
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 14505, 6, True
ProcCarregaToolBar2 Me, 14505, 6, True

SSTab1.Tab = 0
ListView1.CheckBoxes = False
RelacionamentoSimultaneo = False
With frmFaturamento_Prod_Serv
'Finalidade de emissão norma, propria e de entrada

    If Left(.cmbFinalidade_emissao, 1) = 2 Or (Faturamento_NF_Saida = True And .Opt_saida = True) Then
        With ListView1
            .Height = 4080
            .ColumnHeaders(3).Width = 0
            .ColumnHeaders(4).Width = 0
            .ColumnHeaders(5).Width = 0
            .ColumnHeaders(9).Width = 3000
            .ColumnHeaders(10).Width = 0
            .ColumnHeaders(11).Width = 0
            .ColumnHeaders(12).Width = 0
            .ColumnHeaders(13).Width = 0
            .ColumnHeaders(15).Width = 7000
            
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
    If .Opt_saida = True Then
        If Left(.cmbFinalidade_emissao, 2) = 2 Then
            Opt_saida.Value = True
            Opt_entrada.Enabled = False
        Else
            Opt_entrada.Value = True
        End If
        Texto = "Saída"
        Label5.Caption = "Qtde. saída"
        txtQtde.ToolTipText = "Quantidade de saída."
        ListView1.ColumnHeaders(10).Text = "Qtde. entr."
        ListView1.ColumnHeaders(11).Text = "Qtde. saída"
        Label1.Caption = "Qtde. saída"
        Label1.Left = 4635
        txtQtde3.ToolTipText = "Quantidade de saída."
        If .txtIDProduto = 0 Then
            With ListView1
                .CheckBoxes = True
                .ColumnHeaders(1).Width = 300
                .ColumnHeaders(9).Width = 3255
                .Height = 4050
            End With
            PBLista.Top = 7425
            Frame9.Top = 6810
            'Frame2.Visible = False
            RelacionamentoSimultaneo = True
        End If
    Else
'        If Faturamento_NF_Saida = False Then
'            If Left(.cmbFinalidade_emissao, 2) = 2 Then
'                Opt_saida.Enabled = False
'                Opt_entrada.Value = True
'            Else
'                Opt_saida.Value = True
'            End If
'        Else
'            'Opt_saida.Value = True
'            'Opt_entrada.Enabled = False
'        End If
        Texto = "Entrada"
        Label5.Caption = "Qtde. entr."
        txtQtde.ToolTipText = "Quantidade de entrada."
        ListView1.ColumnHeaders(10).Text = "Qtde. saída"
        ListView1.ColumnHeaders(11).Text = "Qtde. entr."
        Label1.Caption = "Qtde. entrada"
        txtQtde.Locked = False
        Label1.Left = 4515
        txtQtde3.ToolTipText = "Quantidade de entrada."
    End If
    If Formulario = "Faturamento/Nota fiscal/Própria" Then
        Caption = "Nota fiscal - Própria - Relacionamento de nota fiscal (Nota fiscal : " & .txtNFiscal & " - " & Texto & ")"
    ElseIf Formulario = "Faturamento/Nota fiscal/Terceiros" Then
        Caption = "Nota fiscal - Terceiros - Relacionamento de nota fiscal (Nota fiscal : " & .txtNFiscal & " - " & Texto & ")"
    ElseIf Formulario = "Estoque/Ordem de faturamento" Then
        Caption = "Ordem de faturamento - Relacionamento de nota fiscal (Ordem : " & .TxtID & " - " & Texto & ")"
    Else
        Caption = "Nota fiscal - Relacionamento de nota fiscal (Nota fiscal : " & .txtNFiscal & " - " & Texto & ")"
    End If
    txtCodinterno = .txtCod_Produto
    txtDescricao = .txtDescricao_Produto
    txtQtde = .txtQtd
End With
cmbfiltrarpor = "Nota fiscal"
ProcFiltrar
ProcCarregaListaRelacionada

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcCarregaListaRelacionada()
On Error GoTo tratar_erro

ListView2.ListItems.Clear
With frmFaturamento_Prod_Serv
    Qtde = IIf(txtQtde = "", 0, txtQtde)
    quantidade = 0
    
    If Left(.cmbFinalidade_emissao, 1) = 2 Or (frmFaturamento_Prod_Serv.Opt_saida = False And Faturamento_NF_Saida = True) Then
        TextoFiltro = "ID_nota = " & .TxtID & " or ID_nota_relacionada = " & .TxtID
    Else
        TextoFiltro = "ID_nota = " & .TxtID & " and ID_produto = " & .txtIDProduto & " or ID_nota_relacionada = " & .TxtID & " and ID_produto_relacionada = " & .txtIDProduto
    End If
    
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from Faturamento_Relacionamento where " & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        PBLista1.Min = 0
        PBLista1.Max = TBAbrir.RecordCount
        PBLista1.Value = 1
        Contador = 0
        Do While TBAbrir.EOF = False
            With ListView2.ListItems
                .Add , , TBAbrir!ID
                
                With frmFaturamento_Prod_Serv
                    If TBAbrir!ID_nota = .TxtID Then
                        If Left(.cmbFinalidade_emissao, 1) = 2 Or TBAbrir!id_produto_relacionada = 0 Or (frmFaturamento_Prod_Serv.Opt_saida = False And Faturamento_NF_Saida = True) Then TextoFiltro = "NF.ID = " & TBAbrir!ID_nota_relacionada Else TextoFiltro = "NFP.Int_codigo = " & TBAbrir!id_produto_relacionada
                    Else
                        If Left(.cmbFinalidade_emissao, 1) = 2 Or TBAbrir!ID_produto = 0 Or (frmFaturamento_Prod_Serv.Opt_saida = False And Faturamento_NF_Saida = True) Then TextoFiltro = "NF.ID = " & TBAbrir!ID_nota Else TextoFiltro = "NFP.Int_codigo = " & TBAbrir!ID_produto
                    End If
                End With
                Set TBFI = CreateObject("adodb.recordset")
                TBFI.Open "Select NF.dt_DataEmissao, NF.int_NotaFiscal, NF.int_TipoNota, NF.aplicacao, NF.txt_Razao_Nome, NFP.dbl_ValorUnitario, NFP.Unidade_com from tbl_Dados_Nota_Fiscal NF LEFT JOIN tbl_Detalhes_Nota NFP ON NFP.ID_nota = NF.ID where " & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
                If TBFI.EOF = False Then
                    .Item(.Count).SubItems(1) = IIf(IsNull(TBFI!dt_DataEmissao), "", (Format(TBFI!dt_DataEmissao, "dd/mm/yy")))
                    .Item(.Count).SubItems(2) = IIf(IsNull(TBFI!int_NotaFiscal), "", TBFI!int_NotaFiscal)
                    If TBFI!int_TipoNota = 1 Then TipoNF2 = "Saída" Else TipoNF2 = "Entrada"
                    If TBFI!Aplicacao = "P" Then TipoNF = "Própria" Else TipoNF = "Terceiros"
                    .Item(.Count).SubItems(3) = TipoNF & "\" & TipoNF2
                    .Item(.Count).SubItems(4) = IIf(IsNull(TBFI!txt_Razao_Nome), "", TBFI!txt_Razao_Nome)
                    
                    If Left(frmFaturamento_Prod_Serv.cmbFinalidade_emissao, 1) <> 2 Then
                        .Item(.Count).SubItems(5) = IIf(IsNull(TBAbrir!Qtde), 0, Format(TBAbrir!Qtde, "###,##0.0000"))
                        .Item(.Count).SubItems(6) = IIf(IsNull(TBFI!dbl_ValorUnitario), 0, Format(TBFI!dbl_ValorUnitario, "###,##0.00000"))
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

If frmFaturamento_Prod_Serv.Opt_saida = True And Faturamento_NF_Saida = True Then ProcCarregaListaDevolucao
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcCarregaListaDevolucao()
On Error GoTo tratar_erro

With frmFaturamento_Prod_Serv
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from Faturamento_Relacionamento where ID_nota = " & .TxtID & " and ID_produto = 0 or ID_nota_relacionada = " & .TxtID & " and ID_produto_relacionada = 0", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        PBLista1.Min = 0
        PBLista1.Max = TBAbrir.RecordCount
        PBLista1.Value = 1
        Contador = 0
        Do While TBAbrir.EOF = False
            With ListView2.ListItems
                .Add , , TBAbrir!ID
                
                With frmFaturamento_Prod_Serv
                    If TBAbrir!ID_nota = .TxtID Then
                        TextoFiltro = "NF.ID = " & TBAbrir!ID_nota_relacionada
                    Else
                        TextoFiltro = "NF.ID = " & TBAbrir!ID_nota
                    End If
                End With
                Set TBFI = CreateObject("adodb.recordset")
                TBFI.Open "Select NF.dt_DataEmissao, NF.int_NotaFiscal, NF.int_TipoNota, NF.Aplicacao, NF.txt_Razao_Nome, NFP.dbl_ValorUnitario, NFP.Unidade_com from tbl_Dados_Nota_Fiscal NF LEFT JOIN tbl_Detalhes_Nota NFP ON NFP.ID_nota = NF.ID where " & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
                If TBFI.EOF = False Then
                    .Item(.Count).SubItems(1) = IIf(IsNull(TBFI!dt_DataEmissao), "", (Format(TBFI!dt_DataEmissao, "dd/mm/yy")))
                    .Item(.Count).SubItems(2) = IIf(IsNull(TBFI!int_NotaFiscal), "", TBFI!int_NotaFiscal)
                    If TBFI!int_TipoNota = 1 Then TipoNF2 = "Saída" Else TipoNF2 = "Entrada"
                    If TBFI!Aplicacao = "P" Then TipoNF = "Própria" Else TipoNF = "Terceiros"
                    .Item(.Count).SubItems(3) = TipoNF & "\" & TipoNF2
                    .Item(.Count).SubItems(4) = IIf(IsNull(TBFI!txt_Razao_Nome), "", TBFI!txt_Razao_Nome)
                    .Item(.Count).SubItems(5) = ""
                    .Item(.Count).SubItems(6) = ""
                End If
                TBFI.Close
            End With
            TBAbrir.MoveNext
            Contador = Contador + 1
            PBLista1.Value = Contador
        Loop
    End If
End With
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
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
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Frame2.Visible = True Then
    With ListView1
        If .ListItems.Count = 0 Then Exit Sub
        txtCodinterno = .SelectedItem.ListSubItems(2)
        txtDescricao = .SelectedItem.ListSubItems(4)
        txtQtde1 = .SelectedItem.ListSubItems(11)
    End With
    Frame2.Enabled = True
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
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
                    If FunVerificaRegistroValidadoSemMsg("tbl_Dados_Nota_Fiscal", "ID = " & frmFaturamento_Prod_Serv.TxtID, True) = False Then
                        .ListItems.Item(InitFor).Checked = False
                        GoTo Proximo
                    End If
                End If
                If .ListItems.Item(InitFor).ListSubItems(5) < 0 And .ListItems.Item(InitFor).ListSubItems(5) <> "" Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
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
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ListView2_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With ListView2
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If frmFaturamento_Prod_Serv.txtNFiscal <> "" Then
                If FunVerificaRegistroValidado("tbl_Dados_Nota_Fiscal", "ID = " & frmFaturamento_Prod_Serv.TxtID, "nota fiscal", "este relacionamento", "excluir", False, True) = False Then
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
                End If
            End If
            
            If .ListItems.Item(InitFor).ListSubItems(5) < 0 And .ListItems.Item(InitFor).ListSubItems(5) <> "" Then
                USMsgBox "Não é permitido excluir este tipo de relacionamento.", vbExclamation, "CAPRIND v5.0"
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
        End If
    Next InitFor
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ListView2_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Left(frmFaturamento_Prod_Serv.cmbFinalidade_emissao, 1) = 2 Or ListView2.ListItems.Count = 0 Then Exit Sub
txt_ID = ListView2.SelectedItem
txtQtde1 = ListView2.SelectedItem.ListSubItems(5)
If ListView2.SelectedItem.ListSubItems(5) = "" Or (frmFaturamento_Prod_Serv.Opt_saida = False And Faturamento_NF_Saida = True) Then
    Frame2.Enabled = False
Else
    Frame2.Enabled = True
    txtQtde1.SetFocus
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcCarregaLista()
On Error GoTo tratar_erro

lblRegistros.Caption = "Nº de reg.: 0"
lblPaginas.Caption = "Página: 0 de: 0"
ListView1.ListItems.Clear
If StrSqlLocProdPadrao = "" Then Exit Sub

Set TBLocalizar_produto_padrao = CreateObject("adodb.recordset")
TBLocalizar_produto_padrao.Open StrSqlLocProdPadrao, Conexao, adOpenKeyset, adLockReadOnly

Debug.Print StrSqlLocProdPadrao

If TBLocalizar_produto_padrao.EOF = False Then ProcExibePagina (1)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
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
    With ListView1.ListItems
        .Add , , TBLocalizar_produto_padrao!ID
        If Left(frmFaturamento_Prod_Serv.cmbFinalidade_emissao, 1) <> 2 Then
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLocalizar_produto_padrao!Int_codigo), "", TBLocalizar_produto_padrao!Int_codigo)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLocalizar_produto_padrao!int_Cod_Produto), "", TBLocalizar_produto_padrao!int_Cod_Produto)
            
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
            
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLocalizar_produto_padrao!txt_Descricao), "", TBLocalizar_produto_padrao!txt_Descricao)
            .Item(.Count).SubItems(9) = IIf(IsNull(TBLocalizar_produto_padrao!int_Qtd), 0, Format(TBLocalizar_produto_padrao!int_Qtd, "###,##0.0000"))
            .Item(.Count).SubItems(10) = Format(IIf(IsNull(TBLocalizar_produto_padrao!int_Qtd), 0, TBLocalizar_produto_padrao!int_Qtd) - IIf(IsNull(TBLocalizar_produto_padrao!Saldo), 0, TBLocalizar_produto_padrao!Saldo), "###,##0.0000")
            .Item(.Count).SubItems(11) = Format(IIf(IsNull(TBLocalizar_produto_padrao!Saldo), 0, TBLocalizar_produto_padrao!Saldo), "###,##0.0000")
            .Item(.Count).SubItems(12) = Format(IIf(IsNull(TBLocalizar_produto_padrao!dbl_ValorUnitario), 0, TBLocalizar_produto_padrao!dbl_ValorUnitario), "###,##0.00000")
            .Item(.Count).SubItems(13) = IIf(IsNull(TBLocalizar_produto_padrao!Unidade_com), "", TBLocalizar_produto_padrao!Unidade_com)
            .Item(.Count).SubItems(14) = IIf(IsNull(TBLocalizar_produto_padrao!Chave_acesso), "", TBLocalizar_produto_padrao!Chave_acesso)
        End If
        .Item(.Count).SubItems(5) = IIf(IsNull(TBLocalizar_produto_padrao!dt_DataEmissao), "", Format(TBLocalizar_produto_padrao!dt_DataEmissao, "dd/mm/yy"))
        .Item(.Count).SubItems(6) = IIf(IsNull(TBLocalizar_produto_padrao!int_NotaFiscal), "", TBLocalizar_produto_padrao!int_NotaFiscal)
        If IsNull(TBLocalizar_produto_padrao!TipoNF) = False Then
            If TBLocalizar_produto_padrao!TipoNF = "M1" Then TipoNF2 = "Produto(s)"
            If TBLocalizar_produto_padrao!TipoNF = "SA" Then TipoNF2 = "Serviço(s)"
            If TBLocalizar_produto_padrao!TipoNF = "M1SA" Then TipoNF2 = "Prod./Serv."
        End If
        .Item(.Count).SubItems(7) = TipoNF2
        .Item(.Count).SubItems(8) = IIf(IsNull(TBLocalizar_produto_padrao!txt_Razao_Nome), "", TBLocalizar_produto_padrao!txt_Razao_Nome)
        .Item(.Count).SubItems(14) = IIf(IsNull(TBLocalizar_produto_padrao!Chave_acesso), "", TBLocalizar_produto_padrao!Chave_acesso)
      
    End With
    TBLocalizar_produto_padrao.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista.Value = Contador
Loop
lblRegistros.Caption = "Nº de reg.: " & ContadorReg
If TBLocalizar_produto_padrao.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "Pág.: 1 de: " & TBLocalizar_produto_padrao.PageCount
ElseIf TBLocalizar_produto_padrao.AbsolutePage = adPosEOF Then
        lblPaginas.Caption = "Pág.: " & TBLocalizar_produto_padrao.PageCount & " de: " & TBLocalizar_produto_padrao.PageCount
    Else
        lblPaginas.Caption = "Pág.: " & TBLocalizar_produto_padrao.AbsolutePage - 1 & " de: " & TBLocalizar_produto_padrao.PageCount
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Opt_entrada_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
ProcCarregaComboFiltrarPor (Opt_entrada.Value)
ProcCorrigeNomeColuna (Opt_entrada.Value)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub opt_Saida_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
ProcCarregaComboFiltrarPor (Opt_entrada.Value)
ProcCorrigeNomeColuna (Opt_entrada.Value)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Optfim_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Optinicio_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Optmeio_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub optProduto_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub optProduto_servico_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub OptServico_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error GoTo tratar_erro

txt_ID = ""
Select Case SSTab1.Tab
    Case 0:
        If ListView1.Visible = True Then ListView1.SetFocus
        If Left(frmFaturamento_Prod_Serv.cmbFinalidade_emissao, 1) <> 2 And Not (frmFaturamento_Prod_Serv.Opt_saida = False And Faturamento_NF_Saida = True) Then
            With Frame2
                If RelacionamentoSimultaneo = True Then .Visible = False Else .Visible = True
                If ListView1.ListItems.Count = 0 Then
                    .Enabled = False
                    txtQtde1 = ""
                Else
                    .Enabled = True
                End If
                .Top = 6840
            End With
            Label7.Caption = "Qtde. à relacionar"
            txtQtde1.ToolTipText = "Quantidade à relacionar."
        End If
        ProcCarregaLista
    Case 1:
        ListView2.SetFocus
        If Left(frmFaturamento_Prod_Serv.cmbFinalidade_emissao, 1) <> 2 And Not (frmFaturamento_Prod_Serv.Opt_saida = False And Faturamento_NF_Saida = True) Then
            With Frame2
                .Visible = True
                .Enabled = False
                .Top = 1290
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
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
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
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub txtQtde1_LostFocus()
On Error GoTo tratar_erro

txtQtde1 = Format(txtQtde1, "###,##0.0000")
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
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
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub txtTexto_LostFocus()
On Error GoTo tratar_erro

If cmbfiltrarpor = "Nota fiscal" And txtTexto <> "" Then txtTexto = FunTamanhoTextoZeroEsq(txtTexto, 9)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub



Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcFiltrar
    Case 2: ProcSalvar True
    'Case 4: ProcAjuda
    Case 5: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub USToolBar2_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcSalvar False
    Case 2: ProcExcluir
    'Case 4: ProcAjuda
    Case 5: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcCarregaComboFiltrarPor(Entrada As Boolean)
On Error GoTo tratar_erro

With cmbfiltrarpor
    .Clear
    .AddItem "Código de referência"
    .AddItem "Código interno"
    .AddItem "Descrição"
    If Entrada = True Then .AddItem "Emitente" Else .AddItem "Destinatário"
    .AddItem "Nosso número"
    .AddItem "Nota fiscal"
    .AddItem "Pedido Cliente"
    .AddItem "Pedido interno/Pedido de compra"
    .Text = "Nota fiscal"
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcCorrigeNomeColuna(Entrada As Boolean)
On Error GoTo tratar_erro

With ListView1.ColumnHeaders
    If Entrada = True Then .Item(9) = "Emitente" Else .Item(9) = "Destinatário"
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Public Sub ProcDevolucao(ID_nota_relacionada As Long)
On Error GoTo tratar_erro

Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select * from Faturamento_Relacionamento where ID_nota <> " & frmFaturamento_Prod_Serv.TxtID & " and ID_nota_relacionada = " & ID_nota_relacionada & " Or ID_nota = " & ID_nota_relacionada & " And ID_nota_relacionada <> " & frmFaturamento_Prod_Serv.TxtID, Conexao, adOpenKeyset, adLockReadOnly
Do While TBProduto.EOF = False
    
    If TBProduto!ID_nota = ID_nota_relacionada Then
        id_produto_entrada = TBProduto!id_produto_relacionada
    Else
        id_produto_entrada = TBProduto!ID_produto
    End If
 
Set TBEstoque = CreateObject("adodb.recordset")
TBEstoque.Open "Select * from Faturamento_Relacionamento", Conexao, adOpenKeyset, adLockOptimistic
TBEstoque.AddNew
TBEstoque!ID_nota = TBProduto!ID_nota
TBEstoque!ID_produto = TBProduto!ID_produto
TBEstoque!ID_nota_relacionada = TBProduto!ID_nota_relacionada
TBEstoque!id_produto_relacionada = TBProduto!id_produto_relacionada
TBEstoque!Qtde = TBProduto!Qtde * -1
TBEstoque.Update
TBEstoque.Close

Set TBEstoque = CreateObject("adodb.recordset")
TBEstoque.Open "Select * from tbl_Detalhes_Nota  WHERE Int_codigo = " & id_produto_entrada, Conexao, adOpenKeyset, adLockOptimistic
If TBEstoque.EOF = False Then
TBEstoque!Saldo = TBEstoque!Saldo + TBProduto!Qtde
TBEstoque.Update
End If
TBEstoque.Close
TBProduto.MoveNext
Loop
TBProduto.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub
