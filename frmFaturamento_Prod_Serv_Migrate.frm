VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2014.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form frmFaturamento_Prod_Serv_Migrate 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Administrativo - Faturamento - Nota fiscal - Dados da NFe"
   ClientHeight    =   10035
   ClientLeft      =   1770
   ClientTop       =   1665
   ClientWidth     =   15360
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
   Icon            =   "frmFaturamento_Prod_Serv_Migrate.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   10035
   ScaleWidth      =   15360
   WindowState     =   2  'Maximizado
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   99
      ScreenHeight    =   768
      ScreenWidth     =   1360
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
   Begin VB.Timer Timer_status_NFe 
      Interval        =   10000
      Left            =   5190
      Top             =   7140
   End
   Begin DrawSuite2014.USProgressBar PBLista 
      Height          =   255
      Left            =   55
      TabIndex        =   55
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
   Begin TabDlg.SSTab SStab_nfe 
      Height          =   10035
      Left            =   0
      TabIndex        =   36
      Top             =   0
      Width           =   15360
      _ExtentX        =   27093
      _ExtentY        =   17701
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
      TabCaption(0)   =   "Dados principais"
      TabPicture(0)   =   "frmFaturamento_Prod_Serv_Migrate.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ListaNota"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "USToolBar1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtID_nota"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtID_entrega"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame9"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Txt_ID_cobranca"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Lista de produtos"
      TabPicture(1)   =   "frmFaturamento_Prod_Serv_Migrate.frx":105E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame_comb_lub"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "txtID_item"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "FrameCST"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "USToolBar2"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "listaProdutos"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      Begin VB.TextBox Txt_ID_cobranca 
         Alignment       =   2  'Centralizar
         Height          =   315
         Left            =   2910
         TabIndex        =   67
         Text            =   "0"
         Top             =   7530
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.Frame Frame_comb_lub 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Dados para combustível e lubrificante"
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
         Height          =   915
         Left            =   -74945
         TabIndex        =   60
         Top             =   2250
         Width           =   15195
         Begin VB.TextBox txtDescANP 
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
            Left            =   2610
            MaxLength       =   60
            TabIndex        =   77
            ToolTipText     =   "Descrição do produto da ANP."
            Top             =   450
            Width           =   9150
         End
         Begin VB.ComboBox Cmb_tipo_produto 
            Appearance      =   0  'Flat
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
            ItemData        =   "frmFaturamento_Prod_Serv_Migrate.frx":107A
            Left            =   12720
            List            =   "frmFaturamento_Prod_Serv_Migrate.frx":1090
            Locked          =   -1  'True
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   26
            TabStop         =   0   'False
            ToolTipText     =   "Tipo do produto."
            Top             =   450
            Width           =   2295
         End
         Begin VB.ComboBox Cmb_codigo_ANP 
            Appearance      =   0  'Flat
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
            ItemData        =   "frmFaturamento_Prod_Serv_Migrate.frx":10F0
            Left            =   180
            List            =   "frmFaturamento_Prod_Serv_Migrate.frx":10F2
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   24
            ToolTipText     =   "Código do produto da ANP."
            Top             =   450
            Width           =   2415
         End
         Begin VB.ComboBox Cmb_UF_consumo 
            Appearance      =   0  'Flat
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
            ItemData        =   "frmFaturamento_Prod_Serv_Migrate.frx":10F4
            Left            =   11790
            List            =   "frmFaturamento_Prod_Serv_Migrate.frx":10F6
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   25
            ToolTipText     =   "UF de consumo."
            Top             =   450
            Width           =   915
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparente
            Caption         =   "Descrição do produto da ANP"
            Height          =   195
            Left            =   6135
            TabIndex        =   78
            Top             =   240
            Width           =   2100
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparente
            Caption         =   "Tipo do produto"
            Height          =   195
            Left            =   13297
            TabIndex        =   63
            Top             =   240
            Width           =   1140
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparente
            Caption         =   "UF cons."
            Height          =   195
            Left            =   11925
            TabIndex        =   62
            Top             =   240
            Width           =   630
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparente
            Caption         =   "Código do produto da ANP"
            Height          =   195
            Left            =   435
            TabIndex        =   61
            Top             =   240
            Width           =   1905
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   60
         TabIndex        =   57
         Top             =   9090
         Width           =   15195
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
            Left            =   2730
            TabIndex        =   28
            Text            =   "30"
            ToolTipText     =   "Número de registros por página."
            Top             =   180
            Width           =   555
         End
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
            ItemData        =   "frmFaturamento_Prod_Serv_Migrate.frx":10F8
            Left            =   6960
            List            =   "frmFaturamento_Prod_Serv_Migrate.frx":1102
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   187
            Width           =   1965
         End
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
            TabIndex        =   30
            ToolTipText     =   "Número da página."
            Top             =   180
            Width           =   555
         End
         Begin DrawSuite2014.USButton cmdPagProx 
            Height          =   315
            Left            =   11760
            TabIndex        =   34
            ToolTipText     =   "Próxima página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmFaturamento_Prod_Serv_Migrate.frx":1123
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
            Left            =   11220
            TabIndex        =   33
            ToolTipText     =   "Página anterior."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmFaturamento_Prod_Serv_Migrate.frx":48C7
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
            Left            =   10110
            TabIndex        =   31
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
            Left            =   10680
            TabIndex        =   32
            ToolTipText     =   "Primeira página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmFaturamento_Prod_Serv_Migrate.frx":83D0
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
            Left            =   12300
            TabIndex        =   35
            ToolTipText     =   "Última página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmFaturamento_Prod_Serv_Migrate.frx":C4BF
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
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparente
            Caption         =   "registros por página"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   3360
            TabIndex        =   74
            Top             =   240
            Width           =   1440
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparente
            Caption         =   "Carregar"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   2040
            TabIndex        =   65
            Top             =   240
            Width           =   645
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparente
            Caption         =   "Operação da lista"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   1
            Left            =   5610
            TabIndex        =   64
            Top             =   240
            Width           =   1260
         End
         Begin VB.Label lblRegistros 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparente
            Caption         =   "Nº de registros: 0"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   180
            TabIndex        =   59
            Top             =   240
            Width           =   1275
         End
         Begin VB.Label lblPaginas 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparente
            Caption         =   "Página: 0 de: 0"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   13050
            TabIndex        =   58
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.TextBox txtID_entrega 
         Alignment       =   2  'Centralizar
         Height          =   315
         Left            =   2550
         TabIndex        =   41
         Text            =   "0"
         Top             =   7530
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.TextBox txtID_nota 
         Alignment       =   2  'Centralizar
         Height          =   315
         Left            =   2130
         TabIndex        =   40
         Text            =   "0"
         Top             =   7530
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.TextBox txtID_item 
         Alignment       =   2  'Centralizar
         Height          =   335
         Left            =   -72870
         TabIndex        =   37
         Text            =   "0"
         ToolTipText     =   "id do produto."
         Top             =   4950
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   2685
         Left            =   55
         TabIndex        =   38
         Top             =   1320
         Width           =   15195
         Begin VB.ComboBox cmbForma_pagamento 
            Appearance      =   0  'Flat
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
            ItemData        =   "frmFaturamento_Prod_Serv_Migrate.frx":FD4B
            Left            =   12990
            List            =   "frmFaturamento_Prod_Serv_Migrate.frx":FD58
            Style           =   2  'Dropdown List
            TabIndex        =   4
            ToolTipText     =   "Indicador da forma de pagamento."
            Top             =   390
            Width           =   2025
         End
         Begin VB.ComboBox cmbFormaPag 
            Appearance      =   0  'Flat
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
            ItemData        =   "frmFaturamento_Prod_Serv_Migrate.frx":FD8C
            Left            =   10980
            List            =   "frmFaturamento_Prod_Serv_Migrate.frx":FDB4
            Style           =   2  'Dropdown List
            TabIndex        =   3
            ToolTipText     =   "Indicador da forma de pagamento."
            Top             =   390
            Width           =   1995
         End
         Begin VB.CheckBox chkCodRef 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Utilizar código de referência na DANFE"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   11850
            TabIndex        =   9
            Top             =   1058
            Width           =   3315
         End
         Begin VB.ComboBox Cmb_presenca_comprador 
            Appearance      =   0  'Flat
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
            ItemData        =   "frmFaturamento_Prod_Serv_Migrate.frx":FEA3
            Left            =   1740
            List            =   "frmFaturamento_Prod_Serv_Migrate.frx":FEBC
            Style           =   2  'Dropdown List
            TabIndex        =   6
            ToolTipText     =   "Indicador de presença do comprador no estabelecimento comercial no momento da operação."
            Top             =   990
            Width           =   4485
         End
         Begin VB.ComboBox Cmb_consumidor 
            Appearance      =   0  'Flat
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
            ItemData        =   "frmFaturamento_Prod_Serv_Migrate.frx":FFCD
            Left            =   180
            List            =   "frmFaturamento_Prod_Serv_Migrate.frx":FFD7
            Style           =   2  'Dropdown List
            TabIndex        =   5
            ToolTipText     =   "Operação com consumidor final."
            Top             =   990
            Width           =   1545
         End
         Begin VB.CheckBox Chk_DA_cobranca 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Imprimir nos dados adicionais"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   12390
            TabIndex        =   15
            Top             =   1980
            Value           =   1  'Marcado
            Width           =   2685
         End
         Begin VB.CheckBox Chk_DA_entrega 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Imprimir nos dados adicionais"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   5040
            TabIndex        =   13
            Top             =   1980
            Value           =   1  'Marcado
            Width           =   2685
         End
         Begin VB.ComboBox cmbEntrega 
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
            ItemData        =   "frmFaturamento_Prod_Serv_Migrate.frx":FFED
            Left            =   180
            List            =   "frmFaturamento_Prod_Serv_Migrate.frx":FFEF
            Style           =   2  'Dropdown List
            TabIndex        =   12
            ToolTipText     =   "Endereço de entrega."
            Top             =   2190
            Width           =   7425
         End
         Begin VB.ComboBox Cmb_cobranca 
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
            ItemData        =   "frmFaturamento_Prod_Serv_Migrate.frx":FFF1
            Left            =   7600
            List            =   "frmFaturamento_Prod_Serv_Migrate.frx":FFF3
            Style           =   2  'Dropdown List
            TabIndex        =   14
            ToolTipText     =   "Endereço de cobrança"
            Top             =   2190
            Width           =   7415
         End
         Begin VB.ComboBox Cmb_enviar_DANFE 
            Appearance      =   0  'Flat
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
            ItemData        =   "frmFaturamento_Prod_Serv_Migrate.frx":FFF5
            Left            =   10560
            List            =   "frmFaturamento_Prod_Serv_Migrate.frx":FFFF
            Style           =   2  'Dropdown List
            TabIndex        =   8
            ToolTipText     =   "Enviar DANFE por e-mail."
            Top             =   990
            Width           =   1185
         End
         Begin VB.TextBox Txt_texto_canhoto_DANFE 
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
            Left            =   7600
            MaxLength       =   60
            TabIndex        =   11
            ToolTipText     =   "Texto do canhoto da DANFE."
            Top             =   1590
            Width           =   7415
         End
         Begin VB.TextBox Txt_titulo_canhoto_DANFE 
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
            MaxLength       =   60
            TabIndex        =   10
            ToolTipText     =   "Título do canhoto da DANFE."
            Top             =   1590
            Width           =   7395
         End
         Begin VB.ComboBox Cmb_arquivos_XML_enviados_email 
            Appearance      =   0  'Flat
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
            ItemData        =   "frmFaturamento_Prod_Serv_Migrate.frx":1000D
            Left            =   6240
            List            =   "frmFaturamento_Prod_Serv_Migrate.frx":10020
            Style           =   2  'Dropdown List
            TabIndex        =   7
            ToolTipText     =   "Arquivos XML enviados por e-mail."
            Top             =   990
            Width           =   4305
         End
         Begin VB.ComboBox Cmb_forma_de_emissao 
            Appearance      =   0  'Flat
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
            ItemData        =   "frmFaturamento_Prod_Serv_Migrate.frx":100A0
            Left            =   180
            List            =   "frmFaturamento_Prod_Serv_Migrate.frx":100B9
            Style           =   2  'Dropdown List
            TabIndex        =   1
            ToolTipText     =   "Finalidade de emissão."
            Top             =   390
            Width           =   8775
         End
         Begin VB.ComboBox cmbFinalidade_emissao 
            Appearance      =   0  'Flat
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
            ItemData        =   "frmFaturamento_Prod_Serv_Migrate.frx":102EA
            Left            =   8970
            List            =   "frmFaturamento_Prod_Serv_Migrate.frx":102FA
            Style           =   2  'Dropdown List
            TabIndex        =   2
            ToolTipText     =   "Finalidade de emissão."
            Top             =   390
            Width           =   1995
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparente
            Caption         =   "Ind. da forma de pagto."
            Height          =   195
            Left            =   13132
            TabIndex        =   76
            Top             =   180
            Width           =   1740
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparente
            Caption         =   "Forma de pagamento"
            Height          =   195
            Left            =   11212
            TabIndex        =   75
            Top             =   180
            Width           =   1530
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparente
            Caption         =   "Presença do comprador no estabelecimento"
            Height          =   195
            Index           =   4
            Left            =   2415
            TabIndex        =   71
            Top             =   780
            Width           =   3135
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparente
            Caption         =   "Consumidor final"
            Height          =   195
            Index           =   3
            Left            =   360
            TabIndex        =   70
            Top             =   780
            Width           =   1185
         End
         Begin VB.Label Label12 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparente
            Caption         =   "Endereço de cobrança"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   8
            Left            =   10500
            TabIndex        =   69
            Top             =   1980
            Width           =   1605
         End
         Begin VB.Label Label12 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparente
            Caption         =   "Endereço de entrega"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   0
            Left            =   3030
            TabIndex        =   68
            Top             =   1980
            Width           =   1515
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparente
            Caption         =   "Enviar DANFE"
            Height          =   195
            Index           =   2
            Left            =   10657
            TabIndex        =   66
            Top             =   780
            Width           =   990
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparente
            Caption         =   "Texto canhoto da DANFE"
            Height          =   195
            Index           =   6
            Left            =   10400
            TabIndex        =   52
            Top             =   1380
            Width           =   1815
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparente
            Caption         =   "Título canhoto da DANFE"
            Height          =   195
            Index           =   5
            Left            =   2985
            TabIndex        =   51
            Top             =   1380
            Width           =   1785
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparente
            Caption         =   "Arquivos XML enviados por e-mail"
            Height          =   195
            Index           =   1
            Left            =   7192
            TabIndex        =   50
            Top             =   780
            Width           =   2400
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparente
            Caption         =   "Forma de emissão"
            Height          =   195
            Left            =   3922
            TabIndex        =   49
            Top             =   180
            Width           =   1290
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparente
            Caption         =   "Finalidade de emissão"
            Height          =   195
            Index           =   0
            Left            =   9172
            TabIndex        =   39
            Top             =   180
            Width           =   1590
         End
      End
      Begin VB.Frame FrameCST 
         BackColor       =   &H00E0E0E0&
         Caption         =   "CST ICMS"
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
         Height          =   915
         Left            =   -74945
         TabIndex        =   42
         Top             =   1330
         Width           =   15195
         Begin VB.ComboBox cmbModalidade_determinacao_ST 
            Appearance      =   0  'Flat
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
            ItemData        =   "frmFaturamento_Prod_Serv_Migrate.frx":1033F
            Left            =   7590
            List            =   "frmFaturamento_Prod_Serv_Migrate.frx":10355
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   23
            ToolTipText     =   "Modalidade de determinação da BC ST."
            Top             =   450
            Width           =   7455
         End
         Begin VB.ComboBox cmbModalidade_determinacao 
            Appearance      =   0  'Flat
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
            ItemData        =   "frmFaturamento_Prod_Serv_Migrate.frx":10404
            Left            =   180
            List            =   "frmFaturamento_Prod_Serv_Migrate.frx":10414
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   22
            ToolTipText     =   "Modalidade de determinação da BC."
            Top             =   450
            Width           =   7395
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparente
            Caption         =   "Modalidade de determinação da BC ST"
            Height          =   195
            Left            =   9945
            TabIndex        =   44
            Top             =   240
            Width           =   2745
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparente
            Caption         =   "Modalidade de determinação da BC"
            Height          =   195
            Left            =   2617
            TabIndex        =   43
            Top             =   240
            Width           =   2520
         End
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
         ForeColor       =   &H00000000&
         Height          =   840
         Left            =   55
         TabIndex        =   45
         Top             =   4020
         Width           =   15195
         Begin VB.ComboBox cmbUF_embarque 
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
            Height          =   330
            ItemData        =   "frmFaturamento_Prod_Serv_Migrate.frx":10482
            Left            =   5280
            List            =   "frmFaturamento_Prod_Serv_Migrate.frx":104DA
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   19
            ToolTipText     =   "UF."
            Top             =   390
            Width           =   690
         End
         Begin VB.TextBox txtLocal_embarque 
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
            Left            =   1800
            MaxLength       =   60
            TabIndex        =   18
            ToolTipText     =   "Local onde ocorrerá o embarque dos produtos."
            Top             =   390
            Width           =   3465
         End
         Begin VB.TextBox txtSerie 
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
            Left            =   1200
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   17
            TabStop         =   0   'False
            ToolTipText     =   "Série."
            Top             =   390
            Width           =   585
         End
         Begin VB.TextBox Txt_chave_acesso 
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
            Left            =   10680
            Locked          =   -1  'True
            MaxLength       =   44
            TabIndex        =   21
            TabStop         =   0   'False
            ToolTipText     =   "Chave de acesso NFe."
            Top             =   390
            Width           =   4335
         End
         Begin VB.TextBox txtNota 
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
            Left            =   150
            Locked          =   -1  'True
            MaxLength       =   60
            TabIndex        =   16
            TabStop         =   0   'False
            ToolTipText     =   "Número da NFe."
            Top             =   390
            Width           =   1035
         End
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
            Height          =   315
            Left            =   5970
            Locked          =   -1  'True
            MaxLength       =   60
            TabIndex        =   20
            TabStop         =   0   'False
            ToolTipText     =   "Status NFe."
            Top             =   390
            Width           =   4695
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparente
            Caption         =   "Local onde ocorrerá o embarque dos produtos"
            Height          =   195
            Index           =   1
            Left            =   1875
            TabIndex        =   73
            Top             =   180
            Width           =   3315
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparente
            Caption         =   "UF"
            Height          =   195
            Left            =   5535
            TabIndex        =   72
            Top             =   180
            Width           =   195
         End
         Begin VB.Label Label12 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparente
            Caption         =   "Série"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   7
            Left            =   1305
            TabIndex        =   53
            Top             =   180
            Width           =   360
         End
         Begin VB.Label Label12 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparente
            Caption         =   "Chave de acesso"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   4
            Left            =   12225
            TabIndex        =   48
            Top             =   180
            Width           =   1230
         End
         Begin VB.Label Label12 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparente
            Caption         =   "Nota fiscal"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   3
            Left            =   292
            TabIndex        =   47
            Top             =   180
            Width           =   750
         End
         Begin VB.Label Label12 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparente
            Caption         =   "Status"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   2
            Left            =   7875
            TabIndex        =   46
            Top             =   180
            Width           =   465
         End
      End
      Begin DrawSuite2014.USToolBar USToolBar1 
         Height          =   975
         Left            =   60
         TabIndex        =   54
         Top             =   330
         Width           =   15195
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
         ButtonLeft2     =   40
         ButtonTop2      =   2
         ButtonWidth2    =   38
         ButtonHeight2   =   21
         ButtonUseMaskColor2=   0   'False
         ButtonCaption3  =   "Liberar"
         ButtonEnabled3  =   0   'False
         ButtonIconSize3 =   32
         ButtonToolTipText3=   "Liberar para envio (F7)"
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
         ButtonLeft3     =   80
         ButtonTop3      =   2
         ButtonWidth3    =   41
         ButtonHeight3   =   21
         ButtonUseMaskColor3=   0   'False
         ButtonCaption4  =   "Cancelar liberação"
         ButtonEnabled4  =   0   'False
         ButtonIconSize4 =   32
         ButtonToolTipText4=   "Cancelar liberação (F4)"
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
         ButtonLeft4     =   123
         ButtonTop4      =   2
         ButtonWidth4    =   96
         ButtonHeight4   =   21
         ButtonUseMaskColor4=   0   'False
         ButtonCaption5  =   "Atualizar status"
         ButtonEnabled5  =   0   'False
         ButtonIconSize5 =   32
         ButtonToolTipText5=   "Atualizar status (F8)"
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
         ButtonLeft5     =   221
         ButtonTop5      =   2
         ButtonWidth5    =   83
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
         ButtonLeft6     =   306
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft7     =   310
         ButtonTop7      =   2
         ButtonWidth7    =   36
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft8     =   348
         ButtonTop8      =   2
         ButtonWidth8    =   26
         ButtonHeight8   =   21
         ButtonUseMaskColor8=   0   'False
         ButtonEnabled9  =   0   'False
         ButtonIconSize9 =   32
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
         ButtonState9    =   5
         ButtonLeft9     =   376
         ButtonTop9      =   2
         ButtonWidth9    =   24
         ButtonHeight9   =   24
         ButtonUseMaskColor9=   0   'False
         Begin DrawSuite2014.USImageList USImageList1 
            Left            =   9390
            Top             =   90
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmFaturamento_Prod_Serv_Migrate.frx":1054C
            Count           =   1
         End
      End
      Begin MSComctlLib.ListView ListaNota 
         Height          =   4200
         Left            =   60
         TabIndex        =   0
         Top             =   4875
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   7408
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
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "N"
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Empresa"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Object.Tag             =   "D"
            Text            =   "Dt. emissão"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Tag             =   "N"
            Text            =   "Nota fiscal"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "Tipo"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Object.Tag             =   "T"
            Text            =   "Série"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Object.Tag             =   "N"
            Text            =   "Valor total"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Object.Tag             =   "T"
            Text            =   "Destinatário"
            Object.Width           =   8123
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Object.Tag             =   "T"
            Text            =   "Status"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Object.Tag             =   "T"
            Text            =   "Status NFe"
            Object.Width           =   2646
         EndProperty
      End
      Begin DrawSuite2014.USToolBar USToolBar2 
         Height          =   975
         Left            =   -74945
         TabIndex        =   56
         Top             =   330
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
         ButtonCaption1  =   "Salvar"
         ButtonEnabled1  =   0   'False
         ButtonIconSize1 =   32
         ButtonToolTipText1=   "Salvar (F3)"
         ButtonKey1      =   "3"
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
         ButtonKey2      =   "4"
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
         ButtonKey4      =   "6"
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
         ButtonKey5      =   "7"
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
         ButtonKey6      =   "8"
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
         ButtonState6    =   5
         ButtonLeft6     =   153
         ButtonTop6      =   2
         ButtonWidth6    =   24
         ButtonHeight6   =   24
         ButtonUseMaskColor6=   0   'False
         Begin DrawSuite2014.USImageList USImageList2 
            Left            =   13980
            Top             =   210
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmFaturamento_Prod_Serv_Migrate.frx":14C3F
            Count           =   1
         End
      End
      Begin MSComctlLib.ListView listaProdutos 
         Height          =   6530
         Left            =   -74940
         TabIndex        =   27
         Top             =   3180
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   11509
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
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
         NumItems        =   17
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "N"
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Tag             =   "T"
            Text            =   "Cod. interno"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Descrição"
            Object.Width           =   7408
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Object.Tag             =   "N"
            Text            =   "CST de ICMS"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Object.Tag             =   "N"
            Text            =   "CST de IPI"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Object.Tag             =   "N"
            Text            =   "CST de PIS"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Object.Tag             =   "N"
            Text            =   "CST de Cofins"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   7
            Object.Tag             =   "T"
            Text            =   "NCM"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   8
            Object.Tag             =   "T"
            Text            =   "Un."
            Object.Width           =   970
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   9
            Object.Tag             =   "N"
            Text            =   "Qtde."
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   10
            Object.Tag             =   "N"
            Text            =   "Vlr.unit."
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   11
            Object.Tag             =   "N"
            Text            =   "Vlr. total"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   12
            Object.Tag             =   "N"
            Text            =   "ICMS"
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   13
            Object.Tag             =   "N"
            Text            =   "IPI"
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   14
            Object.Tag             =   "N"
            Text            =   "Vlr. IPI"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Object.Tag             =   "N"
            Text            =   "Ordem"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Object.Tag             =   "T"
            Text            =   "Pedido do cliente"
            Object.Width           =   2646
         EndProperty
      End
   End
End
Attribute VB_Name = "frmFaturamento_Prod_Serv_Migrate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Pais As String 'OK
Dim Codigo_pais As Long 'OK
Dim Email As String 'OK
Dim TBLISTA_Faturamento_NFe As ADODB.Recordset 'OK

Sub ProcCarregaListaProdutos()
On Error GoTo tratar_erro

listaProdutos.ListItems.Clear
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select NFP.*, CF.IDIntClasse from tbl_Detalhes_Nota NFP LEFT JOIN tbl_ClassificacaoFiscal CF ON CF.Idclass = NFP.ID_CF where NFP.id_nota = " & txtID_nota.Text & " order by NFP.int_codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBProduto.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBProduto.EOF = False
        With listaProdutos.ListItems
            .Add , , TBProduto!Int_codigo
            .Item(.Count).SubItems(1) = IIf(IsNull(TBProduto!int_Cod_Produto), "", TBProduto!int_Cod_Produto)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBProduto!Txt_descricao), "", TBProduto!Txt_descricao)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBProduto!txt_CST), "", TBProduto!txt_CST)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBProduto!CST_IPI), "", TBProduto!CST_IPI)
            .Item(.Count).SubItems(5) = IIf(IsNull(TBProduto!CST_PIS), "", TBProduto!CST_PIS)
            .Item(.Count).SubItems(6) = IIf(IsNull(TBProduto!CST_Cofins), "", TBProduto!CST_Cofins)
            .Item(.Count).SubItems(7) = IIf(IsNull(TBProduto!IDIntClasse), "", TBProduto!IDIntClasse)
            .Item(.Count).SubItems(8) = IIf(IsNull(TBProduto!txt_Unid), "", TBProduto!txt_Unid)
            .Item(.Count).SubItems(9) = IIf(IsNull(TBProduto!int_Qtd), "", Format(TBProduto!int_Qtd, "###,##0.0000"))
            .Item(.Count).SubItems(10) = IIf(IsNull(TBProduto!dbl_ValorUnitario), "", Format(TBProduto!dbl_ValorUnitario, "###,##0.0000000000"))
            If IsNull(TBProduto!dbl_ValorUnitario) = False Then
                .Item(.Count).SubItems(11) = Format(TBProduto!dbl_ValorUnitario * TBProduto!int_Qtd, "###,##0.00")
            End If
            .Item(.Count).SubItems(12) = IIf(IsNull(TBProduto!int_ICMS), "", TBProduto!int_ICMS)
            .Item(.Count).SubItems(13) = IIf(IsNull(TBProduto!int_IPI), "", TBProduto!int_IPI)
            .Item(.Count).SubItems(14) = IIf(IsNull(TBProduto!dbl_valoripi), "", Format(TBProduto!dbl_valoripi, "###,##0.00"))
            .Item(.Count).SubItems(15) = IIf(IsNull(TBProduto!Ordem), "", TBProduto!Ordem)
            .Item(.Count).SubItems(16) = IIf(IsNull(TBProduto!PCCliente), "", TBProduto!PCCliente)
            TBProduto.MoveNext
            Contador = Contador + 1
            PBLista.Value = Contador
        End With
    Loop
End If
TBProduto.Close

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Sub ProcCarregaListaNota(Pagina As Integer)
On Error GoTo tratar_erro

lblRegistros.Caption = "Nº de registros: 0"
lblPaginas.Caption = "Página: 0 de: 0"
ListaNota.ListItems.Clear
With frmFaturamento_Prod_Serv
    If .Strsql_FaturamentoNFe = "" Then Exit Sub
    Set TBLISTA_Faturamento_NFe = CreateObject("adodb.recordset")
    TBLISTA_Faturamento_NFe.Open .Strsql_FaturamentoNFe, Conexao, adOpenKeyset, adLockReadOnly
    If TBLISTA_Faturamento_NFe.EOF = False Then ProcExibePagina (Pagina)
End With

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Private Sub ProcExibePagina(Pagina)
On Error GoTo tratar_erro

ListaNota.ListItems.Clear
TBLISTA_Faturamento_NFe.PageSize = IIf(txtNreg = "", 30, txtNreg)
TBLISTA_Faturamento_NFe.AbsolutePage = Pagina
TamanhoPagina = TBLISTA_Faturamento_NFe.PageSize
ContadorReg = 1

PBLista.Min = 0
PBLista.Max = FunVerifMaxPBListaPaginacao(TBLISTA_Faturamento_NFe.RecordCount - IIf(Pagina > 1, (TBLISTA_Faturamento_NFe.PageSize * (Pagina - 1)), 0), TBLISTA_Faturamento_NFe.PageSize)
PBLista.Value = 1
Contador = 0
Do While TBLISTA_Faturamento_NFe.EOF = False And (ContadorReg <= TamanhoPagina)
    With ListaNota.ListItems
        .Add , , TBLISTA_Faturamento_NFe!ID
        
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from Empresa where Codigo = " & IIf(IsNull(TBLISTA_Faturamento_NFe!ID_empresa), 0, TBLISTA_Faturamento_NFe!ID_empresa), Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            .Item(.Count).SubItems(1) = IIf(IsNull(TBAbrir!Empresa), "", TBAbrir!Empresa)
        End If
        TBAbrir.Close
        
        .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA_Faturamento_NFe!dt_DataEmissao), "", (Format(TBLISTA_Faturamento_NFe!dt_DataEmissao, "dd/mm/yy")))
        .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA_Faturamento_NFe!int_NotaFiscal), "", TBLISTA_Faturamento_NFe!int_NotaFiscal)
        If IsNull(TBLISTA_Faturamento_NFe!TipoNF) = False Then
            If TBLISTA_Faturamento_NFe!TipoNF = "M1" Then TipoNF2 = "Produto(s)"
            If TBLISTA_Faturamento_NFe!TipoNF = "SA" Then TipoNF2 = "Serviço(s)"
            If TBLISTA_Faturamento_NFe!TipoNF = "M1SA" Then TipoNF2 = "Prod./Serv."
        End If
        .Item(.Count).SubItems(4) = TipoNF2
        .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA_Faturamento_NFe!Serie), "", TBLISTA_Faturamento_NFe!Serie)
        .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA_Faturamento_NFe!dbl_Valor_Total_Nota), "0,00", Format(TBLISTA_Faturamento_NFe!dbl_Valor_Total_Nota, "###,##0.00"))
        .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA_Faturamento_NFe!txt_Razao_Nome), "", TBLISTA_Faturamento_NFe!txt_Razao_Nome)
        .Item(.Count).SubItems(8) = IIf(TBLISTA_Faturamento_NFe!Int_status = 1, "Ativa", "Cancelada")
        .Item(.Count).SubItems(9) = FunVerifStatusNFe(TBLISTA_Faturamento_NFe!ID)
    End With
    TBLISTA_Faturamento_NFe.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista.Value = Contador
Loop
lblRegistros.Caption = "Nº de registros: " & TBLISTA_Faturamento_NFe.RecordCount
If TBLISTA_Faturamento_NFe.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "Página: 1 de: " & TBLISTA_Faturamento_NFe.PageCount
ElseIf TBLISTA_Faturamento_NFe.AbsolutePage = adPosEOF Then
        lblPaginas.Caption = "Página: " & TBLISTA_Faturamento_NFe.PageCount & " de: " & TBLISTA_Faturamento_NFe.PageCount
    Else
        lblPaginas.Caption = "Página: " & TBLISTA_Faturamento_NFe.AbsolutePage - 1 & " de: " & TBLISTA_Faturamento_NFe.PageCount
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Private Sub Cmb_codigo_ANP_Click()
On Error GoTo tratar_erro

If Cmb_codigo_ANP = "" Then Exit Sub
Set TBCodigoDesc = CreateObject("adodb.recordset")
TBCodigoDesc.Open "Select Descricao from Codigos_produtos_ANP WHERE Descricao IS NOT NULL AND codigo = " & Cmb_codigo_ANP, Conexao, adOpenKeyset, adLockReadOnly
If TBCodigoDesc.EOF = False Then txtDescANP = TBCodigoDesc!Descricao
TBCodigoDesc.Close
    
Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Private Sub Cmb_opcao_lista_Click()
On Error GoTo tratar_erro

With ListaNota
    For InitFor = 1 To .ListItems.Count
        .ListItems.Item(InitFor).Checked = False
    Next InitFor
End With

With USToolBar1
    If Cmb_opcao_lista = "Cancelar liberação" Then
        .ButtonState(3) = 5
        .ButtonState(4) = 0
    Else
        .ButtonState(3) = 0
        .ButtonState(4) = 5
    End If
    .Refresh
End With

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Private Sub cmbEntrega_Click()
On Error GoTo tratar_erro
  
If cmbEntrega <> "" Then txtID_entrega = cmbEntrega.ItemData(cmbEntrega.ListIndex) Else txtID_entrega = 0

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Private Sub Cmb_cobranca_Click()
On Error GoTo tratar_erro
  
If Cmb_cobranca <> "" Then Txt_ID_cobranca = Cmb_cobranca.ItemData(Cmb_cobranca.ListIndex) Else Txt_ID_cobranca = 0

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Private Sub ProcAtualizarStatus()
On Error GoTo tratar_erro

If Alterar = False Then
    MsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation
    Exit Sub
End If

qtde_solicitada = ""
If MsgBox("Deseja realmente atualizar o status da(s) nota(s) fiscal(ais)?", vbQuestion + vbYesNo) = vbYes Then
Mensagem1:
    qtde_solicitada = InputBox("Favor informar o número de dias para atualização.")
    If qtde_solicitada = "" Then Exit Sub
    If IsNumeric(qtde_solicitada) = False Then
        MsgBox ("Só é permitido número neste campo."), vbExclamation
        GoTo Mensagem1
    End If
    Qtde = qtde_solicitada
    If Qtde <= 0 Then
        MsgBox ("So é permitido número maior que 0."), vbExclamation
        GoTo Mensagem1
    End If
    data = Date - Qtde
    
    Set TBComponente = CreateObject("adodb.recordset")
    TBComponente.Open "Select NF.ID, NF.int_NotaFiscal, NF.ID_empresa, NF.Serie, NF.Int_status, NFE.Status, NFE.Chave_acesso from tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_Dados_Nota_Fiscal_NFe NFE on NF.ID = NFE.ID_nota where NF.dt_DataEmissao >= '" & Format(data, "Short Date") & "' and NF.Aplicacao = 'P' and NF.int_NotaFiscal IS NOT NULL and NF.TipoNF = 'M1' order by NF.int_NotaFiscal, NF.Serie", Conexao, adOpenKeyset, adLockOptimistic
    'TBComponente.Open "Select NF.ID, NF.int_NotaFiscal, NF.ID_empresa, NF.Serie, NF.Int_status, NFE.Status, NFE.Chave_acesso from tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_Dados_Nota_Fiscal_NFe NFE on NF.ID = NFE.ID_nota where NF.ID = 94", Conexao, adOpenKeyset, adLockOptimistic
    If TBComponente.EOF = False Then
        PBLista.Min = 0
        PBLista.Max = TBComponente.RecordCount
        PBLista.Value = 1
        Contador = 0
        Do While TBComponente.EOF = False
            OF = TBComponente!int_NotaFiscal
            Set TBGravar_NFe_Status = CreateObject("adodb.recordset")
            TBGravar_NFe_Status.Open "Select * from NFE012 where CbdEmpCodigo = " & TBComponente!ID_empresa & " and CbdNtfNumero = " & OF & " and CbdNtfSerie = '" & TBComponente!Serie & "' and CbdSituacao <> 0 order by CbdNtfNumero, CbdNtfSerie", Conexao_NFe, adOpenKeyset, adLockOptimistic
            If TBGravar_NFe_Status.EOF = False Then
                If IsNull(TBGravar_NFe_Status!CbdNFEChaAcesso) = False And TBGravar_NFe_Status!CbdNFEChaAcesso <> "" Then EnviadoTexto = "Imprimir = 'True'" Else EnviadoTexto = "Imprimir = 'False'"
                Conexao.Execute "Update tbl_Dados_Nota_Fiscal Set " & EnviadoTexto & " where ID = " & TBComponente!ID
                
                If IsNull(TBGravar_NFe_Status!CbdStsRetCodigo) = False And TBGravar_NFe_Status!CbdStsRetCodigo <> "" And (TBGravar_NFe_Status!CbdProcStatus = "P" Or TBComponente!Int_status = 2 And TBGravar_NFe_Status!CbdProcStatus = "N") Then
                    TBComponente!status = TBGravar_NFe_Status!CbdStsRetCodigo
                Else
                    If TBComponente!Int_status = 2 Then TBComponente!status = -1
                End If
                If IsNull(TBGravar_NFe_Status!CbdSituacao) = False And TBGravar_NFe_Status!CbdSituacao <> 0 Then
                    Set TBAbrir = CreateObject("adodb.recordset")
                    TBAbrir.Open "Select * from Empresa where codigo = " & TBComponente!ID_empresa, Conexao, adOpenKeyset, adLockOptimistic
                    If TBAbrir.EOF = False Then
                        caminho = IIf(IsNull(TBAbrir!Caminho_Nfe), "", TBAbrir!Caminho_Nfe)
                    End If
                    TBAbrir.Close
                    
                    OF = TBComponente!int_NotaFiscal
                    status = TBGravar_NFe_Status!cbdAcao
                    Contador2 = 2
                    Do While Contador2 > 0
                        Set GerArqPastas = CreateObject("Scripting.FileSystemObject")
                        If GerArqPastas.FileExists(caminho & "\Empresa " & TBComponente!ID_empresa & " - Serie " & TBComponente!Serie & " - Nota " & OF & " - Status " & status & ".bat") = True Then Kill (caminho & "\Empresa " & TBComponente!ID_empresa & " - Serie " & TBComponente!Serie & " - Nota " & OF & " - Status " & status & ".bat")
                        If status = "C" Then status = "E" Else status = "C"
                        Contador2 = Contador2 - 1
                    Loop
                End If
                TBComponente!Chave_acesso = IIf(IsNull(TBGravar_NFe_Status!CbdNFEChaAcesso), "", TBGravar_NFe_Status!CbdNFEChaAcesso)
                TBComponente.Update
            End If
            TBGravar_NFe_Status.Close
            
            TBComponente.MoveNext
            Contador = Contador + 1
            PBLista.Value = Contador
        Loop
    End If
    TBComponente.Close
    MsgBox ("Status das(s) nota(s) fiscal(ais) atualizado(s) com sucesso."), vbInformation
    '==================================
    Modulo = Formulario
    Evento = "Atualizar status da(s) nota(s) fiscal(ais)"
    ID_documento = 0
    Documento = ""
    Documento1 = ""
    ProcGravaEvento
    '==================================
    ProcCarregaListaNota (IIf(DS_RetornarNumeros(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, DS_RetornarNumeros(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Private Sub ProcLiberar()
On Error GoTo tratar_erro

If Alterar = False Then
    MsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso.")
    Exit Sub
End If
Permitido = False
With ListaNota
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If MsgBox("Deseja realmente liberar esta(s) nota(s) fiscal(ais)?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
            End If
            
            Permitido = True
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select * from tbl_Dados_Nota_Fiscal WHERE ID = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                ProcExcluirDadosTabelaGNFe IIf(IsNull(TBAbrir!int_NotaFiscal), 0, TBAbrir!int_NotaFiscal), TBAbrir!Serie
                If frmFaturamento_Prod_Serv.txtid = TBAbrir!ID Then frmFaturamento_Prod_Serv.NFe_liberada = True
                
                '==================================
                Modulo = Formulario
                Evento = "Liberar nota fiscal para envio"
                ID_documento = .ListItems(InitFor)
                If IsNull(TBAbrir!int_NotaFiscal) = True Or TBAbrir!int_NotaFiscal = "" Then NomeCampo = "N° ordem: " & TBAbrir!ID Else NomeCampo = "N° nota: " & TBAbrir!int_NotaFiscal
                Documento = NomeCampo & " - Tipo: " & TBAbrir!TipoNF & " - Série: " & TBAbrir!Serie
                Documento1 = ""
                ProcGravaEvento
                '==================================
                
                Email = ""
                'Verifica email e país
                Set TBClientes = CreateObject("adodb.recordset")
                If TBAbrir!txt_tipocliente = "E" Then
                    'Empresa
                    TBClientes.Open "Select Email, Pais, Codigo_pais from Empresa where Codigo = " & TBAbrir!Id_Int_Cliente, Conexao, adOpenKeyset, adLockOptimistic
                    If TBClientes.EOF = False Then
                        Email = IIf(IsNull(TBClientes!Email), "", TBClientes!Email)
                        Pais = TBClientes!Pais
                        Codigo_pais = TBClientes!Codigo_pais
                    End If
                ElseIf TBAbrir!txt_tipocliente = "JP" Or TBAbrir!txt_tipocliente = "JR" Or TBAbrir!txt_tipocliente = "FP" Or TBAbrir!txt_tipocliente = "FR" Then
                        'Cliente
                        TBClientes.Open "Select Email, Pais, Codigo_pais from Clientes where IDcliente = " & TBAbrir!Id_Int_Cliente & " and NomeRazao = '" & TBAbrir!txt_Razao_Nome & "' and Enviar_NF = 'True'", Conexao, adOpenKeyset, adLockOptimistic
                        If TBClientes.EOF = False Then
                            Email = IIf(IsNull(TBClientes!Email), "", TBClientes!Email)
                            If Email <> "" Then TextoFiltro = " and Email <> '" & Email & "'" Else TextoFiltro = ""
                            
                            Set TBFI = CreateObject("adodb.recordset")
                            TBFI.Open "Select Email from Clientes_Contatos where IDcliente = " & TBAbrir!Id_Int_Cliente & TextoFiltro & " and Enviar_NFe = 'True' and EMail is not null", Conexao, adOpenKeyset, adLockOptimistic
                            If TBFI.EOF = False Then
                                Do While TBFI.EOF = False
                                    If IsNull(TBFI!Email) = False And TBFI!Email <> "" Then
                                        If Email <> "" Then Email = Email & ";" & TBFI!Email Else Email = TBFI!Email
                                    End If
                                    TBFI.MoveNext
                                Loop
                            End If
                            TBFI.Close
                                    
                            Pais = TBClientes!Pais
                            Codigo_pais = TBClientes!Codigo_pais
                        End If
                    Else
                        'Fornecedor
                        TBClientes.Open "Select Email, Pais, Codigo_pais from Compras_fornecedores where IDcliente = " & TBAbrir!Id_Int_Cliente & " and Nome_Razao = '" & TBAbrir!txt_Razao_Nome & "' and Enviar_NF = 'True'", Conexao, adOpenKeyset, adLockOptimistic
                        If TBClientes.EOF = False Then
                            Email = IIf(IsNull(TBClientes!Email), "", TBClientes!Email)
                            If Email <> "" Then TextoFiltro = " and Email <> '" & Email & "'" Else TextoFiltro = ""
                            
                            Set TBFI = CreateObject("adodb.recordset")
                            TBFI.Open "Select Email from Contatos_fornecedor where IdFornecedor = " & TBAbrir!Id_Int_Cliente & TextoFiltro & " and Enviar_NFe = 'True' and Email is not null", Conexao, adOpenKeyset, adLockOptimistic
                            If TBFI.EOF = False Then
                                Do While TBFI.EOF = False
                                    If IsNull(TBFI!Email) = False And TBFI!Email <> "" Then
                                        If Email <> "" Then Email = Email & ";" & TBFI!Email Else Email = TBFI!Email
                                    End If
                                    TBFI.MoveNext
                                Loop
                            End If
                            TBFI.Close
                    
                            Pais = TBClientes!Pais
                            Codigo_pais = TBClientes!Codigo_pais
                        End If
                End If
                TBClientes.Close
                
                'Verifica se tem transportadora na NF para consultar o e-mail
                Email1 = ""
                Set TBAfericao = CreateObject("adodb.recordset")
                TBAfericao.Open "Select CNPJ from Empresa where Codigo = " & TBAbrir!ID_empresa, Conexao, adOpenKeyset, adLockOptimistic
                If TBAfericao.EOF = False Then
                    Set TBFIltro = CreateObject("adodb.recordset")
                    TBFIltro.Open "Select IdIntTransp, txt_Razao from tbl_Dados_Transp where ID_nota = " & TBAbrir!ID & " and TXT_CNPJ <> '" & TBAfericao!CNPJ & "' and TXT_CNPJ <> '" & TBAbrir!txt_CNPJ_CPF & "'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBFIltro.EOF = False Then
                        'Cliente
                        Set TBClientes = CreateObject("adodb.recordset")
                        TBClientes.Open "Select Email from Clientes where IDcliente = " & TBFIltro!IdIntTransp & " and NomeRazao = '" & TBFIltro!txt_Razao & "'", Conexao, adOpenKeyset, adLockOptimistic
                        If TBClientes.EOF = False Then
                            Email1 = IIf(IsNull(TBClientes!Email), "", TBClientes!Email)
                            If Email1 <> "" Then
                                TextoFiltro = " and Email <> '" & Email1 & "'"
                                If Email <> "" Then Email = Email & ";" & Email1 Else Email = Email1
                            Else
                                TextoFiltro = ""
                            End If
                            
                            Set TBFI = CreateObject("adodb.recordset")
                            TBFI.Open "Select Email from Clientes_Contatos where IDcliente = " & TBFIltro!IdIntTransp & TextoFiltro & " and Enviar_NFe = 'True' and EMail is not null", Conexao, adOpenKeyset, adLockOptimistic
                            If TBFI.EOF = False Then
                                Do While TBFI.EOF = False
                                    If IsNull(TBFI!Email) = False And TBFI!Email <> "" Then
                                        If Email <> "" Then Email = Email & ";" & TBFI!Email Else Email = TBFI!Email
                                    End If
                                    TBFI.MoveNext
                                Loop
                            End If
                            TBFI.Close
                        Else
                            'Fornecedor
                            Set TBClientes = CreateObject("adodb.recordset")
                            TBClientes.Open "Select Email from Compras_fornecedores where IDcliente = " & TBFIltro!IdIntTransp & " and Nome_Razao = '" & TBFIltro!txt_Razao & "'", Conexao, adOpenKeyset, adLockOptimistic
                            If TBClientes.EOF = False Then
                                Email1 = IIf(IsNull(TBClientes!Email), "", TBClientes!Email)
                                If Email1 <> "" Then
                                    TextoFiltro = " and Email <> '" & Email1 & "'"
                                    If Email <> "" Then Email = Email & ";" & Email1 Else Email = Email1
                                Else
                                    TextoFiltro = ""
                                End If
                                
                                Set TBFI = CreateObject("adodb.recordset")
                                TBFI.Open "Select Email from Contatos_fornecedor where IdFornecedor = " & TBFIltro!IdIntTransp & TextoFiltro & " and Enviar_NFe = 'True' and Email is not null", Conexao, adOpenKeyset, adLockOptimistic
                                If TBFI.EOF = False Then
                                    Do While TBFI.EOF = False
                                        If IsNull(TBFI!Email) = False And TBFI!Email <> "" Then
                                            If Email <> "" Then Email = Email & ";" & TBFI!Email Else Email = TBFI!Email
                                        End If
                                        TBFI.MoveNext
                                    Loop
                                End If
                                TBFI.Close
                            End If
                        End If
                    End If
                End If
            End If
            TBAbrir.Close
        
            'Gravar dados na tabela do MIGRATE GNFE
            Permitido1 = True
            Set TBMaquinas = CreateObject("adodb.recordset")
            TBMaquinas.Open "Select * from Empresa where Empresa = '" & .ListItems(InitFor).ListSubItems(1) & "' and GNFe = 'True'", Conexao, adOpenKeyset, adLockOptimistic
            If TBMaquinas.EOF = False Then
                ProcGerarNFeMigrate .ListItems(InitFor)
                Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open "Select * from tbl_Dados_Nota_Fiscal where ID = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    OF = TBAbrir!int_NotaFiscal
                    Set TBGravar_NFe = CreateObject("adodb.recordset")
                    TBGravar_NFe.Open "Select * from NFE012 where CbdNtfNumero = " & OF & " and CbdNtfSerie = '" & TBAbrir!Serie & "' and CbdAcao <> 'V'", Conexao_NFe, adOpenKeyset, adLockOptimistic
                    If TBGravar_NFe.EOF = True Then Permitido1 = False
                    TBGravar_NFe.Close
                End If
                TBAbrir.Close
            End If
            TBMaquinas.Close
            If Permitido1 = True Then
                Set TBGravar = CreateObject("adodb.recordset")
                TBGravar.Open "Select * from tbl_dados_nota_fiscal_nfe where ID_nota = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
                If TBGravar.EOF = False Then
                    TBGravar!status = 0
                    txtStatus = "Liberado emissão"
                    TBGravar.Update
                End If
                TBGravar.Close
                
                Set TBMaquinas = CreateObject("adodb.recordset")
                TBMaquinas.Open "Select * from Empresa where Empresa = '" & .ListItems(InitFor).ListSubItems(1) & "' and GNFe = 'True'", Conexao, adOpenKeyset, adLockOptimistic
                If TBMaquinas.EOF = False Then
                    Set TBAbrir = CreateObject("adodb.recordset")
                    TBAbrir.Open "Select * from tbl_Dados_Nota_Fiscal where ID = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
                    If TBAbrir.EOF = False Then
                        OF = TBAbrir!int_NotaFiscal
                        Call ProcCriarBATNFeCCe(TBMaquinas!Caminho_Nfe, TBMaquinas!CODIGO, TBAbrir!Serie, OF, "E")
                    End If
                    TBAbrir.Close
                End If
                TBMaquinas.Close
            End If
        End If
    Next InitFor
End With
If Permitido = False Then
    MsgBox ("Informe a(s) nota(s) fiscal(ais) antes de liberar para envio."), vbExclamation
Else
    MsgBox ("Nota(s) fiscal(ais) liberada(s) para envio com sucesso."), vbInformation
    ProcCarregaListaNota (IIf(DS_RetornarNumeros(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, DS_RetornarNumeros(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
    With frmFaturamento_Prod_Serv
        .ProcCarregaListaNota (IIf(DS_RetornarNumeros(Left(.lblPaginas(1).Caption, Len(.lblPaginas(1).Caption) - 5)) <= 1, 1, DS_RetornarNumeros(Left(.lblPaginas(1).Caption, Len(.lblPaginas(1).Caption) - 5))))
    End With
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Private Sub ProcGerarNFeMigrate(ID_nota As Long)
On Error GoTo tratar_erro

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from tbl_Dados_Nota_Fiscal where ID = " & ID_nota, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
       
    OF = TBAbrir!int_NotaFiscal
    Set TBFI = CreateObject("adodb.recordset")
    TBFI.Open "Select * from tbl_Dados_Nota_Fiscal_NFe where ID_nota = " & ID_nota, Conexao, adOpenKeyset, adLockOptimistic
    If TBFI.EOF = False Then
        Set TBGravar_NFe = CreateObject("adodb.recordset")
        TBGravar_NFe.Open "Select * from CBD001", Conexao_NFe, adOpenKeyset, adLockOptimistic
        TBGravar_NFe.AddNew
        TBGravar_NFe!CbdEmpCodigo = TBMaquinas!CODIGO
        TBGravar_NFe!CbdNtfNumero = OF
        TBGravar_NFe!CbdNtfSerie = TBAbrir!Serie
        
        'Dados para impressão automática
        'TBGravar_NFe!CdbUsuImpPadrao = "ImpressoraDANFE"
        'TBGravar_NFe!CdbUsuImpCont = "ImpressoraDANFE"
        
        FamiliaAntiga = DS_RemoverAcentos(TBMaquinas!Cidade)
        TBGravar_NFe!CbdcUF = FunVerificaCodUF(FamiliaAntiga, TBMaquinas!UF)
        
        TBGravar_NFe!CbdcNF = FunTamanhoTextoZeroEsq(OF, 8)
        
        Set TBCFOP = CreateObject("adodb.recordset")
        TBCFOP.Open "Select CFOP.Txt_descricao from tbl_Detalhes_Nota NFP INNER JOIN tbl_NaturezaOperacao CFOP ON CFOP.IDCountCfop = NFP.ID_CFOP where NFP.ID_nota = " & TBAbrir!ID & " order by NFP.Int_codigo", Conexao, adOpenKeyset, adLockOptimistic
        If TBCFOP.EOF = False Then
            TBGravar_NFe!CbdnatOp = DS_RemoverAcentos(TBCFOP!Txt_descricao)
        End If
        
        TBGravar_NFe!Cbdmod = 55
        
        'Novo layout da Sefaz (3.10)
        TBGravar_NFe!CbddEmi = Format(TBAbrir!dt_DataEmissao, "yyyy-mm-dd") & " " & Left(TBAbrir!Hora_emissao, 8)
        TBGravar_NFe!CbdFusoHorario = FunVerifFusoHorario(False)
        
        If IsNull(TBAbrir!dt_Saida_Entrada) = False And TBAbrir!dt_Saida_Entrada <> "" Then TBGravar_NFe!CbddSaiEnt = Format(TBAbrir!dt_Saida_Entrada, "yyyy-mm-dd")
        If IsNull(TBAbrir!txt_Hora_Saida) = False And TBAbrir!txt_Hora_Saida <> "" Then TBGravar_NFe!CbdhrSaiEnt = Format(TBAbrir!txt_Hora_Saida, "hh:mm:ss")
        If TBAbrir!int_TipoNota = 1 Then TBGravar_NFe!CbdtpNf = 1 Else TBGravar_NFe!CbdtpNf = 0 '1 = Saída 2 = Entrada
        
        'Novo layout da Sefaz (3.10)
        Set TBFIltro = CreateObject("adodb.recordset")
        TBFIltro.Open "Select UF from Empresa where Codigo = " & TBAbrir!ID_empresa, Conexao, adOpenKeyset, adLockOptimistic
        If TBFIltro.EOF = False Then
            UFREM = TBFIltro!UF
            If TBFIltro!UF = TBAbrir!txt_UF Then
                TBGravar_NFe!CbdIdDest = 1
            ElseIf TBAbrir!txt_UF = "EX" Then
                    TBGravar_NFe!CbdIdDest = 3
                Else
                    TBGravar_NFe!CbdIdDest = 2
            End If
        End If
        TBFIltro.Close
        
        FamiliaAntiga = DS_RemoverAcentos(TBMaquinas!Cidade)
        TBGravar_NFe!CbdcMunFg = FunVerificaCodMunicipio(FamiliaAntiga, TBMaquinas!UF)
        
        TBGravar_NFe!CbdtpImp = 1 'DANFE 1 = Retrato - 2 = Paisagem no manual
        TBGravar_NFe!CbdtpEmis = TBFI!Forma_emissao
        TBGravar_NFe!CbdfinNFe = TBFI!Finalidade_emissao
        
        'Novo layout da Sefaz (3.10)
        TBGravar_NFe!CbdindFinal = TBFI!Consumidor_final
        TBGravar_NFe!CbdIndPres = TBFI!Presenca_comprador
        
        TBGravar_NFe!CbdCNPJ_emit = DS_RetornarNumeros(TBMaquinas!CNPJ)
        TBGravar_NFe!CbdxNome = DS_RemoverAcentos(Left(TBMaquinas!Razao, 60))
        TBGravar_NFe!CbdxFant = DS_RemoverAcentos(TBMaquinas!Empresa)
        
        FamiliaAntiga = ""
        If IsNull(TBMaquinas!Tipo_endereco) = False And TBMaquinas!Tipo_endereco <> "" Then FamiliaAntiga = TBMaquinas!Tipo_endereco & ": "
        If FamiliaAntiga <> "" Then FamiliaAntiga = FamiliaAntiga & TBMaquinas!Endereco Else FamiliaAntiga = TBMaquinas!Endereco
        TBGravar_NFe!CbdxLgr = DS_RemoverAcentos(FamiliaAntiga)
        
        TBGravar_NFe!Cbdnro = TBMaquinas!Numero
        
        FamiliaAntiga = ""
        If IsNull(TBMaquinas!Tipo_bairro) = False And TBMaquinas!Tipo_bairro <> "" Then FamiliaAntiga = TBMaquinas!Tipo_bairro & ": "
        If FamiliaAntiga <> "" Then FamiliaAntiga = FamiliaAntiga & TBMaquinas!Bairro Else Bairro = TBMaquinas!Bairro
        TBGravar_NFe!CbdxBairro = DS_RemoverAcentos(FamiliaAntiga)
        
        FamiliaAntiga = DS_RemoverAcentos(TBMaquinas!Cidade)
        TBGravar_NFe!CbdcMun = FunVerificaCodMunicipio(FamiliaAntiga, TBMaquinas!UF)
        
        TBGravar_NFe!CbdxMun = DS_RemoverAcentos(TBMaquinas!Cidade)
        TBGravar_NFe!CbdUF = TBMaquinas!UF
        TBGravar_NFe!CbdCEP = DS_RetornarNumeros(TBMaquinas!CEP)
        TBGravar_NFe!CbdcPais = "1058"
        TBGravar_NFe!CbdxPais = "BRASIL"
        If IsNull(TBMaquinas!Telefone) = False And TBMaquinas!Telefone <> "" Then TBGravar_NFe!Cbdfone = DS_RetornarNumeros(TBMaquinas!Telefone)
        If IsNull(TBMaquinas!ie) = False And TBMaquinas!ie <> "" Then TBGravar_NFe!CbdIE = Left(DS_RetornarNumeros(TBMaquinas!ie), 14)
        If IsNull(TBMaquinas!IM) = False And TBMaquinas!IM <> "" Then
            TBGravar_NFe!CbdIM = DS_RetornarNumeros(TBMaquinas!IM)
            If IsNull(TBMaquinas!CNAE) = False And TBMaquinas!CNAE <> "" Then TBGravar_NFe!CbdCNAE = DS_RetornarNumeros(TBMaquinas!CNAE)
        End If
        TBGravar_NFe!CbdxNome_dest = DS_RemoverAcentos(TBAbrir!txt_Razao_Nome)
        TBGravar_NFe!CbdxLgr_dest = DS_RemoverAcentos(TBAbrir!txt_Endereco)
        If Email <> "" Then TBGravar_NFe!CbdxEmail_dest = DS_RemoverAcentos(Left(Email, 160))
        TBGravar_NFe!Cbdnro_dest = TBAbrir!Numero
        
        Set TBFornecedor = CreateObject("adodb.recordset")
        TBFornecedor.Open "Select Complemento, RG_IM, Nao_contribuinte_ICMS, Pessoa from Compras_fornecedores where IDCliente = " & TBAbrir!Id_Int_Cliente & " and Nome_Razao = '" & TBAbrir!txt_Razao_Nome & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBFornecedor.EOF = False Then
            If (TBFornecedor!Nao_contribuinte_ICMS) = True Then Nao_contribuinte_ICMS = "Sim" Else Nao_contribuinte_ICMS = "Não"
            If IsNull(TBFornecedor!Complemento) = False And TBFornecedor!Complemento <> "" Then TBGravar_NFe!CbdxCpl_dest = TBFornecedor!Complemento
            If TBFornecedor!Pessoa = "JURÍDICA" And IsNull(TBFornecedor!RG_IM) = False And TBFornecedor!RG_IM <> "" Then TBGravar_NFe!CbdIM_dest = TBFornecedor!RG_IM 'Novo layout da Sefaz (3.10)
        Else
            Set TBFornecedor = CreateObject("adodb.recordset")
            TBFornecedor.Open "Select Complemento, RG_IM, Nao_contribuinte_ICMS, Tipo from Clientes where IDCliente = " & TBAbrir!Id_Int_Cliente & " and NomeRazao = '" & TBAbrir!txt_Razao_Nome & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBFornecedor.EOF = False Then
                If (TBFornecedor!Nao_contribuinte_ICMS) = True Then Nao_contribuinte_ICMS = "Sim" Else Nao_contribuinte_ICMS = "Não"
                If IsNull(TBFornecedor!Complemento) = False And TBFornecedor!Complemento <> "" Then TBGravar_NFe!CbdxCpl_dest = TBFornecedor!Complemento
                If Left(TBFornecedor!Tipo, 1) = "J" And IsNull(TBFornecedor!RG_IM) = False And TBFornecedor!RG_IM <> "" Then TBGravar_NFe!CbdIM_dest = TBFornecedor!RG_IM 'Novo layout da Sefaz (3.10)
            Else
                Set TBFornecedor = CreateObject("adodb.recordset")
                TBFornecedor.Open "Select Complemento, IM from Empresa where Codigo = " & TBAbrir!Id_Int_Cliente, Conexao, adOpenKeyset, adLockOptimistic
                If TBFornecedor.EOF = False Then
                    Nao_contribuinte_ICMS = "Não"
                    If IsNull(TBFornecedor!Complemento) = False And TBFornecedor!Complemento <> "" Then TBGravar_NFe!CbdxCpl_dest = TBFornecedor!Complemento
                    If IsNull(TBFornecedor!IM) = False And TBFornecedor!IM <> "" Then TBGravar_NFe!CbdIM_dest = TBFornecedor!IM 'Novo layout da Sefaz (3.10)
                End If
            End If
        End If
        TBFornecedor.Close
        
        TBGravar_NFe!CbdxBairro_dest = DS_RemoverAcentos(TBAbrir!Txt_bairro)
        If IsNull(TBAbrir!txt_UF) = True Or TBAbrir!txt_UF = "" Or TBAbrir!txt_UF = "EX" Then
            TBGravar_NFe!CbdcMun_dest = "9999999"
            TBGravar_NFe!CbdxMun_dest = "EXTERIOR"
            TBGravar_NFe!CbdUF_dest = "EX"
            TBGravar_NFe!CbdIdEstrangeiro = Null 'Novo layout da Sefaz (3.10)
        Else
            If TBAbrir!txt_tipocliente = "E" Or Left(TBAbrir!txt_tipocliente, 1) = "J" Then
                TBGravar_NFe!CbdCNPJ_dest = DS_RetornarNumeros(TBAbrir!txt_CNPJ_CPF)
            Else
                TBGravar_NFe!CbdCPF_dest = DS_RetornarNumeros(TBAbrir!txt_CNPJ_CPF)
            End If
            
            FamiliaAntiga = DS_RemoverAcentos(TBAbrir!txt_Municipio)
            TBGravar_NFe!CbdcMun_dest = FunVerificaCodMunicipio(FamiliaAntiga, TBAbrir!txt_UF)
            
            TBGravar_NFe!CbdxMun_dest = DS_RemoverAcentos(TBAbrir!txt_Municipio)
            TBGravar_NFe!CbdUF_dest = TBAbrir!txt_UF
        End If
        TBGravar_NFe!CbdCEP_dest = Left(DS_RetornarNumeros(TBAbrir!Txt_CEP), 8)
        TBGravar_NFe!CbdcPais_dest = Codigo_pais
        TBGravar_NFe!CbdxPais_dest = Pais
        If IsNull(TBAbrir!txt_Fone_Fax) = False And TBAbrir!txt_Fone_Fax <> "" Then
            Fone = DS_RetornarNumeros(TBAbrir!txt_Fone_Fax)
            If Fone <> "" Then TBGravar_NFe!Cbdfone_dest = Right(Fone, 10)
        End If
        
        If Nao_contribuinte_ICMS = "Sim" Or IsNull(TBAbrir!txt_UF) = True Or TBAbrir!txt_UF = "" Or TBAbrir!txt_UF = "EX" Or IsNull(TBAbrir!txt_IE_Cliente) = True Or TBAbrir!txt_IE_Cliente = "" Then
            TBGravar_NFe!CbdIndIEDest = 9 'Novo layout da Sefaz (3.10)
        Else
            If IsNull(TBAbrir!txt_IE_Cliente) = False And TBAbrir!txt_IE_Cliente <> "" Then
                If TBAbrir!txt_IE_Cliente = "ISENTO" Or TBAbrir!txt_IE_Cliente = "Isento" Or TBAbrir!txt_IE_Cliente = "ISENTA" Or TBAbrir!txt_IE_Cliente = "Isenta" Then
                    TBGravar_NFe!CbdIndIEDest = 2 'Novo layout da Sefaz (3.10)
                Else
                    TBGravar_NFe!CbdIndIEDest = 1 'Novo layout da Sefaz (3.10)
                    TBGravar_NFe!CbdIE_dest = Left(DS_RetornarNumeros(TBAbrir!txt_IE_Cliente), 14)
                End If
            End If
        End If
        
        Set TBClientes = CreateObject("adodb.recordset")
        TBClientes.Open "Select CFOP.* from tbl_Detalhes_Nota NFP INNER JOIN tbl_NaturezaOperacao CFOP ON CFOP.IDCountCfop = NFP.ID_CFOP where NFP.ID_nota = " & TBAbrir!ID & " and CFOP.Suframa = 'True'", Conexao, adOpenKeyset, adLockOptimistic
        If TBClientes.EOF = False Then
            Set TBClientes = CreateObject("adodb.recordset")
            TBClientes.Open "Select * from Clientes where IDCliente = " & TBAbrir!Id_Int_Cliente & " and Suframa is not null", Conexao, adOpenKeyset, adLockOptimistic
            If TBClientes.EOF = False Then
                If TBClientes!Suframa <> "" Then TBGravar_NFe!CbdISUF = Left(DS_RetornarNumeros(TBClientes!Suframa), 9)
            End If
        End If
        TBClientes.Close
        
        If Email <> "" Then TBGravar_NFe!CbdEmail_dest = DS_RemoverAcentos(Left(Email, 60))
        
        Set TBTotaisnota = CreateObject("adodb.recordset")
        TBTotaisnota.Open "Select * from tbl_Totais_Nota where ID_nota = " & ID_nota, Conexao, adOpenKeyset, adLockOptimistic
        If TBTotaisnota.EOF = False Then
            TBGravar_NFe!CbdvBC_ttlnfe = IIf(IsNull(TBTotaisnota!dbl_Base_ICMS), 0, TBTotaisnota!dbl_Base_ICMS)
            TBGravar_NFe!CbdvICMS_ttlnfe = IIf(IsNull(TBTotaisnota!dbl_Valor_ICMS), 0, TBTotaisnota!dbl_Valor_ICMS)
            
            'TBGravar_NFe!CbdvICMSDeson_ttlnfe = IIf(IsNull(TBTotaisnota!Valor_total_ICMS_desonerado), 0, TBTotaisnota!Valor_total_ICMS_desonerado) 'Novo layout da Sefaz (3.10) - Não é obrigatório
            
            TBGravar_NFe!CbdvBCST_ttlnfe = IIf(IsNull(TBTotaisnota!dbl_Base_ICMS_Subst), 0, TBTotaisnota!dbl_Base_ICMS_Subst)
            TBGravar_NFe!CbdvST_ttlnfe = IIf(IsNull(TBTotaisnota!dbl_Valor_ICMS_Subst), 0, TBTotaisnota!dbl_Valor_ICMS_Subst)
            TBGravar_NFe!CbdvProd_ttlnfe = IIf(IsNull(TBTotaisnota!dbl_Valor_Total_Produtos), 0, TBTotaisnota!dbl_Valor_Total_Produtos) + IIf(IsNull(TBTotaisnota!dbl_Valor_Total_Nota_Serv), 0, TBTotaisnota!dbl_Valor_Total_Nota_Serv)
            TBGravar_NFe!CbdvFrete_ttlnfe = IIf(IsNull(TBTotaisnota!dbl_Valor_Frete), 0, TBTotaisnota!dbl_Valor_Frete)
            TBGravar_NFe!CbdvSeg_ttlnfe = IIf(IsNull(TBTotaisnota!dbl_Valor_Seguro), 0, TBTotaisnota!dbl_Valor_Seguro)
            TBGravar_NFe!CbdvDesc_ttlnfe = IIf(IsNull(TBTotaisnota!Valor_total_desconto), 0, TBTotaisnota!Valor_total_desconto) + IIf(IsNull(TBTotaisnota!Valor_total_desconto_SUFRAMA), 0, TBTotaisnota!Valor_total_desconto_SUFRAMA)
            TBGravar_NFe!CbdvII_ttlnfe = IIf(IsNull(TBTotaisnota!Valor_total_II), 0, TBTotaisnota!Valor_total_II)
            TBGravar_NFe!CbdvIPI_ttlnfe = IIf(IsNull(TBTotaisnota!dbl_Valor_Total_IPI), 0, TBTotaisnota!dbl_Valor_Total_IPI)
            If TBFI!Finalidade_emissao = 4 And TBMaquinas!Simples = True Then TBGravar_NFe!CbdvIPIDevol_tttlnfe = IIf(IsNull(TBTotaisnota!dbl_Valor_Total_IPI), 0, TBTotaisnota!dbl_Valor_Total_IPI) 'Novo layout da Sefaz (4.0)
            
            TBGravar_NFe!CbdvPIS_ttlnfe = IIf(IsNull(TBTotaisnota!Total_PIS_prod), 0, TBTotaisnota!Total_PIS_prod)
            TBGravar_NFe!CbdvCOFINS_ttlnfe = IIf(IsNull(TBTotaisnota!Total_Cofins_prod), 0, TBTotaisnota!Total_Cofins_prod)
            TBGravar_NFe!CbdvOutro = IIf(IsNull(TBTotaisnota!dbl_Desp_Adicionais), 0, TBTotaisnota!dbl_Desp_Adicionais)
            TBGravar_NFe!CbdvNF = IIf(IsNull(TBTotaisnota!dbl_Valor_Total_Nota), 0, TBTotaisnota!dbl_Valor_Total_Nota)
            If IsNull(TBTotaisnota!Valor_total_aprox_tributos) = False And TBTotaisnota!Valor_total_aprox_tributos <> "" And TBTotaisnota!Valor_total_aprox_tributos <> "0" Then TBGravar_NFe!CBDvTotTrib_ttlnfe = TBTotaisnota!Valor_total_aprox_tributos
            TBGravar_NFe!CbdvBC_ttlnfe_iss = IIf(IsNull(TBTotaisnota!dbl_Valor_Total_Nota_Serv), 0, TBTotaisnota!dbl_Valor_Total_Nota_Serv)
            TBGravar_NFe!CbdvISS = IIf(IsNull(TBTotaisnota!dbl_valor_total_iss), 0, TBTotaisnota!dbl_valor_total_iss)
            TBGravar_NFe!CbdvPIS_servttlnfe = IIf(IsNull(TBTotaisnota!Total_PIS_serv), 0, TBTotaisnota!Total_PIS_serv)
            TBGravar_NFe!CbdvCOFINS_servttlnfe = IIf(IsNull(TBTotaisnota!Total_Cofins_serv), 0, TBTotaisnota!Total_Cofins_serv)
            TBGravar_NFe!CbdvRetPIS = IIf(IsNull(TBTotaisnota!Total_retencao_PIS), 0, TBTotaisnota!Total_retencao_PIS)
            TBGravar_NFe!CbdvRetCOFINS_servttlnfe = IIf(IsNull(TBTotaisnota!Total_retencao_Cofins), 0, TBTotaisnota!Total_retencao_Cofins)
            TBGravar_NFe!CbdvRetCSLL = IIf(IsNull(TBTotaisnota!Total_CSLL_serv), 0, TBTotaisnota!Total_CSLL_serv)
            TBGravar_NFe!CbdvBCIRRF = IIf(IsNull(TBTotaisnota!dbl_Valor_Total_Nota_Serv), 0, TBTotaisnota!dbl_Valor_Total_Nota_Serv)
            TBGravar_NFe!CbdvIRRF = IIf(IsNull(TBTotaisnota!Total_IRRF_serv), 0, TBTotaisnota!Total_IRRF_serv)
            TBGravar_NFe!CbdvFCPUFDest_ttlnfe = IIf(IsNull(TBTotaisnota!Valor_total_ICMS_FCP), 0, TBTotaisnota!Valor_total_ICMS_FCP)
            'TBGravar_NFe!CbdvFCP_ttlnfe = IIf(IsNull(TBTotaisnota!Valor_total_ICMS_FCP), 0, TBTotaisnota!Valor_total_ICMS_FCP) 'Novo layout da Sefaz (4.0)
            TBGravar_NFe!CbdvICMSUFDest_ttlnfe = IIf(IsNull(TBTotaisnota!Valor_total_ICMS_INT_UF_dest), 0, TBTotaisnota!Valor_total_ICMS_INT_UF_dest)
            TBGravar_NFe!CbdvICMSUFRemet_ttlnfe = IIf(IsNull(TBTotaisnota!Valor_total_ICMS_INT_UF_rem), 0, TBTotaisnota!Valor_total_ICMS_INT_UF_rem)

            'Forma de pagamento Novo layout da Sefaz (4.0)
            Set TBContas = CreateObject("adodb.recordset")
            TBContas.Open "Select * from tbl_Detalhes_Recebimento where ID_nota = " & ID_nota & " order by ID", Conexao, adOpenKeyset, adLockOptimistic
            If TBContas.EOF = False Then
                TBGravar_NFe!CbdnFat = TBAbrir!int_NotaFiscal
                TBGravar_NFe!CbdvOrig = IIf(IsNull(TBTotaisnota!Valor_total_receber_pagar), 0, TBTotaisnota!Valor_total_receber_pagar) + IIf(IsNull(TBTotaisnota!Valor_total_desconto), 0, TBTotaisnota!Valor_total_desconto)
                TBGravar_NFe!CbdvDesc_cob = IIf(IsNull(TBTotaisnota!Valor_total_desconto), 0, TBTotaisnota!Valor_total_desconto)
                TBGravar_NFe!CbdvLiq = IIf(IsNull(TBTotaisnota!Valor_total_receber_pagar), 0, TBTotaisnota!Valor_total_receber_pagar)
            End If
            TBContas.Close
            
            'Forma de pagamento Novo layout da Sefaz (4.0)
            Set TBGravar_NFe1 = CreateObject("adodb.recordset")
            TBGravar_NFe1.Open "Select * from CBD001PAG", Conexao_NFe, adOpenKeyset, adLockOptimistic
            TBGravar_NFe1.AddNew
            TBGravar_NFe1!CbdEmpCodigo = TBMaquinas!CODIGO
            TBGravar_NFe1!CbdNtfSerie = TBAbrir!Serie
            TBGravar_NFe1!CbdNtfNumero = OF
            TBGravar_NFe1!CbdPagseq = 1
            TBGravar_NFe1!CbdtPag = IIf(IsNull(TBFI!FormaPagto), "15", TBFI!FormaPagto)
            TBGravar_NFe1!CbdindPag_pag = TBFI!Forma_pagamento
            If TBFI!FormaPagto <> 90 Then TBGravar_NFe1!CbdvPag = IIf(IsNull(TBTotaisnota!dbl_Valor_Total_Nota), 0, TBTotaisnota!dbl_Valor_Total_Nota)
            TBGravar_NFe1.Update
            TBGravar_NFe1.Close
            
            DAPartilhaICMS = ""
            If TBGravar_NFe!CbdvFCPUFDest_ttlnfe <> 0 Then DAPartilhaICMS = "Partilha ICMS operação interestadual consumidor final, disposto na Emenda constitucional 87/2015. Valor ICMS para UF destino (" & TBAbrir!txt_UF & "): R$" & Format(TBGravar_NFe!CbdvICMSUFDest_ttlnfe, "###,##0.00") & ". Valor FCP para o destino: R$" & Format(TBGravar_NFe!CbdvFCPUFDest_ttlnfe, "###,##0.00") & ". Valor ICMS UF remetente (" & UFREM & "): R$" & Format(TBGravar_NFe!CbdvICMSUFRemet_ttlnfe, "###,##0.00") & "."
        End If
        TBTotaisnota.Close
        Set TBTransporte = CreateObject("adodb.recordset")
        TBTransporte.Open "Select * from tbl_Dados_Transp where ID_Nota = " & ID_nota, Conexao, adOpenKeyset, adLockOptimistic
        If TBTransporte.EOF = False Then
            TBGravar_NFe!CbdmodFrete = TBTransporte!txt_Frete_Conta 'Frete Novo layout da Sefaz (4.0)
            
            Familiatext = ""
            If IsNull(TBTransporte!txt_CNPJ) = False And TBTransporte!txt_CNPJ <> "" Then
                Set TBFornecedor = CreateObject("adodb.recordset")
                TBFornecedor.Open "Select * from Compras_fornecedores where IDCliente = " & TBTransporte!IdIntTransp & " and Nome_Razao = '" & TBTransporte!txt_Razao & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBFornecedor.EOF = False Then
                    If Left(TBFornecedor!Pessoa, 1) = "J" Then TBGravar_NFe!CbdCNPJ_transp = DS_RetornarNumeros(TBTransporte!txt_CNPJ) Else TBGravar_NFe!CbdCPF_transp = DS_RetornarNumeros(TBTransporte!txt_CNPJ)
                    If IsNull(TBTransporte!txt_Endereco) = False And TBTransporte!txt_Endereco <> "" Then Familiatext = TBTransporte!txt_Endereco
                    If IsNull(TBTransporte!int_numero) = False And TBTransporte!int_numero <> "" Then
                        If Familiatext <> "" Then Familiatext = Familiatext & ", " & TBTransporte!int_numero Else Familiatext = TBTransporte!int_numero
                    End If
                    If IsNull(TBFornecedor!Bairro) = False And TBFornecedor!Bairro <> "" Then
                        If Familiatext <> "" Then Familiatext = Familiatext & " - " & TBFornecedor!Bairro Else Familiatext = TBFornecedor!Bairro
                    End If
                    'If IsNull(TBTransporte!txt_Municipio) = False And TBTransporte!txt_Municipio <> "" Then
                        'If Familiatext <> "" Then Familiatext = Familiatext & " - " & TBTransporte!txt_Municipio Else Familiatext = TBTransporte!txt_Municipio
                    'End If
                    'If IsNull(TBTransporte!txt_UF) = False And TBTransporte!txt_UF <> "" Then
                        'If Familiatext <> "" Then Familiatext = Familiatext & " - " & TBTransporte!txt_UF Else Familiatext = TBTransporte!txt_UF
                    'End If
                Else
                    Set TBFornecedor = CreateObject("adodb.recordset")
                    TBFornecedor.Open "Select * from Clientes where IDCliente = " & TBTransporte!IdIntTransp & " and NomeRazao = '" & TBTransporte!txt_Razao & "'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBFornecedor.EOF = False Then
                        If Left(TBFornecedor!Tipo, 1) = "J" Then TBGravar_NFe!CbdCNPJ_transp = DS_RetornarNumeros(TBTransporte!txt_CNPJ) Else TBGravar_NFe!CbdCPF_transp = DS_RetornarNumeros(TBTransporte!txt_CNPJ)
                        If IsNull(TBTransporte!txt_Endereco) = False And TBTransporte!txt_Endereco <> "" Then Familiatext = TBTransporte!txt_Endereco
                        If IsNull(TBTransporte!int_numero) = False And TBTransporte!int_numero <> "" Then
                            If Familiatext <> "" Then Familiatext = Familiatext & ", " & TBTransporte!int_numero Else Familiatext = TBTransporte!int_numero
                        End If
                        If IsNull(TBFornecedor!Bairro) = False And TBFornecedor!Bairro <> "" Then
                            If Familiatext <> "" Then Familiatext = Familiatext & " - " & TBFornecedor!Bairro Else Familiatext = TBFornecedor!Bairro
                        End If
'                        If IsNull(TBTransporte!txt_Municipio) = False And TBTransporte!txt_Municipio <> "" Then
'                            If Familiatext <> "" Then Familiatext = Familiatext & " - " & TBTransporte!txt_Municipio Else Familiatext = TBTransporte!txt_Municipio
'                        End If
'                        If IsNull(TBTransporte!txt_UF) = False And TBTransporte!txt_UF <> "" Then
'                            If Familiatext <> "" Then Familiatext = Familiatext & " - " & TBTransporte!txt_UF Else Familiatext = TBTransporte!txt_UF
'                        End If
                    Else
                        Set TBFornecedor = CreateObject("adodb.recordset")
                        TBFornecedor.Open "Select * from Empresa where Codigo = " & TBTransporte!IdIntTransp, Conexao, adOpenKeyset, adLockOptimistic
                        If TBFornecedor.EOF = False Then
                            TBGravar_NFe!CbdCNPJ_transp = DS_RetornarNumeros(TBTransporte!txt_CNPJ)
                            If IsNull(TBTransporte!txt_Endereco) = False And TBTransporte!txt_Endereco <> "" Then Familiatext = TBTransporte!txt_Endereco
                            If IsNull(TBTransporte!int_numero) = False And TBTransporte!int_numero <> "" Then
                                If Familiatext <> "" Then Familiatext = Familiatext & ", " & TBTransporte!int_numero Else Familiatext = TBTransporte!int_numero
                            End If
                            If IsNull(TBFornecedor!Bairro) = False And TBFornecedor!Bairro <> "" Then
                                If Familiatext <> "" Then Familiatext = Familiatext & " - " & TBFornecedor!Bairro Else Familiatext = TBFornecedor!Bairro
                            End If
'                            If IsNull(TBTransporte!txt_Municipio) = False And TBTransporte!txt_Municipio <> "" Then
'                                If Familiatext <> "" Then Familiatext = Familiatext & " - " & TBTransporte!txt_Municipio Else Familiatext = TBTransporte!txt_Municipio
'                            End If
'                            If IsNull(TBTransporte!txt_UF) = False And TBTransporte!txt_UF <> "" Then
'                                If Familiatext <> "" Then Familiatext = Familiatext & " - " & TBTransporte!txt_UF Else Familiatext = TBTransporte!txt_UF
'                            End If
                        End If
                    End If
                End If
                TBFornecedor.Close
            End If
            TBGravar_NFe!CbdxNome_transp = DS_RemoverAcentos(TBTransporte!txt_Razao)
            If IsNull(TBTransporte!txt_IE) = False And TBTransporte!txt_IE <> "" Then
                TBGravar_NFe!CbdIE_transp = Left(DS_RetornarNumeros(TBTransporte!txt_IE), 14)
                TBGravar_NFe!CbdUF_transp = TBTransporte!txt_UF
            End If
            If Familiatext <> "" Then
                Familiatext = Left(Familiatext, 60)
                TBGravar_NFe!CbdxEnder = DS_RemoverAcentos(Familiatext)
            End If
            TBGravar_NFe!CbdxMun_transp = DS_RemoverAcentos(TBTransporte!txt_Municipio)
                            
            'FamiliaAntiga = DS_RemoverAcentos(TBTransporte!txt_Municipio)
            'TBGravar_NFe!CbdcMunFG_transp = FunVerificaCodMunicipio(FamiliaAntiga, TBTransporte!txt_uf)
            
            'If Len(TBAbrir!int_CFOP) > 5 Then
                'If Left(TBAbrir!int_CFOP, 5) <> "5.902" And Left(TBAbrir!int_CFOP, 5) <> "6.902" Then TBGravar_NFe!CbdCFOP_transp = DS_RetornarNumeros(Left(TBAbrir!int_CFOP, 5))
                'If Right(TBAbrir!int_CFOP, 5) <> "5.902" And Right(TBAbrir!int_CFOP, 5) <> "6.902" Then TBGravar_NFe!CbdCFOP_transp = DS_RetornarNumeros(Right(TBAbrir!int_CFOP, 5))
            'Else
                'TBGravar_NFe!CbdCFOP_transp = DS_RetornarNumeros(TBAbrir!int_CFOP)
            'End If
            If UFREM = TBAbrir!txt_UF Then
                TBGravar_NFe!Cbdplaca = TBTransporte!txt_Placa
                TBGravar_NFe!CbdUF_veictransp = TBTransporte!txt_UF_Placa
            End If
            TBGravar_NFe!CbdRNTC = IIf(IsNull(TBTransporte!Codigo_ANTT), "", TBTransporte!Codigo_ANTT)
            
            If TBAbrir!int_TipoNota = 1 And (IsNull(TBAbrir!txt_UF) = True Or TBAbrir!txt_UF = "" Or TBAbrir!txt_UF = "EX") Then
                TBGravar_NFe!CbdUFEmbarq = TBTransporte!UF_embarque
                TBGravar_NFe!CbdxLocEmbarq = TBTransporte!Local_embarque
            End If
            
            If TBMaquinas!Simples = True Then
                TBGravar_NFe!CbdCRT = 1
            ElseIf TBMaquinas!Simples1 = True Then
                TBGravar_NFe!CbdCRT = 2
            Else
                TBGravar_NFe!CbdCRT = 3
            End If
            
            Set TBGravar_NFe1 = CreateObject("adodb.recordset")
            TBGravar_NFe1.Open "Select * from CBD001VOL", Conexao_NFe, adOpenKeyset, adLockOptimistic
            TBGravar_NFe1.AddNew
            TBGravar_NFe1!CbdEmpCodigo = TBMaquinas!CODIGO
            TBGravar_NFe1!CbdNtfSerie = TBAbrir!Serie
            TBGravar_NFe1!CbdNtfNumero = OF
            TBGravar_NFe1!CbdSeqVol = 1 'Novo layout da Sefaz (3.10)
            TBGravar_NFe1!CbdnVol = IIf(IsNull(TBTransporte!Numeracao), 0, TBTransporte!Numeracao)
            TBGravar_NFe1!CbdqVol = IIf(IsNull(TBTransporte!int_Qtd_Transp), 0, TBTransporte!int_Qtd_Transp)
            TBGravar_NFe1!Cbdesp = IIf(IsNull(TBTransporte!txt_Especie), "", TBTransporte!txt_Especie)
            TBGravar_NFe1!Cbdmarca = IIf(IsNull(TBTransporte!txt_Marca), "", TBTransporte!txt_Marca)
            TBGravar_NFe1!CbdpesoL_transp = IIf(IsNull(TBTransporte!dbl_Peso_Liquido), 0, TBTransporte!dbl_Peso_Liquido)
            TBGravar_NFe1!CbdpesoB_transp = IIf(IsNull(TBTransporte!dbl_Peso_Bruto), 0, TBTransporte!dbl_Peso_Bruto)
            TBGravar_NFe1.Update
            TBGravar_NFe1.Close
            
        End If
        TBTransporte.Close

        Familiatext = ""
        DadosAdicionaisTexto = ""
        Set TBControleNF = CreateObject("adodb.recordset")
        TBControleNF.Open "Select * from tbl_DadosAdicionais where ID_nota = " & ID_nota, Conexao, adOpenKeyset, adLockOptimistic
        If TBControleNF.EOF = False Then
            If IsNull(TBControleNF!mem_corpo) = False And TBControleNF!mem_corpo <> "" Then TBGravar_NFe!CbdinfAdFisco = DS_RemoverAcentos(Trim(TBControleNF!mem_corpo)) Else TBGravar_NFe!CbdinfAdFisco = Null
            If IsNull(TBControleNF!mem_DadosAdicionais) = False And TBControleNF!mem_DadosAdicionais <> "" Then DadosAdicionaisTexto = DS_RemoverAcentos(Trim(TBControleNF!mem_DadosAdicionais)) Else DadosAdicionaisTexto = ""
        End If
        TBControleNF.Close
        
        endereco_entrega = ""
        If TBFI!DA_entrega = True Then
            Set TBClientes = CreateObject("adodb.recordset")
            TBClientes.Open "Select * from clientes_entrega where identrega = " & TBFI!ID_entrega, Conexao, adOpenKeyset, adLockOptimistic
            If TBClientes.EOF = False Then
                If IsNull(TBClientes!Tipo_endereco) = False And TBClientes!Tipo_endereco <> "" Then
                    Endereco = TBClientes!Tipo_endereco & ": " & IIf(IsNull(TBClientes!endereco_entrega), "", TBClientes!endereco_entrega)
                Else
                    Endereco = IIf(IsNull(TBClientes!endereco_entrega), "", TBClientes!endereco_entrega)
                End If
                If IsNull(TBClientes!Tipo_bairro) = False And TBClientes!Tipo_bairro <> "" Then
                    Bairro = TBClientes!Tipo_bairro & ": " & IIf(IsNull(TBClientes!bairro_entrega), "", TBClientes!bairro_entrega)
                Else
                    Bairro = IIf(IsNull(TBClientes!bairro_entrega), "", TBClientes!bairro_entrega)
                End If
                endereco_entrega = Endereco & " - " & IIf(IsNull(TBClientes!Numero), "", TBClientes!Numero) & " - " & Bairro & " - " & IIf(IsNull(TBClientes!cidade_entrega), "", TBClientes!cidade_entrega) & " - " & IIf(IsNull(TBClientes!uf_entrega), "", TBClientes!uf_entrega) & " - " & IIf(IsNull(TBClientes!cep_entrega), "", TBClientes!cep_entrega)
            End If
            TBClientes.Close
        End If
        
        Endereco_cobranca = ""
        If TBFI!DA_cobranca = True Then
            Set TBClientes = CreateObject("adodb.recordset")
            TBClientes.Open "Select * from clientes_cobranca where idcobranca = " & TBFI!ID_Cobranca, Conexao, adOpenKeyset, adLockOptimistic
            If TBClientes.EOF = False Then
                If IsNull(TBClientes!Tipo_endereco) = False And TBClientes!Tipo_endereco <> "" Then
                    Endereco = TBClientes!Tipo_endereco & ": " & IIf(IsNull(TBClientes!Endereco_cobranca), "", TBClientes!Endereco_cobranca)
                Else
                    Endereco = IIf(IsNull(TBClientes!Endereco_cobranca), "", TBClientes!Endereco_cobranca)
                End If
                If IsNull(TBClientes!Tipo_bairro) = False And TBClientes!Tipo_bairro <> "" Then
                    Bairro = TBClientes!Tipo_bairro & ": " & IIf(IsNull(TBClientes!bairro_cobranca), "", TBClientes!bairro_cobranca)
                Else
                    Bairro = IIf(IsNull(TBClientes!bairro_cobranca), "", TBClientes!bairro_cobranca)
                End If
                Endereco_cobranca = Endereco & " - " & IIf(IsNull(TBClientes!Numero), "", TBClientes!Numero) & " - " & Bairro & " - " & IIf(IsNull(TBClientes!cidade_cobranca), "", TBClientes!cidade_cobranca) & " - " & IIf(IsNull(TBClientes!uf_cobranca), "", TBClientes!uf_cobranca) & " - " & IIf(IsNull(TBClientes!cep_cobranca), "", TBClientes!cep_cobranca)
            End If
            TBClientes.Close
        End If
                        
        If DadosAdicionaisTexto <> "" Or endereco_entrega <> "" Or Endereco_cobranca <> "" Or DAPartilhaICMS <> "" Then
            If DadosAdicionaisTexto <> "" Then Familiatext = DadosAdicionaisTexto
            If endereco_entrega <> "" Then
               If Familiatext <> "" Then Familiatext = Familiatext & "|Endereço de entrega: " & endereco_entrega Else Familiatext = "Endereço de entrega: " & endereco_entrega
            End If
            If Endereco_cobranca <> "" Then
               If Familiatext <> "" Then Familiatext = Familiatext & "|Endereço de cobrança: " & Endereco_cobranca Else Familiatext = "Endereço de cobrança: " & Endereco_cobranca
            End If
            If DAPartilhaICMS <> "" Then
               If Familiatext <> "" Then Familiatext = Familiatext & "|" & DAPartilhaICMS Else Familiatext = DAPartilhaICMS
            End If
            TBGravar_NFe!CbdinfCpl = DS_RemoverAcentos(LTrim(Trim(Familiatext)))
        Else
            TBGravar_NFe!CbdinfCpl = Null
        End If
        
        TBGravar_NFe!CbdEmailArquivos = TBFI!Enviar_Email
        TBGravar_NFe!CbdTitGenerico = IIf(IsNull(TBFI!Titulo_canhoto_DANFE), "", TBFI!Titulo_canhoto_DANFE)
        TBGravar_NFe!CbdTxtGenerico = IIf(IsNull(TBFI!Texto_canhoto_DANFE), "", TBFI!Texto_canhoto_DANFE)
        If IsNull(TBFI!Enviar_DANFE_email) = False And TBFI!Enviar_DANFE_email <> "" Then TBGravar_NFe!CbdNfeEmailDANFE = TBFI!Enviar_DANFE_email Else TBGravar_NFe!CbdNfeEmailDANFE = "S"
        
        'SCAN
        If TBFI!Forma_emissao = 3 Then
            If TBAbrir!Serie > "899" Then
                TBGravar_NFe!CbdNtfNumeroSCAN = OF
                TBGravar_NFe!CbdNtfSerieSCAN = TBAbrir!Serie
            ElseIf Len(TBAbrir!Serie) < 3 Then
                    TBGravar_NFe!CbdNtfNumeroSCAN = OF
                    Select Case Len(TBAbrir!Serie)
                        Case 1: TBGravar_NFe!CbdNtfSerieSCAN = "90" & TBAbrir!Serie
                        Case 2: TBGravar_NFe!CbdNtfSerieSCAN = "9" & TBAbrir!Serie
                    End Select
            End If
        End If
        
        TBGravar_NFe.Update
        TBGravar_NFe.Close
    End If
    TBFI.Close
    
    'Novo layout da Sefaz (3.10)
    'Autorização para obter XML
    Set TBGravar_NFe = CreateObject("adodb.recordset")
    TBGravar_NFe.Open "Select * from CBD001AUTXML", Conexao_NFe, adOpenKeyset, adLockOptimistic
    TBGravar_NFe.AddNew
    TBGravar_NFe!CbdEmpCodigo = TBMaquinas!CODIGO
    TBGravar_NFe!CbdNtfSerie = TBAbrir!Serie
    TBGravar_NFe!CbdNtfNumero = OF
    TBGravar_NFe!CbdSeqAut = 1
    TBGravar_NFe.Update

    'Transportadora
    Set TBTransporte = CreateObject("adodb.recordset")
    TBTransporte.Open "Select CF.Pessoa, DT.txt_CNPJ from tbl_Dados_Transp DT INNER JOIN Compras_fornecedores CF ON CF.IDCliente = DT.IdIntTransp and CF.Nome_Razao = DT.txt_Razao where DT.ID_Nota = " & ID_nota & " and DT.txt_CNPJ IS NOT NULL and DT.txt_CNPJ <> N'' and DT.enviarXML = 'True'", Conexao, adOpenKeyset, adLockOptimistic
    If TBTransporte.EOF = False Then
        Set TBGravar_NFe = CreateObject("adodb.recordset")
        TBGravar_NFe.Open "Select * from CBD001AUTXML", Conexao_NFe, adOpenKeyset, adLockOptimistic
        TBGravar_NFe.AddNew
        TBGravar_NFe!CbdEmpCodigo = TBMaquinas!CODIGO
        TBGravar_NFe!CbdNtfSerie = TBAbrir!Serie
        TBGravar_NFe!CbdNtfNumero = OF
        TBGravar_NFe!CbdSeqAut = 2
        If TBTransporte!txt_CNPJ <> TBAbrir!txt_CNPJ_CPF Then
            If Left(TBTransporte!Pessoa, 1) = "J" Then TBGravar_NFe!CbdCNPJ_aut = DS_RetornarNumeros(TBTransporte!txt_CNPJ) Else TBGravar_NFe!CbdCPF_aut = DS_RetornarNumeros(TBTransporte!txt_CNPJ)
        End If
        TBGravar_NFe.Update
    End If
    
    TBGravar_NFe.Close
    
    'Produtos
    Contador2 = 1
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select * from tbl_Detalhes_Nota where ID_nota = " & ID_nota & " order by Int_codigo", Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
        Do While TBProduto.EOF = False
            Set TBGravar_NFe = CreateObject("adodb.recordset")
            TBGravar_NFe.Open "Select * from CBD001DET", Conexao_NFe, adOpenKeyset, adLockOptimistic
            TBGravar_NFe.AddNew
            TBGravar_NFe!CbdEmpCodigo = TBMaquinas!CODIGO
            TBGravar_NFe!CbdNtfSerie = TBAbrir!Serie
            TBGravar_NFe!CbdNtfNumero = OF
            TBGravar_NFe!CbdnItem = Contador2
            
            'Verifica se é para utilizar o código de referência na DANFE
            Set TBCodigoDesc = CreateObject("adodb.recordset")
            TBCodigoDesc.Open "Select CodRef from tbl_Dados_Nota_Fiscal_NFe where ID_Nota = " & ID_nota, Conexao, adOpenKeyset, adLockReadOnly
            If TBCodigoDesc.EOF = False Then
                If TBCodigoDesc!CodRef = False Or IsNull(TBCodigoDesc!CodRef) = True Then
                    TBGravar_NFe!CbdcProd = DS_RemoverAcentos(TBProduto!int_Cod_Produto)
                    Set TBCodigoDesc = CreateObject("adodb.recordset")
                    TBCodigoDesc.Open "Select Codigo_ref_desc_DANFE from empresa where codigo = " & TBAbrir!ID_empresa, Conexao, adOpenKeyset, adLockReadOnly
                    If TBCodigoDesc!Codigo_ref_desc_DANFE = True Then
                        CodRef = 2
                    Else
                        CodRef = 0
                    End If
                Else
                    TBGravar_NFe!CbdcProd = DS_RemoverAcentos(TBProduto!N_referencia)
                    CodRef = 1
                End If
            End If
            TBCodigoDesc.Close
            
            If TBProduto!GTIN = "" Then
                TBGravar_NFe!CbdcEAN = "SEM GTIN"
                TBGravar_NFe!CbdcEANTrib = "SEM GTIN"
            Else
                TBGravar_NFe!CbdcEAN = DS_RemoverAcentos(IIf(IsNull(TBProduto!GTIN), "SEM GTIN", TBProduto!GTIN))
                TBGravar_NFe!CbdcEANTrib = DS_RemoverAcentos(IIf(IsNull(TBProduto!GTIN), "SEM GTIN", TBProduto!GTIN))
            End If
            
            CompLetra = 0
            If IsNull(TBProduto!N_referencia) = False And TBProduto!N_referencia <> "" And TBProduto!N_referencia <> TBProduto!int_Cod_Produto And CodRef = 2 Then CompLetra = Len(Trim(TBProduto!N_referencia)) + 3
            If IsNull(TBProduto!Complemento_descricao) = False And TBProduto!Complemento_descricao <> "" Then CompLetra = Len(Trim(TBProduto!Complemento_descricao)) + 3
            If IsNull(TBProduto!PCCliente) = False And TBProduto!PCCliente <> "" Then CompLetra = CompLetra + Len(Trim(TBProduto!PCCliente)) + 8
            If IsNull(TBProduto!N_item) = False And TBProduto!N_item <> "" Then CompLetra = CompLetra + Len(Trim(TBProduto!N_item)) + 11
                        
            DesenhoProduto = Left(Trim(TBProduto!Txt_descricao), 120 - CompLetra)
            If CodRef = 2 Then
                If IsNull(TBProduto!N_referencia) = False And TBProduto!N_referencia <> "" And TBProduto!N_referencia <> TBProduto!int_Cod_Produto Then
                    DesenhoProduto = "(" & TBProduto!N_referencia & ") - " & DesenhoProduto
                Else
                    If Len(TBAbrir!txt_tipocliente) = 2 Then TipoFiltro = "C" Else TipoFiltro = "F"
                    Set TBItem = CreateObject("adodb.recordset")
                    TBItem.Open "Select IA.N_Referencia from item_aplicacoes IA INNER JOIN projproduto P ON IA.codproduto = P.codproduto where P.Desenho = '" & TBProduto!int_Cod_Produto & "' and IA.ID_cliente_forn = " & TBAbrir!Id_Int_Cliente & " and IA.Tipo = '" & TipoFiltro & "' and IA.N_Referencia IS NOT NULL and IA.N_Referencia <> '" & TBProduto!int_Cod_Produto & "'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBItem.EOF = False Then
                        DesenhoProduto = "(" & TBItem!N_referencia & ") - " & DesenhoProduto
                    End If
                    TBItem.Close
                End If
            End If
            If IsNull(TBProduto!Complemento_descricao) = False And TBProduto!Complemento_descricao <> "" Then DesenhoProduto = DesenhoProduto & " - " & Trim(TBProduto!Complemento_descricao)
            If IsNull(TBProduto!PCCliente) = False And TBProduto!PCCliente <> "" Then DesenhoProduto = DesenhoProduto & " - Ped. " & Trim(TBProduto!PCCliente)
            If IsNull(TBProduto!N_item) = False And TBProduto!N_item <> "" Then DesenhoProduto = DesenhoProduto & " - N. item " & Trim(TBProduto!N_item)
            
            TBGravar_NFe!CbdxProd = DS_RemoverAcentos(Left(DesenhoProduto, 120))
            
            Set TBControleNF = CreateObject("adodb.recordset")
            TBControleNF.Open "Select IDIntClasse, CEST from tbl_ClassificacaoFiscal where Idclass = " & TBProduto!ID_CF, Conexao, adOpenKeyset, adLockOptimistic
            If TBControleNF.EOF = False Then
                TBGravar_NFe!CbdNCM = DS_RetornarNumeros(TBControleNF!IDIntClasse)
                If IsNull(TBControleNF!CEST) = False Then TBGravar_NFe!CbdCEST = DS_RetornarNumeros(TBControleNF!CEST)
            End If
            
            Set TBControleNF = CreateObject("adodb.recordset")
            TBControleNF.Open "Select id_CFOP, Devolucao from tbl_NaturezaOperacao where IDCountCfop = " & TBProduto!ID_CFOP, Conexao, adOpenKeyset, adLockOptimistic
            If TBControleNF.EOF = False Then
                If Len(TBControleNF!ID_CFOP) > 5 Then
                    If TBProduto!retorno = True Then
                        If Left(TBControleNF!ID_CFOP, 5) = "5.902" Or Left(TBControleNF!ID_CFOP, 5) = "6.902" Or Left(TBControleNF!ID_CFOP, 5) = "5.916" Or Left(TBControleNF!ID_CFOP, 5) = "6.916" Or Left(TBControleNF!ID_CFOP, 5) = "5.925" Or Left(TBControleNF!ID_CFOP, 5) = "6.925" Then
                            CFOP_Produto = DS_RetornarNumeros(Left(TBControleNF!ID_CFOP, 5))
                        ElseIf Right(TBControleNF!ID_CFOP, 5) = "5.902" Or Right(TBControleNF!ID_CFOP, 5) = "6.902" Or Right(TBControleNF!ID_CFOP, 5) = "5.916" Or Right(TBControleNF!ID_CFOP, 5) = "6.916" Or Right(TBControleNF!ID_CFOP, 5) = "5.925" Or Right(TBControleNF!ID_CFOP, 5) = "6.925" Then
                                CFOP_Produto = DS_RetornarNumeros(Right(TBControleNF!ID_CFOP, 5))
                        End If
                    Else
                        If Left(TBControleNF!ID_CFOP, 5) <> "5.902" And Left(TBControleNF!ID_CFOP, 5) <> "6.902" And Left(TBControleNF!ID_CFOP, 5) <> "5.916" And Left(TBControleNF!ID_CFOP, 5) <> "6.916" And Left(TBControleNF!ID_CFOP, 5) <> "5.925" And Left(TBControleNF!ID_CFOP, 5) <> "6.925" Then
                            CFOP_Produto = DS_RetornarNumeros(Left(TBControleNF!ID_CFOP, 5))
                        ElseIf Right(TBControleNF!ID_CFOP, 5) <> "5.902" And Right(TBControleNF!ID_CFOP, 5) <> "6.902" And Right(TBControleNF!ID_CFOP, 5) <> "5.916" And Right(TBControleNF!ID_CFOP, 5) <> "6.916" And Right(TBControleNF!ID_CFOP, 5) <> "5.925" And Right(TBControleNF!ID_CFOP, 5) <> "6.925" Then
                                CFOP_Produto = DS_RetornarNumeros(Right(TBControleNF!ID_CFOP, 5))
                        End If
                    End If
                Else
                    CFOP_Produto = DS_RetornarNumeros(TBControleNF!ID_CFOP)
                End If
                TBGravar_NFe!CbdCFOP = CFOP_Produto
                If TBControleNF!Devolucao = True And TBMaquinas!Simples = True Then TBGravar_NFe!CbdvIPIDevol = TBProduto!dbl_valoripi 'Novo layout da Sefaz (4.0)
            End If
            TBControleNF.Close
            
            TBGravar_NFe!CbduCOM = DS_RemoverAcentos(TBProduto!Unidade_com)
            TBGravar_NFe!CbdqCOM = TBProduto!int_Qtd
            TBGravar_NFe!CbdvUnCom = TBProduto!dbl_ValorUnitario
            TBGravar_NFe!CbdvProd = TBProduto!dbl_ValorTotal
            TBGravar_NFe!CbduTrib = DS_RemoverAcentos(TBProduto!Unidade_com)
            TBGravar_NFe!CbdqTrib = TBProduto!int_Qtd
            TBGravar_NFe!CbdvUnTrib = TBProduto!dbl_ValorUnitario
            TBGravar_NFe!CbdvFrete = IIf(IsNull(TBProduto!Valor_frete), 0, TBProduto!Valor_frete)
            TBGravar_NFe!CbdvSeg = IIf(IsNull(TBProduto!Valor_seguro), 0, TBProduto!Valor_seguro)
            TBGravar_NFe!CbdvDesc = IIf(IsNull(TBProduto!Valor_desconto), 0, TBProduto!Valor_desconto) + IIf(IsNull(TBProduto!Valor_desconto_SUFRAMA), 0, TBProduto!Valor_desconto_SUFRAMA)
            TBGravar_NFe!CbdvOutro_item = IIf(IsNull(TBProduto!Valor_acessorias), 0, TBProduto!Valor_acessorias)
            If IsNull(TBProduto!Valor_aprox_tributos) = False And TBProduto!Valor_aprox_tributos <> "" And TBProduto!Valor_aprox_tributos <> "0" Then TBGravar_NFe!CBDvTotTrib = TBProduto!Valor_aprox_tributos
            
            If TBProduto!retorno = True Then
                Set TBFIltro = CreateObject("adodb.recordset")
                TBFIltro.Open "Select * from tbl_NaturezaOperacao where IDCountCfop = " & TBProduto!ID_CFOP & " and Soma_retorno_totalnf = 'True'", Conexao, adOpenKeyset, adLockOptimistic
                If TBFIltro.EOF = False Then
                    TBGravar_NFe!CbdIndTot = 1 'O valor do produto compõe o valor total da NF
                Else
                    TBGravar_NFe!CbdIndTot = 0 'O valor do produto não compõe o valor total da NF
                End If
                TBFIltro.Close
            Else
                TBGravar_NFe!CbdIndTot = 1 'O valor do produto compõe o valor total da NF
            End If
            
            If IsNull(TBProduto!Inf_adicionais_prod) = False And TBProduto!Inf_adicionais_prod <> "" Then TBGravar_NFe!CbdinfAdProd = Trim(TBProduto!Inf_adicionais_prod) 'Novo layout da Sefaz (3.10)
            If IsNull(TBProduto!N_item) = False And TBProduto!N_item <> "" Then
                NItemTexto = DS_RetornarNumeros(TBProduto!N_item)
                TBGravar_NFe!CbdnItemPed = IIf(NItemTexto = "", Null, NItemTexto)
            End If
            If IsNull(TBProduto!PCCliente) = False And TBProduto!PCCliente <> "" Then TBGravar_NFe!CbdxPed_item = Left(Trim(TBProduto!PCCliente), 15)
            
            'ICMS
            If IsNull(TBProduto!txt_CST) = False And TBProduto!txt_CST <> "" Then
                Set TBCST = CreateObject("adodb.recordset")
                If Len(TBProduto!txt_CST) = 4 Then FimCST = Right(TBProduto!txt_CST, 3) Else FimCST = Right(TBProduto!txt_CST, 2)
                TBCST.Open "Select * from tbl_Detalhes_Nota_CST_ICMS where id_item = " & TBProduto!Int_codigo, Conexao, adOpenKeyset, adLockOptimistic
                If TBCST.EOF = False Then
                    If FimCST = "00" Or FimCST = 10 Or FimCST = "20" Or FimCST = "51" Or FimCST = 70 Or FimCST = 90 Or FimCST = "900" Then
                        TBGravar_NFe!CbdvBCICMS = TBCST!Valor_BC
                        TBGravar_NFe!CbdvICMS = TBCST!Valor_ICMS
                    End If
                    If FimCST = "10" Or FimCST = "60" Or FimCST = "90" Or FimCST = "201" Or FimCST = "202" Or FimCST = "203" Or FimCST = "900" Then
                        TBGravar_NFe!CbdvBCICMSST = TBCST!Valor_BC_ST
                        TBGravar_NFe!CbdvICMSST = TBCST!Valor_ICMS_ST
                    End If
                    TBGravar_NFe!CbdvBCUFDest = IIf(IsNull(TBCST!Valor_BC_ICMS_UF_dest), 0, TBCST!Valor_BC_ICMS_UF_dest)
                    If IsNull(TBCST!Valor_ICMS_INT_UF_dest) = False And TBCST!Valor_ICMS_INT_UF_dest > 0 Then
                        If IsNull(TBAbrir!txt_UF) = True Or TBAbrir!txt_UF = "" Or TBAbrir!txt_UF = "EX" Then
                            TBGravar_NFe!CbdpICMSInter = 4
                        Else
                            ProcBuscaTributos IIf(IsNull(TBProduto!ID_CF), 0, TBProduto!ID_CF)
                            ProcVerificaRegiao TBAbrir!txt_UF, TBAbrir!Id_Int_Cliente, TBAbrir!txt_Razao_Nome
                            TBGravar_NFe!CbdpICMSInter = vRegiao(0, 1)
                        End If
                        Set TBFIltro = CreateObject("adodb.recordset")
                        TBFIltro.Open "Select ICMS_interno from regioes where UF = '" & TBAbrir!txt_UF & "'", Conexao, adOpenKeyset, adLockOptimistic
                        If TBFIltro.EOF = False Then
                            TBGravar_NFe!CbdpICMSUFDest = IIf(IsNull(TBFIltro!ICMS_interno), 0, TBFIltro!ICMS_interno)
                        End If
                        TBFIltro.Close
                    Else
                        TBGravar_NFe!CbdpICMSInter = 0
                        TBGravar_NFe!CbdpICMSUFDest = 0
                    End If
                    TBGravar_NFe!CbdpICMSInterPart = IIf(IsNull(TBCST!Percentual_provisorio), 0, TBCST!Percentual_provisorio)
                    TBGravar_NFe!CbdvICMSUFDest = IIf(IsNull(TBCST!Valor_ICMS_INT_UF_dest), 0, TBCST!Valor_ICMS_INT_UF_dest)
                    TBGravar_NFe!CbdvICMSUFRemet = IIf(IsNull(TBCST!Valor_ICMS_INT_UF_rem), 0, TBCST!Valor_ICMS_INT_UF_rem)
                    TBGravar_NFe!CbdpFCPUFDest = IIf(IsNull(TBCST!Percentual_FCP), 0, TBCST!Percentual_FCP)
                    TBGravar_NFe!CbdvFCPUFDest = IIf(IsNull(TBCST!Valor_ICMS_FCP), 0, TBCST!Valor_ICMS_FCP)
                End If
                TBCST.Close
            End If
            
            If IsNull(TBProduto!Codigo_enquadramento_IPI) = False Then TBGravar_NFe!CbdcEnq = TBProduto!Codigo_enquadramento_IPI
            
            Set TBCiclo = CreateObject("adodb.recordset")
            TBCiclo.Open "Select * from tbl_Detalhes_Nota_NFe where ID_item = " & TBProduto!Int_codigo, Conexao, adOpenKeyset, adLockOptimistic
            If TBCiclo.EOF = False Then
                
                If IsNull(TBCiclo!Codigo_ANP) = False And TBCiclo!Codigo_ANP <> "" Then
                    TBGravar_NFe!CbdcProdANP = TBCiclo!Codigo_ANP
                    TBGravar_NFe!CbddescANP = TBCiclo!Descricao_ANP
                    TBGravar_NFe!CbdUFcons = TBCiclo!UF_consumo
                    TBGravar_NFe!CbdnTipoItem = TBCiclo!Tipo_produto
                End If
                
                If IsNull(TBCiclo!Local_desembaraco) = False And TBCiclo!Local_desembaraco <> "" Then
                    TBGravar_NFe!CbdvBC_imp = TBCiclo!Valor_BC_importacao
                    TBGravar_NFe!CbdvDespAdu = TBCiclo!Valor_despesas
                    TBGravar_NFe!CbdvII = TBCiclo!Valor_imposto_importacao
                    TBGravar_NFe!CbdvIOF = TBCiclo!Valor_imposto_OperacoesFinanceiras
                End If
                                
                If IsNull(TBCiclo!Documento_importacao) = False And TBCiclo!Documento_importacao <> "" And IsNull(TBCiclo!Numero_adicao) = False And TBCiclo!Numero_adicao <> "" And IsNull(TBCiclo!Numero_sequencial) = False And TBCiclo!Numero_sequencial <> "" And IsNull(TBCiclo!Codigo_fabricante) = False And TBCiclo!Codigo_fabricante <> "" Then
                    Set TBGravar_NFe1 = CreateObject("adodb.recordset")
                    TBGravar_NFe1.Open "Select * from CBD001DETADICOES", Conexao_NFe, adOpenKeyset, adLockOptimistic
                    TBGravar_NFe1.AddNew
                    TBGravar_NFe1!CbdEmpCodigo = TBMaquinas!CODIGO
                    TBGravar_NFe1!CbdNtfSerie = TBAbrir!Serie
                    TBGravar_NFe1!CbdNtfNumero = OF
                    TBGravar_NFe1!CbdnItem = Contador2
                    TBGravar_NFe1!CbdnDI = TBCiclo!Documento_importacao
                    TBGravar_NFe1!CbdnAdicao = TBCiclo!Numero_adicao
                    TBGravar_NFe1!CbdnSeqAdic = TBCiclo!Numero_sequencial
                    TBGravar_NFe1!CbdcFabricante = TBCiclo!Codigo_fabricante
                    TBGravar_NFe1.Update
                    TBGravar_NFe1.Close
                    
                    Set TBGravar_NFe1 = CreateObject("adodb.recordset")
                    TBGravar_NFe1.Open "Select * from CBD001DETDI", Conexao_NFe, adOpenKeyset, adLockOptimistic
                    TBGravar_NFe1.AddNew
                    TBGravar_NFe1!CbdEmpCodigo = TBMaquinas!CODIGO
                    TBGravar_NFe1!CbdNtfSerie = TBAbrir!Serie
                    TBGravar_NFe1!CbdNtfNumero = OF
                    TBGravar_NFe1!CbdnItem = Contador2
                    TBGravar_NFe1!CbdnDI = TBCiclo!Documento_importacao
                    TBGravar_NFe1!CbddDi = Format(TBCiclo!Data_registro, "yyyy-mm-dd")
                    TBGravar_NFe1!CbdxLocDesemb = TBCiclo!Local_desembaraco
                    TBGravar_NFe1!CbdUFDesemb = TBCiclo!UF_desembaraco
                    TBGravar_NFe1!CbdcExportador = TBCiclo!Codigo_exportador
                    TBGravar_NFe1!CbddDesemb = Format(TBCiclo!Data_desembaraco, "yyyy-mm-dd")
                    
                    'Novo layout da Sefaz (3.10)
                    TBGravar_NFe1!CbdTpViaTransp = TBCiclo!Via_transp
                    If TBCiclo!Via_transp = 1 And IsNull(TBCiclo!Valor_AFRMM) = False Then TBGravar_NFe1!CbdvAFRMM = TBCiclo!Valor_AFRMM
                    TBGravar_NFe1!CbdTpIntermedio = TBCiclo!Forma_imp
                    If TBCiclo!Forma_imp <> 1 Then
                        TBGravar_NFe1!CbdCNPJ_adq = DS_RetornarNumeros(TBMaquinas!CNPJ)
                        TBGravar_NFe1!CbdUFTerceiro = TBMaquinas!UF
                    End If
                    
                    TBGravar_NFe1.Update
                    TBGravar_NFe1.Close
                End If
            End If
           
            'Novo layout da Sefaz (3.10)
            'Grupo de informações de exportação para o item
            If Left(CFOP_Produto, 1) = 7 Then
                Set TBGravar_NFe1 = CreateObject("adodb.recordset")
                TBGravar_NFe1.Open "Select * from CBD001DETEXPORT", Conexao_NFe, adOpenKeyset, adLockOptimistic
                TBGravar_NFe1.AddNew
                TBGravar_NFe1!CbdEmpCodigo = TBMaquinas!CODIGO
                TBGravar_NFe1!CbdNtfSerie = TBAbrir!Serie
                TBGravar_NFe1!CbdNtfNumero = OF
                TBGravar_NFe1!CbdnItem = Contador2
                Set TBCiclo = CreateObject("adodb.recordset")
                TBCiclo.Open "Select * from CBD001DETEXPORT where CbdEmpCodigo = " & TBMaquinas!CODIGO & " and CbdNtfSerie = " & TBAbrir!Serie & " and CbdNtfNumero = " & OF, Conexao_NFe, adOpenKeyset, adLockOptimistic
                If TBCiclo.EOF = False Then
                    TBGravar_NFe1!CbdSeqExport = TBCiclo.RecordCount + 1
                Else
                    TBGravar_NFe1!CbdSeqExport = 1
                End If
                TBCiclo.Close
                TBGravar_NFe1.Update
                TBGravar_NFe1.Close
            End If
            
            'ICMS
            If IsNull(TBProduto!txt_CST) = False And TBProduto!txt_CST <> "" Then
                If Len(TBProduto!txt_CST) = 4 Then FimCST = Right(TBProduto!txt_CST, 3) Else FimCST = Right(TBProduto!txt_CST, 2)
                Set TBCST = CreateObject("adodb.recordset")
                TBCST.Open "Select * from tbl_Detalhes_Nota_CST_ICMS where id_item = " & TBProduto!Int_codigo, Conexao, adOpenKeyset, adLockOptimistic
                If TBCST.EOF = False Then
                    Set TBGravar_NFe1 = CreateObject("adodb.recordset")
                    TBGravar_NFe1.Open "Select * from CBD001DETICMSNORMALST", Conexao_NFe, adOpenKeyset, adLockOptimistic
                    TBGravar_NFe1.AddNew
                    TBGravar_NFe1!CbdEmpCodigo = TBMaquinas!CODIGO
                    TBGravar_NFe1!CbdNtfSerie = TBAbrir!Serie
                    TBGravar_NFe1!CbdNtfNumero = OF
                    TBGravar_NFe1!CbdnItem = Contador2
                    TBGravar_NFe1!CbdCST = FimCST
                    TBGravar_NFe1!Cbdorig = TBCST!Origem_mercadoria
                    If FimCST = "00" Or FimCST = "10" Or FimCST = "20" Or FimCST = "51" Or FimCST = "70" Or FimCST = "90" Or FimCST = "900" Then
                        TBGravar_NFe1!CbdmodBC = TBCST!Modalidade_determinacao
                        TBGravar_NFe1!CbdvBC = TBCST!Valor_BC
                        TBGravar_NFe1!CbdpICMS = TBProduto!int_ICMS
                        TBGravar_NFe1!CbdvICMS_icms = TBCST!Valor_ICMS
                        If FimCST = "51" Then
                            TBGravar_NFe1!CbdvICMSOP = TBCST!Valor_ICMS
                            TBGravar_NFe1!CbdpDif = TBCST!Percentual_ICMS_DIF
                            TBGravar_NFe1!CbdvICMSDif = TBCST!Valor_ICMS_DIF
                        End If
                        'Novo layout da Sefaz (4.0)
                        If FimCST <> "900" Then
                            'TBGravar_NFe1!CbdvFCP = TBCST!Valor_ICMS_FCP
                            'TBGravar_NFe1!CbdpFCP = TBCST!Percentual_FCP
                            'TBGravar_NFe1!CbdvBCFCP = TBCST!Valor_BC_ICMS_UF_dest
                        End If
                    End If
                    If FimCST = "10" Or FimCST = "60" Or FimCST = "70" Or FimCST = "90" Or FimCST = "201" Or FimCST = "202" Or FimCST = "203" Or FimCST = "900" Then
                        If FimCST <> "60" Then
                            TBGravar_NFe1!CbdmodBCST = IIf(IsNull(TBCST!Modalidade_determinacao_ST), 4, TBCST!Modalidade_determinacao_ST)
                            TBGravar_NFe1!CbdpRedBCST = TBCST!Percentual_reducao_BC_ST
                            TBGravar_NFe1!CbdpICMSST = TBCST!Aliquota_imposto_ST
                        End If
                        TBGravar_NFe1!CbdvBCST = TBCST!Valor_BC_ST
                        TBGravar_NFe1!CbdvICMSST_icms = TBCST!Valor_ICMS_ST
                    End If
                    If FimCST = "20" Or FimCST = "51" Or FimCST = "70" Or FimCST = "90" Or FimCST = "900" Then TBGravar_NFe1!CbdpRedBC = TBCST!Percentual_reducao_BC
                    If FimCST = "101" Or FimCST = "201" Or FimCST = "900" Then
                        TBGravar_NFe1!CbdpCredSN = TBCST!ICMS_SN
                        TBGravar_NFe1!CbdvCredICMSSN = TBCST!Valor_ICMS_SN
                    End If
                    
                    'Novo layout da Sefaz (3.10) Não é obrigatório
                    'TBGravar_NFe1!CbdmotDesICMS = TBCST!Motivo_ICMS_desonerado
                    'TBGravar_NFe1!CbdvICMSDeson = TBCST!Valor_ICMS_desonerado
                    
                    TBGravar_NFe1.Update
                    TBGravar_NFe1.Close
                End If
                TBCST.Close
            End If
              
            'IPI
            If IsNull(TBProduto!CST_IPI) = False And TBProduto!CST_IPI <> "" Then
                FimCST = Right(TBProduto!CST_IPI, 2)
                Set TBCST = CreateObject("adodb.recordset")
                TBCST.Open "Select * from tbl_Detalhes_Nota_CST_IPI where id_item = " & TBProduto!Int_codigo, Conexao, adOpenKeyset, adLockOptimistic
                If TBCST.EOF = False Then
                    Set TBGravar_NFe1 = CreateObject("adodb.recordset")
                    TBGravar_NFe1.Open "Select * from CBD001DETIPI", Conexao_NFe, adOpenKeyset, adLockOptimistic
                    TBGravar_NFe1.AddNew
                    TBGravar_NFe1!CbdEmpCodigo = TBMaquinas!CODIGO
                    TBGravar_NFe1!CbdNtfSerie = TBAbrir!Serie
                    TBGravar_NFe1!CbdNtfNumero = OF
                    TBGravar_NFe1!CbdnItem = Contador2
                    TBGravar_NFe1!CbdCST_IPI = FimCST
                    If FimCST = "00" Or FimCST = "49" Or FimCST = "50" Or FimCST = "99" Then
                        TBGravar_NFe1!CbdvBC_IPI = TBCST!Valor_BC
                        TBGravar_NFe1!CbdpIPI = TBProduto!int_IPI
                        If IsNull(TBProduto!dbl_ValorTotal) = True Or TBProduto!dbl_ValorTotal = 0 Then
                            TBGravar_NFe1!CbdvIPI = Format((TBCST!Valor_BC * TBProduto!int_IPI) / 100, "0.00")
                        Else
                            TBGravar_NFe1!CbdvIPI = Format(TBProduto!dbl_valoripi, "0.00")
                        End If
                    End If
                    TBGravar_NFe1.Update
                    TBGravar_NFe1.Close
                End If
                TBCST.Close
            End If
            
            'PIS
            If IsNull(TBProduto!CST_PIS) = False And TBProduto!CST_PIS <> "" Then
                FimCST = Right(TBProduto!CST_PIS, 2)
                Set TBCST = CreateObject("adodb.recordset")
                TBCST.Open "Select * from tbl_Detalhes_Nota_CST_PIS where id_item = " & TBProduto!Int_codigo, Conexao, adOpenKeyset, adLockOptimistic
                If TBCST.EOF = False Then
                    Set TBGravar_NFe1 = CreateObject("adodb.recordset")
                    TBGravar_NFe1.Open "Select * from CBD001DETPIS", Conexao_NFe, adOpenKeyset, adLockOptimistic
                    TBGravar_NFe1.AddNew
                    TBGravar_NFe1!CbdEmpCodigo = TBMaquinas!CODIGO
                    TBGravar_NFe1!CbdNtfSerie = TBAbrir!Serie
                    TBGravar_NFe1!CbdNtfNumero = OF
                    TBGravar_NFe1!CbdnItem = Contador2
                    TBGravar_NFe1!CbdCST_pis = FimCST
                    If FimCST = "01" Or FimCST = "02" Or FimCST = "03" Or FimCST = "49" Or FimCST = "98" Or FimCST = "99" Then TBGravar_NFe1!CbdvBC_pis = TBCST!Valor_BC
                    TBGravar_NFe1!CbdpPIS = TBProduto!PIS_Prod
                    TBGravar_NFe1!CbdvPIS = TBProduto!Total_PIS_prod
                    TBGravar_NFe1!CbdqBCprod_pis = TBProduto!int_Qtd
                    TBGravar_NFe1!CbdvAliqProd_pis = TBProduto!PIS_Prod
                    TBGravar_NFe1.Update
                    TBGravar_NFe1.Close
                End If
                TBCST.Close
            End If
            
            'Cofins
            If IsNull(TBProduto!CST_Cofins) = False And TBProduto!CST_Cofins <> "" Then
                FimCST = Right(TBProduto!CST_Cofins, 2)
                Set TBCST = CreateObject("adodb.recordset")
                TBCST.Open "Select * from tbl_Detalhes_Nota_CST_Cofins where id_item = " & TBProduto!Int_codigo, Conexao, adOpenKeyset, adLockOptimistic
                If TBCST.EOF = False Then
                    Set TBGravar_NFe1 = CreateObject("adodb.recordset")
                    TBGravar_NFe1.Open "Select * from CBD001DETCOFINS", Conexao_NFe, adOpenKeyset, adLockOptimistic
                    TBGravar_NFe1.AddNew
                    TBGravar_NFe1!CbdEmpCodigo = TBMaquinas!CODIGO
                    TBGravar_NFe1!CbdNtfSerie = TBAbrir!Serie
                    TBGravar_NFe1!CbdNtfNumero = OF
                    TBGravar_NFe1!CbdnItem = Contador2
                    TBGravar_NFe1!CbdCST_cofins = TBProduto!CST_Cofins
                    If FimCST = "01" Or FimCST = "03" Or FimCST = "49" Or FimCST = "98" Or FimCST = "99" Then TBGravar_NFe1!CbdvBC_cofins = TBCST!Valor_BC
                    TBGravar_NFe1!CbdpCOFINS = TBProduto!Cofins_Prod
                    TBGravar_NFe1!CbdvCOFINS = TBProduto!Total_Cofins_prod
                    TBGravar_NFe1!CbdqBCProd_cofins = TBProduto!int_Qtd
                    TBGravar_NFe1!CbdvAliqProd_cofins = TBProduto!Cofins_Prod
                    TBGravar_NFe1.Update
                    TBGravar_NFe1.Close
                End If
                TBCST.Close
            End If
            TBGravar_NFe.Update
            TBGravar_NFe.Close
            Contador2 = Contador2 + 1
            TBProduto.MoveNext
        Loop
    End If
    
    Contador2 = 0
    Set TBContas = CreateObject("adodb.recordset")
    TBContas.Open "Select * from tbl_Detalhes_Recebimento where ID_nota = " & ID_nota & " order by ID", Conexao, adOpenKeyset, adLockOptimistic
    If TBContas.EOF = False Then
        
        'Novo layout da Sefaz (4.0) - NÃO ACEITA SOMA DAS DUPLICATAS DIFERENTE DO VALOR TOTAL DA NOTA
        'Verifica valor total da nota e valor total a receber/pagar
'        Set TBTotaisnota = CreateObject("adodb.recordset")
'        TBTotaisnota.Open "Select * from tbl_Totais_Nota where ID_nota = " & ID_nota, Conexao, adOpenKeyset, adLockOptimistic
'        If TBTotaisnota.EOF = False Then
'            Valor = IIf(IsNull(TBTotaisnota!dbl_Valor_Total_Nota), 0, TBTotaisnota!dbl_Valor_Total_Nota)
'            ValorTotal = Valor
'            Valor1 = IIf(IsNull(TBTotaisnota!Valor_total_receber_pagar), 0, TBTotaisnota!Valor_total_receber_pagar)
'        End If
'        TBTotaisnota.Close
        
        Do While TBContas.EOF = False
            Set TBGravar_NFe = CreateObject("adodb.recordset")
            TBGravar_NFe.Open "Select * from CBD001DUPLICATAS", Conexao_NFe, adOpenKeyset, adLockOptimistic
            TBGravar_NFe.AddNew
            TBGravar_NFe!CbdEmpCodigo = TBMaquinas!CODIGO
            TBGravar_NFe!CbdNtfSerie = TBAbrir!Serie
            TBGravar_NFe!CbdNtfNumero = OF
            Contador2 = Contador2 + 1
            TBGravar_NFe!CbdDupSeq = Contador2
            TBGravar_NFe!CbdnDup = Left(TBContas!txt_Parcela, 3)
            TBGravar_NFe!CbddVenc = Format(TBContas!dt_Vencimento, "yyyy-mm-dd")
            TBGravar_NFe!CbdvDup = TBContas!dbl_Valor
'            If Contador2 = TBContas.RecordCount Then
'                TBGravar_NFe!CbdvDup = ValorTotal
'            Else
'                'Verifica percentual da parcela
'                Valor3 = (TBContas!dbl_Valor / Valor1) * 100
'
'                TBGravar_NFe!CbdvDup = Format((Valor * Valor3) / 100, "0.00")
'                ValorTotal = Format(ValorTotal - TBGravar_NFe!CbdvDup, "0.00")
'            End If
           
            TBGravar_NFe.Update
            TBGravar_NFe.Close
            TBContas.MoveNext
        Loop
    End If
    TBContas.Close
    
    Contador2 = 0
    Set TBCiclo = CreateObject("adodb.recordset")
    TBCiclo.Open "Select ID_nota_relacionada AS ID from Faturamento_Relacionamento where ID_nota = " & ID_nota & " group by ID_nota_relacionada", Conexao, adOpenKeyset, adLockOptimistic
    If TBCiclo.EOF = True Then
        Set TBCiclo = CreateObject("adodb.recordset")
        TBCiclo.Open "Select ID_nota AS ID from Faturamento_Relacionamento where ID_nota_relacionada = " & ID_nota & " group by ID_nota", Conexao, adOpenKeyset, adLockOptimistic
    End If
    Do While TBCiclo.EOF = False
        Set TBGravar_NFe = CreateObject("adodb.recordset")
        TBGravar_NFe.Open "Select * from CBD001NREF", Conexao_NFe, adOpenKeyset, adLockOptimistic
        TBGravar_NFe.AddNew
        TBGravar_NFe!CbdEmpCodigo = TBMaquinas!CODIGO
        TBGravar_NFe!CbdNtfSerie = TBAbrir!Serie
        TBGravar_NFe!CbdNtfNumero = OF
        Contador2 = Contador2 + 1
        TBGravar_NFe!cbdrefSeq = Contador2
        Set TBCarteira = CreateObject("adodb.recordset")
        TBCarteira.Open "Select ID, int_NotaFiscal, txt_Municipio, txt_UF, dt_DataEmissao, txt_CNPJ_CPF, Modelo, Serie from tbl_Dados_Nota_Fiscal where ID = " & TBCiclo!ID, Conexao, adOpenKeyset, adLockOptimistic
        If TBCarteira.EOF = False Then
            Set TBTempo = CreateObject("adodb.recordset")
            TBTempo.Open "Select Chave_acesso from tbl_Dados_Nota_Fiscal_NFe where ID_nota = " & TBCarteira!ID & " and Chave_acesso IS NOT NULL and Chave_acesso <> N''", Conexao, adOpenKeyset, adLockOptimistic
            If TBTempo.EOF = False Then
                TBGravar_NFe!CbdrefNFe = TBTempo!Chave_acesso
            Else
                IDpedido = TBCarteira!int_NotaFiscal
                FamiliaAntiga = DS_RemoverAcentos(TBCarteira!txt_Municipio)
                TBGravar_NFe!CbdcUF_refNFE = FunVerificaCodUF(FamiliaAntiga, TBCarteira!txt_UF)
                TBGravar_NFe!CbdAAMM = Format(TBCarteira!dt_DataEmissao, "YYMM")
                TBGravar_NFe!CbdCNPJ = DS_RetornarNumeros(TBCarteira!txt_CNPJ_CPF)
                TBGravar_NFe!Cbdmod_refNFE = IIf(Left(TBCarteira!Modelo, 2) = "1B", 1, Left(TBCarteira!Modelo, 2))
                TBGravar_NFe!Cbdserie_refNFE = TBCarteira!Serie
                TBGravar_NFe!CbdnNF_refNFE = IDpedido
            End If
            TBTempo.Close
        End If
        TBCarteira.Close
        TBGravar_NFe.Update
        TBGravar_NFe.Close
        TBCiclo.MoveNext
    Loop
    TBCiclo.Close
    
    Set TBTransporte = CreateObject("adodb.recordset")
    TBTransporte.Open "Select * from tbl_Dados_Transp where ID_Nota = " & ID_nota, Conexao, adOpenKeyset, adLockOptimistic
    If TBTransporte.EOF = False Then
        Set TBGravar_NFe = CreateObject("adodb.recordset")
        TBGravar_NFe.Open "Select * from CBD001REBOQUE", Conexao_NFe, adOpenKeyset, adLockOptimistic
        TBGravar_NFe.AddNew
        TBGravar_NFe!CbdEmpCodigo = TBMaquinas!CODIGO
        TBGravar_NFe!CbdNtfSerie = TBAbrir!Serie
        TBGravar_NFe!CbdNtfNumero = OF
        TBGravar_NFe!Cbdplaca_rebtransp = TBTransporte!txt_Placa
        TBGravar_NFe!CbdUF_rebtransp = TBTransporte!txt_UF
        TBGravar_NFe.Update
        TBGravar_NFe.Close
    End If
    TBTransporte.Close
    
    Set TBGravar_NFe = CreateObject("adodb.recordset")
    TBGravar_NFe.Open "Select * from NFE012", Conexao_NFe, adOpenKeyset, adLockOptimistic
    TBGravar_NFe.AddNew
    TBGravar_NFe!CbdEmpCodigo = TBMaquinas!CODIGO
    TBGravar_NFe!CbdNtfSerie = TBAbrir!Serie
    TBGravar_NFe!CbdNtfNumero = OF
    TBGravar_NFe!cbdAcao = "E"
    TBGravar_NFe!CbdSituacao = 0
    TBGravar_NFe!CbdProcStatus = "N"
    TBGravar_NFe!CbdNFEChaAcesso = ""
    TBGravar_NFe.Update
    TBGravar_NFe.Close
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Private Sub ProcExcluirProduto()
On Error GoTo tratar_erro

If Excluir = False Then
    MsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso.")
    Exit Sub
End If
Permitido = False
With listaProdutos
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If MsgBox("Deseja realmente excluir estes dados do(s) produto(s)/item(ns) da NFe?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
            End If
            Permitido = True
            
            Conexao.Execute "UPDATE tbl_Detalhes_Nota_NFe Set Codigo_ANP = Null, UF_consumo = Null, Tipo_produto = Null WHERE ID_item = " & .ListItems(InitFor)
            Conexao.Execute "UPDATE tbl_Detalhes_Nota_CST_ICMS Set Modalidade_determinacao = Null, Modalidade_determinacao_ST = Null WHERE id_item = " & .ListItems(InitFor)

            '==================================
            Modulo = Formulario
            Evento = "Excluir dados do produto da nota fiscal"
            ID_documento = .ListItems(InitFor)
            
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select * from tbl_Dados_nota_fiscal WHERE id = " & txtID_nota, Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                If IsNull(TBAbrir!int_NotaFiscal) = True Or TBAbrir!int_NotaFiscal = "" Then NomeCampo = "N° ordem: " & TBAbrir!ID Else NomeCampo = "N° nota: " & TBAbrir!int_NotaFiscal
                Documento = NomeCampo & " - Tipo: " & TBAbrir!TipoNF & " - Série: " & TBAbrir!Serie
            End If
            TBAbrir.Close
            Documento1 = "Cód. interno: " & .ListItems(InitFor).ListSubItems(1)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    MsgBox ("Informe o(s) produto(s)/item(ns) da nota fiscal antes de excluir estes dados."), vbExclamation
Else
    MsgBox ("Dados do(s) produto(s)/item(ns) da nota fiscal excluído(s) com sucesso."), vbInformation
    ProcLimpacamposProdutos
    ProcCarregaListaProdutos
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Private Sub cmbFinalidade_emissao_Click()
On Error GoTo tratar_erro

With Cmb_presenca_comprador
    If Left(cmbFinalidade_emissao, 1) = 2 Or Left(cmbFinalidade_emissao, 1) = 3 Then
        .Text = "0 - Não se aplica"
        .Locked = True
        .TabStop = False
    Else
        .ListIndex = -1
        .Locked = False
        .TabStop = True
    End If
End With

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Private Sub cmdPagAnt_Click()
On Error GoTo tratar_erro

If DS_RetornarNumeros(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Faturamento_NFe.AbsolutePage <> 2 Then
    If TBLISTA_Faturamento_NFe.AbsolutePage = -3 Then
        ProcExibePagina (TBLISTA_Faturamento_NFe.PageCount - 1)
    Else
        TBLISTA_Faturamento_NFe.AbsolutePage = TBLISTA_Faturamento_NFe.AbsolutePage - 2
        ProcExibePagina (TBLISTA_Faturamento_NFe.AbsolutePage)
    End If
Else
    ProcExibePagina (1)
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Private Sub cmdPagIr_Click()
On Error GoTo tratar_erro

If txtPagIr = "" Then Exit Sub
Quant = DS_RetornarNumeros(Right(lblPaginas.Caption, 4))
If Quant <= 1 Or txtPagIr > Quant Then Exit Sub
If txtPagIr.Text >= 1 And txtPagIr.Text <= Quant Then
    TBLISTA_Faturamento_NFe.AbsolutePage = txtPagIr.Text
    ProcExibePagina (TBLISTA_Faturamento_NFe.AbsolutePage)
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Private Sub cmdPagPrim_Click()
On Error GoTo tratar_erro

If DS_RetornarNumeros(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Faturamento_NFe.AbsolutePage = 1
ProcExibePagina (TBLISTA_Faturamento_NFe.AbsolutePage)

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Private Sub cmdPagProx_Click()
On Error GoTo tratar_erro

If DS_RetornarNumeros(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Faturamento_NFe.AbsolutePage <> -3 Then
    If TBLISTA_Faturamento_NFe.AbsolutePage = 1 Then
        ProcExibePagina (2)
    Else
        ProcExibePagina (TBLISTA_Faturamento_NFe.AbsolutePage)
    End If
Else
    ProcExibePagina (TBLISTA_Faturamento_NFe.PageCount)
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Private Sub cmdPagUlt_Click()
On Error GoTo tratar_erro

If DS_RetornarNumeros(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Faturamento_NFe.AbsolutePage = TBLISTA_Faturamento_NFe.PageCount
ProcExibePagina (TBLISTA_Faturamento_NFe.AbsolutePage)

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15195, 9, True
ProcCarregaToolBar2 Me, 15195, 6, True

If Formulario = "Faturamento/Nota fiscal/Própria" Then
    Caption = "Administrativo - Faturamento - Nota fiscal - Própria - Dados da NFe"
ElseIf Formulario = "Faturamento/Nota fiscal/Terceiros" Then
        Caption = "Administrativo - Faturamento - Nota fiscal - Terceiros - Dados da NFe"
    ElseIf Formulario = "Estoque/Ordem de faturamento" Then
            Caption = "Estoque - Ordem de faturamento - Dados da NFe"
        Else
            Caption = "Estoque - Nota fiscal - Dados da NFe"
End If
ProcLimpaVariaveisPrincipais
ProcRemoveObjetosResize Me

Cmb_opcao_lista = "Liberar"
ProcCarregaListaNota (1)
ProcLimpaCampos
With frmFaturamento_Prod_Serv
    If .txtid <> "" And .txtid <> "0" And .txtDtValidacao <> "" Then
        txtID_nota = .txtid
        txtNota = IIf(.txtNFiscal = "", Null, .txtNFiscal)
        txtSerie = .txtSerie
        ProcCarregaEntrega
        ProcCarregaCobranca
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from tbl_Dados_Nota_Fiscal_NFe where id_nota = " & txtID_nota, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            ProcPuxaDados
        End If
        TBAbrir.Close
        
        UF_transp = ""
        Cidade = ""
        Set TBFI = CreateObject("adodb.recordset")
        TBFI.Open "Select Empresa.* from Empresa INNER JOIN tbl_Dados_Nota_Fiscal ON Empresa.Codigo = tbl_Dados_Nota_Fiscal.ID_empresa where tbl_Dados_Nota_Fiscal.ID = " & txtID_nota, Conexao, adOpenKeyset, adLockOptimistic
        If TBFI.EOF = False Then
            UF_transp = IIf(IsNull(TBFI!UF), "", TBFI!UF)
            Cidade = IIf(IsNull(TBFI!Cidade), "", TBFI!Cidade)
        End If
        
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "select * from tbl_dados_transp where ID_Nota = " & txtID_nota, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Frame2.Enabled = True
            If IsNull(TBAbrir!UF_embarque) = False Then
                cmbUF_embarque = TBAbrir!UF_embarque
            Else
                If UF_transp <> "" Then cmbUF_embarque = UF_transp
            End If
            txtLocal_embarque = IIf(IsNull(TBAbrir!Local_embarque), Cidade, TBAbrir!Local_embarque)
        Else
            Frame2.Enabled = False
        End If
        TBAbrir.Close
    End If
End With
SStab_nfe.Tab = 0
Timer_status_NFe.Enabled = True

With Cmb_codigo_ANP
    .Clear
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from Codigos_produtos_ANP", Conexao, adOpenKeyset, adLockReadOnly
    If TBAbrir.EOF = False Then
        Do While TBAbrir.EOF = False
            .AddItem TBAbrir!CODIGO
            TBAbrir.MoveNext
        Loop
    End If
    TBAbrir.Close
End With
ProcCarregaComboUF Cmb_UF_consumo, "UF is not null", ""

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Private Sub ProcCancelar()
On Error GoTo tratar_erro

If Excluir = False Then
    MsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso.")
    Exit Sub
End If
Permitido = False
With ListaNota
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If MsgBox("Deseja realmente cancelar a liberação desta(s) nota(s) fiscal(ais)?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
            End If
            Permitido = True
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select * from tbl_Dados_Nota_Fiscal WHERE ID = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False Then
                TBFI!Imprimir = False
                TBFI.Update
                
                If frmFaturamento_Prod_Serv.txtid = TBFI!ID Then frmFaturamento_Prod_Serv.NFe_liberada = False
                
                '==================================
                Modulo = Formulario
                Evento = "Cancelar liberação da nota fiscal"
                ID_documento = .ListItems(InitFor)
                If IsNull(TBFI!int_NotaFiscal) = True Or TBFI!int_NotaFiscal = "" Then NomeCampo = "N° ordem: " & TBFI!ID Else NomeCampo = "N° nota: " & TBFI!int_NotaFiscal
                Documento = NomeCampo & " - Tipo: " & TBFI!TipoNF & " - Série: " & TBFI!Serie
                Documento1 = ""
                ProcGravaEvento
                '==================================
                
                Conexao.Execute "UPDATE tbl_Dados_Nota_Fiscal_NFe Set Status = Null, Chave_acesso = Null WHERE id_nota = " & .ListItems(InitFor)
                
                OF = IIf(IsNull(TBFI!int_NotaFiscal), 0, TBFI!int_NotaFiscal)
                Set TBMaquinas = CreateObject("adodb.recordset")
                TBMaquinas.Open "Select * from Empresa where Codigo = " & TBFI!ID_empresa & " and GNFe = 'True'", Conexao, adOpenKeyset, adLockOptimistic
                If TBMaquinas.EOF = False Then
                    caminho = TBMaquinas!Caminho_Nfe & "\Empresa " & TBFI!ID_empresa & " - Serie " & TBFI!Serie & " - Nota " & OF & " - Status E.bat"
                    Caminho1 = TBMaquinas!Caminho_Nfe & "\Empresa " & TBFI!ID_empresa & " - Serie " & TBFI!Serie & " - Nota " & OF & " - Status C.bat"
                    Set GerArqPastas = CreateObject("Scripting.FileSystemObject")
                    If GerArqPastas.FileExists(caminho) = True Then Kill caminho
                    If GerArqPastas.FileExists(Caminho1) = True Then Kill Caminho1
                End If
                TBMaquinas.Close
                
                ProcExcluirDadosTabelaGNFe OF, TBFI!Serie
                
'                Set TBProduto = CreateObject("adodb.recordset")
'                TBProduto.Open "Select * from tbl_dados_transp Where id_nota = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
'                If TBProduto.EOF = False Then
'                    TBProduto!UF_embarque = Null
'                    TBProduto!Local_embarque = Null
'                    TBProduto.Update
'                End If
'                TBProduto.Close
            End If
            TBFI.Close
        End If
    Next InitFor
End With
If Permitido = False Then
    MsgBox ("Informe a(s) nota(s) fiscal(ais) antes de cancelar a liberação."), vbExclamation
Else
    MsgBox ("Liberação da(s) nota(s) fiscal(ais) canceladas(s) com sucesso."), vbInformation
    ProcLimpaCampos
    txtNota = ""
    txtSerie = ""
    ProcCarregaListaNota (IIf(DS_RetornarNumeros(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, DS_RetornarNumeros(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
    With frmFaturamento_Prod_Serv
        .ProcCarregaListaNota (IIf(DS_RetornarNumeros(Left(.lblPaginas(1).Caption, Len(.lblPaginas(1).Caption) - 5)) <= 1, 1, DS_RetornarNumeros(Left(.lblPaginas(1).Caption, Len(.lblPaginas(1).Caption) - 5))))
    End With
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Private Sub ProcExcluirDadosTabelaGNFe(Nota As Long, Serie As String)
On Error GoTo tratar_erro

Conexao_NFe.Execute "DELETE from CBD001 where CbdNtfNumero = " & Nota & " and CbdNtfSerie = '" & Serie & "'"
Conexao_NFe.Execute "DELETE from CBD001AUTXML where CbdNtfNumero = " & Nota & " and CbdNtfSerie = '" & Serie & "'"
Conexao_NFe.Execute "DELETE from CBD001DET where CbdNtfNumero = " & Nota & " and CbdNtfSerie = '" & Serie & "'"
Conexao_NFe.Execute "DELETE from CBD001DETADICOES where CbdNtfNumero = " & Nota & " and CbdNtfSerie = '" & Serie & "'"
Conexao_NFe.Execute "DELETE from CBD001DETCOFINS where CbdNtfNumero = " & Nota & " and CbdNtfSerie = '" & Serie & "'"
Conexao_NFe.Execute "DELETE from CBD001DETDI where CbdNtfNumero = " & Nota & " and CbdNtfSerie = '" & Serie & "'"
Conexao_NFe.Execute "DELETE from CBD001DETEXPORT where CbdNtfNumero = " & Nota & " and CbdNtfSerie = '" & Serie & "'"
Conexao_NFe.Execute "DELETE from CBD001DETICMSNORMALST where CbdNtfNumero = " & Nota & " and CbdNtfSerie = '" & Serie & "'"
Conexao_NFe.Execute "DELETE from CBD001DETIPI where CbdNtfNumero = " & Nota & " and CbdNtfSerie = '" & Serie & "'"
Conexao_NFe.Execute "DELETE from CBD001DETPIS where CbdNtfNumero = " & Nota & " and CbdNtfSerie = '" & Serie & "'"
Conexao_NFe.Execute "DELETE from CBD001DUPLICATAS where CbdNtfNumero = " & Nota & " and CbdNtfSerie = '" & Serie & "'"
Conexao_NFe.Execute "DELETE from CBD001NREF where CbdNtfNumero = " & Nota & " and CbdNtfSerie = '" & Serie & "'"
Conexao_NFe.Execute "DELETE from CBD001PROCREF where CbdNtfNumero = " & Nota & " and CbdNtfSerie = '" & Serie & "'"
Conexao_NFe.Execute "DELETE from CBD001REBOQUE where CbdNtfNumero = " & Nota & " and CbdNtfSerie = '" & Serie & "'"
Conexao_NFe.Execute "DELETE from CBD001VOL where CbdNtfNumero = " & Nota & " and CbdNtfSerie = '" & Serie & "'"
Conexao_NFe.Execute "DELETE from CBD001PAG where CbdNtfNumero = " & Nota & " and CbdNtfSerie = '" & Serie & "'"

Conexao_NFe.Execute "DELETE from NFE001 where NtfNumero = " & Nota & " and NtfSerie = '" & Serie & "'"
Conexao_NFe.Execute "DELETE from NFE001AUTXML where NtfNumero = " & Nota & " and NtfSerie = '" & Serie & "'"
Conexao_NFe.Execute "DELETE from NFE001DET where NtfNumero = " & Nota & " and NtfSerie = '" & Serie & "'"
Conexao_NFe.Execute "DELETE from NFE001DETADICOES where NtfNumero = " & Nota & " and NtfSerie = '" & Serie & "'"
Conexao_NFe.Execute "DELETE from NFE001DETCOFINS where NtfNumero = " & Nota & " and NtfSerie = '" & Serie & "'"
Conexao_NFe.Execute "DELETE from NFE001DETDI where NtfNumero = " & Nota & " and NtfSerie = '" & Serie & "'"
Conexao_NFe.Execute "DELETE from NFE001DETEXPORT where NtfNumero = " & Nota & " and NtfSerie = '" & Serie & "'"
Conexao_NFe.Execute "DELETE from NFE001DETIPI where NtfNumero = " & Nota & " and NtfSerie = '" & Serie & "'"
Conexao_NFe.Execute "DELETE from NFE001DETPIS where NtfNumero = " & Nota & " and NtfSerie = '" & Serie & "'"
Conexao_NFe.Execute "DELETE from NFE001DUPLICATAS where NtfNumero = " & Nota & " and NtfSerie = '" & Serie & "'"
Conexao_NFe.Execute "DELETE from NFE001ICMS_NORMAL_E_ST where NtfNumero = " & Nota & " and NtfSerie = '" & Serie & "'"
Conexao_NFe.Execute "DELETE from NFE001NFREF where NtfNumero = " & Nota & " and NtfSerie = '" & Serie & "'"
Conexao_NFe.Execute "DELETE from NFE001PROCREF where NtfNumero = " & Nota & " and NtfSerie = '" & Serie & "'"
Conexao_NFe.Execute "DELETE from NFE001PROCREF where NtfNumero = " & Nota & " and NtfSerie = '" & Serie & "'"
Conexao_NFe.Execute "DELETE from NFE001REBOQUE where NtfNumero = " & Nota & " and NtfSerie = '" & Serie & "'"
Conexao_NFe.Execute "DELETE from NFE001VOL where NtfNumero = " & Nota & " and NtfSerie = '" & Serie & "'"
Conexao_NFe.Execute "DELETE from NFE010 where NtfNumero = " & Nota & " and NtfSerie = '" & Serie & "'"
Conexao_NFe.Execute "DELETE from NFE010LOGSTATUS where NtfNumero = " & Nota & " and NtfSerie = '" & Serie & "'"
Conexao_NFe.Execute "DELETE from NFE012 where CbdNtfNumero = " & Nota & " and CbdNtfSerie = '" & Serie & "'"
Conexao_NFe.Execute "DELETE from NFE012LOGSTATUS where CbdNtfNumero = " & Nota & " and CbdNtfSerie = '" & Serie & "'"
Conexao_NFe.Execute "DELETE from NFE001PAG where NtfNumero = " & Nota & " and NtfSerie = '" & Serie & "'"

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Private Sub ProcVerificaStatus(ID_nota As Long)
On Error GoTo tratar_erro

Permitido = True
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select ID_empresa, int_NotaFiscal, Serie from tbl_dados_nota_fiscal where id = " & ID_nota, Conexao, adOpenKeyset, adLockReadOnly
If TBAbrir.EOF = False Then
    OF = TBAbrir!int_NotaFiscal
    Set TBGravar_NFe_Status = CreateObject("adodb.recordset")
    TBGravar_NFe_Status.Open "Select CbdStsRetCodigo from NFE012 where CbdEmpCodigo = " & TBAbrir!ID_empresa & " and CbdNtfNumero = " & OF & " and CbdNtfSerie = '" & TBAbrir!Serie & "' order by CbdNtfNumero, CbdNtfSerie", Conexao_NFe, adOpenKeyset, adLockReadOnly
    If TBGravar_NFe_Status.EOF = False Then
        If TBGravar_NFe_Status!CbdStsRetCodigo = 100 Or TBGravar_NFe_Status!CbdStsRetCodigo = 101 Then
            Permitido = False
            Select Case TBGravar_NFe_Status!CbdStsRetCodigo
                Case "100": NomeCampo = "autorizada no SEFAZ" 'Autorizado o uso da NF-e
                Case "101": NomeCampo = "cancelada no SEFAZ" 'Cancelamento de NF-e homologado"
            End Select
        End If
    End If
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

Timer_status_NFe.Enabled = False
Unload Me

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case SStab_nfe.Tab
    Case 0:
        Select Case KeyCode
            Case vbKeyF2: ProcFiltrar
            Case vbKeyF3: ProcSalvar
            Case vbKeyF4: ProcCancelar
            Case vbKeyF7: ProcLiberar
            Case vbKeyF8: ProcAtualizarStatus
            Case vbKeyEscape: ProcSair
        End Select
    Case 1:
        Select Case KeyCode
            Case vbKeyF3: ProcSalvarProduto
            Case vbKeyF4: ProcExcluirProduto
            Case vbKeyEscape: ProcSair
        End Select
End Select

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

ProcCarregaListaNota (1)

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Private Sub ProcSalvar()
On Error GoTo tratar_erro
  
If Alterar = False Then
    MsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation
    Exit Sub
End If
Acao = "salvar"
If txtID_nota = 0 Then
    MsgBox ("Informe a nota antes de salvar."), vbExclamation
    Exit Sub
End If
If txtStatus <> "" Then
    MsgBox ("Não é permitido salvar, pois esta nota fiscal já foi liberada para envio."), vbExclamation
    Exit Sub
End If

'Verifica se é NFSe
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * FROM tbl_Dados_Nota_Fiscal WHERE ID = " & txtID_nota & " and TipoNF = 'SA'", Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = False Then
    MsgBox ("Não é permitido salvar, pois esta é uma nota fiscal de serviço."), vbExclamation
    TBGravar.Close
    Exit Sub
End If
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * From tbl_Dados_Nota_Fiscal where id = " & txtID_nota & " and DtValidacao IS NULL", Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = False Then
    MsgBox ("Não é permitido salvar, pois esta nota fiscal ainda não foi validada."), vbExclamation
    TBGravar.Close
    Exit Sub
End If
TBGravar.Close

If Cmb_forma_de_emissao = "" Then
    NomeCampo = "a forma de emissão"
    ProcVerificaAcao
    Cmb_forma_de_emissao.SetFocus
    Exit Sub
End If
If cmbFinalidade_emissao = "" Then
    NomeCampo = "a finalidade de emissão"
    ProcVerificaAcao
    cmbFinalidade_emissao.SetFocus
    Exit Sub
End If
If cmbFormaPag = "" Then
    NomeCampo = "a forma de pagamento"
    ProcVerificaAcao
    cmbFormaPag.SetFocus
    Exit Sub
End If
If Cmb_consumidor = "" Then
    NomeCampo = "o tipo do consumidor"
    ProcVerificaAcao
    Cmb_consumidor.SetFocus
    Exit Sub
End If
If Cmb_presenca_comprador = "" Then
    NomeCampo = "a presença do comprador"
    ProcVerificaAcao
    Cmb_presenca_comprador.SetFocus
    Exit Sub
End If

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * FROM tbl_Dados_Nota_Fiscal where ID = " & txtID_nota & " and txt_tipocliente <> 'E'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    Endereco_NF = IIf(IsNull(TBAbrir!txt_Endereco), "", TBAbrir!txt_Endereco) & " - " & IIf(IsNull(TBAbrir!Numero), "", TBAbrir!Numero) & " - " & IIf(IsNull(TBAbrir!Txt_bairro), "", TBAbrir!Txt_bairro) & " - " & IIf(IsNull(TBAbrir!txt_Municipio), "", TBAbrir!txt_Municipio) & " - " & IIf(IsNull(TBAbrir!txt_UF), "", TBAbrir!txt_UF) & " - " & IIf(IsNull(TBAbrir!Txt_CEP), "", TBAbrir!Txt_CEP)
    If cmbEntrega = "" Then
        NomeCampo = "o endereço de entrega"
        ProcVerificaAcao
        cmbEntrega.SetFocus
        Exit Sub
    Else
        If cmbEntrega <> Endereco_NF And Chk_DA_entrega.Value = 0 Then
            If MsgBox("O endereço de entrega é diferente do endereço principal, deseja prosseguir sem imprimir o endereço de entrega nos dados adicionais?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        ElseIf cmbEntrega = Endereco_NF And Chk_DA_entrega.Value = 1 Then
                If MsgBox("O endereço de entrega é igual o endereço principal, deseja prosseguir imprimindo o endereço de entrega nos dados adicionais?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        End If
    End If
    If Cmb_cobranca = "" Then
        NomeCampo = "o endereço de cobrança"
        ProcVerificaAcao
        Cmb_cobranca.SetFocus
        Exit Sub
    Else
        If Cmb_cobranca <> Endereco_NF And Chk_DA_cobranca.Value = 0 Then
            If MsgBox("O endereço de cobrança é diferente do endereço principal, deseja prosseguir sem imprimir o endereço de cobrança nos dados adicionais?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        ElseIf Cmb_cobranca = Endereco_NF And Chk_DA_cobranca.Value = 1 Then
                If MsgBox("O endereço de cobrança é igual o endereço principal, deseja prosseguir imprimindo o endereço de cobrança nos dados adicionais?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        End If
    End If
End If
TBAbrir.Close

If Frame2.Enabled = True Then
    If txtLocal_embarque = "" Then
        NomeCampo = "o local de embarque"
        ProcVerificaAcao
        txtLocal_embarque.SetFocus
        Exit Sub
    End If
    If cmbUF_embarque = "" Then
        NomeCampo = "o UF"
        ProcVerificaAcao
        cmbUF_embarque.SetFocus
        Exit Sub
    End If
End If
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * FROM tbl_Dados_Nota_Fiscal_NFe WHERE ID_nota = " & txtID_nota, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = False Then
    MsgBox ("Alteração efetuada com sucesso."), vbInformation
    Evento = "Alterar dados da nota fiscal"
Else
    TBGravar.AddNew
    MsgBox ("Novos dados da nota fiscal cadastrados com sucesso."), vbInformation
    Evento = "Novos dados da nota fiscal"
    TBGravar!ID_nota = txtID_nota
End If
ProcEnviaDados
TBGravar.Update
TBGravar.Close
'==================================
Modulo = Formulario
ID_documento = txtID_nota
With frmFaturamento_Prod_Serv
    .ProcVerificaTipoNF False
    If .txtNFiscal = "" Then NomeCampo = "N° ordem: " & .txtid Else NomeCampo = "N° nota: " & .txtNFiscal
    Documento = NomeCampo & " - Tipo: " & TipoNF & " - Série: " & .txtSerie
End With
Documento1 = ""
ProcGravaEvento
'==================================

Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select * from tbl_dados_transp Where id_nota = " & txtID_nota, Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    TBProduto!UF_embarque = cmbUF_embarque
    TBProduto!Local_embarque = txtLocal_embarque
    '==================================
    Evento = "Alterar transportadora"
    ID_documento = TBProduto!ID
    With frmFaturamento_Prod_Serv
        .ProcVerificaTipoNF False
        If .txtNFiscal = "" Then NomeCampo = "N° ordem: " & .txtid Else NomeCampo = "N° nota: " & .txtNFiscal
        Documento = NomeCampo & " - Tipo: " & TipoNF & " - Série: " & .txtSerie
    End With
    Documento1 = "Transportadora: " & TBProduto!txt_Razao
    ProcGravaEvento
    '==================================
    TBProduto.Update
End If
TBProduto.Close
ProcCarregaListaNota (IIf(DS_RetornarNumeros(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, DS_RetornarNumeros(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
If CodigoLista <> 0 And ListaNota.ListItems.Count <> 0 Then
    ListaNota.SelectedItem = ListaNota.ListItems(CodigoLista)
    ListaNota.SetFocus
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Sub ProcEnviaDados()
On Error GoTo tratar_erro

If cmbForma_pagamento <> "" Then TBGravar!Forma_pagamento = Left(cmbForma_pagamento, 1) Else TBGravar!Forma_pagamento = Null
TBGravar!FormaPagto = Left(cmbFormaPag, 2)
TBGravar!Forma_emissao = Left(Cmb_forma_de_emissao, 1)
TBGravar!Finalidade_emissao = Left(cmbFinalidade_emissao, 1)
TBGravar!Consumidor_final = Left(Cmb_consumidor, 1)
TBGravar!Presenca_comprador = Left(Cmb_presenca_comprador, 1)
TBGravar!Enviar_Email = Left(Cmb_arquivos_XML_enviados_email, 1)
TBGravar!Titulo_canhoto_DANFE = Txt_titulo_canhoto_DANFE
TBGravar!Texto_canhoto_DANFE = Txt_titulo_canhoto_DANFE

TBGravar!ID_entrega = txtID_entrega
If Chk_DA_entrega.Value = 1 Then TBGravar!DA_entrega = True Else TBGravar!DA_entrega = False
TBGravar!ID_Cobranca = Txt_ID_cobranca
If Chk_DA_cobranca.Value = 1 Then TBGravar!DA_cobranca = True Else TBGravar!DA_cobranca = False

TBGravar!Enviar_DANFE_email = IIf(Cmb_enviar_DANFE = "", "S", Left(Cmb_enviar_DANFE, 1))
TBGravar!status = Null
If chkCodRef.Value = 1 Then TBGravar!CodRef = True Else TBGravar!CodRef = False

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

Timer_status_NFe.Enabled = True

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Private Sub ListaNota_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With ListaNota
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If Cmb_opcao_lista = "Cancelar liberação" Then
                    ProcVerificaStatus .ListItems(InitFor)
                    If Permitido = False Then
                        .ListItems.Item(InitFor).Checked = False
                        GoTo Proximo
                    End If
                Else
                    Set TBAbrir = CreateObject("adodb.recordset")
                    TBAbrir.Open "Select * from tbl_dados_nota_fiscal_nfe where ID_nota = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
                    If TBAbrir.EOF = False Then
                        If TBAbrir!status <> "" Or IsNull(TBAbrir!status) = False Then
                            .ListItems.Item(InitFor).Checked = False
                            TBAbrir.Close
                            GoTo Proximo
                        End If
                    End If
                    TBAbrir.Close
                    ProcVerifLiberacao .ListItems(InitFor), .ListItems(InitFor).ListSubItems(1), False
                    If Permitido = False Then
                        .ListItems.Item(InitFor).Checked = False
                        GoTo Proximo
                    End If
                    'Verifica se a cidade está cadastrada corretamente
                    Set TBAbrir = CreateObject("adodb.recordset")
                    TBAbrir.Open "Select * from tbl_Dados_Nota_Fiscal where ID = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
                    If TBAbrir.EOF = False Then
                        FamiliaAntiga = DS_RemoverAcentos(TBAbrir!txt_Municipio)
                        
                        If IsNull(TBAbrir!txt_UF) = False And TBAbrir!txt_UF <> "" And TBAbrir!txt_UF <> "EX" Then
                            Set TBFI = CreateObject("adodb.recordset")
                            TBFI.Open "Select * from CEP where Municipio = '" & FamiliaAntiga & "'", Conexao, adOpenKeyset, adLockOptimistic
                            If TBFI.EOF = True Then
                                .ListItems.Item(InitFor).Checked = False
                                TBFI.Close
                                GoTo Proximo
                            End If
                            Set TBFI = CreateObject("adodb.recordset")
                            TBFI.Open "Select * from CEP where Sigla_UF = '" & TBAbrir!txt_UF & "'", Conexao, adOpenKeyset, adLockOptimistic
                            If TBFI.EOF = True Then
                                .ListItems.Item(InitFor).Checked = False
                                TBFI.Close
                                GoTo Proximo
                            End If
                            Set TBFI = CreateObject("adodb.recordset")
                            TBFI.Open "Select * from CEP where Municipio = '" & FamiliaAntiga & "' and Sigla_UF = '" & TBAbrir!txt_UF & "'", Conexao, adOpenKeyset, adLockOptimistic
                            If TBFI.EOF = True Then
                                .ListItems.Item(InitFor).Checked = False
                                TBFI.Close
                                GoTo Proximo
                            End If
                        End If
                        
                        'Verifica se tem país cadastrado
                        Set TBClientes = CreateObject("adodb.recordset")
                        If TBAbrir!txt_tipocliente = "E" Then
                            'Empresa
                            TBClientes.Open "Select * from Empresa where Codigo = " & TBAbrir!Id_Int_Cliente, Conexao, adOpenKeyset, adLockOptimistic
                            If TBClientes.EOF = False Then
                                If IsNull(TBClientes!Codigo_pais) = True Or TBClientes!Codigo_pais = "" Then
                                    .ListItems.Item(InitFor).Checked = False
                                    TBClientes.Close
                                    GoTo Proximo
                                End If
                            End If
                        ElseIf TBAbrir!txt_tipocliente = "JP" Or TBAbrir!txt_tipocliente = "JR" Or TBAbrir!txt_tipocliente = "FP" Or TBAbrir!txt_tipocliente = "FR" Then
                                'Cliente
                                TBClientes.Open "Select * from Clientes where IDcliente = " & TBAbrir!Id_Int_Cliente & " and NomeRazao = '" & TBAbrir!txt_Razao_Nome & "'", Conexao, adOpenKeyset, adLockOptimistic
                                If TBClientes.EOF = False Then
                                    If IsNull(TBClientes!Codigo_pais) = True Or TBClientes!Codigo_pais = "" Then
                                        .ListItems.Item(InitFor).Checked = False
                                        TBClientes.Close
                                        GoTo Proximo
                                    End If
                                End If
                            Else
                                'Fornecedor
                                Set TBClientes = CreateObject("adodb.recordset")
                                TBClientes.Open "Select * from Compras_fornecedores where IDcliente = " & TBAbrir!Id_Int_Cliente & " and Nome_Razao = '" & TBAbrir!txt_Razao_Nome & "'", Conexao, adOpenKeyset, adLockOptimistic
                                If TBClientes.EOF = False Then
                                    If IsNull(TBClientes!Codigo_pais) = True Or TBClientes!Codigo_pais = "" Then
                                        .ListItems.Item(InitFor).Checked = False
                                        TBClientes.Close
                                        GoTo Proximo
                                    End If
                                End If
                            End If
                        TBClientes.Close
                    End If
                    TBAbrir.Close
                    
                    Set TBMaquinas = CreateObject("adodb.recordset")
                    TBMaquinas.Open "Select * from Empresa where Empresa = '" & ListaNota.SelectedItem.ListSubItems(1) & "' and GNFe = 'True'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBMaquinas.EOF = False Then
                        'Verifica se esta preenchido o caminho para salvar o arquivo de envio da NFe
                        If IsNull(TBMaquinas!Caminho_Nfe) = True Or TBMaquinas!Caminho_Nfe = "" Then
                            .ListItems.Item(InitFor).Checked = False
                            TBMaquinas.Close
                            GoTo Proximo
                        End If
                        'Verificar se o caminho existe
                        If GerArqPastas.FolderExists(TBMaquinas!Caminho_Nfe) = False Then
                            .ListItems.Item(InitFor).Checked = False
                            TBMaquinas.Close
                            GoTo Proximo
                        End If
                    End If
                    TBMaquinas.Close
                    
                    'Verifica se é nota fiscal de devolução ou complementar e se esta referenciado a nota fiscal
                    Set TBMaquinas = CreateObject("adodb.recordset")
                    TBMaquinas.Open "Select NFE.ID from tbl_Dados_Nota_Fiscal_NFe NFE LEFT JOIN Faturamento_Relacionamento FR ON FR.ID_nota = NFE.ID_nota where NFE.ID_nota = " & .ListItems(InitFor) & " and NFE.Finalidade_emissao <> 1 and NFE.Finalidade_emissao <> 3 and FR.ID IS NULL", Conexao, adOpenKeyset, adLockOptimistic
                    If TBMaquinas.EOF = False Then
                        .ListItems.Item(InitFor).Checked = False
                        TBMaquinas.Close
                        GoTo Proximo
                    End If
                    TBMaquinas.Close
                    
                    'Verifica se o clinte é fisico e esta com cnpj e vice e versa
                    Set TBMaquinas = CreateObject("adodb.recordset")
                    TBMaquinas.Open "Select txt_tipocliente, txt_CNPJ_CPF from tbl_Dados_Nota_Fiscal where ID = " & .ListItems(InitFor) & " and txt_uf <> 'EX'", Conexao, adOpenKeyset, adLockReadOnly
                    If TBMaquinas.EOF = False Then
                        If Left(TBMaquinas!txt_tipocliente, 1) = "J" And Len(TBMaquinas!txt_CNPJ_CPF) < 14 Then
                            .ListItems.Item(InitFor).Checked = False
                            TBMaquinas.Close
                            GoTo Proximo
                        ElseIf Left(TBMaquinas!txt_tipocliente, 1) = "F" And Len(TBMaquinas!txt_CNPJ_CPF) > 14 Then
                            .ListItems.Item(InitFor).Checked = False
                            TBMaquinas.Close
                            GoTo Proximo
                        End If
                    End If
                    TBMaquinas.Close
                End If
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView ListaNota, ColumnHeader
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Private Sub ListaNota_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With ListaNota
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True And .ListItems.Item(InitFor) = Item Then
            If Cmb_opcao_lista = "Cancelar liberação" Then
                ProcVerificaStatus .ListItems(InitFor)
                If Permitido = False Then
                    MsgBox ("Não é permitido cancelar a liberação desta nota fiscal, pois a mesma está " & NomeCampo & "."), vbExclamation
                    .ListItems.Item(InitFor).Checked = False
                End If
            Else
                Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open "Select * from tbl_dados_nota_fiscal_nfe where ID_nota = " & Item, Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    If TBAbrir!status <> "" Or IsNull(TBAbrir!status) = False Then
                        MsgBox ("Não é permitido liberar esta nota fiscal para envio, pois a mesma já foi liberada."), vbExclamation
                        .ListItems.Item(InitFor).Checked = False
                        TBAbrir.Close
                        Exit Sub
                    End If
                End If
                ProcVerifLiberacao .ListItems(InitFor), .ListItems(InitFor).ListSubItems(1), True
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
                End If
                'Verifica se a cidade está cadastrada corretamente
                Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open "Select * from tbl_Dados_Nota_Fiscal where ID = " & Item, Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    FamiliaAntiga = DS_RemoverAcentos(TBAbrir!txt_Municipio)
                    
                    If IsNull(TBAbrir!txt_UF) = False And TBAbrir!txt_UF <> "" And TBAbrir!txt_UF <> "EX" Then
                        Set TBFI = CreateObject("adodb.recordset")
                        TBFI.Open "Select * from CEP where Municipio = '" & FamiliaAntiga & "'", Conexao, adOpenKeyset, adLockOptimistic
                        If TBFI.EOF = True Then
                            MsgBox ("Não é permitido liberar esta nota fiscal para envio, pois a nota esta com a cidade errada."), vbExclamation
                            .ListItems.Item(InitFor).Checked = False
                            TBFI.Close
                            Exit Sub
                        End If
                        Set TBFI = CreateObject("adodb.recordset")
                        TBFI.Open "Select * from CEP where Sigla_UF = '" & TBAbrir!txt_UF & "'", Conexao, adOpenKeyset, adLockOptimistic
                        If TBFI.EOF = True Then
                            MsgBox ("Não é permitido liberar esta nota fiscal para envio, pois a nota esta com o estado errado."), vbExclamation
                            .ListItems.Item(InitFor).Checked = False
                            TBFI.Close
                            Exit Sub
                        End If
                        Set TBFI = CreateObject("adodb.recordset")
                        TBFI.Open "Select * from CEP where Municipio = '" & FamiliaAntiga & "' and Sigla_UF = '" & TBAbrir!txt_UF & "'", Conexao, adOpenKeyset, adLockOptimistic
                        If TBFI.EOF = True Then
                            MsgBox ("Não é permitido liberar esta nota fiscal para envio, pois não existe o munícipio " & FamiliaAntiga & " no estado " & UF & " na tabela CEP."), vbExclamation
                            .ListItems.Item(InitFor).Checked = False
                            TBFI.Close
                            Exit Sub
                        End If
                    End If
                    
                    'Verifica se tem país cadastrado
                    If TBAbrir!txt_tipocliente = "JP" Or TBAbrir!txt_tipocliente = "JR" Or TBAbrir!txt_tipocliente = "FP" Or TBAbrir!txt_tipocliente = "FR" Then
                        'Cliente
                        Set TBClientes = CreateObject("adodb.recordset")
                        TBClientes.Open "Select * from Clientes where IDcliente = " & TBAbrir!Id_Int_Cliente & " and NomeRazao = '" & TBAbrir!txt_Razao_Nome & "'", Conexao, adOpenKeyset, adLockOptimistic
                        If TBClientes.EOF = False Then
                            If IsNull(TBClientes!Codigo_pais) = True Or TBClientes!Codigo_pais = "" Then
                                MsgBox ("Não é permitido liberar esta nota fiscal para envio, pois este cliente não tem país cadastrado."), vbExclamation
                                .ListItems.Item(InitFor).Checked = False
                                TBClientes.Close
                                Exit Sub
                            End If
                        End If
                    Else
                        'Fornecedor
                        Set TBClientes = CreateObject("adodb.recordset")
                        TBClientes.Open "Select * from Compras_fornecedores where IDcliente = " & TBAbrir!Id_Int_Cliente & " and Nome_Razao = '" & TBAbrir!txt_Razao_Nome & "'", Conexao, adOpenKeyset, adLockOptimistic
                        If TBClientes.EOF = False Then
                            If IsNull(TBClientes!Codigo_pais) = True Or TBClientes!Codigo_pais = "" Then
                                MsgBox ("Não é permitido liberar esta nota fiscal para envio, pois este fornecedor não tem país cadastrado."), vbExclamation
                                .ListItems.Item(InitFor).Checked = False
                                TBClientes.Close
                                Exit Sub
                            End If
                        End If
                    End If
                    TBClientes.Close
                    
                    'Verifica se tem foi gerado as dúplicatas quando for CFOP de vendas ou mão de obra
                    Set TBFI = CreateObject("adodb.recordset")
                    TBFI.Open "Select CFOP.* from tbl_Detalhes_Nota NFP INNER JOIN tbl_NaturezaOperacao CFOP ON CFOP.IDCountCfop = NFP.ID_CFOP where NFP.ID_nota = " & TBAbrir!ID & " and (CFOP.Vendas = 'True' or CFOP.MaoObra = 'True')", Conexao, adOpenKeyset, adLockOptimistic
                    If TBFI.EOF = False Then
                        Set TBFIltro = CreateObject("adodb.recordset")
                        TBFIltro.Open "Select * from tbl_Detalhes_Recebimento where ID_nota = " & TBAbrir!ID, Conexao, adOpenKeyset, adLockOptimistic
                        If TBFIltro.EOF = True Then
                            If MsgBox("A(s) duplicata(s) ainda não foi(ram) gerada(s), deseja prosseguir assim mesmo?", vbQuestion + vbYesNo) = vbNo Then
                                .ListItems.Item(InitFor).Checked = False
                                TBFI.Close
                                TBFIltro.Close
                                Exit Sub
                            End If
                        End If
                        TBFIltro.Close
                    End If
                    TBFI.Close
                End If
                TBAbrir.Close
                
                Set TBMaquinas = CreateObject("adodb.recordset")
                TBMaquinas.Open "Select * from Empresa where Empresa = '" & ListaNota.SelectedItem.ListSubItems(1) & "' and GNFe = 'True'", Conexao, adOpenKeyset, adLockOptimistic
                If TBMaquinas.EOF = False Then
                    'Verifica se esta preenchido o caminho para salvar o arquivo de envio da NFe
                    If IsNull(TBMaquinas!Caminho_Nfe) = True Or TBMaquinas!Caminho_Nfe = "" Then
                        MsgBox ("Não é permitido liberar a nota fiscal para envio, pois não foi informado o caminho onde será armazenado os aquivos para envio."), vbExclamation
                        .ListItems.Item(InitFor).Checked = False
                        Exit Sub
                    End If
                    'Verificar se o caminho existe
                    If GerArqPastas.FolderExists(TBMaquinas!Caminho_Nfe) = False Then
                        MsgBox ("Não é permitido liberar a nota fiscal para envio, pois não foi encontrado o caminho " & TBMaquinas!Caminho_Nfe & ", onde será armazenado os aquivos para envio."), vbExclamation
                        .ListItems.Item(InitFor).Checked = False
                        Exit Sub
                    End If
                End If
                TBMaquinas.Close
                
                'Verifica se é nota fiscal de devolução ou complementar e se esta referenciado a nota fiscal
                Set TBMaquinas = CreateObject("adodb.recordset")
                TBMaquinas.Open "Select NFE.ID from tbl_Dados_Nota_Fiscal_NFe NFE LEFT JOIN Faturamento_Relacionamento FR ON FR.ID_nota = NFE.ID_nota where NFE.ID_nota = " & Item & " and NFE.Finalidade_emissao <> 1 and NFE.Finalidade_emissao <> 3 and FR.ID IS NULL", Conexao, adOpenKeyset, adLockOptimistic
                If TBMaquinas.EOF = False Then
                    MsgBox ("Não é permitido liberar a nota fiscal para envio, pois não foi feito o relacionamento."), vbExclamation
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
                End If
                TBMaquinas.Close
                
                'Verifica se o clinte é fisico e esta com cnpj e vice e versa
                Set TBMaquinas = CreateObject("adodb.recordset")
                TBMaquinas.Open "Select txt_tipocliente, txt_CNPJ_CPF from tbl_Dados_Nota_Fiscal where ID = " & Item & " and txt_uf <> 'EX'", Conexao, adOpenKeyset, adLockReadOnly
                If TBMaquinas.EOF = False Then
                    If Left(TBMaquinas!txt_tipocliente, 1) = "J" And Len(TBMaquinas!txt_CNPJ_CPF) < 14 Then
                        MsgBox ("Não é permitido liberar a nota fiscal para envio, pois o CNPJ do destinatario esta errado."), vbExclamation
                        TBMaquinas.Close
                        .ListItems.Item(InitFor).Checked = False
                        Exit Sub
                    ElseIf Left(TBMaquinas!txt_tipocliente, 1) = "F" And Len(TBMaquinas!txt_CNPJ_CPF) > 14 Then
                        MsgBox ("Não é permitido liberar a nota fiscal para envio, pois o CPF do destinatario esta errado."), vbExclamation
                        TBMaquinas.Close
                        .ListItems.Item(InitFor).Checked = False
                        Exit Sub
                    End If
                End If
                TBMaquinas.Close
            End If
        End If
    Next InitFor
End With

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Private Sub ListaNota_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If ListaNota.ListItems.Count = 0 Then Exit Sub
ProcLimpaCampos
CodigoLista = ListaNota.SelectedItem.Index
txtID_nota = ListaNota.SelectedItem
txtNota = ListaNota.SelectedItem.SubItems(3)
txtSerie = ListaNota.SelectedItem.SubItems(5)

ProcCarregaEntrega
ProcCarregaCobranca

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from tbl_Dados_Nota_Fiscal_NFe where id_nota = " & txtID_nota, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    ProcPuxaDados
End If
TBAbrir.Close

UF_transp = ""
Cidade = ""

Set TBAbrir = CreateObject("adodb.recordset")
If ListaNota.SelectedItem.ListSubItems(1) = "" Then
    TBAbrir.Open "Select * from empresa", Conexao, adOpenKeyset, adLockOptimistic
Else
    TBAbrir.Open "Select * from empresa where Empresa = '" & ListaNota.SelectedItem.ListSubItems(1) & "'", Conexao, adOpenKeyset, adLockOptimistic
End If
If TBAbrir.EOF = False Then
    UF_transp = TBAbrir!UF
    Cidade = TBAbrir!Cidade
End If
TBAbrir.Close

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from tbl_dados_transp where ID_Nota = " & txtID_nota, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    Frame2.Enabled = True
    If IsNull(TBAbrir!UF_embarque) = False Then
        cmbUF_embarque = TBAbrir!UF_embarque
    Else
        If UF_transp <> "" Then cmbUF_embarque = UF_transp
    End If
    txtLocal_embarque = IIf(IsNull(TBAbrir!Local_embarque), Cidade, TBAbrir!Local_embarque)
Else
    Frame2.Enabled = False
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Sub ProcLimpaCampos()
On Error GoTo tratar_erro

Cmb_forma_de_emissao = "1 - Normal"
cmbFinalidade_emissao = "1 - Normal"
cmbFormaPag.ListIndex = -1
Cmb_consumidor.ListIndex = -1
Cmb_presenca_comprador.ListIndex = -1
Cmb_arquivos_XML_enviados_email = "2 - Arquivo XML de Compartilhamento da NFe"
txtID_entrega = 0
cmbEntrega.ListIndex = -1
Txt_ID_cobranca = 0
Cmb_cobranca.ListIndex = -1
Cmb_enviar_DANFE = "Sim"
txtLocal_embarque = ""
cmbUF_embarque.ListIndex = -1
txtStatus = ""
Txt_chave_acesso = ""
chkCodRef.Value = 0
txtID_nota = 0

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Sub ProcLimpacamposProdutos()
On Error GoTo tratar_erro

txtID_item = 0
cmbModalidade_determinacao.ListIndex = -1
cmbModalidade_determinacao_ST.ListIndex = -1
Cmb_codigo_ANP.ListIndex = -1
Cmb_UF_consumo.ListIndex = -1
Cmb_tipo_produto.ListIndex = -1
txtDescANP = ""

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Private Sub ListaProdutos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView listaProdutos, ColumnHeader

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Private Sub listaProdutos_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With listaProdutos
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            ProcVerificaStatus .ListItems(InitFor)
            If Permitido = False Then
                MsgBox ("Não é permitido excluir os dados do produto desta nota fiscal, pois a mesma está " & NomeCampo & "."), vbExclamation
                .ListItems.Item(InitFor).Checked = False
            End If
        End If
    Next InitFor
End With

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Private Sub ListaProdutos_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If listaProdutos.ListItems.Count = 0 Then Exit Sub
ProcLimpacamposProdutos
CodigoLista1 = listaProdutos.SelectedItem.Index
txtID_item = listaProdutos.SelectedItem
If Len(listaProdutos.SelectedItem.SubItems(3)) = 4 Then Quant = Right(listaProdutos.SelectedItem.SubItems(3), 3) Else Quant = Right(listaProdutos.SelectedItem.SubItems(3), 2)
With cmbModalidade_determinacao_ST
    If Quant = "00" Or Quant = "10" Or Quant = "20" Or Quant = "51" Or Quant = "70" Or Quant = "90" Or Quant = "201" Or Quant = "202" Or Quant = "900" Then
        FrameCST.Enabled = True
        If Quant = "00" Or Quant = "20" Or Quant = "51" Then
            .Locked = True
            .TabStop = False
        Else
            .Locked = False
            .TabStop = True
        End If
    Else
        FrameCST.Enabled = False
    End If
End With

If Quant = "00" Or Quant = "10" Or Quant = "20" Or Quant = "51" Or Quant = "70" Or Quant = "90" Or Quant = "201" Or Quant = "202" Or Quant = "900" Then
    Set TBCST = CreateObject("adodb.recordset")
    TBCST.Open "select * from tbl_Detalhes_Nota_CST_ICMS where id_item = " & txtID_item, Conexao, adOpenKeyset, adLockOptimistic
    If TBCST.EOF = False Then
        With cmbModalidade_determinacao
            Select Case TBCST!Modalidade_determinacao
                Case "0": .Text = "0 - Margem valor agregado (%)"
                Case "1": .Text = "1 - Pauta (valor)"
                Case "2": .Text = "2 - Preço tabelado máx. (valor)"
                Case "3": .Text = "3 - valor da operação"
            End Select
        End With
        With cmbModalidade_determinacao_ST
            If Quant = "10" Or Quant = "70" Or Quant = "90" Or Quant = "201" Or Quant = "202" Or Quant = "900" Then
                Select Case TBCST!Modalidade_determinacao_ST
                    Case "0": .Text = "0 - Preço tabelado ou máximo sugerido"
                    Case "1": .Text = "1 - Lista negativa (valor)"
                    Case "2": .Text = "2 - Lista positiva (valor)"
                    Case "3": .Text = "3 - Lista neutra (valor)"
                    Case "4": .Text = "4 - Margem valor agregado (%)"
                    Case "5": .Text = "5 - Pauta (valor)"
                End Select
            End If
        End With
    End If
    TBCST.Close
End If

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select CFOP.* from tbl_Detalhes_Nota NFP INNER JOIN tbl_NaturezaOperacao CFOP ON NFP.ID_cfop = CFOP.IDCountCfop where NFP.Int_codigo = " & txtID_item & " and (Right(CFOP.id_CFOP, 3) = '651' or Right(CFOP.id_CFOP, 3) = '652' or Right(CFOP.id_CFOP, 3) = '653' or Right(CFOP.id_CFOP, 3) = '654' or Right(CFOP.id_CFOP, 3) = '655' or Right(CFOP.id_CFOP, 3) = '656' or Right(CFOP.id_CFOP, 3) = '657' or Right(CFOP.id_CFOP, 3) = '658' or Right(CFOP.id_CFOP, 3) = '659' or Right(CFOP.id_CFOP, 3) = '660' or Right(CFOP.id_CFOP, 3) = '661' or Right(CFOP.id_CFOP, 3) = '662' or Right(CFOP.id_CFOP, 3) = '663' or Right(CFOP.id_CFOP, 3) = '664' or Right(CFOP.id_CFOP, 3) = '665' or Right(CFOP.id_CFOP, 3) = '666')", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    Frame_comb_lub.Enabled = True
    Cmb_tipo_produto = "4 - Combustível"
    Set TBFI = CreateObject("adodb.recordset")
    TBFI.Open "Select NFPe.*, CPANP.Descricao from tbl_Detalhes_Nota_NFe NFPe INNER JOIN Codigos_produtos_ANP CPANP ON NFPe.Codigo_ANP = CPANP.Codigo where NFPe.ID_item = " & txtID_item, Conexao, adOpenKeyset, adLockOptimistic
    If TBFI.EOF = False Then
        If IsNull(TBFI!Codigo_ANP) = False And TBFI!Codigo_ANP <> "" Then Cmb_codigo_ANP = TBFI!Codigo_ANP
        If IsNull(TBFI!UF_consumo) = False And TBFI!UF_consumo <> "" Then Cmb_UF_consumo = TBFI!UF_consumo
        If IsNull(TBFI!Tipo_produto) = False And TBFI!Tipo_produto <> "" Then
            With Cmb_tipo_produto
                Select Case TBFI!Tipo_produto
                    Case 0: .Text = "0 - Produtos"
                    Case 1: .Text = "1 - Veículos"
                    Case 2: .Text = "2 - Medicamentos"
                    Case 3: .Text = "3 - Armamentos"
                    Case 4: .Text = "4 - Combustível"
                    Case 5: .Text = "5 - Serviço"
                End Select
            End With
        End If
    End If
    TBFI.Close
Else
    Frame_comb_lub.Enabled = False
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Private Sub SStab_nfe_Click(PreviousTab As Integer)
On Error GoTo tratar_erro

If txtID_nota = 0 Then
    SStab_nfe.Tab = 0
    Exit Sub
End If
Select Case SStab_nfe.Tab
    Case 0: 'Dados da nota
        If ListaNota.Visible = True Then ListaNota.SetFocus
    Case 1: 'Lista de produtos
        listaProdutos.SetFocus
        ProcLimpacamposProdutos
        ProcCarregaListaProdutos
End Select
    
Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Sub ProcPuxaDados()
On Error GoTo tratar_erro

With cmbForma_pagamento
    Select Case TBAbrir!Forma_pagamento
        Case "0": .Text = "0 - pagamento à vista"
        Case "1": .Text = "1 - pagamento à prazo"
    End Select
End With
With cmbFormaPag
    Select Case TBAbrir!FormaPagto
        Case "01": .Text = "01 - Dinheiro"
        Case "02": .Text = "02 - Cheque"
        Case "03": .Text = "03 - Cartão de Crédito"
        Case "04": .Text = "04 - Cartão de Débito"
        Case "05": .Text = "05 - Crédito Loja"
        Case "10": .Text = "10 - Vale Alimentação"
        Case "11": .Text = "11 - Vale Refeição"
        Case "12": .Text = "12 - Vale Presente"
        Case "13": .Text = "13 - Vale Combustível"
        Case "15": .Text = "15 - Boleto Bancário"
        Case "90": .Text = "90 - Sem pagamento"
        Case "99": .Text = "99 - Outros"
    End Select
End With
With Cmb_forma_de_emissao
    Select Case TBAbrir!Forma_emissao
        Case "1": .Text = "1 - Normal"
        Case "2": .Text = "2 - Conting. FS - emissão c/ impressão do DANFE em Formulário de Segurança"
        Case "3": .Text = "3 - Conting. SCAN - emissão no Sistema de Contingência do Ambiente Nacional (SCAN)"
        Case "4": .Text = "4 - Conting. DPEC - emissão c/ envio da Declaração Prévia de Emissão em Contingência (DPEC)"
        Case "5": .Text = "5 - Conting. FS-DA - emissão c/ impr. do DANFE em Formul. de Segurança p/ Impr. de Doc. Aux. de Doc. Fiscal Eletr. (FS-DA)"
        Case "6": .Text = "6 - Contingência SVC-AN - emissão em contingência na SEFAZ Virtual de Contingência"
        Case "7": .Text = "7 - Contingência SVC-RS - emissão em contingência na SEFAZ Virtual de Contingência"
    End Select
End With
With cmbFinalidade_emissao
    Select Case TBAbrir!Finalidade_emissao
        Case "1": .Text = "1 - Normal"
        Case "2": .Text = "2 - Complementar"
        Case "3": .Text = "3 - Ajuste"
        Case "4": .Text = "4 - Devolução/Retorno"
    End Select
End With
With Cmb_consumidor
    Select Case TBAbrir!Consumidor_final
        Case "0": .Text = "0 - Não"
        Case "1": .Text = "1 - Sim"
    End Select
End With
With Cmb_presenca_comprador
    Select Case TBAbrir!Presenca_comprador
        Case "0": .Text = "0 - Não se aplica"
        Case "1": .Text = "1 - Operação presencial"
        Case "2": .Text = "2 - Operação não presencial, pela Internet"
        Case "3": .Text = "3 - Operação não presencial, teleatendimento"
        Case "4": .Text = "4 - NFC-e em operação com entrega em domicílio"
        Case "9": .Text = "9 - Operação não presencial, outros"
    End Select
End With
With Cmb_arquivos_XML_enviados_email
    Select Case TBAbrir!Enviar_Email
        Case "1": .Text = "1 - Arquivo XML da NF-e"
        Case "2": .Text = "2 - Arquivo XML de Compartilhamento da NFe"
        Case "3": .Text = "3 - Ambos"
        Case "4": .Text = "4 - Nenhum"
        Case "5": .Text = "5 - Usar dos Parâmetros Gerais"
    End Select
End With
Txt_titulo_canhoto_DANFE = IIf(IsNull(TBAbrir!Titulo_canhoto_DANFE), "", TBAbrir!Titulo_canhoto_DANFE)
Txt_titulo_canhoto_DANFE = IIf(IsNull(TBAbrir!Texto_canhoto_DANFE), "", TBAbrir!Texto_canhoto_DANFE)
If IsNull(TBAbrir!Enviar_DANFE_email) = False And TBAbrir!Enviar_DANFE_email <> "" Then
    If TBAbrir!Enviar_DANFE_email = "S" Then Cmb_enviar_DANFE = "Sim" Else Cmb_enviar_DANFE = "Não"
Else
    Cmb_enviar_DANFE = "Sim"
End If

txtID_entrega = IIf(IsNull(TBAbrir!ID_entrega), 0, TBAbrir!ID_entrega)
If TBAbrir!DA_entrega = False Then Chk_DA_entrega.Value = 0 Else Chk_DA_entrega.Value = 1
Txt_ID_cobranca = IIf(IsNull(TBAbrir!ID_Cobranca), 0, TBAbrir!ID_Cobranca)
If TBAbrir!DA_cobranca = False Then Chk_DA_cobranca.Value = 0 Else Chk_DA_cobranca.Value = 1
txtStatus = FunVerifStatusNFe(TBAbrir!ID_nota)
Txt_chave_acesso = IIf(IsNull(TBAbrir!Chave_acesso), "", TBAbrir!Chave_acesso)

If IsNull(TBAbrir!CodRef) = True Then
    Set TBCiclo = CreateObject("adodb.recordset")
    TBCiclo.Open "Select Codigo_ref_DANFE from empresa where codigo = " & frmFaturamento_Prod_Serv.Cmb_empresa.ItemData(frmFaturamento_Prod_Serv.Cmb_empresa.ListIndex), Conexao, adOpenKeyset, adLockOptimistic
    If TBCiclo.EOF = False Then
        If TBCiclo!Codigo_ref_DANFE = False Or IsNull(TBCiclo!Codigo_ref_DANFE) = True Then chkCodRef.Value = 0 Else chkCodRef.Value = 1
    End If
Else
    If TBAbrir!CodRef = False Then chkCodRef.Value = 0 Else chkCodRef.Value = 1
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Sub ProcCarregaEntrega()
On Error GoTo tratar_erro

With cmbEntrega
    .Clear
    Set TBFIltro = CreateObject("adodb.recordset")
    TBFIltro.Open "Select ID, Id_Int_Cliente, txt_Razao_Nome from tbl_Dados_Nota_Fiscal where id = " & txtID_nota, Conexao, adOpenKeyset, adLockReadOnly
    If TBFIltro.EOF = False Then
        
        'Verifica se é cliente ou fornecedor
        Set TBFI = CreateObject("adodb.recordset")
        TBFI.Open "Select * from Clientes where IDCliente = " & TBFIltro!Id_Int_Cliente & " and NomeRazao = '" & TBFIltro!txt_Razao_Nome & "'", Conexao, adOpenKeyset, adLockReadOnly
        If TBFI.EOF = False Then Tipo = "C" Else Tipo = "F"
        TBFI.Close
        
        Permitido = True
        Set TBVendas = CreateObject("adodb.recordset")
        If Tipo = "C" Then
            TextoID = ""
            TBVendas.Open "Select VC.* from (vendas_comercial VC INNER JOIN vendas_proposta VP ON VP.Cotacao = VC.Cotacao) INNER JOIN tbl_proposta_nota PN ON PN.proposta = VP.Ncotacao and PN.Revisao = VP.Revisao where PN.ID_nota = " & TBFIltro!ID & " and VC.ID_entrega IS NOT NULL and VC.ID_entrega <> 0 order by VC.ID_entrega", Conexao, adOpenKeyset, adLockReadOnly
            If TBVendas.EOF = False Then
                Permitido = False
                Do While TBVendas.EOF = False
                    If TextoID <> TBVendas!ID_entrega Then
                        .AddItem TBVendas!Local_entrega
                        .ItemData(cmbEntrega.NewIndex) = TBVendas!ID_entrega
                        TextoID = TBVendas!ID_entrega
                    End If
                    TBVendas.MoveNext
                Loop
            End If
            TBVendas.Close
        Else
            TBVendas.Open "Select CC.* from (Compras_comercial CC INNER JOIN Compras_pedido CP ON CP.IDpedido = CC.IDpedido) INNER JOIN tbl_proposta_nota PN ON PN.proposta = CP.Pedido and PN.Revisao = 0 where PN.ID_nota = " & TBFIltro!ID & " and CC.ID_entrega IS NOT NULL and CC.ID_entrega <> 0 and CC.localentrega IS NOT NULL", Conexao, adOpenKeyset, adLockReadOnly
            If TBVendas.EOF = False Then
                Permitido = False
                .AddItem TBVendas!localentrega
                .ItemData(cmbEntrega.NewIndex) = TBVendas!ID_entrega
                txtID_entrega = TBVendas!ID_entrega
                .Text = TBVendas!localentrega
            End If
            TBVendas.Close
        End If
        
        If Permitido = True Then
            Set TBClientes = CreateObject("adodb.recordset")
            TBClientes.Open "Select * from clientes_entrega where idcliente = " & TBFIltro!Id_Int_Cliente & " and Tipo = '" & Tipo & "'", Conexao, adOpenKeyset, adLockReadOnly
            If TBClientes.EOF = False Then
                Do While TBClientes.EOF = False
                    If IsNull(TBClientes!Tipo_endereco) = False And TBClientes!Tipo_endereco <> "" Then
                        Endereco = TBClientes!Tipo_endereco & ": " & IIf(IsNull(TBClientes!endereco_entrega), "", TBClientes!endereco_entrega)
                    Else
                        Endereco = IIf(IsNull(TBClientes!endereco_entrega), "", TBClientes!endereco_entrega)
                    End If
                    If IsNull(TBClientes!Tipo_bairro) = False And TBClientes!Tipo_bairro <> "" Then
                        Bairro = TBClientes!Tipo_bairro & ": " & IIf(IsNull(TBClientes!bairro_entrega), "", TBClientes!bairro_entrega)
                    Else
                        Bairro = IIf(IsNull(TBClientes!bairro_entrega), "", TBClientes!bairro_entrega)
                    End If
                    Endereco2 = Endereco & " - " & IIf(IsNull(TBClientes!Numero), "", TBClientes!Numero) & " - " & Bairro & " - " & IIf(IsNull(TBClientes!cidade_entrega), "", TBClientes!cidade_entrega) & " - " & IIf(IsNull(TBClientes!uf_entrega), "", TBClientes!uf_entrega) & " - " & IIf(IsNull(TBClientes!cep_entrega), "", TBClientes!cep_entrega)
                    .AddItem Endereco2
                    .ItemData(cmbEntrega.NewIndex) = TBClientes!identrega
                    txtID_entrega = IIf(IsNull(TBClientes!identrega), 0, TBClientes!identrega)
                    .Text = Endereco2
                    TBClientes.MoveNext
                Loop
            End If
        End If
        
        identrega = 0
        Set TBAcessos = CreateObject("adodb.recordset")
        TBAcessos.Open "Select ID_entrega from tbl_Dados_Nota_Fiscal_NFe where id_nota = " & txtID_nota & " and ID_entrega IS NOT NULL", Conexao, adOpenKeyset, adLockReadOnly
        If TBAcessos.EOF = False Then
            Set TBClientes = CreateObject("adodb.recordset")
            TBClientes.Open "Select * from clientes_entrega where identrega = " & TBAcessos!ID_entrega, Conexao, adOpenKeyset, adLockOptimistic
            If TBClientes.EOF = False Then
                txtID_entrega = TBClientes!identrega
                identrega = TBClientes!identrega
                
                If IsNull(TBClientes!Tipo_endereco) = False And TBClientes!Tipo_endereco <> "" Then
                    Endereco = TBClientes!Tipo_endereco & ": " & IIf(IsNull(TBClientes!endereco_entrega), "", TBClientes!endereco_entrega)
                Else
                    Endereco = IIf(IsNull(TBClientes!endereco_entrega), "", TBClientes!endereco_entrega)
                End If
                If IsNull(TBClientes!Tipo_bairro) = False And TBClientes!Tipo_bairro <> "" Then
                    Bairro = TBClientes!Tipo_bairro & ": " & IIf(IsNull(TBClientes!bairro_entrega), "", TBClientes!bairro_entrega)
                Else
                    Bairro = IIf(IsNull(TBClientes!bairro_entrega), "", TBClientes!bairro_entrega)
                End If
                
                Endereco = Endereco & " - " & IIf(IsNull(TBClientes!Numero), "", TBClientes!Numero) & " - " & Bairro & " - " & IIf(IsNull(TBClientes!cidade_entrega), "", TBClientes!cidade_entrega) & " - " & IIf(IsNull(TBClientes!uf_entrega), "", TBClientes!uf_entrega) & " - " & IIf(IsNull(TBClientes!cep_entrega), "", TBClientes!cep_entrega)
                .Text = Endereco
            End If
            TBClientes.Close
        End If
    End If
    TBFIltro.Close
End With

Exit Sub
tratar_erro:
    If Err.Number = 383 Then
        With cmbEntrega
            .AddItem Endereco
            .ItemData(cmbEntrega.NewIndex) = identrega
            .Text = Endereco
        End With
        Exit Sub
    End If
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Sub ProcCarregaCobranca()
On Error GoTo tratar_erro

With Cmb_cobranca
    .Clear
    Set TBFIltro = CreateObject("adodb.recordset")
    TBFIltro.Open "Select ID, Id_Int_Cliente, txt_Razao_Nome from tbl_Dados_Nota_Fiscal where id = " & txtID_nota, Conexao, adOpenKeyset, adLockReadOnly
    If TBFIltro.EOF = False Then
        
        'Verifica se é cliente ou fornecedor
        Set TBFI = CreateObject("adodb.recordset")
        TBFI.Open "Select * from Clientes where IDCliente = " & TBFIltro!Id_Int_Cliente & " and NomeRazao = '" & TBFIltro!txt_Razao_Nome & "'", Conexao, adOpenKeyset, adLockReadOnly
        If TBFI.EOF = False Then Tipo = "C" Else Tipo = "F"
        TBFI.Close
        
        Permitido = True
        If Tipo = "C" Then
            TextoID = ""
            Set TBVendas = CreateObject("adodb.recordset")
            TBVendas.Open "Select VC.* from (vendas_comercial VC INNER JOIN vendas_proposta VP ON VP.Cotacao = VC.Cotacao) INNER JOIN tbl_proposta_nota PN ON PN.proposta = VP.Ncotacao and PN.Revisao = VP.Revisao where PN.ID_nota = " & TBFIltro!ID & " and VC.ID_cobranca IS NOT NULL and VC.ID_cobranca <> 0 order by VC.ID_cobranca", Conexao, adOpenKeyset, adLockReadOnly
            If TBVendas.EOF = False Then
                Permitido = False
                Do While TBVendas.EOF = False
                    If TextoID <> TBVendas!ID_Cobranca Then
                        .AddItem TBVendas!Local_cobranca
                        .ItemData(Cmb_cobranca.NewIndex) = TBVendas!ID_Cobranca
                        TextoID = TBVendas!ID_Cobranca
                    End If
                    TBVendas.MoveNext
                Loop
            End If
            TBVendas.Close
        End If
        
        If Permitido = True Then
            Set TBClientes = CreateObject("adodb.recordset")
            TBClientes.Open "Select * from clientes_cobranca where idcliente = " & TBFIltro!Id_Int_Cliente & " and Tipo = '" & Tipo & "'", Conexao, adOpenKeyset, adLockReadOnly
            If TBClientes.EOF = False Then
                Do While TBClientes.EOF = False
                    If IsNull(TBClientes!Tipo_endereco) = False And TBClientes!Tipo_endereco <> "" Then
                        Endereco = TBClientes!Tipo_endereco & ": " & IIf(IsNull(TBClientes!Endereco_cobranca), "", TBClientes!Endereco_cobranca)
                    Else
                        Endereco = IIf(IsNull(TBClientes!Endereco_cobranca), "", TBClientes!Endereco_cobranca)
                    End If
                    If IsNull(TBClientes!Tipo_bairro) = False And TBClientes!Tipo_bairro <> "" Then
                        Bairro = TBClientes!Tipo_bairro & ": " & IIf(IsNull(TBClientes!bairro_cobranca), "", TBClientes!bairro_cobranca)
                    Else
                        Bairro = IIf(IsNull(TBClientes!bairro_cobranca), "", TBClientes!bairro_cobranca)
                    End If
                    Endereco2 = Endereco & " - " & IIf(IsNull(TBClientes!Numero), "", TBClientes!Numero) & " - " & Bairro & " - " & IIf(IsNull(TBClientes!cidade_cobranca), "", TBClientes!cidade_cobranca) & " - " & IIf(IsNull(TBClientes!uf_cobranca), "", TBClientes!uf_cobranca) & " - " & IIf(IsNull(TBClientes!cep_cobranca), "", TBClientes!cep_cobranca)
                    .AddItem Endereco2
                    .ItemData(Cmb_cobranca.NewIndex) = TBClientes!idcobranca
                    Txt_ID_cobranca = IIf(IsNull(TBClientes!idcobranca), 0, TBClientes!idcobranca)
                    .Text = Endereco2
                    TBClientes.MoveNext
                Loop
            End If
        End If
        idcobranca = 0
        Set TBAcessos = CreateObject("adodb.recordset")
        TBAcessos.Open "Select ID_Cobranca from tbl_Dados_Nota_Fiscal_NFe where id_nota = " & txtID_nota & " and ID_Cobranca IS NOT NULL", Conexao, adOpenKeyset, adLockReadOnly
        If TBAcessos.EOF = False Then
            Set TBClientes = CreateObject("adodb.recordset")
            TBClientes.Open "Select * from clientes_cobranca where idcobranca = " & TBAcessos!ID_Cobranca, Conexao, adOpenKeyset, adLockReadOnly
            If TBClientes.EOF = False Then
                Txt_ID_cobranca = TBClientes!idcobranca
                idcobranca = TBClientes!idcobranca
                
                If IsNull(TBClientes!Tipo_endereco) = False And TBClientes!Tipo_endereco <> "" Then
                    Endereco = TBClientes!Tipo_endereco & ": " & IIf(IsNull(TBClientes!Endereco_cobranca), "", TBClientes!Endereco_cobranca)
                Else
                    Endereco = IIf(IsNull(TBClientes!Endereco_cobranca), "", TBClientes!Endereco_cobranca)
                End If
                If IsNull(TBClientes!Tipo_bairro) = False And TBClientes!Tipo_bairro <> "" Then
                    Bairro = TBClientes!Tipo_bairro & ": " & IIf(IsNull(TBClientes!bairro_cobranca), "", TBClientes!bairro_cobranca)
                Else
                    Bairro = IIf(IsNull(TBClientes!bairro_cobranca), "", TBClientes!bairro_cobranca)
                End If
                Endereco = Endereco & " - " & IIf(IsNull(TBClientes!Numero), "", TBClientes!Numero) & " - " & Bairro & " - " & IIf(IsNull(TBClientes!cidade_cobranca), "", TBClientes!cidade_cobranca) & " - " & IIf(IsNull(TBClientes!uf_cobranca), "", TBClientes!uf_cobranca) & " - " & IIf(IsNull(TBClientes!cep_cobranca), "", TBClientes!cep_cobranca)
                .Text = Endereco
            End If
            TBClientes.Close
        End If
    End If
    TBFIltro.Close
End With

Exit Sub
tratar_erro:
    If Err.Number = 383 Then
        With Cmb_cobranca
            .AddItem Endereco
            .ItemData(Cmb_cobranca.NewIndex) = idcobranca
            .Text = Endereco
        End With
        Exit Sub
    End If
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Private Sub Timer_status_NFe_Timer()
On Error GoTo tratar_erro

'If Timer_status_NFe.Enabled = True Then ProcAtualizaStatusNFe
    
Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
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
    MsgBox ("Descrição do erro : " + Error()), vbCritical
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
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Sub ProcSalvarProduto()
On Error GoTo tratar_erro

If Alterar = False Then
    MsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation
    Exit Sub
End If
Acao = "salvar"
If txtID_item = 0 Then
    NomeCampo = "Produto"
    ProcVerificaAcao
    Exit Sub
End If

If FrameCST.Enabled = True Then
    If Len(listaProdutos.SelectedItem.SubItems(3)) = 4 Then Quant = Right(listaProdutos.SelectedItem.SubItems(3), 3) Else Quant = Right(listaProdutos.SelectedItem.SubItems(3), 2)
    If cmbModalidade_determinacao <> "" Or cmbModalidade_determinacao_ST <> "" Then
        Set TBCST = CreateObject("adodb.recordset")
        TBCST.Open "Select * from tbl_Detalhes_Nota_CST_ICMS where id_item = " & txtID_item, Conexao, adOpenKeyset, adLockOptimistic
        If TBCST.EOF = True Then
            TBCST.AddNew
            If cmbModalidade_determinacao <> "" Then TBCST!Modalidade_determinacao = Left(cmbModalidade_determinacao, 1) Else TBCST!Modalidade_determinacao = Null
            If cmbModalidade_determinacao <> "" And (Quant = "10" Or Quant = "70" Or Quant = "90" Or Quant = "201" Or Quant = "202" Or Quant = "900") Then TBCST!Modalidade_determinacao_ST = Left(cmbModalidade_determinacao_ST, 1) Else TBCST!Modalidade_determinacao_ST = Null
            TBCST.Update
        Else
            'Pode ter mais de um com o mesmo ID do produto (nota de importação)
            If cmbModalidade_determinacao <> "" Then
                Conexao.Execute "Update tbl_Detalhes_Nota_CST_ICMS Set Modalidade_determinacao = " & Left(cmbModalidade_determinacao, 1) & " where id_item = " & txtID_item
            Else
                Conexao.Execute "Update tbl_Detalhes_Nota_CST_ICMS Set Modalidade_determinacao = NULL where id_item = " & txtID_item
            End If
            If cmbModalidade_determinacao_ST <> "" Then
                If Quant = "10" Or Quant = "70" Or Quant = "90" Or Quant = "201" Or Quant = "202" Or Quant = "900" Then Conexao.Execute "Update tbl_Detalhes_Nota_CST_ICMS Set Modalidade_determinacao_ST = " & Left(cmbModalidade_determinacao_ST, 1) & " where id_item = " & txtID_item
            Else
                Conexao.Execute "Update tbl_Detalhes_Nota_CST_ICMS Set Modalidade_determinacao_ST = NULL where id_item = " & txtID_item
            End If
        End If
    End If
End If

If Frame_comb_lub.Enabled = True Then
    If Cmb_codigo_ANP = "" Then
        NomeCampo = "o código do produto da ANP"
        ProcVerificaAcao
        Cmb_codigo_ANP.SetFocus
        Exit Sub
    End If
    If Cmb_UF_consumo = "" Then
        NomeCampo = "a UF de consumo"
        ProcVerificaAcao
        Cmb_UF_consumo.SetFocus
        Exit Sub
    End If
    If Cmb_tipo_produto = "" Then
        NomeCampo = "o tipo do produto"
        ProcVerificaAcao
        Cmb_tipo_produto.SetFocus
        Exit Sub
    End If
    
    Set TBFI = CreateObject("adodb.recordset")
    TBFI.Open "Select * from tbl_Detalhes_Nota_NFe where ID_item = " & txtID_item, Conexao, adOpenKeyset, adLockOptimistic
    If TBFI.EOF = True Then TBFI.AddNew
    TBFI!Id_Item = txtID_item
    TBFI!ID_nota = txtID_nota
    TBFI!Codigo_ANP = Cmb_codigo_ANP
    TBFI!Descricao_ANP = txtDescANP
    TBFI!UF_consumo = Cmb_UF_consumo
    TBFI!Tipo_produto = Left(Cmb_tipo_produto, 1)
    TBFI.Update
    TBFI.Close
End If

MsgBox ("Alteração efetuada com sucesso."), vbInformation
'==================================
Modulo = Formulario
Evento = "Alterar dados do produto da nota fiscal"
ID_documento = listaProdutos.SelectedItem
With frmFaturamento_Prod_Serv
    .ProcVerificaTipoNF False
    If .txtNFiscal = "" Then NomeCampo = "N° ordem: " & .txtid Else NomeCampo = "N° nota: " & .txtNFiscal
    Documento = NomeCampo & " - Tipo: " & TipoNF & " - Série: " & .txtSerie
End With
Documento1 = "Cód. interno: " & listaProdutos.SelectedItem.ListSubItems(1)
ProcGravaEvento
'==================================
If CodigoLista1 <> 0 And listaProdutos.ListItems.Count <> 0 Then
    listaProdutos.SelectedItem = listaProdutos.ListItems(CodigoLista1)
    listaProdutos.SetFocus
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Sub ProcVerifLiberacao(ID_nota As Long, Empresa As String, Mensagem As Boolean)
On Error GoTo tratar_erro

Permitido = True
Familiatext = ""

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from tbl_Dados_Nota_Fiscal_NFe where id_nota = " & ID_nota, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = True Then
    If Mensagem = True Then MsgBox ("Salve os dados da NF-e antes de liberar para envio."), vbExclamation
    Permitido = False
    Exit Sub
Else
    If IsNull(TBAbrir!Forma_emissao) = True Or TBAbrir!Forma_emissao = "" Then
        If Mensagem = True Then MsgBox ("Salve os dados da NF-e antes de liberar para envio."), vbExclamation
        Permitido = False
        Exit Sub
    End If
    If TBAbrir!Finalidade_emissao = 4 Then
        Set TBFI = CreateObject("adodb.recordset")
        TBFI.Open "Select Id from tbl_dados_transp where id_nota = " & ID_nota & " and txt_Frete_Conta = 9", Conexao, adOpenKeyset, adLockOptimistic
        If TBFI.EOF = False Then
            If Mensagem = True Then MsgBox ("É necessário cadastrar a transportadora antes de liberar para envio."), vbExclamation
            Permitido = False
            Exit Sub
        Else
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select Id from tbl_dados_transp where id_nota = " & ID_nota, Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = True Then
                If Mensagem = True Then MsgBox ("É necessário cadastrar a transportadora antes de liberar para envio."), vbExclamation
                Permitido = False
                Exit Sub
            End If
        End If
        TBFI.Close
    End If
End If

'Dados da nota fiscal
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from tbl_dados_nota_fiscal where id = " & ID_nota, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    If TBAbrir!Serie = "" Or IsNull(TBAbrir!Serie) = True Then
        Familiatext = "Série da NF"
        Permitido = False
    End If
    If TBAbrir!txt_UF <> "" And IsNull(TBAbrir!txt_UF) = False And TBAbrir!txt_UF <> "EX" Then
        If TBAbrir!txt_CNPJ_CPF = "" Or IsNull(TBAbrir!txt_CNPJ_CPF) = True Then
            If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & "CNPJ do destinatário da NF" Else Familiatext = "CNPJ do destinatário da NF"
            Permitido = False
        End If
    End If
    If TBAbrir!Id_Int_Cliente = "" Or IsNull(TBAbrir!Id_Int_Cliente) = True Or TBAbrir!txt_Razao_Nome = "" Or IsNull(TBAbrir!txt_Razao_Nome) = True Then
        If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & "destinatário da NF" Else Familiatext = "Destinatário da NF"
        Permitido = False
    End If
    If TBAbrir!txt_Endereco = "" Or IsNull(TBAbrir!txt_Endereco) = True Then
        If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & "Endereço do destinatário da NF" Else Familiatext = "Endereço do destinatário da NF"
        Permitido = False
    End If
    If TBAbrir!Numero = "" Or IsNull(TBAbrir!Numero) = True Then
        If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & "Número do destinatário da NF" Else Familiatext = "Número do destinatário da NF"
        Permitido = False
    End If
    If TBAbrir!Txt_bairro = "" Or IsNull(TBAbrir!Txt_bairro) = True Then
        If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & "Bairro do destinatário da NF" Else Familiatext = "Bairro do destinatário da NF"
        Permitido = False
    End If
    If TBAbrir!Txt_CEP = "" Or IsNull(TBAbrir!Txt_CEP) = True Then
        If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & "CEP do destinatário da NF" Else Familiatext = "CEP do destinatário da NF"
        Permitido = False
    End If
    If TBAbrir!txt_UF = "" Or IsNull(TBAbrir!txt_UF) = True Then
        If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & "UF do destinatário da NF" Else Familiatext = "UF do destinatário da NF"
        Permitido = False
    End If
End If

'Itens da nota
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from tbl_Detalhes_Nota where ID_Nota = " & ID_nota, Conexao, adOpenKeyset, adLockOptimistic
Do While TBAbrir.EOF = False
    Set TBFI = CreateObject("adodb.recordset")
    'TBFI.Open "Select Codigo_ref_DANFE from Empresa where Empresa = '" & Empresa & "' and Codigo_ref_DANFE = 'True'", Conexao, adOpenKeyset, adLockOptimistic
    TBFI.Open "Select CodRef from tbl_Dados_Nota_Fiscal_NFe where id_nota = " & ID_nota & " and CodRef = 'True'", Conexao, adOpenKeyset, adLockOptimistic
    If TBFI.EOF = False Then
        If TBAbrir!N_referencia = "" Or IsNull(TBAbrir!N_referencia) = True Then
            If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & "Código de referência do produto " & TBAbrir!int_Cod_Produto Else Familiatext = "Código de referência do produto " & TBAbrir!int_Cod_Produto
            Permitido = False
        End If
    End If
    TBFI.Close
    If TBAbrir!ID_CFOP = "" Or IsNull(TBAbrir!ID_CFOP) = True Then
        If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & "CFOP do produto " & TBAbrir!int_Cod_Produto Else Familiatext = "CFOP do produto " & TBAbrir!int_Cod_Produto
        Permitido = False
    End If
    If TBAbrir!ID_CF = "0" Or IsNull(TBAbrir!ID_CF) = True Then
        If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & "Código da classificação fiscal do produto " & TBAbrir!int_Cod_Produto Else Familiatext = "Código da classificação fiscal do produto " & TBAbrir!int_Cod_Produto
        Permitido = False
    End If
    If TBAbrir!txt_CST = "" Or IsNull(TBAbrir!txt_CST) = True Then
        If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & "CST de ICMS do produto " & TBAbrir!int_Cod_Produto Else Familiatext = "CST de ICMS do produto " & TBAbrir!int_Cod_Produto
        Permitido = False
    End If
    If TBAbrir!CST_IPI = "" Or IsNull(TBAbrir!CST_IPI) = True Then
        If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & "CST de IPI do produto " & TBAbrir!int_Cod_Produto Else Familiatext = "CST de IPI do produto " & TBAbrir!int_Cod_Produto
        Permitido = False
    End If
    If TBAbrir!CST_PIS = "" Or IsNull(TBAbrir!CST_PIS) = True Then
        If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & "CST de PIS do produto " & TBAbrir!int_Cod_Produto Else Familiatext = "CST de PIS do produto " & TBAbrir!int_Cod_Produto
        Permitido = False
    End If
    If TBAbrir!CST_Cofins = "" Or IsNull(TBAbrir!CST_Cofins) = True Then
        If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & "CST de Cofins do produto " & TBAbrir!int_Cod_Produto Else Familiatext = "CST de Cofins do produto " & TBAbrir!int_Cod_Produto
        Permitido = False
    End If
    TBAbrir.MoveNext
Loop

'Dados do transporte
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from tbl_dados_transp Where id_nota = " & ID_nota, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = True Then
    If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & "Frete por conta na transportadora" Else Familiatext = "Frete por conta na transportadora"
    Permitido = False
Else
    If TBAbrir!txt_Frete_Conta <> 0 And TBAbrir!txt_Frete_Conta <> 3 And TBAbrir!txt_Frete_Conta <> 9 Then
        If TBAbrir!txt_Razao = "" Or IsNull(TBAbrir!txt_Razao) = True Then
            If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & "Razão social da transportadora" Else Familiatext = "Razão social da transportadora"
            Permitido = False
        End If
        If TBAbrir!txt_Endereco = "" Or IsNull(TBAbrir!txt_Endereco) = True Then
            If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & "Endereço da transportadora" Else Familiatext = "Endereço da transportadora"
            Permitido = False
        End If
        If TBAbrir!int_numero = "" Or IsNull(TBAbrir!int_numero) = True Then
            If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & "Número da transportadora" Else Familiatext = "Número da transportadora"
            Permitido = False
        End If
        If TBAbrir!txt_Municipio = "" Or IsNull(TBAbrir!txt_Municipio) = True Then
            If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & "Cidade da transportadora" Else Familiatext = "Cidade da transportadora"
            Permitido = False
        End If
        If TBAbrir!txt_UF = "" Or IsNull(TBAbrir!txt_UF) = True Then
            If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & "UF da transportadora" Else Familiatext = "UF da transportadora"
            Permitido = False
        End If
        If TBAbrir!txt_CNPJ = "" Or IsNull(TBAbrir!txt_CNPJ) = True Then
            If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & "CNPJ da transportadora" Else Familiatext = "CNPJ da transportadora"
            Permitido = False
        End If
        If TBAbrir!txt_Placa <> "" And IsNull(TBAbrir!txt_Placa) = False Then
            If TBAbrir!txt_UF_Placa = "" Or IsNull(TBAbrir!txt_UF_Placa) = True Then
                If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & "UF da placa do veículo da transportadora" Else Familiatext = "UF da placa do veículo da transportadora"
                Permitido = False
            End If
        End If
        If TBAbrir!UF_embarque = "" Or IsNull(TBAbrir!UF_embarque) = True Then
            If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & "UF de embarque da transportadora" Else Familiatext = "UF de embarque da transportadora"
            Permitido = False
        End If
        If TBAbrir!Local_embarque = "" Or IsNull(TBAbrir!Local_embarque) = True Then
            If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & "Local de embarque da transportadora" Else Familiatext = "Local de embarque da transportadora"
            Permitido = False
        End If
    End If
End If

'Dados da nota fiscal nf-e
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from tbl_Dados_Nota_Fiscal_NFe where id_nota = " & ID_nota, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    If TBAbrir!FormaPagto = "" Or IsNull(TBAbrir!FormaPagto) = True Then
        If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & "Forma de pagamento" Else Familiatext = "Forma de pagamento"
        Permitido = False
    End If
    If TBAbrir!Forma_emissao = "" Or IsNull(TBAbrir!Forma_emissao) = True Then
        If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & "Forma de emissão" Else Familiatext = "Forma de emissão"
        Permitido = False
    End If
    If TBAbrir!Finalidade_emissao = "" Or IsNull(TBAbrir!Finalidade_emissao) = True Then
        If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & "Finalidade de emissão" Else Familiatext = "Finalidade de emissão"
        Permitido = False
    End If
    If TBAbrir!Enviar_Email = "" Or IsNull(TBAbrir!Enviar_Email) = True Then
        If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & "Arquivo que devera ser enviado por e-mail" Else Familiatext = "Arquivo que devera ser enviado por e-mail"
        Permitido = False
    End If
    If TBAbrir!ID_entrega = "" Or IsNull(TBAbrir!ID_entrega) = True Then
        If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & "Endereço de entrega" Else Familiatext = "Endereço de entrega"
        Permitido = False
    End If
End If
TBAbrir.Close

If Permitido = False And Mensagem = True Then
    MsgBox ("Informe o(s) campo(s) antes de liberar a NF para envio: " & vbCrLf & Familiatext), vbInformation
    Exit Sub
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal Key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcFiltrar
    Case 2: ProcSalvar
    Case 3: ProcLiberar
    Case 4: ProcCancelar
    Case 5: ProcAtualizarStatus
    'Case 7: ProcAjuda
    Case 8: ProcSair
End Select

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Private Sub USToolBar2_ButtonClick(ByVal ButtonIndex As Integer, ByVal Key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcSalvarProduto
    Case 2: ProcExcluirProduto
    'Case 4: ProcAjuda
    Case 5: ProcSair
End Select

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub
