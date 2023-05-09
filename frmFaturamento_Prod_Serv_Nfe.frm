VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2014.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form frmFaturamento_Prod_Serv_NFe 
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
   Icon            =   "frmFaturamento_Prod_Serv_Nfe.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   10035
   ScaleWidth      =   15360
   WindowState     =   2  'Maximizado
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   600
      Top             =   90
   End
   Begin DrawSuite2014.USProgressBar PBLista 
      Height          =   255
      Left            =   60
      TabIndex        =   70
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
   Begin TabDlg.SSTab SStab_nfe 
      Height          =   10035
      Left            =   0
      TabIndex        =   31
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
      TabPicture(0)   =   "frmFaturamento_Prod_Serv_Nfe.frx":1042
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
      TabPicture(1)   =   "frmFaturamento_Prod_Serv_Nfe.frx":105E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "listaProdutos"
      Tab(1).Control(1)=   "USToolBar2"
      Tab(1).Control(2)=   "FrameCST"
      Tab(1).Control(3)=   "txtID_item"
      Tab(1).Control(4)=   "Frame_comb_lub"
      Tab(1).ControlCount=   5
      Begin VB.TextBox Txt_ID_cobranca 
         Alignment       =   2  'Centralizar
         Height          =   315
         Left            =   2910
         TabIndex        =   56
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
         TabIndex        =   51
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
            TabIndex        =   66
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
            ItemData        =   "frmFaturamento_Prod_Serv_Nfe.frx":107A
            Left            =   12720
            List            =   "frmFaturamento_Prod_Serv_Nfe.frx":1090
            Locked          =   -1  'True
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   22
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
            ItemData        =   "frmFaturamento_Prod_Serv_Nfe.frx":10F0
            Left            =   180
            List            =   "frmFaturamento_Prod_Serv_Nfe.frx":10F2
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   20
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
            ItemData        =   "frmFaturamento_Prod_Serv_Nfe.frx":10F4
            Left            =   11790
            List            =   "frmFaturamento_Prod_Serv_Nfe.frx":10F6
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   21
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
            TabIndex        =   67
            Top             =   240
            Width           =   2100
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparente
            Caption         =   "Tipo do produto"
            Height          =   195
            Left            =   13297
            TabIndex        =   54
            Top             =   240
            Width           =   1140
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparente
            Caption         =   "UF cons."
            Height          =   195
            Left            =   11925
            TabIndex        =   53
            Top             =   240
            Width           =   630
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparente
            Caption         =   "Código do produto da ANP"
            Height          =   195
            Left            =   435
            TabIndex        =   52
            Top             =   240
            Width           =   1905
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   60
         TabIndex        =   48
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
            TabIndex        =   24
            Text            =   "30"
            ToolTipText     =   "Número de registros por página."
            Top             =   180
            Width           =   555
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
            TabIndex        =   25
            ToolTipText     =   "Número da página."
            Top             =   180
            Width           =   555
         End
         Begin DrawSuite2014.USButton cmdPagProx 
            Height          =   315
            Left            =   11760
            TabIndex        =   29
            ToolTipText     =   "Próxima página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmFaturamento_Prod_Serv_Nfe.frx":10F8
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
            TabIndex        =   28
            ToolTipText     =   "Página anterior."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmFaturamento_Prod_Serv_Nfe.frx":489C
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
            TabIndex        =   26
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
            TabIndex        =   27
            ToolTipText     =   "Primeira página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmFaturamento_Prod_Serv_Nfe.frx":83A5
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
            TabIndex        =   30
            ToolTipText     =   "Última página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmFaturamento_Prod_Serv_Nfe.frx":C494
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
            TabIndex        =   63
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
            TabIndex        =   55
            Top             =   240
            Width           =   645
         End
         Begin VB.Label lblRegistros 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparente
            Caption         =   "Nº de registros: 0"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   180
            TabIndex        =   50
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
            TabIndex        =   49
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.TextBox txtID_entrega 
         Alignment       =   2  'Centralizar
         Height          =   315
         Left            =   2550
         TabIndex        =   36
         Text            =   "0"
         Top             =   7530
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.TextBox txtID_nota 
         Alignment       =   2  'Centralizar
         Height          =   315
         Left            =   2130
         TabIndex        =   35
         Text            =   "0"
         Top             =   7530
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.TextBox txtID_item 
         Alignment       =   2  'Centralizar
         Height          =   335
         Left            =   -72870
         TabIndex        =   32
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
         TabIndex        =   33
         Top             =   1320
         Width           =   15195
         Begin VB.TextBox txtMotivo 
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
            Height          =   1560
            Left            =   9990
            Locked          =   -1  'True
            MaxLength       =   5000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   68
            TabStop         =   0   'False
            ToolTipText     =   "Motivo do cancelamento da nota fiscal."
            Top             =   990
            Width           =   4995
         End
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
            ItemData        =   "frmFaturamento_Prod_Serv_Nfe.frx":FD20
            Left            =   2190
            List            =   "frmFaturamento_Prod_Serv_Nfe.frx":FD2D
            Style           =   2  'Dropdown List
            TabIndex        =   4
            ToolTipText     =   "Indicador da forma de pagamento."
            Top             =   990
            Width           =   2175
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
            ItemData        =   "frmFaturamento_Prod_Serv_Nfe.frx":FD61
            Left            =   180
            List            =   "frmFaturamento_Prod_Serv_Nfe.frx":FD89
            Style           =   2  'Dropdown List
            TabIndex        =   3
            ToolTipText     =   "Forma de pagamento."
            Top             =   990
            Width           =   1995
         End
         Begin VB.CheckBox chkCodRef 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Utilizar código de referência na DANFE"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   11850
            TabIndex        =   7
            Top             =   458
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
            ItemData        =   "frmFaturamento_Prod_Serv_Nfe.frx":FE78
            Left            =   5790
            List            =   "frmFaturamento_Prod_Serv_Nfe.frx":FE91
            Style           =   2  'Dropdown List
            TabIndex        =   6
            ToolTipText     =   "Indicador de presença do comprador no estabelecimento comercial no momento da operação."
            Top             =   990
            Width           =   4155
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
            ItemData        =   "frmFaturamento_Prod_Serv_Nfe.frx":FFA2
            Left            =   4380
            List            =   "frmFaturamento_Prod_Serv_Nfe.frx":FFAC
            Style           =   2  'Dropdown List
            TabIndex        =   5
            ToolTipText     =   "Operação com consumidor final."
            Top             =   990
            Width           =   1395
         End
         Begin VB.CheckBox Chk_DA_cobranca 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Imprimir nos dados adicionais"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   7470
            TabIndex        =   11
            Top             =   1980
            Value           =   1  'Marcado
            Width           =   2685
         End
         Begin VB.CheckBox Chk_DA_entrega 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Imprimir nos dados adicionais"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   7470
            TabIndex        =   9
            Top             =   1380
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
            ItemData        =   "frmFaturamento_Prod_Serv_Nfe.frx":FFC2
            Left            =   180
            List            =   "frmFaturamento_Prod_Serv_Nfe.frx":FFC4
            Style           =   2  'Dropdown List
            TabIndex        =   8
            ToolTipText     =   "Endereço de entrega."
            Top             =   1590
            Width           =   9765
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
            ItemData        =   "frmFaturamento_Prod_Serv_Nfe.frx":FFC6
            Left            =   180
            List            =   "frmFaturamento_Prod_Serv_Nfe.frx":FFC8
            Style           =   2  'Dropdown List
            TabIndex        =   10
            ToolTipText     =   "Endereço de cobrança"
            Top             =   2190
            Width           =   9765
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
            ItemData        =   "frmFaturamento_Prod_Serv_Nfe.frx":FFCA
            Left            =   180
            List            =   "frmFaturamento_Prod_Serv_Nfe.frx":FFE3
            Style           =   2  'Dropdown List
            TabIndex        =   1
            ToolTipText     =   "Forma de emissão."
            Top             =   390
            Width           =   8715
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
            ItemData        =   "frmFaturamento_Prod_Serv_Nfe.frx":10214
            Left            =   8910
            List            =   "frmFaturamento_Prod_Serv_Nfe.frx":10224
            Style           =   2  'Dropdown List
            TabIndex        =   2
            ToolTipText     =   "Finalidade de emissão."
            Top             =   390
            Width           =   2775
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparente
            Caption         =   "Motivo do cancelamento"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   26
            Left            =   11580
            TabIndex        =   69
            Top             =   780
            Width           =   1740
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparente
            Caption         =   "Ind. da forma de pagto."
            Height          =   195
            Left            =   2407
            TabIndex        =   65
            Top             =   780
            Width           =   1740
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparente
            Caption         =   "Forma de pagamento"
            Height          =   195
            Left            =   412
            TabIndex        =   64
            Top             =   780
            Width           =   1530
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparente
            Caption         =   "Presença do comprador no estabelecimento"
            Height          =   195
            Index           =   4
            Left            =   6300
            TabIndex        =   60
            Top             =   780
            Width           =   3135
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparente
            Caption         =   "Consumidor final"
            Height          =   195
            Index           =   3
            Left            =   4485
            TabIndex        =   59
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
            Left            =   4260
            TabIndex        =   58
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
            Left            =   4305
            TabIndex        =   57
            Top             =   1380
            Width           =   1515
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparente
            Caption         =   "Forma de emissão"
            Height          =   195
            Left            =   3892
            TabIndex        =   44
            Top             =   180
            Width           =   1290
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparente
            Caption         =   "Finalidade de emissão"
            Height          =   195
            Index           =   0
            Left            =   9502
            TabIndex        =   34
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
         TabIndex        =   37
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
            ItemData        =   "frmFaturamento_Prod_Serv_Nfe.frx":10269
            Left            =   7590
            List            =   "frmFaturamento_Prod_Serv_Nfe.frx":1027F
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   19
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
            ItemData        =   "frmFaturamento_Prod_Serv_Nfe.frx":1032E
            Left            =   180
            List            =   "frmFaturamento_Prod_Serv_Nfe.frx":1033E
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   18
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
            TabIndex        =   39
            Top             =   240
            Width           =   2745
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparente
            Caption         =   "Modalidade de determinação da BC"
            Height          =   195
            Left            =   2617
            TabIndex        =   38
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
         TabIndex        =   40
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
            ItemData        =   "frmFaturamento_Prod_Serv_Nfe.frx":103AC
            Left            =   5280
            List            =   "frmFaturamento_Prod_Serv_Nfe.frx":10404
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   15
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
            TabIndex        =   14
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
            TabIndex        =   13
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
            TabIndex        =   17
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
            TabIndex        =   12
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
            TabIndex        =   16
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
            TabIndex        =   62
            Top             =   180
            Width           =   3315
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparente
            Caption         =   "UF"
            Height          =   195
            Left            =   5535
            TabIndex        =   61
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
            TabIndex        =   45
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
            TabIndex        =   43
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
            TabIndex        =   42
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
            TabIndex        =   41
            Top             =   180
            Width           =   465
         End
      End
      Begin DrawSuite2014.USToolBar USToolBar1 
         Height          =   975
         Left            =   60
         TabIndex        =   46
         Top             =   330
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
         ButtonCaption2  =   "Enviar"
         ButtonEnabled2  =   0   'False
         ButtonIconSize2 =   32
         ButtonToolTipText2=   "Enviar NFe para o Sefaz (F6)"
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
         ButtonWidth2    =   38
         ButtonHeight2   =   21
         ButtonUseMaskColor2=   0   'False
         ButtonCaption3  =   "Cancelar"
         ButtonEnabled3  =   0   'False
         ButtonIconSize3 =   32
         ButtonToolTipText3=   "Cancelar nota fiscal (F4)"
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
         ButtonLeft3     =   82
         ButtonTop3      =   2
         ButtonWidth3    =   50
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
         ButtonLeft4     =   134
         ButtonTop4      =   2
         ButtonWidth4    =   51
         ButtonHeight4   =   21
         ButtonUseMaskColor4=   0   'False
         ButtonCaption5  =   "Consultar status"
         ButtonEnabled5  =   0   'False
         ButtonIconSize5 =   32
         ButtonToolTipText5=   "Consultar status da nota fiscal.."
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
         ButtonLeft5     =   187
         ButtonTop5      =   2
         ButtonWidth5    =   87
         ButtonHeight5   =   21
         ButtonUseMaskColor5=   0   'False
         ButtonCaption6  =   "Log nota"
         ButtonEnabled6  =   0   'False
         ButtonIconSize6 =   32
         ButtonToolTipText6=   "Consulta logs da nota fiscal."
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
         ButtonLeft6     =   276
         ButtonTop6      =   2
         ButtonWidth6    =   50
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
         ButtonLeft7     =   328
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
         ButtonLeft8     =   332
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
         ButtonLeft9     =   370
         ButtonTop9      =   2
         ButtonWidth9    =   26
         ButtonHeight9   =   21
         ButtonUseMaskColor9=   0   'False
         ButtonEnabled10 =   0   'False
         ButtonIconSize10=   32
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
         ButtonState10   =   5
         ButtonLeft10    =   398
         ButtonTop10     =   2
         ButtonWidth10   =   24
         ButtonHeight10  =   24
         ButtonUseMaskColor10=   0   'False
         Begin DrawSuite2014.USImageList USImageList1 
            Left            =   9390
            Top             =   90
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmFaturamento_Prod_Serv_Nfe.frx":10476
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
            Object.Width           =   0
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
            Object.Width           =   3528
         EndProperty
      End
      Begin DrawSuite2014.USToolBar USToolBar2 
         Height          =   975
         Left            =   -74945
         TabIndex        =   47
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
            Img1            =   "frmFaturamento_Prod_Serv_Nfe.frx":15B69
            Count           =   1
         End
      End
      Begin MSComctlLib.ListView listaProdutos 
         Height          =   6530
         Left            =   -74940
         TabIndex        =   23
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
Attribute VB_Name = "frmFaturamento_Prod_Serv_NFe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim IDempresa_NF As String
Dim TipoXML As Integer
Dim Pais As String
Dim Codigo_pais As Long
Dim UF_transp As String
Dim Cidade As String
Dim Email As String
Dim TBLISTA_Faturamento_NFe As ADODB.Recordset

Dim CnpjNF As String
Dim DAPartilhaICMS As String
Dim TextoCancelamento As String

Dim objDom As DOMDocument50
Dim objEnviar As IXMLDOMElement
Dim objIde As IXMLDOMElement
Dim objNFRef As IXMLDOMElement
Dim objNFRefItem As IXMLDOMElement
Dim objEmit As IXMLDOMElement
Dim objEnderEmit As IXMLDOMElement
Dim objDest As IXMLDOMElement
Dim objEnderDest As IXMLDOMElement
Dim objAutXML As IXMLDOMElement
Dim objDet As IXMLDOMElement
Dim objDetItem As IXMLDOMElement
Dim objProd As IXMLDOMElement
Dim objDetDI As IXMLDOMElement
Dim objDetDIItem As IXMLDOMElement
Dim objDetAdicoes As IXMLDOMElement
Dim objDetAdicoesItem As IXMLDOMElement
Dim objComb As IXMLDOMElement
Dim objImposto As IXMLDOMElement
Dim objICMS As IXMLDOMElement
Dim objICMSUFDest As IXMLDOMElement
Dim objIPI As IXMLDOMElement
Dim objImpostoDevol As IXMLDOMElement
Dim objIPIDevol As IXMLDOMElement
Dim objII As IXMLDOMElement
Dim objCSTIPI As IXMLDOMElement
Dim objPis As IXMLDOMElement
Dim objCofins As IXMLDOMElement
Dim objTotal As IXMLDOMElement
Dim objICMStot As IXMLDOMElement
Dim objRetTrib As IXMLDOMElement
Dim objTransp As IXMLDOMElement
Dim objTransporta As IXMLDOMElement
Dim objVeicTransp As IXMLDOMElement
Dim objReboque As IXMLDOMElement
Dim objReboqueItem As IXMLDOMElement
Dim objVol As IXMLDOMElement
Dim objVolItem As IXMLDOMElement
Dim objCobr As IXMLDOMElement
Dim objFat As IXMLDOMElement
Dim objDup As IXMLDOMElement
Dim objDupItem As IXMLDOMElement
Dim objPag As IXMLDOMElement
Dim objPagItem As IXMLDOMElement
Dim objInfAdic As IXMLDOMElement
Dim objExporta As IXMLDOMElement
Dim objCompra As IXMLDOMElement

Public Sub procEnviar()
On Error GoTo tratar_erro

If Alterar = False Then
    MsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso."), vbExclamation
    Exit Sub
End If

If txtID_nota = 0 Then
    MsgBox ("Informe a nota fiscal antes de enviar."), vbExclamation
    Exit Sub
End If

Set TBproducao = CreateObject("adodb.recordset")
TBproducao.Open "Select status from tbl_Dados_Nota_Fiscal_NFe WHERE ID_nota = " & txtID_nota & " AND status IN (100,101)", Conexao, adOpenKeyset, adLockReadOnly
If TBproducao.EOF = False Then
    MsgBox ("Não é permitido enviar, pois a mesma já foi enviada."), vbExclamation
    TBproducao.Close
    Exit Sub
End If
TBproducao.Close

Acao = "enviar"
If funVerificaMigrate = False Then Exit Sub
If funVerificacaoEnviar = False Then Exit Sub

If MsgBox("Deseja realmente enviar esta nota fiscal?", vbQuestion + vbYesNo) = vbYes Then
    Set TBproducao = CreateObject("adodb.recordset")
    TBproducao.Open "Select NF.*, T.*, E.Simples, E.Simples1, E.Cultural, E.CNPJ, E.CNAE, E.Razao, E.Empresa, E.IM, E.ie, E.Tipo_endereco, E.Endereco, E.Numero as numeroEmpresa, E.Complemento, E.Tipo_bairro, E.Bairro, E.Cidade, E.UF, E.CEP, E.Telefone, E.Email, NFE.Consumidor_final, NFE.Presenca_comprador, NFE.Forma_emissao, NFE.Finalidade_emissao, NFE.Enviar_Email, NFE.Forma_pagamento, NFE.FormaPagto, NFE.DA_entrega, NFE.DA_cobranca, NFE.ID_entrega, NFE.ID_Cobranca from (tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_Totais_Nota T ON NF.ID = T.ID_nota) INNER JOIN tbl_Dados_Nota_Fiscal_NFe NFE ON NFE.ID_Nota = NF.ID INNER JOIN Empresa E ON NF.ID_empresa = E.Codigo WHERE NF.ID = " & txtID_nota, Conexao, adOpenKeyset, adLockReadOnly
    If TBproducao.EOF = False Then
        NomeArquivo = "NF" & txtNota & txtSerie
        procMontaEmail
        procCriarXML
        
        TipoXML = 1
        procAcionaTimer
    End If
    TBproducao.Close
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Private Sub procCriarXML()
On Error GoTo tratar_erro

Set objDom = New DOMDocument50

'nó Envio (A01)
Set objEnviar = objDom.createElement("Envio")
objDom.appendChild objEnviar
'Abre enviar ===================================================================================================
    objEnviar.appendChild objDom.createElement("ModeloDocumento")
    objEnviar.childNodes(0).Text = "NFe"
    objEnviar.appendChild objDom.createElement("Versao")
    objEnviar.childNodes(1).Text = "4.0"
    objEnviar.appendChild objDom.createElement("ChaveParceiro") 'Chave da caprind que a MIgrate emite
    objEnviar.childNodes(2).Text = "TsDpg/TtLpSXBO5uVUMM3w=="
    objEnviar.appendChild objDom.createElement("ChaveAcesso") 'Chave do cliente que a migrate emite
   
    procIdentificacaoXML
    procEmitenteXML
    procDestinatarioXML
    
    'Autorização para baixar o XML, hoje esta só a tranportadora
    Set TBTransporte = CreateObject("adodb.recordset")
    TBTransporte.Open "Select CF.Pessoa, DT.txt_CNPJ from tbl_Dados_Transp DT INNER JOIN Compras_fornecedores CF ON CF.IDCliente = DT.IdIntTransp and CF.Nome_Razao = DT.txt_Razao where DT.ID_Nota = " & txtID_nota & " and DT.txt_CNPJ IS NOT NULL and DT.txt_CNPJ <> N'' and DT.enviarXML = 1 AND txt_CNPJ <> '" & TBproducao!txt_CNPJ_CPF & "'", Conexao, adOpenKeyset, adLockReadOnly
    If TBTransporte.EOF = False Then
        'nó autXML dentro de Envio (A01)
        Set objAutXML = objDom.createElement("autXML")
        objEnviar.appendChild objAutXML
        'Abre autXML=================================================================================================
            objAutXML.appendChild objDom.createElement("CNPJ_aut") '0
            objAutXML.appendChild objDom.createElement("CPF_aut") '1
            If Left(TBTransporte!Pessoa, 1) = "J" Then objAutXML.childNodes(0).Text = DS_RetornarNumeros(TBTransporte!txt_CNPJ) Else objAutXML.childNodes(1).Text = DS_RetornarNumeros(TBTransporte!txt_CNPJ)
        'Fecha autXML================================================================================================
    End If
    TBTransporte.Close
    
    procProdutosXML
    procTotaisXML
    procTransporteXML
    
    'Forma de pagamento Novo layout da Sefaz (4.0)
    Set TBContas = CreateObject("adodb.recordset")
    TBContas.Open "Select * from tbl_Detalhes_Recebimento where ID_nota = " & txtID_nota & " order by ID", Conexao, adOpenKeyset, adLockReadOnly
    If TBContas.EOF = False Then
        'no cobr (Z01) dentro de enviar (A01)
        Set objCobr = objDom.createElement("cobr")
        objEnviar.appendChild objCobr
        'Abre cobr==================================================================================================
            'no fat (Z02) dentro de Cobr (Z01)
            Set objFat = objDom.createElement("fat")
            objCobr.appendChild objFat
            'Abre Fat==================================================================================================
                objFat.appendChild objDom.createElement("nFat") '0
                objFat.childNodes(0).Text = txtNota
                objFat.appendChild objDom.createElement("vOrig") '1
                objFat.childNodes(1).Text = Replace(IIf(IsNull(TBproducao!Valor_total_receber_pagar), 0, TBproducao!Valor_total_receber_pagar) + IIf(IsNull(TBproducao!Valor_total_desconto), 0, TBproducao!Valor_total_desconto), ",", ".")
                objFat.appendChild objDom.createElement("vDesc_cob") '2
                If IsNull(TBproducao!Valor_total_desconto) = False And TBproducao!Valor_total_desconto > 0 Then
                    objFat.childNodes(2).Text = Replace(TBproducao!Valor_total_desconto, ",", ".")
                Else
                    objFat.childNodes(2).Text = 0
                End If
                objFat.appendChild objDom.createElement("vLiq") '3
                objFat.childNodes(3).Text = Replace(IIf(IsNull(TBproducao!Valor_total_receber_pagar), 0, TBproducao!Valor_total_receber_pagar), ",", ".")
            'Fecha Fat=================================================================================================
            
            'no dup (Za01) dentro de Cobr (Z01)
            Set objDup = objDom.createElement("dup")
            objCobr.appendChild objDup
            'Abre Dup==================================================================================================
                Do While TBContas.EOF = False
                    'no dupItem (Za02) dentro de Cobr (Za01)
                    Set objDupItem = objDom.createElement("dupItem")
                    objDup.appendChild objDupItem
                    'Abre DupItem==================================================================================================
                        objDupItem.appendChild objDom.createElement("nDup") '0
                        objDupItem.childNodes(0).Text = Left(TBContas!txt_Parcela, 3)
                        objDupItem.appendChild objDom.createElement("dVenc") '1
                        objDupItem.childNodes(1).Text = Format(TBContas!dt_Vencimento, "yyyy-mm-dd")
                        objDupItem.appendChild objDom.createElement("vDup") '2
                        objDupItem.childNodes(2).Text = Replace(TBContas!dbl_Valor, ",", ".")
                    'Fecha DupItem=================================================================================================
                    TBContas.MoveNext
                Loop
            'Fecha Dup=================================================================================================
        'Fecha cobr=================================================================================================
    End If
    TBContas.Close
    
    'no Pag (AA01) dentro de Enviar (A01)
    Set objPag = objDom.createElement("pag")
    objEnviar.appendChild objPag
    'Abre Pag====================================================================================================
        'no pagItem (AA02) dentro de Pag (AA01)
        Set objPagItem = objDom.createElement("pagItem")
        objPag.appendChild objPagItem
        'Abre Pag====================================================================================================
            If IsNull(TBproducao!Forma_pagamento) = False Then
                objPagItem.appendChild objDom.createElement("indPag_pag") '0
                objPagItem.getElementsByTagName("indPag_pag").Item(0).Text = TBproducao!Forma_pagamento
            End If
            objPagItem.appendChild objDom.createElement("tPag") '1
            objPagItem.getElementsByTagName("tPag").Item(0).Text = IIf(IsNull(TBproducao!FormaPagto), "15", TBproducao!FormaPagto)
            If TBproducao!FormaPagto <> 90 Then
                objPagItem.appendChild objDom.createElement("vPag") '2
                objPagItem.getElementsByTagName("vPag").Item(0).Text = Replace(IIf(IsNull(TBproducao!dbl_Valor_Total_Nota), 0, TBproducao!dbl_Valor_Total_Nota), ",", ".")
            End If
        'Fecha Pag===================================================================================================
    'Fecha Pag===================================================================================================
    
    procAdicionaisXML

    If TBproducao!int_TipoNota = 1 And (IsNull(TBproducao!txt_UF) = True Or TBproducao!txt_UF = "" Or TBproducao!txt_UF = "EX") Then
        Set TBTransporte = CreateObject("adodb.recordset")
        TBTransporte.Open "Select UF_embarque, Local_embarque from tbl_Dados_Transp where ID_Nota = " & txtID_nota, Conexao, adOpenKeyset, adLockReadOnly
        If TBTransporte.EOF = False Then
            'nó exporta (AD01) dentro de Enviar (A01)
            Set objExporta = objDom.createElement("exporta")
            objEnviar.appendChild objExporta
            'Abre exporta====================================================================================================
                If IsNull(TBTransporte!UF_embarque) = False Then
                    objExporta.appendChild objDom.createElement("UFEmbarq")
                    objExporta.getElementsByTagName("UFEmbarq").Item(0).Text = TBTransporte!UF_embarque
                End If
                If IsNull(TBTransporte!Local_embarque) = False Then
                    objExporta.appendChild objDom.createElement("xLocEmbarq")
                    objExporta.getElementsByTagName("xLocEmbarq").Item(0).Text = TBTransporte!Local_embarque
                End If
            'Fecha exporta===================================================================================================
        End If
        TBTransporte.Close
    End If
    
    'no Compra (AE01) dentro de Enviar (A01)
    'Set objCompra = objDom.createElement("compra")
    'objEnviar.appendChild objCompra
    'Abre Compra====================================================================================================
    'Fecha Compra===================================================================================================
'Fecha Enviar====================================================================================================
                
objDom.Save (DiretorioEnvio & "/" & NomeArquivo & ".xml")

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

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

ProcCarregaToolBar1 Me, 15195, 10, True
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

ProcCarregaListaNota (1)
ProcLimpaCampos
With frmFaturamento_Prod_Serv
    If .TxtID <> "" And .TxtID <> "0" And .txtDtValidacao <> "" Then
        txtID_nota = .TxtID
        txtNota = IIf(.txtNFiscal = "", Null, .txtNFiscal)
        txtSerie = .txtSerie
        ProcCarregaEntrega
        ProcCarregaCobranca
        ProcPuxaDados
        procCarregaEmpresa
        procCarregaTransp
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
cmbFinalidade_emissao.Text = "1 - Normal"
Cmb_forma_de_emissao.Text = "1 - Normal"
cmbFormaPag.Text = "01 - Dinheiro"
cmbForma_pagamento.Text = "1 - pagamento à prazo"
Cmb_consumidor.Text = "1 - Sim"
Cmb_presenca_comprador.Text = "0 - Não se aplica"

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
            Case vbKeyF3: ProcSalvar
            Case vbKeyF4: ProcCancelar
            Case vbKeyF6: procEnviar
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

Set TBproducao = CreateObject("adodb.recordset")
TBproducao.Open "Select status from tbl_Dados_Nota_Fiscal_NFe WHERE ID_nota = " & txtID_nota & " AND status IN (100,101)", Conexao, adOpenKeyset, adLockReadOnly
If TBproducao.EOF = False Then
    MsgBox ("Não é permitido salvar, pois a mesma já foi enviada."), vbExclamation
    TBproducao.Close
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
    If .txtNFiscal = "" Then NomeCampo = "N° ordem: " & .TxtID Else NomeCampo = "N° nota: " & .txtNFiscal
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
        If .txtNFiscal = "" Then NomeCampo = "N° ordem: " & .TxtID Else NomeCampo = "N° nota: " & .txtNFiscal
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

Private Sub ListaNota_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If ListaNota.ListItems.Count = 0 Then Exit Sub
ProcLimpaCampos
CodigoLista = ListaNota.SelectedItem.index
txtID_nota = ListaNota.SelectedItem
txtNota = ListaNota.SelectedItem.SubItems(3)
txtSerie = ListaNota.SelectedItem.SubItems(5)

ProcCarregaEntrega
ProcCarregaCobranca
ProcPuxaDados
procCarregaEmpresa
procCarregaTransp

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
            
            Permitido = True
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select status from tbl_dados_nota_fiscal_NF where id = " & ID_nota & " AND status IN (100,101)", Conexao, adOpenKeyset, adLockReadOnly
            If TBAbrir.EOF = False Then
                Permitido = False
                Select Case TBGravar_NFe_Status!CbdStsRetCodigo
                    Case "100": NomeCampo = "autorizada no SEFAZ" 'Autorizado o uso da NF-e
                    Case "101": NomeCampo = "cancelada no SEFAZ" 'Cancelamento de NF-e homologado"
                End Select
            End If
            TBAbrir.Close
            
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
CodigoLista1 = listaProdutos.SelectedItem.index
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

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from tbl_Dados_Nota_Fiscal_NFe where id_nota = " & txtID_nota, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
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
End If
TBAbrir.Close

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
    If .txtNFiscal = "" Then NomeCampo = "N° ordem: " & .TxtID Else NomeCampo = "N° nota: " & .txtNFiscal
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

Function funVerifLiberacao(Mensagem As Boolean) As Boolean
On Error GoTo tratar_erro

funVerifLiberacao = True

Familiatext = ""
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from tbl_Dados_Nota_Fiscal_NFe where id_nota = " & txtID_nota, Conexao, adOpenKeyset, adLockReadOnly
If TBAbrir.EOF = True Then
    If Mensagem = True Then MsgBox ("Salve os dados da NF-e antes de liberar para envio."), vbExclamation
    funVerifLiberacao = False
    Exit Function
Else
    If IsNull(TBAbrir!Forma_emissao) = True Or TBAbrir!Forma_emissao = "" Then
        If Mensagem = True Then MsgBox ("Salve os dados da NF-e antes de liberar para envio."), vbExclamation
        funVerifLiberacao = False
        Exit Function
    End If
    If TBAbrir!Finalidade_emissao = 4 Then
        Set TBFI = CreateObject("adodb.recordset")
        TBFI.Open "Select Id from tbl_dados_transp where id_nota = " & txtID_nota & " and txt_Frete_Conta = 9", Conexao, adOpenKeyset, adLockReadOnly
        If TBFI.EOF = False Then
            If Mensagem = True Then MsgBox ("Frete inválido para o tipo de nota."), vbExclamation
            funVerifLiberacao = False
            Exit Function
        Else
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select Id from tbl_dados_transp where id_nota = " & txtID_nota, Conexao, adOpenKeyset, adLockReadOnly
            If TBFI.EOF = True Then
                If Mensagem = True Then MsgBox ("É necessário cadastrar a transportadora antes de liberar para envio."), vbExclamation
                funVerifLiberacao = False
                Exit Function
            End If
        End If
        TBFI.Close
    End If
End If

'Dados da nota fiscal
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from tbl_dados_nota_fiscal where id = " & txtID_nota, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    If TBAbrir!Serie = "" Or IsNull(TBAbrir!Serie) = True Then
        Familiatext = "Série da NF"
        funVerifLiberacao = False
    End If
    If TBAbrir!txt_UF <> "" And IsNull(TBAbrir!txt_UF) = False And TBAbrir!txt_UF <> "EX" Then
        If TBAbrir!txt_CNPJ_CPF = "" Or IsNull(TBAbrir!txt_CNPJ_CPF) = True Then
            If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & "CNPJ do destinatário da NF" Else Familiatext = "CNPJ do destinatário da NF"
            funVerifLiberacao = False
        End If
    End If
    If TBAbrir!Id_Int_Cliente = "" Or IsNull(TBAbrir!Id_Int_Cliente) = True Or TBAbrir!txt_Razao_Nome = "" Or IsNull(TBAbrir!txt_Razao_Nome) = True Then
        If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & "destinatário da NF" Else Familiatext = "Destinatário da NF"
        funVerifLiberacao = False
    End If
    If TBAbrir!txt_Endereco = "" Or IsNull(TBAbrir!txt_Endereco) = True Then
        If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & "Endereço do destinatário da NF" Else Familiatext = "Endereço do destinatário da NF"
        funVerifLiberacao = False
    End If
    If TBAbrir!Numero = "" Or IsNull(TBAbrir!Numero) = True Then
        If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & "Número do destinatário da NF" Else Familiatext = "Número do destinatário da NF"
        funVerifLiberacao = False
    End If
    If TBAbrir!Txt_bairro = "" Or IsNull(TBAbrir!Txt_bairro) = True Then
        If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & "Bairro do destinatário da NF" Else Familiatext = "Bairro do destinatário da NF"
        funVerifLiberacao = False
    End If
    If TBAbrir!Txt_CEP = "" Or IsNull(TBAbrir!Txt_CEP) = True Then
        If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & "CEP do destinatário da NF" Else Familiatext = "CEP do destinatário da NF"
        funVerifLiberacao = False
    End If
    If TBAbrir!txt_UF = "" Or IsNull(TBAbrir!txt_UF) = True Then
        If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & "UF do destinatário da NF" Else Familiatext = "UF do destinatário da NF"
        funVerifLiberacao = False
    End If
End If

'Itens da nota
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from tbl_Detalhes_Nota where ID_Nota = " & txtID_nota, Conexao, adOpenKeyset, adLockOptimistic
Do While TBAbrir.EOF = False
    Set TBFI = CreateObject("adodb.recordset")
    'TBFI.Open "Select Codigo_ref_DANFE from Empresa where Empresa = '" & Empresa & "' and Codigo_ref_DANFE = 'True'", Conexao, adOpenKeyset, adLockOptimistic
    TBFI.Open "Select CodRef from tbl_Dados_Nota_Fiscal_NFe where id_nota = " & txtID_nota & " and CodRef = 'True'", Conexao, adOpenKeyset, adLockOptimistic
    If TBFI.EOF = False Then
        If TBAbrir!N_referencia = "" Or IsNull(TBAbrir!N_referencia) = True Then
            If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & "Código de referência do produto " & TBAbrir!int_Cod_Produto Else Familiatext = "Código de referência do produto " & TBAbrir!int_Cod_Produto
            funVerifLiberacao = False
        End If
    End If
    TBFI.Close
    If TBAbrir!ID_CFOP = "" Or IsNull(TBAbrir!ID_CFOP) = True Then
        If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & "CFOP do produto " & TBAbrir!int_Cod_Produto Else Familiatext = "CFOP do produto " & TBAbrir!int_Cod_Produto
        funVerifLiberacao = False
    End If
    If TBAbrir!ID_CF = "0" Or IsNull(TBAbrir!ID_CF) = True Then
        If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & "Código da classificação fiscal do produto " & TBAbrir!int_Cod_Produto Else Familiatext = "Código da classificação fiscal do produto " & TBAbrir!int_Cod_Produto
        funVerifLiberacao = False
    End If
    If TBAbrir!txt_CST = "" Or IsNull(TBAbrir!txt_CST) = True Then
        If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & "CST de ICMS do produto " & TBAbrir!int_Cod_Produto Else Familiatext = "CST de ICMS do produto " & TBAbrir!int_Cod_Produto
        funVerifLiberacao = False
    End If
    If TBAbrir!CST_IPI = "" Or IsNull(TBAbrir!CST_IPI) = True Then
        If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & "CST de IPI do produto " & TBAbrir!int_Cod_Produto Else Familiatext = "CST de IPI do produto " & TBAbrir!int_Cod_Produto
        funVerifLiberacao = False
    End If
    If TBAbrir!CST_PIS = "" Or IsNull(TBAbrir!CST_PIS) = True Then
        If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & "CST de PIS do produto " & TBAbrir!int_Cod_Produto Else Familiatext = "CST de PIS do produto " & TBAbrir!int_Cod_Produto
        funVerifLiberacao = False
    End If
    If TBAbrir!CST_Cofins = "" Or IsNull(TBAbrir!CST_Cofins) = True Then
        If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & "CST de Cofins do produto " & TBAbrir!int_Cod_Produto Else Familiatext = "CST de Cofins do produto " & TBAbrir!int_Cod_Produto
        funVerifLiberacao = False
    End If
    TBAbrir.MoveNext
Loop

'Dados do transporte
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from tbl_dados_transp Where id_nota = " & txtID_nota, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = True Then
    If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & "Frete por conta na transportadora" Else Familiatext = "Frete por conta na transportadora"
    funVerifLiberacao = False
Else
    If TBAbrir!txt_Frete_Conta <> 0 And TBAbrir!txt_Frete_Conta <> 3 And TBAbrir!txt_Frete_Conta <> 9 Then
        If TBAbrir!txt_Razao = "" Or IsNull(TBAbrir!txt_Razao) = True Then
            If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & "Razão social da transportadora" Else Familiatext = "Razão social da transportadora"
            funVerifLiberacao = False
        End If
        If TBAbrir!txt_Endereco = "" Or IsNull(TBAbrir!txt_Endereco) = True Then
            If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & "Endereço da transportadora" Else Familiatext = "Endereço da transportadora"
            funVerifLiberacao = False
        End If
        If TBAbrir!int_numero = "" Or IsNull(TBAbrir!int_numero) = True Then
            If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & "Número da transportadora" Else Familiatext = "Número da transportadora"
            funVerifLiberacao = False
        End If
        If TBAbrir!txt_Municipio = "" Or IsNull(TBAbrir!txt_Municipio) = True Then
            If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & "Cidade da transportadora" Else Familiatext = "Cidade da transportadora"
            funVerifLiberacao = False
        End If
        If TBAbrir!txt_UF = "" Or IsNull(TBAbrir!txt_UF) = True Then
            If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & "UF da transportadora" Else Familiatext = "UF da transportadora"
            funVerifLiberacao = False
        End If
        If TBAbrir!txt_CNPJ = "" Or IsNull(TBAbrir!txt_CNPJ) = True Then
            If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & "CNPJ da transportadora" Else Familiatext = "CNPJ da transportadora"
            funVerifLiberacao = False
        End If
        If TBAbrir!txt_Placa <> "" And IsNull(TBAbrir!txt_Placa) = False Then
            If TBAbrir!txt_UF_Placa = "" Or IsNull(TBAbrir!txt_UF_Placa) = True Then
                If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & "UF da placa do veículo da transportadora" Else Familiatext = "UF da placa do veículo da transportadora"
                funVerifLiberacao = False
            End If
        End If
        If TBAbrir!UF_embarque = "" Or IsNull(TBAbrir!UF_embarque) = True Then
            If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & "UF de embarque da transportadora" Else Familiatext = "UF de embarque da transportadora"
            funVerifLiberacao = False
        End If
        If TBAbrir!Local_embarque = "" Or IsNull(TBAbrir!Local_embarque) = True Then
            If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & "Local de embarque da transportadora" Else Familiatext = "Local de embarque da transportadora"
            funVerifLiberacao = False
        End If
    End If
End If

'Dados da nota fiscal nf-e
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from tbl_Dados_Nota_Fiscal_NFe where id_nota = " & txtID_nota, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    If TBAbrir!FormaPagto = "" Or IsNull(TBAbrir!FormaPagto) = True Then
        If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & "Forma de pagamento" Else Familiatext = "Forma de pagamento"
        funVerifLiberacao = False
    End If
    If TBAbrir!Forma_emissao = "" Or IsNull(TBAbrir!Forma_emissao) = True Then
        If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & "Forma de emissão" Else Familiatext = "Forma de emissão"
        funVerifLiberacao = False
    End If
    If TBAbrir!Finalidade_emissao = "" Or IsNull(TBAbrir!Finalidade_emissao) = True Then
        If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & "Finalidade de emissão" Else Familiatext = "Finalidade de emissão"
        funVerifLiberacao = False
    End If
'    If TBAbrir!Enviar_Email = "" Or IsNull(TBAbrir!Enviar_Email) = True Then
'        If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & "Arquivo que devera ser enviado por e-mail" Else Familiatext = "Arquivo que devera ser enviado por e-mail"
'        funVerifLiberacao = False
'    End If
    If TBAbrir!ID_entrega = "" Or IsNull(TBAbrir!ID_entrega) = True Then
        If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & "Endereço de entrega" Else Familiatext = "Endereço de entrega"
        funVerifLiberacao = False
    End If
End If
TBAbrir.Close

If funVerifLiberacao = False And Mensagem = True Then MsgBox ("Informe o(s) campo(s) antes de liberar a NF para envio: " & vbCrLf & Familiatext), vbInformation

Exit Function
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Function

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcSalvar
    Case 2: procEnviar
    Case 3: ProcCancelar
    Case 4: ProcImprimir
    Case 5: procConsultar
    Case 6: procLogErros
    'Case 8: ProcAjuda
    Case 9: ProcSair
End Select

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Private Sub USToolBar2_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
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

Public Sub procLogErros()
On Error GoTo tratar_erro

If txtID_nota = 0 Then
    MsgBox ("Informe a nota fiscal antes de consultar log de erros."), vbExclamation
    Exit Sub
End If

Acao = "verificar o log"
If funVerificaMigrate = False Then Exit Sub

Sit_REG = 1
frmFaturamento_Prod_Serv_NFSe_Log.Show 1
    
Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Function funVerificaMigrate() As Boolean
On Error GoTo tratar_erro
funVerificaMigrate = False


If DiretorioEnvio = "" Then
    NomeCampo = "o diretório de envio no cadastro da empresa"
    ProcVerificaAcao
    Exit Function
End If

If DiretorioRetorno = "" Then
    NomeCampo = "o diretório de retorno no cadastro da empresa"
    ProcVerificaAcao
    Exit Function
End If

If DiretorioXMLDanfe = "" Then
    NomeCampo = "o diretório de XML e Danfe no cadastro da empresa"
    ProcVerificaAcao
    Exit Function
End If

funVerificaMigrate = True

Exit Function
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Function

Public Sub procLerRetornoXML()
On Error GoTo tratar_erro
Dim doc As New DOMDocument50
Dim success As Boolean
Dim statusXML As String
Dim chaveAcessoXML As String
Dim cnpjXML As String, NotaXML As String, SerieXML As String

'tipo
'1 - envio
'2 - cancelamento
'3 - consulta

'Retorno de envio 000000
'Retorno de cancelamento 11011101
cnpjXML = DS_RetornarNumeros(CnpjNF)
NotaXML = FunTamanhoTextoZeroEsq(txtNota, 9)
SerieXML = FunTamanhoTextoZeroEsq(txtSerie, 5)
formaXML = False

success = doc.Load(DiretorioRetorno & "\NFe\" & cnpjXML & NotaXML & SerieXML & IIf(TipoXML = 2, "11011101", "00000000") & "-ret.xml")
If success = False Then
    MsgBox "Não foi possível obter retorno da Sefaz, favor consultar o log de erros.", vbExclamation
    If TipoXML <> 2 Then statusXML = 0 Else statusXML = ""
Else
    Dim NodeStatus As IXMLDOMNode
    Dim NodeDescricao As IXMLDOMNode
    Dim NodeChave As IXMLDOMNode

    Set NodeStatus = doc.selectSingleNode("/Documento/DocSitCodigo")
    Set NodeDescricao = doc.selectSingleNode("/Documento/DocSitDescricao")
    If NodeStatus Is Nothing Then
        Set NodeStatus = doc.selectSingleNode("/Documento/Situacao/SitCodigo")
        Set NodeDescricao = doc.selectSingleNode("/Documento/Situacao/SitDescricao")
    End If
    
    MsgBox (NodeStatus.Text & " - " & NodeDescricao.Text & "."), vbInformation
    statusXML = NodeStatus.Text
    If TipoXML = 2 And statusXML <> 101 Then statusXML = 100 'verifica se é cancelamento e se cancelou mesmo, caso não alterou mantem o status 100 de aprovado
    
    Set NodeChave = doc.selectSingleNode("/Documento/DocChaAcesso")
    chaveAcessoXML = NodeChave.Text
End If

If statusXML <> "" Then
    Conexao.Execute "Update tbl_dados_nota_fiscal_NFe Set Status = " & statusXML & ", chave_acesso = '" & chaveAcessoXML & "' where id_nota = " & txtID_nota
    If TipoXML = 2 Then Conexao.Execute "Update tbl_dados_nota_fiscal Set Obs = '" & TextoCancelamento & "' where id = " & txtID_nota
    
    If statusXML = 101 Then
        Conexao.Execute "Update tbl_dados_nota_fiscal Set Int_status = 2 where id = " & txtID_nota
        procCancelarTabelas
    End If
    
    If statusXML = 100 Then Conexao.Execute "Update tbl_dados_nota_fiscal Set Imprimir = 1 where id = " & txtID_nota
    txtStatus = FunVerifStatusNFe(txtID_nota)
    Txt_chave_acesso = chaveAcessoXML
    
    If TipoXML = 1 And statusXML = 100 Or TipoXML = 2 And statusXML = 101 Then
        If MsgBox("Deseja visualizar a Danfe?", vbQuestion + vbYesNo) = vbYes Then procAbrirNotaPDF "NFe", CnpjNF, txtNota, txtSerie, DiretorioXMLDanfe, False
    End If
End If

ProcCarregaListaNota (IIf(DS_RetornarNumeros(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, DS_RetornarNumeros(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
With frmFaturamento_Prod_Serv
    .ProcCarregaListaNota (IIf(DS_RetornarNumeros(Left(.lblPaginas(1).Caption, Len(.lblPaginas(1).Caption) - 5)) <= 1, 1, DS_RetornarNumeros(Left(.lblPaginas(1).Caption, Len(.lblPaginas(1).Caption) - 5))))
End With

Exit Sub
tratar_erro:
    If Err.Number = 91 Then
        MsgBox "Não foi possível ter um retorno da Sefaz, favor tentar mais tarde.", vbExclamation
    Else
        MsgBox ("Descrição do erro : " + Error()), vbCritical
    End If
End Sub

Public Sub procConsultar()
On Error GoTo tratar_erro

If Alterar = False Then
    MsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso."), vbExclamation
    Exit Sub
End If

If txtID_nota = 0 Then
    MsgBox ("Informe a nota fiscal antes de consultar o status."), vbExclamation
    Exit Sub
End If

Acao = "consultar"
If funVerificaMigrate = False Then Exit Sub

If MsgBox("Deseja consultar esta nota fiscal?", vbQuestion + vbYesNo) = vbYes Then
    NomeArquivo = "NF" & txtNota & txtSerie & "CON"
    procConsultarXML
    TipoXML = 3
    procAcionaTimer
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Public Sub procConsultarXML()
On Error GoTo tratar_erro

Dim objDom As DOMDocument50
Dim objConsultar As IXMLDOMElement
Dim objConsulta As IXMLDOMElement
Dim objParametrosConsulta As IXMLDOMElement
      
Set objDom = New DOMDocument50
'nó Consultar
Set objConsultar = objDom.createElement("Consultar")
objDom.appendChild objConsultar
'Abre Consultar======================================================================================================
    'filhos dentro de Consultar
    objConsultar.appendChild objDom.createElement("ChaveParceiro") 'Chave da caprind que a Migrate emite
    objConsultar.childNodes(0).Text = "TsDpg/TtLpSXBO5uVUMM3w=="
    objConsultar.appendChild objDom.createElement("ChaveAcesso") 'Chave do cliente que a Migrate emite
    
    'nó Consulta
    Set objConsulta = objDom.createElement("Consulta")
    objConsultar.appendChild objConsulta
    'Abre Consulta===================================================================================================
        objConsulta.appendChild objDom.createElement("ModeloDocumento")
        objConsulta.childNodes(0).Text = "NFe"
        objConsulta.appendChild objDom.createElement("Versao")
        objConsulta.childNodes(1).Text = "4.0"
        objConsulta.appendChild objDom.createElement("tpAmb")
        objConsulta.childNodes(2).Text = 1 '1-produção 2-Homologação
        objConsulta.appendChild objDom.createElement("CnpjEmissor")
        objConsulta.childNodes(3).Text = DS_RetornarNumeros(CnpjNF)
        objConsulta.appendChild objDom.createElement("NumeroInicial")
        objConsulta.childNodes(4).Text = txtNota
        objConsulta.appendChild objDom.createElement("NumeroFinal")
        objConsulta.childNodes(5).Text = txtNota
        objConsulta.appendChild objDom.createElement("Serie")
        objConsulta.childNodes(6).Text = txtSerie
        objConsulta.appendChild objDom.createElement("ChaveAcesso")
        objConsulta.childNodes(7).Text = Txt_chave_acesso
        objConsulta.appendChild objDom.createElement("DataEmissaoInicial")
        objConsulta.appendChild objDom.createElement("DataEmissaoFinal")
    'Fecha Consulta==================================================================================================
    
    'nó ParametrosConsulta
    Set objParametrosConsulta = objDom.createElement("ParametrosConsulta")
    objConsultar.appendChild objParametrosConsulta
    'Abre ParametrosConsulta===================================================================================================
        objParametrosConsulta.appendChild objDom.createElement("Situacao") '0
        'objParametrosConsulta.childNodes(0).Text = "S"
        objParametrosConsulta.appendChild objDom.createElement("XMLCompleto") '1
        objParametrosConsulta.childNodes(1).Text = "S"
        objParametrosConsulta.appendChild objDom.createElement("XMLLink") '2
        objParametrosConsulta.appendChild objDom.createElement("PDFBase64") '3
        objParametrosConsulta.childNodes(3).Text = "S"
        objParametrosConsulta.appendChild objDom.createElement("PDFLink") '4
        objParametrosConsulta.appendChild objDom.createElement("Eventos") '5
        objParametrosConsulta.childNodes(5).Text = "S"
    'Fecha ParametrosConsulta==================================================================================================
'Fecha Consultar=====================================================================================================
objDom.Save (DiretorioEnvio & "/" & NomeArquivo & ".xml")

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Public Sub procProdutosXML()
On Error GoTo tratar_erro

'nó det dentro de Envio (A01)
Set objDet = objDom.createElement("det")
objEnviar.appendChild objDet
'Abre det=================================================================================================
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select N.*, NFE.Documento_importacao, NFE.Numero_adicao, NFE.Numero_sequencial, NFE.Codigo_fabricante, NFE.Data_registro, NFE.Data_desembaraco, NFE.Local_desembaraco, NFE.UF_desembaraco, NFE.Codigo_exportador, NFE.Via_transp, NFE.Valor_AFRMM, NFE.Forma_imp, NFE.Tipo_produto, NFE.Codigo_ANP, NFE.UF_consumo, NFE.Descricao_ANP from tbl_Detalhes_Nota N LEFT JOIN tbl_Detalhes_Nota_NFe NFE ON N.Int_codigo = NFE.ID_item where N.ID_nota = " & txtID_nota & " order by N.Int_codigo", Conexao, adOpenKeyset, adLockReadOnly
    Do While TBProduto.EOF = False
        'nó detItem dentro de Det
        Set objDetItem = objDom.createElement("detItem")
        objDet.appendChild objDetItem
        'Abre detItem==================================================================================================
            
            If IsNull(TBProduto!Inf_adicionais_prod) = False And TBProduto!Inf_adicionais_prod <> "" Then
                objDetItem.appendChild objDom.createElement("infADProd")
                objDetItem.getElementsByTagName("infADProd").Item(0).Text = Trim(TBProduto!Inf_adicionais_prod)
            End If
        
            'nó prod dentro de DetItem
            Set objProd = objDom.createElement("prod")
            objDetItem.appendChild objProd
            'Abre prod==================================================================================================
                objProd.appendChild objDom.createElement("cProd") '0
                'Verifica se é para utilizar o código de referência na DANFE
                Set TBCodigoDesc = CreateObject("adodb.recordset")
                TBCodigoDesc.Open "Select CodRef from tbl_Dados_Nota_Fiscal_NFe where ID_Nota = " & txtID_nota, Conexao, adOpenKeyset, adLockReadOnly
                If TBCodigoDesc.EOF = False Then
                    If TBCodigoDesc!CodRef = False Or IsNull(TBCodigoDesc!CodRef) = True Then
                        objProd.getElementsByTagName("cProd").Item(0).Text = Trim(FunTiraAcentosTexto(TBProduto!int_Cod_Produto))
                        Set TBCodigoDesc = CreateObject("adodb.recordset")
                        TBCodigoDesc.Open "Select Codigo_ref_desc_DANFE from empresa where codigo = " & TBproducao!ID_empresa, Conexao, adOpenKeyset, adLockReadOnly
                        If TBCodigoDesc!Codigo_ref_desc_DANFE = True Then
                            CodRef = 2
                        Else
                            CodRef = 0
                        End If
                    Else
                        objProd.getElementsByTagName("cProd").Item(0).Text = Trim(FunTiraAcentosTexto(TBProduto!N_referencia))
                        CodRef = 1
                    End If
                End If
                TBCodigoDesc.Close
                
                objProd.appendChild objDom.createElement("cEAN") '1
                objProd.appendChild objDom.createElement("xProd") '2
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
                        If Len(TBproducao!txt_tipocliente) = 2 Then TipoFiltro = "C" Else TipoFiltro = "F"
                        Set TBItem = CreateObject("adodb.recordset")
                        TBItem.Open "Select IA.N_Referencia from item_aplicacoes IA INNER JOIN projproduto P ON IA.codproduto = P.codproduto where P.Desenho = '" & TBProduto!int_Cod_Produto & "' and IA.ID_cliente_forn = " & TBproducao!Id_Int_Cliente & " and IA.Tipo = '" & TipoFiltro & "' and IA.N_Referencia IS NOT NULL and IA.N_Referencia <> '" & TBProduto!int_Cod_Produto & "'", Conexao, adOpenKeyset, adLockReadOnly
                        If TBItem.EOF = False Then
                            DesenhoProduto = "(" & TBItem!N_referencia & ") - " & DesenhoProduto
                        End If
                        TBItem.Close
                    End If
                End If
                If IsNull(TBProduto!Complemento_descricao) = False And TBProduto!Complemento_descricao <> "" Then DesenhoProduto = DesenhoProduto & " - " & Trim(TBProduto!Complemento_descricao)
                If IsNull(TBProduto!PCCliente) = False And TBProduto!PCCliente <> "" Then DesenhoProduto = DesenhoProduto & " - Ped. " & Trim(TBProduto!PCCliente)
                If IsNull(TBProduto!N_item) = False And TBProduto!N_item <> "" Then DesenhoProduto = DesenhoProduto & " - N. item " & Trim(TBProduto!N_item)
                objProd.getElementsByTagName("xProd").Item(0).Text = FunTiraAcentosTexto(Left(DesenhoProduto, 120)) 'Descrição
                
                objProd.appendChild objDom.createElement("NCM") '3
                
                Set TBControleNF = CreateObject("adodb.recordset")
                TBControleNF.Open "Select id_CFOP, Devolucao from tbl_NaturezaOperacao where IDCountCfop = " & IIf(IsNull(TBProduto!ID_CFOP), 0, TBProduto!ID_CFOP), Conexao, adOpenKeyset, adLockReadOnly
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
                    objProd.appendChild objDom.createElement("CFOP") '4
                    objProd.getElementsByTagName("CFOP").Item(0).Text = CFOP_Produto 'CFOP
                    
                    If TBControleNF!Devolucao = True Then Devolucao = True Else Devolucao = False
                End If
                TBControleNF.Close
                
                objProd.appendChild objDom.createElement("uCOM") '5
                objProd.getElementsByTagName("uCOM").Item(0).Text = FunTiraAcentosTexto(TBProduto!Unidade_com)
                objProd.appendChild objDom.createElement("qCOM") '6
                objProd.getElementsByTagName("qCOM").Item(0).Text = Replace(TBProduto!int_Qtd, ",", ".")
                objProd.appendChild objDom.createElement("vUnCom") '7
                objProd.getElementsByTagName("vUnCom").Item(0).Text = Replace(TBProduto!dbl_ValorUnitario, ",", ".")
                objProd.appendChild objDom.createElement("vProd") '8
                objProd.getElementsByTagName("vProd").Item(0).Text = Replace(TBProduto!dbl_ValorTotal, ",", ".")
                
                objProd.appendChild objDom.createElement("cEANTrib") '9
                If TBProduto!GTIN = "" Or IsNull(TBProduto!GTIN) = True Then
                    objProd.getElementsByTagName("cEAN").Item(0).Text = "SEM GTIN"
                    objProd.getElementsByTagName("cEANTrib").Item(0).Text = "SEM GTIN"
                Else
                    objProd.getElementsByTagName("cEAN").Item(0).Text = FunTiraAcentosTexto(TBProduto!GTIN)
                    objProd.getElementsByTagName("cEANTrib").Item(0).Text = FunTiraAcentosTexto(TBProduto!GTIN)
                End If
                
                objProd.appendChild objDom.createElement("uTrib") '10
                objProd.getElementsByTagName("uTrib").Item(0).Text = FunTiraAcentosTexto(TBProduto!Unidade_com)
                objProd.appendChild objDom.createElement("qTrib") '11
                objProd.getElementsByTagName("qTrib").Item(0).Text = Replace(TBProduto!int_Qtd, ",", ".")
                objProd.appendChild objDom.createElement("vUnTrib") '12
                objProd.getElementsByTagName("vUnTrib").Item(0).Text = Replace(TBProduto!dbl_ValorUnitario, ",", ".")
                objProd.appendChild objDom.createElement("vFrete") '13
                objProd.getElementsByTagName("vFrete").Item(0).Text = Replace(IIf(IsNull(TBProduto!Valor_frete), 0, TBProduto!Valor_frete), ",", ".")
                objProd.appendChild objDom.createElement("vSeg") '14
                objProd.getElementsByTagName("vSeg").Item(0).Text = Replace(IIf(IsNull(TBProduto!Valor_seguro), 0, TBProduto!Valor_seguro), ",", ".")
                objProd.appendChild objDom.createElement("vDesc") '15
                objProd.getElementsByTagName("vDesc").Item(0).Text = Replace(IIf(IsNull(TBProduto!Valor_desconto), 0, TBProduto!Valor_desconto) + IIf(IsNull(TBProduto!Valor_desconto_SUFRAMA), 0, TBProduto!Valor_desconto_SUFRAMA), ",", ".")
                objProd.appendChild objDom.createElement("vOutro_item") '16
                objProd.getElementsByTagName("vOutro_item").Item(0).Text = Replace(IIf(IsNull(TBProduto!Valor_acessorias), 0, TBProduto!Valor_acessorias), ",", ".")
                
                objProd.appendChild objDom.createElement("indTot") '17
                If TBProduto!retorno = True Then
                    Set TBFIltro = CreateObject("adodb.recordset")
                    TBFIltro.Open "Select * from tbl_NaturezaOperacao where IDCountCfop = " & TBProduto!ID_CFOP & " and Soma_retorno_totalnf = 1", Conexao, adOpenKeyset, adLockReadOnly
                    If TBFIltro.EOF = False Then
                        objProd.getElementsByTagName("indTot").Item(0).Text = 1 'O valor do produto compõe o valor total da NF
                    Else
                        objProd.getElementsByTagName("indTot").Item(0).Text = 0 'O valor do produto não compõe o valor total da NF
                    End If
                    TBFIltro.Close
                Else
                    objProd.getElementsByTagName("indTot").Item(0).Text = 1 'O valor do produto compõe o valor total da NF
                End If
                
                objProd.appendChild objDom.createElement("nTipoItem") '18
                objProd.getElementsByTagName("nTipoItem").Item(0).Text = IIf(IsNull(TBProduto!Tipo_produto), "0", TBProduto!Tipo_produto)
                
                If IsNull(TBProduto!PCCliente) = False And TBProduto!PCCliente <> "" Then
                    objProd.appendChild objDom.createElement("xPed_item")
                    objProd.getElementsByTagName("xPed_item").Item(0).Text = Left(Trim(TBProduto!PCCliente), 15)
                End If
                
                If IsNull(TBProduto!N_item) = False And TBProduto!N_item <> "" Then
                    objProd.appendChild objDom.createElement("nItemPed") '20
                    objProd.getElementsByTagName("nItemPed").Item(0).Text = DS_RetornarNumeros(TBProduto!N_item)
                End If
                
                Set TBControleNF = CreateObject("adodb.recordset")
                TBControleNF.Open "Select IDIntClasse, CEST from tbl_ClassificacaoFiscal where Idclass = " & TBProduto!ID_CF, Conexao, adOpenKeyset, adLockReadOnly
                If TBControleNF.EOF = False Then
                    objProd.appendChild objDom.createElement("NCM") '3
                    If TBControleNF!IDIntClasse = "0000.00.00" Then
                        objProd.getElementsByTagName("NCM").Item(0).Text = "00" 'CFOP
                    Else
                        objProd.getElementsByTagName("NCM").Item(0).Text = DS_RetornarNumeros(TBControleNF!IDIntClasse) 'CFOP
                    End If
                    If IsNull(TBControleNF!CEST) = False Then
                        objProd.appendChild objDom.createElement("CEST") '21
                        objProd.getElementsByTagName("CEST").Item(0).Text = DS_RetornarNumeros(TBControleNF!CEST)
                    End If
                End If
                
                If IsNull(TBProduto!Documento_importacao) = False And TBProduto!Documento_importacao <> "" And IsNull(TBProduto!Numero_adicao) = False And TBProduto!Numero_adicao <> "" And IsNull(TBProduto!Numero_sequencial) = False And TBProduto!Numero_sequencial <> "" And IsNull(TBProduto!Codigo_fabricante) = False And TBProduto!Codigo_fabricante <> "" Then
                    'nó detDI dentro de Prod
                    Set objDetDI = objDom.createElement("detDI")
                    objProd.appendChild objDetDI
                    'Abre objDetDI==================================================================================================
                        'nó DetDIItem dentro de DetDI
                        Set objDetDIItem = objDom.createElement("detDIItem")
                        objDetDI.appendChild objDetDIItem
                        'Abre objDetDIItem==================================================================================================
                            objDetDIItem.appendChild objDom.createElement("nDI") '0
                            objDetDIItem.childNodes(0).Text = TBProduto!Documento_importacao
                            objDetDIItem.appendChild objDom.createElement("dDi") '1
                            objDetDIItem.childNodes(1).Text = Format(TBProduto!Data_registro, "yyyy-mm-dd")
                            objDetDIItem.appendChild objDom.createElement("xLocDesemb") '2
                            objDetDIItem.childNodes(2).Text = TBProduto!Local_desembaraco
                            objDetDIItem.appendChild objDom.createElement("UFDesemb") '3
                            objDetDIItem.childNodes(3).Text = TBProduto!UF_desembaraco
                            objDetDIItem.appendChild objDom.createElement("cExportador") '4
                            objDetDIItem.childNodes(4).Text = TBProduto!Codigo_exportador
                            objDetDIItem.appendChild objDom.createElement("dDesemb") '5
                            objDetDIItem.childNodes(5).Text = Format(TBProduto!Data_desembaraco, "yyyy-mm-dd")
                            objDetDIItem.appendChild objDom.createElement("tpViaTransp") '6
                            objDetDIItem.childNodes(6).Text = TBProduto!Via_transp
                            objDetDIItem.appendChild objDom.createElement("vAFRMM") '7
                            If TBProduto!Via_transp = 1 And IsNull(TBProduto!Valor_AFRMM) = False Then objDetDIItem.childNodes(7).Text = Replace(TBProduto!Valor_AFRMM, ",", ".")
                            objDetDIItem.appendChild objDom.createElement("tpIntermedio") '8
                            objDetDIItem.childNodes(8).Text = TBProduto!Forma_imp
                            objDetDIItem.appendChild objDom.createElement("CNPJ_adq") '9
                            objDetDIItem.appendChild objDom.createElement("UFTerceiro") '10
                            If TBProduto!Forma_imp <> 1 Then
                                objDetDIItem.childNodes(9).Text = DS_RetornarNumeros(TBproducao!CNPJ)
                                objDetDIItem.childNodes(10).Text = TBproducao!UF
                            End If
                        'Fecha objDetDIItem=================================================================================================
                        
                        'nó detAdicoes dentro de DetDIItem
                        Set objDetAdicoes = objDom.createElement("detAdicoes")
                        objDetDIItem.appendChild objDetAdicoes
                        'Abre objDetAdicoes==================================================================================================
                            'nó detAdicoes dentro de DetDI
                            Set objDetAdicoesItem = objDom.createElement("detAdicoesItem")
                            objDetAdicoes.appendChild objDetAdicoesItem
                            'Abre objDetAdicoesItem==================================================================================================
                                objDetAdicoesItem.appendChild objDom.createElement("nAdicao") '0
                                objDetAdicoesItem.childNodes(0).Text = TBProduto!Numero_adicao
                                objDetAdicoesItem.appendChild objDom.createElement("nSeqAdic") '1
                                objDetAdicoesItem.childNodes(1).Text = TBProduto!Numero_sequencial
                                objDetAdicoesItem.appendChild objDom.createElement("cFabricante") '2
                                objDetAdicoesItem.childNodes(2).Text = TBProduto!Codigo_fabricante
                            'Fecha objDetAdicoesItem=================================================================================================
                        'Fecha objDetAdicoes=================================================================================================
                    'Fecha objDetDI=================================================================================================
                End If
                
                If IsNull(TBProduto!Codigo_ANP) = False And TBProduto!Codigo_ANP <> "" Then
                    'nó comb dentro de prod
                    Set objComb = objDom.createElement("comb")
                    objProd.appendChild objComb
                    'Abre comb==================================================================================================
                        objComb.appendChild objDom.createElement("cProdANP") '0
                        objComb.childNodes(0).Text = TBProduto!Codigo_ANP
                        objComb.appendChild objDom.createElement("UFcons") '1
                        objComb.childNodes(1).Text = TBProduto!UF_consumo
                        objComb.appendChild objDom.createElement("descANP") '2
                        objComb.childNodes(2).Text = TBProduto!Descricao_ANP
                    'Fecha comb=================================================================================================
                End If
            'Fecha prod=================================================================================================
        
            'nó imposto dentro de objDetItem
            Set objImposto = objDom.createElement("imposto")
            objDetItem.appendChild objImposto
            'Abre objImposto==================================================================================================
                If IsNull(TBProduto!Valor_aprox_tributos) = False And TBProduto!Valor_aprox_tributos <> "" And TBProduto!Valor_aprox_tributos <> "0" Then
                    objImposto.appendChild objDom.createElement("vTotTrib") '0
                    objImposto.getElementsByTagName("vTotTrib").Item(0).Text = Replace(Format(TBProduto!Valor_aprox_tributos, "0.#0"), ",", ".")
                End If
                
                If IsNull(TBProduto!txt_CST) = False And TBProduto!txt_CST <> "" Then
                    If Len(TBProduto!txt_CST) = 4 Then FimCST = Right(TBProduto!txt_CST, 3) Else FimCST = Right(TBProduto!txt_CST, 2)
                    Set TBCST = CreateObject("adodb.recordset")
                    TBCST.Open "Select * from tbl_Detalhes_Nota_CST_ICMS where id_item = " & TBProduto!Int_codigo, Conexao, adOpenKeyset, adLockReadOnly
                    If TBCST.EOF = False Then
                        'nó objICMS dentro de objImposto
                        Set objICMS = objDom.createElement("ICMS")
                        objImposto.appendChild objICMS
                        'Abre ICMS==================================================================================================
                            ContadorReg = 0 'Usado para saber
                            objICMS.appendChild objDom.createElement("orig") '0
                            objICMS.getElementsByTagName("orig").Item(0).Text = TBCST!Origem_mercadoria 'orig
                            objICMS.appendChild objDom.createElement("CST")
                            objICMS.getElementsByTagName("CST").Item(0).Text = FimCST 'CST
                            
                            If FimCST = "00" Or FimCST = "10" Or FimCST = "20" Or FimCST = "51" Or FimCST = "70" Or FimCST = "90" Or FimCST = "900" Then
                                If IsNull(TBCST!Modalidade_determinacao) = False Then
                                    objICMS.appendChild objDom.createElement("modBC")
                                    objICMS.getElementsByTagName("modBC").Item(0).Text = TBCST!Modalidade_determinacao 'modBC
                                End If
                                objICMS.appendChild objDom.createElement("vBC")
                                objICMS.getElementsByTagName("vBC").Item(0).Text = Replace(TBCST!Valor_BC, ",", ".") 'vBC
                                objICMS.appendChild objDom.createElement("pICMS")
                                objICMS.getElementsByTagName("pICMS").Item(0).Text = Replace(TBProduto!int_ICMS, ",", ".") 'pICMS
                                objICMS.appendChild objDom.createElement("vICMS_icms")
                                objICMS.getElementsByTagName("vICMS_icms").Item(0).Text = Replace(TBCST!Valor_ICMS, ",", ".") 'vICMS_icms
                                
                                If FimCST = "51" Then
                                    objICMS.appendChild objDom.createElement("vICMSOp")
                                    objICMS.getElementsByTagName("vICMSOp").Item(0).Text = Replace(TBCST!Valor_ICMS, ",", ".") 'vICMSOp
                                    objICMS.appendChild objDom.createElement("pDif")
                                    objICMS.getElementsByTagName("pDif").Item(0).Text = Replace(TBCST!Percentual_ICMS_DIF, ",", ".") 'pDif
                                    objICMS.appendChild objDom.createElement("vICMSDif")
                                    objICMS.getElementsByTagName("vICMSDif").Item(0).Text = Replace(TBCST!Valor_ICMS_DIF, ",", ".") 'vICMSDif
                                End If
                            End If
                            If FimCST = "10" Or FimCST = "60" Or FimCST = "70" Or FimCST = "90" Or FimCST = "201" Or FimCST = "202" Or FimCST = "203" Or FimCST = "900" Then
                                If FimCST <> "60" Then
                                    objICMS.appendChild objDom.createElement("modBCST")
                                    objICMS.getElementsByTagName("modBCST").Item(0).Text = IIf(IsNull(TBCST!Modalidade_determinacao_ST), 4, TBCST!Modalidade_determinacao_ST) 'modBCST
                                    objICMS.appendChild objDom.createElement("pRedBCST")
                                    objICMS.getElementsByTagName("pRedBCST").Item(0).Text = Replace(TBCST!Percentual_reducao_BC_ST, ",", ".") 'pRedBCST
                                    objICMS.appendChild objDom.createElement("pICMSST")
                                    objICMS.getElementsByTagName("pICMSST").Item(0).Text = Replace(TBCST!Aliquota_imposto_ST, ",", ".") 'pICMSST
                                End If
                                objICMS.appendChild objDom.createElement("vBCST")
                                objICMS.getElementsByTagName("vBCST").Item(0).Text = Replace(TBCST!Valor_BC_ST, ",", ".") 'vBCST
                                objICMS.appendChild objDom.createElement("vICMSST_icms")
                                objICMS.getElementsByTagName("vICMSST_icms").Item(0).Text = Replace(TBCST!Valor_ICMS_ST, ",", ".") 'vICMSST_icms
                            End If
                            If FimCST = "20" Or FimCST = "51" Or FimCST = "70" Or FimCST = "90" Or FimCST = "900" Then
                                objICMS.appendChild objDom.createElement("pRedBC")
                                objICMS.getElementsByTagName("pRedBC").Item(0).Text = Replace(TBCST!Percentual_reducao_BC, ",", ".") 'pRedBC
                            End If
                            If FimCST = "101" Or FimCST = "201" Or FimCST = "900" Then
                                objICMS.appendChild objDom.createElement("pCredSN")
                                objICMS.getElementsByTagName("pCredSN").Item(0).Text = Replace(TBCST!ICMS_SN, ",", ".") 'pCredSN
                                objICMS.appendChild objDom.createElement("vCredICMSSN")
                                objICMS.getElementsByTagName("vCredICMSSN").Item(0).Text = Replace(TBCST!Valor_ICMS_SN, ",", ".") 'vCredICMSSN
                            End If
                        'Fecha ICMS=================================================================================================
                        
                        'nó objICMSUFDest dentro de objImposto
                        Set objICMSUFDest = objDom.createElement("ICMSUFDest")
                        objImposto.appendChild objICMSUFDest
                        'Abre ICMSUFDest==================================================================================================
                            objICMSUFDest.appendChild objDom.createElement("vBCUFDest")
                            objICMSUFDest.getElementsByTagName("vBCUFDest").Item(0).Text = Replace(IIf(IsNull(TBCST!Valor_BC_ICMS_UF_dest), 0, TBCST!Valor_BC_ICMS_UF_dest), ",", ".")
                            objICMSUFDest.appendChild objDom.createElement("pFCPUFDest") '1
                            objICMSUFDest.getElementsByTagName("pFCPUFDest").Item(0).Text = Replace(IIf(IsNull(TBCST!Percentual_FCP), 0, TBCST!Percentual_FCP), ",", ".")
                            
                            objICMSUFDest.appendChild objDom.createElement("pICMSUFDest") '2
                            objICMSUFDest.appendChild objDom.createElement("pICMSInter") '3
                            objICMSUFDest.appendChild objDom.createElement("pICMSInterPart") '4
                            objICMSUFDest.getElementsByTagName("pICMSInterPart").Item(0).Text = Replace(IIf(IsNull(TBCST!Percentual_provisorio), 0, TBCST!Percentual_provisorio), ",", ".")
                            objICMSUFDest.appendChild objDom.createElement("vFCPUFDest") '5
                            objICMSUFDest.getElementsByTagName("vFCPUFDest").Item(0).Text = Replace(IIf(IsNull(TBCST!Valor_ICMS_FCP), 0, TBCST!Valor_ICMS_FCP), ",", ".")
                            objICMSUFDest.appendChild objDom.createElement("vICMSUFDest") '6
                            objICMSUFDest.getElementsByTagName("vICMSUFDest").Item(0).Text = Replace(IIf(IsNull(TBCST!Valor_ICMS_INT_UF_dest), 0, TBCST!Valor_ICMS_INT_UF_dest), ",", ".")
                            objICMSUFDest.appendChild objDom.createElement("vICMSUFRemet") '7
                            objICMSUFDest.getElementsByTagName("vICMSUFRemet").Item(0).Text = Replace(IIf(IsNull(TBCST!Valor_ICMS_INT_UF_rem), 0, TBCST!Valor_ICMS_INT_UF_rem), ",", ".")
                            
                            If IsNull(TBCST!Valor_ICMS_INT_UF_dest) = False And TBCST!Valor_ICMS_INT_UF_dest > 0 Then
                                If IsNull(TBproducao!txt_UF) = True Or TBproducao!txt_UF = "" Or TBproducao!txt_UF = "EX" Then
                                    objICMSUFDest.getElementsByTagName("pICMSInter").Item(0).Text = 4
                                Else
                                    ProcBuscaTributos IIf(IsNull(TBProduto!ID_CF), 0, TBProduto!ID_CF)
                                    ProcVerificaRegiao TBproducao!txt_UF, TBproducao!Id_Int_Cliente, TBproducao!txt_Razao_Nome
                                    objICMSUFDest.getElementsByTagName("pICMSInter").Item(0).Text = vRegiao(0, 1)
                                End If
                                Set TBFIltro = CreateObject("adodb.recordset")
                                TBFIltro.Open "Select ICMS_interno from regioes where UF = '" & TBproducao!txt_UF & "'", Conexao, adOpenKeyset, adLockOptimistic
                                If TBFIltro.EOF = False Then
                                    objICMSUFDest.getElementsByTagName("pICMSUFDest").Item(0).Text = IIf(IsNull(TBFIltro!ICMS_interno), 0, TBFIltro!ICMS_interno)
                                Else
                                    objICMSUFDest.getElementsByTagName("pICMSUFDest").Item(0).Text = "0.00"
                                End If
                                TBFIltro.Close
                            Else
                                objICMSUFDest.getElementsByTagName("pICMSInter").Item(0).Text = "0.00"
                                objICMSUFDest.getElementsByTagName("pICMSUFDest").Item(0).Text = "0.00"
                            End If
                        'Fecha ICMSUFDest=================================================================================================
                    End If
                    TBCST.Close
                End If
                
                If IsNull(TBProduto!CST_IPI) = False And TBProduto!CST_IPI <> "" Then
                    FimCST = Right(TBProduto!CST_IPI, 2)
                    Set TBCST = CreateObject("adodb.recordset")
                    TBCST.Open "Select * from tbl_Detalhes_Nota_CST_IPI where id_item = " & TBProduto!Int_codigo, Conexao, adOpenKeyset, adLockReadOnly
                    If TBCST.EOF = False Then
                        'nó objIPI dentro de objImposto
                        Set objIPI = objDom.createElement("IPI")
                        objImposto.appendChild objIPI
                        'Abre IPI==================================================================================================
                            objIPI.appendChild objDom.createElement("cEnq")
                            If IsNull(TBProduto!Codigo_enquadramento_IPI) = False Then
                                objIPI.getElementsByTagName("cEnq").Item(0).Text = TBProduto!Codigo_enquadramento_IPI
                            Else
                                objIPI.getElementsByTagName("cEnq").Item(0).Text = "999"
                            End If
                            
                            'nó CSTIPI dentro de IPI
                            Set objCSTIPI = objDom.createElement("CSTIPI")
                            objIPI.appendChild objCSTIPI
                            'Abre CSTIPI==================================================================================================
                                objCSTIPI.appendChild objDom.createElement("CST_IPI") '0
                                objCSTIPI.getElementsByTagName("CST_IPI").Item(0).Text = FimCST
                                'objCSTIPI.appendChild objDom.createElement("qUnid_IPI") '2
                                'objCSTIPI.appendChild objDom.createElement("vUnid_IPI") '3

                                If FimCST = "00" Or FimCST = "49" Or FimCST = "50" Or FimCST = "99" And TBproducao!Simples = False Then
                                    objCSTIPI.appendChild objDom.createElement("vBC_IPI")
                                    objCSTIPI.getElementsByTagName("vBC_IPI").Item(0).Text = Replace(TBCST!Valor_BC, ",", ".")
                                    objCSTIPI.appendChild objDom.createElement("pIPI")
                                    objCSTIPI.getElementsByTagName("pIPI").Item(0).Text = Replace(TBProduto!int_IPI, ",", ".")
                                    
                                    objCSTIPI.appendChild objDom.createElement("vIPI")
                                    If IsNull(TBProduto!dbl_ValorTotal) = True Or TBProduto!dbl_ValorTotal = 0 Then
                                        objCSTIPI.getElementsByTagName("vIPI").Item(0).Text = Replace(Format((TBCST!Valor_BC * TBProduto!int_IPI) / 100, "0.00"), ",", ".")
                                    Else
                                        objCSTIPI.getElementsByTagName("vIPI").Item(0).Text = Replace(Format(TBProduto!dbl_valoripi, "0.00"), ",", ".")
                                    End If
                                End If
                            'Fecha CSTIPI=================================================================================================
                        'Fecha IPI=================================================================================================
                    End If
                    TBCST.Close
                End If
                
                'II
                If IsNull(TBProduto!Local_desembaraco) = False And TBProduto!Local_desembaraco <> "" Then
                    'nó II dentro de objImposto
                    Set objII = objDom.createElement("II")
                    objImposto.appendChild objII
                    'Abre objII==================================================================================================
                        objII.appendChild objDom.createElement("vBC_imp")
                        objII.childNodes(0).Text = Replace(TBProduto!Valor_BC_importacao, ",", ".")
                        objII.appendChild objDom.createElement("vDespAdu")
                        objII.childNodes(1).Text = Replace(TBProduto!Valor_despesas, ",", ".")
                        objII.appendChild objDom.createElement("vII")
                        objII.childNodes(2).Text = Replace(TBProduto!Valor_imposto_importacao, ",", ".")
                        objII.appendChild objDom.createElement("vIOF")
                        objII.childNodes(3).Text = Replace(TBProduto!Valor_imposto_OperacoesFinanceiras, ",", ".")
                    'Fecha objII==================================================================================================
                End If
                
                'PIS
                If IsNull(TBProduto!CST_PIS) = False And TBProduto!CST_PIS <> "" Then
                    FimCST = Right(TBProduto!CST_PIS, 2)
                    Set TBCST = CreateObject("adodb.recordset")
                    TBCST.Open "Select * from tbl_Detalhes_Nota_CST_PIS where id_item = " & TBProduto!Int_codigo, Conexao, adOpenKeyset, adLockReadOnly
                    If TBCST.EOF = False Then
                        'nó objPis dentro de objImposto
                        Set objPis = objDom.createElement("PIS")
                        objImposto.appendChild objPis
                        'Abre Pis==================================================================================================
                            objPis.appendChild objDom.createElement("CST_pis")
                            objPis.getElementsByTagName("CST_pis").Item(0).Text = FimCST
                            If FimCST = "01" Or FimCST = "02" Or FimCST = "03" Or FimCST = "49" Or FimCST = "98" Or FimCST = "99" Then
                                objPis.appendChild objDom.createElement("vBC_pis")
                                objPis.getElementsByTagName("vBC_pis").Item(0).Text = Replace(TBCST!Valor_BC, ",", ".")
                            End If
                            objPis.appendChild objDom.createElement("pPIS")
                            objPis.getElementsByTagName("pPIS").Item(0).Text = Replace(TBProduto!PIS_Prod, ",", ".")
                            objPis.appendChild objDom.createElement("vPIS")
                            objPis.getElementsByTagName("vPIS").Item(0).Text = Replace(TBProduto!Total_PIS_prod, ",", ".")
                            objPis.appendChild objDom.createElement("qBCprod_pis")
                            objPis.getElementsByTagName("qBCprod_pis").Item(0).Text = Replace(TBProduto!int_Qtd, ",", ".")
                            objPis.appendChild objDom.createElement("vAliqProd_pis")
                            objPis.getElementsByTagName("vAliqProd_pis").Item(0).Text = Replace(TBProduto!PIS_Prod, ",", ".")
                        'Fecha Pis=================================================================================================
                    End If
                    TBCST.Close
                End If
                
                'Cofins
                If IsNull(TBProduto!CST_Cofins) = False And TBProduto!CST_Cofins <> "" Then
                    FimCST = Right(TBProduto!CST_Cofins, 2)
                    Set TBCST = CreateObject("adodb.recordset")
                    TBCST.Open "Select * from tbl_Detalhes_Nota_CST_Cofins where id_item = " & TBProduto!Int_codigo, Conexao, adOpenKeyset, adLockReadOnly
                    If TBCST.EOF = False Then
                        'nó cofins dentro de objImposto (O01)
                        Set objCofins = objDom.createElement("COFINS")
                        objImposto.appendChild objCofins
                        'Abre Pis==================================================================================================
                            objCofins.appendChild objDom.createElement("CST_cofins")
                            objCofins.getElementsByTagName("CST_cofins").Item(0).Text = FimCST
                            If FimCST = "01" Or FimCST = "03" Or FimCST = "49" Or FimCST = "98" Or FimCST = "99" Then
                                objCofins.appendChild objDom.createElement("vBC_cofins")
                                objCofins.getElementsByTagName("vBC_cofins").Item(0).Text = Replace(TBCST!Valor_BC, ",", ".")
                            End If
                            objCofins.appendChild objDom.createElement("pCOFINS")
                            objCofins.getElementsByTagName("pCOFINS").Item(0).Text = Replace(TBProduto!Cofins_Prod, ",", ".")
                            objCofins.appendChild objDom.createElement("vCOFINS")
                            objCofins.getElementsByTagName("vCOFINS").Item(0).Text = Replace(TBProduto!Total_Cofins_prod, ",", ".")
                            objCofins.appendChild objDom.createElement("qBCProd_cofins")
                            objCofins.getElementsByTagName("qBCProd_cofins").Item(0).Text = Replace(TBProduto!int_Qtd, ",", ".")
                            objCofins.appendChild objDom.createElement("vAliqProd_cofins")
                            objCofins.getElementsByTagName("vAliqProd_cofins").Item(0).Text = Replace(TBProduto!Cofins_Prod, ",", ".")
                        'Fecha Pis=================================================================================================
                    End If
                    TBCST.Close
                End If
            'Abre objImposto==================================================================================================
            
            'Verificar depois
            If Devolucao = True And TBproducao!Simples = True And TBProduto!dbl_valoripi > 0 Then
                'nó impostoDevol dentro de detItem (G02)
                Set objImpostoDevol = objDom.createElement("impostoDevol")
                objDetItem.appendChild objImpostoDevol
                'Abre objImpostoDevol==================================================================================================
                    objImpostoDevol.appendChild objDom.createElement("pDevol")
                    objImpostoDevol.childNodes(0).Text = "100"
                    
                    'nó IPIDevol dentro de impostoDevol
                    Set objIPIDevol = objDom.createElement("IPIDevol")
                    objImpostoDevol.appendChild objIPIDevol
                    'Abre IPIDevol==================================================================================================
                        objIPIDevol.appendChild objDom.createElement("vIPIDevol")
                        objIPIDevol.childNodes(0).Text = Replace(TBProduto!dbl_valoripi, ",", ".")
                    'Fecha IPIDevol=================================================================================================
                'Fecha objImpostoDevol=================================================================================================
            End If
        'Fecha detItem=================================================================================================
        TBProduto.MoveNext
    Loop
'Fecha det================================================================================================

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Public Sub procIdentificacaoXML()
On Error GoTo tratar_erro

'no ide dentro de Envio (A01)
Set objIde = objDom.createElement("ide")
objEnviar.appendChild objIde
'Abre ide=================================================================================================
    'filhos ide
    objIde.appendChild objDom.createElement("cNF")
    objIde.getElementsByTagName("cNF").Item(0).Text = FunTamanhoTextoZeroEsq(txtNota, 8)
    
    FamiliaAntiga = FunTiraAcentosTexto(TBproducao!Cidade) 'Empresa
    
    objIde.appendChild objDom.createElement("cUF")
    objIde.getElementsByTagName("cUF").Item(0).Text = FunVerificaCodUF(FamiliaAntiga, TBproducao!UF)
    
    Set TBCFOP = CreateObject("adodb.recordset")
    TBCFOP.Open "Select CFOP.Txt_descricao from tbl_Detalhes_Nota NFP INNER JOIN tbl_NaturezaOperacao CFOP ON CFOP.IDCountCfop = NFP.ID_CFOP where NFP.ID_nota = " & txtID_nota & " order by NFP.Int_codigo", Conexao, adOpenKeyset, adLockReadOnly
    If TBCFOP.EOF = False Then
        objIde.appendChild objDom.createElement("natOp")
        objIde.getElementsByTagName("natOp").Item(0).Text = FunTiraAcentosTexto(TBCFOP!Txt_descricao)
    End If
    TBCFOP.Close
    
    objIde.appendChild objDom.createElement("mod")
    objIde.getElementsByTagName("mod").Item(0).Text = 55
    objIde.appendChild objDom.createElement("serie")
    objIde.getElementsByTagName("serie").Item(0).Text = txtSerie
    objIde.appendChild objDom.createElement("nNF")
    objIde.getElementsByTagName("nNF").Item(0).Text = Format(txtNota, "0")
    objIde.appendChild objDom.createElement("dhEmi")
    objIde.getElementsByTagName("dhEmi").Item(0).Text = Format(TBproducao!dt_DataEmissao, "yyyy-mm-dd") & "T" & Left(TBproducao!Hora_emissao, 8)
    objIde.appendChild objDom.createElement("fusoHorario")
    objIde.getElementsByTagName("fusoHorario").Item(0).Text = FunVerifFusoHorario(True)
    objIde.appendChild objDom.createElement("dhSaiEnt")
    objIde.getElementsByTagName("dhSaiEnt").Item(0).Text = Format(TBproducao!dt_DataEmissao, "yyyy-mm-dd") & "T" & Left(TBproducao!Hora_emissao, 8)
    objIde.appendChild objDom.createElement("tpNf")
    If TBproducao!int_TipoNota = 1 Then
        objIde.getElementsByTagName("tpNf").Item(0).Text = 1
    Else
        objIde.getElementsByTagName("tpNf").Item(0).Text = 0
    End If
            
    objIde.appendChild objDom.createElement("idDest")
    If TBproducao!UF = TBproducao!txt_UF Then
        objIde.getElementsByTagName("idDest").Item(0).Text = 1
    ElseIf TBproducao!txt_UF = "EX" Then
        objIde.getElementsByTagName("idDest").Item(0).Text = 3
    Else
        objIde.getElementsByTagName("idDest").Item(0).Text = 2
    End If
    
    objIde.appendChild objDom.createElement("indFinal")
    objIde.getElementsByTagName("indFinal").Item(0).Text = TBproducao!Consumidor_final
    objIde.appendChild objDom.createElement("indPres")
    objIde.getElementsByTagName("indPres").Item(0).Text = TBproducao!Presenca_comprador
    objIde.appendChild objDom.createElement("cMunFg")
    objIde.getElementsByTagName("cMunFg").Item(0).Text = FunVerificaCodMunicipio(FamiliaAntiga, TBproducao!UF)
    objIde.appendChild objDom.createElement("tpImp")
    objIde.getElementsByTagName("tpImp").Item(0).Text = 1 'DANFE 1 = Retrato - 2 = Paisagem no manual
    objIde.appendChild objDom.createElement("tpEmis")
    objIde.getElementsByTagName("tpEmis").Item(0).Text = TBproducao!Forma_emissao
    objIde.appendChild objDom.createElement("tpAmb")
    objIde.getElementsByTagName("tpAmb").Item(0).Text = 1 '1-produção 2-Homologação
    'objIde.appendChild objDom.createElement("xJust")
    'objIde.childNodes(17).Text = ""
    'objIde.appendChild objDom.createElement("dhCont")
    'objIde.childNodes(18).Text = ""
    objIde.appendChild objDom.createElement("finNFe")
    objIde.getElementsByTagName("finNFe").Item(0).Text = TBproducao!Finalidade_emissao
    If Email <> "" Then
        objIde.appendChild objDom.createElement("EmailArquivos")
        objIde.getElementsByTagName("EmailArquivos").Item(0).Text = Trim(Email)
    End If
    'objIde.appendChild objDom.createElement("NumeroPedido")
    'objIde.childNodes(19).Text = ""
        
    'nó objNFRef dentro de Ide
    Set objNFRef = objDom.createElement("NFRef")
    objIde.appendChild objNFRef
    'Abre NFRef=================================================================================================
        'filhos NFRef
        'Contador2 = 0
        Set TBCiclo = CreateObject("adodb.recordset")
        TBCiclo.Open "Select ID_nota_relacionada AS ID from Faturamento_Relacionamento where ID_nota = " & txtID_nota & " group by ID_nota_relacionada", Conexao, adOpenKeyset, adLockReadOnly
        If TBCiclo.EOF = True Then
            Set TBCiclo = CreateObject("adodb.recordset")
            TBCiclo.Open "Select ID_nota AS ID from Faturamento_Relacionamento where ID_nota_relacionada = " & txtID_nota & " group by ID_nota", Conexao, adOpenKeyset, adLockReadOnly
        End If
        Do While TBCiclo.EOF = False
            'nó NFRefItem dentro de NFRef
            Set objNFRefItem = objDom.createElement("NFRefItem")
            objNFRef.appendChild objNFRefItem
            'Abre NFRefItem=================================================================================================
                'filhos NFRefItem
                objNFRefItem.appendChild objDom.createElement("refNFe") '0
                'Contador2 = Contador2 + 1
                'TBGravar_NFe!cbdrefSeq = Contador2
                Set TBCarteira = CreateObject("adodb.recordset")
                TBCarteira.Open "Select ID, int_NotaFiscal, txt_Municipio, txt_UF, dt_DataEmissao, txt_CNPJ_CPF, Modelo, Serie from tbl_Dados_Nota_Fiscal where ID = " & TBCiclo!ID, Conexao, adOpenKeyset, adLockOptimistic
                If TBCarteira.EOF = False Then
                    Set TBTempo = CreateObject("adodb.recordset")
                    TBTempo.Open "Select Chave_acesso from tbl_Dados_Nota_Fiscal_NFe where ID_nota = " & TBCarteira!ID & " and Chave_acesso IS NOT NULL and Chave_acesso <> N''", Conexao, adOpenKeyset, adLockOptimistic
                    If TBTempo.EOF = False Then
                        objNFRefItem.childNodes(0).Text = TBTempo!Chave_acesso
                    Else
                        objNFRefItem.appendChild objDom.createElement("cUF_refNFE") '1
                        objNFRefItem.appendChild objDom.createElement("AAMM") '2
                        objNFRefItem.appendChild objDom.createElement("CNPJ") '3
                        objNFRefItem.appendChild objDom.createElement("CPF") '4
                        objNFRefItem.appendChild objDom.createElement("mod_refNFE") '5
                        objNFRefItem.appendChild objDom.createElement("serie_refNFE") '6
                        objNFRefItem.appendChild objDom.createElement("IE_refNFP") '7
                        'objNFRefItem.appendChild objDom.createElement("RefCte") '8
                        'objNFRefItem.appendChild objDom.createElement("mod_refECF") '9
                        'objNFRefItem.appendChild objDom.createElement("nECF_refECF") '10
                        'objNFRefItem.appendChild objDom.createElement("nCOO_refECF") '11
                    
                        IDpedido = TBCarteira!int_NotaFiscal
                        FamiliaAntiga = FunTiraAcentosTexto(TBCarteira!txt_Municipio)
                        objNFRefItem.childNodes(1).Text = FunVerificaCodUF(FamiliaAntiga, TBCarteira!txt_UF)
                        objNFRefItem.childNodes(2).Text = Format(TBCarteira!dt_DataEmissao, "YYMM")
                        objNFRefItem.childNodes(3).Text = DS_RetornarNumeros(TBCarteira!txt_CNPJ_CPF)
                        objNFRefItem.childNodes(5).Text = IIf(Left(TBCarteira!Modelo, 2) = "1B", 1, Left(TBCarteira!Modelo, 2))
                        objNFRefItem.childNodes(6).Text = TBCarteira!Serie
                        objNFRefItem.childNodes(7).Text = IDpedido
                    End If
                    TBTempo.Close
                End If
                TBCarteira.Close
            'Fecha NFRef====================================================================================================
            TBCiclo.MoveNext
        Loop
        TBCiclo.Close
    'Fecha NFRef===================================================================
'Fecha ide===================================================================

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Public Sub procEmitenteXML()
On Error GoTo tratar_erro

'no emit dentro de Envio (A01)
Set objEmit = objDom.createElement("emit")
objEnviar.appendChild objEmit
'Abre emit=================================================================================================
    objEmit.appendChild objDom.createElement("CNPJ_emit")
    objEmit.getElementsByTagName("CNPJ_emit").Item(0).Text = DS_RetornarNumeros(TBproducao!CNPJ)
    'objEmit.appendChild objDom.createElement("CPF_emit")
    'objemit.childNodes(1).Text = ""
    objEmit.appendChild objDom.createElement("xNome")
    objEmit.getElementsByTagName("xNome").Item(0).Text = Trim(FunTiraAcentosTexto(Left(TBproducao!Razao, 60)))
    objEmit.appendChild objDom.createElement("xFant")
    objEmit.getElementsByTagName("xFant").Item(0).Text = Trim(FunTiraAcentosTexto(TBproducao!Empresa))
    If IsNull(TBproducao!IM) = False And TBproducao!IM <> "" Then
        objEmit.appendChild objDom.createElement("IM")
        objEmit.getElementsByTagName("IM").Item(0).Text = DS_RetornarNumeros(TBproducao!IM)
        If IsNull(TBproducao!CNAE) = False And TBproducao!CNAE <> "" Then
            objEmit.appendChild objDom.createElement("CNAE")
            objEmit.getElementsByTagName("CNAE").Item(0).Text = DS_RetornarNumeros(TBproducao!CNAE)
        End If
    End If
    If IsNull(TBproducao!ie) = False And TBproducao!ie <> "" Then
        objEmit.appendChild objDom.createElement("IE")
        objEmit.getElementsByTagName("IE").Item(0).Text = IIf(TBproducao!ie = "ISENTO", "ISENTO", Left(DS_RetornarNumeros(TBproducao!ie), 14))
    End If
    
    objEmit.appendChild objDom.createElement("CRT")
    If TBproducao!Simples = True Then
        objEmit.getElementsByTagName("CRT").Item(0).Text = 1
    ElseIf TBproducao!Simples1 = True Then
        objEmit.getElementsByTagName("CRT").Item(0).Text = 2
    Else
        objEmit.getElementsByTagName("CRT").Item(0).Text = 3
    End If
    
    'nó enderEmit dentro de emit
    Set objEnderEmit = objDom.createElement("enderEmit")
    objEmit.appendChild objEnderEmit
    'Abre enderEmit=================================================================================================
        objEnderEmit.appendChild objDom.createElement("xLgr")
        FamiliaAntiga = ""
        If IsNull(TBproducao!Tipo_endereco) = False And TBproducao!Tipo_endereco <> "" Then FamiliaAntiga = TBproducao!Tipo_endereco & ": "
        If FamiliaAntiga <> "" Then FamiliaAntiga = FamiliaAntiga & TBproducao!Endereco Else FamiliaAntiga = TBproducao!Endereco
        objEnderEmit.getElementsByTagName("xLgr").Item(0).Text = Trim(FunTiraAcentosTexto(FamiliaAntiga))
        
        objEnderEmit.appendChild objDom.createElement("nro")
        objEnderEmit.getElementsByTagName("nro").Item(0).Text = TBproducao!numeroEmpresa
        
        If IsNull(TBproducao!Complemento) = False Then
            objEnderEmit.appendChild objDom.createElement("xCpl") '2
            objEnderEmit.getElementsByTagName("xCpl").Item(0).Text = Trim(TBproducao!Complemento)
        End If
        
        objEnderEmit.appendChild objDom.createElement("xBairro") '3
        FamiliaAntiga = ""
        If IsNull(TBproducao!Tipo_bairro) = False And TBproducao!Tipo_bairro <> "" Then FamiliaAntiga = TBproducao!Tipo_bairro & ": "
        If FamiliaAntiga <> "" Then FamiliaAntiga = FamiliaAntiga & TBproducao!Bairro Else Bairro = TBproducao!Bairro
        objEnderEmit.getElementsByTagName("xBairro").Item(0).Text = Trim(FunTiraAcentosTexto(FamiliaAntiga))
        
        objEnderEmit.appendChild objDom.createElement("cMun") '4
        FamiliaAntiga = FunTiraAcentosTexto(TBproducao!Cidade)
        objEnderEmit.getElementsByTagName("cMun").Item(0).Text = FunVerificaCodMunicipio(FamiliaAntiga, TBproducao!UF)
        
        objEnderEmit.appendChild objDom.createElement("xMun") '5
        objEnderEmit.getElementsByTagName("xMun").Item(0).Text = FunTiraAcentosTexto(TBproducao!Cidade)
        objEnderEmit.appendChild objDom.createElement("UF") '6
        objEnderEmit.getElementsByTagName("UF").Item(0).Text = TBproducao!UF
        objEnderEmit.appendChild objDom.createElement("CEP") '7
        objEnderEmit.getElementsByTagName("CEP").Item(0).Text = DS_RetornarNumeros(TBproducao!CEP)
        objEnderEmit.appendChild objDom.createElement("cPais") '8
        objEnderEmit.getElementsByTagName("cPais").Item(0).Text = "1058"
        objEnderEmit.appendChild objDom.createElement("xPais") '9
        objEnderEmit.getElementsByTagName("xPais").Item(0).Text = "BRASIL"
        If IsNull(TBproducao!Telefone) = False And TBproducao!Telefone <> "" Then
            objEnderEmit.appendChild objDom.createElement("fone") '10
            objEnderEmit.getElementsByTagName("fone").Item(0).Text = DS_RetornarNumeros(TBproducao!Telefone)
        End If
        If IsNull(TBproducao!Email) = False And TBproducao!Email <> "" Then
            objEnderEmit.appendChild objDom.createElement("Email") '11
            objEnderEmit.getElementsByTagName("Email").Item(0).Text = TBproducao!Email
        End If
    'Fecha enderEmit================================================================================================
'Fecha emit================================================================================================

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Public Sub procDestinatarioXML()
On Error GoTo tratar_erro

'no dest dentro de Envio (A01)
Set objDest = objDom.createElement("dest")
objEnviar.appendChild objDest
'Abre objDest=================================================================================================
    
    If IsNull(TBproducao!txt_UF) = True Or TBproducao!txt_UF = "" Or TBproducao!txt_UF = "EX" Then
        'objDest.appendChild objDom.createElement("idEstrangeiro") '2
        'objdest.childNodes(2).Text = ""
    Else
        
        If TBproducao!txt_tipocliente = "E" Or Left(TBproducao!txt_tipocliente, 1) = "J" Then
            objDest.appendChild objDom.createElement("CNPJ_dest") '0
            objDest.getElementsByTagName("CNPJ_dest").Item(0).Text = DS_RetornarNumeros(TBproducao!txt_CNPJ_CPF) 'CNPJ
        Else
            objDest.appendChild objDom.createElement("CPF_dest") '1
            objDest.getElementsByTagName("CPF_dest").Item(0).Text = DS_RetornarNumeros(TBproducao!txt_CNPJ_CPF) 'CPF
        End If
    End If
    
    objDest.appendChild objDom.createElement("xNome_dest") '3
    objDest.getElementsByTagName("xNome_dest").Item(0).Text = FunTiraAcentosTexto(TBproducao!txt_Razao_Nome)
    'objDest.getElementsByTagName("xNome_dest").Item(0).Text = "NF-E EMITIDA EM AMBIENTE DE HOMOLOGACAO - SEM VALOR FISCAL" 'Usado para testes em homologação
    
    'Verifica se tem suframa
    
    Set TBClientes = CreateObject("adodb.recordset")
    TBClientes.Open "Select CFOP.* from tbl_Detalhes_Nota NFP INNER JOIN tbl_NaturezaOperacao CFOP ON CFOP.IDCountCfop = NFP.ID_CFOP where NFP.ID_nota = " & txtID_nota & " and CFOP.Suframa = 'True'", Conexao, adOpenKeyset, adLockReadOnly
    If TBClientes.EOF = False Then
        Set TBClientes = CreateObject("adodb.recordset")
        TBClientes.Open "Select * from Clientes where IDCliente = " & TBproducao!Id_Int_Cliente & " and Suframa is not null", Conexao, adOpenKeyset, adLockReadOnly
        If TBClientes.EOF = False Then
            If TBClientes!Suframa <> "" Then
                objDest.appendChild objDom.createElement("ISUF") '5
                objDest.getElementsByTagName("ISUF").Item(0).Text = Left(DS_RetornarNumeros(TBClientes!Suframa), 9)
            End If
        End If
    End If
    TBClientes.Close
    
    'nó enderDest dentro de dest
    Set objEnderDest = objDom.createElement("enderDest")
    objDest.appendChild objEnderDest
    'Abre enderDest=================================================================================================
        objEnderDest.appendChild objDom.createElement("nro_dest") '0
        objEnderDest.getElementsByTagName("nro_dest").Item(0).Text = TBproducao!Numero
        objEnderDest.appendChild objDom.createElement("xBairro_dest") '2
        objEnderDest.getElementsByTagName("xBairro_dest").Item(0).Text = Trim(FunTiraAcentosTexto(TBproducao!Txt_bairro))
        objEnderDest.appendChild objDom.createElement("xLgr_dest") '4
        objEnderDest.getElementsByTagName("xLgr_dest").Item(0).Text = Trim(FunTiraAcentosTexto(TBproducao!txt_Endereco))
        
        objEnderDest.appendChild objDom.createElement("xPais_dest") '5
        objEnderDest.appendChild objDom.createElement("cMun_dest") '6
        objEnderDest.appendChild objDom.createElement("xMun_dest") '7
        objEnderDest.appendChild objDom.createElement("UF_dest") '8
        
        objEnderDest.appendChild objDom.createElement("CEP_dest") '9
        objEnderDest.getElementsByTagName("CEP_dest").Item(0).Text = Left(DS_RetornarNumeros(TBproducao!Txt_CEP), 8)
        objEnderDest.appendChild objDom.createElement("cPais_dest") '10
        
        If IsNull(TBproducao!txt_Fone_Fax) = False And TBproducao!txt_Fone_Fax <> "" Then
            objEnderDest.appendChild objDom.createElement("fone_dest") '11
            objEnderDest.getElementsByTagName("fone_dest").Item(0).Text = Right(DS_RetornarNumeros(TBproducao!txt_Fone_Fax), 10)
        End If
        
        If IsNull(TBproducao!txt_UF) = True Or TBproducao!txt_UF = "" Or TBproducao!txt_UF = "EX" Then
            objEnderDest.getElementsByTagName("cMun_dest").Item(0).Text = "9999999"
            objEnderDest.getElementsByTagName("xMun_dest").Item(0).Text = "EXTERIOR"
            objEnderDest.getElementsByTagName("UF_dest").Item(0).Text = "EX"
        Else
            FamiliaAntiga = FunTiraAcentosTexto(TBproducao!txt_Municipio)
            objEnderDest.getElementsByTagName("cMun_dest").Item(0).Text = FunVerificaCodMunicipio(FamiliaAntiga, TBproducao!txt_UF)
            objEnderDest.getElementsByTagName("xMun_dest").Item(0).Text = FunTiraAcentosTexto(TBproducao!txt_Municipio)
            objEnderDest.getElementsByTagName("UF_dest").Item(0).Text = TBproducao!txt_UF
        End If
        
        Set TBFornecedor = CreateObject("adodb.recordset")
        TBFornecedor.Open "Select Email, Codigo_pais, Pais, Complemento, RG_IM, Nao_contribuinte_ICMS, Pessoa from Compras_fornecedores where IDCliente = " & TBproducao!Id_Int_Cliente & " and Nome_Razao = '" & TBproducao!txt_Razao_Nome & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBFornecedor.EOF = False Then
            If (TBFornecedor!Nao_contribuinte_ICMS) = True Then Nao_contribuinte_ICMS = "Sim" Else Nao_contribuinte_ICMS = "Não"
            If IsNull(TBFornecedor!Complemento) = False And TBFornecedor!Complemento <> "" Then
                objEnderDest.appendChild objDom.createElement("xCpl_dest") '1
                objEnderDest.getElementsByTagName("xCpl_dest").Item(0).Text = Trim(TBFornecedor!Complemento)
            End If
            If TBFornecedor!Pessoa = "JURÍDICA" And IsNull(TBFornecedor!RG_IM) = False And TBFornecedor!RG_IM <> "" Then
                objDest.appendChild objDom.createElement("IM_dest") '7
                objDest.getElementsByTagName("IM_dest").Item(0).Text = TBFornecedor!RG_IM 'Este campo IM não esta dentro de objEnderDest esta um nivel acima em objDest
            End If
            objEnderDest.getElementsByTagName("xPais_dest").Item(0).Text = TBFornecedor!Pais
            objEnderDest.getElementsByTagName("cPais_dest").Item(0).Text = TBFornecedor!Codigo_pais
            If IsNull(TBFornecedor!Email) = False Then
                objEnderDest.appendChild objDom.createElement("xEmail_dest") '3
                objEnderDest.getElementsByTagName("xEmail_dest").Item(0).Text = TBFornecedor!Email
            End If
        Else
            Set TBFornecedor = CreateObject("adodb.recordset")
            TBFornecedor.Open "Select Email, Codigo_pais, Pais, Complemento, RG_IM, Nao_contribuinte_ICMS, Tipo from Clientes where IDCliente = " & TBproducao!Id_Int_Cliente & " and NomeRazao = '" & TBproducao!txt_Razao_Nome & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBFornecedor.EOF = False Then
                If (TBFornecedor!Nao_contribuinte_ICMS) = True Then Nao_contribuinte_ICMS = "Sim" Else Nao_contribuinte_ICMS = "Não"
                If IsNull(TBFornecedor!Complemento) = False And TBFornecedor!Complemento <> "" Then
                    objEnderDest.appendChild objDom.createElement("xCpl_dest") '1
                    objEnderDest.getElementsByTagName("xCpl_dest").Item(0).Text = Trim(TBFornecedor!Complemento)
                End If
                If Left(TBFornecedor!Tipo, 1) = "J" And IsNull(TBFornecedor!RG_IM) = False And TBFornecedor!RG_IM <> "" Then
                    objDest.appendChild objDom.createElement("IM_dest") '7
                    objDest.getElementsByTagName("IM_dest").Item(0).Text = TBFornecedor!RG_IM 'Este campo IM não esta dentro de objEnderDest esta um nivel acima em objDest
                End If
                objEnderDest.getElementsByTagName("xPais_dest").Item(0).Text = TBFornecedor!Pais
                objEnderDest.getElementsByTagName("cPais_dest").Item(0).Text = TBFornecedor!Codigo_pais
                If IsNull(TBFornecedor!Email) = False Then
                    objEnderDest.appendChild objDom.createElement("xEmail_dest") '3
                    objEnderDest.getElementsByTagName("xEmail_dest").Item(0).Text = TBFornecedor!Email
                End If
            Else
                Set TBFornecedor = CreateObject("adodb.recordset")
                TBFornecedor.Open "Select email, Codigo_pais, Pais, Complemento, IM from Empresa where Codigo = " & TBproducao!Id_Int_Cliente, Conexao, adOpenKeyset, adLockOptimistic
                If TBFornecedor.EOF = False Then
                    Nao_contribuinte_ICMS = "Não"
                    If IsNull(TBFornecedor!Complemento) = False And TBFornecedor!Complemento <> "" Then
                        objEnderDest.appendChild objDom.createElement("xCpl_dest") '1
                        objEnderDest.getElementsByTagName("xCpl_dest").Item(0).Text = Trim(TBFornecedor!Complemento)
                    End If
                    If IsNull(TBFornecedor!IM) = False And TBFornecedor!IM <> "" Then
                        objDest.appendChild objDom.createElement("IM_dest") '7
                        objDest.getElementsByTagName("IM_dest").Item(0).Text = TBFornecedor!IM 'Este campo IM não esta dentro de objEnderDest esta um nivel acima em objDest
                    End If
                    objEnderDest.getElementsByTagName("xPais_dest").Item(0).Text = TBFornecedor!Pais
                    objEnderDest.getElementsByTagName("cPais_dest").Item(0).Text = TBFornecedor!Codigo_pais
                    If IsNull(TBFornecedor!Email) = False Then
                        objEnderDest.appendChild objDom.createElement("xEmail_dest") '3
                        objEnderDest.getElementsByTagName("xEmail_dest").Item(0).Text = TBFornecedor!Email
                    End If
                End If
            End If
        End If
        TBFornecedor.Close
    'Fecha enderDest================================================================================================
    
    If Nao_contribuinte_ICMS = "Sim" Or IsNull(TBproducao!txt_UF) = True Or TBproducao!txt_UF = "" Or TBproducao!txt_UF = "EX" Or IsNull(TBproducao!txt_IE_Cliente) = True Or TBproducao!txt_IE_Cliente = "" Then
        objDest.appendChild objDom.createElement("indIEDest") 'IE
        objDest.getElementsByTagName("indIEDest").Item(0).Text = 9
    Else
        If IsNull(TBproducao!txt_IE_Cliente) = False And TBproducao!txt_IE_Cliente <> "" Then
            objDest.appendChild objDom.createElement("indIEDest") 'IE
            If TBproducao!txt_IE_Cliente = "ISENTO" Or TBproducao!txt_IE_Cliente = "Isento" Or TBproducao!txt_IE_Cliente = "ISENTA" Or TBproducao!txt_IE_Cliente = "Isenta" Then
                objDest.getElementsByTagName("indIEDest").Item(0).Text = 2
            Else
                objDest.getElementsByTagName("indIEDest").Item(0).Text = 1
                objDest.appendChild objDom.createElement("IE_dest") '4
                objDest.getElementsByTagName("IE_dest").Item(0).Text = Left(DS_RetornarNumeros(TBproducao!txt_IE_Cliente), 14)
            End If
        End If
    End If
'Fecha objDest================================================================================================

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Public Sub procTotaisXML()
On Error GoTo tratar_erro

Set TBTotaisnota = CreateObject("adodb.recordset")
TBTotaisnota.Open "Select * from tbl_Totais_Nota where ID_nota = " & txtID_nota, Conexao, adOpenKeyset, adLockReadOnly
If TBTotaisnota.EOF = False Then
    'nó total dentro de Enviar (A01)
    Set objTotal = objDom.createElement("total")
    objEnviar.appendChild objTotal
    'Abre objTotal==================================================================================================
        'nó ICMStot dentro de total
        Set objICMStot = objDom.createElement("ICMStot")
        objTotal.appendChild objICMStot
        'Abre objICMStot==================================================================================================
            objICMStot.appendChild objDom.createElement("vBC_ttlnfe") '0
            objICMStot.childNodes(0).Text = Replace(IIf(IsNull(TBTotaisnota!dbl_Base_ICMS), 0, TBTotaisnota!dbl_Base_ICMS), ",", ".")
            objICMStot.appendChild objDom.createElement("vICMS_ttlnfe") '1
            objICMStot.childNodes(1).Text = Replace(IIf(IsNull(TBTotaisnota!dbl_Valor_ICMS), 0, TBTotaisnota!dbl_Valor_ICMS), ",", ".")
            'objICMStot.appendChild objDom.createElement("vICMSDeson_ttlnfe") '2
            'objICMStot.childNodes(2).Text = IIf(IsNull(TBTotaisnota!Valor_total_ICMS_desonerado), "0.00", TBTotaisnota!Valor_total_ICMS_desonerado) 'Novo layout da Sefaz (3.10) - Não é obrigatório
            objICMStot.appendChild objDom.createElement("vBCST_ttlnfe") '2
            objICMStot.childNodes(2).Text = Replace(IIf(IsNull(TBTotaisnota!dbl_Base_ICMS_Subst), "0.00", TBTotaisnota!dbl_Base_ICMS_Subst), ",", ".")
            objICMStot.appendChild objDom.createElement("vST_ttlnfe") '3
            objICMStot.childNodes(3).Text = Replace(IIf(IsNull(TBTotaisnota!dbl_Valor_ICMS_Subst), "0.00", TBTotaisnota!dbl_Valor_ICMS_Subst), ",", ".")
            objICMStot.appendChild objDom.createElement("vProd_ttlnfe") '4
            objICMStot.childNodes(4).Text = Replace(IIf(IsNull(TBTotaisnota!dbl_Valor_Total_Produtos), 0, TBTotaisnota!dbl_Valor_Total_Produtos) + IIf(IsNull(TBTotaisnota!dbl_Valor_Total_Nota_Serv), 0, TBTotaisnota!dbl_Valor_Total_Nota_Serv), ",", ".")
            objICMStot.appendChild objDom.createElement("vFrete_ttlnfe") '5
            objICMStot.childNodes(5).Text = Replace(IIf(IsNull(TBTotaisnota!dbl_Valor_Frete), "0.00", TBTotaisnota!dbl_Valor_Frete), ",", ".")
            objICMStot.appendChild objDom.createElement("vSeg_ttlnfe") '6
            objICMStot.childNodes(6).Text = Replace(IIf(IsNull(TBTotaisnota!dbl_Valor_Seguro), "0.00", TBTotaisnota!dbl_Valor_Seguro), ",", ".")
            objICMStot.appendChild objDom.createElement("vDesc_ttlnfe") '7
            objICMStot.childNodes(7).Text = Replace(IIf(IsNull(TBTotaisnota!Valor_total_desconto), 0, TBTotaisnota!Valor_total_desconto) + IIf(IsNull(TBTotaisnota!Valor_total_desconto_SUFRAMA), 0, TBTotaisnota!Valor_total_desconto_SUFRAMA), ",", ".")
            objICMStot.appendChild objDom.createElement("vII_ttlnfe") '8
            objICMStot.childNodes(8).Text = Replace(IIf(IsNull(TBTotaisnota!Valor_total_II), "0.00", TBTotaisnota!Valor_total_II), ",", ".")
            objICMStot.appendChild objDom.createElement("vIPI_ttlnfe") '9
            If TBproducao!Simples = False Then
                objICMStot.childNodes(9).Text = Replace(IIf(IsNull(TBTotaisnota!dbl_Valor_Total_IPI), "0.00", TBTotaisnota!dbl_Valor_Total_IPI), ",", ".")
            Else
                objICMStot.childNodes(9).Text = 0
            End If
            objICMStot.appendChild objDom.createElement("vPIS_ttlnfe") '10
            objICMStot.childNodes(10).Text = Replace(IIf(IsNull(TBTotaisnota!Total_PIS_prod), "0.00", TBTotaisnota!Total_PIS_prod), ",", ".")
            objICMStot.appendChild objDom.createElement("vCOFINS_ttlnfe") '11
            objICMStot.childNodes(11).Text = Replace(IIf(IsNull(TBTotaisnota!Total_Cofins_prod), "0.00", TBTotaisnota!Total_Cofins_prod), ",", ".")
            objICMStot.appendChild objDom.createElement("vCOFINS_ttlnfe") '12
            objICMStot.childNodes(12).Text = Replace(IIf(IsNull(TBTotaisnota!Total_Cofins_prod), "0.00", TBTotaisnota!Total_Cofins_prod), ",", ".")
            objICMStot.appendChild objDom.createElement("vOutro") '13
            objICMStot.childNodes(13).Text = Replace(IIf(IsNull(TBTotaisnota!dbl_Desp_Adicionais), "0.00", TBTotaisnota!dbl_Desp_Adicionais), ",", ".")
            objICMStot.appendChild objDom.createElement("vNF") '14
            objICMStot.childNodes(14).Text = Replace(IIf(IsNull(TBTotaisnota!dbl_Valor_Total_Nota), "0.00", TBTotaisnota!dbl_Valor_Total_Nota), ",", ".")
            objICMStot.appendChild objDom.createElement("vTotTrib_ttlnfe") '15
            If IsNull(TBTotaisnota!Valor_total_aprox_tributos) = False And TBTotaisnota!Valor_total_aprox_tributos <> "" And TBTotaisnota!Valor_total_aprox_tributos <> "0" Then objICMStot.childNodes(15).Text = Replace(TBTotaisnota!Valor_total_aprox_tributos, ",", ".") Else objICMStot.childNodes(15).Text = "0.00"
            objICMStot.appendChild objDom.createElement("vFCPUFDest_ttlnfe") '16
            objICMStot.childNodes(16).Text = Replace(IIf(IsNull(TBTotaisnota!Valor_total_ICMS_FCP), "0.00", TBTotaisnota!Valor_total_ICMS_FCP), ",", ".")
            objICMStot.appendChild objDom.createElement("vICMSUFDest_ttlnfe") '17
            objICMStot.childNodes(17).Text = Replace(IIf(IsNull(TBTotaisnota!Valor_total_ICMS_INT_UF_dest), "0.00", TBTotaisnota!Valor_total_ICMS_INT_UF_dest), ",", ".")
            objICMStot.appendChild objDom.createElement("vICMSUFRemet_ttlnfe") '18
            objICMStot.childNodes(18).Text = Replace(IIf(IsNull(TBTotaisnota!Valor_total_ICMS_INT_UF_rem), "0.00", TBTotaisnota!Valor_total_ICMS_INT_UF_rem), ",", ".")
            
            objICMStot.appendChild objDom.createElement("vIPIDevol_ttlnfe") '19
            If TBproducao!Simples = True And Devolucao = True Then
                objICMStot.childNodes(19).Text = Replace(IIf(IsNull(TBTotaisnota!dbl_Valor_Total_IPI), "0.00", TBTotaisnota!dbl_Valor_Total_IPI), ",", ".")
            Else
                objICMStot.childNodes(19).Text = "0.00"
            End If
            'objICMStot.appendChild objDom.createElement("vFCP_ttlnfe") '20
            'objICMStot.appendChild objDom.createElement("vFCPST_ttlnfe") '21
            'objICMStot.appendChild objDom.createElement("vFCPSTRet_ttlnfe") '22
            'objICMStot.appendChild objDom.createElement("vIPIDevol_ttlnfe") '23
            
            DAPartilhaICMS = ""
            If objICMStot.childNodes(17).Text > 0 Then DAPartilhaICMS = "Partilha ICMS operação interestadual consumidor final, disposto na Emenda constitucional 87/2015. Valor ICMS para UF destino (" & TBproducao!txt_UF & "): R$" & Format(objICMStot.childNodes(18).Text, "###,##0.00") & ". Valor FCP para o destino: R$" & Format(objICMStot.childNodes(17).Text, "###,##0.00") & ". Valor ICMS UF remetente (" & TBproducao!UF & "): R$" & Format(objICMStot.childNodes(19).Text, "###,##0.00") & "."
        'Fecha objICMStot=================================================================================================
        
        'nó RetTrib dentro de Total
        Set objRetTrib = objDom.createElement("retTrib")
        objTotal.appendChild objRetTrib
        'Abre RetTrib==================================================================================================
            objRetTrib.appendChild objDom.createElement("vRetPIS") '0
            objRetTrib.childNodes(0).Text = Replace(IIf(IsNull(TBTotaisnota!Total_retencao_PIS), "0.00", TBTotaisnota!Total_retencao_PIS), ",", ".")
            objRetTrib.appendChild objDom.createElement("vRetCOFINS_servttlnfe") '1
            objRetTrib.childNodes(1).Text = Replace(IIf(IsNull(TBTotaisnota!Total_retencao_Cofins), "0.00", TBTotaisnota!Total_retencao_Cofins), ",", ".")
            objRetTrib.appendChild objDom.createElement("vRetCSLL") '2
            objRetTrib.childNodes(2).Text = Replace(IIf(IsNull(TBTotaisnota!Total_CSLL_serv), "0.00", TBTotaisnota!Total_CSLL_serv), ",", ".")
            objRetTrib.appendChild objDom.createElement("vBCIRRF") '3
            objRetTrib.childNodes(3).Text = Replace(IIf(IsNull(TBTotaisnota!dbl_Valor_Total_Nota_Serv), "0.00", TBTotaisnota!dbl_Valor_Total_Nota_Serv), ",", ".")
            objRetTrib.appendChild objDom.createElement("vIRRF") '4
            objRetTrib.childNodes(4).Text = Replace(IIf(IsNull(TBTotaisnota!Total_IRRF_serv), "0.00", TBTotaisnota!Total_IRRF_serv), ",", ".")
            'objRetTrib.appendChild objDom.createElement("vBCRetPrev") '5
            'objRetTrib.appendChild objDom.createElement("vRetPrev") '6
        'Fecha RetTrib=================================================================================================
    'Fecha objTotal=================================================================================================
End If
TBTotaisnota.Close

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Public Sub procTransporteXML()
On Error GoTo tratar_erro

Set TBTransporte = CreateObject("adodb.recordset")
TBTransporte.Open "Select * from tbl_Dados_Transp where ID_Nota = " & txtID_nota, Conexao, adOpenKeyset, adLockReadOnly
If TBTransporte.EOF = False Then
    'no transp (Y01) dentro de Enviar (A01)
    Set objTransp = objDom.createElement("transp")
    objEnviar.appendChild objTransp
    'Abre transp==================================================================================================
        objTransp.appendChild objDom.createElement("modFrete") '0
        objTransp.getElementsByTagName("modFrete").Item(0).Text = TBTransporte!txt_Frete_Conta 'Frete Novo layout da Sefaz (4.0)
        'objTransp.appendChild objDom.createElement("balsa") '1
        'objTransp.appendChild objDom.createElement("vagao") '2
        
        'no transporta dentro de transp
        Set objTransporta = objDom.createElement("transporta")
        objTransp.appendChild objTransporta
        'Abre Transporta==================================================================================================
            
            objTransporta.appendChild objDom.createElement("xNome_transp") '2
            objTransporta.getElementsByTagName("xNome_transp").Item(0).Text = Trim(FunTiraAcentosTexto(TBTransporte!txt_Razao))
            objTransporta.appendChild objDom.createElement("xEnder") '4
            objTransporta.appendChild objDom.createElement("xMun_transp") '5
            objTransporta.getElementsByTagName("xMun_transp").Item(0).Text = FunTiraAcentosTexto(TBTransporte!txt_Municipio)
            If IsNull(TBTransporte!txt_IE) = False And TBTransporte!txt_IE <> "" Then
                objTransporta.appendChild objDom.createElement("IE_transp")
                objTransporta.getElementsByTagName("IE_transp").Item(0).Text = IIf(TBTransporte!txt_IE = "ISENTO", "ISENTO", Left(DS_RetornarNumeros(TBTransporte!txt_IE), 14))
                objTransporta.appendChild objDom.createElement("UF_transp")
                objTransporta.getElementsByTagName("UF_transp").Item(0).Text = TBTransporte!txt_UF
            End If
            
            Familiatext = ""
            If IsNull(TBTransporte!txt_CNPJ) = False And TBTransporte!txt_CNPJ <> "" Then
                Set TBFornecedor = CreateObject("adodb.recordset")
                TBFornecedor.Open "Select * from Compras_fornecedores where IDCliente = " & TBTransporte!IdIntTransp & " and Nome_Razao = '" & TBTransporte!txt_Razao & "'", Conexao, adOpenKeyset, adLockReadOnly
                If TBFornecedor.EOF = False Then
                    If Left(TBFornecedor!Pessoa, 1) = "J" Then
                        objTransporta.appendChild objDom.createElement("CNPJ_transp") '0
                        objTransporta.getElementsByTagName("CNPJ_transp").Item(0).Text = DS_RetornarNumeros(TBTransporte!txt_CNPJ)
                    Else
                        objTransporta.appendChild objDom.createElement("CPF_transp") '1
                        objTransporta.getElementsByTagName("CPF_transp").Item(0).Text = DS_RetornarNumeros(TBTransporte!txt_CNPJ)
                    End If
                    If IsNull(TBTransporte!txt_Endereco) = False And TBTransporte!txt_Endereco <> "" Then Familiatext = TBTransporte!txt_Endereco
                    If IsNull(TBTransporte!int_numero) = False And TBTransporte!int_numero <> "" Then
                        If Familiatext <> "" Then Familiatext = Familiatext & ", " & TBTransporte!int_numero Else Familiatext = TBTransporte!int_numero
                    End If
                    If IsNull(TBFornecedor!Bairro) = False And TBFornecedor!Bairro <> "" Then
                        If Familiatext <> "" Then Familiatext = Familiatext & " - " & TBFornecedor!Bairro Else Familiatext = TBFornecedor!Bairro
                    End If
                Else
                    Set TBFornecedor = CreateObject("adodb.recordset")
                    TBFornecedor.Open "Select * from Clientes where IDCliente = " & TBTransporte!IdIntTransp & " and NomeRazao = '" & TBTransporte!txt_Razao & "'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBFornecedor.EOF = False Then
                        If Left(TBFornecedor!Tipo, 1) = "J" Then
                            objTransporta.appendChild objDom.createElement("CNPJ_transp") '0
                            objTransporta.getElementsByTagName("CNPJ_transp").Item(0).Text = DS_RetornarNumeros(TBTransporte!txt_CNPJ)
                        Else
                            objTransporta.appendChild objDom.createElement("CPF_transp") '1
                            objTransporta.getElementsByTagName("CPF_transp").Item(0).Text = DS_RetornarNumeros(TBTransporte!txt_CNPJ)
                        End If
                        If IsNull(TBTransporte!txt_Endereco) = False And TBTransporte!txt_Endereco <> "" Then Familiatext = TBTransporte!txt_Endereco
                        If IsNull(TBTransporte!int_numero) = False And TBTransporte!int_numero <> "" Then
                            If Familiatext <> "" Then Familiatext = Familiatext & ", " & TBTransporte!int_numero Else Familiatext = TBTransporte!int_numero
                        End If
                        If IsNull(TBFornecedor!Bairro) = False And TBFornecedor!Bairro <> "" Then
                            If Familiatext <> "" Then Familiatext = Familiatext & " - " & TBFornecedor!Bairro Else Familiatext = TBFornecedor!Bairro
                        End If
                    Else
                        Set TBFornecedor = CreateObject("adodb.recordset")
                        TBFornecedor.Open "Select * from Empresa where Codigo = " & TBTransporte!IdIntTransp, Conexao, adOpenKeyset, adLockOptimistic
                        If TBFornecedor.EOF = False Then
                            objTransporta.appendChild objDom.createElement("CNPJ_transp") '0
                            objTransporta.getElementsByTagName("CNPJ_transp").Item(0).Text = DS_RetornarNumeros(TBTransporte!txt_CNPJ)
                            If IsNull(TBTransporte!txt_Endereco) = False And TBTransporte!txt_Endereco <> "" Then Familiatext = TBTransporte!txt_Endereco
                            If IsNull(TBTransporte!int_numero) = False And TBTransporte!int_numero <> "" Then
                                If Familiatext <> "" Then Familiatext = Familiatext & ", " & TBTransporte!int_numero Else Familiatext = TBTransporte!int_numero
                            End If
                            If IsNull(TBFornecedor!Bairro) = False And TBFornecedor!Bairro <> "" Then
                                If Familiatext <> "" Then Familiatext = Familiatext & " - " & TBFornecedor!Bairro Else Familiatext = TBFornecedor!Bairro
                            End If
                        End If
                    End If
                End If
                TBFornecedor.Close
                If Familiatext <> "" Then
                    Familiatext = Left(Familiatext, 60)
                    objTransporta.getElementsByTagName("xEnder").Item(0).Text = Trim(FunTiraAcentosTexto(Familiatext))
                End If
            End If
        'Fecha Transporta=================================================================================================
        
        If TBTransporte!txt_Placa <> "" And IsNull(TBTransporte!txt_Placa) = False And TBproducao!UF = TBproducao!txt_UF Then
            'no VeicTransp dentro de Transp (Y01)
            Set objVeicTransp = objDom.createElement("veicTransp")
            objTransp.appendChild objVeicTransp
            'Abre VeicTransp==================================================================================================
                objVeicTransp.appendChild objDom.createElement("placa")
                objVeicTransp.getElementsByTagName("placa").Item(0).Text = TBTransporte!txt_Placa
                If TBTransporte!txt_UF_Placa <> "" And IsNull(TBTransporte!txt_UF_Placa) = False Then
                    objVeicTransp.appendChild objDom.createElement("UF_veictransp") '1
                    objVeicTransp.getElementsByTagName("UF_veictransp").Item(0).Text = TBTransporte!txt_UF_Placa
                End If
                If IsNull(TBTransporte!Codigo_ANTT) = False Then
                    objVeicTransp.appendChild objDom.createElement("RNTC") '2
                    objVeicTransp.getElementsByTagName("RNTC").Item(0).Text = TBTransporte!Codigo_ANTT
                End If
            'Fecha VeicTransp=================================================================================================
            
            'no Reboque dentro de Transp (Y01)
            Set objReboque = objDom.createElement("reboque")
            objTransp.appendChild objReboque
            'Abre reboque==================================================================================================
                'no ReboqueItem dentro de Reboque
                Set objReboqueItem = objDom.createElement("reboqueItem")
                objReboque.appendChild objReboqueItem
                'Abre ReboqueItem==================================================================================================
                    objReboqueItem.appendChild objDom.createElement("placa_rebtransp")
                    objReboqueItem.getElementsByTagName("placa_rebtransp").Item(0).Text = TBTransporte!txt_Placa
                    If TBTransporte!txt_UF_Placa <> "" And IsNull(TBTransporte!txt_UF_Placa) = False Then
                        objReboqueItem.appendChild objDom.createElement("UF_rebtransp")
                        objReboqueItem.getElementsByTagName("UF_rebtransp").Item(0).Text = TBTransporte!txt_UF_Placa
                    End If
                    If IsNull(TBTransporte!Codigo_ANTT) = False Then
                        objReboqueItem.appendChild objDom.createElement("RNTC_rebtransp")
                        objReboqueItem.getElementsByTagName("RNTC_rebtransp").Item(0).Text = TBTransporte!Codigo_ANTT
                    End If
                'Fecha ReboqueItem=================================================================================================
            'Fecha reboque=================================================================================================
        End If
        
        'no vol dentro de Transp (Y01)
        Set objVol = objDom.createElement("vol")
        objTransp.appendChild objVol
        'Abre Vol==================================================================================================
            'no volItem dentro de Vol
            Set objVolItem = objDom.createElement("volItem")
            objVol.appendChild objVolItem
            'Abre volItem==================================================================================================
                objVolItem.appendChild objDom.createElement("qVol") '0
                objVolItem.getElementsByTagName("qVol").Item(0).Text = Replace(IIf(IsNull(TBTransporte!int_Qtd_Transp), 0, TBTransporte!int_Qtd_Transp), ",", ".")
                If IsNull(TBTransporte!txt_Especie) = False Then
                    objVolItem.appendChild objDom.createElement("esp") '1
                    objVolItem.getElementsByTagName("esp").Item(0).Text = Trim(TBTransporte!txt_Especie)
                End If
                If IsNull(TBTransporte!txt_Marca) = False Then
                    objVolItem.appendChild objDom.createElement("marca") '2
                    objVolItem.getElementsByTagName("esp").Item(0).Text = Trim(TBTransporte!txt_Marca)
                End If
                objVolItem.appendChild objDom.createElement("nVol") '3
                objVolItem.getElementsByTagName("nVol").Item(0).Text = IIf(IsNull(TBTransporte!Numeracao), 0, TBTransporte!Numeracao)
                objVolItem.appendChild objDom.createElement("pesoL_transp") '4
                objVolItem.getElementsByTagName("pesoL_transp").Item(0).Text = Replace(IIf(IsNull(TBTransporte!dbl_Peso_Liquido), 0, TBTransporte!dbl_Peso_Liquido), ",", ".")
                objVolItem.appendChild objDom.createElement("pesoB_transp") '5
                objVolItem.getElementsByTagName("pesoB_transp").Item(0).Text = Replace(IIf(IsNull(TBTransporte!dbl_Peso_Bruto), 0, TBTransporte!dbl_Peso_Bruto), ",", ".")
            'Fecha volItem=================================================================================================
        'Fecha Vol=================================================================================================
    'Fecha transp=================================================================================================
End If
TBTransporte.Close

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Public Sub procAdicionaisXML()
On Error GoTo tratar_erro

'no infAdic (AB01) dentro de Enviar (A01)
Set objInfAdic = objDom.createElement("infAdic")
objEnviar.appendChild objInfAdic
'Abre InfAdic====================================================================================================
    
    objInfAdic.appendChild objDom.createElement("infCpl") '1
    
    Familiatext = ""
    DadosAdicionaisTexto = ""
    Set TBControleNF = CreateObject("adodb.recordset")
    TBControleNF.Open "Select * from tbl_DadosAdicionais where ID_nota = " & txtID_nota, Conexao, adOpenKeyset, adLockReadOnly
    If TBControleNF.EOF = False Then
        If IsNull(TBControleNF!mem_corpo) = False And TBControleNF!mem_corpo <> "" Then
            objInfAdic.appendChild objDom.createElement("infAdFisco") '0
            objInfAdic.getElementsByTagName("infAdFisco").Item(0).Text = FunTiraAcentosTexto(Trim(TBControleNF!mem_corpo))
        End If
        
        If IsNull(TBControleNF!mem_DadosAdicionais) = False And TBControleNF!mem_DadosAdicionais <> "" Then
            DadosAdicionaisTexto = FunTiraAcentosTexto(Trim(TBControleNF!mem_DadosAdicionais))
        Else
            DadosAdicionaisTexto = ""
        End If
    End If
    TBControleNF.Close
    
    endereco_entrega = ""
    If TBproducao!DA_entrega = True Then
        Set TBClientes = CreateObject("adodb.recordset")
        TBClientes.Open "Select * from clientes_entrega where identrega = " & TBproducao!ID_entrega, Conexao, adOpenKeyset, adLockReadOnly
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
    If TBproducao!DA_cobranca = True Then
        Set TBClientes = CreateObject("adodb.recordset")
        TBClientes.Open "Select * from clientes_cobranca where idcobranca = " & TBproducao!ID_Cobranca, Conexao, adOpenKeyset, adLockReadOnly
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
        objInfAdic.getElementsByTagName("infCpl").Item(0).Text = FunTiraAcentosTexto(LTrim(Trim(Familiatext)))
    End If
'Fecha InfAdic===================================================================================================

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Function funVerificacaoEnviar() As Boolean
On Error GoTo tratar_erro

funVerificacaoEnviar = True

If funVerifLiberacao(True) = False Then
    funVerificacaoEnviar = False
    Exit Function
End If

'Verifica se a cidade está cadastrada corretamente
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from tbl_Dados_Nota_Fiscal where ID = " & txtID_nota, Conexao, adOpenKeyset, adLockReadOnly
If TBAbrir.EOF = False Then
    FamiliaAntiga = FunTiraAcentosTexto(TBAbrir!txt_Municipio)
    
    If IsNull(TBAbrir!txt_UF) = False And TBAbrir!txt_UF <> "" And TBAbrir!txt_UF <> "EX" Then
        Set TBFI = CreateObject("adodb.recordset")
        TBFI.Open "Select * from CEP where Municipio = '" & FamiliaAntiga & "'", Conexao, adOpenKeyset, adLockReadOnly
        If TBFI.EOF = True Then
            MsgBox ("Não é permitido liberar esta nota fiscal para envio, pois a nota esta com a cidade errada."), vbExclamation
            funVerificacaoEnviar = False
            TBFI.Close
            Exit Function
        End If
        
        Set TBFI = CreateObject("adodb.recordset")
        TBFI.Open "Select * from CEP where Sigla_UF = '" & TBAbrir!txt_UF & "'", Conexao, adOpenKeyset, adLockReadOnly
        If TBFI.EOF = True Then
            MsgBox ("Não é permitido liberar esta nota fiscal para envio, pois a nota esta com o estado errado."), vbExclamation
            funVerificacaoEnviar = False
            TBFI.Close
            Exit Function
        End If
        
        Set TBFI = CreateObject("adodb.recordset")
        TBFI.Open "Select * from CEP where Municipio = '" & FamiliaAntiga & "' and Sigla_UF = '" & TBAbrir!txt_UF & "'", Conexao, adOpenKeyset, adLockReadOnly
        If TBFI.EOF = True Then
            MsgBox ("Não é permitido liberar esta nota fiscal para envio, pois não existe o munícipio " & FamiliaAntiga & " no estado " & UF & " na tabela CEP."), vbExclamation
            funVerificacaoEnviar = False
            TBFI.Close
            Exit Function
        End If
    End If
    
    'Verifica se tem país cadastrado
    If TBAbrir!txt_tipocliente = "JP" Or TBAbrir!txt_tipocliente = "JR" Or TBAbrir!txt_tipocliente = "FP" Or TBAbrir!txt_tipocliente = "FR" Then
        'Cliente
        Set TBClientes = CreateObject("adodb.recordset")
        TBClientes.Open "Select * from Clientes where IDcliente = " & TBAbrir!Id_Int_Cliente & " and NomeRazao = '" & TBAbrir!txt_Razao_Nome & "'", Conexao, adOpenKeyset, adLockReadOnly
        If TBClientes.EOF = False Then
            If IsNull(TBClientes!Codigo_pais) = True Or TBClientes!Codigo_pais = "" Then
                MsgBox ("Não é permitido liberar esta nota fiscal para envio, pois este cliente não tem país cadastrado."), vbExclamation
                funVerificacaoEnviar = False
                TBClientes.Close
                Exit Function
            End If
        End If
    Else
        'Fornecedor
        Set TBClientes = CreateObject("adodb.recordset")
        TBClientes.Open "Select * from Compras_fornecedores where IDcliente = " & TBAbrir!Id_Int_Cliente & " and Nome_Razao = '" & TBAbrir!txt_Razao_Nome & "'", Conexao, adOpenKeyset, adLockReadOnly
        If TBClientes.EOF = False Then
            If IsNull(TBClientes!Codigo_pais) = True Or TBClientes!Codigo_pais = "" Then
                MsgBox ("Não é permitido liberar esta nota fiscal para envio, pois este fornecedor não tem país cadastrado."), vbExclamation
                funVerificacaoEnviar = False
                TBClientes.Close
                Exit Function
            End If
        End If
    End If
    TBClientes.Close
    
    'Verifica se tem foi gerado as dúplicatas quando for CFOP de vendas ou mão de obra
    Set TBFI = CreateObject("adodb.recordset")
    TBFI.Open "Select CFOP.* from tbl_Detalhes_Nota NFP INNER JOIN tbl_NaturezaOperacao CFOP ON CFOP.IDCountCfop = NFP.ID_CFOP where NFP.ID_nota = " & TBAbrir!ID & " and (CFOP.Vendas = 'True' or CFOP.MaoObra = 'True')", Conexao, adOpenKeyset, adLockReadOnly
    If TBFI.EOF = False Then
        Set TBFIltro = CreateObject("adodb.recordset")
        TBFIltro.Open "Select * from tbl_Detalhes_Recebimento where ID_nota = " & TBAbrir!ID, Conexao, adOpenKeyset, adLockReadOnly
        If TBFIltro.EOF = True Then
            If MsgBox("A(s) duplicata(s) ainda não foi(ram) gerada(s), deseja prosseguir assim mesmo?", vbQuestion + vbYesNo) = vbNo Then
                funVerificacaoEnviar = False
                TBFI.Close
                TBFIltro.Close
                Exit Function
            End If
        End If
        TBFIltro.Close
    End If
    TBFI.Close
End If
TBAbrir.Close

Set TBMaquinas = CreateObject("adodb.recordset")
TBMaquinas.Open "Select * from Empresa where Empresa = '" & ListaNota.SelectedItem.ListSubItems(1) & "' and GNFe = 'True'", Conexao, adOpenKeyset, adLockReadOnly
If TBMaquinas.EOF = False Then
    'Verifica se esta preenchido o caminho para salvar o arquivo de envio da NFe
    If IsNull(TBMaquinas!Caminho_Nfe) = True Or TBMaquinas!Caminho_Nfe = "" Then
        MsgBox ("Não é permitido liberar a nota fiscal para envio, pois não foi informado o caminho onde será armazenado os aquivos para envio."), vbExclamation
        funVerificacaoEnviar = False
        Exit Function
    End If
    'Verificar se o caminho existe
    If GerArqPastas.FolderExists(TBMaquinas!Caminho_Nfe) = False Then
        MsgBox ("Não é permitido liberar a nota fiscal para envio, pois não foi encontrado o caminho " & TBMaquinas!Caminho_Nfe & ", onde será armazenado os aquivos para envio."), vbExclamation
        funVerificacaoEnviar = False
        Exit Function
    End If
End If
TBMaquinas.Close

'Verifica se é nota fiscal de devolução ou complementar e se esta referenciado a nota fiscal
Set TBMaquinas = CreateObject("adodb.recordset")
TBMaquinas.Open "Select NFE.ID from tbl_Dados_Nota_Fiscal_NFe NFE LEFT JOIN Faturamento_Relacionamento FR ON FR.ID_nota = NFE.ID_nota where NFE.ID_nota = " & txtID_nota & " and NFE.Finalidade_emissao <> 1 and NFE.Finalidade_emissao <> 3 and FR.ID IS NULL", Conexao, adOpenKeyset, adLockReadOnly
If TBMaquinas.EOF = False Then
    MsgBox ("Não é permitido liberar a nota fiscal para envio, pois não foi feito o relacionamento."), vbExclamation
    funVerificacaoEnviar = False
    Exit Function
End If
TBMaquinas.Close

'Verifica se o clinte é fisico e esta com cnpj e vice e versa
Set TBMaquinas = CreateObject("adodb.recordset")
TBMaquinas.Open "Select txt_tipocliente, txt_CNPJ_CPF from tbl_Dados_Nota_Fiscal where ID = " & txtID_nota & " and txt_uf <> 'EX'", Conexao, adOpenKeyset, adLockReadOnly
If TBMaquinas.EOF = False Then
    If Left(TBMaquinas!txt_tipocliente, 1) = "J" And Len(TBMaquinas!txt_CNPJ_CPF) < 14 Then
        MsgBox ("Não é permitido liberar a nota fiscal para envio, pois o CNPJ do destinatario esta errado."), vbExclamation
        TBMaquinas.Close
        funVerificacaoEnviar = False
        Exit Function
    ElseIf Left(TBMaquinas!txt_tipocliente, 1) = "F" And Len(TBMaquinas!txt_CNPJ_CPF) > 14 Then
        MsgBox ("Não é permitido liberar a nota fiscal para envio, pois o CPF do destinatario esta errado."), vbExclamation
        TBMaquinas.Close
        funVerificacaoEnviar = False
        Exit Function
    End If
End If
TBMaquinas.Close

Exit Function
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Function

Public Sub procMontaEmail()
On Error GoTo tratar_erro

Email = ""
'Verifica email e país
Set TBClientes = CreateObject("adodb.recordset")
If TBproducao!txt_tipocliente = "E" Then
    'Empresa
    TBClientes.Open "Select Email, Pais, Codigo_pais from Empresa where Codigo = " & TBproducao!Id_Int_Cliente, Conexao, adOpenKeyset, adLockReadOnly
    If TBClientes.EOF = False Then
        Email = IIf(IsNull(TBClientes!Email), "", TBClientes!Email)
        Pais = TBClientes!Pais
        Codigo_pais = TBClientes!Codigo_pais
    End If
ElseIf TBproducao!txt_tipocliente = "JP" Or TBproducao!txt_tipocliente = "JR" Or TBproducao!txt_tipocliente = "FP" Or TBproducao!txt_tipocliente = "FR" Then
        'Cliente
        TBClientes.Open "Select Email, Pais, Codigo_pais from Clientes where IDcliente = " & TBproducao!Id_Int_Cliente & " and NomeRazao = '" & TBproducao!txt_Razao_Nome & "' and Enviar_NF = 'True'", Conexao, adOpenKeyset, adLockReadOnly
        If TBClientes.EOF = False Then
            Email = IIf(IsNull(TBClientes!Email), "", TBClientes!Email)
            If Email <> "" Then TextoFiltro = " and Email <> '" & Email & "'" Else TextoFiltro = ""
            
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select Email from Clientes_Contatos where IDcliente = " & TBproducao!Id_Int_Cliente & TextoFiltro & " and Enviar_NFe = 'True' and EMail is not null", Conexao, adOpenKeyset, adLockReadOnly
            If TBFI.EOF = False Then
                Do While TBFI.EOF = False
                    If IsNull(TBFI!Email) = False And TBFI!Email <> "" Then
                        If Email <> "" Then Email = Email & "; " & TBFI!Email Else Email = TBFI!Email
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
        TBClientes.Open "Select Email, Pais, Codigo_pais from Compras_fornecedores where IDcliente = " & TBproducao!Id_Int_Cliente & " and Nome_Razao = '" & TBproducao!txt_Razao_Nome & "' and Enviar_NF = 'True'", Conexao, adOpenKeyset, adLockReadOnly
        If TBClientes.EOF = False Then
            Email = IIf(IsNull(TBClientes!Email), "", TBClientes!Email)
            If Email <> "" Then TextoFiltro = " and Email <> '" & Email & "'" Else TextoFiltro = ""
            
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select Email from Contatos_fornecedor where IdFornecedor = " & TBproducao!Id_Int_Cliente & TextoFiltro & " and Enviar_NFe = 'True' and Email is not null", Conexao, adOpenKeyset, adLockReadOnly
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
TBAfericao.Open "Select CNPJ from Empresa where Codigo = " & TBproducao!ID_empresa, Conexao, adOpenKeyset, adLockReadOnly
If TBAfericao.EOF = False Then
    Set TBFIltro = CreateObject("adodb.recordset")
    TBFIltro.Open "Select IdIntTransp, txt_Razao from tbl_Dados_Transp where ID_nota = " & TBproducao!ID & " and TXT_CNPJ <> '" & TBAfericao!CNPJ & "' and TXT_CNPJ <> '" & TBproducao!txt_CNPJ_CPF & "'", Conexao, adOpenKeyset, adLockReadOnly
    If TBFIltro.EOF = False Then
        'Cliente
        Set TBClientes = CreateObject("adodb.recordset")
        TBClientes.Open "Select Email from Clientes where IDcliente = " & TBFIltro!IdIntTransp & " and NomeRazao = '" & TBFIltro!txt_Razao & "'", Conexao, adOpenKeyset, adLockReadOnly
        If TBClientes.EOF = False Then
            Email1 = IIf(IsNull(TBClientes!Email), "", TBClientes!Email)
            If Email1 <> "" Then
                TextoFiltro = " and Email <> '" & Email1 & "'"
                If Email <> "" Then Email = Email & ";" & Email1 Else Email = Email1
            Else
                TextoFiltro = ""
            End If
            
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select Email from Clientes_Contatos where IDcliente = " & TBFIltro!IdIntTransp & TextoFiltro & " and Enviar_NFe = 'True' and EMail is not null", Conexao, adOpenKeyset, adLockReadOnly
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
            TBClientes.Open "Select Email from Compras_fornecedores where IDcliente = " & TBFIltro!IdIntTransp & " and Nome_Razao = '" & TBFIltro!txt_Razao & "'", Conexao, adOpenKeyset, adLockReadOnly
            If TBClientes.EOF = False Then
                Email1 = IIf(IsNull(TBClientes!Email), "", TBClientes!Email)
                If Email1 <> "" Then
                    TextoFiltro = " and Email <> '" & Email1 & "'"
                    If Email <> "" Then Email = Email & ";" & Email1 Else Email = Email1
                Else
                    TextoFiltro = ""
                End If
                
                Set TBFI = CreateObject("adodb.recordset")
                TBFI.Open "Select Email from Contatos_fornecedor where IdFornecedor = " & TBFIltro!IdIntTransp & TextoFiltro & " and Enviar_NFe = 'True' and Email is not null", Conexao, adOpenKeyset, adLockReadOnly
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

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Public Sub procCarregaEmpresa()
On Error GoTo tratar_erro

UF_transp = ""
Cidade = ""
Set TBFI = CreateObject("adodb.recordset")
TBFI.Open "Select E.UF, E.Cidade, E.CNPJ E.Caminho_Nfe, E.Caminho_XMLDanfe, E.Caminho_RetornoNfe, N.Obs from Empresa E INNER JOIN tbl_Dados_Nota_Fiscal N ON E.Codigo = N.ID_empresa where N.ID = " & txtID_nota, Conexao, adOpenKeyset, adLockReadOnly
If TBFI.EOF = False Then
    UF_transp = IIf(IsNull(TBFI!UF), "", TBFI!UF)
    Cidade = IIf(IsNull(TBFI!Cidade), "", TBFI!Cidade)
    CnpjNF = IIf(IsNull(TBFI!CNPJ), "", TBFI!CNPJ)
    DiretorioEnvio = IIf(IsNull(TBFI!Caminho_Nfe), "", TBFI!Caminho_Nfe)
    DiretorioXMLDanfe = IIf(IsNull(TBFI!Caminho_XMLDanfe), "", TBFI!Caminho_XMLDanfe)
    DiretorioRetorno = IIf(IsNull(TBFI!Caminho_RetornoNfe), "", TBFI!Caminho_RetornoNfe)
    txtMotivo = IIf(IsNull(TBFI!Obs), "", TBFI!Obs)
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Public Sub ProcImprimir()
On Error GoTo tratar_erro

If txtID_nota = 0 Then
    MsgBox ("Informe a nota fiscal antes de consultar o status."), vbExclamation
    Exit Sub
End If

Set TBproducao = CreateObject("adodb.recordset")
TBproducao.Open "Select status from tbl_Dados_Nota_Fiscal_NFe WHERE ID_nota = " & txtID_nota & " AND status NOT IN (100,101)", Conexao, adOpenKeyset, adLockReadOnly
If TBproducao.EOF = False Then
    MsgBox ("Só é possível visualizar impressão de notas autorizadas ou canceladas."), vbExclamation
    TBproducao.Close
    Exit Sub
End If
TBproducao.Close

Set TBproducao = CreateObject("adodb.recordset")
TBproducao.Open "Select E.CNPJ from tbl_Dados_Nota_Fiscal NF INNER JOIN Empresa E ON NF.ID_empresa = E.Codigo WHERE NF.ID = " & txtID_nota, Conexao, adOpenKeyset, adLockReadOnly
If TBproducao.EOF = False Then
    procAbrirNotaPDF "NFe", TBproducao!CNPJ, txtNota, txtSerie, DiretorioXMLDanfe, False
End If
TBproducao.Close
   
Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Public Sub procCarregaTransp()
On Error GoTo tratar_erro

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from tbl_dados_transp where ID_Nota = " & txtID_nota, Conexao, adOpenKeyset, adLockReadOnly
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

Private Sub ProcCancelar()
On Error GoTo tratar_erro

If txtID_nota = 0 Then
    MsgBox ("Informe a nota fiscal antes de cancelar."), vbExclamation
    Exit Sub
End If

Set TBproducao = CreateObject("adodb.recordset")
TBproducao.Open "Select status from tbl_Dados_Nota_Fiscal_NFe WHERE ID_nota = " & txtID_nota & " AND status <> 100", Conexao, adOpenKeyset, adLockReadOnly
If TBproducao.EOF = False Then
    MsgBox ("Só é possível cancelar notas aprovadas."), vbExclamation
    TBproducao.Close
    Exit Sub
End If

Acao = "cancelar"
If procVerificaMigrate = False Then Exit Sub

Set TBEstoque = CreateObject("adodb.recordset")
TBEstoque.Open "Select EM.Idoperacao from (tbl_dados_nota_fiscal NF INNER JOIN tbl_Detalhes_Nota NFP ON NF.ID = NFP.ID_nota) INNER JOIN Estoque_movimentacao EM ON EM.ID_prod_NF = NFP.Int_codigo where NF.ID = " & txtID_nota & " and NF.Aplicacao = 'P' and EM.Documento = '" & txtNota & "' and (EM.Operacao = 'SAIDA_NOTA' or EM.Operacao = 'SAIDA_NOTA_PARCIAL')", Conexao, adOpenKeyset, adLockReadOnly
If TBEstoque.EOF = False Then
    MsgBox ("Não é permitido cancelar esta nota fiscal, pois a mesma já baixou estoque."), vbExclamation
    TBEstoque.Close
    Exit Sub
End If
TBEstoque.Close
                  
If MsgBox("Deseja cancelar esta nota fiscal?", vbQuestion + vbYesNo) = vbYes Then
    
Mensagem:
    TextoCancelamento = InputBox("Favor informar o motivo do cancelamento.")
    If TextoCancelamento = "" Then Exit Sub
    If Len(TextoCancelamento) < 15 Then
        MsgBox ("O motivo deve possuir um valor minimo de 15 caracteres!"), vbExclamation
        GoTo Mensagem
    End If
    NomeArquivo = "NF" & txtNota & txtSerie & "C"
    procCancelarXML
    TipoXML = 2
    procAcionaTimer
End If
    
Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Public Sub procCancelarXML()
On Error GoTo tratar_erro

Dim objDom As DOMDocument50
Dim objEnvioEvento As IXMLDOMElement
Dim objEvento As IXMLDOMElement
Dim objEveInf As IXMLDOMElement
Dim objEvedet As IXMLDOMElement

Set objDom = New DOMDocument50

'nó EnvioEvento
Set objEnvioEvento = objDom.createElement("EnvioEvento")
objDom.appendChild objEnvioEvento
'Abre EnvioEvento======================================================================================================
    objEnvioEvento.appendChild objDom.createElement("ModeloDocumento")
    objEnvioEvento.childNodes(0).Text = "NFe"
    objEnvioEvento.appendChild objDom.createElement("Versao")
    objEnvioEvento.childNodes(1).Text = "4.00"
    objEnvioEvento.appendChild objDom.createElement("ChaveParceiro") 'Chave da caprind que a Migrate emite
    objEnvioEvento.childNodes(2).Text = "TsDpg/TtLpSXBO5uVUMM3w=="
    objEnvioEvento.appendChild objDom.createElement("ChaveAcesso") 'Chave do cliente que a migrate emite
    
    'nó Evento
    Set objEvento = objDom.createElement("Evento")
    objEnvioEvento.appendChild objEvento
    'Abre Evento===================================================================================================
        objEvento.appendChild objDom.createElement("NtfCnpjEmissor")
        objEvento.childNodes(0).Text = DS_RetornarNumeros(CnpjNF)
        objEvento.appendChild objDom.createElement("NtfNumero")
        objEvento.childNodes(1).Text = Format(txtNota, "0")
        objEvento.appendChild objDom.createElement("NtfSerie")
        objEvento.childNodes(2).Text = txtSerie
        objEvento.appendChild objDom.createElement("tpAmb")
        objEvento.childNodes(3).Text = 1 '1-Produção 2-homologação
        
        'nó EveInf
        Set objEveInf = objDom.createElement("EveInf")
        objEvento.appendChild objEveInf
        'Abre EveInf===================================================================================================
            objEveInf.appendChild objDom.createElement("EveDh")
            objEveInf.childNodes(0).Text = Format(Now, "yyyy-mm-dd") & "T" & Format(Now, "HH:mm:ss")
            objEveInf.appendChild objDom.createElement("EveFusoHorario")
            objEveInf.childNodes(1).Text = FunVerifFusoHorario(True)
            objEveInf.appendChild objDom.createElement("EveTp")
            objEveInf.childNodes(2).Text = "110111"
            objEveInf.appendChild objDom.createElement("EvenSeq")
            objEveInf.childNodes(3).Text = "1"
            
            'nó EveInf
            Set objEvedet = objDom.createElement("Evedet")
            objEveInf.appendChild objEvedet
            'Abre EveInf===================================================================================================
                objEvedet.appendChild objDom.createElement("EveDesc")
                objEvedet.childNodes(0).Text = "Cancelamento"
                objEvedet.appendChild objDom.createElement("EvenProt")
                objEvedet.childNodes(1).Text = "0"
                objEvedet.appendChild objDom.createElement("EvexJust")
                objEvedet.childNodes(2).Text = TextoCancelamento
            'Fecha EveInf==================================================================================================
        'Fecha EveInf==================================================================================================
    'Fecha Evento==================================================================================================
'Fecha EnvioEvento===============================================================================================================
                
objDom.Save (DiretorioEnvio & "/" & NomeArquivo & ".xml")

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Function procVerificaMigrate() As Boolean
On Error GoTo tratar_erro
procVerificaMigrate = False

If DiretorioEnvio = "" Then
    NomeCampo = "o diretório de envio no cadastro da empresa"
    ProcVerificaAcao
    Exit Function
End If

If DiretorioRetorno = "" Then
    NomeCampo = "o diretório de retorno no cadastro da empresa"
    ProcVerificaAcao
    Exit Function
End If

If DiretorioXMLDanfe = "" Then
    NomeCampo = "o diretório de XML e Danfe no cadastro da empresa"
    ProcVerificaAcao
    Exit Function
End If

procVerificaMigrate = True

Exit Function
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Function

Public Sub procCancelarTabelas()
On Error GoTo tratar_erro

ProcExcluirRelacionamentoNF txtID_nota
ProcExcluirArquivosRemessa txtID_nota

Set TBproducao = CreateObject("adodb.recordset")
TBproducao.Open "Select int_TipoNota, txt_tipocliente from tbl_Dados_Nota_Fiscal WHERE ID = " & txtID_nota, Conexao, adOpenKeyset, adLockReadOnly
If TBproducao.EOF = False Then
    ProcExcluirContas txtID_nota, IIf(TBproducao!int_TipoNota = 1, True, False), TBproducao!txt_tipocliente
End If
TBproducao.Close

Conexao.Execute "DELETE from ECEV from Estoque_Controle_Empenho_Vendas ECEV INNER JOIN tbl_Detalhes_Nota NFP ON NFP.Int_codigo = ECEV.ID_faturamento where NFP.ID_nota = " & txtID_nota

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Private Sub Timer_Timer()
On Error GoTo tratar_erro

PBLista.Value = PBLista.Value + 1
If Dir(DiretorioEnvio & "/" & NomeArquivo & ".xml") = "" Or PBLista.Value >= 200 Then
    Timer.Enabled = False
    procLerRetornoXML
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Public Sub procAcionaTimer()
On Error GoTo tratar_erro

PBLista.Min = 0
PBLista.Max = 200
PBLista.Value = 0
Timer.Enabled = True

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub
